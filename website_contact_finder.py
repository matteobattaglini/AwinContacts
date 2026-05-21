#!/usr/bin/env python3
"""
Website Contact Finder for missing Awin advertisers.

Strategy (no extra API keys needed):
  1. Parse the Awin program description (already in col 12) for any name/email mentions
  2. Scrape the company website:
       a. Homepage → find links to team / about / affiliate / partner pages
       b. Scrape those pages → extract names + roles from team sections
       c. Also extract any email addresses
  3. Filter:
       - Keep only company-domain emails (reject @awin.com, @agency, @gmail …)
       - Keep only people with a marketing/affiliate/digital/ecommerce role
       - Reject anyone whose email or LinkedIn suggests they work for Awin or an external agency
  4. If a name is found without email → still write it (RocketReach finds email from name+domain)
  5. Write results to Excel
"""

import re
import time
import threading
import requests
import openpyxl
from openpyxl.styles import Font
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed

# ── Config ────────────────────────────────────────────────────────────────────

INPUT_FILE  = "/home/user/AwinContacts/Awin_Contacts_complete.xlsx"
OUTPUT_FILE = "/home/user/AwinContacts/Awin_Contacts_web.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "it-IT,it;q=0.9,en;q=0.8",
}
TIMEOUT = 12

# URL path fragments → likely affiliate/team/contact page
PAGE_KEYWORDS = [
    "affiliat", "partner", "publisher", "collabora",
    "team", "chi-siamo", "about-us", "about", "who-we-are",
    "chi siamo", "chisiamo", "nostro-team", "our-team",
    "contatti", "contact", "contacto", "kontakt",
    "press", "media", "management", "leadership",
    "marketing", "lavora-con-noi", "work-with-us",
]

# Job title keywords to accept a contact
VALID_ROLE_KW = [
    "affiliat", "partner", "digital", "performance", "marketing",
    "ecommerce", "e-commerce", "growth", "brand", "cmo",
    "chief marketing", "communication", "comunicazione",
    "online", "web", "commerciale", "sales director",
    "head of", "responsabile", "direttore", "director", "manager",
    "seo", "sem", "acquisition", "retention", "crm",
]

# Domains that indicate Awin staff, external agencies, or generic mail
REJECT_DOMAINS = {
    "awin.com", "affilinet.com", "tradedoubler.com", "zanox.com",
    "webgains.com", "partnerize.com", "rakuten.com", "cj.com",
    "gmail.com", "yahoo.com", "yahoo.it", "hotmail.com",
    "outlook.com", "icloud.com", "libero.it", "virgilio.it",
    "tiscali.it", "alice.it", "tin.it",
    "pec.it", "legalmail.it", "aruba.it", "arubapec.it",
}

# Role-specific email prefixes — high priority even without a name
ROLE_PREFIXES = {
    "affiliat", "affiliate", "partner", "partnership",
    "marketing", "digital", "performance", "ecommerce",
    "brand", "growth", "acquisition", "media",
    "press", "collabora", "publisher",
}

# Company names / words in email local-part that suggest external agency
AGENCY_SIGNALS = [
    "agency", "agenzia", "studio", "consulen", "freelanc",
    "digital-agency", "web-agency",
]

EMAIL_RE = re.compile(r'\b[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}\b')

# First Last — two capitalised tokens, 2-25 chars each
NAME_RE = re.compile(
    r'\b([A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝ][a-zàáâãäåæçèéêëìíîïðñòóôõöùúûüý]{1,24})'
    r'[ \xa0]+'
    r'([A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝ][a-zàáâãäåæçèéêëìíîïðñòóôõöùúûüý]{1,24})\b'
)

# Words that look like names but aren't
FALSE_NAME_TOKENS = {
    "Cookie", "Cookies", "Privacy", "Terms", "Powered", "About", "Contact",
    "Contatto", "Contatti", "Company", "Follow", "Share", "Read", "More",
    "Learn", "Click", "Download", "Subscribe", "Iscriviti", "Sign", "Login",
    "Register", "Back", "Next", "Previous", "Home", "Shop", "Cart", "Menu",
    "Search", "Sale", "New", "Best", "Free", "View", "All", "See", "Get",
    "Our", "Your", "The", "This", "That", "With", "From", "Into", "Over",
    "Under", "After", "Before", "During", "Awin", "Azienda", "Italia",
    "Europe", "Global", "International", "Group", "Brand", "Marketing",
    "Digital", "Online", "Social", "Media", "Manager", "Director", "Team",
    "Staff", "Department", "Customer", "Service", "Style", "Concierge",
    "Luxury", "Personalized", "Collaboriamo", "Nostri", "Sustainability",
    "Bonjour", "Newsletter", "Scopri", "Clicca", "Leggi", "Inizia",
    "Email", "Invia", "Chatta",
}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _base_domain(domain: str) -> str:
    domain = domain.lower().strip().lstrip("www.")
    parts = domain.split(".")
    if len(parts) >= 3 and parts[-2] in ("co", "com", "org", "net", "gov", "edu"):
        return ".".join(parts[-3:])
    return ".".join(parts[-2:]) if len(parts) >= 2 else domain


def _email_prefix(email: str) -> str:
    return email.split("@")[0].lower()


def _is_company_email(email: str, company_domain: str) -> bool:
    """
    Accept email if:
    - Domain is not a blacklisted generic/Awin/agency domain
    - Domain shares the same brand token as the company domain
      e.g. marketing@aviliagroup.it accepted for aviliahome.it
           because both share 'avilia'
    """
    ed = email.split("@", 1)[-1].lower()
    if ed in REJECT_DOMAINS:
        return False
    if any(sig in ed for sig in AGENCY_SIGNALS):
        return False
    # Exact base-domain match
    if _base_domain(ed) == _base_domain(company_domain):
        return True
    # Shared brand token: longest token in company domain that appears in email domain
    brand_tokens = re.split(r'[\.\-_]', company_domain.lower())
    brand_tokens = [t for t in brand_tokens if len(t) >= 4 and t not in
                    {"shop","store","it","com","eu","net","org","www"}]
    return any(tok in ed for tok in brand_tokens)


def _role_score(text: str) -> int:
    tl = text.lower()
    return sum(1 for kw in VALID_ROLE_KW if kw in tl)


def _is_valid_name(first: str, last: str) -> bool:
    if first in FALSE_NAME_TOKENS or last in FALSE_NAME_TOKENS:
        return False
    if len(first) < 2 or len(last) < 2:
        return False
    if first.lower() == last.lower():
        return False
    return True


def _fetch(url: str) -> str | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        if r.status_code == 200:
            ct = r.headers.get("content-type", "")
            if "html" in ct or not ct:
                return r.text
    except Exception:
        pass
    return None


def _candidate_pages(soup: BeautifulSoup, base_url: str) -> list[str]:
    base_host = urlparse(base_url).netloc
    seen, out = set(), []
    for a in soup.find_all("a", href=True):
        href  = a.get("href", "").strip()
        text  = a.get_text(" ", strip=True).lower()
        full  = urljoin(base_url, href)
        p     = urlparse(full)
        if p.netloc and p.netloc != base_host:
            continue
        path_lower = p.path.lower()
        if any(kw in path_lower or kw in text for kw in PAGE_KEYWORDS):
            if full not in seen and full != base_url:
                seen.add(full)
                out.append(full)
    return out[:12]


def _parse_team_blocks(soup: BeautifulSoup) -> list[dict]:
    """
    Find structured team/person blocks in the page.
    Returns list of {first, last, role} dicts.
    """
    people = []

    # Strategy A: elements with class/id suggesting a person card
    person_selectors = [
        "[class*='team']", "[class*='member']", "[class*='person']",
        "[class*='staff']", "[class*='card']", "[class*='profil']",
        "[class*='people']", "[class*='author']",
    ]
    for sel in person_selectors:
        for block in soup.select(sel):
            block_text = block.get_text(" ", strip=True)
            names = NAME_RE.findall(block_text)
            for first, last in names:
                if _is_valid_name(first, last):
                    people.append({"first": first, "last": last,
                                   "role": block_text, "score": _role_score(block_text)})

    # Strategy B: heading tags followed by a role paragraph
    for tag in soup.find_all(["h2", "h3", "h4", "h5"]):
        text = tag.get_text(" ", strip=True)
        names = NAME_RE.findall(text)
        if not names:
            continue
        # Get the next sibling for role context
        sibling = tag.find_next_sibling()
        role_text = sibling.get_text(" ", strip=True) if sibling else ""
        combined  = text + " " + role_text
        for first, last in names:
            if _is_valid_name(first, last):
                people.append({"first": first, "last": last,
                               "role": combined, "score": _role_score(combined)})

    return people


def _extract_from_description(description: str) -> list[dict]:
    """
    Parse the Awin program description for any name + email mentions.
    e.g. 'Contact our affiliate manager Mario Rossi at mario@brand.it'
    """
    results = []
    emails  = EMAIL_RE.findall(description)
    for email in emails:
        idx   = description.find(email)
        chunk = description[max(0, idx - 200): idx + 50]
        names = NAME_RE.findall(chunk)
        for first, last in names:
            if _is_valid_name(first, last):
                results.append({"first": first, "last": last,
                                 "email": email.lower(), "source": "awin_description"})
    # Also look for names near contact keywords even without email
    for kw in ["contact", "contatt", "manager", "responsabile", "referente"]:
        for m in re.finditer(kw, description, re.I):
            chunk = description[m.start(): m.start() + 120]
            for first, last in NAME_RE.findall(chunk):
                if _is_valid_name(first, last):
                    results.append({"first": first, "last": last,
                                     "email": "", "source": "awin_description"})
    return results


# ── Core ──────────────────────────────────────────────────────────────────────

def _collect_emails_from_page(soup: BeautifulSoup, domain: str, source_url: str) -> list[dict]:
    """
    Collect company-domain emails from a parsed page.
    Primary: mailto: links (explicit, intentional contact points).
    Secondary: email addresses visible in text.
    """
    results = []

    # ── Primary: mailto links ────────────────────────────────────────────────
    for a in soup.find_all("a", href=re.compile(r'^mailto:', re.I)):
        raw = a["href"].replace("mailto:", "").split("?")[0].strip().lower()
        if not EMAIL_RE.match(raw):
            continue
        if not _is_company_email(raw, domain):
            continue
        ctx  = (a.parent or a).get_text(" ", strip=True)
        prefix = _email_prefix(raw)
        is_role = any(rp in prefix for rp in ROLE_PREFIXES)
        score = (5 if is_role else 1) + _role_score(ctx)
        results.append({
            "email": raw, "ctx": ctx, "score": score,
            "is_role": is_role, "source": source_url,
        })

    # ── Secondary: plain text emails ─────────────────────────────────────────
    page_text = soup.get_text(" ", strip=True)
    for email in EMAIL_RE.findall(page_text):
        el = email.lower()
        if not _is_company_email(el, domain):
            continue
        if any(r["email"] == el for r in results):
            continue  # already captured via mailto
        prefix = _email_prefix(el)
        is_role = any(rp in prefix for rp in ROLE_PREFIXES)
        if not is_role:
            continue  # skip generic info@ from plain text (too noisy)
        idx   = page_text.find(email)
        ctx   = page_text[max(0, idx - 200): idx + 50]
        score = (4 if is_role else 0) + _role_score(ctx)
        results.append({
            "email": el, "ctx": ctx, "score": score,
            "is_role": is_role, "source": source_url,
        })

    return results


def scrape_company(company: dict) -> dict:
    """
    Sources (in priority order):
      1. Awin description — direct email/name mentions
      2. Affiliate/partner/contact/collab page — mailto links
      3. Homepage — mailto links + role-prefix plain emails
    Accepts company-domain AND parent/group domain emails.
    Rejects: Awin, agencies, generic mail providers.
    """
    website     = company["website"]
    domain      = company["domain"]
    description = company.get("description", "") or ""

    raw_emails: list[dict] = []   # {email, ctx, score, is_role, source}
    pages_checked = 0

    # ── Step 1: Awin description ─────────────────────────────────────────────
    for hit in _extract_from_description(description):
        email = hit.get("email", "")
        if email:
            el = email.lower()
            if _is_company_email(el, domain):
                raw_emails.append({
                    "email": el, "ctx": description[:200],
                    "score": 15, "is_role": True, "source": "awin_description",
                })

    # ── Step 2: Website ──────────────────────────────────────────────────────
    html = _fetch(website)
    if html:
        pages_checked += 1
        soup_home = BeautifulSoup(html, "html.parser")
        extra_pages = _candidate_pages(soup_home, website)

        for url in [website] + extra_pages:
            if url != website:
                html = _fetch(url)
                if not html:
                    continue
                pages_checked += 1
            soup = BeautifulSoup(html, "html.parser")
            for tag in soup.select("script, style, [class*='cookie'], [class*='banner']"):
                tag.decompose()

            raw_emails.extend(_collect_emails_from_page(soup, domain, url))

            if len(raw_emails) >= 8:
                break

    # ── Rank + dedup emails ───────────────────────────────────────────────────
    seen_em: set[str] = set()
    ranked: list[dict] = []
    for em in sorted(raw_emails, key=lambda x: -x["score"]):
        if em["email"] not in seen_em:
            seen_em.add(em["email"])
            ranked.append(em)

    # ── Build contacts (email only; name left blank for RocketReach) ──────────
    contacts = []
    for em in ranked[:2]:
        # Try to find a name in the context text
        first, last = "", ""
        for f, l in NAME_RE.findall(em["ctx"]):
            if _is_valid_name(f, l):
                first, last = f, l
                break
        contacts.append({
            "first":  first,
            "last":   last,
            "email":  em["email"],
            "title":  "",
            "score":  em["score"],
            "source": em["source"],
        })

    return {
        **company,
        "contacts":      contacts,
        "pages_checked": pages_checked,
    }


# ── Excel ────────────────────────────────────────────────────────────────────

def load_missing_companies(xlsx_path: str) -> list[dict]:
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    company_rows: dict[str, list] = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 6).value
        if name:
            company_rows.setdefault(name, []).append(r)

    companies, seen = [], set()
    for r in range(2, ws.max_row + 1):
        name    = ws.cell(r, 6).value
        domain  = ws.cell(r, 8).value
        website = ws.cell(r, 13).value
        desc    = ws.cell(r, 12).value or ""
        awin    = ws.cell(r, 17).value or ""
        if not name or not domain or name in seen:
            continue
        seen.add(name)
        rows = company_rows.get(name, [])
        if rows and ws.cell(rows[0], 1).value:
            continue
        companies.append({
            "name":        name,
            "domain":      domain,
            "website":     website or f"https://www.{domain}",
            "description": desc,
            "awin_profile": awin,
            "rows":        rows,
        })
    return companies


def write_results(xlsx_path: str, output_path: str, results: list[dict]):
    import shutil
    shutil.copy2(xlsx_path, output_path)

    wb = openpyxl.load_workbook(output_path)
    ws = wb.active

    company_rows: dict[str, list] = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 6).value
        if name:
            company_rows.setdefault(name, []).append(r)

    normal = Font(name="Calibri", size=10)
    written = 0

    for result in results:
        if not result["contacts"]:
            continue
        name = result["name"]
        rows = company_rows.get(name, [])

        for row_idx, contact in zip(rows, result["contacts"]):
            if ws.cell(row_idx, 1).value:
                continue

            ws.cell(row_idx, 1).value = contact["first"];  ws.cell(row_idx, 1).font = normal
            ws.cell(row_idx, 2).value = contact["last"];   ws.cell(row_idx, 2).font = normal
            if contact["email"]:
                ws.cell(row_idx, 3).value = contact["email"]; ws.cell(row_idx, 3).font = normal
            if contact["title"]:
                ws.cell(row_idx, 7).value = contact["title"][:100]; ws.cell(row_idx, 7).font = normal
            written += 1

    wb.save(output_path)
    return written


# ── Runner ───────────────────────────────────────────────────────────────────

def run():
    companies = load_missing_companies(INPUT_FILE)
    total     = len(companies)
    print(f"Aziende da processare: {total}", flush=True)
    print(f"Fonti: descrizione Awin + sito web (team/about/affiliate pages)\n", flush=True)

    results = []
    hits    = 0
    lock    = threading.Lock()
    done    = [0]

    def _process(company):
        result = scrape_company(company)
        with lock:
            done[0] += 1
            found = len(result["contacts"])
            if found:
                parts = []
                for c in result["contacts"]:
                    name_str  = f"{c['first']} {c['last']}".strip() or "(no name)"
                    email_str = c['email'] or "(no email)"
                    parts.append(f"{name_str} <{email_str}>")
                print(f"  [{done[0]:>3}/{total}] HIT  {result['name']:<42} "
                      f"{'  |  '.join(parts)}", flush=True)
            elif done[0] % 25 == 0:
                print(f"  [{done[0]:>3}/{total}] — {done[0]} processate", flush=True)
        return result

    with ThreadPoolExecutor(max_workers=15) as ex:
        futures = {ex.submit(_process, c): c for c in companies}
        for future in as_completed(futures):
            r = future.result()
            results.append(r)
            if r["contacts"]:
                hits += 1

    print(f"\n{'='*60}")
    print(f"Completato")
    print(f"  Aziende processate : {total}")
    print(f"  Hit (≥1 contatto)  : {hits} ({hits/total*100:.0f}%)")
    print(f"{'='*60}\n")

    written = write_results(INPUT_FILE, OUTPUT_FILE, results)
    print(f"Scritto: {written} righe → {OUTPUT_FILE}")
    return results


if __name__ == "__main__":
    run()
