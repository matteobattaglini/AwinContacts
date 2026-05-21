#!/usr/bin/env python3
"""
Surfe Contact Finder for Awin Advertisers
Finds Affiliate / Digital Marketing contacts for each company.
"""

import requests
import time
import itertools
import openpyxl
from openpyxl.styles import Font
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

# ── API keys (round-robin) ───────────────────────────────────────────────────
API_KEYS = [
    "TL1fKLZMEE7rdWEKaIe_W-vmUvxkyP0fecMTLsbFJKA",  # reset stanotte
    # "rMY_QurAxUvl6jDeeM9xN01q5IMFBfrH5AUJJLdafe4",  # ancora esaurita
    # "KwGtwE1QeCfHU1G3NjkMyGFqf2PT36NEC5taWt0iFgY",  # ancora esaurita
    "xjs0s_35PkhFgU8L4Z1SU-m_urZsZd55nqFX8xp2noA",
    "FtZWylT6Cfe2S3f9UbV528opGBjsJ0fSYEHMzajXebE",
]
SURFE_URL = "https://api.surfe.com/v2/people/search"

# ── Job title search terms (Surfe expands semantically) ─────────────────────
JOB_TITLES_HIGH = [
    "Affiliate Marketing",
    "Digital Marketing",
    "Performance Marketing",
]
JOB_TITLES_LOW = [
    "Marketing Manager",
    "Chief Marketing Officer",
    "Marketing Specialist",
    "E-commerce Manager",
    "E-commerce Director",
]
JOB_TITLES = JOB_TITLES_HIGH + JOB_TITLES_LOW

# Client-side priority ranking (lower index = higher priority)
PRIORITY_KEYWORDS = [
    "affiliate",
    "digital marketing",
    "performance marketing",
    "online marketing",
    "cmo",
    "chief marketing",
    "marketing manager",
    "marketing specialist",
    "marketing",
    "ecommerce",
    "e-commerce",
]

# ── Country TLD → ISO-2 code ─────────────────────────────────────────────────
TLD_COUNTRY = {
    "it": "it", "de": "de", "fr": "fr", "es": "es",
    "nl": "nl", "pl": "pl", "pt": "pt", "be": "be",
    "se": "se", "dk": "dk", "fi": "fi", "no": "no",
    "ch": "ch", "at": "at", "ro": "ro", "cz": "cz",
    "hu": "hu", "gr": "gr", "ru": "ru", "tr": "tr",
    "uk": "gb", "ie": "ie", "br": "br", "au": "au",
    "jp": "ja", "in": "in", "mx": "mx",
}

# ── Thread-safe state ────────────────────────────────────────────────────────
_lock = threading.Lock()
_call_counter = [0]
_key_cycle = itertools.cycle(API_KEYS)
_key_lock = threading.Lock()


def _next_key() -> str:
    with _key_lock:
        return next(_key_cycle)


# ── Domain helpers ───────────────────────────────────────────────────────────

def parse_domain(domain: str):
    """
    Returns (original, com_fallback, country_code).
    country_code is the ISO-2 code detected from the TLD, or None.
    """
    domain = domain.lower().strip().removeprefix("www.")

    if domain.endswith(".co.uk"):
        base = domain[:-len(".co.uk")].split(".")[-1]
        return domain, f"{base}.com", "gb"

    parts = domain.split(".")
    tld = parts[-1]

    if tld in TLD_COUNTRY:
        country = TLD_COUNTRY[tld]
        base = parts[-2]
        return domain, f"{base}.com", country

    if tld == "com" and len(parts) > 2:
        # Subdomain like offer.alibaba.com → alibaba.com
        return domain, ".".join(parts[-2:]), None

    return domain, domain, None


# ── Surfe search ─────────────────────────────────────────────────────────────

def surfe_search(domain: str, country: str = None,
                 job_titles: list = None,
                 limit: int = 5, retries: int = 3) -> list:
    """Call Surfe API and return list of people dicts. Thread-safe."""
    people_filter: dict = {"jobTitles": job_titles if job_titles is not None else JOB_TITLES}
    if country:
        people_filter["countries"] = [country]

    payload = {
        "limit": limit,
        "people": people_filter,
        "companies": {"domains": [domain]},
    }
    key = _next_key()
    with _lock:
        _call_counter[0] += 1
    time.sleep(0.05)  # minimal stagger to avoid simultaneous burst start

    for attempt in range(retries):
        try:
            r = requests.post(
                SURFE_URL,
                json=payload,
                headers={"Authorization": f"Bearer {key}"},
                timeout=20,
            )
            if r.status_code in (429, 403):
                wait = 5 * (attempt + 1)   # 5s, 10s, 15s
                time.sleep(wait)
                continue
            if r.status_code == 200:
                return r.json().get("people", [])
            return []
        except Exception:
            time.sleep(1)
    return []


def title_priority(job_title: str) -> int:
    jt = job_title.lower()
    for i, kw in enumerate(PRIORITY_KEYWORDS):
        if kw in jt:
            return i
    return len(PRIORITY_KEYWORDS)


import re as _re
_EMOJI_RE = _re.compile(
    "[\U0001F000-\U0001FFFF"
    "\U00002702-\U000027B0"
    "\U000024C2-\U0001F251"
    "\U0001F900-\U0001F9FF"
    "☀-⛿✀-➿]+",
    flags=_re.UNICODE,
)

def _clean(name: str) -> str:
    """Strip emojis, zero-width chars and stray punctuation from a name."""
    if not name:
        return ""
    name = _EMOJI_RE.sub("", name)
    # Remove zero-width joiners, variation selectors and other invisible chars
    name = _re.sub(r"[​-‏  ️͏­]", "", name)
    name = _re.sub(r"[✓✔☑✅»«|•·—–]", "", name)
    return _re.sub(r"\s+", " ", name).strip()


def _split_name(first: str, last: str):
    """
    Ensure first and last are populated.
    Surfe sometimes puts emoji (or the whole name) in firstName and
    leaves lastName empty, or vice-versa.
    """
    first = _clean(first)
    last  = _clean(last)

    # If firstName is empty after cleaning, promote from lastName
    if not first and last:
        parts = last.split()
        first = parts[0]
        last  = " ".join(parts[1:]) if len(parts) > 1 else ""

    # If lastName is empty and firstName has multiple tokens, split off last word
    if first and not last and " " in first:
        parts = first.split()
        first = " ".join(parts[:-1])
        last  = parts[-1]

    return first, last


def _filter_by_country(people: list, country: str) -> list:
    """Keep only contacts whose country matches (or is unknown)."""
    return [
        p for p in people
        if not p.get("country") or p.get("country", "").lower() == country
    ]


def _dedup(people: list) -> list:
    """Remove duplicate contacts by LinkedIn URL (keeps first occurrence)."""
    seen, out = set(), []
    for p in people:
        li = p.get("linkedInUrl", "")
        key = li if li else id(p)
        if key not in seen:
            seen.add(key)
            out.append(p)
    return out


def find_contacts(company: dict) -> dict:
    """
    Search Surfe for up to 2 contacts with country-aware priority.

    Strategy: always check BOTH the original domain and the .com fallback
    (when a country TLD is detected), then merge and deduplicate results.

      1. Search original domain, post-filter by country for country-TLD domains.
      2. Also search .com WITH explicit country filter (when applicable).
      3. Merge + dedup by LinkedIn URL, sort by priority, take top 2.

    For .com / no-country domains a single search is made (no fallback).
    """
    domain = company["domain"]
    original, com_fallback, country = parse_domain(domain)

    # Step 1: original domain, 5 candidates, post-filter by country
    people = surfe_search(original, limit=5)
    if country:
        people = _filter_by_country(people, country)

    used_fallback = False

    # Step 2: always also search .com with country filter when applicable
    if country and com_fallback != original:
        more = surfe_search(com_fallback, country=country, limit=5)
        if more:
            used_fallback = True
            # Merge: append contacts not already in people (dedup after)
            people = people + more

    # Dedup (Surfe can return the same person twice across calls)
    people = _dedup(people)
    people.sort(key=lambda p: title_priority(p.get("jobTitle", "")))

    contacts = []
    for p in people[:2]:
        fn, ln = _split_name(p.get("firstName", ""), p.get("lastName", ""))
        surfe_domain = p.get("companyDomain", "")
        rocketreach_domain = surfe_domain if surfe_domain else (
            com_fallback if used_fallback else original
        )
        contacts.append({
            "first_name":         fn,
            "last_name":          ln,
            "linkedin":           p.get("linkedInUrl", ""),
            "title":              p.get("jobTitle", ""),
            "person_country":     p.get("country", ""),
            "rocketreach_domain": rocketreach_domain,
        })

    return {
        **company,
        "contacts":       contacts,
        "used_fallback":  used_fallback,
    }


# ── Excel helpers ────────────────────────────────────────────────────────────

def load_companies(xlsx_path: str, start: int = 0, limit: int = None) -> list:
    """
    Load unique companies from Excel, skipping those already filled
    (First Name present in row 1 of the pair → manually handled).
    """
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    # Collect all rows per company (pairs)
    company_rows: dict[str, list] = {}
    for row_idx in range(2, ws.max_row + 1):
        name = ws.cell(row_idx, 6).value
        if name:
            company_rows.setdefault(name, []).append(row_idx)

    companies = []
    seen = set()
    for row_idx in range(2, ws.max_row + 1):
        name = ws.cell(row_idx, 6).value
        domain = ws.cell(row_idx, 8).value
        if not name or not domain or name in seen:
            continue
        seen.add(name)

        # Skip if first contact row already has a First Name (already processed)
        rows = company_rows.get(name, [])
        if rows and ws.cell(rows[0], 1).value:
            continue

        companies.append({"name": name, "domain": domain})

    # Apply pagination
    companies = companies[start:]
    if limit:
        companies = companies[:limit]
    return companies


def country_to_language(country_code: str) -> str:
    mapping = {
        "it": "it",
        "de": "de", "at": "de", "ch": "de",
        "fr": "fr", "be": "fr",
        "es": "es", "mx": "es",
        "nl": "nl",
        "pl": "pl",
        "pt": "pt", "br": "pt",
        "se": "sv",
        "dk": "da",
        "fi": "fi",
        "no": "no",
        "ro": "ro",
        "cz": "cs",
        "hu": "hu",
        "gr": "el",
    }
    return mapping.get((country_code or "").lower(), "en")


def update_excel(xlsx_path: str, results: list, output_path: str):
    """
    Fill contact rows in the Excel.
    - Respects existing manual entries (skips rows with First Name).
    - Updates Company Domain (col 8) with RocketReach domain when
      different from original, so email lookup uses the right domain.
    - Writes person country (col 9) for Brevo language segmentation.
    """
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    # Map company name → list of row indices
    company_rows: dict[str, list] = {}
    for row_idx in range(2, ws.max_row + 1):
        name = ws.cell(row_idx, 6).value
        if name:
            company_rows.setdefault(name, []).append(row_idx)

    blue   = Font(color="0563C1", underline="single", name="Calibri", size=10)
    normal = Font(name="Calibri", size=10)

    for result in results:
        name     = result["name"]
        contacts = result["contacts"]
        rows     = company_rows.get(name, [])

        for row_idx, contact in zip(rows, contacts):
            # Never overwrite a manually filled row
            if ws.cell(row_idx, 1).value:
                continue

            ws.cell(row_idx, 1).value = contact["first_name"]
            ws.cell(row_idx, 1).font  = normal
            ws.cell(row_idx, 2).value = contact["last_name"]
            ws.cell(row_idx, 2).font  = normal
            ws.cell(row_idx, 7).value = contact["title"]
            ws.cell(row_idx, 9).value = contact["person_country"]

            # Update Company Domain with Surfe's actual domain (for RocketReach)
            rr_domain = contact["rocketreach_domain"]
            if rr_domain:
                ws.cell(row_idx, 8).value = rr_domain
                # Mirror on the paired row (keep both rows consistent)
                paired = [r for r in rows if r != row_idx]
                if paired and not ws.cell(paired[0], 1).value:
                    ws.cell(paired[0], 8).value = rr_domain

            if contact["linkedin"]:
                ws.cell(row_idx, 5).value = contact["linkedin"]
                ws.cell(row_idx, 5).font  = blue

    wb.save(output_path)


# ── Batch runner ─────────────────────────────────────────────────────────────

def run_batch(start: int = 0, limit: int = None,
              output_suffix: str = "batch1",
              input_file: str = None):
    """Process companies without contacts and write results to Excel."""
    xlsx = input_file or "/home/user/AwinContacts/Awin_Contacts.xlsx"
    out  = f"/home/user/AwinContacts/Awin_Contacts_{output_suffix}.xlsx"

    print(f"Loading companies (offset={start}, limit={limit})…", flush=True)
    companies = load_companies(xlsx, start=start, limit=limit)
    print(f"  → {len(companies)} companies to process\n", flush=True)

    results     = []
    hits        = 0
    fallbacks   = 0
    calls_start = _call_counter[0]
    done        = [0]

    def _process(company):
        result = find_contacts(company)
        with _lock:
            done[0] += 1
            found = len(result["contacts"])
            if found:
                print(f"  [{done[0]:>3}/{len(companies)}] HIT  "
                      f"{result['name'][:40]:<40} "
                      f"({_call_counter[0] - calls_start} calls so far)",
                      flush=True)
            elif done[0] % 25 == 0:
                print(f"  [{done[0]:>3}/{len(companies)}] "
                      f"— {_call_counter[0] - calls_start} calls so far",
                      flush=True)
        return result

    with ThreadPoolExecutor(max_workers=20) as ex:  # 20 worker = burst max Surfe, ~2 req/s, limite 10 req/s
        futures = {ex.submit(_process, c): c for c in companies}
        for future in as_completed(futures):
            r = future.result()
            results.append(r)
            if r["contacts"]:
                hits += 1
            if r["used_fallback"]:
                fallbacks += 1

    calls_used = _call_counter[0] - calls_start
    results.sort(key=lambda r: r["name"].lower())

    print(f"\n{'='*60}")
    print(f"BATCH COMPLETATO")
    print(f"  Aziende processate : {len(companies)}")
    print(f"  Hit (≥1 contatto)  : {hits} ({hits/len(companies)*100:.0f}%)")
    print(f"  Fallback .com      : {fallbacks}")
    print(f"  Chiamate API       : {calls_used}")
    print(f"{'='*60}\n")

    print(f"Scrittura Excel → {out}")
    update_excel(xlsx, results, out)
    print("Fatto.")
    return results


if __name__ == "__main__":
    run_batch(
        output_suffix="batch5",
        input_file="/home/user/AwinContacts/Awin_Contacts_complete.xlsx",
    )
