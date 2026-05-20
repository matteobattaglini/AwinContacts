#!/usr/bin/env python3
"""
Surfe Contact Finder for Awin Advertisers
Finds Affiliate/Digital Marketing contacts for each company.
"""

import requests
import json
import time
import itertools
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import re

# ── API keys ────────────────────────────────────────────────────────────────
API_KEYS = [
    "TL1fKLZMEE7rdWEKaIe_W-vmUvxkyP0fecMTLsbFJKA",
    "rMY_QurAxUvl6jDeeM9xN01q5IMFBfrH5AUJJLdafe4",
    "KwGtwE1QeCfHU1G3NjkMyGFqf2PT36NEC5taWt0iFgY",
]
SURFE_URL = "https://api.surfe.com/v2/people/search"

# ── Job title priorities ─────────────────────────────────────────────────────
JOB_TITLES = [
    "Affiliate Marketing",
    "Digital Marketing",
    "Performance Marketing",
]

# Keywords to rank results client-side (order = priority)
PRIORITY_KEYWORDS = [
    "affiliate",
    "digital marketing",
    "performance marketing",
    "online marketing",
    "marketing",
]

# ── Country TLD → ISO code ───────────────────────────────────────────────────
TLD_COUNTRY = {
    "it": "it", "de": "de", "fr": "fr", "es": "es",
    "nl": "nl", "pl": "pl", "pt": "pt", "be": "be",
    "se": "se", "dk": "dk", "fi": "fi", "no": "no",
    "ch": "ch", "at": "at", "ro": "ro", "cz": "cz",
    "hu": "hu", "gr": "gr", "ru": "ru", "tr": "tr",
    "uk": "gb", "ie": "ie", "br": "br", "au": "au",
    "jp": "ja", "in": "in", "mx": "mx",
}

lock = threading.Lock()
call_counter = [0]
# Round-robin key distributor
key_cycle = itertools.cycle(API_KEYS)
key_lock = threading.Lock()


def next_key():
    with key_lock:
        return next(key_cycle)


def parse_domain(domain: str):
    """
    Returns (original, com_fallback, country_code).
    country_code is None for generic TLDs (.com, .net, .eu, .org).
    """
    domain = domain.lower().strip().lstrip("www.")

    # Handle .co.uk
    if domain.endswith(".co.uk"):
        base = domain[: -len(".co.uk")].split(".")[-1]
        return domain, f"{base}.com", "gb"

    parts = domain.split(".")
    tld = parts[-1]

    if tld in TLD_COUNTRY:
        country = TLD_COUNTRY[tld]
        # Base = second-to-last part (handles subdomains like conti.credit-agricole.it)
        base = parts[-2]
        com_fallback = f"{base}.com"
        return domain, com_fallback, country

    if tld == "com" and len(parts) > 2:
        # Subdomain: offer.alibaba.com → alibaba.com
        root = ".".join(parts[-2:])
        return domain, root, None

    return domain, domain, None


def title_priority(job_title: str) -> int:
    """Lower = higher priority."""
    jt = job_title.lower()
    for i, kw in enumerate(PRIORITY_KEYWORDS):
        if kw in jt:
            return i
    return len(PRIORITY_KEYWORDS)


def surfe_search(domain: str, limit: int = 10, retries: int = 3) -> list:
    """Single Surfe API call. Returns list of person dicts."""
    payload = {
        "limit": limit,
        "people": {"jobTitles": JOB_TITLES},
        "companies": {"domains": [domain]},
    }

    key = next_key()
    with lock:
        call_counter[0] += 1

    for attempt in range(retries):
        try:
            r = requests.post(
                SURFE_URL,
                json=payload,
                headers={"Authorization": f"Bearer {key}"},
                timeout=20,
            )
            if r.status_code == 429:
                wait = 2 ** attempt
                time.sleep(wait)
                continue
            if r.status_code != 200:
                return []
            return r.json().get("people", [])
        except Exception:
            time.sleep(1)
    return []


def find_contacts(company: dict) -> dict:
    """
    Find up to 2 prioritised contacts for a company.
    Strategy:
      1. Search with original domain
      2. If 0 results and domain has country TLD → try .com (no country filter,
         since affiliate teams are often global even for local brands)
    """
    domain = company["domain"]
    original, com_fallback, country = parse_domain(domain)

    people = surfe_search(original)

    used_fallback = False
    if not people and country and com_fallback != original:
        people = surfe_search(com_fallback)   # no country filter
        used_fallback = True

    # Sort by priority
    people.sort(key=lambda p: title_priority(p.get("jobTitle", "")))

    contacts = []
    for p in people[:2]:
        contacts.append({
            "first_name": p.get("firstName", ""),
            "last_name": p.get("lastName", ""),
            "linkedin": p.get("linkedInUrl", ""),
            "title": p.get("jobTitle", ""),
            "country": p.get("country", ""),
        })

    return {
        **company,
        "contacts": contacts,
        "used_fallback": used_fallback,
        "resolved_domain": com_fallback if used_fallback else original,
    }


def load_companies(xlsx_path: str, limit: int = None) -> list:
    """Load unique companies from Excel (2-row pairs)."""
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    seen = set()
    companies = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = row[5]
        domain = row[7]
        if name and domain and name not in seen:
            seen.add(name)
            companies.append({"name": name, "domain": domain})
        if limit and len(companies) >= limit:
            break
    return companies


def update_excel(xlsx_path: str, results: list, output_path: str):
    """Write contacts into the Excel, filling the 2-row slots per company."""
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    # Build lookup: company_name → row indices (1-based) of its two rows
    company_rows = {}
    for row_idx in range(2, ws.max_row + 1):
        name = ws.cell(row_idx, 6).value
        if name:
            if name not in company_rows:
                company_rows[name] = []
            company_rows[name].append(row_idx)

    blue = Font(color="0563C1", underline="single", name="Calibri", size=10)
    normal = Font(name="Calibri", size=10)

    for result in results:
        name = result["name"]
        contacts = result["contacts"]
        rows = company_rows.get(name, [])

        for slot, (row_idx, contact) in enumerate(zip(rows, contacts)):
            # Skip if row already has a First Name (manually filled)
            if ws.cell(row_idx, 1).value:
                continue

            ws.cell(row_idx, 1).value = contact["first_name"]
            ws.cell(row_idx, 1).font = normal
            ws.cell(row_idx, 2).value = contact["last_name"]
            ws.cell(row_idx, 2).font = normal
            ws.cell(row_idx, 7).value = contact["title"]    # Title column
            ws.cell(row_idx, 9).value = contact["country"]  # Country column

            # LinkedIn as hyperlink
            linkedin = contact["linkedin"]
            if linkedin:
                ws.cell(row_idx, 5).value = linkedin
                ws.cell(row_idx, 5).font = blue

    wb.save(output_path)


def run_test(limit: int = 100):
    """Test run on first N companies."""
    xlsx = "/home/user/AwinContacts/Awin_Contacts.xlsx"
    print(f"Loading first {limit} companies from Excel...")
    companies = load_companies(xlsx, limit=limit)
    print(f"Loaded {len(companies)} companies\n")

    results = []
    hits = 0
    fallbacks = 0
    calls_start = call_counter[0]

    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = {executor.submit(find_contacts, c): c for c in companies}
        for future in as_completed(futures):
            result = future.result()
            results.append(result)
            found = len(result["contacts"])
            if found > 0:
                hits += 1
            if result["used_fallback"]:
                fallbacks += 1

    calls_used = call_counter[0] - calls_start
    results.sort(key=lambda r: r["name"].lower())

    print(f"\n{'='*65}")
    print(f"RISULTATI TEST — prime {limit} aziende")
    print(f"{'='*65}")
    print(f"Hit (≥1 contatto trovato):  {hits}/{len(companies)} = {hits/len(companies)*100:.0f}%")
    print(f"Fallback .com usati:        {fallbacks}")
    print(f"Chiamate API totali:        {calls_used}")
    print(f"{'='*65}\n")

    print(f"{'AZIENDA':<35} {'CONTATTO 1':<30} {'TITOLO':<35} {'FB'}")
    print("-" * 110)
    for r in results:
        c1 = r["contacts"][0] if r["contacts"] else None
        c2 = r["contacts"][1] if len(r["contacts"]) > 1 else None
        fb = "↩.com" if r["used_fallback"] else ""
        if c1:
            print(f"  {r['name'][:33]:<33} {(c1['first_name']+' '+c1['last_name'])[:28]:<30} {c1['title'][:33]:<35} {fb}")
            if c2:
                print(f"  {'':33} {(c2['first_name']+' '+c2['last_name'])[:28]:<30} {c2['title'][:33]:<35}")
        else:
            print(f"  {r['name'][:33]:<33} {'—':<30} {'nessun risultato':<35} {fb}")

    # Write to Excel
    out = "/home/user/AwinContacts/Awin_Contacts_test100.xlsx"
    update_excel(xlsx, results, out)
    print(f"\nFile aggiornato: {out}")
    return results


if __name__ == "__main__":
    run_test(limit=100)
