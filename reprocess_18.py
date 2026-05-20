#!/usr/bin/env python3
"""
Targeted reprocessing for the 18 companies that got non-Italian contacts.
Strategy: always try original domain first, then force country='it' on .com fallback.
Reads from Awin_Contacts_batch1_v1.xlsx (the current best file),
writes corrected results to Awin_Contacts_batch1_v2.xlsx.
"""

import sys
import os

# Reuse everything from contact_finder
sys.path.insert(0, os.path.dirname(__file__))
from contact_finder import (
    parse_domain, surfe_search, find_contacts,
    title_priority, _split_name, _lock, _call_counter,
    update_excel, _next_key,
)
import openpyxl
import time
from openpyxl.styles import Font
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

# The 18 companies that had non-Italian contacts in v1
TARGET_COMPANIES = {
    "3DMakerpro (EU)",
    "Acer IT",
    "AliExpress IT",
    "Allies of Skin EU",
    "Avon IT",
    "Beatbot IT",
    "Blizzard Gear Store IT",
    "Costa Crociere IT",
    "Decathlon IT",
    "Decathlon Rent IT",
    "Delonghi IT",
    "Enjoy the Wood",
    "Europcar IT",
    "Fiverr Affiliates",
    "Flightnetwork",
    "Folletto IT",
    "Goldcar",
    "Honor IT",
}

INPUT_FILE  = "/home/user/AwinContacts/Awin_Contacts_batch1_v1.xlsx"
OUTPUT_FILE = "/home/user/AwinContacts/Awin_Contacts_batch1_v2.xlsx"


def find_contacts_forced_it(company: dict) -> dict:
    """
    Like find_contacts() but always uses country='it' as the fallback filter.
    For companies whose original domain returns 0 results, this ensures we
    search for Italian-based people at the .com domain.
    """
    domain = company["domain"]
    original, com_fallback, country = parse_domain(domain)

    # Force Italian market targeting for these programs
    target_country = country if country else "it"

    people = surfe_search(original)

    used_fallback = False
    if not people:
        # Try .com variant with Italian country filter
        if com_fallback != original:
            people = surfe_search(com_fallback, country=target_country)
            used_fallback = True
        elif not people:
            # Domain is already .com-like — retry same domain with country filter
            people = surfe_search(original, country=target_country)
            used_fallback = True

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
        "contacts":      contacts,
        "used_fallback": used_fallback,
    }


def load_target_companies(xlsx_path: str) -> list:
    """Load only the 18 target companies from the Excel file."""
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    companies = []
    seen = set()
    for row_idx in range(2, ws.max_row + 1):
        name   = ws.cell(row_idx, 6).value
        domain = ws.cell(row_idx, 8).value
        if not name or not domain or name in seen:
            continue
        seen.add(name)
        if name in TARGET_COMPANIES:
            companies.append({"name": name, "domain": domain})

    return companies


def clear_contacts_in_excel(xlsx_path: str, output_path: str, target_names: set):
    """
    Clear existing contact data for target companies so update_excel can refill them.
    Clears cols 1,2,5,7,9 for rows belonging to these companies.
    """
    import shutil
    shutil.copy2(xlsx_path, output_path)

    wb = openpyxl.load_workbook(output_path)
    ws = wb.active

    for row_idx in range(2, ws.max_row + 1):
        name = ws.cell(row_idx, 6).value
        if name in target_names:
            ws.cell(row_idx, 1).value = None  # First Name
            ws.cell(row_idx, 2).value = None  # Last Name
            ws.cell(row_idx, 5).value = None  # LinkedIn
            ws.cell(row_idx, 7).value = None  # Title
            ws.cell(row_idx, 9).value = None  # Country

    wb.save(output_path)
    print(f"Cleared contact slots for {len(target_names)} companies in {output_path}")


def main():
    print(f"Targeted reprocessing: {len(TARGET_COMPANIES)} companies", flush=True)
    print(f"Input:  {INPUT_FILE}", flush=True)
    print(f"Output: {OUTPUT_FILE}\n", flush=True)

    companies = load_target_companies(INPUT_FILE)
    found_names = {c["name"] for c in companies}
    missing = TARGET_COMPANIES - found_names
    if missing:
        print(f"WARNING: not found in Excel: {missing}", flush=True)
    print(f"Found {len(companies)} companies to process:\n", flush=True)
    for c in companies:
        print(f"  {c['name']:<45} {c['domain']}", flush=True)
    print(flush=True)

    # Step 1: copy v1 → v2 and clear the target rows
    clear_contacts_in_excel(INPUT_FILE, OUTPUT_FILE, TARGET_COMPANIES)

    # Step 2: fetch contacts
    results     = []
    hits        = 0
    calls_start = _call_counter[0]
    done        = [0]
    lock_local  = threading.Lock()

    def _process(company):
        result = find_contacts_forced_it(company)
        with lock_local:
            done[0] += 1
            found = len(result["contacts"])
            status = "HIT " if found else "MISS"
            print(f"  [{done[0]:>2}/{len(companies)}] {status} "
                  f"{result['name']:<45} "
                  f"→ {result['contacts'][0]['first_name'] + ' ' + result['contacts'][0]['last_name'] if found else result['domain']} "
                  f"[{result['contacts'][0].get('person_country','?') if found else ''}]",
                  flush=True)
        return result

    with ThreadPoolExecutor(max_workers=2) as ex:
        futures = {ex.submit(_process, c): c for c in companies}
        for future in as_completed(futures):
            r = future.result()
            results.append(r)
            if r["contacts"]:
                hits += 1

    calls_used = _call_counter[0] - calls_start

    print(f"\n{'='*60}", flush=True)
    print(f"Completato", flush=True)
    print(f"  Aziende processate : {len(companies)}", flush=True)
    print(f"  Hit (≥1 contatto)  : {hits}/{len(companies)}", flush=True)
    print(f"  Chiamate API       : {calls_used}", flush=True)
    print(f"{'='*60}\n", flush=True)

    print(f"Scrittura → {OUTPUT_FILE}", flush=True)
    update_excel(OUTPUT_FILE, results, OUTPUT_FILE)
    print("Fatto.", flush=True)


if __name__ == "__main__":
    main()
