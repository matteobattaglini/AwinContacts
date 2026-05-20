#!/usr/bin/env python3
"""
Awin Advertiser Directory Scraper
Fetches all advertisers from awin.com IT directory and their public profile data.
"""

import requests
import json
import time
import re
from urllib.parse import urlparse
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

ALGOLIA_APP_ID = "C1W1Y0AAMV"
ALGOLIA_API_KEY = "e2f2e353dd25e86d66044683f3711e6a"
ALGOLIA_INDEX = "awin-website-automatic_advertiser-directory_it"
ALGOLIA_URL = "https://c1w1y0aamv-dsn.algolia.net/1/indexes/*/queries"

PROFILE_BASE = "https://ui.awin.com/merchant-profile/{}"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9,en-US;q=0.8",
}

lock = threading.Lock()
progress_counter = [0]


def get_all_advertisers():
    """Fetch all advertisers from Algolia in pages."""
    params = {
        "x-algolia-agent": "Algolia for JavaScript (5.52.1); Browser",
        "x-algolia-api-key": ALGOLIA_API_KEY,
        "x-algolia-application-id": ALGOLIA_APP_ID,
    }

    # First request to get total pages
    payload = {
        "requests": [{
            "indexName": ALGOLIA_INDEX,
            "params": "query=&hitsPerPage=50&page=0&distinct=true",
        }]
    }
    r = requests.post(ALGOLIA_URL, params=params, json=payload, timeout=30)
    data = r.json()
    result = data["results"][0]
    nb_pages = result["nbPages"]
    nb_hits = result["nbHits"]
    print(f"Found {nb_hits} advertisers across {nb_pages} pages")

    all_hits = list(result["hits"])

    # Fetch remaining pages
    for page in range(1, nb_pages):
        payload = {
            "requests": [{
                "indexName": ALGOLIA_INDEX,
                "params": f"query=&hitsPerPage=50&page={page}&distinct=true",
            }]
        }
        r = requests.post(ALGOLIA_URL, params=params, json=payload, timeout=30)
        hits = r.json()["results"][0]["hits"]
        all_hits.extend(hits)
        print(f"  Page {page + 1}/{nb_pages}: {len(all_hits)} advertisers fetched")
        time.sleep(0.1)

    return all_hits


def clean_url(url):
    """Fix malformed URLs like 'http://facebook.com/https://...'"""
    if not url:
        return url
    m = re.search(r"(https?://(?!.*https?://).+)", url)
    if m:
        return m.group(1)
    return url


def get_profile_data(merchant_id):
    """Fetch public profile page and extract website, Facebook, cookie period."""
    url = PROFILE_BASE.format(merchant_id)
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code != 200:
            return {}
        soup = BeautifulSoup(r.text, "html.parser")

        # Cookie / attribution period
        cookie_period = ""
        pag_h2 = soup.find("h2", class_="margin-top-lg")
        if pag_h2 and "Pagamenti" in pag_h2.get_text():
            p_el = pag_h2.find_next("p", class_="margin-bottom")
            if p_el:
                cookie_period = p_el.get_text(strip=True)

        # Website and social links
        website = ""
        facebook = ""
        instagram = ""
        twitter = ""
        links_div = soup.find("div", class_="list-group list-group-icons")
        if links_div:
            for a in links_div.find_all("a", href=True):
                href = a["href"]
                icon = a.find("i")
                icon_class = " ".join(icon.get("class", [])) if icon else ""
                if "fa-globe" in icon_class:
                    website = href
                elif "fa-facebook" in icon_class:
                    facebook = href
                elif "fa-instagram" in icon_class:
                    instagram = href
                elif "fa-twitter" in icon_class or "fa-x-twitter" in icon_class:
                    twitter = href
                elif not website and href.startswith("http"):
                    website = href

        return {
            "website": clean_url(website),
            "facebook": clean_url(facebook),
            "instagram": clean_url(instagram),
            "twitter": clean_url(twitter),
            "cookie_period": cookie_period,
        }

    except Exception as e:
        return {}


def extract_domain(url):
    """Extract clean domain from URL."""
    if not url:
        return ""
    try:
        url = clean_url(url)
        parsed = urlparse(url if url.startswith("http") else "https://" + url)
        domain = parsed.netloc.lower()
        domain = re.sub(r"^www\.", "", domain)
        return domain
    except Exception:
        return url


def fetch_advertiser(hit):
    """Fetch profile data for a single advertiser."""
    merchant_id = hit["id"]
    profile_data = get_profile_data(merchant_id)

    website = profile_data.get("website", "")
    domain = extract_domain(website)

    sectors = []
    for s in hit.get("sectors", []):
        if s.get("level", 0) == 1:
            sectors.append(s.get("sectorName", ""))
    if not sectors:
        for s in hit.get("sectors", []):
            name = s.get("sectorName", "")
            if name:
                sectors.append(name)
                break

    with lock:
        progress_counter[0] += 1
        if progress_counter[0] % 50 == 0 or progress_counter[0] <= 5:
            print(f"  [{progress_counter[0]}] {hit['name'][:50]} -> {domain}")

    return {
        "id": merchant_id,
        "name": hit.get("name", ""),
        "description": hit.get("description", ""),
        "sector": ", ".join(sectors),
        "join_date": hit.get("joinDate", "")[:10] if hit.get("joinDate") else "",
        "website": website,
        "domain": domain,
        "facebook": profile_data.get("facebook", ""),
        "instagram": profile_data.get("instagram", ""),
        "twitter": profile_data.get("twitter", ""),
        "cookie_period": profile_data.get("cookie_period", ""),
        "profile_url": f"https://ui.awin.com/merchant-profile/{merchant_id}",
    }


def build_excel(advertisers, output_path):
    """Create Excel file with 2 rows per advertiser."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Awin Contacts"

    # Header style
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=10, name="Calibri")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin = Side(border_style="thin", color="CCCCCC")
    cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Columns
    headers = [
        "First Name",
        "Last Name",
        "Work Email",
        "Personal Email",
        "LinkedIn URL",
        "Company",
        "Title",
        "Company Domain",
        "Country",
        "Commissioni (Cookie Period)",
        "Settore",
        "Descrizione",
        "Sito Web",
        "Facebook",
        "Instagram",
        "Twitter / X",
        "Awin Profile",
        "Data Iscrizione Awin",
    ]

    ws.row_dimensions[1].height = 35
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = cell_border

    # Alternating row colors
    fill_a = PatternFill(start_color="EBF3FB", end_color="EBF3FB", fill_type="solid")
    fill_b = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    data_font = Font(name="Calibri", size=10)
    data_align = Alignment(vertical="center", wrap_text=False)

    row = 2
    for i, adv in enumerate(advertisers):
        fill = fill_a if i % 2 == 0 else fill_b

        # Row 1 of the pair: company info + full detail columns
        row1_data = [
            "",                         # First Name
            "",                         # Last Name
            "",                         # Work Email
            "",                         # Personal Email
            "",                         # LinkedIn URL
            adv["name"],               # Company
            "",                         # Title
            adv["domain"],             # Company Domain
            "",                         # Country
            adv["cookie_period"],      # Commissioni
            adv["sector"],             # Settore
            adv["description"],        # Descrizione
            adv["website"],            # Sito Web
            adv["facebook"],           # Facebook
            adv["instagram"],          # Instagram
            adv["twitter"],            # Twitter
            adv["profile_url"],        # Awin Profile
            adv["join_date"],          # Data Iscrizione
        ]

        # Row 2 of the pair: contact placeholder (name + domain only)
        row2_data = [
            "",                         # First Name
            "",                         # Last Name
            "",                         # Work Email
            "",                         # Personal Email
            "",                         # LinkedIn URL
            adv["name"],               # Company (repeated)
            "",                         # Title
            adv["domain"],             # Company Domain (repeated)
            "",                         # Country
            "",                         # Commissioni (blank)
            "",                         # Settore (blank)
            "",                         # Descrizione (blank)
            "",                         # Sito Web (blank)
            "",                         # Facebook (blank)
            "",                         # Instagram (blank)
            "",                         # Twitter (blank)
            "",                         # Awin Profile (blank)
            "",                         # Data (blank)
        ]

        for col_idx, value in enumerate(row1_data, 1):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.fill = fill
            cell.font = data_font
            cell.alignment = data_align
            cell.border = cell_border

        for col_idx, value in enumerate(row2_data, 1):
            cell = ws.cell(row=row + 1, column=col_idx, value=value)
            cell.fill = fill
            cell.font = data_font
            cell.alignment = data_align
            cell.border = cell_border

        row += 2

    # Column widths
    col_widths = {
        1: 15,   # First Name
        2: 15,   # Last Name
        3: 28,   # Work Email
        4: 28,   # Personal Email
        5: 35,   # LinkedIn
        6: 35,   # Company
        7: 20,   # Title
        8: 30,   # Company Domain
        9: 12,   # Country
        10: 22,  # Commissioni
        11: 25,  # Settore
        12: 55,  # Descrizione
        13: 40,  # Sito Web
        14: 35,  # Facebook
        15: 35,  # Instagram
        16: 35,  # Twitter
        17: 40,  # Awin Profile
        18: 18,  # Data Iscrizione
    }
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Freeze header row
    ws.freeze_panes = "A2"

    wb.save(output_path)
    print(f"\nSaved: {output_path}")
    print(f"Rows: {row - 1} ({len(advertisers)} advertisers × 2)")


def main():
    print("=== Awin Advertiser Directory Scraper ===\n")

    print("Step 1: Fetching all advertisers from Algolia...")
    hits = get_all_advertisers()
    print(f"Total: {len(hits)} advertisers\n")

    print("Step 2: Fetching profile pages (concurrent)...")
    advertisers = []

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(fetch_advertiser, hit): hit for hit in hits}
        for future in as_completed(futures):
            result = future.result()
            if result:
                advertisers.append(result)
            time.sleep(0.02)

    # Sort by name
    advertisers.sort(key=lambda x: x["name"].lower())

    print(f"\nStep 3: Building Excel with {len(advertisers)} advertisers...")
    output_path = "/home/user/AwinContacts/Awin_Contacts.xlsx"
    build_excel(advertisers, output_path)

    print("\nDone!")


if __name__ == "__main__":
    main()
