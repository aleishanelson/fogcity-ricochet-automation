"""
One-off fix: clear the bad May 18 rows (9815–9865) then re-upload
with correct SKUs from the Inventory Summary tab.
"""
import os, json, re, logging, base64 as _b64
from datetime import date, timedelta, datetime, timezone
from collections import defaultdict

from playwright.sync_api import sync_playwright
from google.oauth2 import service_account
from googleapiclient.discovery import build

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                    handlers=[logging.StreamHandler()])
log = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────
RICOCHET_URL   = "https://tradingpost.ricoconsign.com"
EMAIL          = os.environ["RICOCHET_EMAIL"]
PASSWORD       = os.environ["RICOCHET_PASSWORD"]
_sa_raw        = os.environ["GOOGLE_SA_JSON"].strip()
try:
    SA_JSON = json.loads(_b64.b64decode(_sa_raw).decode())
except Exception:
    SA_JSON = json.loads(_sa_raw)
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
FOG_CITY_TAB   = "Fog City Sales"

# May 18 only
TARGET_DATE    = date(2026, 5, 18)
DATE_LABEL     = "5/18 - 5/18 ricochet export"
CLEAR_RANGE    = f"'{FOG_CITY_TAB}'!A9815:F9865"   # rows with bad SKUs

PDT            = timezone(timedelta(hours=-7))

# ── Sheets service ────────────────────────────────────────────────────────────
def get_sheets():
    creds = service_account.Credentials.from_service_account_info(
        SA_JSON, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return build("sheets", "v4", credentials=creds).spreadsheets()

# ── SKU lookup ────────────────────────────────────────────────────────────────
def build_sku_lookup(sheets):
    result = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'Inventory Summary'!B:E",
    ).execute()
    lookup = {}
    for row in result.get("values", [])[1:]:
        if len(row) >= 4:
            name = str(row[0]).strip()
            sku  = str(row[3]).strip()
            if name and sku:
                lookup[name.lower()] = sku
    log.info(f"SKU lookup: {len(lookup)} entries")
    return lookup

def find_sku(name, lookup):
    key = name.strip().lower()
    if key in lookup:
        return lookup[key]
    for k, v in lookup.items():
        if key in k or k in key:
            return v
    return ""

# ── Normalize (copy from main script) ────────────────────────────────────────
def normalize(name):
    nl = name.strip().lower()
    if re.search(r'sticker.{0,15}11|3.{0,5}for.{0,5}\$?11.{0,10}sticker', nl): return 'Stickers- 3 for $11'
    if re.search(r'postcard.{0,15}11|3.{0,5}for.{0,5}\$?11.{0,10}postcard', nl): return 'Postcards 3 for $11'
    if re.search(r'retro.{0,8}(gg|golden gate).{0,20}(magnet|magent)', nl): return 'Retro GG Travel Poster magnet'
    if re.search(r'(gg|golden gate).{0,10}travel.{0,10}(poster.{0,5})?magnet', nl): return 'Golden Gate Travel Poster Magnet'
    if re.search(r'(acrylic|die.?cut).{0,20}(golden gate|gg).{0,20}magnet|(golden gate|gg).{0,20}(acrylic|die.?cut).{0,20}magnet', nl): return 'Golden Gate Acrylic Die Cut Magnet'
    if re.search(r'(acrylic|die.?cut).{0,20}ferry building|ferry building.{0,20}(acrylic|die.?cut)', nl): return 'Ferry Building Acrylic Die Cut Magnet'
    if re.search(r'(acrylic|die.?cut).{0,20}fisherm|fisherm.{0,20}(acrylic|die.?cut)', nl): return 'Fishermans Wharf Acrylic Die Cut Magnet'
    if re.search(r'(acrylic|die.?cut).{0,20}painted lad|painted lad.{0,15}(acrylic|die.?cut)', nl): return 'ACRYLIC Painted Ladies Magnet'
    if re.search(r'sf (illustrated|aleisha).{0,10}(landmark.{0,5})?magnet', nl): return 'SF Illustrated Landmarks Magnet'
    if re.search(r'ferry building.{0,10}travel.{0,10}magnet', nl): return 'Ferry Building Travel Poster Magnet'
    if re.search(r'retro.{0,10}painted lad.{0,20}(poster.{0,5})?magnet', nl): return 'Retro Painted Ladies Poster Magnet'
    if re.search(r'retro.{0,8}ferry building.{0,20}(poster|retro).{0,10}magnet', nl): return 'Retro Ferry Building Poster Magnet'
    if re.search(r'take the scenic route magnet|scenic route magnet', nl): return 'Take The Scenic Route Magnet'
    if re.search(r'home sweet sf magnet', nl): return 'Home Sweet SF Magnet'
    if re.search(r'(sf|san francisco).{0,10}(pink.{0,5})?icons? magnet', nl): return 'SF Pink Icons Magnet'
    if re.search(r'city by the bay.{0,10}(local notion.{0,5})?magnet', nl): return 'City by the Bay Local Notion Magnet'
    if re.search(r'pink city by the bay.{0,10}(circle.{0,5})?magnet', nl): return 'Pink City By The Bay Circle Magnet'
    if re.search(r'cable car.{0,10}magnet|yellow.{0,5}cable.?car', nl): return 'Yellow Cable Car Magnet'
    if re.search(r'blue sf waves? magnet|sfwave magnet', nl): return 'Blue SF Waves Magnet'
    if re.search(r'san francisco.{0,10}(golden gate.{0,10})?retro postcard|sfgg', nl): return 'San Francisco Golden Gate Bridge Retro Postcard'
    if re.search(r'ferry building retro postcard', nl): return 'Ferry Building Retro Postcard'
    if re.search(r'fisherm.{0,10}(wharf.{0,5})?retro postcard', nl): return "Fisherman's Wharf Retro Postcard"
    if re.search(r'lombard.{0,10}(street.{0,5})?retro postcard', nl): return 'Lombard Street Retro Postcard'
    if re.search(r'painted ladies retro postcard', nl): return 'Painted Ladies Retro Postcard'
    if re.search(r'alcatraz.{0,10}retro postcard', nl): return 'Alcatraz Island Retro Postcard'
    if re.search(r'blue.{0,10}icons? postcard|sf blue icons postcard', nl): return 'SF Blue Icons Postcard'
    if re.search(r'pink.{0,10}icons? postcard', nl): return 'Pink Icons Postcard'
    if re.search(r'west coast best coast sticker', nl): return 'West Coast Best Coast sticker'
    if re.search(r'i come with baggage', nl): return 'I Come With Baggage Sticker'
    if re.search(r'sunny.{0,10}75', nl): return 'Sunny and 75 in SF Card'
    if re.search(r'(sf|san francisco).{0,10}(landmark.{0,5})?sticker sheet', nl): return 'San Francisco Landmark Sticker Sheet'
    if re.search(r'uc (berkeley|santa barbara).{0,10}tea towel|uc berkeley tea', nl): return 'UC Berkeley Tea Towels'
    if re.search(r'ucla.{0,10}tea towel', nl): return 'UCLA Tea Towels'
    if re.search(r'lake tahoe.{0,10}tea towel', nl): return 'Lake Tahoe Tea Towel'
    if re.search(r'sfo.{0,10}(luggage.{0,5})?keychain', nl): return 'SFO Luggage Tag Keychain'
    if re.search(r'acrylic keychain|sf.{0,5}(icon.{0,5})?keychain', nl): return 'Acrylic keychain'
    if re.search(r'home sweet (sf|san francisco)', nl):
        if 'tote' in nl: return 'Home Sweet San Francisco Tote'
        if 'sticker' in nl: return 'Home Sweet Home Sticker'
        return 'Home Sweet San Francisco Art Print 8x8'
    if re.search(r'(sf|san francisco).{0,10}city name.{0,10}sticker', nl): return 'SF City Name With Golden Gate Sticker'
    if re.search(r'bon voyage sticker', nl): return 'Bon Voyage Sticker'
    if re.search(r'proud tourist sticker', nl): return 'Proud Tourist sticker'
    return name.strip()

# ── Scrape Ricochet ───────────────────────────────────────────────────────────
def scrape_ricochet():
    log.info("Scraping Ricochet…")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page    = browser.new_page()
        page.goto(f"{RICOCHET_URL}/login", wait_until="networkidle")
        page.fill('input[type="text"], input[name="username"]', EMAIL)
        page.fill('input[type="password"]', PASSWORD)
        page.click('button[type="submit"], button:has-text("Login")')
        page.wait_for_url(f"{RICOCHET_URL}/dashboard", timeout=15_000)
        page.click('text=Payout')
        page.wait_for_selector('table tr', timeout=10_000)
        page.wait_for_load_state("networkidle")
        rows = page.evaluate("""() => {
            const rows = document.querySelectorAll('table tr');
            return Array.from(rows).map(r =>
                Array.from(r.querySelectorAll('td,th')).map(c => c.innerText.trim())
            ).filter(r => r.length > 0);
        }""")
        browser.close()

    headers = [h.lower().replace("/","_").replace(" ","_") for h in rows[0]]
    records = [dict(zip(headers, r)) for r in rows[1:] if len(r) == len(headers)]
    log.info(f"Scraped {len(records)} records")
    return records

# ── Filter to May 18 ──────────────────────────────────────────────────────────
def filter_may18(records):
    out = []
    for r in records:
        try:
            d = datetime.strptime(r.get("sold",""), "%m-%d-%Y").date()
        except ValueError:
            continue
        if d == TARGET_DATE:
            out.append(r)
    log.info(f"May 18 records: {len(out)}")
    return out

# ── Merge ─────────────────────────────────────────────────────────────────────
def merge_rows(records):
    groups = defaultdict(lambda: {"item":"","sku":"","qty":0})
    for r in records:
        raw = r.get("item","").strip()
        can = normalize(raw)
        key = can.lower()
        groups[key]["item"] = can
        groups[key]["qty"] += 1
        if not groups[key]["sku"]:
            groups[key]["sku"] = r.get("sku","").strip()
    return sorted(groups.values(), key=lambda x: x["item"].lower())

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    log.info("=== fix_may18 starting ===")
    sheets     = get_sheets()
    sku_lookup = build_sku_lookup(sheets)

    # Step 1: Clear old May 18 rows
    log.info(f"Clearing {CLEAR_RANGE}…")
    sheets.values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=CLEAR_RANGE,
        body={}
    ).execute()
    log.info("Cleared.")

    # Step 2: Scrape and filter to May 18
    raw       = scrape_ricochet()
    filtered  = filter_may18(raw)
    if not filtered:
        log.warning("No May 18 sales found — nothing to write.")
        return

    merged = merge_rows(filtered)
    log.info(f"Merged: {len(merged)} items, {sum(m['qty'] for m in merged)} units")

    # Step 3: Write fresh rows starting at 9815
    new_values = [
        ["May", 2026, m["item"],
         find_sku(m["item"], sku_lookup) or m["sku"],
         m["qty"], DATE_LABEL]
        for m in merged
    ]

    write_range = f"'{FOG_CITY_TAB}'!A9815:F{9815 + len(new_values) - 1}"
    sheets.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=write_range,
        valueInputOption="USER_ENTERED",
        body={"values": new_values},
    ).execute()

    log.info(f"✅ Wrote {len(new_values)} rows to {write_range}")
    for m in new_values:
        log.info(f"   {m[2][:40]:40s}  SKU={m[3]}  qty={m[4]}")
    log.info("=== Done ===")

if __name__ == "__main__":
    main()
