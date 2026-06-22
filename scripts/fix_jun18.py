"""
One-off fix: clear the bad June 18 rows (10784-10810) and re-upload
with correct SKUs using the OVERRIDES from the updated ricochet_sync.py.
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
RICOCHET_URL  = "https://tradingpost.ricoconsign.com"
EMAIL         = os.environ["RICOCHET_EMAIL"]
PASSWORD      = os.environ["RICOCHET_PASSWORD"]
_sa_raw       = os.environ["GOOGLE_SA_JSON"].strip()
try:
    SA_JSON = json.loads(_b64.b64decode(_sa_raw).decode())
except Exception:
    SA_JSON = json.loads(_sa_raw)
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
FOG_CITY_TAB   = "Fog City Sales"

TARGET_DATE  = date(2026, 6, 18)
DATE_LABEL   = "6/18 - 6/18 ricochet export"
CLEAR_RANGE  = f"'{FOG_CITY_TAB}'!A10784:G10820"  # rows with bad SKUs
WRITE_START  = 10784

PDT = timezone(timedelta(hours=-7))

# ── Sheets service ────────────────────────────────────────────────────────────
def get_sheets():
    creds = service_account.Credentials.from_service_account_info(
        SA_JSON, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return build("sheets", "v4", credentials=creds).spreadsheets()

# ── SKU lookup from Inventory Summary ────────────────────────────────────────
def build_sku_lookup(sheets):
    result = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'Inventory Summary'!A:E",
    ).execute()
    all_lookup = {}
    by_category = {}
    for row in result.get("values", [])[1:]:
        if len(row) < 5:
            continue
        category  = str(row[0]).strip()
        item_name = str(row[1]).strip()
        sku       = str(row[4]).strip()
        if not item_name or not sku:
            continue
        key = item_name.lower()
        all_lookup[key] = sku
        if category:
            by_category.setdefault(category, {})[key] = sku
    log.info(f"SKU lookup: {len(all_lookup)} entries across {len(by_category)} categories")
    return {"all": all_lookup, "by_category": by_category}

def find_sku(item_name: str, lookup: dict) -> str:
    key = item_name.strip().lower()
    all_lookup  = lookup.get("all", lookup) if isinstance(lookup, dict) else lookup
    by_category = lookup.get("by_category", {}) if isinstance(lookup, dict) else {}

    def _match(d, k):
        if k in d: return d[k]
        for lname, sku in d.items():
            if k in lname or lname in k: return sku
        return ""

    OVERRIDES = {
        # ── 6/18/2026 fix items ───────────────────────────────────────────────
        "acrylic keychain":                                  "KC-SF-01",
        "sf acrylic keychain":                               "KC-SF-01",
        "golden gate icon keychain (icon series)":           "KC-SF-01",
        "acrylic painted ladies magnet":                     "MAG-SF-PLADIES-CL",
        "bay area map print 8x10":                           "BAYAREA_BW_8x10",
        "bay area map print 9x12":                           "BAYAREA_BW_9x12",
        "bay area map print 11x14":                          "BAYAREA_BW_11x14",
        "bay area map print 12x16":                          "BAYAREA_BW_12x16",
        "california state sticker (black and white)":        "CALIFORNIASTATE_BLOCKFONT_BW",
        "california state sticker black and white":          "CALIFORNIASTATE_BLOCKFONT_BW",
        "chicago map print 8x10":                            "CHICAGO_BW_8x10",
        "chicago map print 9x12":                            "CHICAGO_BW_9x12",
        "chicago map print 11x14":                           "CHICAGO_BW_11x14",
        "ferry building acrylic die cut magnet":             "MAG-AC-SF-FERRYB",
        "ferry building retro postcard":                     "FERRYBUILDING_RETRO_PC",
        "golden gate acrylic die cut magnet":                "MAG-AC-SF-GGB",
        "golden gate bridge sticker (pink)":                 "GGBRIDGE_PINK_STICKER",
        "golden gate travel poster magnet":                  "MAGNET_GOLDENGATETRAVELPOSTER",
        "home sweet home sticker":                           "HOMESWEETSANFRANCISCO_STICKER",
        "magnet set san francisco":                          "SANFRANCISCOICONS_MAGNETSET",
        "postcards 3 for $11":                               "postcards3for11",
        "retro ferry building poster magnet":                "MAG-SF-RETRO-FB",
        "retro golden gate bridge poster sticker":           "RETRO_GGB_STICKER",
        "retro painted ladies poster magnet":                "MAG-SF-RETRO-PL",
        "san francisco golden gate bridge retro postcard":   "SFGGBRIDGE_RETRO_PCARD",
        "santa clara university campus map print 8x10":      "SCU_BW_8x10_CURSIVE",
        "santa clara university campus map print":           "SCU_BW_8x10_CURSIVE",
        "sf city by the bay dad hat - cream":                "DH-CITYBYTHEBAY-CREAM",
        "sf city by the bay dad hat":                        "DH-CITYBYTHEBAY-CREAM",
        "sf city name with golden gate sticker":             "SFCITYNAME_STICKER",
        # ── Pre-existing overrides carried from fix_may18.py ─────────────────
        "alcatraz island postcard":                          "ALCATRAZ_RETRO_PC",
        "fishermans wharf acrylic die cut magnet":           "MAG-AC-SF-FW",
        "fishermans wharf sf postcard":                      "FISHERMANSWHARF_RETRO_PC",
        "fisherman's wharf sticker":                         "FW_ILLUSTRATED_STICKER",
        "golden gate bridge acrylic":                        "MAG-AC-SF-GGB",
        "home sweet sf magnet":                              "MAGNET_HOMESWEETSF",
        "retro gg travel poster magnet":                     "MAG-SF-RETRO-GGB",
        "santa clara university campus map print 8x10":      "SCU_BW_8x10_CURSIVE",
        "sf icon tote":                                      "SFICONS_TOTE",
        "sf landmark magnet":                                "MAG-SF-LDMKS",
        "twist and turn card":                               "TWISTSANDTURNS_GCARD",
        "window seat card":                                  "WINDOWSEAT_A2_GREETINGCARD",
        "sunny and 75 in sf card":                           "LOVEYOUMORETHANSUNNYSF_A2CARD",
        "stickers- 3 for $11":                               "STICKERS_3FOR11",
        "stickers 3 for $11":                                "STICKERS_3FOR11",
        "3 for 11":                                          "STICKERS_3FOR11",
    }

    if key in OVERRIDES: return OVERRIDES[key]
    for k, v in OVERRIDES.items():
        if key in k or k in key: return v

    # Category-filtered search
    CATEGORY_HINTS = {
        "magnet": {"Magnets"},
        "sticker": {"Stickers","Sticker","Sticker deal","Sticker Sheet","Sticker Book"},
        "tote": {"Totes"},
        "card": {"Greeting Card","Card Pack"},
        "print": {"City Print","School Prints","Film Print","Landmark"},
        "keychain": {"Keychains","Keychain"},
        "postcard": {"Postcards","Postcard deal"},
        "tea towel": {"Tea Towel"},
        "towel": {"Tea Towel"},
    }
    for keyword, categories in CATEGORY_HINTS.items():
        if keyword in key:
            cat_pool = {}
            for cat in categories:
                cat_pool.update(by_category.get(cat, {}))
            if cat_pool:
                r = _match(cat_pool, key)
                if r: return r
            break

    return _match(all_lookup, key)

# ── Normalize (mirrors ricochet_sync.py) ─────────────────────────────────────
def normalize(name: str) -> str:
    nl = name.strip().lower()
    if re.search(r'sticker.{0,15}11|3.{0,5}for.{0,5}\$?11.{0,10}sticker', nl): return 'Stickers- 3 for $11'
    if re.search(r'postcard.{0,15}11|3.{0,5}for.{0,5}\$?11.{0,10}postcard', nl): return 'Postcards 3 for $11'
    if re.search(r'retro.{0,8}(gg|golden.?gate).{0,20}(magnet|magent)', nl): return 'Retro GG Travel Poster magnet'
    if re.search(r'(gg|golden.?gate).{0,10}travel.{0,10}(poster.{0,5})?magnet', nl): return 'Golden Gate Travel Poster Magnet'
    if re.search(r'(acrylic|die.?cut).{0,20}(golden gate|gg).{0,20}magnet|(golden gate|gg).{0,20}(acrylic|die.?cut).{0,20}magnet', nl): return 'Golden Gate Acrylic Die Cut Magnet'
    if re.search(r'(acrylic|die.?cut).{0,20}ferry building|ferry building.{0,20}(acrylic|die.?cut)', nl): return 'Ferry Building Acrylic Die Cut Magnet'
    if re.search(r'(acrylic|die.?cut).{0,20}fisherm|fisherm.{0,20}(acrylic|die.?cut)', nl): return 'Fishermans Wharf Acrylic Die Cut Magnet'
    if re.search(r'(acrylic|die.?cut).{0,20}painted lad|painted lad.{0,15}(acrylic|die.?cut)', nl): return 'ACRYLIC Painted Ladies Magnet'
    if re.search(r'(gg|golden.?gate).{0,15}travel.{0,5}(poster.{0,5})?card', nl): return 'Golden Gate Travel Poster Card'
    if re.search(r'(gg|golden.?gate).{0,20}travel.{0,20}poster.{0,5}8x10', nl): return 'Golden Gate Travel Poster 8x10'
    if re.search(r'(gg|golden.?gate).{0,20}travel.{0,20}poster.{0,5}11x14', nl): return 'Golden Gate Travel Poster 11x14'
    if re.search(r'(gg|golden.?gate).{0,20}travel.{0,20}poster.{0,5}12x16', nl): return 'Golden Gate Travel Poster 12x16'
    if re.search(r'ferry building.{0,30}(8x10|9x12|11x14|12x16)', nl): return 'Ferry Building Travel Poster Print'
    if re.search(r'(retro.{0,10})?ferry building.{0,20}(poster.{0,5})?sticker', nl): return 'Retro Ferry Building Poster Sticker'
    if re.search(r'ferry building.{0,20}travel.{0,10}(poster.{0,5})?magnet', nl): return 'Ferry Building Travel Poster Magnet'
    if re.search(r'(retro.{0,8})?ferry building.{0,20}(poster|retro).{0,10}magnet', nl): return 'Retro Ferry Building Poster Magnet'
    if re.search(r'retro.{0,10}painted lad.{0,20}(poster.{0,5})?magnet', nl): return 'Retro Painted Ladies Poster Magnet'
    if re.search(r'retro.{0,8}(gg|golden.?gate).{0,20}(poster.{0,5})?sticker', nl): return 'Retro Golden Gate Bridge Poster Sticker'
    if re.search(r'(gg|golden.?gate).{0,10}travel.{0,10}(poster.{0,5})?sticker', nl): return 'Golden Gate Travel Poster Sticker'
    if re.search(r'san francisco.{0,10}(golden gate.{0,10})?retro postcard|sfgg', nl): return 'San Francisco Golden Gate Bridge Retro Postcard'
    if re.search(r'ferry building retro postcard', nl): return 'Ferry Building Retro Postcard'
    if re.search(r'fisherm.{0,10}(wharf.{0,5})?retro postcard', nl): return "Fisherman's Wharf Retro Postcard"
    if re.search(r'lombard.{0,10}(street.{0,5})?retro postcard', nl): return 'Lombard Street Retro Postcard'
    if re.search(r'painted ladies retro postcard', nl): return 'Painted Ladies Retro Postcard'
    if re.search(r'alcatraz.{0,10}retro postcard', nl): return 'Alcatraz Island Retro Postcard'
    if re.search(r'blue.{0,10}icons? postcard|sf blue icons postcard', nl): return 'SF Blue Icons Postcard'
    if re.search(r'pink.{0,10}icons? postcard|sf icons postcard', nl): return 'Pink Icons Postcard'
    if re.search(r'sf.{0,5}(map|icons?|city name).{0,5}tote|(san francisco|sf).{0,10}tote', nl):
        if 'map' in nl: return 'SF Map tote'
        if 'icon' in nl: return 'SF Icons Tote Bag'
        if 'city' in nl: return 'San Francisco city tote'
        return 'SF Icons Tote Bag'
    if re.search(r'(sf|san francisco).{0,10}(landmark.{0,5})?sticker sheet', nl): return 'San Francisco Landmark Sticker Sheet'
    if re.search(r'sunny.{0,10}75', nl): return 'Sunny and 75 in SF Card'
    if re.search(r"i'?d escape alcatraz", nl): return "I'd Escape Alcatraz Card"
    if re.search(r'window seat', nl): return "I'd Give Up My Window Seat card"
    if re.search(r'twi?st.{0,8}turn', nl): return 'Thru twists and turns greeting card'
    if re.search(r'sweet.{0,15}(birthday|bday)|wishing.{0,10}sweet', nl): return 'Wishing You A Sweet Birthday Cake Card'
    if re.search(r'sfo.{0,10}(luggage.{0,5})?keychain', nl): return 'SFO Luggage Tag Keychain'
    if re.search(r'acrylic keychain|sf.{0,5}(icon.{0,5})?keychain|golden gate icon keychain', nl): return 'Acrylic keychain'
    if re.search(r'home sweet (sf|san francisco|home)', nl):
        if 'magnet' in nl: return 'Home Sweet SF Magnet'
        if 'sticker' in nl: return 'Home Sweet Home Sticker'
        if 'tote' in nl: return 'Home Sweet San Francisco Tote'
        if 'card' in nl: return 'Home Sweet SF Greeting Card'
        return 'Home Sweet San Francisco Art Print 8x8'
    if re.search(r'(sf|san francisco).{0,10}city name.{0,10}(with golden gate.{0,5})?sticker', nl): return 'SF City Name With Golden Gate Sticker'
    if re.search(r'(sf|san francisco).{0,10}map sticker', nl): return 'SF Map Sticker'
    if re.search(r'west coast best coast sticker', nl): return 'West Coast Best Coast sticker'
    if re.search(r'(ca|california) state sticker.{0,10}(black|b.?w)', nl): return 'California State Sticker (Black and White)'
    if re.search(r'(ca|california) state sticker.{0,10}blue', nl): return 'California State Sticker (Blue)'
    if re.search(r'golden gate bridge sticker', nl): return 'Golden Gate Bridge Sticker (pink)'
    if re.search(r'(sf|san francisco).{0,10}(icons?.{0,5})?magnet set|magnet set.{0,10}(sf|san francisco)', nl): return 'Magnet set San Francisco'
    if re.search(r'(sf|san francisco).{0,10}block font magnet', nl): return 'SF Block Font Magnet'
    if re.search(r'city by the bay.{0,10}(local notion.{0,5})?magnet', nl): return 'City by the Bay Local Notion Magnet'
    if re.search(r'pink city by the bay.{0,10}(circle.{0,5})?magnet', nl): return 'Pink City By The Bay Circle Magnet'
    if re.search(r'cable car.{0,10}magnet|yellow.{0,5}cable.?car', nl): return 'Yellow Cable Car Magnet'
    if re.search(r'take the scenic route magnet', nl): return 'Take The Scenic Route Magnet'
    if re.search(r'(sf|san francisco).{0,10}(pink.{0,5})?icons? magnet', nl): return 'SF Pink Icons Magnet'
    if re.search(r'(sf.{0,5})?city by the bay.{0,10}(dad.{0,5})?hat|city by the bay hat', nl): return 'SF City By The Bay Dad Hat - Cream'
    if re.search(r'sf fog dad hat', nl): return 'SF Fog Dad Hat - Light Blue'
    if re.search(r'bay area.{0,10}(map.{0,5})?tea towel', nl): return 'The Bay Area tea towel'
    if re.search(r'(san francisco|sf).{0,10}(map.{0,5})?tea towel', nl): return 'San Francisco Tea Towel'
    if re.search(r'california (state|map).{0,10}tea towel|state of california tea towel|california tea towel', nl): return 'California Map tea towel'
    if re.search(r'uc (berkeley|santa barbara).{0,10}tea towel', nl): return 'UC Berkeley Tea Towels'
    if re.search(r'ucla.{0,10}tea towel', nl): return 'UCLA Tea Towels'
    if re.search(r'lake tahoe.{0,10}tea towel', nl): return 'Lake Tahoe Tea Towel'
    if re.search(r'santa clara university.{0,10}(campus.{0,5})?map print', nl): return 'Santa Clara University CAMPUS Map Print 8x10'
    return name.strip()

# ── Scrape Ricochet ───────────────────────────────────────────────────────────
def scrape_ricochet():
    log.info("Scraping Ricochet...")
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
    log.info(f"Headers: {headers}")
    records = [dict(zip(headers, r)) for r in rows[1:] if len(r) == len(headers)]
    log.info(f"Scraped {len(records)} records")
    return records

# ── Filter to Jun 18 ─────────────────────────────────────────────────────────
def filter_jun18(records):
    out = []
    for r in records:
        try:
            d = datetime.strptime(r.get("sold", ""), "%m-%d-%Y").date()
        except ValueError:
            continue
        if d == TARGET_DATE:
            out.append(r)
    log.info(f"Jun 18 records: {len(out)}")
    return out

# ── Merge ─────────────────────────────────────────────────────────────────────
def merge_rows(records):
    groups = defaultdict(lambda: {"item": "", "sku": "", "qty": 0, "revenue": 0.0})
    for r in records:
        raw = r.get("item", "").strip()
        can = normalize(raw)
        key = can.lower()
        groups[key]["item"] = can
        groups[key]["qty"] += 1
        if not groups[key]["sku"]:
            groups[key]["sku"] = r.get("sku", "").strip()
        try:
            price_str = r.get("aged_price", "0").replace("$","").replace(",","").strip()
            groups[key]["revenue"] += float(price_str) if price_str else 0.0
        except (ValueError, AttributeError):
            pass
    return sorted(groups.values(), key=lambda x: x["item"].lower())

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    log.info("=== fix_jun18 starting ===")
    sheets     = get_sheets()
    sku_lookup = build_sku_lookup(sheets)

    # Step 1: Clear old bad rows
    log.info(f"Clearing {CLEAR_RANGE}...")
    sheets.values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=CLEAR_RANGE,
        body={}
    ).execute()
    log.info("Cleared.")

    # Step 2: Scrape and filter to Jun 18
    raw      = scrape_ricochet()
    filtered = filter_jun18(raw)
    if not filtered:
        log.warning("No Jun 18 sales found — nothing to write.")
        return

    merged = merge_rows(filtered)
    log.info(f"Merged: {len(merged)} items, {sum(m['qty'] for m in merged)} units")

    # Step 3: Write fresh rows starting at WRITE_START
    new_values = [
        ["June", 2026, m["item"],
         find_sku(m["item"], sku_lookup) or m["sku"],
         m["qty"], DATE_LABEL, round(m.get("revenue", 0.0), 2)]
        for m in merged
    ]

    write_range = f"'{FOG_CITY_TAB}'!A{WRITE_START}:G{WRITE_START + len(new_values) - 1}"
    sheets.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=write_range,
        valueInputOption="USER_ENTERED",
        body={"values": new_values},
    ).execute()

    log.info(f"Wrote {len(new_values)} rows to {write_range}")
    for m in new_values:
        log.info(f"  {m[2][:45]:45s} SKU={m[3]} qty={m[4]} rev={m[6]}")
    log.info("=== Done ===")

if __name__ == "__main__":
    main()
