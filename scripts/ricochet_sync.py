"""
ricochet_sync.py
Daily job: scrape Ricochet payout data → merge/deduplicate → write to Google Sheet.
Runs via GitHub Actions at 11 pm PDT every day.

Date logic:
  1. Read the last populated cell in column F of the Fog City Sales tab.
  2. Parse the end date from that label (e.g. "5/1 - 5/17 ricochet export" → May 17).
  3. Set the upload window to: (that date + 1 day) through yesterday.
  4. If column F is empty or unparseable, fall back to the 1st of the current month.
  5. If the sheet is already current (last end date == yesterday), exit with no changes.

Required environment variables (stored as GitHub Secrets):
  RICOCHET_EMAIL      artbyaleisha@gmail.com
  RICOCHET_PASSWORD   (Ricochet password)
  GOOGLE_SA_JSON      Full JSON of a Google service-account key with Sheets editor access
  SPREADSHEET_ID      1S-gUNDVezLyAA30yUk1_u22vzNb5YYV0cCeDMOh_nBY
"""

import os, json, re, logging
from datetime import datetime, timedelta, date, timezone
from collections import defaultdict

from playwright.sync_api import sync_playwright
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.StreamHandler(), logging.FileHandler("sync.log")],
)
log = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────
RICOCHET_URL   = "https://tradingpost.ricoconsign.com"
EMAIL          = os.environ["RICOCHET_EMAIL"]
PASSWORD       = os.environ["RICOCHET_PASSWORD"]
import base64 as _b64
_sa_raw        = os.environ["GOOGLE_SA_JSON"].strip()
# Secret is stored as base64 to avoid newline escaping issues in GitHub Actions.
try:
    SA_JSON = json.loads(_b64.b64decode(_sa_raw).decode())
except Exception:
    SA_JSON = json.loads(_sa_raw)
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
FOG_CITY_TAB   = "Fog City Sales"
SOURCE_LABEL   = "ricochet export"

PDT       = timezone(timedelta(hours=-7))
TODAY     = datetime.now(PDT).date()
YESTERDAY = TODAY - timedelta(days=1)


# ── Step 1: Connect to Sheets ─────────────────────────────────────────────────

def get_sheets_service():
    creds = service_account.Credentials.from_service_account_info(
        SA_JSON,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds).spreadsheets()


# ── Step 1.5: Build SKU lookup from Inventory Summary tab ────────────────────

def build_sku_lookup(sheets) -> dict:
    """
    Read Inventory Summary tab cols B (item name) and E (SKU).
    Return a dict of lowercase item name → SKU so we can replace
    Ricochet's SKUs with the correct ones from your spreadsheet.
    """
    result = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'Inventory Summary'!B:E",
    ).execute()

    rows = result.get("values", [])
    lookup = {}
    for row in rows[1:]:   # skip header
        if len(row) < 4:
            continue
        item_name = str(row[0]).strip()
        sku       = str(row[3]).strip() if len(row) > 3 else ""
        if item_name and sku:
            lookup[item_name.lower()] = sku

    log.info(f"SKU lookup loaded: {len(lookup)} entries from Inventory Summary")
    return lookup


def find_sku(item_name: str, lookup: dict) -> str:
    """
    Try to find a matching SKU from the lookup table.
    Tries exact match first, then partial substring match.
    Falls back to hardcoded overrides for items not in Inventory Summary.
    Returns empty string if no match found.
    """
    key = item_name.strip().lower()

    # Exact match against Inventory Summary
    if key in lookup:
        return lookup[key]

    # Partial match — item name contains or is contained in a lookup key
    for lookup_name, sku in lookup.items():
        if key in lookup_name or lookup_name in key:
            return sku

    # Hardcoded overrides: items missing from or mismatched in Inventory Summary
    OVERRIDES = {
        # Postcards
        "alcatraz island postcard":                      "ALCATRAZ_RETRO_PC",
        "alcatraz island retro postcard":                "ALCATRAZ_RETRO_PC",
        "fishermans wharf sf postcard":                  "FISHERMANSWHARF_RETRO_PC",
        "fisherman's wharf retro postcard":              "FISHERMANSWHARF_RETRO_PC",
        "ferry building retro postcard":                 "FERRYBUILDING_RETRO_PC",
        "lombard street retro postcard":                 "LOMBARDST_RETRO_PC",
        "painted ladies retro postcard":                 "PAINTEDLADIES_RETRO_PC",
        "pink icons postcard":                           "SFICONS_POSTCARD_PINK",
        "art by aleisha postcards - pink":               "SFICONS_POSTCARD_PINK",
        "sf blue icons postcard":                        "SFICONS_BLUE_4x6",
        "blue icons postcard":                           "SFICONS_BLUE_4x6",
        "golden gate bridge retro postcard":             "SFGGBRIDGE_RETRO_PCARD",
        "san francisco golden gate bridge retro postcard": "SFGGBRIDGE_RETRO_PCARD",
        "postcards 3 for $11":                           "postcards3for11",
        "postcards- 3 for $11":                          "postcards3for11",
        # Magnets
        # ACRYLIC die-cut magnets — separate product line from flat magnets
        # canonical names match what normalize() returns
        "fishermans wharf acrylic die cut magnet":       "MAG-AC-SF-FW",
        "acrylic fishermans wharf":                      "MAG-AC-SF-FW",
        "acrylic fisherman's wharf magnet":              "MAG-AC-SF-FW",
        "acrylic fishermans wharf magnet":               "MAG-AC-SF-FW",
        "ferry building acrylic die cut magnet":         "MAG-AC-SF-FERRYB",
        "acrylic ferry building magnet":                 "MAG-AC-SF-FERRYB",
        "acrylic ferry building":                        "MAG-AC-SF-FERRYB",
        "golden gate acrylic die cut magnet":            "MAG-AC-SF-GGB",
        "acrylic golden gate bridge magnet":             "MAG-AC-SF-GGB",
        "acrylic golden gate magnet":                    "MAG-AC-SF-GGB",
        "acrylic gg magnet":                             "MAG-AC-SF-GGB",
        "golden gate bridge acrylic die cut magnet":     "MAG-AC-SF-GGB",
        "gg acrylic die cut magnet":                     "MAG-AC-SF-GGB",
        "golden gate bridge acrylic":                    "MAG-AC-SF-GGB",  # NOT a print
        "acrylic painted ladies magnet":                 "MAG-SF-PLADIES-CL",
        "acrylic painted ladies":                        "MAG-SF-PLADIES-CL",
        "sf houses acrylic die cut magnet":              "MAG-AC-SF-HOUSES",
        "acrylic sf houses magnet":                      "MAG-AC-SF-HOUSES",
        "sf house acrylic magnet":                       "MAG-AC-SF-HOUSES",
        "sfo luggage tag acrylic magnet":                "MAG-AC-SF-SFO",
        "acrylic sfo luggage tag":                       "MAG-AC-SF-SFO",
        "home sweet sf magnet":                          "MAGNET_HOMESWEETSF",
        "home sweet home magnet":                        "MAGNET_HOMESWEETSF",
        "retro gg travel poster magnet":                 "MAG-SF-RETRO-GGB",
        "retrogg bridge travel poster magent":           "MAG-SF-RETRO-GGB",
        "retro golden gate bridge poster magnet":        "MAG-SF-RETRO-GGB",
        "retro golden gate poster magnet":               "MAG-SF-RETRO-GGB",
        "retro ferry building poster magnet":            "MAG-SF-RETRO-FB",
        "retro painted ladies poster magnet":            "MAG-SF-RETRO-PL",
        "ferry building travel poster magnet":           "FERRYBUILDINGTRAVELPOSTER_MAGNET",
        "ferry building":                                "FERRYBUILDINGTRAVELPOSTER_MAGNET",
        "sf illustrated landmarks magnet":               "MAG-SF-LDMKS",
        "sf landmark magnet":                            "MAG-SF-LDMKS",
        "sf illustrated landmark":                       "MAG-SF-LDMKS",
        "sf pink icons magnet":                          "MAG-SF-PINKICONS",
        "sf block font magnet":                          "MAG-SF-BLOCKFONT",
        "cable car magnet":                              "MAG-SF-CABLECAR",
        "yellow cable car magnet":                       "MAG-SF-CABLECAR",
        "take the scenic route magnet":                  "TAKETHESCENICROUTE_49MILE_MAGNET",
        "sfo luggage tag magnet":                        "SFO_LUGGAGETAG_MAGNET",
        "vintage sfo luggage tag magnet":                "SFO_LUGGAGETAG_MAGNET",
        "pink city by the bay circle magnet":            "MAGNET_SFCITYBYTHEBAY_PINK",
        "city by the bay local notion magnet":           "MAGNET_SFCITYBYTHEBAY_LOCALNOTION",
        # Stickers
        "stickers 3 for $11":                           "STICKERS_3FOR11",
        "stickers- 3 for $11":                          "STICKERS_3FOR11",
        "3 for 11":                                     "STICKERS_3FOR11",
        "golden gate bridge sticker (pink)":             "GGBRIDGE_PINK_STICKER",
        "golden gate sticker (pink)":                    "GGBRIDGE_PINK_STICKER",
        "gg bridge sticker (pink)":                      "GGBRIDGE_PINK_STICKER",
        "golden gate travel sticker":                    "GOLDENGATETRAVELPOSTER_STICKER",
        "golden gate bridge travel sticker":             "GOLDENGATETRAVELPOSTER_STICKER",
        "gg travel sticker":                             "GOLDENGATETRAVELPOSTER_STICKER",
        "retro golden gate bridge poster sticker":       "RETRO_GGB_STICKER",
        "retro golden gate poster sticker":              "RETRO_GGB_STICKER",
        "retro gg poster sticker":                       "RETRO_GGB_STICKER",
        "retro sf ferry building poster sticker":        "RETROSFFERRYBUILDING_STICKER",
        "illustrated fisherman's wharf sticker":         "FW_ILLUSTRATED_STICKER",
        "fisherman's wharf sticker":                     "FW_ILLUSTRATED_STICKER",
        "illustrated golden gate bridge sticker":        "GGB_ILLUSTRATED_STICKER",
        "illustrated golden gate sticker":               "GGB_ILLUSTRATED_STICKER",
        "illustrated gg sticker":                        "GGB_ILLUSTRATED_STICKER",
        "illustrated ferry building sticker":            "FB_ILLUSTRATED_STICKER",
        "sf landmark sticker sheet":                     "SS-SF-LDMKS",
        "san francisco landmark sticker sheet":          "SS-SF-LDMKS",
        "sf map sticker":                                "SFMAP_STICKER",
        "sf city name sticker":                          "SFCITYNAME_STICKER",
        "sf pennant sticker":                            "SFPENNANT_STICKER",
        "home sweet home sticker":                       "HOMESWEETSANFRANCISCO_STICKER",
        "west coast best coast sticker":                 "WESTCOASTBESTCOAST_CIRCLE_STICKER",
        "proud tourist sticker":                         "PROUDTOURIST_STICKER",
        "bon voyage sticker":                            "BONVOYAGE_STICKER",
        "i come with baggage sticker":                   "ICOMEWITHBAGGAGE_STICKER",
        "pink sf city by the bay circle sticker":        "SFCITYBYTHEBAY_PINKCIRCLE_STICKER",
        # Totes
        "sf icons tote":                                 "SFICONS_TOTE",
        "sf icon tote":                                  "SFICONS_TOTE",
        "san francisco icons tote":                      "SFICONS_TOTE",
        "sf city name tote":                             "SF_BLOCKFONT_TOTE",
        "home sweet home tote":                          "TOTE_HOMESWEETSF",
        "sf map tote":                                   "SF_MAP_TOTE",
        # Keychains
        "sfo luggage tag keychain":                      "KC-SFO-LUGGAGETAG",
        "acrylic keychain":                              "KC-SFO-LUGGAGETAG",
        # Cards
        "sunny and 75 in sf card":                       "LOVEYOUMORETHANSUNNYSF_A2CARD",
        "sunny and 75 sf card":                          "LOVEYOUMORETHANSUNNYSF_A2CARD",
        "i love you more than a":                        "LOVEYOUMORETHANSUNNYSF_A2CARD",
        "twist and turn card":                           "TWISTSANDTURNS_GCARD",
        "thru twists and turns card":                    "TWISTSANDTURNS_GCARD",
        "twists and turns greeting card":                "TWISTSANDTURNS_GCARD",
        "window seat card":                              "WINDOWSEAT_A2_GREETINGCARD",
        "art by aleisha cards":                          "WINDOWSEAT_A2_GREETINGCARD",  # flag
        # Pencil Pouches
        "sf icons pencil pouch - natural":               "pp-sf-cn-02",
        "sf icons pencil pouch - blue":                  "PP-SF-CB-01",
        "pencil pouches - sf icons pencil pouch - blue": "PP-SF-CB-01",
        # Tea Towels
        "uc berkeley tea towels":                        "TT-CAMPUS-BERKELEY",
        "uc berkeley tea towel":                         "TT-CAMPUS-BERKELEY",
        "uc santa barbara tea towel":                    "TT-CAMPUS-SANTABARBARA",
        "ucla tea towels":                               "TT-CAMPUS-UCLA",
        "lake tahoe tea towel":                          "TT-LAKETAHOE",
        "cal poly slo tea towel":                        "TT-CAMPUS-CALPOLYSLO",
        "san francisco map tea towel":                   "SANFRANCISCO_MAP_DISHTOWEL",
        "bay area tea towel":                            "BAYAREA_MAP_TEATOWEL",
        "state of california tea towel":                 "STATEOFCALIFORNIA_MAP_TEATOWEL",
        "santa clara university tea towel":              "TT-CAMPUS-SANTACLARA",
        # Prints (common mismatches seen in Fog City Sales)
        "home sweet san francisco art print 8x8":        "HOMESWEETSF_8x8",
        "home sweet home 8x8":                           "HOMESWEETSF_8x8",
        "home sweet sf art print 8x8":                   "HOMESWEETSF_8x8",
        "home sweet san francisco art print 8x10":       "HOMESWEETSF_8x10",
        "home sweet home 8x10":                          "HOMESWEETSF_8x10",
        "home sweet sf art print 8x10":                  "HOMESWEETSF_8x10",
        "home sweet san francisco art print 11x15":      "HOMESWEETSF_11x15",
        "home sweet home 11x15":                         "HOMESWEETSF_11x15",
        "home sweet sf art print 11x15":                 "HOMESWEETSF_11x15",
        "home sweet sf magnet":                          "MAGNET_HOMESWEETSF",
        "home sweet home magnet":                        "MAGNET_HOMESWEETSF",
        "home sweet sf sticker":                         "HOMESWEETSANFRANCISCO_STICKER",
        "home sweet home sticker":                       "HOMESWEETSANFRANCISCO_STICKER",
        "home sweet sf tote":                            "TOTE_HOMESWEETSF",
        "home sweet home tote":                          "TOTE_HOMESWEETSF",
        "golden gate travel poster - 8x10":              "GOLDENGATETRAVELPOSTER_8x10",
        "golden gate bridge travel poster - 8x10":       "GOLDENGATETRAVELPOSTER_8x10",
        "gg travel poster - 8x10":                       "GOLDENGATETRAVELPOSTER_8x10",
        "golden gate travel poster - 11x14":             "GOLDENGATETRAVELPOSTER_11x14",
        "golden gate bridge travel poster - 11x14":      "GOLDENGATETRAVELPOSTER_11x14",
        "gg travel poster - 11x14":                      "GOLDENGATETRAVELPOSTER_11x14",
        "ferry building travel poster - 8x10":           "FERRYBUILDINGTRAVELPOSTER_8x10",
        "ferry building travel poster - 11x14":          "FERRYBUILDINGTRAVELPOSTER_11x14",
        "santa clara university campus map print 8x10":  "SCU_BW_8x10_CURSIVE",
        "stanford campus map print 8x10":               "STANFORD_BW_8x10",
    }
    if key in OVERRIDES:
        return OVERRIDES[key]

    # Partial match against overrides
    for override_name, sku in OVERRIDES.items():
        if key in override_name or override_name in key:
            return sku

    return ""


# ── Step 2: Determine upload window from last column F entry ──────────────────

def get_window_start(sheets) -> date:
    """
    Read all values in column F of Fog City Sales.
    Walk backwards to find the last non-empty cell.
    Parse the END date from its label, e.g.:
        "5/1 - 5/17 ricochet export"  →  May 17  →  window starts May 18
        "5/17 ricochet export"         →  May 17  →  window starts May 18
        "5/1-5/17 ricochet export"     →  May 17  →  window starts May 18

    Falls back to the 1st of YESTERDAY's month if column F is empty or unparseable.
    """
    result = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{FOG_CITY_TAB}'!F:F",
    ).execute()

    col_f = result.get("values", [])

    # Find the last non-empty cell walking backwards
    last_label = None
    for row in reversed(col_f):
        if row and str(row[0]).strip():
            last_label = str(row[0]).strip()
            break

    if not last_label:
        fallback = YESTERDAY.replace(day=1)
        log.warning(f"Column F is empty — falling back to start of month: {fallback}")
        return fallback

    log.info(f"Last column F value: '{last_label}'")

    # Extract all date tokens of the form M/D or M/D/YYYY
    date_tokens = re.findall(r'\b(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\b', last_label)

    if not date_tokens:
        fallback = YESTERDAY.replace(day=1)
        log.warning(f"No date found in '{last_label}' — falling back to {fallback}")
        return fallback

    # The LAST date token is the end of the range
    month_s, day_s, year_s = date_tokens[-1]
    month = int(month_s)
    day   = int(day_s)

    if year_s:
        year = int(year_s)
        if year < 100:
            year += 2000
    else:
        # Infer year: use current year; if candidate is in the future, step back one year
        year      = YESTERDAY.year
        candidate = date(year, month, day)
        if candidate > YESTERDAY:
            year -= 1

    last_upload_end = date(year, month, day)
    window_start    = last_upload_end + timedelta(days=1)

    log.info(f"Last upload ended {last_upload_end} → new window starts {window_start}")
    return window_start


# ── Step 3: Scrape Ricochet payout table ──────────────────────────────────────

def scrape_ricochet() -> list[dict]:
    """Log in to Ricochet and return all payout rows as a list of dicts."""
    log.info("Launching headless browser to scrape Ricochet…")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page    = browser.new_page()

        # Login
        page.goto(f"{RICOCHET_URL}/login", wait_until="networkidle")
        page.fill('input[type="text"], input[name="username"], input[placeholder*="sername"]', EMAIL)
        page.fill('input[type="password"]', PASSWORD)
        page.click('button[type="submit"], button:has-text("Login")')
        page.wait_for_url(f"{RICOCHET_URL}/dashboard", timeout=15_000)
        log.info("Logged in successfully.")

        # Go to Payout tab
        page.click('text=Payout')
        page.wait_for_selector('table tr', timeout=10_000)
        page.wait_for_load_state("networkidle")

        rows = page.evaluate("""
            () => {
                const rows = document.querySelectorAll('table tr');
                const data = [];
                rows.forEach(row => {
                    const cells = row.querySelectorAll('td, th');
                    const r = Array.from(cells).map(c => c.innerText.trim());
                    if (r.length > 0) data.push(r);
                });
                return data;
            }
        """)
        browser.close()

    if not rows:
        raise RuntimeError("No rows returned from Ricochet payout table.")

    # Header: Item, SKU, Sale, Agreed, Cost/Split, Aged Price, Discounts, Sold, Amount
    headers = [h.lower().replace("/", "_").replace(" ", "_") for h in rows[0]]
    log.info(f"Headers: {headers}  |  Total rows (incl. header): {len(rows)}")

    records = [dict(zip(headers, row)) for row in rows[1:] if len(row) == len(headers)]
    log.info(f"Raw records: {len(records)}")
    return records


# ── Step 4: Filter to the computed date window ────────────────────────────────

def filter_by_date(records: list[dict], window_start: date, window_end: date) -> list[dict]:
    """Keep only rows where 'sold' date is within [window_start, window_end]."""
    filtered = []
    for r in records:
        try:
            d = datetime.strptime(r.get("sold", ""), "%m-%d-%Y").date()
        except ValueError:
            continue
        if window_start <= d <= window_end:
            r["_date"] = d
            filtered.append(r)
    log.info(f"After date filter ({window_start} – {window_end}): {len(filtered)} rows")
    return filtered


# ── Step 5: Normalize item names ──────────────────────────────────────────────

def normalize(name: str) -> str:
    nl = name.strip().lower()

    # 3-for-$11 bundles
    if re.search(r'sticker.{0,15}11|11.{0,15}sticker|3.{0,5}for.{0,5}\$?11.{0,10}sticker|sticker.{0,10}3.{0,5}for', nl):
        return 'Stickers- 3 for $11'
    if re.search(r'postcard.{0,15}11|11.{0,15}postcard|3.{0,5}for.{0,5}\$?11.{0,10}postcard|postcard.{0,10}3.{0,5}for', nl):
        return 'Postcards 3 for $11'

    # Retro GG / GG Travel Poster magnets
    if re.search(r'retro.{0,8}(gg|golden.?gate).{0,20}(magnet|magent)', nl):
        return 'Retro GG Travel Poster magnet'
    if re.search(r'(gg|golden.?gate).{0,10}travel.{0,10}(poster.{0,5})?magnet', nl):
        return 'Golden Gate Travel Poster Magnet'

    # Retro GG / GG Travel Poster stickers
    if re.search(r'retro.{0,8}(gg|golden.?gate).{0,20}(poster.{0,5})?sticker', nl):
        return 'Retro Golden Gate Bridge Poster Sticker'
    if re.search(r'(gg|golden.?gate).{0,10}travel.{0,10}(poster.{0,5})?sticker', nl):
        return 'Golden Gate Travel Poster Sticker'

    # Ferry Building magnets / sticker
    if re.search(r'(retro.{0,8})?ferry building.{0,20}(poster|retro).{0,10}magnet', nl):
        return 'Retro Ferry Building Poster Magnet'
    if re.search(r'ferry building.{0,10}travel.{0,10}magnet', nl):
        return 'Ferry Building Travel Poster Magnet'
    if re.search(r'(retro.{0,8})?ferry building.{0,20}(poster|retro).{0,10}sticker', nl):
        return 'Illustrated Ferry Building Landmark Sticker'

    # Retro Painted Ladies magnet
    if re.search(r'retro.{0,10}painted lad.{0,20}(poster.{0,5})?magnet', nl):
        return 'Retro Painted Ladies Poster Magnet'

    # ACRYLIC die-cut magnets
    if re.search(r'(acrylic|die.?cut).{0,20}(golden gate|gg).{0,20}magnet|(golden gate|gg).{0,20}(acrylic|die.?cut).{0,20}magnet|golden gate acrylic magnet', nl):
        return 'Golden Gate Acrylic Die Cut Magnet'
    if re.search(r'(acrylic|die.?cut).{0,20}ferry building.{0,20}magnet|ferry building.{0,20}(acrylic|die.?cut).{0,20}magnet|acrylic ferry building', nl):
        return 'Ferry Building Acrylic Die Cut Magnet'
    if re.search(r'(acrylic|die.?cut).{0,20}fisherm|fisherm.{0,20}(acrylic|die.?cut)', nl):
        return 'Fishermans Wharf Acrylic Die Cut Magnet'
    if re.search(r'(acrylic|die.?cut).{0,20}painted lad|painted lad.{0,15}(acrylic|die.?cut)|die cut painted ladies magnet|painted ladies acrylic magnet', nl):
        return 'ACRYLIC Painted Ladies Magnet'
    if re.search(r'(sf|san francisco).{0,10}house.{0,10}(acrylic|die.?cut|magnet)|acrylic.{0,15}(sf|san francisco).{0,10}house|sf houses acrylic', nl):
        return 'SF Houses Acrylic Die Cut Magnet'
    if re.search(r'(acrylic|die.?cut).{0,20}sfo|sfo.{0,20}(acrylic|die.?cut)|sfo luggage tag (acrylic|die)', nl):
        return 'SFO Luggage Tag Acrylic Magnet'

    # GG Acrylic print (not magnet)
    if re.search(r'golden gate.{0,10}(bridge.{0,5})?acrylic(?!.{0,15}magnet)', nl):
        return 'Golden Gate Bridge acrylic'

    # SF Illustrated Landmark Magnet
    if re.search(r'sf (illustrated|aleisha).{0,10}(landmark.{0,5})?magnet|aleisha magnet', nl):
        return 'SF Illustrated Landmarks Magnet'

    # Magnet sets
    if re.search(r'(sf|san francisco).{0,10}(icons?.{0,5})?magnet set|magnet set.{0,10}(sf|san francisco)|magnet set$', nl):
        return 'Magnet set San Francisco'

    # Hats
    if re.search(r'sf fog dad hat', nl): return 'SF Fog Dad Hat - Light Blue'
    if re.search(r'sf felt.{0,10}(lettered.{0,5})?dad hat', nl): return 'SF Felt Lettered Dad Hat - Navy'
    if re.search(r'sf bridge (rope hat|hat).{0,10}red', nl): return 'SF Bridge Rope Hat - Red'
    if re.search(r'sf bridge (rope hat|hat).{0,10}white', nl): return 'SF Bridge Rope Hat - White'
    if re.search(r'sf bridge (rope hat|hat).{0,10}green', nl): return 'SF Bridge Rope Hat - Green'
    if re.search(r'sf bridge (rope hat|hat).{0,10}navy', nl): return 'SF Bridge Rope Hat - Navy'
    if re.search(r'sf bridge.{0,5}(rope|dad) hat|sf bridge dad hat|sf bridge hat', nl): return 'Sf Bridge Dad hat'
    if re.search(r'(sf |san francisco )?city by the bay.{0,10}(dad )?hat|city by the bay hat', nl): return 'SF City By The Bay Dad Hat - Cream'
    if re.search(r'(sf |san francisco )?(fuzzy patch.{0,5})?trucker hat|sf trucker', nl): return 'SF Fuzzy Patch Trucker Hat - Navy'
    if re.search(r'sf.{0,5}(fog|felt|bridge|trucker|rope|dad).{0,10}hat', nl): return 'SF Fog Dad Hat - Light Blue'

    # Cards
    if re.search(r"i'?d escape alcatraz", nl): return "I'd Escape Alcatraz Card"
    if re.search(r"i'?d climb.{0,15}(hill|any)", nl) or 'id climb' in nl: return "I'd Climb Any Hill For You greeting card"
    if re.search(r'window seat', nl): return "I'd Give Up My Window Seat card"
    if re.search(r'twi?st.{0,8}turn', nl): return 'Thru twists and turns greeting card'
    if re.search(r'sunny.{0,10}75', nl): return 'Sunny and 75 in SF Card'
    if re.search(r'sweet.{0,15}(birthday|bday)|wishing.{0,10}sweet|sweetbday', nl): return 'Wishing You A Sweet Birthday Cake Card'
    if re.search(r'golden gate.{0,15}(travel.{0,5})?card|gg.{0,5}travel.{0,5}card|ggtravel card', nl): return 'Golden Gate Travel Poster Card'
    if re.search(r'ferry building.{0,10}(travel.{0,5})?card', nl): return 'Art By Aleisha Cards'
    if re.search(r'(sf|san francisco).{0,10}map.{0,10}(greeting.{0,5})?card', nl): return 'San Francisco Map Greeting Card'
    if re.search(r'bay area.{0,10}map.{0,10}(greeting.{0,5})?card', nl): return 'Bay Area Map Greeting card'

    # Postcards (single)
    if re.search(r'san francisco.{0,10}(golden gate.{0,10})?retro postcard|sfgg|sf golden gate.{0,10}retro', nl):
        return 'San Francisco Golden Gate Bridge Retro Postcard'
    if re.search(r'blue.{0,10}icons? postcard|blue sf icons? postcard|sf icon postcard blue|sf icons blue postcard|sf blue icons postcard', nl):
        return 'SF Blue Icons Postcard'
    if re.search(r'pink.{0,10}icons? postcard|sf icons postcard$|sf icons postcard.{0,5}pink', nl):
        return 'Pink Icons Postcard'
    if re.search(r'ferry building retro postcard', nl): return 'Ferry Building Retro Postcard'
    if re.search(r'fisherm.{0,10}(wharf.{0,5})?retro postcard', nl): return "Fisherman's Wharf Retro Postcard"
    if re.search(r'lombard.{0,10}(street.{0,5})?retro postcard|retro lombard', nl): return 'Lombard Street Retro Postcard'
    if re.search(r'painted ladies retro postcard|the painted ladies postcard', nl): return 'Painted Ladies Retro Postcard'
    if re.search(r'alcatraz.{0,10}retro postcard|alcatraz island retro postcard', nl): return 'Alcatraz Island Retro Postcard'

    # Sticker sheets
    if re.search(r'(sf|san francisco).{0,10}(landmark.{0,5})?sticker sheet|landmark sticker sheet|sf (vinyl|landmarks) sticker', nl):
        return 'San Francisco Landmark Sticker Sheet'

    # Totes
    if re.search(r'sf.{0,5}(map|icons?|city name).{0,5}tote|(san francisco|sf).{0,10}tote', nl):
        if 'map' in nl:        return 'SF Map tote'
        if 'icon' in nl:       return 'SF Icons Tote Bag'
        if 'city' in nl or 'blockfont' in nl: return 'San Francisco city tote'
        if 'home sweet' in nl: return 'Home Sweet San Francisco Tote'
        return 'SF Icons Tote Bag'

    # Tea towels
    if re.search(r'bay area.{0,10}(map.{0,5})?tea towel|the bay area tea towel', nl): return 'The Bay Area tea towel'
    if re.search(r'(san francisco|sf).{0,10}(map.{0,5})?tea towel|tea towel san francisco', nl): return 'San Francisco Tea Towel'
    if re.search(r'california (state|map).{0,10}tea towel|state of california tea towel|california tea towel|california map tea towel', nl): return 'California Map tea towel'
    if re.search(r'pacific coast.{0,10}tea towel', nl): return 'Pacific Coast Highway Tea Towel'
    if re.search(r'(sf|san francisco).{0,10}snowglobe.{0,10}tea towel', nl): return 'SF Snowglobe Tea Towel'
    if re.search(r'uc (berkeley|santa barbara).{0,10}tea towel|uc berkeley tea', nl): return 'UC Berkeley Tea Towels'
    if re.search(r'ucla.{0,10}tea towel', nl): return 'UCLA Tea Towels'
    if re.search(r'lake tahoe.{0,10}tea towel', nl): return 'Lake Tahoe Tea Towel'

    # Pencil pouches
    if re.search(r'(blue.{0,5})?(sf|san francisco).{0,10}(blue.{0,5})?pou?ch|sf blue pou?ch|blue pencil pou?ch|sf (icons?|pencil) pouch.{0,5}blue|pencil pou?ch', nl):
        return 'SF blue pouch'
    if re.search(r'(natural.{0,5})?pencil pou?ch|natural pou?ch', nl): return 'Natural Pencil Pouches'

    # Keychains
    if re.search(r'green painted lad.{0,10}keychain|painted lad.{0,10}green.{0,10}keychain|painted lady keychain.{0,5}green', nl): return 'Painted Lady Keychain - Green'
    if re.search(r'yellow painted lad.{0,10}keychain|painted lad.{0,10}yellow.{0,10}keychain|painted lady keychain.{0,5}yellow', nl): return 'Painted Lady Keychain - Yellow'
    if re.search(r'pink painted lad.{0,10}keychain|painted lad.{0,10}pink.{0,10}keychain|painted lady keychain.{0,5}pink', nl): return 'Painted Lady Keychain - Pink'
    if re.search(r'blue painted lad.{0,10}keychain|painted lad.{0,10}blue.{0,10}keychain|painted lady keychain.{0,5}blue', nl): return 'Painted Lady Keychain - Blue'
    if re.search(r'sfo.{0,10}(luggage.{0,5})?keychain|sfo keychain', nl): return 'SFO Luggage Tag Keychain'
    if re.search(r'acrylic keychain|sf.{0,5}(icon.{0,5})?keychain|california (icon.{0,5})?keychain|golden gate icon keychain|blue victorian house', nl): return 'Acrylic keychain'

    # SFO luggage tag magnet / sticker
    if re.search(r'sfo luggage (die cut|tag acrylic magnet|tag magnet|sticker)|vntage sfo|vintage sfo', nl): return 'SFO Luggage Tag Acrylic Magnet'

    # Home Sweet SF / Home Sweet Home
    if re.search(r'home sweet (sf|san francisco|home)', nl):
        if 'magnet' in nl: return 'Home Sweet SF Magnet'
        if 'sticker' in nl: return 'Home Sweet Home Sticker'
        if 'tote' in nl:   return 'Home Sweet San Francisco Tote'
        return 'Home Sweet San Francisco Art Print 8x8'

    # Stickers
    if re.search(r'(sf|san francisco).{0,10}city name.{0,10}(with golden gate.{0,5})?sticker|sfcityname', nl): return 'SF City Name With Golden Gate Sticker'
    if re.search(r'(sf|san francisco).{0,10}map sticker|sfmap sticker', nl): return 'SF Map Sticker'
    if re.search(r'bay area map sticker', nl): return 'Bay Area Map Sticker'
    if re.search(r'west coast best coast sticker', nl): return 'West Coast Best Coast sticker'
    if re.search(r'vintage.{0,5}sfo luggage tag sticker|luggagetag sfo sticker', nl): return 'Vintage SFO Luggage Tag Sticker'
    if re.search(r'(ca|california) state sticker.{0,10}(black|b.?w)|b.?w california sticker|california state.{0,10}(black|b.?w)', nl): return 'California State Sticker (Black and White)'
    if re.search(r'(ca|california) state sticker.{0,10}blue|california state.{0,10}blue|blue (ca|california) sticker', nl): return 'California State Sticker (Blue)'
    if re.search(r'(red.{0,5})?(sf|san francisco).{0,5}pill sticker|pill.{0,5}sf|san francisco pill|red pill sticker|california pill', nl): return 'Red San Francisco Pill Sticker'
    if re.search(r'proud tourist sticker', nl): return 'Proud Tourist sticker'
    if re.search(r'i come with baggage', nl): return 'I Come With Baggage Sticker'
    if re.search(r'bon voyage sticker', nl): return 'Bon Voyage Sticker'
    if re.search(r'b.?w city by the bay sticker|white sf city by the bay|sfcitybythebay.{0,5}circle.{0,5}sticker', nl): return 'White SF City By The Bay Circle Sticker'
    if re.search(r'pink sf city by the bay.{0,10}circle sticker|pink city by the bay.{0,10}sticker|pink city sticker', nl): return 'Pink SF City By The Bay Circle Sticker'
    if re.search(r'(sf|san francisco).{0,10}(national park|pennant).{0,5}sticker', nl): return 'San Francisco Pennant Sticker'
    if re.search(r'golden gate bridge sticker', nl): return 'Golden Gate Bridge Sticker (pink)'
    if re.search(r'illustrated.{0,20}(golden gate|gg).{0,20}(landmark.{0,5})?sticker', nl): return 'Illustrated Golden Gate Bridge Landmark Sticker'
    if re.search(r'illustrated.{0,20}fisherm.{0,20}(landmark.{0,5})?sticker', nl): return "Illustrated Fisherman's Wharf Landmark sticker"
    if re.search(r'retro golden gate bridge poster sticker|retro gg.{0,5}(bridge.{0,5})?sticker', nl): return 'Retro Golden Gate Bridge Poster Sticker'

    # Magnets (non-acrylic)
    if re.search(r'(sf|san francisco).{0,10}(pink.{0,5})?icons? magnet|sf pink icons? magnet', nl): return 'SF Pink Icons Magnet'
    if re.search(r'(sf|san francisco).{0,10}block font magnet', nl): return 'SF Block Font Magnet'
    if re.search(r'city by the bay.{0,10}(local notion.{0,5})?magnet|local notion.{0,10}city by the bay', nl): return 'City by the Bay Local Notion Magnet'
    if re.search(r'red greetings from sf.{0,10}(local notion.{0,5})?magnet', nl): return 'Red Greetings From SF Magnet'
    if re.search(r'cable car.{0,10}(fridgedoor.{0,5})?magnet|yellow.{0,5}cable.?car.{0,5}magnet|retro cablecar.{0,5}(poster.{0,5})?magnet|cablecar.{0,10}magnet|cable car magnet', nl): return 'Yellow Cable Car Magnet'
    if re.search(r'pink city by the bay.{0,10}(circle.{0,5})?magnet|pink city.{0,5}(by the bay.{0,5})?magnet|sf pink city circle magnet|magnet.{0,10}sfcitybythebay', nl): return 'Pink City By The Bay Circle Magnet'
    if re.search(r'home sweet sf magnet|home sweet (san francisco|home).{0,5}magnet', nl): return 'Home Sweet SF Magnet'
    if re.search(r'take the scenic route magnet|scenic route magnet', nl): return 'Take The Scenic Route Magnet'
    if re.search(r'california.{0,10}(palm tree.{0,5})?magnet.{0,10}(fridgedoor)?|fridgedoor.{0,5}california', nl): return 'California Palm Tree Magnet'
    if re.search(r'red golden gate bridge.{0,5}square magnet', nl): return 'Red Golden Gate Bridge Square Magnet'
    if re.search(r'(blue.{0,5})?sf.{0,5}waves?.{0,5}magnet|sfwave magnet', nl): return 'Blue SF Waves Magnet'

    return name.strip()


# ── Step 6: Merge by canonical name ───────────────────────────────────────────

def merge_rows(records: list[dict]) -> list[dict]:
    """Group by canonical item name, sum unit counts, keep first-seen SKU."""
    groups: dict[str, dict] = defaultdict(lambda: {"item": "", "sku": "", "qty": 0})
    for r in records:
        raw_name  = r.get("item", "").strip()
        canonical = normalize(raw_name)
        key       = canonical.lower()
        g         = groups[key]
        g["item"] = canonical
        g["qty"]  += 1      # each raw row = 1 unit sold
        if not g["sku"]:
            g["sku"] = r.get("sku", "").strip()

    merged = sorted(groups.values(), key=lambda x: x["item"].lower())
    log.info(f"After merge: {len(merged)} unique items, {sum(m['qty'] for m in merged)} total units")
    return merged


# ── Step 7: Find any existing rows for this exact date range ──────────────────

def find_existing_block(sheets, date_range_label: str):
    """
    Scan column F for rows that exactly match date_range_label.
    Returns (first_row, last_row) 1-indexed, or (None, None) if not present.
    """
    result = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{FOG_CITY_TAB}'!F:F",
    ).execute()

    col_f     = result.get("values", [])
    first_row = last_row = None

    for i, row in enumerate(col_f):
        sheet_row = i + 1
        cell_val  = str(row[0]).strip() if row else ""
        if cell_val == date_range_label:
            if first_row is None:
                first_row = sheet_row
            last_row = sheet_row

    log.info(f"Existing block for '{date_range_label}': rows {first_row}–{last_row}")
    return first_row, last_row


# ── Step 8: Write to Google Sheet ─────────────────────────────────────────────

def write_to_sheet(sheets, merged: list[dict], date_range_label: str,
                   first_row, last_row, sku_lookup: dict = None):
    month_name = YESTERDAY.strftime("%B")
    year       = YESTERDAY.year

    def resolve_sku(item_name: str, ricochet_sku: str) -> tuple[str, bool]:
        """Return (sku, needs_review). needs_review=True means no match found."""
        if sku_lookup:
            sheet_sku = find_sku(item_name, sku_lookup)
            if sheet_sku:
                return sheet_sku, False
        return ricochet_sku, True  # fell back to Ricochet SKU — flag for review

    rows_needing_review = []
    new_values = []
    for i, m in enumerate(merged):
        sku, needs_review = resolve_sku(m["item"], m["sku"])
        new_values.append([month_name, year, m["item"], sku, m["qty"], date_range_label])
        if needs_review:
            rows_needing_review.append(i)

    if first_row is not None:
        # Replace the existing block in place, padding with blanks if needed
        existing_count = last_row - first_row + 1
        blanks_needed  = max(0, existing_count - len(new_values))
        padded         = new_values + [["", "", "", "", "", ""]] * blanks_needed
        range_str      = f"'{FOG_CITY_TAB}'!A{first_row}:F{first_row + len(padded) - 1}"
        log.info(f"Replacing existing {existing_count} rows at {range_str} "
                 f"with {len(new_values)} data rows + {blanks_needed} blanks")
    else:
        # Append after the last occupied row in column A
        result     = sheets.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{FOG_CITY_TAB}'!A:A",
        ).execute()
        last_occ   = len(result.get("values", []))
        append_row = last_occ + 1
        padded     = new_values
        range_str  = f"'{FOG_CITY_TAB}'!A{append_row}:F{append_row + len(padded) - 1}"
        log.info(f"Appending {len(new_values)} rows starting at row {append_row}")

    sheets.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_str,
        valueInputOption="USER_ENTERED",
        body={"values": padded},
    ).execute()

    log.info(f"✅ Sheet updated — {len(merged)} items, "
             f"{sum(m['qty'] for m in merged)} total units written.")

    # Yellow-highlight any rows where SKU couldn't be matched (needs manual review)
    if rows_needing_review:
        start_row = first_row if first_row else append_row
        yellow = {"red": 1.0, "green": 0.95, "blue": 0.0}
        requests = []
        for i in rows_needing_review:
            row_idx = start_row + i - 1  # 0-indexed for batchUpdate
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": 1018380031,  # Fog City Sales gid
                        "startRowIndex": row_idx,
                        "endRowIndex":   row_idx + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex":   6,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": yellow
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor",
                }
            })
        sheets.batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests},
        ).execute()
        log.warning(f"⚠️  {len(rows_needing_review)} rows highlighted yellow — SKU not matched, needs review.")


# ── Dashboard JSON export ─────────────────────────────────────────────────────

def build_dashboard_json(sheets):
    """
    Read the last 30 days of Fog City Sales data from the sheet,
    compute top 10 items and hot/cold trends, write data.json.
    """
    log.info("Building dashboard JSON…")

    result = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{FOG_CITY_TAB}'!A1:F25000",
    ).execute()
    all_rows = result.get("values", [])

    # Parse each row: A=Month, B=Year, C=Item, D=SKU, E=Qty, F=Source
    import re as _re
    from collections import defaultdict as _dd

    def parse_source_dates(source, year):
        tokens = _re.findall(r'\b(\d{1,2})/(\d{1,2})\b', str(source))
        if not tokens:
            return None, None
        try:
            s = date(year, int(tokens[0][0]),  int(tokens[0][1]))
            e = date(year, int(tokens[-1][0]), int(tokens[-1][1]))
            if e < s:
                e = date(year + 1, int(tokens[-1][0]), int(tokens[-1][1]))
            return s, e
        except Exception:
            return None, None

    # Group rows by source label
    source_groups = {}
    for row in all_rows[1:]:
        if len(row) < 6:
            continue
        try:
            year_val = int(float(row[1])) if row[1] else 0
        except Exception:
            year_val = 0
        if year_val < 2025:
            continue
        item   = str(row[2]).strip() if row[2] else ''
        try:
            qty = int(float(row[4])) if row[4] else 0
        except Exception:
            qty = 0
        source = str(row[5]).strip() if row[5] else ''
        if not item or not source or qty <= 0:
            continue
        key = f"{year_val}|{source}"
        if key not in source_groups:
            s, e = parse_source_dates(source, year_val)
            source_groups[key] = {'start': s, 'end': e, 'items': _dd(int)}
        source_groups[key]['items'][item] += qty

    # Determine reference date = latest end date in the sheet
    valid = [(k, g) for k, g in source_groups.items() if g['end']]
    if not valid:
        log.warning("No dated source groups found — skipping dashboard JSON.")
        return
    ref_date = max(g['end'] for _, g in valid)

    # Windows
    window_10_start = ref_date - timedelta(days=9)   # last 10 days
    window_30_start = ref_date - timedelta(days=29)  # last 30 days
    prev_start      = ref_date - timedelta(days=30)  # day 11–30 (prior 20 days)

    recent_10 = _dd(float)
    recent_30 = _dd(float)
    prev_20   = _dd(float)

    for _, g in valid:
        s, e = g['start'], g['end']
        span = max((e - s).days + 1, 1)

        # Last 10 days
        if e >= window_10_start and s <= ref_date:
            ol_s = max(s, window_10_start)
            ol_e = min(e, ref_date)
            ratio = ((ol_e - ol_s).days + 1) / span
            for item, qty in g['items'].items():
                recent_10[item] += qty * ratio

        # Last 30 days
        if e >= window_30_start and s <= ref_date:
            ol_s = max(s, window_30_start)
            ol_e = min(e, ref_date)
            ratio = ((ol_e - ol_s).days + 1) / span
            for item, qty in g['items'].items():
                recent_30[item] += qty * ratio

        # Prior 20 days (days 11–30)
        if e >= prev_start and s < window_10_start:
            ol_s = max(s, prev_start)
            ol_e = min(e, window_10_start - timedelta(days=1))
            if ol_e >= ol_s:
                ratio = ((ol_e - ol_s).days + 1) / span
                for item, qty in g['items'].items():
                    prev_20[item] += qty * ratio

    # Normalize to daily rates
    def daily(d, days):
        return {k: v / days for k, v in d.items()}

    rate_10 = daily(recent_10, 10)
    rate_20 = daily(prev_20, 20)

    # Clean up verbose Ricochet product name prefixes
    prefixes = [
        'Art By Aleisha Magnets - ', 'Art by Aleisha Magnets - ',
        'Art By Aleisha Stickers - ', 'Art by Aleisha Stickers - ',
        'Art by Aleisha Keychains - ', 'Art By Aleisha Keychains - ',
        'Pencil Pouches - ', 'Art By Aleisha Postcards - ',
        'Art by Aleisha Postcards - ', 'Art By Aleisha Cards - ',
        'Art by Aleisha Cards - ', 'Art By Aleisha Tea Towels - ',
        'Art by Aleisha Tea Towels - ', 'Art By Aleisha Hats - ',
        'Art by Aleisha Hats - ',
    ]

    def clean(name):
        for p in prefixes:
            name = name.replace(p, '')
        return name.strip()

    def merge_canonical(d):
        out = _dd(float)
        for k, v in d.items():
            out[clean(normalize(k))] += v
        return out

    r10 = merge_canonical(recent_10)
    p20 = merge_canonical(prev_20)
    rt10 = daily(r10, 10)
    rt20 = daily(p20, 20)

    # Top 10 by recent 10-day qty
    top10 = sorted(r10.items(), key=lambda x: -x[1])[:10]

    # Trends
    all_items = set(list(rt10.keys()) + list(rt20.keys()))
    trends = {}
    for item in all_items:
        rr = rt10.get(item, 0)
        rp = rt20.get(item, 0)
        qr = r10.get(item, 0)
        qp = p20.get(item, 0)
        if qr < 2 and qp < 2:
            continue
        if rp > 0.01:
            pct = (rr - rp) / rp * 100
        elif rr > 0:
            pct = 150
        else:
            pct = 0
        trends[item] = {'pct': pct, 'qty_r': qr, 'qty_p': qp}

    hot  = sorted([(k, v) for k, v in trends.items() if v['pct'] >= 50 and v['qty_r'] >= 2],
                  key=lambda x: -x[1]['pct'])[:6]
    cold = sorted([(k, v) for k, v in trends.items() if v['pct'] <= -50 and v['qty_p'] >= 3],
                  key=lambda x: x[1]['pct'])[:6]

    period_label = (
        f"{window_10_start.strftime('%-m/%-d')} – "
        f"{ref_date.strftime('%-m/%-d, %Y')}"
    )

    data = {
        "updated":      ref_date.strftime("%-B %-d, %Y"),
        "period_label": period_label,
        "total_units":  round(sum(v for _, v in top10)),
        "top10": [
            {"name": clean(normalize(k)), "qty": round(v)}
            for k, v in top10
        ],
        "hot": [
            {
                "name":       k,
                "qty_recent": round(v['qty_r']),
                "qty_prev":   round(v['qty_p']),
                "pct":        round(v['pct']),
            }
            for k, v in hot
        ],
        "cold": [
            {
                "name":       k,
                "qty_recent": round(v['qty_r']),
                "qty_prev":   round(v['qty_p']),
                "pct":        round(v['pct']),
            }
            for k, v in cold
        ],
    }

    with open("data.json", "w") as f:
        json.dump(data, f, indent=2)

    log.info(f"✅ data.json written — top10: {len(top10)} items, "
             f"hot: {len(hot)}, cold: {len(cold)}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    log.info("=== Ricochet daily sync starting ===")
    log.info(f"Today (PDT): {TODAY}  |  Yesterday: {YESTERDAY}")

    # Connect to Sheets first — needed to determine the date window
    sheets = get_sheets_service()

    # Load SKU lookup from Inventory Summary tab
    sku_lookup = build_sku_lookup(sheets)

    # Determine window: day after last upload → yesterday
    window_start = get_window_start(sheets)
    window_end   = YESTERDAY

    if window_start > window_end:
        log.info(
            f"Sheet is already up to date through {window_end} "
            f"(last upload ended {window_end}). Nothing to do."
        )
        build_dashboard_json(sheets)
        return

    date_range_label = (
        f"{window_start.month}/{window_start.day} - "
        f"{window_end.month}/{window_end.day} {SOURCE_LABEL}"
    )
    log.info(f"Date range: {window_start} → {window_end}  |  Label: '{date_range_label}'")

    # Scrape Ricochet and filter to window
    raw_records = scrape_ricochet()
    filtered    = filter_by_date(raw_records, window_start, window_end)

    if not filtered:
        log.warning(
            f"No Ricochet sales found between {window_start} and {window_end}. "
            "Nothing to write."
        )
        return

    # Merge duplicates into one row per product
    merged = merge_rows(filtered)

    # Write to sheet (replace existing block or append)
    first_row, last_row = find_existing_block(sheets, date_range_label)
    write_to_sheet(sheets, merged, date_range_label, first_row, last_row, sku_lookup)

    # Always regenerate dashboard JSON after updating the sheet
    build_dashboard_json(sheets)

    log.info("=== Done ===")


if __name__ == "__main__":
    main()
