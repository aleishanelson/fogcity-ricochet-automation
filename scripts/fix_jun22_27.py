"""
fix_jun22_27.py
One-off backfill: re-write the 6/22-6/25 and 6/27 Fog City Sales rows
with correct per-item revenue (Agreed price) and correct inventory SKUs.

Background:
- 6/22-6/25: written before the agreed-price fix; revenue column has only
  a single total in the last row of each block instead of per-item values.
- 6/26: written after the fix - correct, no changes needed.
- 6/27: written with old OVERRIDES; several items got Ricochet numeric
  SKUs (3037, 00304K, 0030W9) instead of inventory SKUs, and
  revenue is missing entirely.

Row ranges (1-indexed, from sheet inspection 2026-06-30):
  6/22  rows 10913-10933  (21 rows)
  6/23  rows 10934-10958  (25 rows)
  6/24  rows 10959-10979  (21 rows)
  6/25  rows 10980-11005  (26 rows)
  6/26  rows 11006-11034  SKIP (correct)
  6/27  rows 11035-end    (find end dynamically)
"""

import os, json
import base64 as _b64

from google.oauth2 import service_account
from googleapiclient.discovery import build

_sa_raw = os.environ["GOOGLE_SA_JSON"].strip()
try:
    SA_JSON = json.loads(_b64.b64decode(_sa_raw).decode())
except Exception:
    SA_JSON = json.loads(_sa_raw)
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
FOG_CITY_TAB   = "Fog City Sales"
SHEET_GID      = 1018380031

def get_sheets():
    creds = service_account.Credentials.from_service_account_info(
        SA_JSON, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return build("sheets", "v4", credentials=creds).spreadsheets()

OVERRIDES = {
    "alcatraz island postcard": "ALCATRAZ_RETRO_PC",
    "alcatraz island retro postcard": "ALCATRAZ_RETRO_PC",
    "fishermans wharf sf postcard": "FISHERMANSWHARF_RETRO_PC",
    "fisherman's wharf retro postcard": "FISHERMANSWHARF_RETRO_PC",
    "ferry building retro postcard": "FERRYBUILDING_RETRO_PC",
    "lombard street retro postcard": "LOMBARDST_RETRO_PC",
    "painted ladies retro postcard": "PAINTEDLADIES_RETRO_PC",
    "pink icons postcard": "SFICONS_POSTCARD_PINK",
    "sf blue icons postcard": "SFICONS_BLUE_4x6",
    "blue icons postcard": "SFICONS_BLUE_4x6",
    "golden gate bridge retro postcard": "SFGGBRIDGE_RETRO_PCARD",
    "san francisco golden gate bridge retro postcard": "SFGGBRIDGE_RETRO_PCARD",
    "postcards 3 for $11": "postcards3for11",
    "postcards- 3 for $11": "postcards3for11",
    "postcards 3 for $10": "postcards3for11",
    "fishermans wharf acrylic die cut magnet": "MAG-AC-SF-FW",
    "acrylic fishermans wharf": "MAG-AC-SF-FW",
    "ferry building acrylic die cut magnet": "MAG-AC-SF-FERRYB",
    "golden gate acrylic die cut magnet": "MAG-AC-SF-GGB",
    "golden gate bridge acrylic die cut magnet": "MAG-AC-SF-GGB",
    "acrylic painted ladies magnet": "MAG-SF-PLADIES-CL",
    "sf houses acrylic die cut magnet": "MAG-AC-SF-HOUSES",
    "sfo luggage tag acrylic magnet": "MAG-AC-SF-SFO",
    "home sweet sf magnet": "MAGNET_HOMESWEETSF",
    "home sweet home magnet": "MAGNET_HOMESWEETSF",
    "retro gg travel poster magnet": "MAG-SF-RETRO-GGB",
    "retrogg bridge travel poster magent": "MAG-SF-RETRO-GGB",
    "retro golden gate bridge poster magnet": "MAG-SF-RETRO-GGB",
    "retro golden gate poster magnet": "MAG-SF-RETRO-GGB",
    "retro ferry building poster magnet": "MAG-SF-RETRO-FB",
    "retro painted ladies poster magnet": "MAG-SF-RETRO-PL",
    "ferry building travel poster magnet": "FERRYBUILDINGTRAVELPOSTER_MAGNET",
    "sf illustrated landmarks magnet": "MAG-SF-LDMKS",
    "sf landmark magnet": "MAG-SF-LDMKS",
    "sf pink icons magnet": "MAG-SF-PINKICONS",
    "sf block font magnet": "MAG-SF-BLOCKFONT",
    "cable car magnet": "MAG-SF-CABLECAR",
    "yellow cable car magnet": "MAG-SF-CABLECAR",
    "take the scenic route magnet": "TAKETHESCENICROUTE_49MILE_MAGNET",
    "sfo luggage tag magnet": "SFO_LUGGAGETAG_MAGNET",
    "pink city by the bay circle magnet": "MAGNET_SFCITYBYTHEBAY_PINK",
    "city by the bay local notion magnet": "MAGNET_SFCITYBYTHEBAY_PINK",
    "sf magnet set": "SANFRANCISCOICONS_MAGNETSET",
    "san francisco magnet set": "SANFRANCISCOICONS_MAGNETSET",
    "magnet set san francisco": "SANFRANCISCOICONS_MAGNETSET",
    "sf icons magnet set": "SANFRANCISCOICONS_MAGNETSET",
    "west coast best coast magnet": "MAG_WESTBESTCOAST_CIR",
    "stickers 3 for $11": "STICKERS_3FOR11",
    "stickers- 3 for $11": "STICKERS_3FOR11",
    "golden gate bridge sticker (pink)": "GGBRIDGE_PINK_STICKER",
    "retro golden gate bridge poster sticker": "RETRO_GGB_STICKER",
    "retro sf ferry building poster sticker": "RETROSFFERRYBUILDING_STICKER",
    "retro ferry building poster sticker": "RETROSFFERRYBUILDING_STICKER",
    "retro ferry building sticker": "RETROSFFERRYBUILDING_STICKER",
    "illustrated fisherman's wharf sticker": "FW_ILLUSTRATED_STICKER",
    "illustrated fishermans wharf landmark sticker": "FW_ILLUSTRATED_STICKER",
    "illustrated ferry building landmark sticker": "FB_ILLUSTRATED_STICKER",
    "illustrated golden gate bridge sticker": "GGB_ILLUSTRATED_STICKER",
    "sf landmark sticker sheet": "SS-SF-LDMKS",
    "san francisco landmark sticker sheet": "SS-SF-LDMKS",
    "sf map sticker": "SFMAP_STICKER",
    "sf city name sticker": "SFCITYNAME_STICKER",
    "sf pennant sticker": "SFPENNANT_STICKER",
    "home sweet home sticker": "HOMESWEETSANFRANCISCO_STICKER",
    "west coast best coast sticker": "WESTCOASTBESTCOAST_CIRCLE_STICKER",
    "proud tourist sticker": "PROUDTOURIST_STICKER",
    "bon voyage sticker": "BONVOYAGE_STICKER",
    "i come with baggage sticker": "ICOMEWITHBAGGAGE_STICKER",
    "pink sf city by the bay circle sticker": "SFCITYBYTHEBAY_PINKCIRCLE_STICKER",
    "california state sticker (blue)": "CALIFORNIASTATE_BLOCKFONT_BLUE",
    "blue ca sticker": "CALIFORNIASTATE_BLOCKFONT_BLUE",
    "california state sticker (black and white)": "CALIFORNIASTATE_BLOCKFONT_BW",
    "red san francisco pill sticker": "PILL_SF_RED_STICKER",
    "san francisco pill sticker": "PILL_SF_RED_STICKER",
    "vntage sfo luggage tag sticker": "LUGGAGETAG_SFO_STICKER",
    "sf city name with golden gate sticker": "SFCITYNAME_STICKER",
    "sf icons tote": "SFICONS_TOTE",
    "sf city name tote": "SF_BLOCKFONT_TOTE",
    "home sweet home tote": "TOTE_HOMESWEETSF",
    "home sweet san francisco tote": "TOTE_HOMESWEETSF",
    "sf map tote": "SF_MAP_TOTE",
    "sfo luggage tag keychain": "KC-SFO-LUGGAGETAG",
    "acrylic keychain": "KC-SF-01",
    "california icon keychain": "KC-SF-01",
    "california keychain": "KC-SF-01",
    "golden gate icon keychain (icon series)": "KC-SF-01",
    "blue victorian house keychain (icon series)": "KC-SF-01",
    "painted lady keychain - pink": "KC-PAINTEDLADY-PINK",
    "painted lady keychain - blue": "KC-PAINTEDLADY-BLUE",
    "painted lady keychain - green": "KC-PAINTEDLADY-GREEN",
    "painted lady keychain - yellow": "KC-PAINTEDLADY-YELLOW",
    "sunny and 75 in sf card": "LOVEYOUMORETHANSUNNYSF_A2CARD",
    "twists and turns card": "TWISTSANDTURNS_GCARD",
    "thru twists and turns card": "TWISTSANDTURNS_GCARD",
    "thru twists and turns greeting card": "TWISTSANDTURNS_GCARD",
    "window seat card": "WINDOWSEAT_A2_GREETINGCARD",
    "home sweet san francisco card": "HOMESWEETSF_A2CARD",
    "home sweet sf greeting card": "HOMESWEETSF_A2CARD",
    "sweetest birthday greeting card": "SWEETBDAYCAKE_A2_GCARD",
    "wishing you a sweet birthday cake card": "SWEETBDAYCAKE_A2_GCARD",
    "i'd escape alcatraz card": "IDESCAPEALCATRAZFORYOU_A2CARD",
    "i'd climb any hill card": "IDCLIMBANYHILL_A2CARD",
    "sf icons greeting card": "SFICONS_GREETINGCARD",
    "golden gate travel poster card": "GOLDENGATETRAVELPOSTER_A2CARD",
    "sf icons pencil pouch - natural": "pp-sf-cn-02",
    "sf icons pencil pouch - blue": "PP-SF-CB-01",
    "blue pencil pouches": "PP-SF-CB-01",
    "natural pencil pouches": "pp-sf-cn-02",
    "sf blue pouch": "PP-SF-CB-01",
    "uc berkeley tea towel": "TT-CAMPUS-BERKELEY",
    "uc santa barbara tea towel": "TT-CAMPUS-SANTABARBARA",
    "ucla tea towels": "TT-CAMPUS-UCLA",
    "lake tahoe tea towel": "TT-LAKETAHOE",
    "san francisco map tea towel": "SANFRANCISCO_MAP_DISHTOWEL",
    "bay area tea towel": "BAYAREA_MAP_TEATOWEL",
    "california tea towel": "STATEOFCALIFORNIA_MAP_TEATOWEL",
    "home sweet san francisco art print 8x8": "HOMESWEETSF_8x8",
    "home sweet san francisco art print 8x10": "HOMESWEETSF_8x10",
    "home sweet san francisco art print 11x11": "HOMESWEETSF_11x15",
    "golden gate travel poster 8x10": "GOLDENGATE_TRAVELPOSTER_8x10",
    "golden gate travel poster 11x14": "GOLDENGATE_TRAVELPOSTER_11x14",
    "ferry building travel poster - 8x10": "FERRYBUILDINGTRAVELPOSTER_8x10",
    "san francisco map print 8x10": "SF_BW_8x10",
    "bay area map print 8x10": "BAYAREA_BW_8x10",
    "bay area map print 9x12": "BAYAREA_BW_9x12",
    "jersey city map print 8x10": "JERSEYCITY_BW_8x10",
    "los angeles map print 8x10": "LA_BW_8x10",
    "napa valley map print 8x10": "NAPAVALLEY_BW_8x10",
    "minneapolis map print 8x10": "MPLS_BW_8x10",
    "pittsburgh map print 8x10": "PITTSBURGH_BW_8x10",
    "paris map print 8x10": "PARIS_BW_8x10",
    "stanford campus map print 8x10": "STANFORD_BW_8x10",
    "uc berkeley campus map print 8x10": "UCBERKELEY_CAMPUS_BW_8x10",
    "ucla campus map print 8x10": "UCLA_BW_8x10",
    "ohio state university campus map print 8x10": "OHIOSTATE_BW_8x10",
    "purdue university campus map print 8x10": "PURDUE_BW_8x10",
    "austin map print 8x10": "AUSTIN_BW_8x10",
    "cape cod map print 8x10": "CAPECOD_BW_8x10",
    "chicago map print 8x10": "CHICAGO_BW_8x10",
    "santa clara university campus map print 8x10": "SCU_BW_8x10_CURSIVE",
    "fog city dad hat - light blue": "DH-004-LB",
    "fog city dad hat": "DH-004-LB",
    "sf fog dad hat - light blue": "DH-004-LB",
    "sf felt lettered dad hat": "DH-003-NB",
    "sf bridge dad hat - navy": "DH-001-NB",
    "sf city by the bay dad hat - cream": "DH-002-CR",
    "sf city by the bay dad hat": "DH-002-CR",
    "sf fuzzy patch trucker hat - navy": "TH-001-NB",
    "golden gate travel poster magnet": "MAGNET_GOLDENGATETRAVELPOSTER",
    "retro painted ladies travel poster": "MAG-SF-RETRO-PL",
    "bay area map greeting card": "BAYAREAMAP_A2CARD",
}

def find_sku_override(item_name):
    key = item_name.strip().lower()
    if key in OVERRIDES:
        return OVERRIDES[key]
    for oname, osku in OVERRIDES.items():
        if key in oname or oname in key:
            return osku
    return ""

def build_sku_lookup(sheets):
    result = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'Inventory Summary'!A:E",
    ).execute()
    rows = result.get("values", [])
    all_lookup = {}
    for row in rows[1:]:
        if len(row) < 5:
            continue
        item_name = str(row[1]).strip()
        sku = str(row[4]).strip()
        if item_name and sku:
            all_lookup[item_name.lower()] = sku
    return all_lookup

def resolve_sku(item_name, all_lookup):
    sku = find_sku_override(item_name)
    if sku:
        return sku, False
    key = item_name.strip().lower()
    if key in all_lookup:
        return all_lookup[key], False
    for lname, lsku in all_lookup.items():
        if key in lname or lname in key:
            return lsku, False
    return "", True

# Data from Ricochet payout table scraped 2026-06-30
# Each tuple: (item_name, total_revenue_for_date, qty)
BLOCKS = {
    "6/22": {
        "label": "6/22 - 6/22 ricochet export",
        "first_row": 10913,
        "last_row": 10933,
        "data": [
            ("Acrylic keychain", 10.00, 1),
            ("Fisherman's Wharf Retro Postcard", 4.00, 3),
            ("Fishermans Wharf Acrylic Die Cut Magnet", 8.50, 1),
            ("Golden Gate Bridge Sticker (pink)", 4.00, 2),
            ("Golden Gate Travel Poster 8x10", 40.00, 1),
            ("Golden Gate Travel Poster Card", 26.00, 4),
            ("Home Sweet Home Sticker", 4.00, 2),
            ("Home Sweet SF Magnet", 8.00, 3),
            ("Magnet set San Francisco", 42.00, 3),
            ("Painted Ladies Retro Postcard", 16.00, 4),
            ("Pink Icons Postcard", 8.00, 2),
            ("Pink SF City By The Bay Circle Sticker", 8.00, 2),
            ("Postcards 3 for $11", 20.00, 2),
            ("Retro Golden Gate Bridge Poster Sticker", 16.00, 4),
            ("SF Blue Icons Postcard", 4.00, 1),
            ("SF blue pouch", 16.00, 1),
            ("SF City By The Bay Dad Hat - Cream", 34.00, 1),
            ("SF Illustrated Landmarks Magnet", 32.00, 4),
            ("SF Pink Icons Magnet", 8.00, 1),
            ("Stickers- 3 for $11", 99.00, 9),
            ("Wishing You A Sweet Birthday Cake Card", 6.50, 1),
        ],
    },
    "6/23": {
        "label": "6/23 - 6/23 ricochet export",
        "first_row": 10934,
        "last_row": 10958,
        "data": [
            ("Acrylic keychain", 30.00, 3),
            ("Fisherman's Wharf Retro Postcard", 4.00, 1),
            ("Golden Gate Bridge Sticker (pink)", 12.00, 3),
            ("Golden Gate Travel Poster Card", 45.50, 7),
            ("Home Sweet Home Sticker", 20.00, 5),
            ("Home Sweet SF Magnet", 8.00, 1),
            ("Jersey City Map Print 8x10", 20.00, 1),
            ("Lombard Street Retro Postcard", 4.00, 1),
            ("Los Angeles Map Print 8x10", 20.00, 1),
            ("Magnet set San Francisco", 28.00, 2),
            ("Painted Ladies Retro Postcard", 4.00, 1),
            ("Postcards 3 for $11", 40.00, 4),
            ("Retro Ferry Building Poster Magnet", 8.00, 1),
            ("Retro Golden Gate Bridge Poster Sticker", 12.00, 3),
            ("San Francisco Pennant Sticker", 4.00, 1),
            ("SF Block Font Magnet", 8.00, 1),
            ("SF Blue Icons Postcard", 8.00, 2),
            ("SF Illustrated Landmarks Magnet", 16.00, 2),
            ("SF Pink Icons Magnet", 8.00, 1),
            ("SFO Luggage Tag Acrylic Magnet", 42.50, 5),
            ("SFO Luggage Tag Keychain", 13.00, 1),
            ("Stickers- 3 for $11", 110.00, 10),
            ("Sunny and 75 in SF Card", 6.50, 1),
            ("West Coast Best Coast Magnet", 8.00, 1),
        ],
    },
    "6/24": {
        "label": "6/24 - 6/24 ricochet export",
        "first_row": 10959,
        "last_row": 10979,
        "data": [
            ("Acrylic keychain", 20.00, 2),
            ("Fishermans Wharf Acrylic Die Cut Magnet", 8.50, 1),
            ("Golden Gate Acrylic Die Cut Magnet", 8.50, 1),
            ("Golden Gate Travel Poster Card", 26.00, 4),
            ("Home Sweet Home Sticker", 4.00, 1),
            ("Home Sweet SF Magnet", 40.00, 5),
            ("I Come With Baggage Sticker", 4.00, 1),
            ("I'd Escape Alcatraz Card", 6.50, 1),
            ("Illustrated Ferry Building Landmark Sticker", 8.00, 2),
            ("Magnet set San Francisco", 28.00, 2),
            ("Retro Ferry Building Poster Magnet", 16.00, 2),
            ("Retro Golden Gate Bridge Poster Sticker", 4.00, 1),
            ("Retro Painted Ladies Poster Magnet", 8.00, 1),
            ("San Francisco Pennant Sticker", 4.00, 1),
            ("SF blue pouch", 8.00, 1),
            ("SF City Name With Golden Gate Sticker", 4.00, 1),
            ("SF Map Sticker", 4.00, 1),
            ("SFO Luggage Tag Acrylic Magnet", 42.50, 5),
            ("Stickers- 3 for $11", 33.00, 3),
            ("Take The Scenic Route Magnet", 24.00, 3),
            ("Yellow Cable Car Magnet", 8.00, 1),
        ],
    },
    "6/25": {
        "label": "6/25 - 6/25 ricochet export",
        "first_row": 10980,
        "last_row": 11005,
        "data": [
            ("Acrylic keychain", 10.00, 1),
            ("Alcatraz Island Retro Postcard", 4.00, 1),
            ("Bay Area Map Greeting card", 6.50, 1),
            ("City by the Bay Local Notion Magnet", 16.00, 2),
            ("Ferry Building Retro Postcard", 4.00, 1),
            ("Fisherman's Wharf Retro Postcard", 4.00, 1),
            ("Fishermans Wharf Acrylic Die Cut Magnet", 17.00, 2),
            ("Golden Gate Bridge Sticker (pink)", 8.00, 2),
            ("Golden Gate Travel Poster Card", 45.50, 7),
            ("Home Sweet Home Sticker", 12.00, 3),
            ("Home Sweet SF Magnet", 16.00, 2),
            ("Illustrated Ferry Building Landmark Sticker", 4.00, 1),
            ("Lombard Street Retro Postcard", 4.00, 1),
            ("Magnet set San Francisco", 14.00, 1),
            ("Painted Ladies Retro Postcard", 4.00, 1),
            ("Pink SF City By The Bay Circle Sticker", 4.00, 1),
            ("Postcards 3 for $11", 20.00, 2),
            ("Retro Painted Ladies Travel Poster", 8.00, 1),
            ("San Francisco Pennant Sticker", 4.00, 1),
            ("SF Block Font Magnet", 8.00, 1),
            ("SF City Name With Golden Gate Sticker", 8.00, 2),
            ("SF Houses Acrylic Die Cut Magnet", 8.50, 1),
            ("SF Map Sticker", 4.00, 1),
            ("Stickers- 3 for $11", 22.00, 2),
            ("Sunny and 75 in SF Card", 6.50, 1),
            ("West Coast Best Coast Magnet", 8.00, 1),
        ],
    },
    "6/27": {
        "label": "6/27 - 6/27 ricochet export",
        "first_row": 11035,
        "last_row": None,
        "data": [
            ("Bay Area Map Print 9x12", 25.00, 1),
            ("Bon Voyage Sticker", 4.00, 1),
            ("California State Sticker (Black and White)", 4.00, 1),
            ("Ferry Building Retro Postcard", 4.00, 1),
            ("Golden Gate Bridge Sticker (pink)", 8.00, 2),
            ("Golden Gate Travel Poster Card", 45.50, 7),
            ("Home Sweet Home Sticker", 12.00, 3),
            ("Home Sweet San Francisco Art Print 8x8", 40.00, 2),
            ("Home Sweet SF Magnet", 32.00, 4),
            ("I'd Escape Alcatraz Card", 6.50, 1),
            ("Illustrated Ferry Building Landmark Sticker", 8.00, 2),
            ("Magnet set San Francisco", 28.00, 2),
            ("Painted Ladies Retro Postcard", 4.00, 1),
            ("Painted Lady Keychain - Blue", 24.00, 2),
            ("Pink SF City By The Bay Circle Sticker", 4.00, 1),
            ("Postcards 3 for $11", 90.00, 9),
            ("Purdue University Campus Map Print 8x10", 20.00, 1),
            ("Retro Ferry Building Poster Magnet", 16.00, 2),
            ("Retro Painted Ladies Poster Magnet", 8.00, 1),
            ("San Francisco Pennant Sticker", 4.00, 1),
            ("SF Block Font Magnet", 8.00, 1),
            ("SF Blue Icons Postcard", 4.00, 1),
            ("SF blue pouch", 16.00, 1),
            ("SF City By The Bay Dad Hat - Cream", 34.00, 1),
            ("SF Fuzzy Patch Trucker Hat - Navy", 34.00, 1),
            ("SF Icons Tote Bag", 30.00, 1),
            ("SF Map Sticker", 4.00, 1),
            ("SFO Luggage Tag Acrylic Magnet", 8.50, 1),
            ("Stickers- 3 for $11", 99.00, 9),
            ("West Coast Best Coast Magnet", 8.00, 1),
        ],
    },
}

def fix_block(sheets, all_lookup, block_key):
    block = BLOCKS[block_key]
    label = block["label"]
    first_row = block["first_row"]
    data = block["data"]
    yellow = {"red": 1.0, "green": 0.95, "blue": 0.0}

    if block["last_row"] is None:
        col_f = sheets.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{FOG_CITY_TAB}'!F{first_row}:F{first_row+200}",
        ).execute().get("values", [])
        last_row = first_row
        for i, r in enumerate(col_f):
            if r and str(r[0]).strip() == label:
                last_row = first_row + i
        block["last_row"] = last_row
    last_row = block["last_row"]
    existing_count = last_row - first_row + 1

    merged = {}
    for item_name, revenue, qty in data:
        if item_name in merged:
            merged[item_name]["qty"] += qty
            merged[item_name]["revenue"] += revenue
        else:
            merged[item_name] = {"qty": qty, "revenue": revenue}

    new_values = []
    rows_needing_review = []
    for i, (item_name, d) in enumerate(sorted(merged.items())):
        sku, needs_review = resolve_sku(item_name, all_lookup)
        new_values.append([
            "June", 2026, item_name, sku,
            d["qty"], label, round(d["revenue"], 2)
        ])
        if needs_review:
            rows_needing_review.append(i)

    blanks_needed = max(0, existing_count - len(new_values))
    padded = new_values + [["", "", "", "", "", "", ""]] * blanks_needed
    range_str = f"'{FOG_CITY_TAB}'!A{first_row}:G{first_row + len(padded) - 1}"

    sheets.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_str,
        valueInputOption="USER_ENTERED",
        body={"values": padded},
    ).execute()
    print(f"  {block_key}: wrote {len(new_values)} items + {blanks_needed} blank rows -> {range_str}")

    if rows_needing_review:
        requests = []
        for i in rows_needing_review:
            row_0idx = first_row + i - 1
            requests.append({"repeatCell": {
                "range": {
                    "sheetId": SHEET_GID,
                    "startRowIndex": row_0idx,
                    "endRowIndex": row_0idx + 1,
                    "startColumnIndex": 0, "endColumnIndex": 7,
                },
                "cell": {"userEnteredFormat": {"backgroundColor": yellow}},
                "fields": "userEnteredFormat.backgroundColor",
            }})
        sheets.batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests},
        ).execute()
        print(f"  {block_key}: {len(rows_needing_review)} rows flagged yellow for SKU review")


def main():
    print("=== fix_jun22_27.py starting ===")
    sheets = get_sheets()
    all_lookup = build_sku_lookup(sheets)
    for key in ["6/22", "6/23", "6/24", "6/25", "6/27"]:
        print(f"Fixing {key}...")
        fix_block(sheets, all_lookup, key)
    print("=== Done ===")

if __name__ == "__main__":
    main()
