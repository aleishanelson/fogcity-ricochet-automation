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

# Map keywords in item names → the Inventory Summary categories to search first.
# If an item name contains one of these keywords, we only match against rows
# from the corresponding category before falling back to a broader search.
CATEGORY_HINTS = {
    "sticker":      {"Stickers", "Sticker", "Sticker deal", "Sticker Sheet", "Sticker Book"},
    "magnet":       {"Magnets"},
    "tote":         {"Totes"},
    "tea towel":    {"Tea Towel"},
    "towel":        {"Tea Towel"},
    "postcard":     {"Postcards", "Postcard deal"},
    "keychain":     {"Keychains", "Keychain"},
    "card":         {"Greeting Card", "Card Pack"},
    "print":        {"City Print", "School Prints", "Film Print", "Landmark"},
    "pencil pouch": {"Pencil Pouch"},
    "tote bag":     {"Totes"},
    # "poster" intentionally omitted — too ambiguous (magnets, stickers, prints all say poster)
}


def build_sku_lookup(sheets) -> dict:
    """
    Read Inventory Summary tab cols A (category), B (item name), E (SKU).
    Returns:
      lookup["all"]         → {lowercase_name: sku}  (every entry)
      lookup["by_category"] → {category: {lowercase_name: sku}}
    """
    result = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'Inventory Summary'!A:E",
    ).execute()

    rows = result.get("values", [])
    all_lookup = {}
    by_category = {}

    for row in rows[1:]:   # skip header
        if len(row) < 5:
            continue
        category  = str(row[0]).strip()
        item_name = str(row[1]).strip()
        sku       = str(row[4]).strip() if len(row) > 4 else ""
        if not item_name or not sku:
            continue
        key = item_name.lower()
        all_lookup[key] = sku
        if category:
            by_category.setdefault(category, {})[key] = sku

    log.info(f"SKU lookup loaded: {len(all_lookup)} entries across "
             f"{len(by_category)} categories from Inventory Summary")
    return {"all": all_lookup, "by_category": by_category}


def find_sku(item_name: str, lookup: dict) -> str:
    """
    Find the correct SKU for an item name.

    Strategy:
    1. Category-filtered search against the live Inventory Summary tab
    2. Broad search across all categories in Inventory Summary
    3. Hardcoded overrides (fallback only, if not found above)
    """
    key = item_name.strip().lower()
    all_lookup = lookup.get("all", lookup) if isinstance(lookup, dict) else lookup
    by_category = lookup.get("by_category", {}) if isinstance(lookup, dict) else {}

    def _match(d: dict, k: str) -> str:
        """Exact then partial match within a given sub-dict."""
        if k in d:
            return d[k]
        for lookup_name, sku in d.items():
            if k in lookup_name or lookup_name in k:
                return sku
        return ""

    # Hardcoded overrides - checked only as a fallback, AFTER the live
    # Inventory Summary tab lookup (Inventory Summary is the
    # source of truth for SKUs; overrides only apply if the item
    # isn't found there).

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
        # GG Travel Poster PRINTS (Landmark category, GOLDENGATE_TRAVELPOSTER_*)
        "golden gate travel poster 8x10":                "GOLDENGATE_TRAVELPOSTER_8x10",
        "golden gate travel poster (new version) 8x10":  "GOLDENGATE_TRAVELPOSTER_8x10",
        "golden gate travel poster - 8x10":              "GOLDENGATE_TRAVELPOSTER_8x10",
        "golden gate bridge travel poster 8x10":         "GOLDENGATE_TRAVELPOSTER_8x10",
        "golden gate bridge travel poster - 8x10":       "GOLDENGATE_TRAVELPOSTER_8x10",
        "8x10 golden gate travel poster":                "GOLDENGATE_TRAVELPOSTER_8x10",
        "gg travel poster 8x10":                         "GOLDENGATE_TRAVELPOSTER_8x10",
        "gg travel poster - 8x10":                       "GOLDENGATE_TRAVELPOSTER_8x10",
        "golden gate travel poster 11x14":               "GOLDENGATE_TRAVELPOSTER_11x14",
        "golden gate travel poster - 11x14":             "GOLDENGATE_TRAVELPOSTER_11x14",
        "golden gate bridge travel poster 11x14":        "GOLDENGATE_TRAVELPOSTER_11x14",
        "golden gate bridge travel poster - 11x14":      "GOLDENGATE_TRAVELPOSTER_11x14",
        "gg travel poster 11x14":                        "GOLDENGATE_TRAVELPOSTER_11x14",
        "gg travel poster - 11x14":                      "GOLDENGATE_TRAVELPOSTER_11x14",
        "golden gate travel poster 12x16":               "GOLDENGATE_TRAVELPOSTER_12x16",
        "golden gate travel poster - 12x16":             "GOLDENGATE_TRAVELPOSTER_12x16",
        # GG Travel Poster MAGNET (separate SKU)
        "golden gate travel poster magnet":              "MAGNET_GOLDENGATETRAVELPOSTER",
        "golden gate bridge travel poster magnet":       "MAGNET_GOLDENGATETRAVELPOSTER",
        "gg travel poster magnet":                       "MAGNET_GOLDENGATETRAVELPOSTER",
        "golden gate travel (new version)":              "MAGNET_GOLDENGATETRAVELPOSTER",
        # GG Travel Poster STICKER (separate SKU)
        "golden gate travel sticker":                    "GOLDENGATETRAVELPOSTER_STICKER",
        "golden gate bridge travel sticker":             "GOLDENGATETRAVELPOSTER_STICKER",
        "golden gate travel":                            "GOLDENGATETRAVELPOSTER_STICKER",
        "gg travel sticker":                             "GOLDENGATETRAVELPOSTER_STICKER",
        # GG Travel Poster CARD (separate SKU — flag if not found)
        "golden gate travel poster card":                "GOLDENGATETRAVELPOSTER_A2CARD",
        "golden gate travel card":                       "GOLDENGATETRAVELPOSTER_A2CARD",
        "gg travel poster card":                         "GOLDENGATETRAVELPOSTER_A2CARD",
        "ferry building travel poster - 8x10":           "FERRYBUILDINGTRAVELPOSTER_8x10",
        "ferry building travel poster - 11x14":          "FERRYBUILDINGTRAVELPOSTER_11x14",
        "santa clara university campus map print 8x10":  "SCU_BW_8x10_CURSIVE",
        "stanford campus map print 8x10":               "STANFORD_BW_8x10",
        # California products (partial match misses "State of California" prefix)
        "california tea towel":                          "STATEOFCALIFORNIA_MAP_TEATOWEL",
        "califonia tea towel":                           "STATEOFCALIFORNIA_MAP_TEATOWEL",  # typo variant
        "califonia sticker":                             "SFCITYNAME_STICKER",             # typo variant
        # Cards that need exact routing
        "i'd escape alcatraz card":                      "IDESCAPEALCATRAZFORYOU_A2CARD",
        "id escape alcatraz card":                       "IDESCAPEALCATRAZFORYOU_A2CARD",
        "home sweet home greeting card":                 "HOMESWEETSF_A2CARD",
        "home sweet sf greeting card":                   "HOMESWEETSF_A2CARD",
        # Normalized-form overrides — keys match what normalize() returns
        "golden gate travel poster card":                "GOLDENGATETRAVELPOSTER_A2CARD",
        "golden gate travel poster magnet":              "MAGNET_GOLDENGATETRAVELPOSTER",
        "city by the bay local notion magnet":           "MAGNET_SFCITYBYTHEBAY_LOCALNOTION",
        "ferry building acrylic die cut magnet":         "MAG-AC-SF-FERRYB",
        "california map tea towel":                      "STATEOFCALIFORNIA_MAP_TEATOWEL",
        "california tea towel":                          "STATEOFCALIFORNIA_MAP_TEATOWEL",
        # Stickers
        "california state sticker (black and white)":    "CALIFORNIASTATE_BLOCKFONT_BW",
        "california state sticker black and white":      "CALIFORNIASTATE_BLOCKFONT_BW",
        "ca state sticker black and white":              "CALIFORNIASTATE_BLOCKFONT_BW",
        # Magnets — discontinued/renamed products
        "city by the bay local notion magnet":           "MAGNET_SFCITYBYTHEBAYCIRCLE_PINK",  # discontinued → Pink
        "city by the bay magnet":                        "MAGNET_SFCITYBYTHEBAYCIRCLE_PINK",
        "local notion magnet":                           "MAGNET_SFCITYBYTHEBAYCIRCLE_PINK",
        "sf magnet set":                                 "SANFRANCISCOICONS_MAGNETSET",
        "san francisco magnet set":                      "SANFRANCISCOICONS_MAGNETSET",
        "magnet set san francisco":                      "SANFRANCISCOICONS_MAGNETSET",
        # Keychains
        "painted lady keychain - pink":                  "KC-PAINTEDLADY-PINK",
        "painted lady keychain pink":                    "KC-PAINTEDLADY-PINK",
        "painted lady keychain - blue":                  "KC-PAINTEDLADY-BLUE",
        "painted lady keychain blue":                    "KC-PAINTEDLADY-BLUE",
        "painted lady keychain - green":                 "KC-PAINTEDLADY-GREEN",
        "painted lady keychain green":                   "KC-PAINTEDLADY-GREEN",
        # Retro Ferry Building Sticker — NOT a magnet
        "retro ferry building sticker":                  "RETROSFFERRYBUILDING_STICKER",
        "retro sf ferry building sticker":               "RETROSFFERRYBUILDING_STICKER",
        "retro ferry building poster sticker":           "RETROSFFERRYBUILDING_STICKER",
        # Stickers — override Ricochet numeric SKUs
        "california state sticker (blue)":               "CALIFORNIASTATE_BLOCKFONT_BLUE",
        "california state sticker blue":                 "CALIFORNIASTATE_BLOCKFONT_BLUE",
        "ca state sticker blue":                         "CALIFORNIASTATE_BLOCKFONT_BLUE",
        "golden gate sticker pink":                      "GGBRIDGE_PINK_STICKER",
        "golden gate bridge sticker pink":               "GGBRIDGE_PINK_STICKER",
        "golden gate bridge sticker (pink)":             "GGBRIDGE_PINK_STICKER",
        "illustrated fisherman's wharf landmark sticker": "FW_ILLUSTRATED_STICKER",
        "illustrated fishermans wharf landmark sticker": "FW_ILLUSTRATED_STICKER",
        # Totes — do NOT use SF_HP_* for totes
        "home sweet san francisco tote":                 "TOTE_HOMESWEETSF",
        "home sweet sf tote":                            "TOTE_HOMESWEETSF",
        "home sweet home tote":                          "TOTE_HOMESWEETSF",
        # Keychains
        "painted lady keychain - yellow":                "KC-PAINTEDLADY-YELLOW",
        "painted lady keychain yellow":                  "KC-PAINTEDLADY-YELLOW",
        # Bay Area prints
        "bay area map print 8x10":                       "BAYAREA_BW_8x10",
        "bay area map print 9x12":                       "BAYAREA_BW_9x12",
        "bay area map print 11x14":                      "BAYAREA_BW_11x14",
        "bay area map print 12x16":                      "BAYAREA_BW_12x16",
        # SF Map Prints — prevent partial match to SFMAP_STICKER
        "san francisco map print 8x10":                  "SF_BW_8x10",
        "san francisco map print 9x12":                  "SF_BW_9x12",
        "san francisco map print 11x14":                 "SF_BW_11x14",
        "san francisco map print 12x16":                 "SF_BW_12x16",
        # Chicago Map Print — NOT CHICAGO_HP_* (HP = hand painted, different product)
        "chicago map print 8x10":                        "CHICAGO_BW_8x10",
        "chicago map print 9x12":                        "CHICAGO_BW_9x12",
        "chicago map print 11x14":                       "CHICAGO_BW_11x14",
        # USA Map Print
        "usa map print 8x10":                            "USA_BW_8x10",
        "usa map print 9x12":                            "USA_BW_9x12",
        # UC Campus Map Prints — TT-CAMPUS-* are TEA TOWELS, not prints
        "uc berkeley campus map print 8x10":             "UCBERKELEY_CAMPUS_BW_8x10",
        "uc berkeley campus map print":                  "UCBERKELEY_CAMPUS_BW_8x10",
        "uc santa barbara campus map print 8x10":        "UCSANTABARBARA_BW_8x10",
        "uc santa barbara campus map print":             "UCSANTABARBARA_BW_8x10",
        # Ferry Building PRINT (size in name = print, NOT the travel poster magnet)
        "ferry building print 8x10":                     "FERRYBUILDING_TRAVELPOSTER_8x10",
        "ferry building print 11x14":                    "FERRYBUILDING_TRAVELPOSTER_11x14",
        "8x10 ferry building print":                     "FERRYBUILDING_TRAVELPOSTER_8x10",
        "ferry building travel poster print 8x10":       "FERRYBUILDING_TRAVELPOSTER_8x10",
        "ferry building travel poster print 11x14":      "FERRYBUILDING_TRAVELPOSTER_11x14",
        "ferry building travel poster print":            "FERRYBUILDING_TRAVELPOSTER_8x10",  # normalized form
        # Stickers — prevent HP SKU misassignment (HP = hand painted print)
        "red san francisco pill sticker":                "PILL_SF_RED_STICKER",
        "red sf pill sticker":                           "PILL_SF_RED_STICKER",
        "travel poster golden gate sticker":             "GOLDENGATETRAVELPOSTER_STICKER",
        "golden gate bridge sticker":                    "GGBRIDGE_PINK_STICKER",
        # City by the Bay stickers — white/black version is STICKER-2
        "white sf city by the bay circle sticker":       "SFCITYBYTHEBAY_CIRCLE_STICKER-2",
        "city by the bay black/white sticker":           "SFCITYBYTHEBAY_CIRCLE_STICKER-2",
        "city by the bay black and white sticker":       "SFCITYBYTHEBAY_CIRCLE_STICKER-2",
        "the city by the bay black/white":               "SFCITYBYTHEBAY_CIRCLE_STICKER-2",
        # Tea Towels — prevent HP SKU misassignment
        "san francisco tea towel":                       "SANFRANCISCO_MAP_DISHTOWEL",
        "sf tea towel":                                  "SANFRANCISCO_MAP_DISHTOWEL",
        # Pencil Pouches — prevent numeric Ricochet SKU misassignment
        "pencil case natural":                           "pp-sf-cn-02",
        "sf icons pencil pouch natural":                 "pp-sf-cn-02",
        "sf blue pouch":                                 "PP-SF-CB-01",
        "sf icons pencil pouch blue":                    "PP-SF-CB-01",
        "pencil case blue":                              "PP-SF-CB-01",
        # Totes
        "block print tote":                              "SF_BLOCKFONT_TOTE",
        "sf block print tote":                           "SF_BLOCKFONT_TOTE",
      # SF City Name Tote — NOT SF_HP_11x15 (HP = hand painted print)
        "san francisco city tote":                       "SF_BLOCKFONT_TOTE",
        "sf city tote":                                  "SF_BLOCKFONT_TOTE",
        "san francisco city name tote":                  "SF_BLOCKFONT_TOTE",
        # Hats
        "sf fog dad hat - light blue":                   "DH-004-LB",
        "sf fog dad hat light blue":                     "DH-004-LB",
        "fog dad hat light blue":                        "DH-004-LB",
        "fog dad hat - light blue":                      "DH-004-LB",
        # Magnets
        "magnets set":                                   "SANFRANCISCOICONS_MAGNETSET",
        "magnet set":                                    "SANFRANCISCOICONS_MAGNETSET",
        # Cards
        "wishing you a sweet birthday cake card":        "SWEETBDAYCAKE_A2_GCARD",
        "sweet birthday cake card":                      "SWEETBDAYCAKE_A2_GCARD",
        # ── 6/18/2026 fix: items that fell back to raw Ricochet numeric SKUs ──
        # These were correctly named but find_sku() missed them; adding explicit
        # lowercase keys ensures they always resolve to the right inventory SKU.
        "acrylic painted ladies magnet":                     "MAG-SF-PLADIES-CL",
        "acrylic painted ladies magnet (acrylic)":           "MAG-SF-PLADIES-CL",
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
        # ── 6/28/2026 fix: new Ricochet item name variants ────────────────────────
        # Revenue fix applied above: changed aged_price → agreed (Agreed = actual sale price).
        # These entries cover new/changed item names in the Ricochet export format.
        "blue pencil pouches":                    "PP-SF-CB-01",
        "natural pencil pouches":                 "pp-sf-cn-02",
        "home sweet magnet":                      "MAGNET_HOMESWEETSF",
        "golden gate retro postcard":             "SFGGBRIDGE_RETRO_PCARD",
        "retro golden gate postcard":             "SFGGBRIDGE_RETRO_PCARD",
        "vntage sfo luggage tag sticker":         "LUGGAGETAG_SFO_STICKER",
        "twists and turns card":                  "TWISTSANDTURNS_GCARD",
        "home sweet san francisco card":          "HOMESWEETSF_A2CARD",
        "fog city dad hat - light blue":          "DH-004-LB",
        "fog city dad hat":                       "DH-004-LB",
        "retro painted ladies travel poster":     "MAG-SF-RETRO-PL",
        "san francisco pill sticker":             "PILL_SF_RED_STICKER",
        "san francisco. pencil case":             "pp-sf-cn-02",
        "sweetest birthday greeting card":        "SWEETBDAYCAKE_A2_GCARD",
        "retrogg bridge travel poster magent":    "MAG-SF-RETRO-GGB",
        "retro gg bridge travel poster magent":   "MAG-SF-RETRO-GGB",
        "sf icons magnet set":                    "SANFRANCISCOICONS_MAGNETSET",
        "sf city name tote":                      "SF_BLOCKFONT_TOTE",
        "blue ca sticker":                        "CALIFORNIASTATE_BLOCKFONT_BLUE",
        "i'd climb any hill card":                "IDCLIMBANYHILL_A2CARD",
        "i'd climb greeting card":                "IDCLIMBANYHILL_A2CARD",
        "blue painted lady":                      "KC-PAINTEDLADY-BLUE",
        "blue painted lady keychain":             "KC-PAINTEDLADY-BLUE",
        "pink painted lady keychain":             "KC-PAINTEDLADY-PINK",
        "painted lady blue keychain":             "KC-PAINTEDLADY-BLUE",
        "luggage tag keychain":                   "KC-SFO-LUGGAGETAG",
        "sfo luggage tag":                        "SFO_LUGGAGETAG_MAGNET",
        "window seat greeting card":              "WINDOWSEAT_A2_GREETINGCARD",
        "pencil case blue sf":                    "PP-SF-CB-01",
        "sf felt lettered dad hat":               "DH-003-NB",
        "sf bridge dad hat - navy":               "DH-001-NB",
        "sf icons greeting card":                 "SFICONS_GREETINGCARD",
        "sf blue icons postcard":                 "SFBLUEICONS_POSTCARD",
        "illustrated ferry building landmark sticker": "FB_ILLUSTRATED_STICKER",
        "8x10 bay area map":                      "BAYAREA_BW_8x10",
        "8x10 golden gate travel print":          "GOLDENGATE_TRAVELPOSTER_8x10",
        "8x10 sf map print":                      "SF_BW_8x10",
        "blue icons postcard":                    "SFBLUEICONS_POSTCARD",
        "cape cod map print 8x10":                "CAPECOD_BW_8x10",
        "austin map print 8x10":                  "AUSTIN_BW_8x10",
        "ohio state university campus map print 8x10": "OHIOSTATE_BW_8x10",
        "purdue university campus map print 8x10": "PURDUE_BW_8x10",
        "jersey city map print 8x10":             "JERSEYCITY_BW_8x10",
        "los angeles map print 8x10":             "LA_BW_8x10",
        "napa valley map print 8x10":             "NAPAVALLEY_BW_8x10",
        "minneapolis map print 8x10":             "MPLS_BW_8x10",
        "pittsburgh map print 8x10":              "PITTSBURGH_BW_8x10",
        "paris map print 8x10":                   "PARIS_BW_8x10",
        "stanford campus map print 8x10":         "STANFORD_BW_8x10",
        "uc berkeley campus map print 8x10":      "UCBERKELEY_CAMPUS_BW_8x10",
        "ucla campus map print 8x10":             "UCLA_BW_8x10",
        "retro golden gate bridge poster magnet":  "MAG-SF-RETRO-GGB",
        "home sweet san francisco art print 11x11": "HOMESWEETSF_11x15",
        "sf tote block letters":                  "SF_BLOCKFONT_TOTE",
        "postcards 3 for $10":                    "postcards3for11",
    }
    # 1. Category-filtered search - only look in the right product type
    for keyword, categories in CATEGORY_HINTS.items():
        if keyword in key:
            cat_pool = {}
            for cat in categories:
                cat_pool.update(by_category.get(cat, {}))
            if cat_pool:
                result = _match(cat_pool, key)
                if result:
                    return result
            break   # keyword matched - don't try other keywords

    # 2. Broad search across all categories
    result = _match(all_lookup, key)
    if result:
        return result

    # 3. Hardcoded overrides - fallback only, checked last
    if key in OVERRIDES:
        return OVERRIDES[key]
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

    # GG Travel Poster PRINTS — must fire before magnet/sticker/card checks
    if re.search(r'(gg|golden.?gate).{0,20}travel.{0,20}poster.{0,5}8x10|8x10.{0,20}(gg|golden.?gate).{0,20}travel', nl):
        return 'Golden Gate Travel Poster 8x10'
    if re.search(r'(gg|golden.?gate).{0,20}travel.{0,20}poster.{0,5}11x14|11x14.{0,20}(gg|golden.?gate).{0,20}travel', nl):
        return 'Golden Gate Travel Poster 11x14'
    if re.search(r'(gg|golden.?gate).{0,20}travel.{0,20}poster.{0,5}12x16', nl):
        return 'Golden Gate Travel Poster 12x16'

    # GG Travel Poster CARD
    if re.search(r'(gg|golden.?gate).{0,15}travel.{0,5}(poster.{0,5})?card', nl):
        return 'Golden Gate Travel Poster Card'

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

    # Ferry Building magnets / sticker — sticker and travel poster checks before retro magnet
    # Size in name (8x10 etc.) = it's a print, not a magnet
    if re.search(r'ferry building.{0,30}(8x10|9x12|11x14|12x16)', nl):
        return 'Ferry Building Travel Poster Print'
    if re.search(r'(retro.{0,10})?ferry building.{0,20}(poster.{0,5})?sticker', nl):
        return 'Retro Ferry Building Poster Sticker'
    if re.search(r'ferry building.{0,20}travel.{0,10}(poster.{0,5})?magnet', nl):
        return 'Ferry Building Travel Poster Magnet'
    if re.search(r'(retro.{0,8})?ferry building.{0,20}(poster|retro).{0,10}magnet', nl):
        return 'Retro Ferry Building Poster Magnet'

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
        if 'card' in nl:   return 'Home Sweet SF Greeting Card'
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
    """Group by canonical item name, sum unit counts and revenue, keep first-seen SKU."""
    groups: dict[str, dict] = defaultdict(lambda: {"item": "", "sku": "", "qty": 0, "revenue": 0.0})
    for r in records:
        raw_name  = r.get("item", "").strip()
        canonical = normalize(raw_name)
        key       = canonical.lower()
        g         = groups[key]
        g["item"] = canonical
        g["qty"]  += 1      # each raw row = 1 unit sold
        if not g["sku"]:
            g["sku"] = r.get("sku", "").strip()
        # Sum aged price (revenue)
        try:
            price_str = r.get("agreed", "0").replace("$", "").replace(",", "").strip()
            g["revenue"] += float(price_str) if price_str else 0.0
        except (ValueError, AttributeError):
            pass

    merged = sorted(groups.values(), key=lambda x: x["item"].lower())
    total_rev = sum(m["revenue"] for m in merged)
    log.info(f"After merge: {len(merged)} unique items, {sum(m['qty'] for m in merged)} units, ${total_rev:.2f} revenue")
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
        new_values.append([month_name, year, m["item"], sku, m["qty"], date_range_label, round(m.get("revenue", 0.0), 2)])
        if needs_review:
            rows_needing_review.append(i)

    if first_row is not None:
        # Replace the existing block in place, padding with blanks if needed
        existing_count = last_row - first_row + 1
        blanks_needed  = max(0, existing_count - len(new_values))
        padded         = new_values + [["", "", "", "", "", "", ""]] * blanks_needed
        range_str      = f"'{FOG_CITY_TAB}'!A{first_row}:G{first_row + len(padded) - 1}"
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
        range_str  = f"'{FOG_CITY_TAB}'!A{append_row}:G{append_row + len(padded) - 1}"
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
                        "endColumnIndex":   7,
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

def _build_category_trends(cat_recent: dict, cat_prev: dict) -> list:
    """
    Compare category sales last 10 days vs prior 20 days (normalised to daily rate).
    Return up to 3 categories with the most drastic change (up or down),
    minimum threshold to filter noise.
    """
    from collections import defaultdict as _dd2
    all_cats = set(list(cat_recent.keys()) + list(cat_prev.keys()))
    results = []
    for cat in all_cats:
        r = cat_recent.get(cat, 0)
        p = cat_prev.get(cat, 0)
        # Normalise to daily rate
        r_daily = r / 10
        p_daily = p / 20
        # Require at least some minimum activity to avoid noise
        if r < 3 and p < 3:
            continue
        if p_daily > 0.01:
            pct = (r_daily - p_daily) / p_daily * 100
        elif r_daily > 0:
            pct = 150  # new activity
        else:
            pct = -100  # dropped to zero
        results.append({
            "category":   cat,
            "qty_recent": round(r),
            "qty_prev":   round(p),
            "pct":        round(pct),
        })

    # Sort by absolute % change, pick top 3 most dramatic
    results.sort(key=lambda x: -abs(x["pct"]))
    return results[:3]


def build_dashboard_json(sheets):
    """
    Read the last 30 days of Fog City Sales data from the sheet,
    compute top 10 items and hot/cold trends, write data.json.
    Aggregates by SKU (source of truth) and maps to display names
    from the Inventory Summary tab so naming is always consistent.
    """
    log.info("Building dashboard JSON…")

    # Display-friendly category groupings (Inventory Summary category → display label)
    CATEGORY_GROUPS = {
        # Stickers — all sticker types in one bucket
        "Stickers":     "Stickers",
        "Sticker":      "Stickers",
        "Sticker Sheet":"Stickers",
        "Sticker Book": "Stickers",
        "Sticker deal": "Stickers",
        # Magnets
        "Magnets":      "Magnets",
        # Postcards — all postcard types in one bucket
        "Postcards":    "Postcards",
        "Postcard deal":"Postcards",
        # Cards — greeting cards and card packs together
        "Greeting Card":"Cards",
        "Card Pack":    "Cards",
        # Map Prints — city maps and campus maps together
        "City Print":   "Map Prints",
        "School Prints":"Map Prints",
        # Other prints kept separate
        "Film Print":   "Film Prints",
        "Landmark":     "Landmark Prints",
        # Other
        "Totes":        "Totes",
        "Tea Towel":    "Tea Towels",
        "Pencil Pouch": "Pencil Pouches",
        "Keychains":    "Keychains",
        "Keychain":     "Keychains",
    }

    # ── Load Inventory Summary: SKU → display name + category ─────────────────
    inv_result = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'Inventory Summary'!A:E",
    ).execute()
    sku_to_name = {}      # SKU → clean display name
    sku_to_category = {}  # SKU → display category label
    for row in inv_result.get("values", [])[1:]:
        if len(row) >= 5:
            raw_cat   = str(row[0]).strip()
            item_name = str(row[1]).strip()
            sku       = str(row[4]).strip()
            if sku and item_name:
                sku_to_name[sku.upper()] = item_name
            if sku and raw_cat in CATEGORY_GROUPS:
                sku_to_category[sku.upper()] = CATEGORY_GROUPS[raw_cat]

    # ── Load Fog City Sales ───────────────────────────────────────────────────
    result = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{FOG_CITY_TAB}'!A:G",            # open-ended so sheet growth never cuts off data
    ).execute()
    all_rows = result.get("values", [])

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

    # Group rows by source label, keyed by SKU
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
        item_name = str(row[2]).strip() if row[2] else ''
        sku_raw   = str(row[3]).strip() if len(row) > 3 and row[3] else ''
        try:
            qty = int(float(row[4])) if row[4] else 0
        except Exception:
            qty = 0
        source = str(row[5]).strip() if row[5] else ''
        if not source or qty <= 0:
            continue

        # Resolve SKU: use what's in the sheet, fall back to normalizing the name
        sku = sku_raw.upper() if sku_raw else find_sku(item_name, {}).upper()
        if not sku:
            sku = normalize(item_name).upper()  # last resort key

        key = f"{year_val}|{source}"
        if key not in source_groups:
            s, e = parse_source_dates(source, year_val)
            source_groups[key] = {'start': s, 'end': e, 'skus': _dd(int), 'rev': _dd(float)}
        source_groups[key]['skus'][sku] += qty
        # Revenue from col G
        try:
            rev_str = str(row[6]).replace("$", "").replace(",", "").strip() if len(row) > 6 else "0"
            source_groups[key]['rev'][sku] += float(rev_str) if rev_str else 0.0
        except (ValueError, AttributeError):
            pass

    # ── Historical revenue baselines (pre-automation) ────────────────────────
    # Covers all revenue earned before nightly col G scraping began (May 20+).
    # The nightly scraper adds to these each night going forward.
    # Last updated: May 20, 2026
    BASELINE_YTD     = 85344.94   # Jan 1 – May 19, 2026
    BASELINE_MONTHLY =  4098.18   # May 1 – May 19, 2026
    BASELINE_WEEKLY  =  2372.68   # May 13 – May 19, 2026
    valid = [(k, g) for k, g in source_groups.items() if g['end']]
    if not valid:
        log.warning("No dated source groups found — skipping dashboard JSON.")
        return
    ref_date = max(g['end'] for _, g in valid)

    # Windows
    window_10_start = ref_date - timedelta(days=9)
    prev_start      = ref_date - timedelta(days=30)
    week_start      = ref_date - timedelta(days=6)   # last 7 days
    month_start     = ref_date.replace(day=1)         # current month
    ytd_start       = ref_date.replace(month=1, day=1) # year to date

    recent_10  = _dd(float)
    prev_20    = _dd(float)
    rev_10     = _dd(float)   # revenue last 10 days
    rev_week   = _dd(float)   # revenue last 7 days
    rev_month  = _dd(float)   # revenue current month
    rev_ytd    = _dd(float)   # revenue year to date
    rev_yesterday = _dd(float)   # revenue for the most recent day (ref_date)
    cat_recent = _dd(float)   # category label → qty last 10 days
    cat_prev   = _dd(float)   # category label → qty prior 20 days

    for _, g in valid:
        s, e = g['start'], g['end']
        span = max((e - s).days + 1, 1)

        def add_window(bucket_qty, bucket_rev, win_s, win_e):
            if e >= win_s and s <= win_e:
                ol_s = max(s, win_s)
                ol_e = min(e, win_e)
                ratio = ((ol_e - ol_s).days + 1) / span
                for sku, qty in g['skus'].items():
                    bucket_qty[sku] += qty * ratio
                for sku, rev in g['rev'].items():
                    bucket_rev[sku] += rev * ratio

        add_window(recent_10, rev_10,    window_10_start, ref_date)
        add_window(_dd(float), rev_week,  week_start,       ref_date)
        add_window(_dd(float), rev_month, month_start,      ref_date)
        add_window(_dd(float), rev_ytd,   ytd_start,        ref_date)
        add_window(_dd(float), rev_yesterday, ref_date, ref_date)

        # Prev 20 days for trend comparison
        if e >= prev_start and s < window_10_start:
            ol_s = max(s, prev_start)
            ol_e = min(e, window_10_start - timedelta(days=1))
            if ol_e >= ol_s:
                ratio = ((ol_e - ol_s).days + 1) / span
                for sku, qty in g['skus'].items():
                    prev_20[sku] += qty * ratio

        # Category aggregation (last 10 days vs prior 20 days)
        if e >= window_10_start and s <= ref_date:
            ol_s = max(s, window_10_start)
            ol_e = min(e, ref_date)
            ratio = ((ol_e - ol_s).days + 1) / span
            for sku, qty in g['skus'].items():
                cat = sku_to_category.get(sku.upper())
                if cat:
                    cat_recent[cat] += qty * ratio

        if e >= prev_start and s < window_10_start:
            ol_s = max(s, prev_start)
            ol_e = min(e, window_10_start - timedelta(days=1))
            if ol_e >= ol_s:
                ratio = ((ol_e - ol_s).days + 1) / span
                for sku, qty in g['skus'].items():
                    cat = sku_to_category.get(sku.upper())
                    if cat:
                        cat_prev[cat] += qty * ratio

    def daily(d, days):
        return {k: v / days for k, v in d.items()}

    rt10 = daily(recent_10, 10)
    rt20 = daily(prev_20,   20)

    def display_name(sku):
        """Return the clean product name for a SKU from Inventory Summary."""
        return sku_to_name.get(sku.upper(), sku)  # fall back to SKU itself if not found

    # Top 10 by recent 10-day qty
    top10 = sorted(recent_10.items(), key=lambda x: -x[1])[:10]

    # Trends
    all_skus = set(list(rt10.keys()) + list(rt20.keys()))
    trends = {}
    for sku in all_skus:
        rr = rt10.get(sku, 0)
        rp = rt20.get(sku, 0)
        qr = recent_10.get(sku, 0)
        qp = prev_20.get(sku, 0)
        if qr < 2 and qp < 2:
            continue
        if rp > 0.01:
            pct = (rr - rp) / rp * 100
        elif rr > 0:
            pct = 150
        else:
            pct = 0
        trends[sku] = {'pct': pct, 'qty_r': qr, 'qty_p': qp}

    hot  = sorted([(k, v) for k, v in trends.items() if v['pct'] >= 50  and v['qty_r'] >= 2],
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
        "revenue": {
            "yesterday":       round(sum(rev_yesterday.values()),                  2),
            "yesterday_label": ref_date.strftime("%-m/%-d"),
            "weekly":          round(sum(rev_week.values())  + BASELINE_WEEKLY,    2),
            "monthly":         round(sum(rev_month.values()) + BASELINE_MONTHLY,   2),
            "ytd":             round(sum(rev_ytd.values())   + BASELINE_YTD,       2),
        },
        "top10": [
            {"name": display_name(sku), "sku": sku, "qty": round(qty)}
            for sku, qty in top10
        ],
        "hot": [
            {
                "name":       display_name(sku),
                "sku":        sku,
                "qty_recent": round(v['qty_r']),
                "qty_prev":   round(v['qty_p']),
                "pct":        round(v['pct']),
            }
            for sku, v in hot
        ],
        "cold": [
            {
                "name":       display_name(sku),
                "sku":        sku,
                "qty_recent": round(v['qty_r']),
                "qty_prev":   round(v['qty_p']),
                "pct":        round(v['pct']),
            }
            for sku, v in cold
        ],
        "category_trends": _build_category_trends(cat_recent, cat_prev),
    }

    with open("data.json", "w") as f:
        json.dump(data, f, indent=2)



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
