"""
Read all 2026 Fog City Sales rows, cross-reference with Inventory Summary,
identify correct vs wrong SKUs, output a JSON report and write corrections
back to the sheet (highlighting uncertain ones in yellow).
"""
import os, json, re, base64 as _b64, logging

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

_sa_raw = os.environ["GOOGLE_SA_JSON"].strip()
try:
    SA_JSON = json.loads(_b64.b64decode(_sa_raw).decode())
except Exception:
    SA_JSON = json.loads(_sa_raw)
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
FOG_CITY_TAB   = "Fog City Sales"

def get_sheets():
    creds = service_account.Credentials.from_service_account_info(
        SA_JSON, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return build("sheets", "v4", credentials=creds).spreadsheets()

# ── Hardcoded overrides ───────────────────────────────────────────────────────
OVERRIDES = {
    "alcatraz island postcard":                     "ALCATRAZ_RETRO_PC",
    "fishermans wharf acrylic die cut magnet":      "MAG-AC-SF-FW",
    "fishermans wharf sf postcard":                 "FISHERMANSWHARF_RETRO_PC",
    "fisherman's wharf sticker":                    "FW_ILLUSTRATED_STICKER",
    "golden gate acrylic die cut magnet":           "MAG-AC-SF-GGB",
    "golden gate bridge acrylic":                   "MAG-AC-SF-GGB",
    "golden gate bridge sticker (pink)":            "GGBRIDGE_PINK_STICKER",
    "home sweet sf magnet":                         "MAGNET_HOMESWEETSF",
    "i love you more than a":                       "LOVEYOUMORETHANSUNNYSF_A2CARD",
    "retro gg travel poster magnet":                "MAG-SF-RETRO-GGB",
    "retrogg bridge travel poster magent":          "MAG-SF-RETRO-GGB",
    "santa clara university campus map print 8x10": "SCU_BW_8x10_CURSIVE",
    "sf house acrylic magnet":                      "MAG-AC-SF-HOUSES",
    "sf icon tote":                                 "SFICONS_TOTE",
    "sf landmark magnet":                           "MAG-SF-LDMKS",
    "stanford campus map print 8x10":              "STANFORD_BW_8x10",
    "twist and turn card":                          "TWISTSANDTURNS_GCARD",
    "window seat card":                             "WINDOWSEAT_A2_GREETINGCARD",
}

def find_sku(name, inv_lookup):
    key = name.strip().lower()
    if key in inv_lookup:        return inv_lookup[key], "inventory_exact"
    for k, v in inv_lookup.items():
        if key in k or k in key: return v, "inventory_partial"
    if key in OVERRIDES:         return OVERRIDES[key], "override_exact"
    for k, v in OVERRIDES.items():
        if key in k or k in key: return v, "override_partial"
    return None, "no_match"

def is_ricochet_sku(sku):
    """Ricochet SKUs look like 0030T3, 00304G, 3027, 3042 etc."""
    return bool(re.match(r'^0{2,}\d', sku) or re.match(r'^\d{4}$', sku))

def main():
    sheets = get_sheets()

    # Load Inventory Summary lookup
    inv_result = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'Inventory Summary'!B:E"
    ).execute()
    inv_lookup = {}
    for row in inv_result.get("values", [])[1:]:
        if len(row) >= 4 and row[0].strip() and row[3].strip():
            inv_lookup[row[0].strip().lower()] = row[3].strip()
    log.info(f"Inventory lookup: {len(inv_lookup)} entries")

    # Load all Fog City Sales
    fog_result = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{FOG_CITY_TAB}'!A1:F15000"
    ).execute()
    all_rows = fog_result.get("values", [])
    log.info(f"Total Fog City Sales rows: {len(all_rows)}")

    # Find all 2026 rows
    updates  = []   # (row_index_1based, new_sku, confidence)
    unknowns = []   # items with no SKU match
    highlight_rows = []

    for i, row in enumerate(all_rows):
        if i == 0: continue
        if len(row) < 3: continue
        year = str(row[1]).strip() if len(row) > 1 else ""
        if year != "2026": continue

        item    = str(row[2]).strip() if len(row) > 2 else ""
        cur_sku = str(row[3]).strip() if len(row) > 3 else ""
        if not item: continue

        sheet_row = i + 1  # 1-based
        new_sku, method = find_sku(item, inv_lookup)

        if new_sku:
            if cur_sku != new_sku:
                updates.append({
                    "row": sheet_row,
                    "item": item,
                    "old_sku": cur_sku,
                    "new_sku": new_sku,
                    "method": method,
                    "highlight": method in ("inventory_partial", "override_partial")
                })
                if method in ("inventory_partial", "override_partial"):
                    highlight_rows.append(sheet_row)
        else:
            # No match at all — note it and highlight
            if cur_sku and not is_ricochet_sku(cur_sku):
                # SKU looks real already, keep it
                log.info(f"  KEEP existing: {item} → {cur_sku}")
            else:
                unknowns.append({"row": sheet_row, "item": item, "cur_sku": cur_sku})
                highlight_rows.append(sheet_row)

    log.info(f"Updates needed: {len(updates)}")
    log.info(f"No match found: {len(unknowns)}")
    log.info(f"Rows to highlight: {len(highlight_rows)}")

    # ── Apply SKU corrections ─────────────────────────────────────────────────
    batch_data = []
    for u in updates:
        batch_data.append({
            "range": f"'{FOG_CITY_TAB}'!D{u['row']}",
            "values": [[u["new_sku"]]]
        })

    if batch_data:
        sheets.values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "USER_ENTERED", "data": batch_data}
        ).execute()
        log.info(f"✅ Applied {len(batch_data)} SKU corrections")

    # ── Highlight uncertain rows in yellow ────────────────────────────────────
    if highlight_rows:
        # Get the sheet id for Fog City Sales (gid=1018380031)
        sheet_id = 1018380031
        requests = []
        for row in highlight_rows:
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row - 1,
                        "endRowIndex": row,
                        "startColumnIndex": 2,  # col C (item)
                        "endColumnIndex": 4     # col D (sku)
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 1.0, "green": 0.95, "blue": 0.4}
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })
        sheets.batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests}
        ).execute()
        log.info(f"✅ Highlighted {len(highlight_rows)} uncertain rows in yellow")

    # ── Print report ──────────────────────────────────────────────────────────
    print("\n=== SKU CORRECTION REPORT ===")
    print(f"\n✅ CORRECTED ({len(updates)} rows):")
    for u in sorted(updates, key=lambda x: x["item"]):
        flag = " ⚠️  HIGHLIGHTED" if u["highlight"] else ""
        print(f"  Row {u['row']:5d}  {u['item'][:45]:45s}  {u['old_sku']:25s} → {u['new_sku']}{flag}")

    print(f"\n❓ NO MATCH FOUND ({len(unknowns)} rows) — highlighted in yellow:")
    for u in sorted(unknowns, key=lambda x: x["item"]):
        print(f"  Row {u['row']:5d}  {u['item'][:45]:45s}  current: {u['cur_sku']}")

    # Save report
    report = {"corrections": updates, "unknowns": unknowns}
    with open("sku_report.json", "w") as f:
        json.dump(report, f, indent=2)
    log.info("Report saved to sku_report.json")

if __name__ == "__main__":
    main()
