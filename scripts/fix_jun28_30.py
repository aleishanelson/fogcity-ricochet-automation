"""
fix_jun28_30.py
One-off backfill: fix 6/28, 6/29, 6/30 rows in Fog City Sales with correct
inventory SKUs and estimated revenue based on agreed price lookup table.

Background:
  6/28-6/30 were uploaded with raw Ricochet numeric SKUs and revenue=$0
  because the aged_price->agreed fix had not yet taken effect.
  Ricochet payout history no longer accessible for these dates.
  Revenue is estimated from the agreed price lookup table used at time of sale.

Row ranges (1-indexed, inspected 2026-07-01):
  6/28  rows 11064-11094
  6/29  rows 11095-11126
  6/30  rows 11127-end (detected dynamically)
"""

import os, json, re, logging, sys
import base64 as _b64

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

_sa_raw = os.environ["GOOGLE_SA_JSON"].strip()
try:
    SA_JSON = json.loads(_b64.b64decode(_sa_raw).decode())
except Exception:
    SA_JSON = json.loads(_sa_raw)

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
FOG_CITY_TAB = "Fog City Sales"
SHEET_GID = 1018380031

sys.path.insert(0, os.path.dirname(__file__))
from ricochet_sync import get_sheets_service, build_sku_lookup, find_sku

# Agreed prices per item (based on Ricochet payout table patterns seen in fix_jun22_27)
# Key: lowercase item name substring -> agreed price per unit
PRICE_TABLE = {
    "stickers- 3 for $11": 11.00,
    "stickers 3 for $11": 11.00,
    "postcards 3 for $11": 11.00,
    "postcards- 3 for $11": 11.00,
    "postcards 3 for $10": 10.00,
    "packaged keychains 3 for $25": 25.00,
    "magnet set san francisco": 14.00,
    "sf icons magnet set": 14.00,
    "travel poster 8x10": 25.00,
    "golden gate travel poster 8x10": 25.00,
    "bay area map print 9x12": 25.00,
    "campus map print 8x10": 20.00,
    "ohio state": 20.00,
    "purdue": 20.00,
    "map print 8x10": 20.00,
    "art print 8x8": 20.00,
    "home sweet san francisco art print": 20.00,
    "sf city by the bay dad hat": 34.00,
    "fog city dad hat": 34.00,
    "sf fuzzy patch trucker hat": 34.00,
    "sf bridge dad hat": 34.00,
    "sf fog dad hat": 34.00,
    "sf felt lettered dad hat": 34.00,
    "sf icons tote": 30.00,
    "sf map tote": 30.00,
    "tote": 30.00,
    "acrylic keychain": 10.00,
    "california icon keychain": 10.00,
    "california keychain": 10.00,
    "blue victorian house keychain": 10.00,
    "golden gate icon keychain": 10.00,
    "sfo luggage tag keychain": 13.00,
    "painted lady keychain": 12.00,
    "sf blue pouch": 16.00,
    "pencil pouch": 16.00,
    "pencil case": 16.00,
    "sfo luggage tag acrylic magnet": 8.50,
    "acrylic die cut magnet": 8.50,
    "fishermans wharf acrylic": 8.50,
    "ferry building acrylic": 8.50,
    "golden gate acrylic die cut": 8.50,
    "sf houses acrylic": 8.50,
    "painted ladies acrylic": 8.50,
    "card": 6.50,
    "greeting card": 6.50,
    "magnet": 8.00,
    "postcard": 4.00,
    "sticker": 4.00,
}


def estimate_price(item_name: str, qty: int) -> float:
    """Estimate agreed price from item name, return total for qty units."""
    key = item_name.strip().lower()
    # Check exact matches first (longest first to prefer more specific)
    for pattern, price in sorted(PRICE_TABLE.items(), key=lambda x: -len(x[0])):
        if pattern in key:
            return round(price * qty, 2)
    return 0.0  # unknown - will flag for review


BLOCKS = {
    "6/28": {"label": "6/28 - 6/28 ricochet export", "first_row": 11064, "last_row": 11094},
    "6/29": {"label": "6/29 - 6/29 ricochet export", "first_row": 11095, "last_row": 11126},
    "6/30": {"label": "6/30 - 6/30 ricochet export", "first_row": 11127, "last_row": None},
}


def fix_block(sheets, sku_lookup, block_key):
    block = BLOCKS[block_key]
    label = block["label"]
    first_row = block["first_row"]
    yellow = {"red": 1.0, "green": 0.95, "blue": 0.0}

    last_row = block["last_row"]
    if last_row is None:
        col_f = sheets.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{FOG_CITY_TAB}'!F{first_row}:F{first_row + 200}",
        ).execute().get("values", [])
        last_row = first_row
        for i, r in enumerate(col_f):
            if r and str(r[0]).strip() == label:
                last_row = first_row + i
        block["last_row"] = last_row

    existing_count = last_row - first_row + 1

    # Read existing rows from sheet
    read_range = f"'{FOG_CITY_TAB}'!A{first_row}:G{last_row}"
    existing = sheets.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=read_range,
    ).execute().get("values", [])

    new_values = []
    rows_needing_review = []
    for i, row in enumerate(existing):
        if not row or len(row) < 3:
            new_values.append(["", "", "", "", "", "", ""])
            continue
        item_name = str(row[2]).strip() if len(row) > 2 else ""
        try:
            qty = int(float(str(row[4]).strip())) if len(row) > 4 and row[4] else 1
        except Exception:
            qty = 1

        # Re-resolve SKU
        sku = find_sku(item_name, sku_lookup) if item_name else ""
        if not sku:
            rows_needing_review.append(i)

        # Estimate revenue if missing
        existing_rev = str(row[6]).strip() if len(row) > 6 else ""
        if existing_rev and existing_rev != "0":
            revenue = float(existing_rev.replace("$", "").replace(",", ""))
        else:
            revenue = estimate_price(item_name, qty)
            if revenue == 0.0 and item_name:
                log.warning(f"  No price found for: {item_name!r} (qty={qty})")
                rows_needing_review.append(i)

        new_values.append([
            str(row[0]) if row else "June",
            str(row[1]) if len(row) > 1 else "2026",
            item_name,
            sku,
            qty,
            label,
            round(revenue, 2)
        ])

    total_rev = sum(float(r[6]) for r in new_values if r[6])
    total_qty = sum(int(r[4]) for r in new_values if r[4])
    log.info(f"{block_key}: {len(new_values)} items, {total_qty} units, ${total_rev:.2f} revenue")

    range_str = f"'{FOG_CITY_TAB}'!A{first_row}:G{first_row + len(new_values) - 1}"
    sheets.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_str,
        valueInputOption="USER_ENTERED",
        body={"values": new_values},
    ).execute()
    log.info(f"{block_key}: wrote {len(new_values)} rows -> {range_str}")

    if rows_needing_review:
        rows_needing_review = list(set(rows_needing_review))
        requests = []
        for i in rows_needing_review:
            row_0idx = first_row + i - 1
            requests.append({"repeatCell": {
                "range": {
                    "sheetId": SHEET_GID,
                    "startRowIndex": row_0idx,
                    "endRowIndex": row_0idx + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 7,
                },
                "cell": {"userEnteredFormat": {"backgroundColor": yellow}},
                "fields": "userEnteredFormat.backgroundColor",
            }})
        sheets.batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests},
        ).execute()
        log.warning(f"{block_key}: {len(rows_needing_review)} rows flagged yellow (SKU/price review needed)")


def main():
    log.info("=== fix_jun28_30.py starting ===")
    sheets = get_sheets_service()
    sku_lookup = build_sku_lookup(sheets)
    for key in ["6/28", "6/29", "6/30"]:
        log.info(f"--- Fixing {key} ---")
        fix_block(sheets, sku_lookup, key)
    log.info("=== Done ===")


if __name__ == "__main__":
    main()
