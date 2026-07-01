"""
fix_jun28_30.py
One-off backfill: re-scrape Ricochet for 6/28, 6/29, 6/30 and rewrite those rows
with correct inventory SKUs and correct revenue (agreed price).

Background:
  6/28-6/30 were uploaded before the aged_price->agreed fix landed,
  so all three dates have revenue=$0 and raw Ricochet numeric SKUs.

Row ranges (1-indexed, inspected 2026-07-01):
  6/28  rows 11064-11094
  6/29  rows 11095-11126
  6/30  rows 11127-end (detected dynamically)
"""

import os, json, re, logging, sys
import base64 as _b64
from datetime import date

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

# Import shared logic from sibling module
sys.path.insert(0, os.path.dirname(__file__))
from ricochet_sync import (
    get_sheets_service,
    build_sku_lookup,
    find_sku,
    scrape_ricochet,
    filter_by_date,
    merge_rows,
)

BLOCKS = {
    "6/28": {"label": "6/28 - 6/28 ricochet export", "first_row": 11064, "last_row": 11094},
    "6/29": {"label": "6/29 - 6/29 ricochet export", "first_row": 11095, "last_row": 11126},
    "6/30": {"label": "6/30 - 6/30 ricochet export", "first_row": 11127, "last_row": None},
}


def fix_block(sheets, sku_lookup, block_key, all_records):
    block = BLOCKS[block_key]
    label = block["label"]
    first_row = block["first_row"]
    yellow = {"red": 1.0, "green": 0.95, "blue": 0.0}

    import re as _re
    m = _re.search(r"(\d+)/(\d+)", label)
    month, day = int(m.group(1)), int(m.group(2))
    target_date = date(2026, month, day)

    filtered = filter_by_date(all_records, target_date, target_date)
    if not filtered:
        log.warning(f"No Ricochet records for {block_key} -- skipping")
        return

    merged = merge_rows(filtered)
    total_rev = sum(x["revenue"] for x in merged)
    total_qty = sum(x["qty"] for x in merged)
    log.info(f"{block_key}: {len(merged)} items, {total_qty} units, ${total_rev:.2f} revenue")

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
    new_values = []
    rows_needing_review = []
    for i, item in enumerate(merged):
        sku = find_sku(item["item"], sku_lookup)
        if not sku:
            sku = item.get("sku", "")
            rows_needing_review.append(i)
        new_values.append([
            "June", 2026, item["item"], sku,
            item["qty"], label, round(item.get("revenue", 0.0), 2)
        ])

    blanks_needed = max(0, existing_count - len(new_values))
    padded = new_values + [["", "", "", "", "", "", ""]] * blanks_needed
    range_str = f"'{FOG_CITY_TAB}'!A{first_row}:G{first_row + len(padded) - 1}"

    sheets.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_str,
        valueInputOption="USER_ENTERED",
        body={"values": padded},
    ).execute()
    log.info(f"{block_key}: wrote {len(new_values)} rows + {blanks_needed} blanks -> {range_str}")

    if rows_needing_review:
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
        log.warning(f"{block_key}: {len(rows_needing_review)} rows flagged yellow")


def main():
    log.info("=== fix_jun28_30.py starting ===")
    sheets = get_sheets_service()
    sku_lookup = build_sku_lookup(sheets)
    log.info("Scraping Ricochet payout data...")
    all_records = scrape_ricochet()
    log.info(f"Total raw records: {len(all_records)}")
    for key in ["6/28", "6/29", "6/30"]:
        log.info(f"--- Fixing {key} ---")
        fix_block(sheets, sku_lookup, key, all_records)
    log.info("=== Done ===")


if __name__ == "__main__":
    main()
