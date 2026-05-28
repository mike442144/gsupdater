#!/usr/bin/env python3
"""
Add YoY growth sub-section to existing Key Stats in a Google Sheets tab.

Steps:
  1. Idempotency check (skip if sub-sections exist)
  2. Insert 1 row after "Key Stats" header (for "盈利指标" sub-header)
  3. Write "盈利指标" in B column of new row
  4. Copy Key Stats items from B to C, clear B
  5. Insert rows for YoY section, write items, fill formulas

Usage:
    python add_yoy_section.py --sheet "今世缘财务" --spreadsheet-id <id>
    python add_yoy_section.py --sheet "今世缘财务" --spreadsheet-id <id> --dry-run
"""

import sys
import os
import re
import argparse

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')

SUB_SECTION_METRICS = "盈利指标"
SUB_SECTION_YOY = "同比增速"
SUB_SECTIONS = (SUB_SECTION_METRICS, SUB_SECTION_YOY)

YOY_ITEMS = [
    ("Revenue YoY", "=__C__{Total Revenue}/__PC__{Total Revenue}-1"),
    ("Gross Profit YoY", "=__C__{Gross Profit}/__PC__{Gross Profit}-1"),
    ("Operating Income YoY", "=__C__{Operating Income}/__PC__{Operating Income}-1"),
    ("Net Income YoY", "=__C__{Net Income to Company}/__PC__{Net Income to Company}-1"),
    ("扣非净利润 YoY", "=__C__{扣非净利润}/__PC__{扣非净利润}-1"),
]


def col_to_letter(col_idx):
    result = ''
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_service():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    return build('sheets', 'v4', credentials=creds)


def get_sheet_properties(service, spreadsheet_id, sheet_name):
    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId,gridProperties(rowCount,columnCount)))'
    ).execute()
    for s in result.get('sheets', []):
        if s['properties']['title'] == sheet_name:
            return s['properties']
    return None


def find_key_stats_section(service, spreadsheet_id, sheet_name):
    """Find Key Stats section. Returns (ks_start, ks_end, items).

    items: list of (row_0idx, item_name) — from column C (new) or B (old).
    """
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!A1:C100"
    ).execute()
    rows = result.get('values', [])

    ks_start = None
    ks_end = len(rows)
    items = []

    for i, row in enumerate(rows):
        a = row[0].strip().lower() if row and row[0] else ''
        b = row[1].strip() if len(row) > 1 and row[1] else ''
        c = row[2].strip() if len(row) > 2 and row[2] else ''

        if a == 'key stats':
            ks_start = i
        elif ks_start is not None and i > ks_start:
            if a:
                ks_end = i
                break
            item_val = c if c else b
            if item_val and item_val not in SUB_SECTIONS:
                items.append((i, item_val))

    return ks_start, ks_end, items, rows


def find_data_columns(service, spreadsheet_id, sheet_name):
    props = get_sheet_properties(service, spreadsheet_id, sheet_name)
    col_count = props.get('gridProperties', {}).get('columnCount', 200) if props else 200
    end_col = col_to_letter(min(col_count - 1, 200))

    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[f"'{sheet_name}'!A1:{end_col}1"],
        includeGridData=True
    ).execute()

    header_vals = result.get('sheets', [{}])[0].get('data', [{}])[0].get('rowData', [{}])[0].get('values', [])

    data_cols = []
    for j in range(3, len(header_vals)):
        cell = header_vals[j]
        uev = cell.get('userEnteredValue', {})
        fv = cell.get('formattedValue', '')
        text = uev.get('stringValue', fv) or uev.get('numberValue', 0) or fv
        if text and re.search(r'\d{4}', str(text)):
            data_cols.append(j)

    return data_cols


def scan_item_to_row(service, spreadsheet_id, sheet_name):
    """Scan B+C to build item_name → row mapping. C takes priority."""
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!B1:C300"
    ).execute()
    rows = result.get('values', [])
    mapping = {}
    for i, row in enumerate(rows):
        b_val = row[0].strip() if row and row[0] else ''
        c_val = row[1].strip() if len(row) > 1 and row[1] else ''
        val = c_val if c_val else b_val
        if val and val not in SUB_SECTIONS:
            mapping[val.lower()] = i
    return mapping


def resolve_formula(template_str, col_letter, prev_col_letter, item_to_row):
    formula = template_str.replace('__C__', col_letter).replace('__PC__', prev_col_letter)

    def replace_item(match):
        item_name = match.group(1)
        row = item_to_row.get(item_name.lower())
        if row is None:
            return '1'
        return str(row + 1)

    formula = re.sub(r'\{([^}]+)\}', replace_item, formula)
    return formula


def add_yoy_section(service, spreadsheet_id, sheet_name, dry_run=False):
    print(f"\n{'='*60}")
    print(f"Adding YoY section to: {sheet_name}")
    print(f"Spreadsheet: {spreadsheet_id[:30]}...")
    print(f"{'='*60}")

    props = get_sheet_properties(service, spreadsheet_id, sheet_name)
    if not props:
        print(f"ERROR: Sheet '{sheet_name}' not found")
        return False

    sheet_id = props['sheetId']

    # ── Step 1: Idempotency check ──
    ks_start, ks_end, items, rows = find_key_stats_section(service, spreadsheet_id, sheet_name)

    if ks_start is None:
        print("ERROR: Key Stats section not found")
        return False

    for _, row in enumerate(rows[ks_start:ks_end]):
        b_val = row[1].strip() if len(row) > 1 and row[1] else ''
        if b_val in SUB_SECTIONS:
            print(f"  Already has sub-sections ('{b_val}'), skipping")
            return True

    if not items:
        print("  ERROR: No items found in Key Stats section")
        return False

    N = len(items)
    print(f"  Key Stats: rows {ks_start+1}-{ks_end} (1-indexed), {N} items")

    if dry_run:
        print(f"  [DRY RUN] Step 2: Insert 1 row after 'Key Stats' (row {ks_start+1})")
        print(f"  [DRY RUN] Step 3: Write '{SUB_SECTION_METRICS}' in B{ks_start+2}")
        print(f"  [DRY RUN] Step 4: Copy {N} items B→C, clear B")
        print(f"  [DRY RUN] Step 5: Insert {1+len(YOY_ITEMS)} rows for YoY, write formulas")
        return True

    # ── Step 2: Insert 1 row after "Key Stats" header ──
    # This creates a blank row for the "盈利指标" sub-header.
    # Existing items (and their D+ data) shift down by 1 row — safe.
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': [{
            'insertDimension': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': ks_start + 1,
                    'endIndex': ks_start + 2,
                },
                'inheritFromBefore': False,
            }
        }]}
    ).execute()
    print(f"  ✓ Step 2: Inserted 1 row after 'Key Stats'")

    # After insertion, old items shifted from rows ks_start+1..ks_start+N
    # to rows ks_start+2..ks_start+1+N (0-indexed)
    shifted_items_start = ks_start + 2  # first item row (0-indexed)

    # ── Step 3: Write "盈利指标" sub-header in B ──
    sub_header_row = ks_start + 1  # 0-indexed, the newly inserted row
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': [{
            'updateCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': sub_header_row,
                    'endRowIndex': sub_header_row + 1,
                    'startColumnIndex': 1,
                    'endColumnIndex': 2,
                },
                'rows': [{'values': [{
                    'userEnteredValue': {'stringValue': SUB_SECTION_METRICS},
                    'userEnteredFormat': {'textFormat': {'bold': True}}
                }]}],
                'fields': 'userEnteredValue,userEnteredFormat',
            }
        }]}
    ).execute()
    print(f"  ✓ Step 3: Wrote '{SUB_SECTION_METRICS}' in B{sub_header_row+1}")

    # ── Step 4: Copy items B→C, clear B ──
    requests = []

    # Clear B for all shifted item rows
    for idx in range(N):
        row_0idx = shifted_items_start + idx
        requests.append({
            'updateCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': row_0idx,
                    'endRowIndex': row_0idx + 1,
                    'startColumnIndex': 1,
                    'endColumnIndex': 2,
                },
                'rows': [{'values': [{}]}],
                'fields': 'userEnteredValue',
            }
        })

    # Write item names to C column (same rows — D+ data preserved)
    for idx, (_, item_name) in enumerate(items):
        row_0idx = shifted_items_start + idx
        requests.append({
            'updateCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': row_0idx,
                    'endRowIndex': row_0idx + 1,
                    'startColumnIndex': 2,
                    'endColumnIndex': 3,
                },
                'rows': [{'values': [{'userEnteredValue': {'stringValue': item_name}}]}],
                'fields': 'userEnteredValue',
            }
        })

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': requests}
    ).execute()
    print(f"  ✓ Step 4: Copied {N} items B→C, cleared B")

    # ── Step 5: YoY section expansion + formulas ──
    # After steps 2-4, layout is:
    #   ks_start:     "Key Stats" (A)
    #   ks_start+1:   "盈利指标" (B)
    #   ks_start+2 ~ ks_start+1+N: items in C (data in D+)
    #   ks_start+2+N: blank separator (old ks_end row, shifted)
    #   ks_start+3+N: next section (IS, shifted)

    # Insert 1 (blank) + 1 (同比增速) + 5 (YoY items) = 7 rows
    # at position ks_start+2+N (replacing the old blank separator)
    yoy_insert_at = ks_start + 2 + N
    yoy_insert_count = 1 + 1 + len(YOY_ITEMS)  # blank + sub-header + items

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': [{
            'insertDimension': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': yoy_insert_at,
                    'endIndex': yoy_insert_at + yoy_insert_count,
                },
                'inheritFromBefore': False,
            }
        }]}
    ).execute()
    print(f"  ✓ Step 5a: Inserted {yoy_insert_count} rows for YoY at row {yoy_insert_at+1}")

    # Write YoY section: clear inherited junk, then write content
    yoy_requests = []

    # Clear B+C for all new rows
    for row_0idx in range(yoy_insert_at, yoy_insert_at + yoy_insert_count):
        yoy_requests.append({
            'updateCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': row_0idx,
                    'endRowIndex': row_0idx + 1,
                    'startColumnIndex': 1,
                    'endColumnIndex': 3,
                },
                'rows': [{'values': [{}, {}]}],
                'fields': 'userEnteredValue',
            }
        })

    # "同比增速" sub-header in B (after 1 blank row)
    yoy_sub_row = yoy_insert_at + 1
    yoy_requests.append({
        'updateCells': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': yoy_sub_row,
                'endRowIndex': yoy_sub_row + 1,
                'startColumnIndex': 1,
                'endColumnIndex': 2,
            },
            'rows': [{'values': [{
                'userEnteredValue': {'stringValue': SUB_SECTION_YOY},
                'userEnteredFormat': {'textFormat': {'bold': True}}
            }]}],
            'fields': 'userEnteredValue,userEnteredFormat',
        }
    })

    # YoY items in C
    for idx, (item_name, _) in enumerate(YOY_ITEMS):
        row_0idx = yoy_sub_row + 1 + idx
        yoy_requests.append({
            'updateCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': row_0idx,
                    'endRowIndex': row_0idx + 1,
                    'startColumnIndex': 2,
                    'endColumnIndex': 3,
                },
                'rows': [{'values': [{'userEnteredValue': {'stringValue': item_name}}]}],
                'fields': 'userEnteredValue',
            }
        })

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': yoy_requests}
    ).execute()
    print(f"  ✓ Step 5b: Wrote '{SUB_SECTION_YOY}' + {len(YOY_ITEMS)} YoY items")

    # Write YoY formulas
    item_to_row = scan_item_to_row(service, spreadsheet_id, sheet_name)
    data_cols = find_data_columns(service, spreadsheet_id, sheet_name)

    if not data_cols:
        print("  WARNING: No data columns found, skipping YoY formulas")
        return True

    formula_requests = []
    for item_name, template in YOY_ITEMS:
        item_row_0idx = item_to_row.get(item_name.lower())
        if item_row_0idx is None:
            print(f"  WARNING: Item '{item_name}' not found in sheet")
            continue

        for ci, data_col in enumerate(data_cols):
            col_letter = col_to_letter(data_col)
            pc = col_to_letter(max(0, data_cols[ci - 1])) if ci > 0 else col_to_letter(max(0, data_col - 1))

            formula = resolve_formula(template, col_letter, pc, item_to_row)
            cell_value = {
                'userEnteredValue': {'formulaValue': formula},
                'userEnteredFormat': {'numberFormat': {'type': 'PERCENT', 'pattern': '0.0%'}},
            }
            formula_requests.append({
                'updateCells': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': item_row_0idx,
                        'endRowIndex': item_row_0idx + 1,
                        'startColumnIndex': data_col,
                        'endColumnIndex': data_col + 1,
                    },
                    'rows': [{'values': [cell_value]}],
                    'fields': 'userEnteredValue,userEnteredFormat',
                }
            })

    if formula_requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': formula_requests}
        ).execute()
        print(f"  ✓ Step 5c: Wrote {len(formula_requests)} YoY formulas")

    print(f"\nDone! Added YoY section to '{sheet_name}'")
    return True


def main():
    parser = argparse.ArgumentParser(description='Add YoY sub-section to existing Key Stats')
    parser.add_argument('--sheet', required=True, help='Sheet name (e.g. "今世缘财务")')
    parser.add_argument('--spreadsheet-id', required=True, help='Google Spreadsheet ID')
    parser.add_argument('--dry-run', action='store_true', help='Preview changes without writing')
    args = parser.parse_args()

    service = get_service()
    success = add_yoy_section(service, args.spreadsheet_id, args.sheet, dry_run=args.dry_run)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
