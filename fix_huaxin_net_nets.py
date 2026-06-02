#!/usr/bin/env python3
"""Fix Net-nets formula for 华新水泥财务 tab.

The rename script changed the label but didn't update the formula because
formulas use row numbers (e.g. =D15-E25) not text names.

We need to find the Balance Sheet rows for:
  - Total Current Liabilities (old reference)
  - Total Liabilities (new reference)
Then replace the old row number with the new one in the Net-nets formula row.
"""

import sys
import os
import re
import json
import time

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')

# 建材 spreadsheet
SPREADSHEET_ID = '1GsHpzk06-vPt9kDh6dD1tfAqf84BobWAnbkT6S29VV4'
SHEET_NAME = '华新水泥财务'


def scan_sections(rows):
    sections = {}
    current = None
    next_headers = ('key stats', 'income statement', 'balance sheet',
                    'cash flow', 'supplemental', 'business segments')
    for i, row in enumerate(rows):
        a = row[0].strip().lower() if row and row[0] else ''
        if a in next_headers:
            current = a
            sections[current] = {'items': {}}
        elif current:
            b = row[1].strip() if len(row) > 1 and row[1] else ''
            c = row[2].strip() if len(row) > 2 and row[2] else ''
            val = c if c else b
            if val and val not in ('盈利指标', '同比增速'):
                sections[current]['items'][val.lower()] = i
    return sections


def _api_with_retry(func, max_retries=5):
    for attempt in range(max_retries):
        try:
            return func()
        except HttpError as e:
            if e.resp.status == 429 and attempt < max_retries - 1:
                wait = 60 * (attempt + 1)
                print(f"  429, retrying in {wait}s...")
                time.sleep(wait)
            else:
                raise


def main():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    service = build('sheets', 'v4', credentials=creds)

    # Get sheet ID
    result = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        fields='sheets(properties(title,sheetId))'
    ).execute()
    sheet_id = None
    for s in result.get('sheets', []):
        if s['properties']['title'] == SHEET_NAME:
            sheet_id = s['properties']['sheetId']
            break
    if sheet_id is None:
        print(f"Sheet {SHEET_NAME} not found!")
        return

    # 1. Read A1:C300 for sections
    data = _api_with_retry(lambda: service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f'{SHEET_NAME}!A1:C300'
    ).execute())
    rows = data.get('values', [])
    sections = scan_sections(rows)

    bs_items = sections.get('balance sheet', {}).get('items', {})
    ks_items = sections.get('key stats', {}).get('items', {})

    print(f"Balance Sheet items found: {len(bs_items)}")
    print(f"Key Stats items found: {len(ks_items)}")

    # Find BS rows
    tcl_row = bs_items.get('total current liabilities')
    tl_row = bs_items.get('total liabilities')
    nets_row = ks_items.get('net-nets')

    print(f"\nTotal Current Liabilities row (0-based): {tcl_row}, 1-based: {tcl_row+1}")
    print(f"Total Liabilities row (0-based): {tl_row}, 1-based: {tl_row+1}")
    print(f"Net-nets row (0-based): {nets_row}, 1-based: {nets_row+1}")

    # Print nearby BS items for verification
    print("\nNearby BS items:")
    for name, idx in sorted(bs_items.items(), key=lambda x: x[1]):
        if abs(idx - tcl_row) < 5 or abs(idx - tl_row) < 5:
            print(f"  row {idx+1}: {name}")

    if tcl_row is None or tl_row is None or nets_row is None:
        print("Missing required items!")
        return

    # 2. Read grid data for Net-nets row formulas (columns D through GR)
    data = _api_with_retry(lambda: service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        ranges=[f'{SHEET_NAME}!D{nets_row+1}:GR{nets_row+1}'],
        includeGridData=True,
        fields='sheets.data.rowData.values(userEnteredValue),sheets.data.startRow,sheets.data.startColumn'
    ).execute())

    d = data['sheets'][0]['data'][0]
    rd = d.get('rowData', [])
    start_col = d.get('startColumn', 0)
    start_row = d.get('startRow', 0)

    print(f"\nGrid data: startRow={start_row}, startCol={start_col}")

    if not rd:
        print("No row data found!")
        return

    cells = rd[0].get('values', [])
    print(f"Cells in row: {len(cells)}")

    # Debug: print first few formulas
    for col_offset, cell in enumerate(cells[:5]):
        uev = cell.get('userEnteredValue', {})
        formula = uev.get('formulaValue', '')
        if formula:
            actual_col = start_col + col_offset
            print(f"  Col {actual_col} (offset {col_offset}): {formula}")

    # The formula is: =Total Current Assets row - Total Current Liabilities row
    # We need to find the second reference (the subtraction) and replace TCL row with TL row
    # Pattern: a letter followed by the TCL row number
    # scan_sections returns 0-based indices, but GS formulas use 1-based row numbers
    tcl_row_1based = tcl_row + 1
    tl_row_1based = tl_row + 1

    pattern = re.compile(rf'(?<=[A-Z]){tcl_row_1based}(?!\d)')

    updates = []
    for col_offset, cell in enumerate(cells):
        uev = cell.get('userEnteredValue', {})
        formula = uev.get('formulaValue', '')
        if formula and pattern.search(formula):
            new_formula = pattern.sub(str(tl_row_1based), formula)
            if new_formula != formula:
                actual_col = start_col + col_offset
                updates.append((nets_row, actual_col, new_formula))
                print(f"  Col {actual_col}: {formula} -> {new_formula}")

    print(f"\nTotal formulas to update: {len(updates)}")

    if not updates:
        print("No formulas need updating.")
        return

    # 3. Batch update
    requests = []
    for row_idx, col_idx, formula in updates:
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': row_idx,
                    'endRowIndex': row_idx + 1,
                    'startColumnIndex': col_idx,
                    'endColumnIndex': col_idx + 1,
                },
                'cell': {
                    'userEnteredValue': {'formulaValue': formula},
                },
                'fields': 'userEnteredValue',
            }
        })

    try:
        _api_with_retry(lambda: service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={'requests': requests}
        ).execute())
        print(f"Successfully updated {len(updates)} formulas!")
    except HttpError as e:
        print(f"FAILED: {e}")


if __name__ == '__main__':
    main()
