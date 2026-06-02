#!/usr/bin/env python3
"""Fix Net Income and Net-nets formula issues for 特宝生物财务 tab.

1. Net Income bug: Key Stats formulas reference 'Net Income to Company' IS row
   instead of 'Net Income' IS row. Fix by replacing old row number with new.
2. Net-nets bug: Formula uses Total Current Liabilities instead of Total Liabilities.
   Fix by replacing TCL row number with TL row number.
"""

import sys
import os
import re
import time

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')
SPREADSHEET_ID = '1CHoPtKocyOtFCi5o2m3rHvYtNdv4JGTu9M2wvTPYpBg'
SHEET_NAME = '三花智控财务'


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
        fields='sheets.properties(title,sheetId)'
    ).execute()
    sheet_id = None
    for s in result.get('sheets', []):
        if s['properties']['title'] == SHEET_NAME:
            sheet_id = s['properties']['sheetId']
            break
    if sheet_id is None:
        print(f"Sheet {SHEET_NAME} not found!")
        return

    # Read A1:C300 for sections
    data = _api_with_retry(lambda: service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f'{SHEET_NAME}!A1:C300'
    ).execute())
    rows = data.get('values', [])
    sections = scan_sections(rows)

    is_items = sections.get('income statement', {}).get('items', {})
    bs_items = sections.get('balance sheet', {}).get('items', {})
    ks_items = sections.get('key stats', {}).get('items', {})

    # === Net Income fix ===
    nic_row = is_items.get('net income to company')
    ni_row = is_items.get('net income')
    # KS items that use Net Income formula
    ni_ks_items = ['net income', 'net margin', 'roe',
                   'interest and rental exp coverage ratio', 'net income yoy']

    print("=== Net Income fix ===")
    print(f"Net Income to Company row (0-based): {nic_row}, 1-based: {(nic_row or -1)+1}")
    print(f"Net Income row (0-based): {ni_row}, 1-based: {(ni_row or -1)+1}")

    # === Net-nets fix ===
    tcl_row = bs_items.get('total current liabilities')
    tl_row = bs_items.get('total liabilities')
    nets_row = ks_items.get('net-nets')

    print("\n=== Net-nets fix ===")
    print(f"Total Current Liabilities row (0-based): {tcl_row}, 1-based: {(tcl_row or -1)+1}")
    print(f"Total Liabilities row (0-based): {tl_row}, 1-based: {(tl_row or -1)+1}")
    print(f"Net-nets row (0-based): {nets_row}, 1-based: {(nets_row or -1)+1}")

    all_updates = []  # (row_0based, col_0based, new_formula)

    # --- Read grid data for all affected rows at once ---
    # Collect all rows we need to check
    rows_to_check = set()
    if nic_row is not None and ni_row is not None:
        for item in ni_ks_items:
            if item in ks_items:
                rows_to_check.add(ks_items[item])
    if tcl_row is not None and tl_row is not None and nets_row is not None:
        rows_to_check.add(nets_row)

    if not rows_to_check:
        print("\nNo rows need checking!")
        return

    ks_start = min(rows_to_check)
    ks_end = max(rows_to_check)

    data = _api_with_retry(lambda: service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        ranges=[f'{SHEET_NAME}!D{ks_start+1}:GR{ks_end+1}'],
        includeGridData=True,
        fields='sheets.data.rowData.values(userEnteredValue),sheets.data.startRow,sheets.data.startColumn'
    ).execute())

    d = data['sheets'][0]['data'][0]
    rd = d.get('rowData', [])
    start_col = d.get('startColumn', 0)
    start_row = d.get('startRow', 0)

    print(f"\nGrid data: startRow={start_row}, startCol={start_col}, rows={len(rd)}")

    # Build patterns for replacement
    patterns = {}
    if nic_row is not None and ni_row is not None:
        patterns[nic_row + 1] = ni_row + 1  # Net Income -> Net Income
    if tcl_row is not None and tl_row is not None:
        patterns[tcl_row + 1] = tl_row + 1  # Net-nets

    for row_0based in rows_to_check:
        data_row_idx = row_0based - start_row
        if data_row_idx < 0 or data_row_idx >= len(rd):
            continue

        cells = rd[data_row_idx].get('values', [])
        for col_offset, cell in enumerate(cells):
            uev = cell.get('userEnteredValue', {})
            formula = uev.get('formulaValue', '')
            if not formula:
                continue

            new_formula = formula
            for old_row, new_row in patterns.items():
                pat = re.compile(rf'(?<=[A-Z]){old_row}(?!\d)')
                new_formula = pat.sub(str(new_row), new_formula)

            if new_formula != formula:
                actual_col = start_col + col_offset
                all_updates.append((row_0based, actual_col, new_formula))
                if len(all_updates) <= 5:
                    print(f"  Row {row_0based+1} Col {actual_col}: {formula} -> {new_formula}")

    print(f"\nTotal formulas to update: {len(all_updates)}")

    if not all_updates:
        print("No formulas need updating.")
        return

    # Batch update
    requests = []
    for row_idx, col_idx, formula in all_updates:
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
        print(f"Successfully updated {len(all_updates)} formulas!")
    except HttpError as e:
        print(f"FAILED: {e}")


if __name__ == '__main__':
    main()
