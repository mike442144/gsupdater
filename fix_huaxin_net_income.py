#!/usr/bin/env python3
"""Fix Net Income formulas in all company tabs of建材 spreadsheet.

The Key Stats items (Net Income, Net Margin, ROE, Net Income YoY) reference
IS row numbers directly. When both 'Net Income to Company' and 'Net Income'
exist in IS, the formulas wrongly point to NIC row instead of NI row.

Fix: replace old_row (NIC) with new_row (NI) in all formula cells on those KS rows.
"""

import sys
import os
import re
import time

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

creds = Credentials.from_authorized_user_file(os.path.expanduser('~/.hermes/google_token.json'))
service = build('sheets', 'v4', credentials=creds)

SID = '1GsHpzk06-vPt9kDh6dD1tfAqf84BobWAnbkT6S29VV4'

# Get sheet list
result = service.spreadsheets().get(
    spreadsheetId=SID,
    fields='sheets(properties(title,sheetId))'
).execute()

company_sheets = {}
for s in result.get('sheets', []):
    t = s['properties']['title']
    if '财务' in t and t not in ('Summary', 'Template'):
        company_sheets[t] = s['properties']['sheetId']

print(f'Found {len(company_sheets)} company tabs')

# KS items whose formulas reference Net Income
ITEMS_TO_FIX = ['net income', 'net margin', 'roe',
                'interest and rental exp coverage ratio', 'net income yoy']

total_fixed = 0
for sheet_name, sheet_id in company_sheets.items():
    # 1. Find sections
    data = service.spreadsheets().values().get(
        spreadsheetId=SID, range=f'{sheet_name}!A1:C300'
    ).execute()
    rows = data.get('values', [])

    is_items = {}
    ks_items = {}
    current = None
    for i, row in enumerate(rows):
        a = row[0].strip().lower() if row and row[0] else ''
        if a == 'income statement':
            current = 'is'
            continue
        if a == 'key stats':
            current = 'ks'
            continue
        if a in ('balance sheet', 'cash flow'):
            current = None
        if current:
            b = row[1].strip() if len(row) > 1 and row[1] else ''
            c = row[2].strip() if len(row) > 2 and row[2] else ''
            val = c if c else b
            if val and val not in ('盈利指标', '同比增速'):
                if current == 'is':
                    is_items[val.lower()] = i
                elif current == 'ks':
                    ks_items[val.lower()] = i

    # Only fix if IS has BOTH items
    nic_lower = 'net income to company'
    ni_lower = 'net income'
    if nic_lower not in is_items or ni_lower not in is_items:
        continue

    old_row = is_items[nic_lower] + 1  # 1-based row number
    new_row = is_items[ni_lower] + 1
    print(f'{sheet_name}: NIC=row{old_row} -> NI=row{new_row}')

    # 2. For each KS item, find formula cells referencing old_row and fix them
    updates = []
    for item_lower in ITEMS_TO_FIX:
        row_idx = ks_items.get(item_lower)
        if row_idx is None:
            continue

        # Read grid data for this row, columns D to GR
        data = service.spreadsheets().get(
            spreadsheetId=SID,
            ranges=[f'{sheet_name}!D{row_idx+1}:GR{row_idx+1}'],
            includeGridData=True,
            fields='sheets.data.rowData.values(userEnteredValue),sheets.data.startRow,sheets.data.startColumn'
        ).execute()
        d = data['sheets'][0]['data'][0]
        rd = d.get('rowData', [])
        if not rd:
            continue

        start_col = d.get('startColumn', 0)
        cells = rd[0].get('values', [])

        for col_offset, cell in enumerate(cells):
            uev = cell.get('userEnteredValue', {})
            formula = uev.get('formulaValue', '')
            if not formula:
                continue

            # Replace old row number with new row number in cell references
            # Pattern: column letter followed by old_row, but NOT followed by a digit
            # e.g., E80 -> E83, but E801 should not become E831
            pattern = rf'(?<=[A-Z]){old_row}(?!\d)'
            new_formula = re.sub(pattern, str(new_row), formula)

            if new_formula != formula:
                actual_col = start_col + col_offset
                updates.append({
                    'row': row_idx,
                    'col': actual_col,
                    'formula': new_formula,
                })

    if not updates:
        print(f'  {sheet_name}: no formulas to fix')
        continue

    # 3. Batch update with repeatCell
    requests = []
    for upd in updates:
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': upd['row'],
                    'endRowIndex': upd['row'] + 1,
                    'startColumnIndex': upd['col'],
                    'endColumnIndex': upd['col'] + 1,
                },
                'cell': {
                    'userEnteredValue': {'formulaValue': upd['formula']},
                },
                'fields': 'userEnteredValue',
            }
        })

    # Retry on 429
    for attempt in range(5):
        try:
            service.spreadsheets().batchUpdate(
                spreadsheetId=SID,
                body={'requests': requests}
            ).execute()
            print(f'  {sheet_name}: fixed {len(updates)} formulas')
            total_fixed += len(updates)
            break
        except HttpError as e:
            if e.resp.status == 429 and attempt < 4:
                wait = 60 * (attempt + 1)
                print(f'  429, retrying in {wait}s...')
                time.sleep(wait)
            else:
                print(f'  {sheet_name}: FAILED - {e}')
                break

    time.sleep(2)

print(f'\nDone! Total formulas fixed: {total_fixed}')
