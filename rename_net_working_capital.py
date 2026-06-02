#!/usr/bin/env python3
"""One-off script to rename 'Net Working Capital' -> 'Net-nets' in all existing company tabs.

Also updates the formula from Total Current Assets - Total Current Liabilities
to Total Current Assets - Total Liabilities.
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

OLD_NAME = 'Net Working Capital'
NEW_NAME = 'Net-nets'


def get_spreadsheet_ids():
    with open('industry_spreadsheets.json') as f:
        cfg = json.load(f)
    return [data['spreadsheet_id'] for data in cfg.values()]


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


def process_spreadsheet(service, spreadsheet_id):
    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId))'
    ).execute()

    company_sheets = {}
    for s in result.get('sheets', []):
        t = s['properties']['title']
        if '财务' in t and t not in ('Summary', 'Template'):
            company_sheets[t] = s['properties']['sheetId']

    total_renamed = 0
    for sheet_name, sheet_id in company_sheets.items():
        # 1. Read A1:C300 for sections
        data = _api_with_retry(lambda: service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!A1:C300'
        ).execute())
        rows = data.get('values', [])
        sections = scan_sections(rows)
        ks_items = sections.get('key stats', {}).get('items', {})

        if OLD_NAME.lower() not in ks_items:
            continue

        row_idx = ks_items[OLD_NAME.lower()]

        # 2. Check if formula uses "Total Current Liabilities" (old) vs "Total Liabilities" (new)
        # Read grid data for formula cells on this row
        data = _api_with_retry(lambda: service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            ranges=[f'{sheet_name}!D{row_idx+1}:GR{row_idx+1}'],
            includeGridData=True,
            fields='sheets.data.rowData.values(userEnteredValue),sheets.data.startRow,sheets.data.startColumn'
        ).execute())
        d = data['sheets'][0]['data'][0]
        rd = d.get('rowData', [])
        start_col = d.get('startColumn', 0)

        updates = []  # label + formula updates
        has_formula_change = False

        if rd:
            cells = rd[0].get('values', [])
            for col_offset, cell in enumerate(cells):
                uev = cell.get('userEnteredValue', {})
                formula = uev.get('formulaValue', '')
                if formula and 'Total Current Liabilities' in formula:
                    new_formula = formula.replace(
                        'Total Current Liabilities', 'Total Liabilities')
                    updates.append({
                        'row': row_idx,
                        'col': start_col + col_offset,
                        'formula': new_formula,
                    })
                    has_formula_change = True

        # 3. Rename the label (column C, or B if C is empty)
        # Check which column has the label
        label_data = _api_with_retry(lambda: service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            ranges=[f'{sheet_name}!B{row_idx+1}:C{row_idx+1}'],
            includeGridData=True,
            fields='sheets.data.rowData.values(userEnteredValue),sheets.data.startRow,sheets.data.startColumn'
        ).execute())
        ld = label_data['sheets'][0]['data'][0]
        lrd = ld.get('rowData', [])
        label_col = None
        if lrd:
            label_start_col = ld.get('startColumn', 0)
            lcells = lrd[0].get('values', [])
            for co, lc in enumerate(lcells):
                sv = lc.get('userEnteredValue', {}).get('stringValue', '')
                if sv == OLD_NAME:
                    label_col = label_start_col + co
                    break

        requests = []

        # Update label
        if label_col is not None:
            requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': row_idx,
                        'endRowIndex': row_idx + 1,
                        'startColumnIndex': label_col,
                        'endColumnIndex': label_col + 1,
                    },
                    'cell': {
                        'userEnteredValue': {'stringValue': NEW_NAME},
                    },
                    'fields': 'userEnteredValue',
                }
            })

        # Update formulas
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

        if requests:
            try:
                _api_with_retry(lambda: service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={'requests': requests}
                ).execute())
                desc = []
                if label_col is not None:
                    desc.append('label')
                if has_formula_change:
                    desc.append(f'{len(updates)} formulas')
                print(f'  {sheet_name}: renamed {", ".join(desc)}')
                total_renamed += 1
            except HttpError as e:
                print(f'  {sheet_name}: FAILED - {e}')

        time.sleep(2)

    return total_renamed


def main():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    service = build('sheets', 'v4', credentials=creds)

    spreadsheet_ids = get_spreadsheet_ids()
    print(f'Processing {len(spreadsheet_ids)} spreadsheets')

    total = 0
    for sid in spreadsheet_ids:
        try:
            count = process_spreadsheet(service, sid)
            total += count
        except Exception as e:
            print(f'Spreadsheet {sid[:20]}...: {e}')

    print(f'\nDone! Total tabs renamed: {total}')


if __name__ == '__main__':
    main()
