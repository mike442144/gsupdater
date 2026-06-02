#!/usr/bin/env python3
"""One-off script to fix Net Income formula references in existing company tabs.

Changes formulas that wrongly reference 'Net Income to Company' to 'Net Income':
  - Net Income row
  - Net Margin row
  - ROE row
  - Interest and Rental Exp Coverage Ratio row
  - Net Income YoY row
"""

import sys
import os
import json
import time
import logging

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s', datefmt='%H:%M:%S',
                    force=True)


def scan_sections(rows):
    """Scan A1:C rows to find sections and item-to-row mappings."""
    sections = {}
    current = None
    next_headers = ('key stats', 'income statement', 'balance sheet',
                    'cash flow', 'supplemental', 'business segments')
    for i, row in enumerate(rows):
        a = row[0].strip().lower() if row and row[0] else ''
        if a in next_headers:
            current = a
            sections[current] = {'start': i, 'items': {}}
        elif current:
            b = row[1].strip() if len(row) > 1 and row[1] else ''
            c = row[2].strip() if len(row) > 2 and row[2] else ''
            val = c if c else b
            if val and val not in ('盈利指标', '同比增速'):
                sections[current]['items'][val.lower()] = i
    return sections


def get_spreadsheet_ids():
    with open('industry_spreadsheets.json') as f:
        cfg = json.load(f)
    return [data['spreadsheet_id'] for data in cfg.values()]


def _api_with_retry(func, max_retries=5):
    """Call API func with retry for 429."""
    for attempt in range(max_retries):
        try:
            return func()
        except HttpError as e:
            if e.resp.status == 429 and attempt < max_retries - 1:
                wait = 60 * (attempt + 1)
                logging.warning(f"429, retrying in {wait}s...")
                time.sleep(wait)
            else:
                raise


def process_spreadsheet(service, spreadsheet_id):
    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId,gridProperties))'
    ).execute()

    company_sheets = []
    sheet_id_map = {}
    for s in result.get('sheets', []):
        title = s['properties']['title']
        sheet_id_map[title] = s['properties']['sheetId']
        if title not in ('Summary', 'Template') and '财务' in title:
            company_sheets.append(title)

    logging.info(f"{spreadsheet_id[:20]}...: {len(company_sheets)} tabs")

    total_fixed = 0
    for sheet_name in company_sheets:
        # Read A1:C300 for section scanning
        data = _api_with_retry(lambda: service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A1:C300"
        ).execute())
        rows = data.get('values', [])
        sections = scan_sections(rows)

        is_items = sections.get('income statement', {}).get('items', {})
        # Only fix if IS has BOTH items
        if 'net income to company' not in is_items or 'net income' not in is_items:
            continue

        ks_items = sections.get('key stats', {}).get('items', {})
        sheet_id = sheet_id_map[sheet_name]

        # Read grid data for the Key Stats rows to get actual formulas
        # Find the range of Key Stats rows
        ks_start = sections.get('key stats', {}).get('start', 0)
        ks_end = max(ks_items.values()) + 1 if ks_items else ks_start + 50

        # Read grid data for formulas (columns D through GR, rows ks_start+1 to ks_end+1)
        end_col_letter = 'GR'  # col 200
        data = _api_with_retry(lambda: service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            ranges=[f"{sheet_name}!D{ks_start+1}:{end_col_letter}{ks_end+1}"],
            includeGridData=True,
            fields='sheets.data.rowData.values(userEnteredValue)'
        ).execute())

        row_data = data['sheets'][0].get('data', [{}])[0].get('rowData', [])

        updates = []  # (row_idx_0based, col_idx_0based, new_formula)

        for item_lower, row_idx in ks_items.items():
            data_row_idx = row_idx - ks_start
            if 0 <= data_row_idx < len(row_data):
                cells = row_data[data_row_idx].get('values', [])
                for col_idx, cell in enumerate(cells):
                    uev = cell.get('userEnteredValue', {})
                    formula = uev.get('formulaValue', '')
                    if formula and 'Net Income to Company' in formula:
                        new_formula = formula.replace('Net Income to Company', 'Net Income')
                        if new_formula != formula:
                            # col D = index 3 (0-based)
                            updates.append((row_idx, 3 + col_idx, new_formula))

        if not updates:
            continue

        # Build repeatCell requests
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

        if requests:
            try:
                _api_with_retry(lambda: service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={'requests': requests}
                ).execute())
                total_fixed += len(updates)
                logging.info(f"  {sheet_name}: fixed {len(updates)} formulas")
            except HttpError as e:
                logging.error(f"  {sheet_name}: FAILED - {e}")

        time.sleep(2)

    return total_fixed


def main():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    service = build('sheets', 'v4', credentials=creds)

    spreadsheet_ids = get_spreadsheet_ids()
    logging.info(f"Processing {len(spreadsheet_ids)} spreadsheets")

    total = 0
    for sid in spreadsheet_ids:
        try:
            count = process_spreadsheet(service, sid)
            total += count
        except Exception as e:
            logging.error(f"Spreadsheet {sid[:20]}...: {e}")

    logging.info(f"\nDone! Total formulas fixed: {total}")


if __name__ == '__main__':
    main()
