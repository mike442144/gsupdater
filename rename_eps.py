#!/usr/bin/env python3
"""
Rename Key Stats item "Diluted EPS Excl. Extra Items" → "Basic EPS" and update
all formulas (Key Stats + Summary INDIRECT) to reference the Basic EPS row in IS.

Usage:
    python rename_eps.py                          # all industries
    python rename_eps.py --spreadsheet-id <ID>    # single spreadsheet
    python rename_eps.py --dry-run                # preview only
"""

import sys
import os
import re
import json
import time
import argparse

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')
OLD_NAMES = ['Diluted EPS Excl. Extra Items', 'Normalized Diluted EPS']
NEW_NAME = 'Basic EPS'
SECTION_HEADERS = {'income statement', 'balance sheet', 'cash flow',
                   'key stats', 'supplemental', 'business segments'}
SUB_SECTIONS = {'盈利指标', '同比增速'}


def get_service():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    return build('sheets', 'v4', credentials=creds)


def load_industries():
    path = os.path.join(os.path.dirname(__file__), 'industry_spreadsheets.json')
    with open(path) as f:
        return json.load(f)


def col_to_letter(idx):
    result = ''
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _api_read(service, spreadsheet_id, range_str, max_retries=5):
    for attempt in range(max_retries):
        try:
            return service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_str
            ).execute()
        except Exception as e:
            if '429' in str(e) and attempt < max_retries - 1:
                wait = 60 * (attempt + 1)
                print(f"    Rate limited on read, waiting {wait}s...")
                time.sleep(wait)
            else:
                raise


def get_company_sheets(service, spreadsheet_id):
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId))'
    ).execute()
    sheets = []
    for s in meta.get('sheets', []):
        p = s['properties']
        title = p['title']
        if title == 'Summary' or '资本结构' in title:
            continue
        sheets.append((title, p['sheetId']))
    return sheets


def scan_sections(service, spreadsheet_id, sheet_name):
    """Scan a company tab to find section boundaries and item→row mappings per section."""
    result = _api_read(service, spreadsheet_id, f"'{sheet_name}'!A1:C300")
    rows = result.get('values', [])

    is_items = {}
    ks_items = {}
    current_section = None

    for i, row in enumerate(rows):
        a = row[0].strip().lower() if row and row[0] else ''
        b = row[1].strip() if len(row) > 1 and row[1] else ''
        c = row[2].strip() if len(row) > 2 and row[2] else ''

        if a in SECTION_HEADERS:
            current_section = a
            continue

        item_val = c if c else b
        if not item_val or item_val in SUB_SECTIONS:
            continue

        if current_section == 'income statement':
            is_items[item_val] = i
        elif current_section == 'key stats':
            ks_items[item_val] = i

    return is_items, ks_items


def process_company_tab(service, spreadsheet_id, sheet_name, sheet_id, dry_run):
    """Process one company tab: update formulas and rename item.

    Returns (old_ks_row_0idx, new_ks_row_0idx) or (None, None) if skipped.
    """
    is_items, ks_items = scan_sections(service, spreadsheet_id, sheet_name)

    old_is_row = None
    old_ks_row = None
    matched_old_name = None
    for name in OLD_NAMES:
        if name in is_items and name in ks_items:
            old_is_row = is_items[name]
            old_ks_row = ks_items[name]
            matched_old_name = name
            break

    new_is_row = is_items.get(NEW_NAME)

    if old_ks_row is None:
        return None, None

    if old_is_row is None:
        print(f"    WARNING: IS row for '{matched_old_name}' not found, only renaming label")
    if new_is_row is None:
        print(f"    WARNING: IS row for '{NEW_NAME}' not found, only renaming label")

    requests = []
    item_col = None  # will detect from grid data

    if old_is_row is not None and new_is_row is not None:
        old_row_1 = old_is_row + 1
        new_row_1 = new_is_row + 1

        ks_range = f"'{sheet_name}'!A{old_ks_row + 1}:ZZ{old_ks_row + 1}"
        grid = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            ranges=[ks_range],
            includeGridData=True
        ).execute()

        row_data = (grid.get('sheets', [{}])[0]
                    .get('data', [{}])[0]
                    .get('rowData', [{}])[0])

        # Detect label column from grid data (col 1=B, col 2=C)
        for j, cell in enumerate(row_data.get('values', [])):
            sv = cell.get('effectiveValue', {}).get('stringValue', '')
            if sv.strip() in OLD_NAMES:
                item_col = j
                break

        for j, cell in enumerate(row_data.get('values', [])):
            fv = cell.get('userEnteredValue', {}).get('formulaValue', '')
            if not fv or old_row_1 == new_row_1:
                continue
            new_fv = re.sub(
                rf'(?<=[A-Z]){old_row_1}(?!\d)',
                str(new_row_1), fv
            )
            if new_fv != fv:
                requests.append({
                    'updateCells': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': old_ks_row,
                            'endRowIndex': old_ks_row + 1,
                            'startColumnIndex': j,
                            'endColumnIndex': j + 1,
                        },
                        'rows': [{'values': [{
                            'userEnteredValue': {'formulaValue': new_fv},
                        }]}],
                        'fields': 'userEnteredValue',
                    }
                })

    # Fallback: read label columns if not detected from grid
    if item_col is None:
        for col_idx, col in ((1, 'B'), (2, 'C')):
            cell_ref = f"'{sheet_name}'!{col}{old_ks_row + 1}"
            vals = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id, range=cell_ref
            ).execute().get('values', [])
            val = vals[0][0].strip() if vals and vals[0] else ''
            if val in OLD_NAMES:
                item_col = col_idx
                break

    if item_col is not None:
        requests.append({
            'updateCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': old_ks_row,
                    'endRowIndex': old_ks_row + 1,
                    'startColumnIndex': item_col,
                    'endColumnIndex': item_col + 1,
                },
                'rows': [{'values': [{
                    'userEnteredValue': {'stringValue': NEW_NAME},
                }]}],
                'fields': 'userEnteredValue',
            }
        })

    if requests:
        if dry_run:
            for r in requests:
                uc = r.get('updateCells', {})
                rng = uc.get('range', {})
                rows_data = uc.get('rows', [{}])
                val = rows_data[0].get('values', [{}])[0].get('userEnteredValue', {})
                fv = val.get('formulaValue', '')
                sv = val.get('stringValue', '')
                col_start = rng.get('startColumnIndex', 0)
                if fv:
                    print(f"    [{col_to_letter(col_start)}{old_ks_row + 1}] formula → {fv}")
                elif sv:
                    print(f"    [{col_to_letter(col_start)}{old_ks_row + 1}] label → \"{sv}\"")
        else:
            _batch_update_with_retry(service, spreadsheet_id, requests)

    formula_count = len(requests) - (1 if item_col is not None else 0)
    print(f"  {sheet_name}: {formula_count} formulas updated, label renamed")
    return old_ks_row, new_is_row


def _batch_update_with_retry(service, spreadsheet_id, requests, max_retries=5):
    for attempt in range(max_retries):
        try:
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': requests}
            ).execute()
            return
        except Exception as e:
            if '429' in str(e) and attempt < max_retries - 1:
                wait = 60 * (attempt + 1)
                print(f"    Rate limited, waiting {wait}s...")
                time.sleep(wait)
            else:
                raise


def process_spreadsheet(service, spreadsheet_id, industry_name, dry_run):
    print(f"\n{'='*60}")
    print(f"{industry_name} — {spreadsheet_id[:20]}...")
    print(f"{'='*60}")

    sheets = get_company_sheets(service, spreadsheet_id)
    if not sheets:
        print("  No company tabs found")
        return 0, 0

    updated = 0
    skipped = 0

    for i, (sheet_name, sheet_id) in enumerate(sheets):
        if i > 0:
            time.sleep(2)
        old_ks_row, new_is_row = process_company_tab(
            service, spreadsheet_id, sheet_name, sheet_id, dry_run
        )
        if old_ks_row is not None:
            updated += 1
        else:
            skipped += 1

    print(f"  → Updated: {updated}, Skipped: {skipped}")
    return updated, skipped


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--spreadsheet-id', help='Target a single spreadsheet')
    parser.add_argument('--dry-run', action='store_true', help='Preview only')
    args = parser.parse_args()

    service = get_service()
    total_updated = 0
    total_skipped = 0

    if args.spreadsheet_id:
        u, s = process_spreadsheet(service, args.spreadsheet_id, 'Custom', args.dry_run)
        total_updated += u
        total_skipped += s
    else:
        industries = load_industries()
        for name, info in industries.items():
            u, s = process_spreadsheet(service, info['spreadsheet_id'], name, args.dry_run)
            total_updated += u
            total_skipped += s

    action = '[DRY RUN] Would update' if args.dry_run else 'Updated'
    print(f"\n{action}: {total_updated} tabs, Skipped: {total_skipped} tabs")


if __name__ == '__main__':
    main()
