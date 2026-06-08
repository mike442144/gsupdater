#!/usr/bin/env python3
"""
Fix the Summary "自由现金流/净利润" ratio after the Net Income to Company row exists.

Run order: (1) add_net_income_to_company.py inserts the row in each company tab,
(2) the row-index +1 script fixes Summary INDIRECT refs for the shift, (3) this
script renames the ratio "自由现金流/净利润" -> "自由现金流/公司净利润" and repoints its
denominator from Net Income (company-tab row N) to Net Income to Company (row N+1).

Usage:
    python fix_summary_net_income_refs.py --spreadsheet-id <id>
    python fix_summary_net_income_refs.py --spreadsheet-id <id> --dry-run
"""

import sys
import os
import re
import argparse

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')


def get_service():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    return build('sheets', 'v4', credentials=creds)


def find_net_income_row(service, spreadsheet_id, sheet_name):
    """Find the row number of 'Net Income' in the company tab's Key Stats section.

    Returns the 0-indexed row number, or None if not found.
    """
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!A1:C300"
    ).execute()

    rows = result.get('values', [])
    for i, row in enumerate(rows):
        a = row[0].strip().lower() if row and row[0] else ''
        b = row[1].strip() if len(row) > 1 and row[1] else ''
        c = row[2].strip() if len(row) > 2 and row[2] else ''

        if a == 'key stats':
            # Found Key Stats section, now look for Net Income
            for j in range(i + 1, min(i + 30, len(rows))):
                r = rows[j]
                item_val = (r[2].strip() if len(r) > 2 and r[2] else
                           r[1].strip() if len(r) > 1 and r[1] else '')
                if item_val and item_val.lower() == 'net income':
                    return j
    return None


def retarget_net_income_row(formula, ni_row, ni_to_company_row):
    """Repoint a ratio's Net Income denominator to the Net Income to Company row.

    Swaps cell-ref row `ni_row` -> `ni_to_company_row` inside "'...财务'!<refs>"
    INDIRECT bodies. Only the denominator references ni_row (the numerator is the
    FCFF row), so swapping that row number is safe.
    """
    def replace_row(m):
        if int(m.group(2)) == ni_row:
            return f'{m.group(1)}{ni_to_company_row}'
        return m.group(0)

    def fix_indirect(m):
        return m.group(1) + re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)', replace_row, m.group(2))

    return re.sub(r"('[^']*'!)([^)\"]+)", fix_indirect, formula)


def fix_summary(service, spreadsheet_id, dry_run=False):
    print(f"\n{'='*60}")
    print(f"Fixing Summary INDIRECT formulas after Net Income to Company insertion")
    print(f"Spreadsheet: {spreadsheet_id[:30]}...")
    print(f"{'='*60}")

    # Get Summary sheet ID and all company sheet names
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId))'
    ).execute()

    summary_sheet_id = None
    company_sheets = []
    for s in meta.get('sheets', []):
        title = s['properties']['title']
        if title == 'Summary':
            summary_sheet_id = s['properties']['sheetId']
        elif title != 'Summary' and '资本结构' not in title:
            company_sheets.append(title)

    if summary_sheet_id is None:
        print("ERROR: Summary sheet not found")
        return False

    if not company_sheets:
        print("ERROR: No company sheets found")
        return False

    # Find Net Income row in the first company sheet that has one (skips non-financial
    # tabs that come before the financial tabs in MIXED industries). All financial tabs
    # share the same layout, so the first match is representative.
    net_income_row = None
    for example_sheet in company_sheets:
        net_income_row = find_net_income_row(service, spreadsheet_id, example_sheet)
        if net_income_row is not None:
            break

    if net_income_row is None:
        print("ERROR: Could not find Net Income row in any company tab")
        return False

    ni_row = net_income_row + 1             # 1-indexed Net Income row in company tabs
    ni_to_company_row = net_income_row + 2  # 1-indexed Net Income to Company row (Net Income + 1)
    print(f"  Net Income at row {ni_row}, Net Income to Company at row {ni_to_company_row} in company tabs")

    # Read all Summary cells with formulas
    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=["'Summary'"],
        includeGridData=True
    ).execute()

    sheets_data = result.get('sheets', [])
    if not sheets_data:
        print("ERROR: No data returned")
        return False

    row_data = sheets_data[0].get('data', [{}])[0].get('rowData', [])

    # In each "自由现金流/净利润" ratio row: rename the label and repoint the denominator
    # from Net Income to Net Income to Company. The label sits in an early column;
    # the ratio formulas follow in the data columns of the same row.
    fcf_ratio_updates = []
    fcf_label_updates = []
    for i, row in enumerate(row_data):
        values = row.get('values', [])
        is_fcf_ratio = any(c.get('userEnteredValue', {}).get('stringValue')
                           in ('自由现金流/净利润', '自由现金流/公司净利润') for c in values)
        if not is_fcf_ratio:
            continue
        for j, cell in enumerate(values):
            uev = cell.get('userEnteredValue', {})
            if uev.get('stringValue') == '自由现金流/净利润':
                fcf_label_updates.append((i, j, '自由现金流/净利润', '自由现金流/公司净利润'))
                continue
            fv = uev.get('formulaValue', '')
            if fv and 'INDIRECT' in fv:
                new_fv = retarget_net_income_row(fv, ni_row, ni_to_company_row)
                if new_fv != fv:
                    fcf_ratio_updates.append((i, j, fv, new_fv))

    print(f"  Found {len(fcf_ratio_updates)} ratio formulas and {len(fcf_label_updates)} labels to update")

    if not fcf_ratio_updates and not fcf_label_updates:
        print("  Nothing to fix (ratio already points at Net Income to Company?)")
        return True

    # Show samples
    if fcf_ratio_updates:
        print(f"\n  Ratio formula updates (first 3):")
        for i, j, old, new in fcf_ratio_updates[:3]:
            print(f"    Row {i+1} Col {j}:")
            print(f"      OLD: {old}")
            print(f"      NEW: {new}")
        if len(fcf_ratio_updates) > 3:
            print(f"    ... and {len(fcf_ratio_updates) - 3} more")

    if fcf_label_updates:
        print(f"\n  Label updates:")
        for i, j, old, new in fcf_label_updates:
            print(f"    Row {i+1} Col {j}: '{old}' → '{new}'")

    if dry_run:
        print(f"\n  [DRY RUN] Would fix {len(fcf_ratio_updates) + len(fcf_label_updates)} cells")
        return True

    # Build batch update requests
    requests = []

    # Fix FCF ratio formulas
    for i, j, _, new_fv in fcf_ratio_updates:
        requests.append({
            'updateCells': {
                'range': {
                    'sheetId': summary_sheet_id,
                    'startRowIndex': i,
                    'endRowIndex': i + 1,
                    'startColumnIndex': j,
                    'endColumnIndex': j + 1,
                },
                'rows': [{'values': [{
                    'userEnteredValue': {'formulaValue': new_fv},
                }]}],
                'fields': 'userEnteredValue',
            }
        })

    # Fix labels
    for i, j, _, new_label in fcf_label_updates:
        requests.append({
            'updateCells': {
                'range': {
                    'sheetId': summary_sheet_id,
                    'startRowIndex': i,
                    'endRowIndex': i + 1,
                    'startColumnIndex': j,
                    'endColumnIndex': j + 1,
                },
                'rows': [{'values': [{
                    'userEnteredValue': {'stringValue': new_label},
                }]}],
                'fields': 'userEnteredValue',
            }
        })

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': requests}
    ).execute()

    print(f"\n  Updated {len(fcf_ratio_updates)} ratio formulas and {len(fcf_label_updates)} labels")
    return True


def main():
    parser = argparse.ArgumentParser(description='Fix Summary INDIRECT formula row refs after Net Income to Company insertion')
    parser.add_argument('--spreadsheet-id', required=True, help='Google Spreadsheet ID')
    parser.add_argument('--dry-run', action='store_true', help='Preview changes without writing')
    args = parser.parse_args()

    service = get_service()
    success = fix_summary(service, args.spreadsheet_id, dry_run=args.dry_run)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
