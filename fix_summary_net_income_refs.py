#!/usr/bin/env python3
"""
Fix Summary sheet INDIRECT formulas after inserting Net Income to Company row.

When add_net_income_to_company.py inserts the "Net Income to Company" row after
"Net Income", all Key Stats items that were originally after Net Income shift
down by 1 row. Summary INDIRECT formulas still reference old row numbers and need +1.

Also updates the "自由现金流/净利润" ratio to use "Net Income to Company" as denominator.

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


def increment_formula_row(formula, min_row):
    """Increment row numbers only after '!' in INDIRECT string references.

    Only increments rows >= min_row because the insertion happens at min_row.
    """
    def replace_row(m):
        n = int(m.group(2))
        if n < min_row:
            return m.group(0)
        return f'{m.group(1)}{n + 1}'

    def fix_indirect(m):
        prefix = m.group(1)
        cell_ref = m.group(2)
        cell_ref = re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)', replace_row, cell_ref)
        return prefix + cell_ref

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

    # Find Net Income row in the first company sheet (they should all be the same)
    example_sheet = company_sheets[0]
    net_income_row = find_net_income_row(service, spreadsheet_id, example_sheet)

    if net_income_row is None:
        print(f"ERROR: Could not find Net Income row in {example_sheet}")
        return False

    print(f"  Net Income is at row {net_income_row + 1} in company tabs")
    print(f"  Insertion happens after row {net_income_row + 1}, so rows >= {net_income_row + 2} shift down")

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

    # Collect cells that need fixing
    updates = []
    fcf_ratio_updates = []
    fcf_label_updates = []

    for i, row in enumerate(row_data):
        values = row.get('values', [])
        for j, cell in enumerate(values):
            uev = cell.get('userEnteredValue', {})
            fv = uev.get('formulaValue', '')

            # Fix INDIRECT formulas
            if 'INDIRECT' in fv:
                new_fv = increment_formula_row(fv, net_income_row + 2)  # 1-indexed
                if new_fv != fv:
                    updates.append((i, j, fv, new_fv))
                    continue

            # Look for "自由现金流/净利润" label to change to "自由现金流/公司净利润"
            if 'stringValue' in uev and uev['stringValue'] == '自由现金流/净利润':
                fcf_label_updates.append((i, j, '自由现金流/净利润', '自由现金流/公司净利润'))

    print(f"  Found {len(updates)} INDIRECT formulas to fix")
    print(f"  Found {len(fcf_label_updates)} labels to update")

    # Also need to update FCF/Net Income ratio formula to use Net Income to Company
    # First, find where "Net Income to Company" is now (Net Income row + 1)
    net_income_to_company_row = net_income_row + 1  # 0-indexed

    # Now find and update FCF/Net Income formulas
    for i, row in enumerate(row_data):
        values = row.get('values', [])
        for j, cell in enumerate(values):
            uev = cell.get('userEnteredValue', {})
            fv = uev.get('formulaValue', '')

            if fv and 'INDIRECT' in fv:
                # Check if this is a free cash flow / net income ratio
                # We look for formulas that reference Net Income
                if f"!C{net_income_row + 1}" in fv or f"!C{net_income_row + 2}" in fv:
                    # This might be the FCF/Net Income ratio
                    # Replace Net Income reference with Net Income to Company
                    old_net_income_ref = f"!C{net_income_row + 2}"  # 1-indexed after shift
                    new_net_income_ref = f"!C{net_income_to_company_row + 2}"  # Net Income to Company row

                    if old_net_income_ref in fv:
                        new_fv = fv.replace(old_net_income_ref, new_net_income_ref)
                        if new_fv != fv:
                            fcf_ratio_updates.append((i, j, fv, new_fv))

    print(f"  Found {len(fcf_ratio_updates)} FCF/Net Income ratio formulas to update")

    if not updates and not fcf_label_updates and not fcf_ratio_updates:
        print("  Nothing to fix")
        return True

    # Show samples
    if updates:
        print(f"\n  INDIRECT formula fixes (first 3):")
        for i, j, old, new in updates[:3]:
            print(f"    Row {i+1} Col {j}:")
            print(f"      OLD: {old}")
            print(f"      NEW: {new}")
        if len(updates) > 3:
            print(f"    ... and {len(updates) - 3} more")

    if fcf_ratio_updates:
        print(f"\n  FCF/Net Income ratio updates:")
        for i, j, old, new in fcf_ratio_updates:
            print(f"    Row {i+1} Col {j}:")
            print(f"      OLD: {old}")
            print(f"      NEW: {new}")

    if fcf_label_updates:
        print(f"\n  Label updates:")
        for i, j, old, new in fcf_label_updates:
            print(f"    Row {i+1} Col {j}: '{old}' → '{new}'")

    if dry_run:
        total = len(updates) + len(fcf_ratio_updates) + len(fcf_label_updates)
        print(f"\n  [DRY RUN] Would fix {total} cells")
        return True

    # Build batch update requests
    requests = []

    # Fix INDIRECT formulas
    for i, j, _, new_fv in updates:
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

    print(f"\n  Fixed {len(updates)} INDIRECT formulas (row refs +1)")
    print(f"  Updated {len(fcf_ratio_updates)} FCF ratio formulas")
    print(f"  Updated {len(fcf_label_updates)} labels")
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
