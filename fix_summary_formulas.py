#!/usr/bin/env python3
"""
Fix Summary sheet INDIRECT formulas after a row is inserted in the company tabs.

When a row is inserted at position --after-row, all company-tab items at or below
it shift down by 1. Summary INDIRECT formulas still reference old row numbers and
need +1 for rows >= --after-row. Defaults to row 3 (the YoY "盈利指标" sub-header
from add_yoy_section.py); pass --after-row 6 for the Net Income to Company row.

Usage:
    python fix_summary_formulas.py --spreadsheet-id <id>                  # YoY (row 3)
    python fix_summary_formulas.py --spreadsheet-id <id> --after-row 6    # Net Income to Company
    python fix_summary_formulas.py --spreadsheet-id <id> --dry-run
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


def increment_formula_row(formula, after_row):
    """Increment row numbers only after '!' in INDIRECT string references.

    Increments rows >= after_row (the position where the new row was inserted);
    rows above the insertion point are unaffected.
    """
    def replace_row(m):
        n = int(m.group(2))
        if n < after_row:
            return m.group(0)
        return f'{m.group(1)}{n + 1}'

    def fix_indirect(m):
        prefix = m.group(1)
        cell_ref = m.group(2)
        cell_ref = re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)', replace_row, cell_ref)
        return prefix + cell_ref

    return re.sub(r"('[^']*'!)([^)\"]+)", fix_indirect, formula)


def fix_summary(service, spreadsheet_id, dry_run=False, after_row=3):
    print(f"\n{'='*60}")
    print(f"Fixing Summary INDIRECT formulas (rows >= {after_row} shift +1)")
    print(f"Spreadsheet: {spreadsheet_id[:30]}...")
    print(f"{'='*60}")

    # Get Summary sheet ID
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId))'
    ).execute()
    summary_sheet_id = None
    for s in meta.get('sheets', []):
        if s['properties']['title'] == 'Summary':
            summary_sheet_id = s['properties']['sheetId']
            break
    if summary_sheet_id is None:
        print("ERROR: Summary sheet not found")
        return False

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
    for i, row in enumerate(row_data):
        values = row.get('values', [])
        for j, cell in enumerate(values):
            uev = cell.get('userEnteredValue', {})
            fv = uev.get('formulaValue', '')
            if 'INDIRECT' in fv:
                new_fv = increment_formula_row(fv, after_row)
                if new_fv != fv:
                    updates.append((i, j, fv, new_fv))

    print(f"  Found {len(updates)} formulas to fix")
    if not updates:
        print("  Nothing to fix")
        return True

    # Show samples
    for i, j, old, new in updates[:5]:
        print(f"  Row {i+1} Col {j}:")
        print(f"    OLD: {old}")
        print(f"    NEW: {new}")
    if len(updates) > 5:
        print(f"  ... and {len(updates) - 5} more")

    if dry_run:
        print(f"\n  [DRY RUN] Would fix {len(updates)} formulas")
        return True

    # Build batch update requests
    requests = []
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

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': requests}
    ).execute()
    print(f"\n  Fixed {len(updates)} Summary formulas (row refs +1)")
    return True


def main():
    parser = argparse.ArgumentParser(description='Fix Summary INDIRECT formula row refs after YoY insertion')
    parser.add_argument('--spreadsheet-id', required=True, help='Google Spreadsheet ID')
    parser.add_argument('--dry-run', action='store_true', help='Preview changes without writing')
    parser.add_argument('--after-row', type=int, default=3,
                        help='Insertion position; rows >= this shift +1 (default 3 = YoY section)')
    args = parser.parse_args()

    service = get_service()
    success = fix_summary(service, args.spreadsheet_id, dry_run=args.dry_run, after_row=args.after_row)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
