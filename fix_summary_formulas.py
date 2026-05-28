#!/usr/bin/env python3
"""
Fix Summary sheet INDIRECT formulas after YoY section insertion.

When add_yoy_section.py inserts the "盈利指标" sub-header row,
all Key Stats items shift down by 1 row. Summary INDIRECT formulas
still reference old row numbers and need +1.

Usage:
    python fix_summary_formulas.py --spreadsheet-id <id>
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


def increment_formula_row(formula):
    """Increment row numbers only after '!' in INDIRECT string references.

    Only increments rows >= 3 because the YoY insertion happens at row 3
    (after the Key Stats header at row 2). Rows 1-2 are unaffected.
    """
    def replace_row(m):
        n = int(m.group(2))
        if n < 3:
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
    print(f"Fixing Summary INDIRECT formulas")
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
                new_fv = increment_formula_row(fv)
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
    args = parser.parse_args()

    service = get_service()
    success = fix_summary(service, args.spreadsheet_id, dry_run=args.dry_run)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
