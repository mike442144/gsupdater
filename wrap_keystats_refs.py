#!/usr/bin/env python3
"""
Wrap Key Stats formula references to a given item's row in N(), so that CIQ's
'-' nil marker (a text string) is coerced to 0 instead of producing #VALUE!.

Defaults to "Minority Interest" — its balance-sheet cell is frequently '-' for
companies with no minority interest, which breaks any arithmetic Key Stats
formula (e.g. the capital-source ROIC denominator) that references it.

Only formulas inside the Key Stats section are touched. The rewrite is
idempotent: a reference already wrapped as N(<ref>) is left alone, so the
script is safe to re-run.

Usage:
    python wrap_keystats_refs.py --spreadsheet-id <id>
    python wrap_keystats_refs.py --spreadsheet-id <id> --dry-run
    python wrap_keystats_refs.py --spreadsheet-id <id> --item "Minority Interest"
    python wrap_keystats_refs.py                       # all industries
    python wrap_keystats_refs.py --dry-run
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
SECTION_HEADERS = {'income statement', 'balance sheet', 'cash flow',
                   'key stats', 'supplemental', 'business segments'}
DEFAULT_ITEM = "Minority Interest"


def get_service():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    return build('sheets', 'v4', credentials=creds)


def load_industries():
    path = os.path.join(os.path.dirname(__file__), 'industry_spreadsheets.json')
    with open(path) as f:
        return json.load(f)


def _retry(fn, max_retries=5):
    for attempt in range(max_retries):
        try:
            return fn()
        except Exception as e:
            if '429' in str(e) and attempt < max_retries - 1:
                wait = 60 * (attempt + 1)
                print(f"    Rate limited, waiting {wait}s...")
                time.sleep(wait)
            else:
                raise


def get_company_sheets(service, spreadsheet_id):
    meta = _retry(lambda: service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId))'
    ).execute())
    sheets = []
    for s in meta.get('sheets', []):
        p = s['properties']
        title = p['title']
        if title == 'Summary' or '资本结构' in title:
            continue
        sheets.append((title, p['sheetId']))
    return sheets


def locate(service, spreadsheet_id, sheet_name, item_lower):
    """Return (target_rows[1-indexed], ks_start_0idx, ks_end_0idx).

    target_rows: every row whose B/C label equals the item (case-insensitive).
    ks_start/ks_end: 0-indexed inclusive bounds of the Key Stats section body.
    """
    rows = _retry(lambda: service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"'{sheet_name}'!A1:C300"
    ).execute()).get('values', [])

    target_rows = []
    ks_start = ks_end = None
    section = None
    for i, row in enumerate(rows):
        a = row[0].strip().lower() if row and row[0] else ''
        b = row[1].strip() if len(row) > 1 and row[1] else ''
        c = row[2].strip() if len(row) > 2 and row[2] else ''
        if a in SECTION_HEADERS:
            if a == 'key stats':
                ks_start = i + 1
            elif ks_start is not None and ks_end is None:
                ks_end = i - 1
            section = a
            continue
        val = c if c else b
        if val and val.lower() == item_lower:
            target_rows.append(i + 1)
    if ks_start is not None and ks_end is None:
        ks_end = len(rows) - 1
    return target_rows, ks_start, ks_end


def wrap_refs(formula, target_rows):
    r"""Wrap each unwrapped reference to a target row in N(). Idempotent.

    Two lookbehinds guard the match start:
      (?<![A-Za-z0-9$])  the ref must begin at a token boundary, so we never grab
                         a suffix of a multi-letter column (e.g. the inner 'A190'
                         of an already-wrapped 'N(AA190)').
      (?<!N\()           a ref already wrapped as N(<ref>) is left alone.
    """
    for t in target_rows:
        formula = re.sub(rf'(?<![A-Za-z0-9$])(?<!N\()(\$?[A-Z]{{1,3}}\$?{t})(?!\d)',
                         r'N(\1)', formula)
    return formula


def process_tab(service, spreadsheet_id, sheet_name, sheet_id, item_lower, dry_run):
    target_rows, ks_start, ks_end = locate(service, spreadsheet_id, sheet_name, item_lower)
    if ks_start is None:
        print(f"  {sheet_name}: SKIP — no Key Stats section")
        return 'skip'
    if not target_rows:
        print(f"  {sheet_name}: SKIP — '{item_lower}' row not found")
        return 'skip'

    grid = _retry(lambda: service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[f"'{sheet_name}'!A{ks_start + 1}:BZ{ks_end + 1}"],
        includeGridData=True,
        fields='sheets.data.rowData.values.userEnteredValue'
    ).execute())
    rd = grid.get('sheets', [{}])[0].get('data', [{}])[0].get('rowData', [])

    requests = []
    samples = []
    for ri, row in enumerate(rd):
        for ci, cell in enumerate(row.get('values', [])):
            fv = cell.get('userEnteredValue', {}).get('formulaValue', '')
            if not fv:
                continue
            new = wrap_refs(fv, target_rows)
            if new == fv:
                continue
            abs_row = ks_start + ri  # 0-indexed
            requests.append({
                'updateCells': {
                    'range': {'sheetId': sheet_id,
                              'startRowIndex': abs_row, 'endRowIndex': abs_row + 1,
                              'startColumnIndex': ci, 'endColumnIndex': ci + 1},
                    'rows': [{'values': [{'userEnteredValue': {'formulaValue': new}}]}],
                    'fields': 'userEnteredValue',
                }
            })
            if len(samples) < 2:
                samples.append((abs_row + 1, fv, new))

    if not requests:
        print(f"  {sheet_name}: nothing to wrap (rows {target_rows})")
        return 'noop'

    for r1, old, new in samples:
        print(f"    row {r1}: {old}  ->  {new}")
    if dry_run:
        print(f"  {sheet_name}: [DRY RUN] would wrap {len(requests)} cells "
              f"(target rows {target_rows})")
        return 'update'

    _retry(lambda: service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={'requests': requests}).execute())
    print(f"  {sheet_name}: wrapped {len(requests)} cells (target rows {target_rows})")
    return 'update'


def process_spreadsheet(service, spreadsheet_id, label, item_lower, dry_run):
    print(f"\n{'='*60}\n{label} — {spreadsheet_id[:24]}...\n{'='*60}")
    sheets = get_company_sheets(service, spreadsheet_id)
    if not sheets:
        print("  No company tabs found")
        return 0, 0
    updated = other = 0
    for i, (name, sid) in enumerate(sheets):
        if i > 0:
            time.sleep(2)
        if process_tab(service, spreadsheet_id, name, sid, item_lower, dry_run) == 'update':
            updated += 1
        else:
            other += 1
    print(f"  -> Updated: {updated}, Skipped/noop: {other}")
    return updated, other


def main():
    parser = argparse.ArgumentParser(description='Wrap Key Stats refs to an item row in N()')
    parser.add_argument('--spreadsheet-id', help='Target a single spreadsheet')
    parser.add_argument('--item', default=DEFAULT_ITEM, help=f'Item name (default: {DEFAULT_ITEM})')
    parser.add_argument('--dry-run', action='store_true', help='Preview only')
    args = parser.parse_args()

    service = get_service()
    item_lower = args.item.strip().lower()
    tu = to = 0
    if args.spreadsheet_id:
        u, o = process_spreadsheet(service, args.spreadsheet_id, 'Custom', item_lower, args.dry_run)
        tu, to = tu + u, to + o
    else:
        for name, info in load_industries().items():
            u, o = process_spreadsheet(service, info['spreadsheet_id'], name, item_lower, args.dry_run)
            tu, to = tu + u, to + o

    action = '[DRY RUN] Would update' if args.dry_run else 'Updated'
    print(f"\n{action}: {tu} tabs, Skipped/noop: {to} tabs")


if __name__ == '__main__':
    main()
