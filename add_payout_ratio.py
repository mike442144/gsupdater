#!/usr/bin/env python3
"""
Add a "Payout Ratio %" row to each company tab's Key Stats section, right below
the last ROIC row (the bottom of the 盈利指标 numeric block).

    Payout Ratio % = Dividends per Share / Basic EPS

DPS is N()-wrapped so CIQ's '-' nil marker (no-dividend years) coerces to 0
(→ 0%), and the whole thing is wrapped in IFERROR so a zero/negative Basic EPS
yields a blank cell instead of #DIV/0!.

Anchor = last ROIC row (fallback: Total Revenue). Summary INDIRECT formulas only
reference Key Stats rows up to "Total Revenue" / ROIC, so inserting a row after
the last ROIC row does NOT shift any Summary-referenced row — no
fix_summary_formulas pass is needed. Mirrors add_roic_methods.py.

Usage:
    python add_payout_ratio.py --spreadsheet-id <id>            # one spreadsheet
    python add_payout_ratio.py --spreadsheet-id <id> --dry-run  # preview only
    python add_payout_ratio.py                                  # all industries
    python add_payout_ratio.py --dry-run
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
SUB_SECTIONS = {'盈利指标', '同比增速'}

PAYOUT_LABEL = "Payout Ratio %"
# {!Item} -> N(ref): coerces CIQ '-' nil text to 0. IFERROR(...,) -> blank on a
# zero/negative Basic EPS divisor. Only __C__ (current column) is referenced.
PAYOUT_TEMPLATE = "=IFERROR(__C__{!Dividends per Share}/__C__{Basic EPS},)"
PERCENT_FORMAT = {"type": "PERCENT", "pattern": "0.0%"}


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


def scan_tab(service, spreadsheet_id, sheet_name):
    """Return (item_to_row, payout_row, anchor_row, label_col).

    item_to_row: full name->0idx row map (last-wins, all sections) for formula refs.
    payout_row:  0-indexed row of an existing 'Payout Ratio %' Key Stats item, else None.
    anchor_row:  0-indexed row to insert after — the LAST ROIC row, or Total Revenue.
    label_col:   0-indexed column where the anchor's label sits (1=B or 2=C).
    """
    result = _retry(lambda: service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"'{sheet_name}'!A1:C300"
    ).execute())
    rows = result.get('values', [])

    item_to_row = {}
    payout_row = None
    roic_row = None
    roic_col = 2
    total_rev_row = None
    tr_col = 2
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

        item_to_row[item_val.lower()] = i
        if current_section == 'key stats':
            low = item_val.lower()
            if low == PAYOUT_LABEL.lower():
                payout_row = i
            elif low.startswith('roic'):
                # Last ROIC row wins — anchor below the full ROIC block.
                roic_row = i
                roic_col = 2 if c else 1
            elif low == 'total revenue':
                total_rev_row = i
                tr_col = 2 if c else 1

    if roic_row is not None:
        anchor_row, label_col = roic_row, roic_col
    elif total_rev_row is not None:
        anchor_row, label_col = total_rev_row, tr_col
    else:
        anchor_row, label_col = None, 2
    return item_to_row, payout_row, anchor_row, label_col


def find_data_columns(service, spreadsheet_id, sheet_name):
    """Data columns = columns from D onward whose row-1 header contains a 4-digit year."""
    result = _retry(lambda: service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[f"'{sheet_name}'!A1:CV1"],
        includeGridData=True,
        fields='sheets.data.rowData.values(userEnteredValue,formattedValue)'
    ).execute())
    vals = (result.get('sheets', [{}])[0].get('data', [{}])[0]
            .get('rowData', [{}])[0].get('values', []))
    cols = []
    for j in range(3, len(vals)):
        uev = vals[j].get('userEnteredValue', {})
        text = uev.get('stringValue', '') or vals[j].get('formattedValue', '')
        if text and re.search(r'\d{4}', str(text)):
            cols.append(j)
    return cols


def resolve_formula(template, col_letter, item_to_row, sheet_name):
    """Resolve __C__{Item} tokens to a GS formula (current column only).

    '!Item' -> N(<col><row>), coercing CIQ's '-' nil text to 0. A missing item
    falls back to <col>1 and is reported.
    """
    missing = []

    def repl(m):
        name = m.group(2)
        marker = ''
        if name[:1] in ('?', '!'):
            marker, name = name[0], name[1:]
        row = item_to_row.get(name.lower())
        if row is None:
            if marker == '?':
                return '0'
            missing.append(name)
            return col_letter + '1'
        if marker:
            return f'N({col_letter}{row + 1})'
        return f'{col_letter}{row + 1}'

    formula = re.sub(r'(__C__|__PC__)\{([?!]?[^}]+)\}', repl, template)
    if missing:
        print(f"    WARNING [{sheet_name}]: unresolved items {sorted(set(missing))}")
    return formula


def process_tab(service, spreadsheet_id, sheet_name, sheet_id, dry_run):
    item_to_row, payout_row, anchor_row, label_col = scan_tab(
        service, spreadsheet_id, sheet_name)

    if anchor_row is None:
        print(f"  {sheet_name}: SKIP — no ROIC or Total Revenue row in Key Stats")
        return 'skip'

    data_cols = find_data_columns(service, spreadsheet_id, sheet_name)
    if not data_cols:
        print(f"  {sheet_name}: SKIP — no data columns found")
        return 'skip'

    requests = []
    if payout_row is not None:
        # Re-runnable: rewrite formulas in the existing row, no insert/shift.
        target_row = payout_row
        shifted = item_to_row
        mode = 'rewrite'
    else:
        insert_at = anchor_row + 1  # 0-indexed row index for the new row
        target_row = insert_at
        # Items at original row >= insert_at shift down by 1; anchor & above unaffected.
        shifted = {n: (r + 1 if r >= insert_at else r) for n, r in item_to_row.items()}
        mode = 'insert'
        requests.append({
            'insertDimension': {
                'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                          'startIndex': insert_at, 'endIndex': insert_at + 1},
                'inheritFromBefore': True,
            }
        })

    # label
    requests.append({
        'updateCells': {
            'range': {'sheetId': sheet_id,
                      'startRowIndex': target_row, 'endRowIndex': target_row + 1,
                      'startColumnIndex': label_col, 'endColumnIndex': label_col + 1},
            'rows': [{'values': [{'userEnteredValue': {'stringValue': PAYOUT_LABEL}}]}],
            'fields': 'userEnteredValue',
        }
    })
    # formulas across data columns
    last_formula = None
    for dc in data_cols:
        cl = col_to_letter(dc)
        formula = resolve_formula(PAYOUT_TEMPLATE, cl, shifted, sheet_name)
        last_formula = formula
        requests.append({
            'updateCells': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': target_row, 'endRowIndex': target_row + 1,
                          'startColumnIndex': dc, 'endColumnIndex': dc + 1},
                'rows': [{'values': [{
                    'userEnteredValue': {'formulaValue': formula},
                    'userEnteredFormat': {'numberFormat': dict(PERCENT_FORMAT)},
                }]}],
                'fields': 'userEnteredValue,userEnteredFormat',
            }
        })

    if dry_run:
        lc = col_to_letter(label_col)
        last_cl = col_to_letter(data_cols[-1])
        print(f"  {sheet_name}: row {target_row + 1} {lc}{target_row + 1}=\"{PAYOUT_LABEL}\"  "
              f"{last_cl}{target_row + 1}={last_formula}")
        verb = 'rewrite formula in' if mode == 'rewrite' else 'insert 1 row +'
        print(f"  {sheet_name}: [DRY RUN] would {verb} write {len(data_cols)} formulas across cols")
        return 'update'

    _retry(lambda: service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={'requests': requests}).execute())
    verb = ('rewrote formula in existing row' if mode == 'rewrite'
            else f'inserted Payout Ratio % row at {target_row + 1},')
    print(f"  {sheet_name}: {verb} {len(data_cols)} formulas")
    return 'update'


def process_spreadsheet(service, spreadsheet_id, label, dry_run):
    print(f"\n{'='*60}\n{label} — {spreadsheet_id[:24]}...\n{'='*60}")
    sheets = get_company_sheets(service, spreadsheet_id)
    if not sheets:
        print("  No company tabs found")
        return 0, 0
    updated = skipped = 0
    for i, (name, sid) in enumerate(sheets):
        if i > 0:
            time.sleep(2)
        if process_tab(service, spreadsheet_id, name, sid, dry_run) == 'update':
            updated += 1
        else:
            skipped += 1
    print(f"  -> Updated: {updated}, Skipped: {skipped}")
    return updated, skipped


def main():
    parser = argparse.ArgumentParser(description='Add Payout Ratio % row to Key Stats')
    parser.add_argument('--spreadsheet-id', help='Target a single spreadsheet')
    parser.add_argument('--dry-run', action='store_true', help='Preview only')
    args = parser.parse_args()

    service = get_service()
    tu = ts = 0
    if args.spreadsheet_id:
        u, s = process_spreadsheet(service, args.spreadsheet_id, 'Custom', args.dry_run)
        tu, ts = tu + u, ts + s
    else:
        for name, info in load_industries().items():
            u, s = process_spreadsheet(service, info['spreadsheet_id'], name, args.dry_run)
            tu, ts = tu + u, ts + s

    action = '[DRY RUN] Would update' if args.dry_run else 'Updated'
    print(f"\n{action}: {tu} tabs, Skipped: {ts} tabs")


if __name__ == '__main__':
    main()
