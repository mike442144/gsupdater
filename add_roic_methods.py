#!/usr/bin/env python3
"""
Add two additional ROIC rows to each company tab's Key Stats section, right
below the existing ROIC row (which uses the capital-source / financing method):

  - "ROIC (资产法)"      operating-asset (investing) approach, NOPAT / avg invested
                         capital where invested capital = Total Assets - excess cash
                         - non-interest-bearing current liabilities. Two-period avg.
  - "ROIC (Greenblatt)"  pre-tax EBIT / (operating working capital + net fixed assets),
                         tangible capital employed, point-in-time.

The existing ROIC sits at the bottom of the 盈利指标 sub-group. Summary INDIRECT
formulas only reference Key Stats rows up to "Total Revenue" (above ROIC), so
inserting rows after ROIC does NOT shift any Summary-referenced row — no
fix_summary_formulas pass is needed.

Usage:
    python add_roic_methods.py --spreadsheet-id <id>            # one spreadsheet
    python add_roic_methods.py --spreadsheet-id <id> --dry-run  # preview only
    python add_roic_methods.py                                  # all industries
    python add_roic_methods.py --dry-run
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

# New items in display order, with formula templates. __C__ = current data column
# letter, __PC__ = previous data column letter, {Item} = row of that item name.
NEW_ITEMS = [
    # {?Item} = optional: resolves to 0 when CIQ omits that (zero-valued) line.
    ("ROIC (资产法)",
     "=__C__{EBIT}*(1-__C__{!Effective Tax Rate %})/"
     "((__C__{Total Assets}-__C__{?Total Cash & ST Investments}-"
     "(__C__{Total Current Liabilities}-__C__{?Short-term Borrowings}-__C__{?Curr. Port. of Leases}))+"
     "(__PC__{Total Assets}-__PC__{?Total Cash & ST Investments}-"
     "(__PC__{Total Current Liabilities}-__PC__{?Short-term Borrowings}-__PC__{?Curr. Port. of Leases})))*2"),
    # Textbook Greenblatt: current EBIT over BEGINNING-of-period tangible capital
    # (i.e. the prior column's balance sheet, __PC__) — no averaging.
    ("ROIC (Greenblatt)",
     "=__C__{EBIT}/"
     "((__PC__{Total Current Assets}-__PC__{?Total Cash & ST Investments})-"
     "(__PC__{Total Current Liabilities}-__PC__{?Short-term Borrowings}-__PC__{?Curr. Port. of Leases})+"
     "__PC__{Net Property, Plant & Equipment})"),
]
PERCENT_FORMAT = {"type": "PERCENT", "pattern": "0.0%"}

# The pre-existing ROIC row uses the capital-source (financing) method; relabel it
# so the three ROIC rows read in parallel. Anchor detection accepts both names so
# the script stays re-runnable after the relabel.
NEW_ROIC_LABEL = "ROIC (资本来源法)"

# Capital-source ROIC, used only when a tab has NO base ROIC row at all: then we
# anchor on "Total Revenue" and insert all THREE rows (this one + NEW_ITEMS).
CAPITAL_SOURCE_ITEM = (
    NEW_ROIC_LABEL,
    "=__C__{EBIT}*(1-__C__{!Effective Tax Rate %})/"
    "(__C__{Net Debt}+__C__{Common Equity}+__C__{!Minority Interest}+"
    "__PC__{Net Debt}+__PC__{Common Equity}+__PC__{!Minority Interest})*2")
ALL_ITEMS = [CAPITAL_SOURCE_ITEM] + NEW_ITEMS
BASE_ROIC_NAMES = {"roic", NEW_ROIC_LABEL.lower()}


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
    """Return (item_to_row, ks_items, roic_row, label_col).

    item_to_row: full name->0idx row map (last-wins, all sections) for formula refs.
    ks_items: Key Stats name->0idx row map.
    roic_row: 0-indexed row of the 'ROIC' Key Stats item (None if absent).
    label_col: 0-indexed column where the ROIC label sits (1=B or 2=C).
    """
    result = _retry(lambda: service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"'{sheet_name}'!A1:C300"
    ).execute())
    rows = result.get('values', [])

    item_to_row = {}
    ks_items = {}
    roic_row = None
    roic_label = None
    total_rev_row = None
    label_col = 2
    tr_label_col = 2
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
            ks_items[item_val.lower()] = i
            if item_val.lower() in BASE_ROIC_NAMES:
                roic_row = i
                roic_label = item_val
                label_col = 2 if c else 1
            elif item_val.lower() == 'total revenue':
                total_rev_row = i
                tr_label_col = 2 if c else 1

    # When there's no base ROIC, the Total Revenue row is the anchor; use its column.
    if roic_row is None:
        label_col = tr_label_col
    return item_to_row, ks_items, roic_row, roic_label, total_rev_row, label_col


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


def resolve_formula(template, col_letter, prev_col_letter, item_to_row, sheet_name):
    """Resolve a template of __C__{Item} / __PC__{Item} tokens to a GS formula.

    Marker prefixes: '?Item' optional (a missing, CIQ-omitted line resolves to a
    bare 0); '!Item' required but N()-wrapped. Both N()-wrap when the row exists,
    coercing CIQ's '-' nil marker (text) to 0. A missing required item -> <col>1.
    """
    missing = []

    def repl(m):
        col = col_letter if m.group(1) == '__C__' else prev_col_letter
        name = m.group(2)
        marker = ''
        if name[:1] in ('?', '!'):
            marker, name = name[0], name[1:]
        row = item_to_row.get(name.lower())
        if row is None:
            if marker == '?':
                return '0'
            missing.append(name)
            return col + '1'
        if marker:
            return f'N({col}{row + 1})'
        return f'{col}{row + 1}'

    formula = re.sub(r'(__C__|__PC__)\{([?!]?[^}]+)\}', repl, template)
    if missing:
        print(f"    WARNING [{sheet_name}]: unresolved items {sorted(set(missing))}")
    return formula


def process_tab(service, spreadsheet_id, sheet_name, sheet_id, dry_run):
    item_to_row, ks_items, roic_row, roic_label, total_rev_row, label_col = scan_tab(
        service, spreadsheet_id, sheet_name)

    # Anchor: after the base ROIC row if present (add the two new methods); else
    # after Total Revenue (add all three, including capital-source). No anchor → skip.
    if roic_row is not None:
        items, anchor_row, do_relabel = NEW_ITEMS, roic_row, True
    elif total_rev_row is not None:
        items, anchor_row, do_relabel = ALL_ITEMS, total_rev_row, False
    else:
        print(f"  {sheet_name}: SKIP — no ROIC or Total Revenue row in Key Stats")
        return 'skip'

    n = len(items)
    existing = [ks_items.get(label.lower()) for label, _ in items]
    already_present = all(r is not None for r in existing)

    data_cols = find_data_columns(service, spreadsheet_id, sheet_name)
    if not data_cols:
        print(f"  {sheet_name}: SKIP — no data columns found")
        return 'skip'

    requests = []
    # Relabel the base ROIC row (capital-source method) — row index unaffected by
    # any insertion below it.
    relabel = do_relabel and roic_label != NEW_ROIC_LABEL
    if relabel:
        requests.append({
            'updateCells': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': roic_row, 'endRowIndex': roic_row + 1,
                          'startColumnIndex': label_col, 'endColumnIndex': label_col + 1},
                'rows': [{'values': [{'userEnteredValue': {'stringValue': NEW_ROIC_LABEL}}]}],
                'fields': 'userEnteredValue',
            }
        })

    if already_present:
        # Re-runnable: rewrite formulas in the existing rows, no insert/shift.
        target_rows = existing
        shifted = item_to_row
        mode = 'rewrite'
    else:
        insert_at = anchor_row + 1  # 0-indexed row index for the first new row
        target_rows = [insert_at + k for k in range(n)]
        # After inserting n rows at `insert_at`, items at original row >= insert_at
        # shift down by n. The anchor and everything above are unaffected.
        shifted = {name: (r + n if r >= insert_at else r) for name, r in item_to_row.items()}
        mode = 'insert'
        requests.append({
            'insertDimension': {
                'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                          'startIndex': insert_at, 'endIndex': insert_at + n},
                'inheritFromBefore': True,
            }
        })

    # labels + formulas for each new row
    preview = []
    for k, (label, template) in enumerate(items):
        new_row = target_rows[k]  # 0-indexed
        # label
        requests.append({
            'updateCells': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': new_row, 'endRowIndex': new_row + 1,
                          'startColumnIndex': label_col, 'endColumnIndex': label_col + 1},
                'rows': [{'values': [{'userEnteredValue': {'stringValue': label}}]}],
                'fields': 'userEnteredValue',
            }
        })
        # formulas across data columns
        for ci, dc in enumerate(data_cols):
            cl = col_to_letter(dc)
            pc = col_to_letter(data_cols[ci - 1]) if ci > 0 else col_to_letter(max(0, dc - 1))
            formula = resolve_formula(template, cl, pc, shifted, sheet_name)
            requests.append({
                'updateCells': {
                    'range': {'sheetId': sheet_id,
                              'startRowIndex': new_row, 'endRowIndex': new_row + 1,
                              'startColumnIndex': dc, 'endColumnIndex': dc + 1},
                    'rows': [{'values': [{
                        'userEnteredValue': {'formulaValue': formula},
                        'userEnteredFormat': {'numberFormat': dict(PERCENT_FORMAT)},
                    }]}],
                    'fields': 'userEnteredValue,userEnteredFormat',
                }
            })
            if ci == len(data_cols) - 1:
                preview.append((label, col_to_letter(label_col), new_row + 1, cl, formula))

    relabel_note = f' + relabel base ROIC→"{NEW_ROIC_LABEL}"' if relabel else ''
    if dry_run:
        for label, lc, r1, cl, formula in preview:
            print(f"  {sheet_name}: row {r1} {lc}{r1}=\"{label}\"  {cl}{r1}={formula}")
        verb = 'rewrite formulas in' if mode == 'rewrite' else f'insert {n} rows +'
        print(f"  {sheet_name}: [DRY RUN] would {verb} write "
              f"{len(data_cols)} formulas x{n} across cols{relabel_note}")
        return 'update'

    _retry(lambda: service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={'requests': requests}).execute())
    verb = 'rewrote formulas in' if mode == 'rewrite' else f'inserted {n} ROIC rows at {target_rows[0] + 1},'
    print(f"  {sheet_name}: {verb} {len(data_cols)} formulas each{relabel_note}")
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
    parser = argparse.ArgumentParser(description='Add 资产法 & Greenblatt ROIC rows to Key Stats')
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
