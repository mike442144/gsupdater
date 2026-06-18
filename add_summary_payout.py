#!/usr/bin/env python3
"""
Add a "平均派息率" (avg payout ratio) row to each time-window section of the
Summary sheet — the 10-year, 5-year and 3-year blocks.

Each Summary section (e.g. "2016-2025年（10年）") lists averaged metrics that
INDIRECT-reference the company tabs over a year-column window. We append one
row at the BOTTOM of each section:

    平均派息率 = AVERAGE of the per-year Payout Ratio % across the window

where the Payout Ratio % is the row added to each company tab's Key Stats
section by add_payout_ratio.py. The window's start/end columns and the payout
row are detected at runtime (not hardcoded), so this works for any year range
or fiscal layout.

The row is appended below the section's last metric ("自由现金流/公司净利润"),
inheriting that ratio row's format. Insertions run BOTTOM-UP (3yr → 5yr → 10yr)
so each earlier section's row index is untouched by later inserts.

Idempotent: if a section already ends in a 派息/payout row, it is left in place
and its formulas are rewritten in place instead of inserting a duplicate.

Usage:
    python add_summary_payout.py --spreadsheet-id <id>            # one spreadsheet
    python add_summary_payout.py --spreadsheet-id <id> --dry-run  # preview only
    python add_summary_payout.py                                  # all industries
    python add_summary_payout.py --dry-run
"""

import os
import re
import sys
import json
import time
import argparse

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')

PAYOUT_LABEL = "平均派息率"
UNIT = "%"
PERCENT_FORMAT = {"type": "PERCENT", "pattern": "0.0%"}

# col-A rows that delimit a company tab's sections.
SECTION_HEADERS = {'income statement', 'balance sheet', 'cash flow',
                   'key stats', 'supplemental', 'business segments'}
# Summary col-A rows that match the time-window section headers (e.g.
# "2016-2025年（10年）"). Other col-A labels ("价格/净利润", "总债务"…) won't match.
WINDOW_HEADER_RE = re.compile(r'^\d{4}-\d{4}年')
RANGE_RE = re.compile(r'\$([A-Z]{1,3})\$(\d+):\$([A-Z]{1,3})\$(\d+)')


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


def find_key_stats_payout_row(service, spreadsheet_id):
    """Row number (1-indexed) of 'Payout Ratio %' inside a company tab's Key
    Stats section. All company tabs share the same layout, so the first one
    with the row answers for the whole spreadsheet. Returns None if absent."""
    meta = _retry(lambda: service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title))').execute())
    for s in meta.get('sheets', []):
        title = s['properties']['title']
        if title == 'Summary' or '资本结构' in title or '经营数据' in title:
            continue
        result = _retry(lambda: service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=f"'{title}'!A1:C40").execute())
        rows = result.get('values', [])
        section = None
        for i, row in enumerate(rows[:40]):
            a = row[0].strip().lower() if row and row[0] else ''
            if a in SECTION_HEADERS:
                section = a
                continue
            b = row[1].strip() if len(row) > 1 and row[1] else ''
            c = row[2].strip() if len(row) > 2 and row[2] else ''
            val = c or b
            if val.lower() == 'payout ratio %' and section == 'key stats':
                return i + 1
    return None


def parse_metric(formula):
    """From a metric formula like =sum(INDIRECT("'"&F$2&"财务'!$N$5:$W$5")),
    return (prefix, start_col, end_col) where prefix is the INDIRECT argument
    text up to (not including) the range, e.g. "\"'\"&F$2&\"财务'!". The prefix
    carries the column letter AND the per-sheet tab-name suffix (e.g. 财务), so a
    payout formula built from it references the right tab. Returns None if the
    formula has no INDIRECT range we can reuse.
    """
    idx = formula.upper().find('INDIRECT(')
    if idx < 0:
        return None
    arg_start = idx + len('INDIRECT(')
    m = RANGE_RE.search(formula, arg_start)
    if not m:
        return None
    return formula[arg_start:m.start()], m.group(1), m.group(3)


def prefix_for_col(prefix, src_letter, tgt_letter):
    """Adapt a column-specific prefix to another column by swapping the row-2
    reference (F$2 -> E$2). Used when a data column itself has no metric formula
    in a section (sparse sheets) so we borrow a sibling's prefix."""
    if src_letter == tgt_letter:
        return prefix
    return prefix.replace(src_letter + '$2', tgt_letter + '$2')


def scan_summary(service, spreadsheet_id):
    """Inspect the Summary sheet.

    Returns (summary_sheet_id, sections, data_cols) where sections is a list of
    dicts {header_row, last_metric_row, start_col, end_col, ref_prefix,
    ref_letter} and data_cols is the list of 0-indexed company data columns
    (>= E). Window columns/prefix are derived per-section from whichever data
    columns actually carry metric formulas, so sparse layouts and per-sheet
    tab-name suffixes (财务) are handled.
    """
    meta = _retry(lambda: service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=["'Summary'!A1:ZZ200"],
        includeGridData=True,
        fields='sheets(properties(sheetId),data.rowData.values.userEnteredValue)'
    ).execute())
    sheet = meta['sheets'][0]
    summary_sheet_id = sheet['properties']['sheetId']
    row_data = sheet['data'][0].get('rowData', [])

    # data columns: row 1 (0-idx row 0) cells from E onward that are non-empty
    data_cols = []
    r0 = row_data[0].get('values', []) if row_data else []
    for j in range(4, len(r0)):
        uev = r0[j].get('userEnteredValue', {}) if j < len(r0) else {}
        if uev:
            data_cols.append(j)
    if not data_cols:
        return summary_sheet_id, [], []

    def cell(row, i):
        vals = row.get('values', []) if row else []
        return vals[i].get('userEnteredValue', {}) if i < len(vals) else {}

    sections = []
    n = len(row_data)
    i = 0
    while i < n:
        a = cell(row_data[i], 0).get('stringValue', '')
        if a and WINDOW_HEADER_RE.match(a):
            header_row = i
            # metric block: rows after header where col B has a label
            last_metric = header_row
            j = header_row + 1
            while j < n:
                b_val = cell(row_data[j], 1).get('stringValue', '')
                a_val = cell(row_data[j], 0).get('stringValue', '')
                if a_val and WINDOW_HEADER_RE.match(a_val):
                    break  # next window header
                if b_val:
                    last_metric = j
                    j += 1
                else:
                    break
            # window cols + prefix from the first data column whose metric rows
            # in this section yield a parseable INDIRECT range.
            ref = None
            for dc in data_cols:
                for r in range(header_row + 1, last_metric + 1):
                    fv = cell(row_data[r], dc).get('formulaValue', '')
                    if fv:
                        parsed = parse_metric(fv)
                        if parsed:
                            ref = parsed
                            break
                if ref:
                    break
            if ref:
                sections.append({
                    'header_row': header_row,
                    'last_metric_row': last_metric,
                    'start_col': ref[1],
                    'end_col': ref[2],
                    'ref_prefix': ref[0],
                    'ref_letter': col_to_letter(dc),
                })
            else:
                sections.append({
                    'header_row': header_row,
                    'last_metric_row': last_metric,
                    'start_col': None,
                    'end_col': None,
                    'ref_prefix': None,
                    'ref_letter': None,
                })
            i = j
        else:
            i += 1
    return summary_sheet_id, sections, data_cols


def has_payout_in_section(row_data, header_row, last_metric_row):
    """True if the section already contains a 派息/payout label row."""
    for j in range(header_row + 1, last_metric_row + 1):
        vals = row_data[j].get('values', []) if j < len(row_data) else []
        if len(vals) > 1:
            b = vals[1].get('userEnteredValue', {}).get('stringValue', '')
            if b and ('派息' in b or 'payout' in b.lower()):
                return j
    return None


def build_formula(prefix, start_col, end_col, payout_row):
    """prefix already carries the column letter + tab-name suffix up to the
    range, e.g. "\"'\"&F$2&\"财务'!\". We just append the range on payout_row
    and wrap in AVERAGE(INDIRECT(...))."""
    pr = payout_row  # 1-indexed
    # The range sits INSIDE the source's string literal (e.g.
    # "财务'!$N$5:$W$5"), so there's no closing quote until after the range.
    # Append the range on payout_row, then close the string + both functions.
    return f'=AVERAGE(INDIRECT({prefix}${start_col}${pr}:${end_col}${pr}"))'


def process_spreadsheet(service, spreadsheet_id, label, dry_run):
    print(f"\n{'='*60}\n{label} — {spreadsheet_id[:24]}...\n{'='*60}")

    payout_row = find_key_stats_payout_row(service, spreadsheet_id)
    if payout_row is None:
        print("  SKIP — no 'Payout Ratio %' row in any company tab Key Stats "
              "(run add_payout_ratio.py first)")
        return 0, 1
    print(f"  Key Stats Payout Ratio % row: {payout_row}")

    summary_sheet_id, sections, data_cols = scan_summary(service, spreadsheet_id)
    if not sections:
        print("  SKIP — no time-window sections found on Summary")
        return 0, 1
    if not data_cols:
        print("  SKIP — no company data columns found on Summary")
        return 0, 1

    # need raw row data for the payout-exists check
    rd = _retry(lambda: service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, ranges=["'Summary'"], includeGridData=True,
        fields='sheets.data.rowData.values.userEnteredValue').execute())
    row_data = rd['sheets'][0]['data'][0].get('rowData', [])

    requests = []
    inserted = 0
    rewritten = 0
    preview = []

    # BOTTOM-UP so upper section indices stay valid across inserts.
    for sec in sorted(sections, key=lambda s: s['header_row'], reverse=True):
        if sec['start_col'] is None or sec['ref_prefix'] is None:
            print(f"  section @ row {sec['header_row']+1}: SKIP — "
                  "could not parse window columns")
            continue
        existing = has_payout_in_section(
            row_data, sec['header_row'], sec['last_metric_row'])

        if existing is not None:
            target_row = existing  # 0-indexed
            mode = 'rewrite'
        else:
            target_row = sec['last_metric_row'] + 1
            mode = 'insert'
            requests.append({
                'insertDimension': {
                    'range': {'sheetId': summary_sheet_id, 'dimension': 'ROWS',
                              'startIndex': target_row, 'endIndex': target_row + 1},
                    'inheritFromBefore': True,
                }
            })

        # label (col B / idx 1)
        requests.append({
            'updateCells': {
                'range': {'sheetId': summary_sheet_id,
                          'startRowIndex': target_row, 'endRowIndex': target_row + 1,
                          'startColumnIndex': 1, 'endColumnIndex': 2},
                'rows': [{'values': [{'userEnteredValue': {'stringValue': PAYOUT_LABEL}}]}],
                'fields': 'userEnteredValue',
            }
        })
        # unit (col D / idx 3)
        requests.append({
            'updateCells': {
                'range': {'sheetId': summary_sheet_id,
                          'startRowIndex': target_row, 'endRowIndex': target_row + 1,
                          'startColumnIndex': 3, 'endColumnIndex': 4},
                'rows': [{'values': [{'userEnteredValue': {'stringValue': UNIT}}]}],
                'fields': 'userEnteredValue',
            }
        })
        # formulas across data columns — adapt the section's reference prefix to
        # each column's own letter (handles sparse columns + per-sheet tab suffix)
        for dc in data_cols:
            cl = col_to_letter(dc)
            prefix = prefix_for_col(sec['ref_prefix'], sec['ref_letter'], cl)
            formula = build_formula(prefix, sec['start_col'], sec['end_col'], payout_row)
            requests.append({
                'updateCells': {
                    'range': {'sheetId': summary_sheet_id,
                              'startRowIndex': target_row, 'endRowIndex': target_row + 1,
                              'startColumnIndex': dc, 'endColumnIndex': dc + 1},
                    'rows': [{'values': [{
                        'userEnteredValue': {'formulaValue': formula},
                        'userEnteredFormat': {'numberFormat': dict(PERCENT_FORMAT)},
                    }]}],
                    'fields': 'userEnteredValue,userEnteredFormat',
                }
            })

        if mode == 'insert':
            inserted += 1
        else:
            rewritten += 1
        first_dc = col_to_letter(data_cols[0])
        last_dc = col_to_letter(data_cols[-1])
        hdr = (row_data[sec['header_row']].get('values', [{}])[0]
               .get('userEnteredValue', {}).get('stringValue', ''))
        preview.append((hdr, target_row + 1, mode,
                        f"{first_dc}{target_row+1}…{last_dc}{target_row+1}"))

    for hdr, r, mode, span in preview:
        verb = 'rewrite' if mode == 'rewrite' else 'insert '
        print(f"  [{verb}] row {r:>3} {span}  <- {hdr}")

    if not requests:
        print("  Nothing to do")
        return 0, 1

    if dry_run:
        print(f"  [DRY RUN] {len(requests)} request(s) would be sent "
              f"(inserted {inserted}, rewritten {rewritten})")
        return 1, 0

    _retry(lambda: service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={'requests': requests}).execute())
    print(f"  -> Inserted {inserted}, Rewritten {rewritten} sections")
    return 1, 0


def main():
    parser = argparse.ArgumentParser(
        description='Add 平均派息率 row to each Summary time-window section')
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
    print(f"\n{action}: {tu} spreadsheets, Skipped: {ts} spreadsheets")


if __name__ == '__main__':
    main()
