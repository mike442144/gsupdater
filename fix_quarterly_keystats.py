#!/usr/bin/env python3
"""
Fix quarterly-column formulas in company tabs of a spreadsheet.

create_company_tab.py / add_yoy_section.py / add_roic_methods.py /
add_payout_ratio.py used to treat quarter columns ("Q1 2021") like annual
columns, so:
- YoY rows divided by the previous quarter column (QoQ, not YoY); the first
  quarter year even referenced the LTM column.
- ROIC and Payout Ratio formulas were filled into quarter columns where they
  are not meaningful.

Per company tab (every '<name>财务' tab with quarterly IS data — tabs whose
quarter columns are empty scaffolding are skipped, and already-correct
formulas are left untouched):
- Rewrites YoY formulas in quarter columns to use the same quarter of the
  prior year as base. Formula shape and item refs are preserved — only the
  base column letter changes.
- Clears YoY cells whose quarter has no prior-year twin column.
- Clears ROIC* and Payout Ratio* rows (Key Stats and IS sections) in quarter
  columns.

Usage:
    python fix_quarterly_keystats.py --spreadsheet-id <ID> [--dry-run]
"""

import re
import time
import argparse

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import HttpRequest
import os

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')

# ── API pacing (same policy as update_financials.py) ────────────────────────
_orig_execute = HttpRequest.execute
_MIN_CALL_INTERVAL = 1.1
_last_call_ts = [0.0]


def _paced_execute(self, *args, **kwargs):
    for attempt in range(5):
        wait = _MIN_CALL_INTERVAL - (time.monotonic() - _last_call_ts[0])
        if wait > 0:
            time.sleep(wait)
        try:
            resp = _orig_execute(self, *args, **kwargs)
            _last_call_ts[0] = time.monotonic()
            return resp
        except HttpError as e:
            _last_call_ts[0] = time.monotonic()
            if e.resp.status in (429, 500, 503) and attempt < 4:
                pause = 20 * (attempt + 1)
                print(f"  ... HTTP {e.resp.status}, retrying in {pause}s")
                time.sleep(pause)
                continue
            raise


HttpRequest.execute = _paced_execute

YOY_FORMULA_RE = re.compile(
    r'^=(?:IFERROR\(\s*)?([A-Z]{1,3})(\d+)\s*/\s*([A-Z]{1,3})(\d+)\s*-\s*1\s*,?\s*\)?$')


def col_to_letter(col_idx):
    result = ''
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def quarter_header_key(text):
    m = re.match(r'^Q([1-4]) (\d{4})$', str(text).strip())
    return (int(m.group(1)), int(m.group(2))) if m else None


def get_service():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    return build('sheets', 'v4', credentials=creds)


def company_tabs(service, spreadsheet_id):
    """All tab titles ending in '财务' (company tabs)."""
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId,gridProperties(rowCount,columnCount)))'
    ).execute()
    tabs = []
    for s in meta.get('sheets', []):
        title = s['properties']['title']
        if title.endswith('财务'):
            gp = s['properties'].get('gridProperties', {})
            tabs.append((title, s['properties']['sheetId'],
                         gp.get('rowCount', 1000), gp.get('columnCount', 26)))
    return tabs


IS_SECTION_END = ('balance sheet', 'cash flow', 'key stats', 'supplemental',
                  'multiples', 'ratios', 'segments', 'capitalization')


def find_is_section(grid):
    """(start, end) row indexes of the Income Statement section, or (None, None)."""
    start = None
    for i, row in enumerate(grid):
        a = row[0].strip().lower() if row and row[0] else ''
        if a == 'income statement':
            start = i + 1
        elif start is not None and a in IS_SECTION_END:
            return start, i
    return (start, len(grid)) if start is not None else (None, None)


def classify_row(row):
    """Row label (C preferred — Key Stats layout — else B) → 'yoy'/'roic'/'payout'/None."""
    b = row[1].strip() if len(row) > 1 and row[1] else ''
    c = row[2].strip() if len(row) > 2 and row[2] else ''
    label = (c or b).lower()
    if label.endswith(' yoy'):
        return 'yoy'
    if label.startswith('roic'):
        return 'roic'
    if label.startswith('payout ratio'):
        return 'payout'
    return None


def fix_tab(service, spreadsheet_id, title, sheet_id, row_count, col_count, dry_run):
    header = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"'{title}'!A1:{col_to_letter(col_count - 1)}1"
    ).execute().get('values', [[]])[0]

    qcols = {}
    for j, h in enumerate(header):
        qkey = quarter_header_key(h)
        if qkey:
            qcols[qkey] = j
    if not qcols:
        print(f"  {title}: no quarter columns, skip")
        return

    grid = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{title}'!A1:{col_to_letter(col_count - 1)}{row_count}",
        valueRenderOption='FORMULA').execute().get('values', [])

    is_start, is_end = find_is_section(grid)
    if is_start is None:
        print(f"  {title}: no Income Statement section, skip")
        return
    # Only quarter columns that actually hold quarterly IS data — tabs whose
    # quarterly columns were scaffolded but never filled are left alone.
    qcols = {q: j for q, j in qcols.items()
             if any(i < len(grid) and j < len(grid[i]) and grid[i][j] != ''
                    for i in range(is_start, is_end))}
    if not qcols:
        print(f"  {title}: quarter columns hold no IS data, skip")
        return

    requests = []
    stats = {'yoy_fixed': 0, 'yoy_filled': 0, 'yoy_cleared': 0, 'metric_cleared': 0}
    samples = []

    for i, row in enumerate(grid):
        kind = classify_row(row)
        if kind is None:
            continue
        # For YoY rows, derive the referenced item row (and IFERROR shape) from
        # any existing formula in the row — used to backfill empty quarter cells.
        yoy_tmpl = None
        if kind == 'yoy':
            for cell in row:
                m = YOY_FORMULA_RE.match(str(cell)) if cell else None
                if m:
                    yoy_tmpl = (int(m.group(2)), 'IFERROR' in str(cell))
                    break
        for qkey, j in sorted(qcols.items(), key=lambda kv: (kv[0][1], kv[0][0])):
            cell = str(row[j]) if j < len(row) and row[j] != '' else ''
            if kind in ('roic', 'payout'):
                if cell:
                    requests.append(_clear(sheet_id, i, j))
                    stats['metric_cleared'] += 1
                continue
            # YoY row
            base_j = qcols.get((qkey[0], qkey[1] - 1))
            if not cell:
                # Backfill: some tabs never had quarterly YoY formulas
                if base_j is None or yoy_tmpl is None:
                    continue
                ref_row, has_iferror = yoy_tmpl
                cur, base = col_to_letter(j), col_to_letter(base_j)
                new_formula = (f'=IFERROR({cur}{ref_row}/{base}{ref_row}-1,)'
                               if has_iferror else f'={cur}{ref_row}/{base}{ref_row}-1')
                requests.append({
                    'updateCells': {
                        'range': {'sheetId': sheet_id,
                                  'startRowIndex': i, 'endRowIndex': i + 1,
                                  'startColumnIndex': j, 'endColumnIndex': j + 1},
                        'rows': [{'values': [{'userEnteredValue': {'formulaValue': new_formula},
                                              'userEnteredFormat': {'numberFormat': {'type': 'PERCENT', 'pattern': '0.0%'}}}]}],
                        'fields': 'userEnteredValue,userEnteredFormat',
                    }
                })
                stats['yoy_filled'] += 1
                if len(samples) < 3:
                    samples.append(f"    {cur}{i + 1}: (empty) -> {new_formula}")
                continue
            m = YOY_FORMULA_RE.match(cell)
            if not m or m.group(1) != col_to_letter(j):
                print(f"    WARNING: {title} row {i + 1} col {col_to_letter(j)} "
                      f"unexpected formula: {cell!r} — skipped")
                continue
            if base_j is None:
                requests.append(_clear(sheet_id, i, j))
                stats['yoy_cleared'] += 1
                continue
            base_letter = col_to_letter(base_j)
            if m.group(3) == base_letter:
                continue
            new_formula = cell[:m.start(3)] + base_letter + cell[m.end(3):]
            requests.append({
                'updateCells': {
                    'range': {'sheetId': sheet_id,
                              'startRowIndex': i, 'endRowIndex': i + 1,
                              'startColumnIndex': j, 'endColumnIndex': j + 1},
                    'rows': [{'values': [{'userEnteredValue': {'formulaValue': new_formula}}]}],
                    'fields': 'userEnteredValue',
                }
            })
            stats['yoy_fixed'] += 1
            if len(samples) < 3:
                samples.append(f"    {col_to_letter(j)}{i + 1}: {cell} -> {new_formula}")

    print(f"  {title}: YoY fixed {stats['yoy_fixed']}, filled {stats['yoy_filled']}, "
          f"cleared {stats['yoy_cleared']}, ROIC/Payout cleared {stats['metric_cleared']}")
    for s in samples:
        print(s)
    if dry_run or not requests:
        return
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={'requests': requests}).execute()
    print(f"    ✓ {len(requests)} updates applied")


def _clear(sheet_id, row, col):
    return {
        'updateCells': {
            'range': {'sheetId': sheet_id,
                      'startRowIndex': row, 'endRowIndex': row + 1,
                      'startColumnIndex': col, 'endColumnIndex': col + 1},
            'rows': [{'values': [{}]}],
            'fields': 'userEnteredValue',
        }
    }


def main():
    parser = argparse.ArgumentParser(description='Fix quarterly-column YoY/ROIC/Payout formulas')
    parser.add_argument('--spreadsheet-id', required=True)
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()

    service = get_service()
    for title, sheet_id, row_count, col_count in company_tabs(service, args.spreadsheet_id):
        fix_tab(service, args.spreadsheet_id, title, sheet_id, row_count, col_count, args.dry_run)


if __name__ == '__main__':
    main()
