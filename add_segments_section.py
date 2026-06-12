#!/usr/bin/env python3
"""
Build a "<name>运营数据" segments tab for companies in a Google Spreadsheet.

Data source: eastmoney 主营构成 API via ~/Projects/tinyant/eastmoney/mainop.js
(works for A-share codes only — mainop.js builds .SZ/.SH SECUCODEs).

A dedicated tab is created per company (e.g. 茅台 -> 茅台运营数据). Its year
columns mirror the company's 财务 tab (so the two line up); annual segment data
is written underneath in a 3-level hierarchy:

    A            B            C            | 2007 ... 2024 2025
    主营构成
    按行业
                 酒类
                              营业收入       |  ...   170,612
                              收入占比       |  ...    99.8%
                              毛利率         |  ...    92.0%
                              营业成本       |  ...    13,630
                              营业利润       |  ...   156,982
    按产品
                 茅台酒  ...
    按地区
                 国内 / 国外 ...

Annual data only (REPORT_DATE == YYYY-12-31). The tab is rebuilt from scratch on
every run (cleared then rewritten), so the script is idempotent. Items with no
data inside the mirrored year columns (legacy pre-range labels) are dropped.

Usage:
    python add_segments_section.py --sheet-id <id> --codes 600519 --dry-run
    python add_segments_section.py --sheet-id <id> --codes 600519
"""

import sys
import os
import re
import csv
import argparse
import subprocess
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ── Configuration ──────────────────────────────────────────────────────────

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')
EASTMONEY_DIR = os.path.expanduser('~/Projects/tinyant/eastmoney')
MAINOP_SCRIPT = os.path.join(EASTMONEY_DIR, 'mainop.js')

TAB_SUFFIX = '运营数据'      # 茅台财务 -> 茅台运营数据
TAB_TITLE = '主营构成'       # A1 title on the new tab
UNIT_HEADER = 'Unit'         # mirrors the 财务 tab's Unit column (header only)
YEAR_START_COL = 4           # year columns begin at column E; Unit sits in col D
LABEL_COL_WIDTH = 80         # A-D width, matching create_company_tab.py's column C

# Classification render order (MAINOP_TYPE_NAME)
CLASSIFICATION_ORDER = ['按行业', '按产品', '按地区']

# (label, CSV column, number-format type, number-format pattern)
# Ratio columns from the API are already decimals (0.92 => 92.0%).
METRICS = [
    ('营业收入', 'MAIN_BUSINESS_INCOME', 'NUMBER', '#,##0'),
    ('收入占比', 'MBI_RATIO', 'PERCENT', '0.0%'),
    ('毛利率', 'GROSS_RPOFIT_RATIO', 'PERCENT', '0.0%'),
    ('营业成本', 'MAIN_BUSINESS_COST', 'NUMBER', '#,##0'),
    ('营业利润', 'MAIN_BUSINESS_RPOFIT', 'NUMBER', '#,##0'),
]


# ── Helpers ────────────────────────────────────────────────────────────────

def col_to_letter(col_idx):
    """Convert 0-indexed column number to Excel-style letter."""
    result = ''
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def is_ashare(code):
    return bool(re.match(r'^\d{6}$', code))


def report_date_to_year(date_str):
    """'2024-12-31 00:00:00' -> '2024' for annual rows; None otherwise."""
    if not date_str:
        return None
    try:
        dt = datetime.strptime(date_str.split()[0], '%Y-%m-%d')
        if dt.month == 12 and dt.day == 31:
            return str(dt.year)
    except ValueError:
        pass
    return None


def parse_float(v):
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


# ── eastmoney mainop.js ────────────────────────────────────────────────────

def run_mainop(codes):
    """Run mainop.js for comma-separated A-share codes; return CSV rows (dicts)."""
    for f in Path(EASTMONEY_DIR).glob('data/eastmoney_mainop_*.csv'):
        f.unlink()

    # --count is the per-request page size; the default (200) truncates long
    # histories, so request enough to cover all periods back to ~2001.
    cmd = ['node', MAINOP_SCRIPT, '--codes', codes, '--count', '1000']
    result = subprocess.run(cmd, cwd=EASTMONEY_DIR, capture_output=True,
                            text=True, timeout=180)
    if result.returncode != 0:
        print(f"  ERROR: mainop.js failed: {result.stderr}")
        return []

    csv_files = list(Path(EASTMONEY_DIR).glob('data/eastmoney_mainop_*.csv'))
    if not csv_files:
        print("  ERROR: no mainop CSV output found")
        return []

    rows = []
    with open(csv_files[0], 'r', encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            rows.append(row)
    return rows


def build_segments(csv_rows, code):
    """Shape CSV rows for one code into a nested structure.

    Returns:
        {classification: {
            'items': [item_name, ...]  (ordered by latest-year rank),
            'data':  {item: {year: {metric_label: value}}},
        }}
    """
    raw = {}    # raw[classification][item][year][metric_label] = value
    rank = {}   # rank[classification][item] = (year, rank) for latest year seen

    for row in csv_rows:
        if str(row.get('SECURITY_CODE', '')).strip() != str(code):
            continue
        year = report_date_to_year(row.get('REPORT_DATE'))
        if not year:
            continue
        cls = (row.get('MAINOP_TYPE_NAME') or '').strip()
        item = (row.get('ITEM_NAME') or '').strip()
        if not cls or not item:
            continue

        ydata = raw.setdefault(cls, {}).setdefault(item, {}).setdefault(year, {})
        for label, col, _t, _p in METRICS:
            val = parse_float(row.get(col))
            if val is not None:
                ydata[label] = val

        r = parse_float(row.get('RANK'))
        prev = rank.setdefault(cls, {}).get(item)
        if r is not None and (prev is None or year > prev[0]):
            rank[cls][item] = (year, r)

    result = {}
    for cls, items in raw.items():
        ordered = sorted(items.keys(),
                         key=lambda it: (rank.get(cls, {}).get(it, ('', 1e9))[1], it))
        result[cls] = {'items': ordered, 'data': items}
    return result


# ── Google Sheets ──────────────────────────────────────────────────────────

def get_service():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    return build('sheets', 'v4', credentials=creds)


def get_summary_mapping(service, spreadsheet_id):
    """Stock code -> '<name>财务' sheet name, from the Summary tab."""
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="'Summary'!A1:AZ3"
    ).execute()
    rows = result.get('values', [])
    if len(rows) < 2:
        return {}
    codes_row, names_row = rows[0], rows[1]
    mapping = {}
    for j in range(len(codes_row)):
        cde = str(codes_row[j]).strip() if j < len(codes_row) else ''
        name = str(names_row[j]).strip() if j < len(names_row) else ''
        if cde and name:
            mapping[cde] = f"{name}财务"
    return mapping


def resolve_sheet(service, spreadsheet_id, sheet_name):
    """Return (sheet_name, sheetId, rowCount, colCount), with 财务-suffix fallback."""
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId,gridProperties(rowCount,columnCount)))'
    ).execute()
    by_title = {s['properties']['title']: s['properties'] for s in meta.get('sheets', [])}

    candidates = [sheet_name]
    if sheet_name.endswith('财务'):
        candidates.append(sheet_name[:-2])
    for cand in candidates:
        if cand in by_title:
            p = by_title[cand]
            g = p.get('gridProperties', {})
            if cand != sheet_name:
                print(f"  NOTE: using sheet '{cand}' (without 财务 suffix)")
            return cand, p['sheetId'], g.get('rowCount', 1000), g.get('columnCount', 26)
    return None, None, None, None


def get_fiscal_years(service, spreadsheet_id, fin_sheet_name, col_count):
    """Ascending list of 4-digit year headers from the 财务 tab's row 1."""
    end_col = col_to_letter(col_count - 1)
    hdr = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{fin_sheet_name}'!A1:{end_col}1"
    ).execute().get('values', [[]])
    header_row = hdr[0] if hdr else []
    years = {str(v).strip() for v in header_row if re.match(r'^\d{4}$', str(v).strip())}
    return sorted(years)


def find_tab(service, spreadsheet_id, tab_name):
    """Return (sheetId, rowCount, colCount) for tab_name, or (None, ...)."""
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId,gridProperties(rowCount,columnCount)))'
    ).execute()
    for s in meta.get('sheets', []):
        if s['properties']['title'] == tab_name:
            g = s['properties'].get('gridProperties', {})
            return (s['properties']['sheetId'],
                    g.get('rowCount', 1000), g.get('columnCount', 26))
    return None, None, None


def cell(value, fmt_type=None, fmt_pattern=None, bold=False):
    """Build an updateCells value dict."""
    c = {}
    if isinstance(value, str):
        c['userEnteredValue'] = {'stringValue': value}
    else:
        c['userEnteredValue'] = {'numberValue': value}
    fmt = {}
    if fmt_type:
        fmt['numberFormat'] = {'type': fmt_type, 'pattern': fmt_pattern}
    if bold:
        fmt['textFormat'] = {'bold': True}
    if fmt:
        c['userEnteredFormat'] = fmt
    return c


def plan_rows(segments, years):
    """Build planned cell writes for the whole tab.

    year_cols maps each year to its 0-indexed column (D onward). Returns
    (writes, next_free_row, year_cols). Each write:
    (row_1idx, col_0idx, value, fmt_type, fmt_pattern, bold).
    """
    year_cols = {y: YEAR_START_COL + i for i, y in enumerate(years)}
    writes = [(1, 0, TAB_TITLE, None, None, True),
              (1, YEAR_START_COL - 1, UNIT_HEADER, None, None, True)]
    for y, c in year_cols.items():
        writes.append((1, c, y, None, None, True))

    # Top level: metric. Then classification, then item (item row carries values).
    r = 2
    for label, _col, ftype, fpat in METRICS:
        writes.append((r, 0, label, None, None, True))   # metric -> col A
        r += 1
        for cls in CLASSIFICATION_ORDER:
            if cls not in segments:
                continue
            block = segments[cls]
            # Drop items with no data inside the mirrored year columns.
            items = [it for it in block['items']
                     if any(y in block['data'][it] for y in year_cols)]
            if not items:
                continue
            writes.append((r, 1, cls, None, None, True))   # classification -> col B
            r += 1
            for item in items:
                writes.append((r, 2, item, None, None, False))  # item -> col C
                ydata = block['data'][item]
                for y, c in year_cols.items():
                    if y in ydata and label in ydata[y]:
                        val = ydata[y][label]
                        if ftype == 'NUMBER':
                            val = round(val, 2)
                        writes.append((r, c, val, ftype, fpat, False))
                r += 1
    return writes, r, year_cols


def apply_tab(service, spreadsheet_id, tab_name, segments, years, dry_run):
    writes, next_row, year_cols = plan_rows(segments, years)
    needed_cols = YEAR_START_COL + len(years)

    if dry_run:
        print(f"  Tab '{tab_name}': {next_row - 1} rows, "
              f"{needed_cols} cols, {len(writes)} cell writes")
        for (rr, cc, val, ftype, _fp, bold) in writes:
            if cc < YEAR_START_COL:
                indent = '   ' * cc
                tag = ' [bold]' if bold else ''
                print(f"    R{rr:<4} {col_to_letter(cc)}: {indent}{val}{tag}")
            else:
                year = next((y for y, c in year_cols.items() if c == cc), '?')
                shown = f"{val:,.1f}" if isinstance(val, float) else val
                print(f"           └ {year}: {shown}")
        return

    sheet_id, row_count, col_count = find_tab(service, spreadsheet_id, tab_name)
    requests = []

    if sheet_id is None:
        requests.append({'addSheet': {'properties': {
            'title': tab_name,
            'gridProperties': {'rowCount': max(next_row + 5, 100),
                               'columnCount': max(needed_cols, 26)}}}})
    else:
        # Grow grid if needed, then wipe values+formats for a clean rebuild.
        if next_row > row_count or needed_cols > col_count:
            requests.append({'updateSheetProperties': {
                'properties': {'sheetId': sheet_id, 'gridProperties': {
                    'rowCount': max(row_count, next_row + 5),
                    'columnCount': max(col_count, needed_cols)}},
                'fields': 'gridProperties.rowCount,gridProperties.columnCount'}})
        requests.append({'repeatCell': {
            'range': {'sheetId': sheet_id},
            'cell': {}, 'fields': 'userEnteredValue,userEnteredFormat'}})

    if requests:
        resp = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={'requests': requests}).execute()
        if sheet_id is None:
            sheet_id = resp['replies'][0]['addSheet']['properties']['sheetId']
            print(f"  ✓ Created tab '{tab_name}'")

    # Write content.
    content = []
    for (rr, cc, val, ftype, fpat, bold) in writes:
        content.append({'updateCells': {
            'range': {'sheetId': sheet_id,
                      'startRowIndex': rr - 1, 'endRowIndex': rr,
                      'startColumnIndex': cc, 'endColumnIndex': cc + 1},
            'rows': [{'values': [cell(val, ftype, fpat, bold)]}],
            'fields': 'userEnteredValue,userEnteredFormat'}})
    # Set label columns A-D to a uniform width (matches the 财务 tab's column C).
    content.append({'updateDimensionProperties': {
        'range': {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                  'startIndex': 0, 'endIndex': YEAR_START_COL},
        'properties': {'pixelSize': LABEL_COL_WIDTH},
        'fields': 'pixelSize'}})
    # Freeze the header row and the label columns (A-D).
    content.append({'updateSheetProperties': {
        'properties': {'sheetId': sheet_id, 'gridProperties': {
            'frozenRowCount': 1, 'frozenColumnCount': YEAR_START_COL}},
        'fields': 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount'}})
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={'requests': content}).execute()
    print(f"  ✓ Wrote '{tab_name}': {next_row - 1} rows, {len(writes)} cells")


def process_code(service, spreadsheet_id, code, fin_sheet_name, csv_rows, dry_run):
    print(f"\n{'='*60}\n{code}\n{'='*60}")
    segments = build_segments(csv_rows, code)
    if not segments:
        print("  No segment data returned; skipping.")
        return

    fin_name, fin_id, _rc, fin_cols = resolve_sheet(
        service, spreadsheet_id, fin_sheet_name)
    if fin_id is None:
        print(f"  ERROR: 财务 tab '{fin_sheet_name}' not found; skipping.")
        return

    years = get_fiscal_years(service, spreadsheet_id, fin_name, fin_cols)
    if not years:
        print("  ERROR: no year columns found on 财务 tab; skipping.")
        return

    base = fin_name[:-2] if fin_name.endswith('财务') else fin_name
    tab_name = base + TAB_SUFFIX
    print(f"  Mirroring years {years[0]}-{years[-1]} ({len(years)} cols) "
          f"from '{fin_name}' -> '{tab_name}'")
    for cls in CLASSIFICATION_ORDER:
        if cls in segments:
            shown = [it for it in segments[cls]['items']
                     if any(y in segments[cls]['data'][it] for y in years)]
            print(f"    {cls}: {len(shown)} items ({', '.join(shown)})")

    apply_tab(service, spreadsheet_id, tab_name, segments, years, dry_run)


def main():
    parser = argparse.ArgumentParser(description='Build 主营构成 运营数据 tab')
    parser.add_argument('--sheet-id', required=True)
    parser.add_argument('--codes', required=True,
                        help='Comma-separated A-share codes')
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()

    codes = [c.strip() for c in args.codes.split(',') if c.strip()]
    ashare = [c for c in codes if is_ashare(c)]
    skipped = [c for c in codes if not is_ashare(c)]
    if skipped:
        print(f"Skipping non-A-share codes (mainop.js unsupported): {skipped}")
    if not ashare:
        print("No A-share codes to process.")
        return

    service = get_service()
    mapping = get_summary_mapping(service, args.sheet_id)

    print(f"Fetching 主营构成 for: {', '.join(ashare)}")
    csv_rows = run_mainop(','.join(ashare))
    print(f"  Got {len(csv_rows)} CSV rows")

    for code in ashare:
        if code not in mapping:
            print(f"\nSKIP {code}: not found in Summary")
            continue
        process_code(service, args.sheet_id, code, mapping[code], csv_rows,
                     args.dry_run)


if __name__ == '__main__':
    main()
