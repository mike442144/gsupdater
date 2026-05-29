#!/usr/bin/env python3
"""
Rank all companies across rollout industries by EV/EBIT and ROIC.

Data sources:
  - TEV/EBITDA: from Summary sheet (current market multiple)
  - EV: derived as Summary TEV/EBITDA × company-tab EBITDA
  - EBIT, EBITDA, ROIC (资本来源法): from company tabs (LTM preferred, latest annual fallback)

Output: CSV sorted by EV/EBIT ascending (cheaper first).

Usage:
    python rank_companies.py                       # all rollout industries
    python rank_companies.py 互联网 食品            # specific industries
    python rank_companies.py --output rankings.csv  # custom output path
"""

import sys
import os
import re
import json
import csv
import time
import argparse

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')

SAFE = ['互联网', '传媒', '贸易', '纺织服装', '食品', '餐饮', '个人用品',
        '药', '建材', '家电', '汽车', '设备', '交通运输', 'SAAS']

SECTION_HEADERS = {'income statement', 'balance sheet', 'cash flow',
                   'key stats', 'supplemental', 'business segments'}


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


def col_to_letter(idx):
    result = ''
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ── Summary reading ────────────────────────────────────────────────────────

def read_summary(service, spreadsheet_id):
    """Read Summary sheet: codes (row 1), names (row 2), TEV/EBITDA row index.

    Returns (codes, names, tev_ebitda_row, summary_rows).
    tev_ebitda_row is 0-indexed or None if not found.
    """
    result = _retry(lambda: service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="'Summary'!A1:BZ100"
    ).execute())
    rows = result.get('values', [])

    if len(rows) < 3:
        return [], [], None, rows

    codes = rows[0]
    names = rows[1]

    tev_ebitda_row = None
    for i in range(2, len(rows)):
        row = rows[i]
        # Check columns A and B for the label
        for col_idx in range(min(3, len(row))):
            val = str(row[col_idx]).strip().lower() if row[col_idx] else ''
            if val in ('tev/ebitda', 'ev/ebitda', 'tev / ebitda'):
                tev_ebitda_row = i
                break
        if tev_ebitda_row is not None:
            break

    return codes, names, tev_ebitda_row, rows


def build_company_col_map(codes, names):
    """Map stock code and company name → Summary column index."""
    col_map = {}
    for j in range(max(len(codes), len(names))):
        code = str(codes[j]).strip() if j < len(codes) and codes[j] else ''
        name = str(names[j]).strip() if j < len(names) and names[j] else ''
        if code:
            col_map[code] = j
        if name:
            col_map[name] = j
    return col_map


def get_summary_val(summary_rows, row_idx, col_idx):
    """Safely read a cell from Summary data."""
    if row_idx is None or row_idx >= len(summary_rows):
        return None
    row = summary_rows[row_idx]
    if col_idx >= len(row):
        return None
    return row[col_idx]


# ── Company tab reading ───────────────────────────────────────────────────

def scan_company_tab(service, spreadsheet_id, sheet_name):
    """Read a company tab and return (all_rows, item_map, data_cols).

    item_map: {item_name_lower: (0_indexed_row, section)}
    data_cols: list of 0-indexed column indices with year/LTM headers
    """
    result = _retry(lambda: service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!A1:CV300"
    ).execute())
    rows = result.get('values', [])

    # Find item rows
    item_map = {}
    current_section = None
    for i, row in enumerate(rows):
        a = row[0].strip().lower() if row and row[0] else ''
        b = row[1].strip() if len(row) > 1 and row[1] else ''
        c = row[2].strip() if len(row) > 2 and row[2] else ''

        if a in SECTION_HEADERS:
            current_section = a
            continue

        item_val = c if c else b
        if not item_val:
            continue
        item_map[item_val.lower()] = (i, current_section)

    # Find data columns (headers with 4-digit year or LTM)
    data_cols = []
    if rows:
        for j in range(3, len(rows[0])):
            text = str(rows[0][j]).strip() if j < len(rows[0]) and rows[0][j] else ''
            if text and (re.search(r'\d{4}', text) or text.upper().startswith('LTM')):
                data_cols.append(j)

    return rows, item_map, data_cols


def pick_data_col(rows, data_cols):
    """Pick LTM column if present, else the last (rightmost) annual column.

    Returns (col_index, period_label) or (None, None).
    """
    if not data_cols or not rows:
        return None, None

    header = rows[0]
    for dc in data_cols:
        if dc < len(header) and str(header[dc]).strip().upper().startswith('LTM'):
            return dc, 'LTM'

    # Latest annual = rightmost col whose header is a pure 4-digit year
    for dc in reversed(data_cols):
        if dc < len(header):
            text = str(header[dc]).strip()
            if re.fullmatch(r'\d{4}', text):
                return dc, text

    return data_cols[-1], str(header[data_cols[-1]]).strip()


def safe_float(val):
    """Convert cell value to float. Returns None for non-numeric / empty.

    Handles percentage strings like '117.9%' → 1.179 and 'x' unit markers.
    """
    if val is None or val == '' or val == '-':
        return None
    try:
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).replace(',', '').strip()
        if not s or s == 'x':
            return None
        if s.endswith('%'):
            return float(s[:-1]) / 100.0
        return float(s)
    except (ValueError, TypeError):
        return None


def read_cell(rows, row_idx, col_idx):
    """Read a cell value, returning None for out-of-bounds."""
    if row_idx >= len(rows):
        return None
    row = rows[row_idx]
    if col_idx >= len(row):
        return None
    return row[col_idx]


# ── Main processing ───────────────────────────────────────────────────────

def process_spreadsheet(service, spreadsheet_id, industry, results, summary_cache):
    """Process one industry spreadsheet, appending to results."""
    codes, names, tev_ebitda_row, summary_rows = read_summary(
        service, spreadsheet_id)

    if not codes:
        print(f"  {industry}: no companies in Summary")
        return

    col_map = build_company_col_map(codes, names)
    tev_found = tev_ebitda_row is not None

    if not tev_found:
        print(f"  {industry}: WARNING — no TEV/EBITDA row in Summary")

    # Get company tabs
    meta = _retry(lambda: service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title))'
    ).execute())
    tabs = [s['properties']['title'] for s in meta.get('sheets', [])
            if s['properties']['title'] != 'Summary'
            and '资本结构' not in s['properties']['title']]

    print(f"  {industry}: {len(tabs)} tabs, "
          f"TEV/EBITDA row: {tev_ebitda_row + 1 if tev_found else 'N/A'}")

    for tab in tabs:
        company = tab.replace('财务', '')
        try:
            rows, item_map, data_cols = scan_company_tab(
                service, spreadsheet_id, tab)

            if not data_cols:
                print(f"    {company}: skip — no data columns")
                continue

            col, period = pick_data_col(rows, data_cols)
            if col is None:
                print(f"    {company}: skip — no usable period")
                continue

            # Locate items
            ebitda_info = item_map.get('ebitda')
            ebit_info = item_map.get('ebit')
            roic_info = item_map.get('roic (资本来源法)') or item_map.get('roic')

            # Read values at the chosen column
            ebitda_val = safe_float(read_cell(
                rows, ebitda_info[0], col)) if ebitda_info else None
            ebit_val = safe_float(read_cell(
                rows, ebit_info[0], col)) if ebit_info else None
            roic_val = safe_float(read_cell(
                rows, roic_info[0], col)) if roic_info else None

            # TEV/EBITDA from Summary
            summary_col = col_map.get(company)
            tev_ebitda_val = None
            if tev_found and summary_col is not None:
                tev_ebitda_val = safe_float(get_summary_val(
                    summary_rows, tev_ebitda_row, summary_col))

            # Compute EV and EV/EBIT
            ev = None
            ev_ebit = None
            if tev_ebitda_val and ebitda_val:
                ev = tev_ebitda_val * ebitda_val
                if ebit_val and ebit_val != 0:
                    ev_ebit = ev / ebit_val

            results.append({
                'industry': industry,
                'code': next((c for c, n in zip(codes, names)
                              if n.strip() == company), ''),
                'company': company,
                'period': period,
                'tev_ebitda': tev_ebitda_val,
                'ev': ev,
                'ebitda': ebitda_val,
                'ebit': ebit_val,
                'ev_ebit': ev_ebit,
                'roic': roic_val,
            })

            # Status
            parts = []
            if ev_ebit:
                parts.append(f'EV/EBIT={ev_ebit:.1f}')
            if roic_val:
                parts.append(f'ROIC={roic_val:.1%}')
            skip = []
            if not tev_ebitda_val:
                skip.append('no TEV/EBITDA from Summary')
            if not ebitda_val:
                skip.append('no EBITDA')
            if not ebit_val:
                skip.append('no EBIT')
            status = ', '.join(parts) if parts else ', '.join(skip)
            print(f"    {company} [{period}]: {status}")

        except Exception as e:
            print(f"    {company}: ERROR — {e}")

        time.sleep(3)


def main():
    parser = argparse.ArgumentParser(
        description='Rank companies by EV/EBIT and ROIC')
    parser.add_argument('industries', nargs='*',
                        help='Industries to include (default: all rollout)')
    parser.add_argument('--output', '-o', default='rankings.csv',
                        help='Output CSV path (default: rankings.csv)')
    args = parser.parse_args()

    industries = args.industries or SAFE
    all_industries = load_industries()

    service = get_service()
    results = []

    for ind in industries:
        if ind not in all_industries:
            print(f"  WARNING: unknown industry '{ind}', skipping")
            continue
        sid = all_industries[ind]['spreadsheet_id']
        print(f"\n{'=' * 60}\n{ind} — {sid[:30]}...\n{'=' * 60}")
        process_spreadsheet(service, sid, ind, results, {})
        time.sleep(5)  # pause between spreadsheets

    # ── Rankings ──────────────────────────────────────────────────────
    # EV/EBIT: ascending (cheaper = better), only positive values
    ev_ranked = [r for r in results
                 if r['ev_ebit'] is not None and r['ev_ebit'] > 0]
    ev_ranked.sort(key=lambda x: x['ev_ebit'])

    # ROIC: descending (higher return = better), only positive values
    roic_ranked = [r for r in results
                   if r['roic'] is not None and r['roic'] > 0]
    roic_ranked.sort(key=lambda x: -x['roic'])

    # Assign rank numbers — key by (company, industry) to handle duplicates
    ev_rank = {}
    for i, r in enumerate(ev_ranked, 1):
        ev_rank[(r['company'], r['industry'])] = i
        r['ev_rank'] = i

    roic_rank = {}
    for i, r in enumerate(roic_ranked, 1):
        roic_rank[(r['company'], r['industry'])] = i
        r['roic_rank'] = i

    # Combined: companies in BOTH lists, rank sum ascending
    combined = []
    for r in results:
        key = (r['company'], r['industry'])
        if key in ev_rank and key in roic_rank:
            combined.append({
                **r,
                'ev_rank': ev_rank[key],
                'roic_rank': roic_rank[key],
                'combined': ev_rank[key] + roic_rank[key],
            })
    combined.sort(key=lambda x: x['combined'])

    # ── Write CSV ──────────────────────────────────────────────────────
    fieldnames = ['industry', 'code', 'company', 'period',
                  'tev_ebitda', 'ev', 'ebitda', 'ebit',
                  'ev_ebit', 'ev_rank', 'roic', 'roic_rank', 'combined']

    def _round_row(r):
        """Round numeric fields for cleaner CSV output."""
        out = dict(r)
        for key in ('tev_ebitda', 'ev_ebit'):
            if out.get(key) is not None:
                out[key] = round(out[key], 1)
        for key in ('ev', 'ebitda', 'ebit'):
            if out.get(key) is not None:
                out[key] = round(out[key])
        if out.get('roic') is not None:
            out['roic'] = round(out['roic'], 4)
        return out

    output_path = os.path.join(os.path.dirname(__file__), args.output)
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames,
                                extrasaction='ignore')
        writer.writeheader()
        for r in combined:
            writer.writerow(_round_row(r))

    # ── Print: EV/EBIT ranking ────────────────────────────────────────
    print(f"\n{'=' * 90}")
    print(f" EV/EBIT Ranking (low → high, cheaper = better)")
    print(f"{'=' * 90}")
    print(f"{'#':>3}  {'Company':<16} {'Industry':<8} "
          f"{'EV/EBIT':>8} {'ROIC':>8}  Code")
    print(f"{'-' * 60}")
    for r in ev_ranked:
        roic_str = f"{r['roic']:.1%}" if r['roic'] and r['roic'] > 0 else '-'
        print(f"{r['ev_rank']:>3}  {r['company']:<16} {r['industry']:<8} "
              f"{r['ev_ebit']:>8.1f} {roic_str:>8}  {r['code']}")

    # ── Print: ROIC ranking ───────────────────────────────────────────
    print(f"\n{'=' * 90}")
    print(f" ROIC Ranking (high → low, better return)")
    print(f"{'=' * 90}")
    print(f"{'#':>3}  {'Company':<16} {'Industry':<8} "
          f"{'ROIC':>8} {'EV/EBIT':>8}  Code")
    print(f"{'-' * 60}")
    for r in roic_ranked:
        ev_str = f"{r['ev_ebit']:.1f}" if r['ev_ebit'] and r['ev_ebit'] > 0 else '-'
        print(f"{r['roic_rank']:>3}  {r['company']:<16} {r['industry']:<8} "
              f"{r['roic']:>7.1%} {ev_str:>8}  {r['code']}")

    # ── Print: Combined ranking ───────────────────────────────────────
    print(f"\n{'=' * 90}")
    print(f" Combined Ranking (EV/EBIT rank + ROIC rank, lower = better)")
    print(f"{'=' * 90}")
    print(f"{'#':>3}  {'Company':<16} {'Industry':<8} "
          f"{'Combined':>8} {'EV/EBIT':>8} {'R_EV':>5} {'ROIC':>8} {'R_ROIC':>6}  Code")
    print(f"{'-' * 80}")
    for i, r in enumerate(combined, 1):
        print(f"{i:>3}  {r['company']:<16} {r['industry']:<8} "
              f"{r['combined']:>8} {r['ev_ebit']:>8.1f} {r['ev_rank']:>5} "
              f"{r['roic']:>7.1%} {r['roic_rank']:>6}  {r['code']}")

    print(f"\nTotal: {len(results)} companies | "
          f"EV/EBIT ranked: {len(ev_ranked)} | "
          f"ROIC ranked: {len(roic_ranked)} | "
          f"Combined: {len(combined)}")
    print(f"CSV: {output_path}")


if __name__ == '__main__':
    main()
