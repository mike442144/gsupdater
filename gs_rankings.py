#!/usr/bin/env python3
"""
Fetch all ranking data from Google Sheets in ONE pass, output two CSVs:
  - rankings.csv         (EV/EBIT + ROIC ranking)
  - payout_rankings.csv  (5-year aggregate Payout Ratio ranking)

Each company tab is read exactly once; EBITDA/EBIT/ROIC and DPS/EPS are
extracted from the same response.

Usage:
    python gs_rankings.py                       # all rollout industries
    python gs_rankings.py 互联网 食品            # specific industries
    python gs_rankings.py --rankings r.csv --payout p.csv  # custom output paths

Then run combined_ranking.py on the two CSVs to get the master ranking.
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
        '药', '建材', '家电', '汽车', '设备', '交通运输', 'SAAS', '饮料']

SECTION_HEADERS = {'income statement', 'balance sheet', 'cash flow',
                   'key stats', 'supplemental', 'business segments'}
SUB_SECTIONS = {'盈利指标', '同比增速'}


# ── Helpers ─────────────────────────────────────────────────────────────────

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


def safe_float(val):
    """Convert cell value to float. Handles '45.2%' → 0.452, '-' → None."""
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
    if row_idx >= len(rows):
        return None
    row = rows[row_idx]
    if col_idx >= len(row):
        return None
    return row[col_idx]


# ── Summary ─────────────────────────────────────────────────────────────────

def read_summary(service, spreadsheet_id):
    """Returns (codes, names, tev_ebitda_row, summary_rows)."""
    result = _retry(lambda: service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="'Summary'!A1:BZ100"
    ).execute())
    rows = result.get('values', [])
    if len(rows) < 3:
        return [], [], None, rows

    codes, names = rows[0], rows[1]
    tev_ebitda_row = None
    for i in range(2, len(rows)):
        for col_idx in range(min(3, len(rows[i]))):
            val = str(rows[i][col_idx]).strip().lower() if rows[i][col_idx] else ''
            if val in ('tev/ebitda', 'ev/ebitda', 'tev / ebitda'):
                tev_ebitda_row = i
                break
        if tev_ebitda_row is not None:
            break
    return codes, names, tev_ebitda_row, rows


def build_company_col_map(codes, names):
    col_map = {}
    for j in range(max(len(codes), len(names))):
        code = str(codes[j]).strip() if j < len(codes) and codes[j] else ''
        name = str(names[j]).strip() if j < len(names) and names[j] else ''
        if code:
            col_map[code] = j
        if name:
            col_map[name] = j
    return col_map


# ── Company tab scan (single read) ──────────────────────────────────────────

def scan_company_tab(rows):
    """Parse one company tab's rows.

    Returns (item_map, all_data_cols, year_only_cols, header_labels).
      item_map:        {name_lower: row_index}
      all_data_cols:   cols with 4-digit year OR LTM header
      year_only_cols:  cols with pure 4-digit year header (no LTM)
      header_labels:   {col_index: header_str}
    """
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
        if not item_val or item_val in SUB_SECTIONS:
            continue
        item_map[item_val.lower()] = i

    all_data_cols = []
    year_only_cols = []
    header_labels = {}
    if rows:
        for j in range(3, len(rows[0])):
            text = str(rows[0][j]).strip() if j < len(rows[0]) and rows[0][j] else ''
            is_year = bool(re.fullmatch(r'\d{4}', text))
            is_ltm = text.upper().startswith('LTM')
            if text and (is_year or is_ltm):
                all_data_cols.append(j)
                header_labels[j] = text
            if is_year:
                year_only_cols.append(j)

    return item_map, all_data_cols, year_only_cols, header_labels


def pick_valuation_col(rows, data_cols):
    """Pick LTM column if present, else latest annual. For EV/EBIT & ROIC."""
    if not data_cols or not rows:
        return None, None
    header = rows[0]
    for dc in data_cols:
        if dc < len(header) and str(header[dc]).strip().upper().startswith('LTM'):
            return dc, 'LTM'
    for dc in reversed(data_cols):
        if dc < len(header):
            text = str(header[dc]).strip()
            if re.fullmatch(r'\d{4}', text):
                return dc, text
    return data_cols[-1], str(header[data_cols[-1]]).strip()


# ── Per-spreadsheet processing ──────────────────────────────────────────────

def process_spreadsheet(service, spreadsheet_id, industry,
                        ev_results, pay_results):
    codes, names, tev_ebitda_row, summary_rows = read_summary(
        service, spreadsheet_id)
    if not codes:
        print(f"  {industry}: no companies in Summary")
        return

    col_map = build_company_col_map(codes, names)
    tev_found = tev_ebitda_row is not None
    if not tev_found:
        print(f"  {industry}: WARNING — no TEV/EBITDA row in Summary")

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
            # ── Single API call per tab ────────────────────────────
            raw = _retry(lambda: service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{tab}'!A1:CV300"
            ).execute())
            rows = raw.get('values', [])

            item_map, all_data_cols, year_only_cols, header_labels = \
                scan_company_tab(rows)

            if not all_data_cols:
                print(f"    {company}: skip — no data columns")
                continue

            code = next((c for c, n in zip(codes, names)
                         if n.strip() == company), '')

            # ── EV/EBIT + ROIC (LTM or latest annual) ─────────────
            val_col, period = pick_valuation_col(rows, all_data_cols)
            if val_col is not None:
                ebitda_row = item_map.get('ebitda')
                ebit_row = item_map.get('ebit')
                roic_row = item_map.get('roic (资本来源法)') or item_map.get('roic')

                ebitda_val = safe_float(read_cell(rows, ebitda_row, val_col)) if ebitda_row is not None else None
                ebit_val = safe_float(read_cell(rows, ebit_row, val_col)) if ebit_row is not None else None
                roic_val = safe_float(read_cell(rows, roic_row, val_col)) if roic_row is not None else None

                summary_col = col_map.get(company)
                tev_ebitda_val = None
                if tev_found and summary_col is not None:
                    row_data = summary_rows[tev_ebitda_row]
                    if summary_col < len(row_data):
                        tev_ebitda_val = safe_float(row_data[summary_col])

                ev = None
                ev_ebit = None
                if tev_ebitda_val and ebitda_val:
                    ev = tev_ebitda_val * ebitda_val
                    if ebit_val and ebit_val != 0:
                        ev_ebit = ev / ebit_val

                roic_corrected = roic_val
                if (roic_val is not None and roic_val < 0
                        and ebit_val is not None and ebit_val > 0):
                    roic_corrected = -roic_val

                # ── Profit Quality: latest annual 扣非/净利润 ─────────
                net_income_row = item_map.get('net income')
                kf_row = item_map.get('扣非净利润')
                profit_quality = None
                if net_income_row is not None and kf_row is not None and len(year_only_cols) >= 1:
                    latest_yr_col = max(year_only_cols)
                    net_income = safe_float(read_cell(rows, net_income_row, latest_yr_col))
                    kf_net_income = safe_float(read_cell(rows, kf_row, latest_yr_col))
                    if net_income and net_income != 0 and kf_net_income is not None:
                        profit_quality = kf_net_income / net_income
                        # Exclude extreme outliers: <0.5 or >1.5
                        if profit_quality < 0.5 or profit_quality > 1.5:
                            profit_quality = None

                # ── FCF Ratio: latest annual FCFF / Net Income to Company ─
                fcff_row = item_map.get('fcff')
                nic_row = item_map.get('net income to company')
                fcf_ratio = None
                if fcff_row is not None and nic_row is not None and len(year_only_cols) >= 1:
                    latest_yr_col = max(year_only_cols)
                    fcff_val = safe_float(read_cell(rows, fcff_row, latest_yr_col))
                    nic_val = safe_float(read_cell(rows, nic_row, latest_yr_col))
                    if fcff_val is not None and nic_val and nic_val != 0:
                        fcf_ratio = fcff_val / nic_val

                # ── Capex Ratio: |Capital Expenditure| / Cash from Ops. ─
                capex_row = item_map.get('capital expenditure')
                ocf_row = item_map.get('cash from ops.')
                capex_ratio = None
                if capex_row is not None and ocf_row is not None and len(year_only_cols) >= 1:
                    latest_yr_col = max(year_only_cols)
                    capex_val = safe_float(read_cell(rows, capex_row, latest_yr_col))
                    ocf_val = safe_float(read_cell(rows, ocf_row, latest_yr_col))
                    if capex_val is not None and ocf_val and ocf_val != 0:
                        capex_ratio = abs(capex_val) / ocf_val

                ev_results.append({
                    'industry': industry, 'code': code, 'company': company,
                    'period': period,
                    'tev_ebitda': tev_ebitda_val, 'ev': ev,
                    'ebitda': ebitda_val, 'ebit': ebit_val,
                    'ev_ebit': ev_ebit,
                    'roic': roic_val, 'roic_corrected': roic_corrected,
                    'profit_quality': profit_quality,
                    'fcf_ratio': fcf_ratio,
                    'capex_ratio': capex_ratio,
                })

                # status
                parts = []
                if ev_ebit:
                    parts.append(f'EV/EBIT={ev_ebit:.1f}')
                if roic_val:
                    parts.append(f'ROIC={roic_val:.1%}')
                if not parts:
                    parts.append('no EV/ROIC data')
                ev_status = ', '.join(parts)
            else:
                ev_status = 'no usable period'

            # ── Payout Ratio (latest 5 annual years) ──────────────
            dps_row = item_map.get('dividends per share')
            eps_row = item_map.get('basic eps')

            if dps_row is not None and eps_row is not None and len(year_only_cols) >= 5:
                latest5 = sorted(year_only_cols, reverse=True)[:5]
                latest5 = sorted(latest5)

                dps_vals, eps_vals, yr_labels = [], [], []
                for dc in latest5:
                    d = safe_float(read_cell(rows, dps_row, dc))
                    e = safe_float(read_cell(rows, eps_row, dc))
                    dps_vals.append(d if d is not None else 0.0)
                    eps_vals.append(e if e is not None else 0.0)
                    yr_labels.append(header_labels.get(dc, str(dc)))

                total_eps = sum(eps_vals)
                total_dps = sum(dps_vals)

                if total_eps > 0:
                    agg_payout = total_dps / total_eps
                    year_ratios = []
                    for d, e in zip(dps_vals, eps_vals):
                        if e and e > 0:
                            year_ratios.append(min(d / e, 10.0))
                        else:
                            year_ratios.append(None)

                    pay_results.append({
                        'industry': industry, 'company': company,
                        'years': ' → '.join(yr_labels),
                        'total_dps': total_dps, 'total_eps': total_eps,
                        'agg_payout': agg_payout,
                        'yr1': year_ratios[0], 'yr2': year_ratios[1],
                        'yr3': year_ratios[2], 'yr4': year_ratios[3],
                        'yr5': year_ratios[4],
                    })
                    pay_status = f'Payout={agg_payout:.1%}'
                else:
                    pay_status = f'ΣEPS≤0'
            else:
                missing = []
                if dps_row is None:
                    missing.append('no DPS')
                if eps_row is None:
                    missing.append('no EPS')
                if len(year_only_cols) < 5:
                    missing.append(f'only {len(year_only_cols)}yr')
                pay_status = ', '.join(missing)

            print(f"    {company} [{period}]: {ev_status} | {pay_status}")

        except Exception as e:
            print(f"    {company}: ERROR — {e}")

        time.sleep(3)


# ── Ranking & output ────────────────────────────────────────────────────────

def write_ev_csv(results, path):
    ev_ranked = [r for r in results
                 if r['ev_ebit'] is not None
                 and (r['ev_ebit'] > 0
                      or (r['ebit'] is not None and r['ebit'] > 0))]
    ev_ranked.sort(key=lambda x: x['ev_ebit'])

    roic_ranked = [r for r in results
                   if r['roic_corrected'] is not None
                   and r['roic_corrected'] > 0]
    roic_ranked.sort(key=lambda x: -x['roic_corrected'])

    quality_ranked = [r for r in results if r['profit_quality'] is not None]
    quality_ranked.sort(key=lambda x: -x['profit_quality'])

    fcf_ranked = [r for r in results if r['fcf_ratio'] is not None]
    fcf_ranked.sort(key=lambda x: -x['fcf_ratio'])

    capex_ranked = [r for r in results if r['capex_ratio'] is not None]
    capex_ranked.sort(key=lambda x: x['capex_ratio'])

    ev_rank = {}
    for i, r in enumerate(ev_ranked, 1):
        ev_rank[(r['company'], r['industry'])] = i
        r['ev_rank'] = i

    roic_rank = {}
    for i, r in enumerate(roic_ranked, 1):
        roic_rank[(r['company'], r['industry'])] = i
        r['roic_rank'] = i

    quality_rank = {}
    for i, r in enumerate(quality_ranked, 1):
        quality_rank[(r['company'], r['industry'])] = i
        r['quality_rank'] = i

    fcf_rank = {}
    for i, r in enumerate(fcf_ranked, 1):
        fcf_rank[(r['company'], r['industry'])] = i
        r['fcf_rank'] = i

    capex_rank = {}
    for i, r in enumerate(capex_ranked, 1):
        capex_rank[(r['company'], r['industry'])] = i
        r['capex_rank'] = i

    combined = []
    for r in results:
        key = (r['company'], r['industry'])
        if key in ev_rank and key in roic_rank:
            combined.append({
                **r,
                'ev_rank': ev_rank[key],
                'roic_rank': roic_rank[key],
                'quality_rank': quality_rank.get(key),
                'fcf_rank': fcf_rank.get(key),
                'capex_rank': capex_rank.get(key),
                'combined': ev_rank[key] + roic_rank[key],
            })
    combined.sort(key=lambda x: x['combined'])

    fieldnames = ['industry', 'code', 'company', 'period',
                  'tev_ebitda', 'ev', 'ebitda', 'ebit',
                  'ev_ebit', 'ev_rank', 'roic', 'roic_corrected',
                  'roic_rank', 'profit_quality', 'quality_rank',
                  'fcf_ratio', 'fcf_rank',
                  'capex_ratio', 'capex_rank', 'combined']

    def _round(r):
        out = dict(r)
        for k in ('tev_ebitda', 'ev_ebit'):
            if out.get(k) is not None:
                out[k] = round(out[k], 1)
        for k in ('ev', 'ebitda', 'ebit'):
            if out.get(k) is not None:
                out[k] = round(out[k])
        for k in ('roic', 'roic_corrected', 'profit_quality', 'fcf_ratio', 'capex_ratio'):
            if out.get(k) is not None:
                out[k] = round(out[k], 4)
        return out

    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames,
                                extrasaction='ignore')
        writer.writeheader()
        for r in combined:
            writer.writerow(_round(r))

    return ev_ranked, roic_ranked, quality_ranked, fcf_ranked, capex_ranked, combined


def write_payout_csv(results, path):
    ranked = [r for r in results if r['agg_payout'] is not None]
    ranked.sort(key=lambda x: -x['agg_payout'])
    for i, r in enumerate(ranked, 1):
        r['rank'] = i

    fieldnames = ['rank', 'industry', 'company', 'agg_payout',
                  'total_dps', 'total_eps', 'years',
                  'yr1', 'yr2', 'yr3', 'yr4', 'yr5']

    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames,
                                extrasaction='ignore')
        writer.writeheader()
        for r in ranked:
            row = dict(r)
            for k in ('agg_payout', 'yr1', 'yr2', 'yr3', 'yr4', 'yr5'):
                if row.get(k) is not None:
                    row[k] = round(row[k], 4)
            for k in ('total_dps', 'total_eps'):
                if row.get(k) is not None:
                    row[k] = round(row[k], 2)
            writer.writerow(row)

    return ranked


def print_ev_ranking(ev_ranked):
    print(f"\n{'=' * 90}")
    print(f" EV/EBIT Ranking (low → high, cheaper = better)")
    print(f"{'=' * 90}")
    print(f"{'#':>3}  {'Company':<16} {'Industry':<8} "
          f"{'EV/EBIT':>8} {'ROIC':>9} {'Qual%':>7}  Code")
    print(f"{'-' * 60}")
    for r in ev_ranked:
        roic_s = '-'
        if r['roic_corrected'] is not None:
            flipped = r['roic'] is not None and r['roic_corrected'] != r['roic']
            roic_s = f"{r['roic_corrected']:.1%}" + ('*' if flipped else '')
        qual_s = '-'
        if r['profit_quality'] is not None:
            qual_s = f"{r['profit_quality']:.1%}"
        print(f"{r['ev_rank']:>3}  {r['company']:<16} {r['industry']:<8} "
              f"{r['ev_ebit']:>8.1f} {roic_s:>9} {qual_s:>7}  {r['code']}")


def print_roic_ranking(roic_ranked):
    print(f"\n{'=' * 90}")
    print(f" ROIC Ranking (high → low, * = sign-corrected)")
    print(f"{'=' * 90}")
    print(f"{'#':>3}  {'Company':<16} {'Industry':<8} "
          f"{'ROIC':>9} {'EV/EBIT':>8} {'Qual%':>7}  Code")
    print(f"{'-' * 60}")
    for r in roic_ranked:
        flipped = r['roic'] is not None and r['roic_corrected'] != r['roic']
        roic_s = f"{r['roic_corrected']:.1%}" + ('*' if flipped else '')
        ev_s = '-'
        if r['ev_ebit'] is not None:
            if r['ev_ebit'] > 0 or (r['ebit'] is not None and r['ebit'] > 0):
                ev_s = f"{r['ev_ebit']:.1f}"
        qual_s = '-'
        if r['profit_quality'] is not None:
            qual_s = f"{r['profit_quality']:.1%}"
        print(f"{r['roic_rank']:>3}  {r['company']:<16} {r['industry']:<8} "
              f"{roic_s:>9} {ev_s:>8} {qual_s:>7}  {r['code']}")


def print_payout_ranking(ranked):
    print(f"\n{'=' * 110}")
    print(f" 5-Year Aggregate Payout Ratio (ΣDPS/ΣEPS, high → low)")
    print(f"{'=' * 110}")
    print(f"{'#':>3}  {'Company':<20} {'Industry':<10} "
          f"{'Agg%':>8} {'ΣDPS':>8} {'ΣEPS':>8}  "
          f"{'Y1':>8} {'Y2':>8} {'Y3':>8} {'Y4':>8} {'Y5':>8}  Years")
    print(f"{'-' * 110}")
    for r in ranked:
        def p(v):
            if v is None:
                return '-'
            return f"{v:.0%}" if v >= 1.0 else f"{v:.1%}"
        print(f"{r['rank']:>3}  {r['company']:<20} {r['industry']:<10} "
              f"{p(r['agg_payout']):>8} "
              f"{r['total_dps']:>8.2f} {r['total_eps']:>8.2f}  "
              f"{p(r['yr1']):>8} {p(r['yr2']):>8} {p(r['yr3']):>8} "
              f"{p(r['yr4']):>8} {p(r['yr5']):>8}  {r['years']}")


def print_quality_ranking(quality_ranked):
    print(f"\n{'=' * 90}")
    print(f" Profit Quality Ranking (扣非净利润/净利润, high → low)")
    print(f"{'=' * 90}")
    print(f"{'#':>3}  {'Company':<16} {'Industry':<8} "
          f"{'Qual%':>7} {'ROIC':>9} {'EV/EBIT':>8}  Code")
    print(f"{'-' * 60}")
    for r in quality_ranked:
        qual_s = f"{r['profit_quality']:.1%}"
        roic_s = '-'
        if r['roic_corrected'] is not None:
            flipped = r['roic'] is not None and r['roic_corrected'] != r['roic']
            roic_s = f"{r['roic_corrected']:.1%}" + ('*' if flipped else '')
        ev_s = '-'
        if r['ev_ebit'] is not None:
            if r['ev_ebit'] > 0 or (r['ebit'] is not None and r['ebit'] > 0):
                ev_s = f"{r['ev_ebit']:.1f}"
        print(f"{r['quality_rank']:>3}  {r['company']:<16} {r['industry']:<8} "
              f"{qual_s:>7} {roic_s:>9} {ev_s:>8}  {r['code']}")


def print_fcf_ratio_ranking(fcf_ranked):
    print(f"\n{'=' * 90}")
    print(f" FCF Ratio Ranking (自由现金流/公司净利润, high → low)")
    print(f"{'=' * 90}")
    print(f"{'#':>3}  {'Company':<16} {'Industry':<8} "
          f"{'FCF%':>7} {'ROIC':>9} {'EV/EBIT':>8}  Code")
    print(f"{'-' * 60}")
    for r in fcf_ranked:
        fcf_s = f"{r['fcf_ratio']:.1%}"
        roic_s = '-'
        if r['roic_corrected'] is not None:
            flipped = r['roic'] is not None and r['roic_corrected'] != r['roic']
            roic_s = f"{r['roic_corrected']:.1%}" + ('*' if flipped else '')
        ev_s = '-'
        if r['ev_ebit'] is not None:
            if r['ev_ebit'] > 0 or (r['ebit'] is not None and r['ebit'] > 0):
                ev_s = f"{r['ev_ebit']:.1f}"
        print(f"{r['fcf_rank']:>3}  {r['company']:<16} {r['industry']:<8} "
              f"{fcf_s:>7} {roic_s:>9} {ev_s:>8}  {r['code']}")


def print_capex_ratio_ranking(capex_ranked):
    print(f"\n{'=' * 90}")
    print(f" Capex Ratio Ranking (资本开支/经营活动现金流, low → high)")
    print(f"{'=' * 90}")
    print(f"{'#':>3}  {'Company':<16} {'Industry':<8} "
          f"{'Capex%':>7} {'ROIC':>9} {'EV/EBIT':>8}  Code")
    print(f"{'-' * 60}")
    for r in capex_ranked:
        capex_s = f"{r['capex_ratio']:.1%}"
        roic_s = '-'
        if r['roic_corrected'] is not None:
            flipped = r['roic'] is not None and r['roic_corrected'] != r['roic']
            roic_s = f"{r['roic_corrected']:.1%}" + ('*' if flipped else '')
        ev_s = '-'
        if r['ev_ebit'] is not None:
            if r['ev_ebit'] > 0 or (r['ebit'] is not None and r['ebit'] > 0):
                ev_s = f"{r['ev_ebit']:.1f}"
        print(f"{r['capex_rank']:>3}  {r['company']:<16} {r['industry']:<8} "
              f"{capex_s:>7} {roic_s:>9} {ev_s:>8}  {r['code']}")


# ── Main ────────────────────────────────────────────────────────────────────

def main():
    here = os.path.dirname(__file__)
    parser = argparse.ArgumentParser(
        description='Fetch all ranking data from GS in one pass')
    parser.add_argument('industries', nargs='*',
                        help='Industries to include (default: all rollout)')
    parser.add_argument('--rankings', default=os.path.join(here, 'rankings.csv'),
                        help='Output: EV/EBIT + ROIC + Profit Quality CSV')
    parser.add_argument('--payout', default=os.path.join(here, 'payout_rankings.csv'),
                        help='Output: Payout Ratio CSV')
    args = parser.parse_args()

    industries = args.industries or SAFE
    all_industries = load_industries()

    service = get_service()
    ev_results = []
    pay_results = []

    for ind in industries:
        if ind not in all_industries:
            print(f"  WARNING: unknown industry '{ind}', skipping")
            continue
        sid = all_industries[ind]['spreadsheet_id']
        print(f"\n{'=' * 60}\n{ind} — {sid[:30]}...\n{'=' * 60}")
        process_spreadsheet(service, sid, ind, ev_results, pay_results)
        time.sleep(5)

    # ── Write CSVs ────────────────────────────────────────────────
    ev_ranked, roic_ranked, quality_ranked, fcf_ranked, capex_ranked, ev_combined = write_ev_csv(ev_results, args.rankings)
    pay_ranked = write_payout_csv(pay_results, args.payout)

    # ── Print summaries ──────────────────────────────────────────
    print_ev_ranking(ev_ranked)
    print_roic_ranking(roic_ranked)
    print_quality_ranking(quality_ranked)
    print_fcf_ratio_ranking(fcf_ranked)
    print_capex_ratio_ranking(capex_ranked)
    print_payout_ranking(pay_ranked)

    print(f"\n{'=' * 60}")
    print(f"Summary")
    print(f"{'=' * 60}")
    print(f"  Fetched: {len(ev_results)} companies for EV/ROIC/Quality/FCF/Capex, "
          f"{len(pay_results)} for Payout")
    print(f"  EV/EBIT ranked: {len(ev_ranked)}")
    print(f"  ROIC ranked:    {len(roic_ranked)}")
    print(f"  Profit Quality ranked: {len(quality_ranked)}")
    print(f"  FCF Ratio ranked: {len(fcf_ranked)}")
    print(f"  Capex Ratio ranked: {len(capex_ranked)}")
    print(f"  EV+ROIC combined: {len(ev_combined)}")
    print(f"  Payout ranked:  {len(pay_ranked)}")
    print(f"\n  → {args.rankings}")
    print(f"  → {args.payout}")
    print(f"\nNext step: python combined_ranking.py")


if __name__ == '__main__':
    main()
