#!/usr/bin/env python3
"""One-off: audit whether each industry spreadsheet can safely receive the new
ROIC rows (inserted right after the base ROIC row).

Insertion at roic_row+1 shifts every company-tab row below it down by 2. The
Summary tab references company-tab rows by fixed row number via INDIRECT, so the
insert is safe for a spreadsheet only if Summary never references a row strictly
greater than the ROIC row (i.e. max_summary_ref <= min ROIC row across tabs).

Prints one line per spreadsheet plus per-tab ROIC rows; rate-limit tolerant.
"""
import json
import re
import time
import sys
import os

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))
from create_company_tab import get_service

SEC = {'income statement', 'balance sheet', 'cash flow', 'key stats',
       'supplemental', 'business segments'}
BASE_ROIC = {'roic', 'roic (资本来源法)'}


def _retry(fn, tries=6):
    for a in range(tries):
        try:
            return fn()
        except Exception as e:
            if '429' in str(e) and a < tries - 1:
                time.sleep(30 * (a + 1))
            else:
                raise


def max_summary_ref(svc, sid):
    r = _retry(lambda: svc.spreadsheets().get(
        spreadsheetId=sid, ranges=["'Summary'"], includeGridData=True,
        fields='sheets.data.rowData.values.userEnteredValue').execute())
    rd = r['sheets'][0]['data'][0].get('rowData', [])
    mx = 0
    for row in rd:
        for cell in row.get('values', []):
            fv = cell.get('userEnteredValue', {}).get('formulaValue', '')
            for seg in re.finditer(r"财务'!([^)\"&]+)", fv):
                for m in re.finditer(r'\$?[A-Z]{1,3}\$?(\d+)', seg.group(1)):
                    mx = max(mx, int(m.group(1)))
    return mx


def roic_rows(svc, sid):
    meta = _retry(lambda: svc.spreadsheets().get(
        spreadsheetId=sid, fields='sheets.properties.title').execute())
    tabs = [s['properties']['title'] for s in meta['sheets']
            if s['properties']['title'] != 'Summary'
            and '资本结构' not in s['properties']['title']]
    found = {}
    missing = []
    for t in tabs:
        time.sleep(1.2)
        vals = _retry(lambda: svc.spreadsheets().values().get(
            spreadsheetId=sid, range=f"'{t}'!A1:C60").execute()).get('values', [])
        sec = None
        has_keystats = False
        rr = None
        for i, row in enumerate(vals):
            a = row[0].strip().lower() if row and row[0] else ''
            b = row[1].strip() if len(row) > 1 and row[1] else ''
            c = row[2].strip() if len(row) > 2 and row[2] else ''
            if a in SEC:
                sec = a
                if a == 'key stats':
                    has_keystats = True
                continue
            val = c if c else b
            if sec == 'key stats' and val.lower() in BASE_ROIC:
                rr = i + 1
        if rr:
            found[t] = rr
        elif has_keystats:
            missing.append(t)  # has Key Stats but no ROIC (like 农夫山泉)
        # tabs without a Key Stats section are non-financial → ignored
    return found, missing


def main():
    svc = get_service()
    inds = json.load(open(os.path.join(os.path.dirname(__file__), 'industry_spreadsheets.json')))
    print(f"{'industry':10} {'maxSumRef':>9} {'roicRows':>14} {'safe':>6}  missingROIC / notes")
    for name, info in inds.items():
        sid = info['spreadsheet_id']
        try:
            mx = max_summary_ref(svc, sid)
            found, missing = roic_rows(svc, sid)
        except Exception as e:
            print(f"{name:10} ERROR {str(e)[:60]}")
            continue
        rows = sorted(set(found.values()))
        safe = bool(rows) and mx <= min(rows)
        note = ''
        if missing:
            note = f"missingROIC={missing}"
        print(f"{name:10} {mx:>9} {str(rows):>14} {str(safe):>6}  {note}")
        time.sleep(2)


if __name__ == '__main__':
    main()
