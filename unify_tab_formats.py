#!/usr/bin/env python3
"""
Unify number formats across the company tabs ('<name>财务') of a spreadsheet.

- 扣非净利润 row: literal numeric cells -> blue text + #,##0; formula cells
  (LTM) -> #,##0 with the default color.
- Every other literal numeric cell in a data column (year / LTM / quarter
  headers) -> #,##0; EPS-like rows (label contains 'eps' or 'per share') ->
  #,##0.00.
- Formulas, text and cells with a deliberate non-NUMBER format (PERCENT,
  DATE, ...) are never touched; cells already matching the target are
  skipped, so re-runs write nothing.

Usage:
    python unify_tab_formats.py --spreadsheet-id <ID> [--dry-run]
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

BLUE = {'blue': 1}
THOUSANDS = {'type': 'NUMBER', 'pattern': '#,##0'}
TWO_DECIMALS = {'type': 'NUMBER', 'pattern': '#,##0.00'}


def col_to_letter(col_idx):
    result = ''
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_service():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    return build('sheets', 'v4', credentials=creds)


def company_tabs(service, spreadsheet_id):
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


def is_data_header(h):
    h = str(h).strip()
    return (re.match(r'^\d{4}$', h)
            or re.match(r'^Q[1-4] \d{4}$', h)
            or h.upper().startswith('LTM'))


def is_eps_label(label):
    return 'eps' in label or 'per share' in label


def fmt_matches(current, target):
    return (current.get('type') == target['type']
            and current.get('pattern') == target['pattern'])


def sweep_tab(service, spreadsheet_id, title, sheet_id, row_count, col_count, dry_run):
    header = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"'{title}'!A1:{col_to_letter(col_count - 1)}1"
    ).execute().get('values', [[]])[0]
    data_cols = [j for j, h in enumerate(header) if j >= 3 and is_data_header(h)]
    if not data_cols:
        print(f"  {title}: no data columns, skip")
        return

    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{title}'!A1:{col_to_letter(col_count - 1)}{row_count}").execute().get('values', [])
    last_row = max((i + 1 for i, r in enumerate(values) if any(str(c).strip() for c in r)), default=0)
    if last_row < 2:
        print(f"  {title}: empty tab, skip")
        return

    labels = {}
    kcfj_row = None
    for i, r in enumerate(values[:last_row]):
        b = r[1].strip() if len(r) > 1 and r[1] else ''
        c = r[2].strip() if len(r) > 2 and r[2] else ''
        label = (c or b)
        labels[i] = label.lower()
        if label == '扣非净利润' and kcfj_row is None:
            kcfj_row = i

    grid = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[f"'{title}'!A1:{col_to_letter(col_count - 1)}{last_row}"],
        includeGridData=True,
        fields='sheets.data.rowData.values(userEnteredValue,userEnteredFormat)'
    ).execute()
    row_data = (grid.get('sheets', [{}])[0].get('data', [{}])[0].get('rowData', []))

    requests = []
    counts = {'koufei': 0, 'eps': 0, 'number': 0}
    for i, row in enumerate(row_data):
        label = labels.get(i, '')
        eps_row = is_eps_label(label)
        cells = row.get('values', [])
        for j in data_cols:
            if j >= len(cells):
                break
            cell = cells[j]
            uev = cell.get('userEnteredValue', {})
            uef = cell.get('userEnteredFormat', {})
            cur_fmt = uef.get('numberFormat', {})

            if 'numberValue' in uev:
                if i == kcfj_row:
                    target = THOUSANDS
                    target_cell = {'numberFormat': dict(target),
                                   'textFormat': {'foregroundColor': dict(BLUE)}}
                    if fmt_matches(cur_fmt, target) and uef.get(
                            'textFormat', {}).get('foregroundColor', {}).get('blue') == 1:
                        continue
                    fields = 'userEnteredFormat(numberFormat,textFormat.foregroundColor)'
                    counts['koufei'] += 1
                else:
                    target = TWO_DECIMALS if eps_row else THOUSANDS
                    # Only normalize NUMBER-typed or unformatted cells — PERCENT/
                    # DATE etc. are deliberate.
                    if cur_fmt and cur_fmt.get('type') not in ('NUMBER', None):
                        continue
                    if fmt_matches(cur_fmt, target):
                        continue
                    target_cell = {'numberFormat': dict(target)}
                    fields = 'userEnteredFormat(numberFormat)'
                    counts['eps' if eps_row else 'number'] += 1
                requests.append({
                    'updateCells': {
                        'range': {'sheetId': sheet_id,
                                  'startRowIndex': i, 'endRowIndex': i + 1,
                                  'startColumnIndex': j, 'endColumnIndex': j + 1},
                        'rows': [{'values': [{'userEnteredFormat': target_cell}]}],
                        'fields': fields,
                    }
                })
            elif 'formulaValue' in uev and i == kcfj_row:
                # LTM formula cell: thousands, default color
                if fmt_matches(cur_fmt, THOUSANDS):
                    continue
                requests.append({
                    'updateCells': {
                        'range': {'sheetId': sheet_id,
                                  'startRowIndex': i, 'endRowIndex': i + 1,
                                  'startColumnIndex': j, 'endColumnIndex': j + 1},
                        'rows': [{'values': [{'userEnteredFormat': {'numberFormat': dict(THOUSANDS)}}]}],
                        'fields': 'userEnteredFormat(numberFormat)',
                    }
                })
                counts['koufei'] += 1

    verb = 'would reformat' if dry_run else 'reformatted'
    print(f"  {title}: {verb} 扣非 {counts['koufei']}, eps {counts['eps']}, number {counts['number']} cells")
    if dry_run or not requests:
        return
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={'requests': requests}).execute()


def main():
    parser = argparse.ArgumentParser(description='Unify number formats in company tabs')
    parser.add_argument('--spreadsheet-id', required=True)
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()

    service = get_service()
    for title, sheet_id, row_count, col_count in company_tabs(service, args.spreadsheet_id):
        sweep_tab(service, args.spreadsheet_id, title, sheet_id, row_count, col_count, args.dry_run)


if __name__ == '__main__':
    main()
