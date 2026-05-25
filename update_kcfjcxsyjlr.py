#!/usr/bin/env python3
"""
Update 扣非净利润 (non-recurring net profit) for A-share companies in Google Sheets.

Data source: eastmoney API via ~/Projects/tinyant/eastmoney/index.js

Tasks:
1. Fill annual 扣非净利润 in yearly columns (2024, 2025, etc.)
2. Fill quarterly 扣非净利润 in quarterly columns (Q1 2024, Q2 2024, etc.)
3. Add SUM formula for LTM column (sum of latest 4 quarters)
"""

import sys
import os
import re
import json
import subprocess
import csv
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# Configuration
GOOGLE_SHEET_ID = "1huXdbAgYR2xul5CDtOmuoCjBKGwQu69XB9_AcooRPC0"
GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')
EASTMONEY_SCRIPT = os.path.join(os.path.expanduser('~/Projects/tinyant/eastmoney'), 'index.js')
EASTMONEY_DIR = os.path.expanduser('~/Projects/tinyant/eastmoney')


def col_to_letter(col_idx):
    """Convert 0-indexed column number to Excel-style letter."""
    result = ''
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def parse_quarter_from_header(header_val):
    """Parse 'Q1 2024' -> ('Q1', '2024') or None."""
    if not header_val:
        return None
    match = re.match(r'Q(\d)\s+(\d{4})', str(header_val).strip())
    if match:
        return f'Q{match.group(1)}', match.group(2)
    return None


def parse_report_date(date_str):
    """Parse '2024-12-31 00:00:00' -> ('Q4', '2024')."""
    if not date_str:
        return None
    try:
        dt = datetime.strptime(date_str.split()[0], '%Y-%m-%d')
        quarter = f'Q{(dt.month - 1) // 3 + 1}'
        return quarter, str(dt.year)
    except:
        return None


def run_eastmoney(codes, period, count):
    """Run eastmoney script and return CSV data."""
    cmd = ['node', EASTMONEY_SCRIPT, '--codes', codes, '--period', period, '--count', str(count)]
    
    # Remove old CSV
    csv_pattern = Path(EASTMONEY_DIR) / 'data' / 'eastmoney_finance_*.csv'
    for f in Path(EASTMONEY_DIR).glob('data/eastmoney_finance_*.csv'):
        f.unlink()
    
    # Run script
    result = subprocess.run(cmd, cwd=EASTMONEY_DIR, capture_output=True, text=True, timeout=60)
    if result.returncode != 0:
        print(f"  ERROR: eastmoney script failed: {result.stderr}")
        return []
    
    # Read CSV
    csv_files = list(Path(EASTMONEY_DIR).glob('data/eastmoney_finance_*.csv'))
    if not csv_files:
        print(f"  ERROR: No CSV output found")
        return []
    
    data = []
    with open(csv_files[0], 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            data.append(row)
    
    return data


def extract_kcfjcxsyjlr_data(csv_data):
    """Extract annual KCFJCXSYJLR (扣非净利润) from yearly CSV."""
    result = {}
    for row in csv_data:
        if row.get('REPORT_TYPE') == '年报' and row.get('KCFJCXSYJLR'):
            year = row.get('REPORT_YEAR')
            value = row.get('KCFJCXSYJLR')
            if year and value:
                try:
                    result[year] = round(float(value), 2)
                except:
                    pass
    return result


def extract_dedu_parent_profit_data(csv_data):
    """Extract quarterly DEDU_PARENT_PROFIT (单季度扣非净利润) from quarterly CSV."""
    result = {}
    for row in csv_data:
        if not row.get('REPORT_TYPE'):  # Quarterly data has empty REPORT_TYPE
            report_date = row.get('REPORT_DATE')
            value = row.get('DEDU_PARENT_PROFIT')
            if report_date and value:
                parsed = parse_report_date(report_date)
                if parsed:
                    quarter, year = parsed
                    key = f'{quarter}_{year}'
                    try:
                        result[key] = round(float(value), 2)
                    except:
                        pass
    return result


def update_kcfjcxsyjlr(service, spreadsheet_id, sheet_name, stock_code, dry_run=False):
    """Update 扣非净利润 for a single company."""
    print(f"\n{'='*60}")
    print(f"Updating 扣非净利润 for: {sheet_name} (code: {stock_code})")
    print(f"{'='*60}")
    
    # Get sheet ID and grid width
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId,gridProperties(columnCount)))'
    ).execute()
    target_sheet_id = None
    grid_width = 0
    actual_sheet_name = sheet_name
    for s in spreadsheet.get('sheets', []):
        if s['properties']['title'] == sheet_name:
            target_sheet_id = s['properties']['sheetId']
            grid_width = s['properties']['gridProperties']['columnCount']
            break
    
    # Fallback: try without "财务" suffix
    if target_sheet_id is None and sheet_name.endswith('财务'):
        fallback_name = sheet_name[:-2]
        for s in spreadsheet.get('sheets', []):
            if s['properties']['title'] == fallback_name:
                target_sheet_id = s['properties']['sheetId']
                grid_width = s['properties']['gridProperties']['columnCount']
                actual_sheet_name = fallback_name
                print(f"  NOTE: Using sheet name '{fallback_name}' (without '财务' suffix)")
                break
    
    if target_sheet_id is None:
        print(f"  ERROR: Sheet not found")
        return False
    
    # Use actual_sheet_name for all subsequent API calls
    sheet_name = actual_sheet_name
    
    # Build range covering full grid width
    end_col = col_to_letter(grid_width - 1)
    
    # Dynamically find 扣非净利润 row by scanning column B
    b_col_result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!B1:B100"
    ).execute()
    b_rows = b_col_result.get('values', [])
    
    kcfj_row = None
    for i, row in enumerate(b_rows):
        if row and '扣非净利润' in str(row[0]).strip():
            kcfj_row = i + 1  # 1-indexed
            break
    
    if kcfj_row is None:
        print(f"  ERROR: 扣非净利润 not found in column B")
        return False
    
    print(f"  Found 扣非净利润 at row {kcfj_row}")
    
    range_str = f"'{sheet_name}'!A1:{end_col}1"
    range_str2 = f"'{sheet_name}'!A{kcfj_row}:{end_col}{kcfj_row}"
    
    # Read GS header and 扣非净利润 row
    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[range_str, range_str2],
        includeGridData=True
    ).execute()
    
    sheets_data = result.get('sheets', [])
    if not sheets_data:
        print(f"  ERROR: Sheet not found")
        return False
    
    # Parse header
    header_row = sheets_data[0].get('data', [{}])[0].get('rowData', [{}])[0].get('values', [])
    year_cols = {}
    quarter_cols = {}
    ltm_col = None
    
    for j, v in enumerate(header_row):
        fv = v.get('formattedValue', '').strip()
        # Also check stringValue for multi-line text (LTM\n12 months\n...)
        sv = v.get('userEnteredValue', {}).get('stringValue', '')
        combined = sv if sv else fv
        
        if re.match(r'^\d{4}$', fv):
            year_cols[fv] = j
        elif combined.upper().startswith('LTM'):
            ltm_col = j
        else:
            parsed = parse_quarter_from_header(fv)
            if parsed:
                quarter, year = parsed
                quarter_cols[f'{quarter}_{year}'] = j
    
    # Parse existing 扣非净利润 data
    kcfj_row_data = sheets_data[0].get('data', [{}])[1].get('rowData', [{}])[0].get('values', [])
    existing_data = {}
    for j, v in enumerate(kcfj_row_data):
        fv = v.get('formattedValue', '')
        if fv and fv.strip() and fv.strip() != '扣非净利润':
            try:
                existing_data[j] = float(fv.replace(',', ''))
            except:
                pass
    
    print(f"  GS structure: {len(year_cols)} year cols, {len(quarter_cols)} quarter cols, LTM col: {ltm_col}")
    print(f"  Existing 扣非净利润 data: {len(existing_data)} cells")
    
    # Run eastmoney for yearly data
    print(f"  Fetching yearly data from eastmoney...")
    yearly_csv = run_eastmoney(stock_code, 'y', 5)
    yearly_data = extract_kcfjcxsyjlr_data(yearly_csv)
    print(f"    Got {len(yearly_data)} years: {list(yearly_data.keys())}")
    
    # Run eastmoney for quarterly data (20 quarters)
    print(f"  Fetching quarterly data from eastmoney...")
    quarterly_csv = run_eastmoney(stock_code, 'q', 20)
    quarterly_data = extract_dedu_parent_profit_data(quarterly_csv)
    print(f"    Got {len(quarterly_data)} quarters: {list(quarterly_data.keys())}")
    
    # Determine updates needed
    updates = []
    
    # 1. Annual data
    for year, value in yearly_data.items():
        if year in year_cols:
            col_idx = year_cols[year]
            # Check if value differs significantly from existing (tolerance 0.5)
            if col_idx not in existing_data or abs(existing_data[col_idx] - value) > 0.5:
                updates.append((col_idx, value, f'{year} annual'))
    
    # 2. Quarterly data - fill into empty columns (no insertDimension)
    all_quarters_from_eastmoney = sorted(quarterly_data.keys(), key=lambda x: (x.split('_')[1], x.split('_')[0]))
    new_quarters = [q for q in all_quarters_from_eastmoney if q not in quarter_cols]
    
    if new_quarters:
        print(f"  Appending {len(new_quarters)} new quarter columns: {new_quarters[0]} to {new_quarters[-1]}")
        
        # Find the first and last quarter columns — search for empty within the quarter area
        first_quarter_col = min(quarter_cols.values()) if quarter_cols else None
        last_data_col = max(
            list(year_cols.values()) +
            list(quarter_cols.values()) +
            ([ltm_col] if ltm_col is not None else [0])
        )
        
        # Determine search start: first quarter column (to fill gaps), or after last data col
        search_start = first_quarter_col if first_quarter_col is not None else last_data_col + 1
        
        # Find empty columns from search_start onward
        empty_cols = []
        for j in range(search_start, len(header_row)):
            # Skip label columns (A, B, C)
            if j < 3:
                continue
            v = header_row[j]
            fv = v.get('formattedValue', '').strip()
            sv = v.get('userEnteredValue', {}).get('stringValue', '')
            if not fv and not sv:
                empty_cols.append(j)
        
        # Also include columns beyond current header within grid
        for j in range(len(header_row), grid_width):
            empty_cols.append(j)
        
        if len(empty_cols) < len(new_quarters):
            # Need to expand grid
            required_width = grid_width + len(new_quarters) - len(empty_cols) + 5
            if not dry_run:
                expand_request = {
                    'updateSheetProperties': {
                        'properties': {
                            'sheetId': target_sheet_id,
                            'gridProperties': {'columnCount': required_width}
                        },
                        'fields': 'gridProperties.columnCount'
                    }
                }
                service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={'requests': [expand_request]}
                ).execute()
                print(f"  ✓ Expanded grid to {required_width} columns")
                grid_width = required_width
                # Add new empty columns
                for j in range(len(header_row), required_width):
                    if j not in empty_cols:
                        empty_cols.append(j)
        
        # Assign new quarters to empty columns
        header_updates = []
        for i, qkey in enumerate(new_quarters):
            if i >= len(empty_cols):
                print(f"  WARNING: No empty column for {qkey}")
                break
            col_idx = empty_cols[i]
            quarter_cols[qkey] = col_idx
            updates.append((col_idx, quarterly_data[qkey], qkey))
            quarter, year = qkey.split('_')
            header_updates.append((col_idx, f'{quarter} {year}'))
        
        # Write headers
        if not dry_run and header_updates:
            requests = []
            for col_idx, label in header_updates:
                requests.append({
                    'updateCells': {
                        'range': {
                            'sheetId': target_sheet_id,
                            'startRowIndex': 0, 'endRowIndex': 1,
                            'startColumnIndex': col_idx,
                            'endColumnIndex': col_idx + 1,
                        },
                        'rows': [{'values': [{'userEnteredValue': {'stringValue': label}}]}],
                        'fields': 'userEnteredValue',
                    }
                })
            if requests:
                service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={'requests': requests}
                ).execute()
                print(f"  ✓ Wrote {len(requests)} quarter headers")

    # 2b. Backfill empty existing quarter columns with eastmoney data
    backfilled = 0
    for qkey, col_idx in quarter_cols.items():
        if qkey in quarterly_data and col_idx not in existing_data:
            updates.append((col_idx, quarterly_data[qkey], qkey))
            backfilled += 1
    if backfilled:
        print(f"  Backfilled {backfilled} empty quarter columns")

    # 3. LTM formula (sum of latest 4 quarters) - calculate AFTER new columns added
    if ltm_col is not None:
        # Find latest 4 quarters from updated quarter_cols
        sorted_quarters = sorted(quarter_cols.keys(), key=lambda x: (x.split('_')[1], x.split('_')[0]), reverse=True)
        latest_4 = sorted_quarters[:4]
        if len(latest_4) == 4:
            cols = [quarter_cols[q] for q in latest_4]
            col_refs = [f'{col_to_letter(c)}{kcfj_row}' for c in cols]
            formula = f'=SUM({",".join(col_refs)})'
            updates.append((ltm_col, formula, 'LTM formula'))
            print(f"  LTM formula: {formula} (quarters: {latest_4})")
    
    # Apply updates
    if updates:
        print(f"  Writing {len(updates)} updates...")
        
        if dry_run:
            for col_idx, value, label in updates:
                print(f"    [DRY RUN] Col {col_idx} ({label}): {value}")
        else:
            # Use batchUpdate with updateCells to set both value and number format
            # Format "0" displays as integer but stores the full precision value
            requests = []
            for col_idx, value, label in updates:
                cell_value = {}
                if isinstance(value, str) and value.startswith('='):
                    cell_value['userEnteredValue'] = {'formulaValue': value}
                else:
                    cell_value['userEnteredValue'] = {'numberValue': value}
                    cell_value['userEnteredFormat'] = {'numberFormat': {'type': 'NUMBER', 'pattern': '0'}}
                
                requests.append({
                    'updateCells': {
                        'range': {
                            'sheetId': target_sheet_id,
                            'startRowIndex': kcfj_row - 1,
                            'endRowIndex': kcfj_row,
                            'startColumnIndex': col_idx,
                            'endColumnIndex': col_idx + 1,
                        },
                        'rows': [{'values': [cell_value]}],
                        'fields': 'userEnteredValue,userEnteredFormat',
                    }
                })
            
            if requests:
                service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={'requests': requests}
                ).execute()
                print(f"  ✓ Wrote {len(requests)} cells (format: integer display)")
    else:
        print(f"  No updates needed")
    
    return True


def is_ashare(code):
    """Check if stock code is A-share (6-digit number)."""
    return bool(re.match(r'^\d{6}$', code))


def get_summary_mapping(service, spreadsheet_id):
    """Get stock code -> sheet name mapping from Summary sheet."""
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="'Summary'!A1:AZ3"
    ).execute()
    rows = result.get('values', [])
    
    if len(rows) < 2:
        return {}
    
    codes_row = rows[0]
    names_row = rows[1]
    
    mapping = {}
    for j in range(len(codes_row)):
        code = str(codes_row[j]).strip() if j < len(codes_row) else ''
        name = str(names_row[j]).strip() if j < len(names_row) else ''
        if code and name:
            sheet_name = f"{name}财务"
            mapping[code] = sheet_name
    
    return mapping


def main():
    import argparse
    parser = argparse.ArgumentParser(description='Update 扣非净利润 from eastmoney')
    parser.add_argument('--dry-run', action='store_true', help='Preview changes without writing')
    parser.add_argument('--sheet-id', default=GOOGLE_SHEET_ID, help='Google Sheets spreadsheet ID')
    parser.add_argument('--codes', help='Comma-separated stock codes (default: all A-shares from Summary)')
    args = parser.parse_args()
    
    # Initialize Google Sheets API
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    service = build('sheets', 'v4', credentials=creds)
    
    # Get mapping from Summary sheet
    mapping = get_summary_mapping(service, args.sheet_id)
    print(f"Found {len(mapping)} companies in Summary sheet")
    
    # Filter to A-shares
    if args.codes:
        codes_to_process = [c.strip() for c in args.codes.split(',')]
    else:
        codes_to_process = [code for code in mapping.keys() if is_ashare(code)]
    
    print(f"Processing {len(codes_to_process)} A-share companies: {codes_to_process}")
    
    for code in codes_to_process:
        if code not in mapping:
            print(f"  SKIP: {code} not found in Summary")
            continue
        
        sheet_name = mapping[code]
        update_kcfjcxsyjlr(service, args.sheet_id, sheet_name, code, dry_run=args.dry_run)


if __name__ == '__main__':
    main()
