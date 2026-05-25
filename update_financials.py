#!/usr/bin/env python3
"""
Update financial data from CIQ Excel files to Google Sheets.

Strategy: Match by item name (column B in Google Sheets), NOT by row number.
"""

import sys
import os
import re
import json
import argparse

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

import xlrd
from datetime import datetime, timedelta
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build


# ── Configuration ──────────────────────────────────────────────────────────

GOOGLE_SHEET_ID = "1huXdbAgYR2xul5CDtOmuoCjBKGwQu69XB9_AcooRPC0"
GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')

SECTION_MAP = {
    "Income Statement": "Income Statement",
    "Balance Sheet": "Balance Sheet",
    "Cash Flow": "Cash Flow",
}


# ── Helpers ────────────────────────────────────────────────────────────────

def parse_num(v):
    if v is None or v == '' or v == '-':
        return ''
    try:
        return float(v)
    except (ValueError, TypeError):
        return str(v)


def normalize_name(name):
    return str(name).strip().lower()


def serial_to_year(serial):
    """Convert Excel serial date to year."""
    try:
        s = float(serial)
        if 30000 < s < 70000:
            dt = datetime(1899, 12, 30) + timedelta(days=int(s))
            return str(dt.year)
    except (ValueError, TypeError, OverflowError):
        pass
    return None


def extract_header_info(header_str):
    """Extract (year, is_ltm, label) from Excel header.
    
    Returns:
        (year_str_or_None, is_ltm_bool, display_label)
    """
    s = str(header_str).strip().replace('\n', ' ').replace('\r', '')
    if not s:
        return None, False, ''

    # Excel serial date (Balance Sheet)
    year = serial_to_year(s)
    if year:
        return year, False, year

    is_ltm = 'LTM' in s.upper()

    # Find 4-digit year
    match = re.search(r'(\d{4})', s)
    if match:
        year = match.group(1)
        if is_ltm:
            return year, True, s
        return year, False, year

    return None, is_ltm, s


def col_to_letter(col_idx):
    result = ''
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_sheet_grid_size(service, spreadsheet_id, sheet_name):
    """Get actual grid dimensions (row_count, col_count) of a sheet."""
    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId,gridProperties(rowCount,columnCount)))'
    ).execute()
    for s in result.get('sheets', []):
        if s['properties']['title'] == sheet_name:
            gp = s['properties'].get('gridProperties', {})
            return gp.get('rowCount', 1000), gp.get('columnCount', 26)
    return 1000, 26


# ── Read Excel ─────────────────────────────────────────────────────────────

def read_excel_sheet(wb, sheet_name):
    try:
        sheet = wb.sheet_by_name(sheet_name)
    except xlrd.XLRDError:
        return None, None

    # Find title row
    title_row_idx = None
    for i in range(sheet.nrows):
        row = [sheet.cell_value(i, j) for j in range(sheet.ncols)]
        first = str(row[0]).strip().lower() if row[0] else ''
        if first == sheet_name.lower():
            title_row_idx = i
            break

    if title_row_idx is None:
        return None, None

    header_row_idx = title_row_idx + 1
    header_values = [sheet.cell_value(header_row_idx, j) for j in range(sheet.ncols)]

    # Parse data columns
    col_info = []  # list of (excel_col_idx, year, is_ltm, display_label)
    for j in range(1, sheet.ncols):
        year, is_ltm, label = extract_header_info(header_values[j])
        if year:
            col_info.append((j, year, is_ltm, label))

    if not col_info:
        return None, None

    # Find data start
    data_start = header_row_idx + 1
    while data_start < sheet.nrows:
        row = [sheet.cell_value(data_start, j) for j in range(sheet.ncols)]
        first = str(row[0]).strip().lower() if row[0] else ''
        if first == 'currency' or not any(row):
            data_start += 1
        else:
            break

    # Read items
    exact_skip = {'filing date', 'restatement type', 'calculation type'}
    items = {}
    for i in range(data_start, sheet.nrows):
        row = [sheet.cell_value(i, j) for j in range(sheet.ncols)]
        name = ''
        if row[0] and str(row[0]).strip():
            name = str(row[0]).strip()
        elif row[1] and str(row[1]).strip():
            name = str(row[1]).strip()
        else:
            continue

        # Skip section headers and metadata rows
        if name.upper() in ('ASSETS', 'LIABILITIES', 'EQUITY'):
            continue
        if name.lower() in exact_skip:
            continue
        if not name.strip():
            continue

        values = [parse_num(row[j]) for j, _, _, _ in col_info]
        items[normalize_name(name)] = {
            'original_name': name,
            'values': values,
        }

    return items, col_info


# ── Read Google Sheets ─────────────────────────────────────────────────────

def read_gs_section(service, spreadsheet_id, sheet_name, section_header):
    # Dynamic range based on actual grid size
    row_count, col_count = get_sheet_grid_size(service, spreadsheet_id, sheet_name)
    end_col_letter = col_to_letter(col_count - 1)
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!A1:{end_col_letter}{row_count}"
    ).execute()
    rows = result.get('values', [])

    section_start = None
    for i, row in enumerate(rows):
        if row and row[0].strip().lower() == section_header.lower():
            section_start = i
            break

    if section_start is None:
        return None, None, []

    # Find section end
    section_headers = ('income statement', 'balance sheet', 'cash flow',
                       'key stats', 'supplemental', 'multiples', 'ratios',
                       'segments', 'capitalization')
    section_end = len(rows)
    for i in range(section_start + 1, len(rows)):
        row = rows[i]
        if row and row[0].strip().lower() in section_headers:
            section_end = i
            break

    header_row = rows[0] if rows else []

    item_to_row = {}
    for i in range(section_start + 1, section_end):
        if i >= len(rows):
            break
        row = rows[i]
        if len(row) > 1 and row[1].strip():
            name = row[1].strip()
            item_to_row[normalize_name(name)] = i

    return item_to_row, header_row, rows


# ── Key Stats formula copy ─────────────────────────────────────────────────

def copy_key_stats_formulas(service, spreadsheet_id, sheet_name, source_col, target_cols, target_sheet_id, grid_width):
    """Copy formulas from source column to target columns in Key Stats section.
    Also copies numberFormat from source cell to preserve display format."""
    end_col_letter = col_to_letter(grid_width - 1)
    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[f"'{sheet_name}'!A2:{end_col_letter}26"],
        includeGridData=True
    ).execute()

    sheets_data = result.get('sheets', [])
    if not sheets_data:
        return

    data = sheets_data[0].get('data', [])
    if not data:
        return

    row_data = data[0].get('rowData', [])
    
    source_letter = col_to_letter(source_col)
    
    requests = []
    for row_idx, row in enumerate(row_data):
        values = row.get('values', [])
        if source_col < len(values):
            cell = values[source_col]
            val = cell.get('userEnteredValue', {})
            if 'formulaValue' in val:
                original_formula = val['formulaValue']
                gs_row = row_idx + 2  # 1-indexed (header is row 1, data starts row 2)
                
                # Get source cell's number format (if any)
                source_fmt = cell.get('effectiveFormat', {}).get('numberFormat', {})
                
                for target_col in target_cols:
                    target_letter = col_to_letter(target_col)
                    new_formula = re.sub(
                        rf'\b{source_letter}(\d)',
                        rf'{target_letter}\1',
                        original_formula
                    )
                    new_formula = new_formula.replace(f'{source_letter}:{source_letter}', f'{target_letter}:{target_letter}')
                    
                    cell_value = {
                        'userEnteredValue': {'formulaValue': new_formula},
                    }
                    fields = 'userEnteredValue'
                    
                    # Copy number format from source
                    if source_fmt:
                        cell_value['userEnteredFormat'] = {'numberFormat': source_fmt}
                        fields += ',userEnteredFormat'
                    
                    requests.append({
                        'updateCells': {
                            'range': {
                                'sheetId': target_sheet_id,
                                'startRowIndex': gs_row - 1,
                                'endRowIndex': gs_row,
                                'startColumnIndex': target_col,
                                'endColumnIndex': target_col + 1,
                            },
                            'rows': [{'values': [cell_value]}],
                            'fields': fields,
                        }
                    })

    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests}
        ).execute()
        print(f"  ✓ Copied {len(requests)} Key Stats formulas (with format) to columns {', '.join(col_to_letter(c) for c in target_cols)}")


def is_eps_item(norm_name):
    """Check if the item is EPS-related (needs 2 decimal format)."""
    return 'eps' in norm_name or 'per share' in norm_name


# ── Main Logic ─────────────────────────────────────────────────────────────

def process_excel_to_gs(excel_path, gs_sheet_name, spreadsheet_id=None, dry_run=False):
    if spreadsheet_id is None:
        spreadsheet_id = GOOGLE_SHEET_ID

    print(f"\n{'='*60}")
    print(f"Processing: {excel_path}")
    print(f"Spreadsheet: {spreadsheet_id[:30]}...")
    print(f"Target Google Sheet: '{gs_sheet_name}'")
    print(f"{'='*60}")

    wb = xlrd.open_workbook(excel_path)
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    service = build('sheets', 'v4', credentials=creds)
    
    # Resolve sheet name: try exact match, then without "财务" suffix
    spreadsheet_meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId,gridProperties(rowCount,columnCount)))'
    ).execute()
    actual_titles = {s['properties']['title']: s['properties'] for s in spreadsheet_meta.get('sheets', [])}
    
    if gs_sheet_name not in actual_titles:
        # Try without "财务" suffix
        alt_name = gs_sheet_name.replace('财务', '')
        if alt_name in actual_titles:
            print(f"  Note: Sheet '{gs_sheet_name}' not found, using '{alt_name}'")
            gs_sheet_name = alt_name
        else:
            raise ValueError(f"Sheet '{gs_sheet_name}' (or '{alt_name}') not found in spreadsheet")
    
    props = actual_titles[gs_sheet_name]
    target_sheet_id = props['sheetId']
    gp = props.get('gridProperties', {})
    row_count = gp.get('rowCount', 1000)
    col_count = gp.get('columnCount', 26)
    
    # Read GS header (dynamic range based on actual grid width)
    end_col_letter = col_to_letter(col_count - 1)
    header_result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{gs_sheet_name}'!A1:{end_col_letter}1"
    ).execute()
    gs_header = header_result.get('values', [[]])[0]
    
    # Pad gs_header to cover full grid width (API only returns cells with content)
    while len(gs_header) < col_count:
        gs_header.append('')
    
    # ── Column Planning (two-pointer: match by year, no insert/delete) ──
    
    # Parse GS header into list of (col_idx, year_or_None)
    gs_cols = []  # [(col_idx, year_str_or_None), ...]
    for j, val in enumerate(gs_header):
        val = str(val).strip()
        if val:
            match = re.search(r'(\d{4})', val)
            gs_cols.append((j, match.group(1) if match else None))
        else:
            gs_cols.append((j, None))  # empty column
    
    # Read Excel columns from first available sheet (CIQ gives same years for all sheets)
    excel_col_plan = []  # (excel_col_idx, year, is_ltm, label)
    for excel_sheet in SECTION_MAP:
        try:
            s = wb.sheet_by_name(excel_sheet)
        except xlrd.XLRDError:
            continue
        for i in range(s.nrows):
            row = [s.cell_value(i, j) for j in range(s.ncols)]
            first = str(row[0]).strip().lower() if row[0] else ''
            if first == excel_sheet.lower():
                hdr = [s.cell_value(i+1, j) for j in range(s.ncols)]
                for j in range(1, s.ncols):
                    year, is_ltm, label = extract_header_info(hdr[j])
                    if year:
                        excel_col_plan.append((j, year, is_ltm, label))
                break
        if excel_col_plan:
            break
    
    # Two-pointer matching: a = GS pointer (never goes back), b = Excel columns (left to right)
    plan_col_to_gs_col = {}  # excel_col_idx -> gs_col_idx
    header_updates = []       # [(gs_col_idx, new_header_text)]
    
    # Find the first data column (skip label columns A-C)
    # Data columns start where the first year header appears
    a = 0
    for i, (gs_j, gs_year) in enumerate(gs_cols):
        if gs_year is not None:
            a = i
            break
    
    # Find the boundary between annual and quarterly columns
    quarter_boundary = len(gs_cols)
    for i, (gs_j, gs_year) in enumerate(gs_cols):
        val = str(gs_header[gs_j]).strip() if gs_j < len(gs_header) else ''
        if re.match(r'Q[1-4]\s', val):
            quarter_boundary = i
            break
    
    for excel_col_idx, year, is_ltm, label in excel_col_plan:
        # For annual data (not LTM), limit search to before quarterly columns
        search_end = quarter_boundary if not is_ltm else len(gs_cols)
        
        # Scan GS from position `a` forward looking for year match
        matched = False
        for i in range(a, search_end):
            gs_j, gs_year = gs_cols[i]
            if gs_year == year:
                # Verify it's a pure year or LTM header, not a quarter header
                hdr_val = str(gs_header[gs_j]).strip()
                if re.match(r'^\d{4}$', hdr_val) or re.search(r'LTM', hdr_val, re.IGNORECASE) or not hdr_val:
                    plan_col_to_gs_col[excel_col_idx] = gs_j
                    if gs_header[gs_j] != label:
                        header_updates.append((gs_j, label))
                    print(f"  Excel col {excel_col_idx} year {year} ('{label}') → GS col {gs_j} ('{gs_header[gs_j]}') [match]")
                    a = i + 1
                    matched = True
                    break
        
        if not matched:
            # Find next empty column from position `a` forward
            found_empty = False
            for i in range(a, search_end):
                gs_j, gs_year = gs_cols[i]
                if gs_year is None:  # empty column
                    plan_col_to_gs_col[excel_col_idx] = gs_j
                    header_updates.append((gs_j, label))
                    gs_cols[i] = (gs_j, year)  # mark as occupied
                    print(f"  Excel col {excel_col_idx} year {year} ('{label}') → GS col {gs_j} [empty col]")
                    a = i + 1
                    found_empty = True
                    break
            
            if not found_empty:
                print(f"  WARNING: No empty column available for Excel col {excel_col_idx} year {year} ('{label}')")
    
    print(f"  Column mapping: {plan_col_to_gs_col}")
    if header_updates:
        print(f"  Header updates needed: {header_updates}")
    
    # Execute header updates (if not dry-run)
    if not dry_run and header_updates:
        requests = []
        for gs_col, new_label in header_updates:
            requests.append({
                'updateCells': {
                    'range': {
                        'sheetId': target_sheet_id,
                        'startRowIndex': 0, 'endRowIndex': 1,
                        'startColumnIndex': gs_col,
                        'endColumnIndex': gs_col + 1,
                    },
                    'rows': [{'values': [{'userEnteredValue': {'stringValue': new_label}}]}],
                    'fields': 'userEnteredValue',
                }
            })
        if requests:
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': requests}
            ).execute()
            print(f"  ✓ Updated {len(requests)} headers")

    total_updates = 0
    total_matched = 0
    total_unmatched = 0
    key_stats_cols_written = set()  # Track which columns got data in Key Stats

    for excel_sheet, gs_section in SECTION_MAP.items():
        print(f"\n--- {excel_sheet} → {gs_section} ---")

        excel_items, excel_col_info = read_excel_sheet(wb, excel_sheet)
        if excel_items is None:
            print(f"  SKIP: Sheet not found or no data")
            continue

        gs_mapping, gs_header_row, gs_rows = read_gs_section(
            service, spreadsheet_id, gs_sheet_name, gs_section
        )
        if gs_mapping is None:
            print(f"  SKIP: Section not found")
            continue

        # Map Excel columns to GS columns (by excel_col_idx from plan)
        excel_col_to_gs_col = {}
        for excel_col_idx, year, is_ltm, label in excel_col_info:
            if excel_col_idx in plan_col_to_gs_col:
                excel_col_to_gs_col[excel_col_idx] = plan_col_to_gs_col[excel_col_idx]
                print(f"  Excel col {excel_col_idx} ({label}) → GS col {plan_col_to_gs_col[excel_col_idx]}")
            else:
                print(f"  WARNING: Excel col {excel_col_idx} ({label}) → no GS column mapped")

        if not excel_col_to_gs_col:
            continue

        sorted_excel_cols = sorted(excel_col_to_gs_col.keys())

        updates = []  # (gs_col, gs_row_0idx, val, is_eps)
        matched = 0
        unmatched = 0

        for norm_name, item_data in excel_items.items():
            if norm_name in gs_mapping:
                gs_row = gs_mapping[norm_name]  # 0-indexed
                is_eps = is_eps_item(norm_name)
                for excel_col_idx, excel_col in enumerate(sorted_excel_cols):
                    gs_col = excel_col_to_gs_col[excel_col]
                    if excel_col_idx < len(item_data['values']):
                        val = item_data['values'][excel_col_idx]
                        if val != '':
                            updates.append((gs_col, gs_row, val, is_eps))
                            key_stats_cols_written.add(gs_col)
                matched += 1
            else:
                unmatched += 1

        total_matched += matched
        total_unmatched += unmatched

        if dry_run:
            print(f"  [DRY RUN] Would write {len(updates)} cells ({matched} matched, {unmatched} unmatched)")
        else:
            if updates:
                requests = []
                for gs_col, gs_row, val, is_eps in updates:
                    cell_value = {}
                    if isinstance(val, (int, float)):
                        num_fmt = '#,##0.00' if is_eps else '#,##0'
                        cell_value['userEnteredValue'] = {'numberValue': val}
                        cell_value['userEnteredFormat'] = {'numberFormat': {'type': 'NUMBER', 'pattern': num_fmt}}
                    else:
                        cell_value['userEnteredValue'] = {'stringValue': str(val)}
                    
                    fields = 'userEnteredValue'
                    if 'userEnteredFormat' in cell_value:
                        fields += ',userEnteredFormat'
                    
                    requests.append({
                        'updateCells': {
                            'range': {
                                'sheetId': target_sheet_id,
                                'startRowIndex': gs_row,
                                'endRowIndex': gs_row + 1,
                                'startColumnIndex': gs_col,
                                'endColumnIndex': gs_col + 1,
                            },
                            'rows': [{'values': [cell_value]}],
                            'fields': fields,
                        }
                    })
                
                if requests:
                    service.spreadsheets().batchUpdate(
                        spreadsheetId=spreadsheet_id,
                        body={'requests': requests}
                    ).execute()
                    print(f"  ✓ Wrote {len(requests)} cells for {matched} items")
            else:
                print(f"  No data to write ({matched} matched)")

        total_updates += len(updates)
        
        # Special handling for Payout Ratio in Income Statement
        if gs_section == "Income Statement" and not dry_run:
            payout_row = next((v for k, v in gs_mapping.items() if re.match(r'payout ratio', k)), None)
            dps_row = next((v for k, v in gs_mapping.items() if re.match(r'dividends per share', k)), None)
            eps_row = next((v for k, v in gs_mapping.items() if re.match(r'basic eps', k)), None)
            
            if payout_row and dps_row and eps_row:
                # Write formula for each column that was written
                payout_requests = []
                for excel_col_idx, excel_col in enumerate(sorted_excel_cols):
                    gs_col = excel_col_to_gs_col[excel_col]
                    if gs_col in key_stats_cols_written:
                        dps_cell = f'{col_to_letter(gs_col)}{dps_row + 1}'
                        eps_cell = f'{col_to_letter(gs_col)}{eps_row + 1}'
                        formula = f'={dps_cell}/{eps_cell}'
                        
                        payout_requests.append({
                            'updateCells': {
                                'range': {
                                    'sheetId': target_sheet_id,
                                    'startRowIndex': payout_row,
                                    'endRowIndex': payout_row + 1,
                                    'startColumnIndex': gs_col,
                                    'endColumnIndex': gs_col + 1,
                                },
                                'rows': [{'values': [{
                                    'userEnteredValue': {'formulaValue': formula},
                                    'userEnteredFormat': {'numberFormat': {'type': 'PERCENT', 'pattern': '0.0%'}}
                                }]}],
                                'fields': 'userEnteredValue,userEnteredFormat',
                            }
                        })
                
                if payout_requests:
                    service.spreadsheets().batchUpdate(
                        spreadsheetId=spreadsheet_id,
                        body={'requests': payout_requests}
                    ).execute()
                    print(f"  ✓ Wrote {len(payout_requests)} Payout Ratio formulas")

    # Copy Key Stats formulas to new columns
    if not dry_run and key_stats_cols_written:
        # Find the source column (last column with existing formulas before new data)
        # This is the column just before the first new column we wrote to
        first_new_col = min(key_stats_cols_written)
        source_col = first_new_col - 1
        target_cols = sorted(key_stats_cols_written)
        
        if source_col >= 0:
            copy_key_stats_formulas(service, spreadsheet_id, gs_sheet_name, source_col, target_cols, target_sheet_id, col_count)

    print(f"\n{'='*60}")
    print(f"Summary: {total_matched} items matched, {total_unmatched} unmatched, "
          f"{total_updates} cells {'would be ' if dry_run else ''}written")
    print(f"{'='*60}")

    # Update Capital Structure Details from Excel
    if not dry_run:
        update_capital_structure_details(wb, service, spreadsheet_id, gs_sheet_name)


# ── Capital Structure Details ──────────────────────────────────────────────

def update_capital_structure_details(wb, service, spreadsheet_id, gs_sheet_name):
    """Read Capital Structure Details from Excel and write to GS 资本结构 tab.

    If Excel doesn't have the tab, skip. If multiple FY sections exist, use the latest one.
    Delete existing rows for this company in GS first, then insert new data.
    """
    EXCEL_TAB_NAME = "Capital Structure Details"
    GS_TAB_NAME = "资本结构"

    # Check if Excel has the tab
    try:
        sheet = wb.sheet_by_name(EXCEL_TAB_NAME)
    except xlrd.XLRDError:
        return  # Skip if tab doesn't exist

    if sheet.nrows < 5:
        return

    # Extract company Chinese name from gs_sheet_name (e.g. "福耀玻璃财务" -> "福耀玻璃")
    company_name = gs_sheet_name.replace("财务", "")

    # Find all FY sections in the Excel tab
    # Each section starts with a row containing "FY" in col 0 (e.g. "FY 2025 (Dec-31-2025) Capital Structure...")
    sections = []
    for i in range(sheet.nrows):
        val = str(sheet.cell_value(i, 0)).strip()
        if val.startswith("FY ") and "Capital Structure" in val:
            sections.append(i)

    if not sections:
        print("  Capital Structure Details: no FY sections found, skipping")
        return

    # Use the latest (last) section
    latest_start = sections[-1]
    print(f"  Capital Structure Details: found {len(sections)} FY sections, using latest (row {latest_start + 1})")

    # Find the column header row (row with "Description", "Type", etc.)
    hdr_row = None
    for i in range(latest_start, min(latest_start + 5, sheet.nrows)):
        row_vals = [str(sheet.cell_value(i, j)).strip().lower() for j in range(min(sheet.ncols, 10))]
        if "description" in row_vals:
            hdr_row = i
            break

    if hdr_row is None:
        print("  Capital Structure Details: no header row found, skipping")
        return

    # Collect data rows until empty row or next FY section
    data_rows = []
    for i in range(hdr_row + 1, sheet.nrows):
        col_a = str(sheet.cell_value(i, 0)).strip()
        if not col_a:
            # Stop at first empty row after data
            if not data_rows:
                continue  # skip leading empty rows
            break
        if col_a.startswith("FY "):
            break  # next section

        # Collect columns 0-8 from Excel (Description through Convertible), total 9 cols
        row_data = [col_a]  # Description is col A in Excel
        for j in range(1, 9):
            if j < sheet.ncols:
                row_data.append(_format_excel_value(sheet.cell_value(i, j)))
            else:
                row_data.append('')
        data_rows.append(row_data)

    if not data_rows:
        print("  Capital Structure Details: no data rows found, skipping")
        return

    print(f"  Capital Structure Details: {len(data_rows)} data rows from Excel")

    # Build the full GS row: [company_name, col0, col1, ..., col8]
    gs_rows = []
    for row in data_rows:
        gs_rows.append([company_name] + row)

    # Read existing 资本结构 tab data with grid info to get actual row positions
    gs_meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=f"{GS_TAB_NAME}!A1:J10000",
        includeGridData=True
    ).execute()

    existing_values = []
    last_data_row = 0
    for sheet_data in gs_meta.get('sheets', []):
        row_data = sheet_data.get('data', [{}])[0].get('rowData', [])
        for i, row in enumerate(row_data):
            actual_row = i + 1
            cells = row.get('values', [])
            row_vals = []
            for cell in cells:
                row_vals.append(cell.get('formattedValue', ''))
            # Pad to 10 columns
            while len(row_vals) < 10:
                row_vals.append('')
            existing_values.append(row_vals)
            if any(v for v in row_vals):
                last_data_row = actual_row

    # Find row indices (1-based for API) where column A matches company_name
    rows_to_delete = []
    for i, row in enumerate(existing_values):
        if row and len(row) > 0 and str(row[0]).strip() == company_name:
            rows_to_delete.append(i + 1)

    if rows_to_delete:
        print(f"  Capital Structure Details: deleting {len(rows_to_delete)} existing rows for {company_name}")
        # Delete from bottom to top to preserve row indices
        for row_idx in reversed(rows_to_delete):
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': [{
                    'deleteDimension': {
                        'range': {
                            'sheetId': _get_sheet_id(service, spreadsheet_id, GS_TAB_NAME),
                            'dimension': 'ROWS',
                            'startIndex': row_idx - 1,
                            'endIndex': row_idx,
                        }
                    }
                }]}
            ).execute()

        # Re-read the sheet to find the actual last data row after deletion
        gs_meta2 = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            ranges=f"{GS_TAB_NAME}!A1:J10000",
            includeGridData=True
        ).execute()
        last_data_row = 0
        for sheet_data in gs_meta2.get('sheets', []):
            row_data = sheet_data.get('data', [{}])[0].get('rowData', [])
            for i, row in enumerate(row_data):
                actual_row = i + 1
                cells = row.get('values', [])
                row_vals = [c.get('formattedValue', '') for c in cells]
                while len(row_vals) < 10:
                    row_vals.append('')
                if any(v for v in row_vals):
                    last_data_row = actual_row

    # Append after the last row with data
    first_data_row = last_data_row if last_data_row > 0 else 1

    print(f"  Capital Structure Details: writing {len(gs_rows)} rows starting at row {first_data_row + 1}")

    # Write data rows
    update_range = f"{GS_TAB_NAME}!A{first_data_row + 1}:J{first_data_row + len(gs_rows)}"
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=update_range,
        valueInputOption='RAW',
        body={'values': gs_rows}
    ).execute()


def _get_sheet_id(service, spreadsheet_id, sheet_name):
    """Get the numeric sheetId for a given sheet name."""
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in meta.get('sheets', []):
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    raise ValueError(f"Sheet '{sheet_name}' not found")


def _format_excel_value(val):
    """Clean up an Excel cell value for GS display.

    Converts Excel serial dates (float > 30000) to date strings like 'Dec-31-2025'.
    """
    if val is None or val == '':
        return ''
    if isinstance(val, float) and val > 30000 and val < 100000:
        # Excel serial date
        try:
            dt = datetime(1899, 12, 30) + timedelta(days=int(val))
            return dt.strftime('%b-%d-%Y').upper()
        except (ValueError, OverflowError):
            pass
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val).strip()


# ── Build Mapping from Summary Sheet ──────────────────────────────────────

def build_mapping_from_summary(service, spreadsheet_id, summary_sheet='Summary'):
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{summary_sheet}'!A1:AZ2"
    ).execute()
    rows = result.get('values', [])
    
    if len(rows) < 2:
        return {}
    
    codes = rows[0]
    names = rows[1]
    
    mapping = {}
    for j in range(len(codes)):
        code = str(codes[j]).strip() if j < len(codes) else ''
        name = str(names[j]).strip() if j < len(names) else ''
        if code and name:
            mapping[code] = f"{name}财务"
    
    return mapping


def extract_code_from_filename(filename):
    base = os.path.splitext(filename)[0]
    match = re.search(r'(?:SHSE|SZSE|SSE|HKEX|NYSE|NASDAQ)\s+(\d+)', base, re.IGNORECASE)
    if match:
        return match.group(1)
    matches = re.findall(r'\b(\d{4,6})\b', base)
    if matches:
        return matches[-1]
    # Ticker: skip exchange names, find last uppercase 2-5 char token
    exchange_names = {'SHSE', 'SZSE', 'SSE', 'HKEX', 'NYSE', 'NASDAQ'}
    ticker_matches = re.findall(r'\b([A-Z]{2,5})\b', base)
    tickers = [t for t in ticker_matches if t not in exchange_names]
    if tickers:
        return tickers[-1]
    return None


# ── Batch Processing ───────────────────────────────────────────────────────

def batch_process(directory, spreadsheet_ids=None, dry_run=False):
    import glob
    from collections import defaultdict

    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    service = build('sheets', 'v4', credentials=creds)

    if not spreadsheet_ids:
        json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'industry_spreadsheets.json')
        if os.path.exists(json_path):
            with open(json_path, 'r', encoding='utf-8') as f:
                gs_data = json.load(f)
            spreadsheet_ids = [info['spreadsheet_id'] for info in gs_data.values()]
            print(f"Loaded {len(spreadsheet_ids)} spreadsheets from industry_spreadsheets.json")
        else:
            spreadsheet_ids = [GOOGLE_SHEET_ID]
            print(f"Using default spreadsheet: {GOOGLE_SHEET_ID}")

    routing = {}
    for sid in spreadsheet_ids:
        print(f"\nReading Summary sheet for spreadsheet {sid[:20]}...")
        summary_mapping = build_mapping_from_summary(service, sid)
        print(f"  Found {len(summary_mapping)} companies: {list(summary_mapping.keys())}")
        for code, sheet_name in summary_mapping.items():
            routing[code] = (sid, sheet_name)

    xls_files = glob.glob(os.path.join(directory, '*.xls')) + \
                glob.glob(os.path.join(directory, '*.xlsx'))
    xls_files.sort()

    if not xls_files:
        print(f"\nNo Excel files found in {directory}")
        return

    print(f"\nFound {len(xls_files)} Excel files")

    files_by_spreadsheet = defaultdict(list)
    unmatched_files = []

    for xls_path in xls_files:
        filename = os.path.basename(xls_path)
        code = extract_code_from_filename(filename)
        
        if code and code in routing:
            sid, sheet_name = routing[code]
            files_by_spreadsheet[sid].append((xls_path, sheet_name, code))
        else:
            unmatched_files.append((filename, code))

    print(f"\nRouting summary:")
    for sid, files in files_by_spreadsheet.items():
        print(f"  Spreadsheet {sid[:20]}... → {len(files)} files")
        for _, sheet_name, code in files:
            print(f"    {code:8s} → {sheet_name}")
    
    if unmatched_files:
        print(f"\n  Unmatched ({len(unmatched_files)} files):")
        for fname, code in unmatched_files:
            print(f"    {fname} (code='{code}')")

    total_processed = 0
    for sid in spreadsheet_ids:
        if sid not in files_by_spreadsheet:
            continue
        
        files = files_by_spreadsheet[sid]
        print(f"\n{'='*60}")
        print(f"Processing spreadsheet {sid[:20]}... ({len(files)} files)")
        print(f"{'='*60}")
        
        for xls_path, sheet_name, code in files:
            try:
                process_excel_to_gs(xls_path, sheet_name, spreadsheet_id=sid, dry_run=dry_run)
                total_processed += 1
            except Exception as e:
                print(f"\n  ✗ FAILED: {code} ({sheet_name}): {e}")

    print(f"\n{'='*60}")
    print(f"Batch complete: {total_processed} processed, {len(unmatched_files)} unmatched")
    print(f"{'='*60}")


# ── Entry Point ────────────────────────────────────────────────────────────

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Update financials from Excel to Google Sheets')
    parser.add_argument('--spreadsheet-id', help='Google Spreadsheet ID (for single file mode)')
    parser.add_argument('--spreadsheets', help='Comma-separated list of spreadsheet IDs for batch mode')
    parser.add_argument('--dry-run', action='store_true', help='Preview changes without writing')
    parser.add_argument('--batch', action='store_true', help='Batch process all files in directory')
    parser.add_argument('excel_path', nargs='?', default=None, help='Path to the Excel file (or directory for batch)')
    parser.add_argument('gs_sheet_name', nargs='?', default=None, help='Google Sheets sheet name')
    args = parser.parse_args()

    if args.batch:
        directory = args.excel_path or './CIQ_Financials'
        spreadsheet_ids = [s.strip() for s in args.spreadsheets.split(',')] if args.spreadsheets else None
        batch_process(directory, spreadsheet_ids=spreadsheet_ids, dry_run=args.dry_run)
    else:
        if not args.excel_path or not args.gs_sheet_name:
            parser.print_help()
            print("\nError: excel_path and gs_sheet_name are required for single file mode")
            sys.exit(1)

        if not os.path.exists(args.excel_path):
            print(f"ERROR: File not found: {args.excel_path}")
            sys.exit(1)

        process_excel_to_gs(args.excel_path, args.gs_sheet_name, spreadsheet_id=args.spreadsheet_id, dry_run=args.dry_run)
