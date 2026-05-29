#!/usr/bin/env python3
"""
Create a new company tab from scratch in a Google Spreadsheet.

Builds structure from scratch (no template copy):
1. Fill section headers in column A (Key Stats, IS, BS, CF)
2. Fill item names in column B (Key Stats from GS, IS/BS/CF from Excel)
3. Fill year/quarter headers in row 1
4. Write Key Stats formulas by item name resolution

Usage:
    python create_company_tab.py --code 600519 --name "贵州茅台" --excel file.xls
    python create_company_tab.py --code 600519 --name "贵州茅台"  # auto-routes via industry_spreadsheets.json
"""

import sys
import os
import re
import json
import csv
import subprocess
import xlrd
from pathlib import Path
from datetime import datetime

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')

# ── Key Stats item names ──
KEY_STATS_ITEMS = [
    "Net Working Capital",
    "Net Income",
    "Gross Margin",
    "Op. Margin",
    "Net Margin",
    "ROE",
    "Interest Coverage Ratio",
    "Interest and Rental Exp Coverage Ratio",
    "D&A from Company",
    "Net Current Asset",
    "Common Equity",
    "Common Stock",
    "Tangible Book Value",
    "Dividends per Share",
    "Total Liabilities And Equity",
    "Cash from Ops.",
    "Minority Interest",
    "FCFF",
    "Basic EPS",
    "Net Debt",
    "EBITDA",
    "扣非净利润",
    "SBC",
    "Total Revenue",
    "ROIC (资本来源法)",
    "ROIC (资产法)",
    "ROIC (Greenblatt)",
    "Payout Ratio %",
]

# Additional Key Stats items that may appear (from GS observation)
KEY_STATS_SUPPLEMENTAL = [
    "Receivables/Revenue",
    "Unearned Revenue/Revenue",
]

# Key Stats sub-section headers (used as sub-group labels in column B)
KEY_STATS_SUB_SECTIONS = ('盈利指标', '同比增速')

# YoY growth items (new sub-section)
YOY_ITEMS = [
    "Revenue YoY",
    "Gross Profit YoY",
    "Operating Income YoY",
    "Net Income YoY",
    "扣非净利润 YoY",
]

# Formula templates
FORMULA_TEMPLATES = {
    "Net Working Capital": {
        "formula": "=__C__{Total Current Assets}-__C__{Total Current Liabilities}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "Net Income": {
        "formula": "=__C__{Net Income to Company}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "Gross Margin": {
        "formula": "=__C__{Gross Profit}/__C__{Total Revenue}",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
    "Op. Margin": {
        "formula": "=__C__{Operating Income}/__C__{Total Revenue}",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
    "Net Margin": {
        "formula": "=__C__{Net Income to Company}/__C__{Total Revenue}",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
    "ROE": {
        "formula": "=__C__{Net Income to Company}/(__C__{Total Common Equity}+__PC__{Total Common Equity})*2",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
    "Interest Coverage Ratio": {
        "formula": "=__C__{EBIT}/__C__{Interest Expense}*-1",
        "format": {"type": "NUMBER", "pattern": "#,##0.0"},
    },
    "Interest and Rental Exp Coverage Ratio": {
        "formula": "=__C__{Net Income to Company}/(__C__{Interest Expense}-__C__{Net Rental Exp.})*-1+1",
        "format": {"type": "NUMBER", "pattern": "#,##0.0"},
    },
    "D&A from Company": {
        "formula": "=__C__{Depreciation & Amort., Total}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "Net Current Asset": {
        "formula": "=__C__{Total Current Assets}-__C__{Total Current Liabilities}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "Common Equity": {
        "formula": "=__C__{Total Common Equity}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "Common Stock": {
        "formula": "=__C__{Total Shares Out. on Balance Sheet Date}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "Tangible Book Value": {
        "formula": "=__C__{Tangible Book Value}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "Dividends per Share": {
        "formula": "=__C__{Dividends per Share}",
        "format": {"type": "NUMBER", "pattern": "#,##0.00"},
    },
    "Total Liabilities And Equity": {
        "formula": "=__C__{Total Liabilities And Equity}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "Cash from Ops.": {
        "formula": "=__C__{Cash from Ops.}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "Minority Interest": {
        "formula": "=__C__{!Minority Interest}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "FCFF": {
        "formula": "=__C__{Cash from Ops.}+__C__{Capital Expenditure}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "Basic EPS": {
        "formula": "=__C__{Basic EPS}",
        "format": {"type": "NUMBER", "pattern": "#,##0.00"},
    },
    "Net Debt": {
        "formula": "=__C__{Net Debt}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "EBITDA": {
        "formula": "=__C__{EBITDA}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "Total Revenue": {
        "formula": "=__C__{Total Revenue}",
        "format": {"type": "NUMBER", "pattern": "#,##0"},
    },
    "ROIC (资本来源法)": {
        "formula": "=__C__{EBIT}*(1-__C__{!Effective Tax Rate %})/(__C__{Net Debt}+__C__{Common Equity}+__C__{!Minority Interest}+__PC__{Net Debt}+__PC__{Common Equity}+__PC__{!Minority Interest})*2",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
    # Operating-asset (investing) approach: invested capital from the asset side,
    # = Total Assets - excess cash - non-interest-bearing current liabilities. Two-period avg.
    "ROIC (资产法)": {
        "formula": "=__C__{EBIT}*(1-__C__{!Effective Tax Rate %})/((__C__{Total Assets}-__C__{?Total Cash & ST Investments}-(__C__{Total Current Liabilities}-__C__{?Short-term Borrowings}-__C__{?Curr. Port. of Leases}))+(__PC__{Total Assets}-__PC__{?Total Cash & ST Investments}-(__PC__{Total Current Liabilities}-__PC__{?Short-term Borrowings}-__PC__{?Curr. Port. of Leases})))*2",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
    # Greenblatt/McKinsey tangible-capital approach: pre-tax EBIT over operating
    # working capital + net fixed assets (excludes goodwill/intangibles & excess cash).
    # Textbook form: beginning-of-period capital base (prior column, __PC__), no averaging.
    "ROIC (Greenblatt)": {
        "formula": "=__C__{EBIT}/((__PC__{Total Current Assets}-__PC__{?Total Cash & ST Investments})-(__PC__{Total Current Liabilities}-__PC__{?Short-term Borrowings}-__PC__{?Curr. Port. of Leases})+__PC__{Net Property, Plant & Equipment})",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
    # DPS N()-wrapped (no-dividend '-' -> 0); IFERROR blanks a zero/negative EPS.
    "Payout Ratio %": {
        "formula": "=IFERROR(__C__{!Dividends per Share}/__C__{Basic EPS},)",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
    "Revenue YoY": {
        "formula": "=__C__{Total Revenue}/__PC__{Total Revenue}-1",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
    "Gross Profit YoY": {
        "formula": "=__C__{Gross Profit}/__PC__{Gross Profit}-1",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
    "Operating Income YoY": {
        "formula": "=__C__{Operating Income}/__PC__{Operating Income}-1",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
    "Net Income YoY": {
        "formula": "=__C__{Net Income to Company}/__PC__{Net Income to Company}-1",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
    "扣非净利润 YoY": {
        "formula": "=__C__{扣非净利润}/__PC__{扣非净利润}-1",
        "format": {"type": "PERCENT", "pattern": "0.0%"},
    },
}


def col_to_letter(col_idx):
    """Convert 0-indexed column number to Excel-style letter."""
    result = ''
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_service():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    return build('sheets', 'v4', credentials=creds)


def get_sheet_names(service, spreadsheet_id):
    """Get all sheet names and their properties."""
    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(title,sheetId,gridProperties(rowCount,columnCount)))'
    ).execute()
    sheets = {}
    for s in result.get('sheets', []):
        p = s['properties']
        sheets[p['title']] = {
            'sheetId': p['sheetId'],
            'rowCount': p.get('gridProperties', {}).get('rowCount', 1000),
            'colCount': p.get('gridProperties', {}).get('columnCount', 26),
        }
    return sheets


# ── Read Excel sections ──

def read_excel_items(wb, sheet_name):
    """Read item names from an Excel sheet's column A/B (after the header).

    Preserves empty rows for visual grouping in the output.
    Preserves text formatting (bold, italic) from Excel font.
    Returns a list of (name, text_format) tuples — empty name means a blank row.
    text_format is a dict with keys 'bold', 'italic' (bool).
    """
    try:
        sheet = wb.sheet_by_name(sheet_name)
    except xlrd.XLRDError:
        return None

    # Find title row
    title_row_idx = None
    for i in range(sheet.nrows):
        first = str(sheet.cell_value(i, 0)).strip().lower()
        if first == sheet_name.lower():
            title_row_idx = i
            break
    if title_row_idx is None:
        return None

    # Skip header rows (title + currency/etc.)
    data_start = title_row_idx + 1
    while data_start < sheet.nrows:
        val = str(sheet.cell_value(data_start, 0)).strip().lower()
        if val == 'currency' or not val or val == sheet_name.lower():
            data_start += 1
        else:
            break

    # Read items, preserving empty rows and text formatting
    exact_skip = {'filing date', 'restatement type', 'calculation type'}
    section_headers = {'income statement', 'balance sheet', 'cash flow'}
    items = []
    for i in range(data_start, sheet.nrows):
        col_a = str(sheet.cell_value(i, 0)).strip() if sheet.cell_value(i, 0) else ''
        col_b = str(sheet.cell_value(i, 1)).strip() if sheet.cell_value(i, 1) else ''

        # Stop at next section header in column A
        if col_a.lower() in section_headers:
            break

        if col_a and col_a.lower() in exact_skip:
            continue
        if col_a.upper() in ('ASSETS', 'LIABILITIES', 'EQUITY'):
            continue

        # Determine source column and extract formatting
        use_col = 0 if col_a else 1
        name = col_a if col_a else col_b

        text_format = {'bold': False, 'italic': False}
        if name:  # skip empty rows
            xf_idx = sheet.cell_xf_index(i, use_col)
            xf = wb.xf_list[xf_idx]
            font = wb.font_list[xf.font_index]
            text_format['bold'] = font.weight >= 700
            text_format['italic'] = font.italic

        items.append((name, text_format))

    # Trim trailing empty rows
    while items and items[-1][0] == '':
        items.pop()

    return items


def detect_fiscal_year_end(excel_path):
    """Detect fiscal year end month from Excel header.

    Returns:
        (month, day) tuple, or (12, 31) for natural year.
    """
    wb = xlrd.open_workbook(excel_path)
    for sheet_name in ['Income Statement', 'Balance Sheet', 'Cash Flow']:
        try:
            sheet = wb.sheet_by_name(sheet_name)
        except xlrd.XLRDError:
            continue
        for i in range(sheet.nrows):
            first = str(sheet.cell_value(i, 0)).strip().lower()
            if first == sheet_name.lower():
                # Read header row
                for j in range(1, min(sheet.ncols, 4)):
                    hdr = str(sheet.cell_value(i + 1, j)).strip()
                    # Extract date like "May-31-2007" or "Dec-31-2007"
                    match = re.search(r'([A-Z][a-z]{2})-(\d{2})-(\d{4})', hdr)
                    if match:
                        month_str = match.group(1)
                        day = int(match.group(2))
                        months = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                                  'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
                        month = months.get(month_str, 12)
                        return (month, day)
                break
    return (12, 31)


# ── Read Key Stats from existing GS tab ──

def get_key_stats_from_gs(service, spreadsheet_id, template_name):
    """Get Key Stats item names from an existing tab.

    Handles both old layout (items in column B) and new layout
    (sub-section headers in B, items in column C).
    """
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{template_name}!A1:C40"
    ).execute()
    rows = result.get('values', [])

    key_stats_start = None
    for i, row in enumerate(rows):
        if row and row[0].strip().lower() == 'key stats':
            key_stats_start = i
            break

    if key_stats_start is None:
        print(f"  WARNING: Key Stats not found in {template_name}, using defaults")
        return list(KEY_STATS_ITEMS)

    section_headers = ('income statement', 'balance sheet', 'cash flow',
                       'key stats', 'supplemental', 'multiples', 'ratios',
                       'segments', 'capitalization', 'business segments')
    items = []
    for i in range(key_stats_start + 1, len(rows)):
        row = rows[i]
        b_val = row[1].strip() if len(row) > 1 and row[1] else ''
        c_val = row[2].strip() if len(row) > 2 and row[2] else ''

        # Prefer column C (new layout with sub-sections), fall back to B
        item_val = c_val if c_val else b_val

        if not item_val:
            break
        if item_val.lower() in section_headers:
            break
        if item_val in KEY_STATS_SUB_SECTIONS:
            continue
        items.append(item_val)

    if not items:
        return list(KEY_STATS_ITEMS)

    print(f"  Key Stats items from template: {len(items)} items")
    return items


def find_template_sheet(service, spreadsheet_id):
    """Find any existing financial tab to copy Key Stats from."""
    sheets = get_sheet_names(service, spreadsheet_id)
    for name in sheets:
        if name not in ('Summary', '资本结构') and name.endswith('财务'):
            return name
    return None


# ── Build structure ──

def build_sheet_structure(service, spreadsheet_id, target_sheet_name, excel_path):
    """
    Build the sheet from scratch:
    1. Create sheet with adequate grid
    2. Read Excel IS/BS/CF items
    3. Read Key Stats from existing GS tab
    4. Fill section headers in A, items in B
    5. Fill year/quarter headers in row 1
    """
    sheets = get_sheet_names(service, spreadsheet_id)

    # Detect fiscal year end from Excel
    fy_month, fy_day = detect_fiscal_year_end(excel_path)
    is_natural = (fy_month == 12 and fy_day == 31)
    print(f"  Fiscal year end: {fy_month}/{fy_day} ({'natural' if is_natural else 'non-natural'})")

    # Read Excel items
    wb = xlrd.open_workbook(excel_path, formatting_info=True)
    is_items = read_excel_items(wb, 'Income Statement')
    bs_items = read_excel_items(wb, 'Balance Sheet')
    cf_items = read_excel_items(wb, 'Cash Flow')

    if not is_items:
        print("  ERROR: Could not read Income Statement from Excel")
        return False

    print(f"  Excel: IS={len(is_items)}, BS={len(bs_items or [])}, CF={len(cf_items or [])} items")

    # Find Key Stats template
    template = find_template_sheet(service, spreadsheet_id)
    ks_items = []
    if template:
        ks_items = get_key_stats_from_gs(service, spreadsheet_id, template)
    else:
        ks_items = list(KEY_STATS_ITEMS)

    # Convert Key Stats items to (name, text_format) tuples for uniform handling
    ks_items_tuples = [(name, {'bold': False, 'italic': False}) for name in ks_items]
    yoy_items_tuples = [(name, {'bold': False, 'italic': False}) for name in YOY_ITEMS]

    # ── Calculate row positions ──
    # Row 1 (0-idx): headers
    # Row 2 (0-idx=1): "Key Stats" (A2)
    # Row 3 (0-idx=2): "盈利指标" sub-header (B3)
    # Row 4-N: Key Stats items (C4 onwards)
    # Blank row
    # "同比增速" sub-header (B)
    # YoY items (C)
    # Blank row, then "Income Statement", etc.

    sections = []
    current_row = 1  # 0-indexed: row 0 = headers

    # Key Stats: A2 = "Key Stats" (0-idx=1)
    ks_row = 1
    sections.append(('Key Stats', ks_row, []))  # header only, no items in tuple
    current_row = ks_row + 1  # row 2 (0-idx)

    # "盈利指标" sub-header + items in column C
    sections.append(('', current_row, [('盈利指标', {'bold': True, 'italic': False})], 1))  # sub-header in B
    current_row += 1  # items start here (row 3, 0-idx)
    sections.append(('', current_row, ks_items_tuples, 2))  # items in C
    current_row += len(ks_items_tuples)

    # Blank row
    sections.append(('', current_row, []))
    current_row += 1

    # "同比增速" sub-header + YoY items in column C
    sections.append(('', current_row, [('同比增速', {'bold': True, 'italic': False})], 1))  # sub-header in B
    current_row += 1
    sections.append(('', current_row, yoy_items_tuples, 2))  # items in C
    current_row += len(yoy_items_tuples)

    # Blank row + Income Statement
    sections.append(('', current_row, []))  # blank
    current_row += 1
    sections.append(('Income Statement', current_row, is_items, 1))
    current_row += 1 + len(is_items)

    # Blank + Balance Sheet
    sections.append(('', current_row, []))
    current_row += 1
    sections.append(('Balance Sheet', current_row, bs_items, 1))
    current_row += 1 + len(bs_items)

    # Blank + Cash Flow
    sections.append(('', current_row, []))
    current_row += 1
    sections.append(('Cash Flow', current_row, cf_items, 1))
    current_row += 1 + len(cf_items)

    total_rows = current_row + 200  # extra buffer
    total_cols = 200  # plenty for years + quarters

    # 1. Create new sheet
    add_sheet_request = {
        'addSheet': {
            'properties': {
                'title': target_sheet_name,
                'gridProperties': {
                    'rowCount': total_rows,
                    'columnCount': total_cols,
                }
            }
        }
    }
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': [add_sheet_request]}
    ).execute()
    target_sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']
    print(f"  Created sheet '{target_sheet_name}' (id={target_sheet_id})")

    # 2. Build cell writes
    requests = []

    # Section headers (A column) and item names (B or C column)
    for section_tuple in sections:
        section_name = section_tuple[0]
        row_idx = section_tuple[1]
        items = section_tuple[2]
        item_col = section_tuple[3] if len(section_tuple) > 3 else 1

        # Section header in A
        if section_name:
            requests.append({
                'updateCells': {
                    'range': {
                        'sheetId': target_sheet_id,
                        'startRowIndex': row_idx,
                        'endRowIndex': row_idx + 1,
                        'startColumnIndex': 0,
                        'endColumnIndex': 1,
                    },
                    'rows': [{'values': [{'userEnteredValue': {'stringValue': section_name}}]}],
                    'fields': 'userEnteredValue',
                }
            })
        # Items in specified column (B=1 or C=2)
        if items:
            item_requests = []
            for i, item in enumerate(items):
                # Handle both tuple (name, text_format) and plain string
                if isinstance(item, tuple):
                    item_name, text_format = item
                else:
                    item_name = item
                    text_format = {'bold': False, 'italic': False}

                cell_value = {'userEnteredValue': {'stringValue': item_name}}
                fields = 'userEnteredValue'

                if text_format.get('bold') or text_format.get('italic'):
                    fmt = {}
                    if text_format.get('bold'):
                        fmt['bold'] = True
                    if text_format.get('italic'):
                        fmt['italic'] = True
                    cell_value['userEnteredFormat'] = {'textFormat': fmt}
                    fields += ',userEnteredFormat'

                item_requests.append({
                    'updateCells': {
                        'range': {
                            'sheetId': target_sheet_id,
                            'startRowIndex': row_idx + 1 + i,
                            'endRowIndex': row_idx + 2 + i,
                            'startColumnIndex': item_col,
                            'endColumnIndex': item_col + 1,
                        },
                        'rows': [{'values': [cell_value]}],
                        'fields': fields,
                    }
                })
            if item_requests:
                requests.extend(item_requests)

    # Batch section/item writes
    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests}
        ).execute()
        print(f"  ✓ Wrote {len(requests)} section headers and item names")

    # 3. Build year/quarter headers in row 1 (0-indexed row 0)
    header_requests = []

    # D1 = "Unit"
    header_requests.append({
        'updateCells': {
            'range': {
                'sheetId': target_sheet_id,
                'startRowIndex': 0,
                'endRowIndex': 1,
                'startColumnIndex': 3,
                'endColumnIndex': 4,
            },
            'rows': [{'values': [{'userEnteredValue': {'stringValue': 'Unit'}}]}],
            'fields': 'userEnteredValue',
        }
    })

    # E1 (col_idx=4) start from 2007 to current_year - 1
    current_year = datetime.now().year
    fiscal_year_start = 2007
    col_idx = 4  # E

    for year in range(fiscal_year_start, current_year):
        if is_natural:
            label = str(year)
        else:
            month_abbr = datetime(2024, fy_month, 1).strftime('%b')
            day_str = f'{fy_day:02d}'
            label = f'12 months {month_abbr}-{day_str}-{year}'

        header_requests.append({
            'updateCells': {
                'range': {
                    'sheetId': target_sheet_id,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': col_idx,
                    'endColumnIndex': col_idx + 1,
                },
                'rows': [{'values': [{'userEnteredValue': {'stringValue': label}}]}],
                'fields': 'userEnteredValue',
            }
        })
        col_idx += 1

    # LTM column after years (at col X = 24 for 19 years 2007-2025)
    ltm_col = col_idx
    header_requests.append({
        'updateCells': {
            'range': {
                'sheetId': target_sheet_id,
                'startRowIndex': 0,
                'endRowIndex': 1,
                'startColumnIndex': ltm_col,
                'endColumnIndex': ltm_col + 1,
            },
            'rows': [{'values': [{'userEnteredValue': {'stringValue': 'LTM'}}]}],
            'fields': 'userEnteredValue',
        }
    })
    col_idx += 1

    # Quarter headers starting at AA1 (col 26)
    aa_idx = 26
    if col_idx < aa_idx:
        col_idx = aa_idx

    now = datetime.now()
    current_q = (now.month - 1) // 3 + 1
    current_q_year = now.year

    q_year = 2021
    q_num = 1

    while (q_year, q_num) < (current_q_year, current_q):
        label = f'Q{q_num} {q_year}'
        header_requests.append({
            'updateCells': {
                'range': {
                    'sheetId': target_sheet_id,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': col_idx,
                    'endColumnIndex': col_idx + 1,
                },
                'rows': [{'values': [{'userEnteredValue': {'stringValue': label}}]}],
                'fields': 'userEnteredValue',
            }
        })
        col_idx += 1
        q_num += 1
        if q_num > 4:
            q_num = 1
            q_year += 1
    print(f"  ✓ Planned {len(header_requests)} header cells (row 1)")

    if header_requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': header_requests}
        ).execute()
        print(f"  ✓ Wrote {len(header_requests)} header cells")

    # 4. Set column widths and frozen panes
    requests = [
        # Column widths: A=20, B=80, C=80, D=40
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': target_sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 0,
                    'endIndex': 1,
                },
                'properties': {'pixelSize': 20},
                'fields': 'pixelSize',
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': target_sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 1,
                    'endIndex': 2,
                },
                'properties': {'pixelSize': 80},
                'fields': 'pixelSize',
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': target_sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 2,
                    'endIndex': 3,
                },
                'properties': {'pixelSize': 80},
                'fields': 'pixelSize',
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': target_sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 3,
                    'endIndex': 4,
                },
                'properties': {'pixelSize': 40},
                'fields': 'pixelSize',
            }
        },
        # Freeze first 4 columns and first row (headers)
        {
            'updateSheetProperties': {
                'properties': {
                    'sheetId': target_sheet_id,
                    'gridProperties': {
                        'frozenColumnCount': 4,
                        'frozenRowCount': 1,
                    },
                },
                'fields': 'gridProperties.frozenColumnCount,gridProperties.frozenRowCount',
            }
        },
    ]
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': requests}
    ).execute()

    return target_sheet_id, total_cols


def scan_item_to_row(service, spreadsheet_id, sheet_name):
    """Scan columns B and C to build item_name (lowercase) → 0-indexed row mapping.

    For each row, prefers column C (Key Stats items in new layout with sub-sections),
    falls back to column B (IS/BS/CF items and old layout).
    Skips sub-section headers like "盈利指标" and "同比增速".
    """
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!B1:C300"
    ).execute()
    rows = result.get('values', [])
    mapping = {}
    for i, row in enumerate(rows):
        b_val = row[0].strip() if row and row[0] else ''
        c_val = row[1].strip() if len(row) > 1 and row[1] else ''

        # Prefer C (Key Stats new layout), fall back to B (IS/BS/CF)
        val = c_val if c_val else b_val

        if not val:
            continue
        if val in KEY_STATS_SUB_SECTIONS:
            continue
        mapping[val.lower()] = i
    return mapping


def resolve_formula(template_str, col_letter, prev_col_letter, item_to_row):
    """Resolve formula template to actual GS formula.

    {?Item} is optional: a missing (CIQ-omitted, zero) line resolves to a bare 0,
    dropping its column prefix; a missing required item -> <col>1 + warning.
    """
    def replace_item(match):
        col = col_letter if match.group(1) == '__C__' else prev_col_letter
        item_name = match.group(2)
        # Marker prefixes: '?' optional (0 if the row is absent), '!' required but
        # N()-wrapped. Both N()-wrap when the row exists, coercing CIQ's '-' nil
        # marker (a text string) to 0 so arithmetic doesn't yield #VALUE!.
        marker = ''
        if item_name[:1] in ('?', '!'):
            marker, item_name = item_name[0], item_name[1:]
        row = item_to_row.get(item_name.lower())
        if row is None:
            if marker == '?':
                return '0'
            print(f"    WARNING: Item '{item_name}' not found, using row 1")
            return col + '1'
        if marker:
            return f'N({col}{row + 1})'
        return f'{col}{row + 1}'

    return re.sub(r'(__C__|__PC__)\{([?!]?[^}]+)\}', replace_item, template_str)


def find_data_columns(service, spreadsheet_id, sheet_name):
    """Find all data columns (columns with year headers, starting from col D)."""
    sheets = get_sheet_names(service, spreadsheet_id)
    col_count = sheets.get(sheet_name, {}).get('colCount', 200)
    end_col = col_to_letter(min(col_count - 1, 200))

    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[f"{sheet_name}!A1:{end_col}1"],
        includeGridData=True
    ).execute()

    header_vals = result.get('sheets', [{}])[0].get('data', [{}])[0].get('rowData', [{}])[0].get('values', [])

    data_cols = []
    for j in range(3, len(header_vals)):
        cell = header_vals[j]
        uev = cell.get('userEnteredValue', {})
        fv = cell.get('formattedValue', '')
        text = uev.get('stringValue', fv) or uev.get('numberValue', 0) or fv
        if text and re.search(r'\d{4}', str(text)):
            data_cols.append(j)

    return data_cols


def write_key_stats_formulas(service, spreadsheet_id, sheet_name, target_sheet_id, col_count):
    """Write Key Stats formulas by resolving item names to row numbers."""
    item_to_row = scan_item_to_row(service, spreadsheet_id, sheet_name)
    print(f"  Scanned {len(item_to_row)} items in columns B-C")

    # Find Key Stats section and build a LOCAL item→row map for just this section
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A1:C50"
    ).execute()
    rows = result.get('values', [])

    key_stats_row = None
    ks_items = {}  # local map: item_name_lower → 0-indexed row (Key Stats area only)
    next_section_headers = ('income statement', 'balance sheet', 'cash flow',
                            'key stats', 'supplemental', 'business segments')

    for i, row in enumerate(rows):
        a = row[0].strip().lower() if row and row[0] else ''
        b = row[1].strip() if len(row) > 1 and row[1] else ''
        c = row[2].strip() if len(row) > 2 and row[2] else ''

        if a == 'key stats':
            key_stats_row = i
        elif key_stats_row is not None and i > key_stats_row:
            if a in next_section_headers:
                break
            # Prefer C (new layout), fall back to B (old layout)
            item_val = c if c else b
            if item_val and item_val not in KEY_STATS_SUB_SECTIONS:
                ks_items[item_val.lower()] = i

    if key_stats_row is None:
        print("  WARNING: Key Stats section not found")
        return

    print(f"  Key Stats items found: {len(ks_items)}")

    data_cols = find_data_columns(service, spreadsheet_id, sheet_name)
    if not data_cols:
        print("  WARNING: No data columns found")
        return

    requests = []
    for item_name, template in FORMULA_TEMPLATES.items():
        item_name_lower = item_name.lower()

        # Find item row in Key Stats area (local map, not global)
        found_item_row = ks_items.get(item_name_lower)

        if found_item_row is None:
            # Fuzzy match
            for k, v in ks_items.items():
                if item_name_lower in k or k in item_name_lower:
                    found_item_row = v
                    break

        if found_item_row is None:
            continue

        # Resolve formula: item references like {Net Income} resolve to IS/BS/CF rows
        # (the global item_to_row picks up those, not Key Stats)
        for ci, data_col in enumerate(data_cols):
            col_letter = col_to_letter(data_col)
            pc = col_to_letter(max(0, data_cols[ci - 1])) if ci > 0 else col_to_letter(max(0, data_col - 1))

            formula = resolve_formula(template['formula'], col_letter, pc, item_to_row)
            cell_value = {
                'userEnteredValue': {'formulaValue': formula},
                'userEnteredFormat': {'numberFormat': dict(template['format'])},
            }
            requests.append({
                'updateCells': {
                    'range': {
                        'sheetId': target_sheet_id,
                        'startRowIndex': found_item_row,
                        'endRowIndex': found_item_row + 1,
                        'startColumnIndex': data_col,
                        'endColumnIndex': data_col + 1,
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
        print(f"  ✓ Wrote {len(requests)} Key Stats formulas")
    else:
        print("  No Key Stats formulas to write")


def classify_exchange(code):
    """Classify a stock code by exchange.

    Returns 'SSE' (上交所), 'SZSE' (深交所), 'HKEX' (港交所), or None.
    """
    s = str(code).strip()
    if not s.isdigit():
        return None
    if len(s) == 6:
        if s[:3] in ('600', '601', '603', '605', '688', '689', '900'):
            return 'SSE'
        if s[0] in ('0', '3') or s[:3] == '200':
            return 'SZSE'
    if 1 <= len(s) <= 5:
        return 'HKEX'
    return None


def copy_summary_column_formulas(service, spreadsheet_id, summary_sheet_id,
                                  src_col_0idx, dst_col_0idx):
    """Copy formulas/format from src col rows 3+ to dst col on Summary.

    Self-relative refs (e.g. F$2) auto-shift to the destination column.
    """
    src_letter = col_to_letter(src_col_0idx)
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f'Summary!{src_letter}3:{src_letter}1000'
    ).execute()
    src_values = result.get('values', [])
    while src_values and (not src_values[-1] or not str(src_values[-1][0]).strip()):
        src_values.pop()
    if not src_values:
        print(f"  Source column {src_letter} empty below row 2, skipping formula copy")
        return

    last_row_excl = 2 + len(src_values)  # 0-indexed exclusive end
    requests = [{
        'copyPaste': {
            'source': {
                'sheetId': summary_sheet_id,
                'startRowIndex': 2,
                'endRowIndex': last_row_excl,
                'startColumnIndex': src_col_0idx,
                'endColumnIndex': src_col_0idx + 1,
            },
            'destination': {
                'sheetId': summary_sheet_id,
                'startRowIndex': 2,
                'endRowIndex': last_row_excl,
                'startColumnIndex': dst_col_0idx,
                'endColumnIndex': dst_col_0idx + 1,
            },
            'pasteType': 'PASTE_NORMAL',
            'pasteOrientation': 'NORMAL',
        }
    }]
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={'requests': requests}
    ).execute()
    dst_letter = col_to_letter(dst_col_0idx)
    print(f"  ✓ Cloned formulas Summary!{src_letter} → {dst_letter} (rows 3-{last_row_excl})")


def add_to_summary(service, spreadsheet_id, code, name):
    """Add company to Summary sheet and clone formulas from a same-exchange column."""
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range='Summary!A1:AZ2'
    ).execute()
    rows = result.get('values', [])

    max_col = 0
    for row in rows:
        max_col = max(max_col, len(row))

    next_col = max_col
    if rows and code in [str(c).strip() for c in rows[0]]:
        print(f"  Code {code} already in Summary, skipping")
        return

    range_str = f'Summary!{col_to_letter(next_col)}1:{col_to_letter(next_col)}2'
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_str,
        valueInputOption='RAW',
        body={'values': [[code], [name]]}
    ).execute()
    print(f"  ✓ Added {code} ({name}) to Summary column {col_to_letter(next_col)}")

    exchange = classify_exchange(code)
    if exchange is None:
        print(f"  Exchange not identified for code {code}, skipping formula clone")
        return

    codes_row = rows[0] if rows else []
    src_col = None
    for j, c in enumerate(codes_row):
        if j == next_col:
            continue
        if classify_exchange(c) == exchange:
            src_col = j
            break

    if src_col is None:
        print(f"  No existing {exchange} column to clone from, skipping formula clone")
        return

    sheets = get_sheet_names(service, spreadsheet_id)
    summary_sheet_id = sheets.get('Summary', {}).get('sheetId')
    if summary_sheet_id is None:
        print("  WARNING: Summary sheet not found")
        return

    copy_summary_column_formulas(
        service, spreadsheet_id, summary_sheet_id, src_col, next_col
    )


def get_spreadsheet_id_for_code(code):
    """Look up spreadsheet ID from industry_spreadsheets.json."""
    json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'industry_spreadsheets.json')
    if os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        for industry, info in data.items():
            if code in info.get('codes', []):
                return info['spreadsheet_id'], industry
    return None, None


def main():
    import argparse
    parser = argparse.ArgumentParser(description='Create a new company tab from scratch')
    parser.add_argument('--code', required=True, help='Stock code (e.g. 600519)')
    parser.add_argument('--name', required=True, help='Company name (e.g. 贵州茅台)')
    parser.add_argument('--spreadsheet-id', help='Google Spreadsheet ID (auto-routed if omitted)')
    parser.add_argument('--excel', required=True, help='Path to CIQ Excel file')
    parser.add_argument('--sheet-suffix', default='财务', help='Sheet name suffix (default: 财务)')
    parser.add_argument('--dry-run', action='store_true', help='Preview actions without writing')
    args = parser.parse_args()

    service = get_service()

    # Auto-route if spreadsheet ID not provided
    spreadsheet_id = args.spreadsheet_id
    if not spreadsheet_id:
        spreadsheet_id, industry = get_spreadsheet_id_for_code(args.code)
        if spreadsheet_id:
            print(f"Auto-routed to industry: {industry}")
        else:
            print(f"WARNING: Code {args.code} not found in industry_spreadsheets.json")
            spreadsheet_id = "1huXdbAgYR2xul5CDtOmuoCjBKGwQu69XB9_AcooRPC0"

    sheet_name = f"{args.name}{args.sheet_suffix}"
    print(f"Creating tab: '{sheet_name}' in spreadsheet {spreadsheet_id[:20]}...")
    print(f"Excel source: {args.excel}")

    # Check if target sheet already exists
    sheets = get_sheet_names(service, spreadsheet_id)
    if sheet_name in sheets:
        print(f"Sheet '{sheet_name}' already exists, aborting")
        sys.exit(1)

    if args.dry_run:
        wb = xlrd.open_workbook(args.excel)
        is_items = read_excel_items(wb, 'Income Statement')
        bs_items = read_excel_items(wb, 'Balance Sheet')
        cf_items = read_excel_items(wb, 'Cash Flow')
        print(f"Excel: IS={len(is_items or [])}, BS={len(bs_items or [])}, CF={len(cf_items or [])} items (with formatting)")
        print(f"Key Stats: will copy from first available GS tab")
        print("[DRY RUN] No changes made")
        return 0, 200

    # Step 1: Build structure (sections, items, headers)
    target_sheet_id, col_count = build_sheet_structure(service, spreadsheet_id, sheet_name, args.excel)

    # Step 2: Write Key Stats formulas (incl. Payout Ratio % from FORMULA_TEMPLATES)
    write_key_stats_formulas(service, spreadsheet_id, sheet_name, target_sheet_id, col_count)

    # Step 3: Add to Summary
    add_to_summary(service, spreadsheet_id, args.code, args.name)

    print(f"\nDone! Tab '{sheet_name}' created.")
    print("Next: Run update_financials.py to fill data.")


if __name__ == '__main__':
    main()
