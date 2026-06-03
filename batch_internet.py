#!/usr/bin/env python3
"""
Batch add YoY section + ROIC (资产法, Greenblatt) to all internet company tabs.

Usage:
    python batch_internet.py --dry-run   # preview
    python batch_internet.py             # apply
"""

import sys, os, re, json, time

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')

# ── Config ──────────────────────────────────────────────────────
SS_ID = "12PscMVZfOsaDXhgWtoRc-UPUgGOWZSxCL72a3ALZ7pg"  # SAAS

SKIP_SHEETS = {'Summary', '资本结构'}

SECTION_HEADERS = {'income statement', 'balance sheet', 'cash flow',
                   'key stats', 'supplemental', 'business segments'}
SUB_SECTIONS = ('盈利指标', '同比增速')

# ── YoY ─────────────────────────────────────────────────────────
YOY_ITEMS = [
    ("Revenue YoY", "=IFERROR(__C__{Total Revenue}/__PC__{Total Revenue}-1,)"),
    ("Gross Profit YoY", "=IFERROR(__C__{Gross Profit}/__PC__{Gross Profit}-1,)"),
    ("Operating Income YoY", "=IFERROR(__C__{Operating Income}/__PC__{Operating Income}-1,)"),
    ("Net Income YoY", "=IFERROR(__C__{Net Income to Company}/__PC__{Net Income to Company}-1,)"),
    ("扣非净利润 YoY", "=IFERROR(__C__{扣非净利润}/__PC__{扣非净利润}-1,)"),
]

# ── ROIC new methods ────────────────────────────────────────────
NEW_ROIC_LABEL = "ROIC (资本来源法)"

NEW_ITEMS = [
    ("ROIC (资产法)",
     "=__C__{EBIT}*(1-__C__{!Effective Tax Rate %})/"
     "((__C__{Total Assets}-__C__{?Total Cash & ST Investments}-"
     "(__C__{Total Current Liabilities}-__C__{?Short-term Borrowings}-__C__{?Curr. Port. of Leases}))+"
     "(__PC__{Total Assets}-__PC__{?Total Cash & ST Investments}-"
     "(__PC__{Total Current Liabilities}-__PC__{?Short-term Borrowings}-__PC__{?Curr. Port. of Leases})))*2"),
    ("ROIC (Greenblatt)",
     "=__C__{EBIT}/"
     "((__PC__{Total Current Assets}-__PC__{?Total Cash & ST Investments})-"
     "(__PC__{Total Current Liabilities}-__PC__{?Short-term Borrowings}-__PC__{?Curr. Port. of Leases})+"
     "__PC__{Net Property, Plant & Equipment})"),
]
PERCENT_FORMAT = {"type": "PERCENT", "pattern": "0.0%"}

CAPITAL_SOURCE_ITEM = (
    NEW_ROIC_LABEL,
    "=__C__{EBIT}*(1-__C__{!Effective Tax Rate %})/"
    "(__C__{Net Debt}+__C__{Common Equity}+__C__{!Minority Interest}+"
    "__PC__{Net Debt}+__PC__{Common Equity}+__PC__{!Minority Interest})*2")
ALL_ITEMS = [CAPITAL_SOURCE_ITEM] + NEW_ITEMS
BASE_ROIC_NAMES = {"roic", NEW_ROIC_LABEL.lower()}

# ── Helpers ─────────────────────────────────────────────────────

def col_to_letter(col_idx):
    col_idx += 1
    result = ''
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_service():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    return build('sheets', 'v4', credentials=creds)


def get_company_sheets(service):
    result = service.spreadsheets().get(
        spreadsheetId=SS_ID, fields='sheets(properties(title,sheetId))'
    ).execute()
    sheets = []
    for s in result.get('sheets', []):
        p = s['properties']
        title = p['title']
        if title in SKIP_SHEETS or '资本结构' in title:
            continue
        sheets.append((title, p['sheetId']))
    return sheets


def find_data_columns(service, sheet_name):
    result = service.spreadsheets().get(
        spreadsheetId=SS_ID, ranges=[f"'{sheet_name}'!A1:CV1"],
        includeGridData=True,
        fields='sheets.data.rowData.values(userEnteredValue,formattedValue)'
    ).execute()
    vals = (result.get('sheets', [{}])[0].get('data', [{}])[0]
            .get('rowData', [{}])[0].get('values', []))
    cols = []
    for j in range(3, len(vals)):
        uev = vals[j].get('userEnteredValue', {})
        text = uev.get('stringValue', '') or vals[j].get('formattedValue', '')
        if text and re.search(r'\d{4}', str(text)):
            cols.append(j)
    return cols


def scan_item_to_row(service, sheet_name):
    """Scan B+C to build item_name → row mapping. C takes priority."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SS_ID, range=f"'{sheet_name}'!B1:C300"
    ).execute()
    rows = result.get('values', [])
    mapping = {}
    for i, row in enumerate(rows):
        b_val = row[0].strip() if row and row[0] else ''
        c_val = row[1].strip() if len(row) > 1 and row[1] else ''
        val = c_val if c_val else b_val
        if val and val not in SUB_SECTIONS:
            mapping[val.lower()] = i
    return mapping


def find_key_stats(service, sheet_name):
    """Find Key Stats section. Returns (ks_start, items, rows)."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SS_ID, range=f"'{sheet_name}'!A1:C100"
    ).execute()
    rows = result.get('values', [])
    ks_start = None
    items = []
    for i, row in enumerate(rows):
        a = row[0].strip().lower() if row and row[0] else ''
        b = row[1].strip() if len(row) > 1 and row[1] else ''
        c = row[2].strip() if len(row) > 2 and row[2] else ''
        if a == 'key stats':
            ks_start = i
        elif ks_start is not None and i > ks_start:
            if a:
                break
            item_val = c if c else b
            if item_val and item_val not in SUB_SECTIONS:
                items.append((i, item_val))
    return ks_start, items, rows


def resolve_formula(template_str, col_letter, prev_col_letter, item_to_row):
    formula = template_str.replace('__C__', col_letter).replace('__PC__', prev_col_letter)
    def replace_item(match):
        item_name = match.group(1)
        row = item_to_row.get(item_name.lower())
        if row is None:
            return '1'
        return str(row + 1)
    formula = re.sub(r'\{([^}]+)\}', replace_item, formula)
    return formula


def resolve_roic_formula(template, col_letter, prev_col_letter, item_to_row, sheet_name):
    """Resolve __C__{Item} / __PC__{Item} with ? and ! markers."""
    def repl(m):
        col = col_letter if m.group(1) == '__C__' else prev_col_letter
        name = m.group(2)
        marker = ''
        if name[:1] in ('?', '!'):
            marker, name = name[0], name[1:]
        row = item_to_row.get(name.lower())
        if row is None:
            if marker == '?':
                return '0'
            return col + '1'
        return f'N({col}{row + 1})'
    return re.sub(r'(__C__|__PC__)\{([?!]?[^}]+)\}', repl, template)


def scan_tab_for_roic(service, sheet_name):
    """Return (item_to_row, roic_row, total_rev_row, label_col, roic_label)."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SS_ID, range=f"'{sheet_name}'!A1:C300"
    ).execute()
    rows = result.get('values', [])
    item_to_row = {}
    roic_row = None
    roic_label = None
    total_rev_row = None
    label_col = 2
    tr_label_col = 2
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
        item_to_row[item_val.lower()] = i
        if item_val.lower() in BASE_ROIC_NAMES:
            roic_row = i
            roic_label = item_val
            label_col = 2 if c else 1
        elif item_val.lower() == 'total revenue':
            total_rev_row = i
            tr_label_col = 2 if c else 1
    if roic_row is None:
        label_col = tr_label_col
    return item_to_row, roic_row, roic_label, total_rev_row, label_col


# ── YoY section ─────────────────────────────────────────────────

def add_yoy_section(service, sheet_name, sheet_id, dry_run):
    ks_start, items, rows = find_key_stats(service, sheet_name)
    if ks_start is None:
        print(f"  {sheet_name}: SKIP — Key Stats not found")
        return 'skip'

    # Idempotency
    for _, row in enumerate(rows[ks_start:]):
        b_val = row[1].strip() if len(row) > 1 and row[1] else ''
        if b_val in SUB_SECTIONS:
            print(f"  {sheet_name}: YoY already present, skipping")
            return 'skip'

    if not items:
        print(f"  {sheet_name}: SKIP — no Key Stats items")
        return 'skip'

    N = len(items)
    data_cols = find_data_columns(service, sheet_name)
    if not data_cols:
        print(f"  {sheet_name}: SKIP — no data columns")
        return 'skip'

    if dry_run:
        print(f"  {sheet_name}: would insert 盈利指标 + 同比增速 + {len(YOY_ITEMS)} YoY items")
        return 'update'

    # Step 2: Insert 1 row after Key Stats header
    service.spreadsheets().batchUpdate(
        spreadsheetId=SS_ID, body={'requests': [{
            'insertDimension': {
                'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                          'startIndex': ks_start + 1, 'endIndex': ks_start + 2},
                'inheritFromBefore': False,
            }
        }]}).execute()

    # Step 3: Write "盈利指标" sub-header
    sub_header_row = ks_start + 1
    service.spreadsheets().batchUpdate(
        spreadsheetId=SS_ID, body={'requests': [{
            'updateCells': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': sub_header_row, 'endRowIndex': sub_header_row + 1,
                          'startColumnIndex': 1, 'endColumnIndex': 2},
                'rows': [{'values': [{
                    'userEnteredValue': {'stringValue': '盈利指标'},
                    'userEnteredFormat': {'textFormat': {'bold': True}}
                }]}],
                'fields': 'userEnteredValue,userEnteredFormat',
            }
        }]}).execute()

    # Step 4: Copy items B→C, clear B
    reqs = []
    for idx in range(N):
        row_0idx = ks_start + 2 + idx
        reqs.append({
            'updateCells': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': row_0idx, 'endRowIndex': row_0idx + 1,
                          'startColumnIndex': 1, 'endColumnIndex': 2},
                'rows': [{'values': [{}]}],
                'fields': 'userEnteredValue',
            }
        })
    for idx, (_, item_name) in enumerate(items):
        row_0idx = ks_start + 2 + idx
        reqs.append({
            'updateCells': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': row_0idx, 'endRowIndex': row_0idx + 1,
                          'startColumnIndex': 2, 'endColumnIndex': 3},
                'rows': [{'values': [{'userEnteredValue': {'stringValue': item_name}}]}],
                'fields': 'userEnteredValue',
            }
        })
    service.spreadsheets().batchUpdate(
        spreadsheetId=SS_ID, body={'requests': reqs}).execute()

    # Step 5: Insert YoY rows
    yoy_insert_at = ks_start + 2 + N
    yoy_insert_count = 1 + 1 + len(YOY_ITEMS)
    service.spreadsheets().batchUpdate(
        spreadsheetId=SS_ID, body={'requests': [{
            'insertDimension': {
                'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                          'startIndex': yoy_insert_at, 'endIndex': yoy_insert_at + yoy_insert_count},
                'inheritFromBefore': False,
            }
        }]}).execute()

    # Write YoY section content
    yoy_reqs = []
    for row_0idx in range(yoy_insert_at, yoy_insert_at + yoy_insert_count):
        yoy_reqs.append({
            'updateCells': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': row_0idx, 'endRowIndex': row_0idx + 1,
                          'startColumnIndex': 1, 'endColumnIndex': 3},
                'rows': [{'values': [{}, {}]}],
                'fields': 'userEnteredValue',
            }
        })
    yoy_sub_row = yoy_insert_at + 1
    yoy_reqs.append({
        'updateCells': {
            'range': {'sheetId': sheet_id,
                      'startRowIndex': yoy_sub_row, 'endRowIndex': yoy_sub_row + 1,
                      'startColumnIndex': 1, 'endColumnIndex': 2},
            'rows': [{'values': [{
                'userEnteredValue': {'stringValue': '同比增速'},
                'userEnteredFormat': {'textFormat': {'bold': True}}
            }]}],
            'fields': 'userEnteredValue,userEnteredFormat',
        }
    })
    for idx, (item_name, _) in enumerate(YOY_ITEMS):
        row_0idx = yoy_sub_row + 1 + idx
        yoy_reqs.append({
            'updateCells': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': row_0idx, 'endRowIndex': row_0idx + 1,
                          'startColumnIndex': 2, 'endColumnIndex': 3},
                'rows': [{'values': [{'userEnteredValue': {'stringValue': item_name}}]}],
                'fields': 'userEnteredValue',
            }
        })
    service.spreadsheets().batchUpdate(
        spreadsheetId=SS_ID, body={'requests': yoy_reqs}).execute()

    # Write YoY formulas
    item_to_row = scan_item_to_row(service, sheet_name)
    formula_reqs = []
    for item_name, template in YOY_ITEMS:
        item_row = item_to_row.get(item_name.lower())
        if item_row is None:
            print(f"  {sheet_name}: WARNING — YoY item '{item_name}' not found")
            continue
        for ci, dc in enumerate(data_cols):
            cl = col_to_letter(dc)
            pc = col_to_letter(data_cols[ci - 1]) if ci > 0 else col_to_letter(max(0, dc - 1))
            formula = resolve_formula(template, cl, pc, item_to_row)
            formula_reqs.append({
                'updateCells': {
                    'range': {'sheetId': sheet_id,
                              'startRowIndex': item_row, 'endRowIndex': item_row + 1,
                              'startColumnIndex': dc, 'endColumnIndex': dc + 1},
                    'rows': [{'values': [{
                        'userEnteredValue': {'formulaValue': formula},
                        'userEnteredFormat': {'numberFormat': dict(PERCENT_FORMAT)},
                    }]}],
                    'fields': 'userEnteredValue,userEnteredFormat',
                }
            })
    if formula_reqs:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SS_ID, body={'requests': formula_reqs}).execute()

    print(f"  {sheet_name}: ✓ YoY section added ({len(formula_reqs)} formulas)")
    return 'update'


# ── ROIC methods ────────────────────────────────────────────────

def add_roic_methods(service, sheet_name, sheet_id, dry_run):
    item_to_row, roic_row, roic_label, total_rev_row, label_col = scan_tab_for_roic(service, sheet_name)

    if roic_row is not None:
        items, anchor_row, do_relabel = NEW_ITEMS, roic_row, True
    elif total_rev_row is not None:
        items, anchor_row, do_relabel = ALL_ITEMS, total_rev_row, False
    else:
        print(f"  {sheet_name}: SKIP — no ROIC or Total Revenue row")
        return 'skip'

    n = len(items)
    data_cols = find_data_columns(service, sheet_name)
    if not data_cols:
        print(f"  {sheet_name}: SKIP — no data columns")
        return 'skip'

    existing = []
    ks = scan_item_to_row(service, sheet_name)
    for label, _ in items:
        r = ks.get(label.lower())
        existing.append(r)

    already_present = all(r is not None for r in existing)

    if dry_run:
        verb = 'rewrite' if already_present else f'insert {n} rows'
        relabel = ' + relabel→资本来源法' if do_relabel and roic_label != NEW_ROIC_LABEL else ''
        print(f"  {sheet_name}: ROIC would {verb}{relabel}")
        return 'update'

    requests = []
    relabel = do_relabel and roic_label != NEW_ROIC_LABEL

    if relabel:
        requests.append({
            'updateCells': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': roic_row, 'endRowIndex': roic_row + 1,
                          'startColumnIndex': label_col, 'endColumnIndex': label_col + 1},
                'rows': [{'values': [{'userEnteredValue': {'stringValue': NEW_ROIC_LABEL}}]}],
                'fields': 'userEnteredValue',
            }
        })

    if already_present:
        target_rows = existing
        shifted = item_to_row
    else:
        insert_at = anchor_row + 1
        target_rows = [insert_at + k for k in range(n)]
        shifted = {name: (r + n if r >= insert_at else r) for name, r in item_to_row.items()}
        requests.append({
            'insertDimension': {
                'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                          'startIndex': insert_at, 'endIndex': insert_at + n},
                'inheritFromBefore': True,
            }
        })

    for k, (label, template) in enumerate(items):
        new_row = target_rows[k]
        requests.append({
            'updateCells': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': new_row, 'endRowIndex': new_row + 1,
                          'startColumnIndex': label_col, 'endColumnIndex': label_col + 1},
                'rows': [{'values': [{'userEnteredValue': {'stringValue': label}}]}],
                'fields': 'userEnteredValue',
            }
        })
        for ci, dc in enumerate(data_cols):
            cl = col_to_letter(dc)
            pc = col_to_letter(data_cols[ci - 1]) if ci > 0 else col_to_letter(max(0, dc - 1))
            formula = resolve_roic_formula(template, cl, pc, shifted, sheet_name)
            requests.append({
                'updateCells': {
                    'range': {'sheetId': sheet_id,
                              'startRowIndex': new_row, 'endRowIndex': new_row + 1,
                              'startColumnIndex': dc, 'endColumnIndex': dc + 1},
                    'rows': [{'values': [{
                        'userEnteredValue': {'formulaValue': formula},
                        'userEnteredFormat': {'numberFormat': dict(PERCENT_FORMAT)},
                    }]}],
                    'fields': 'userEnteredValue,userEnteredFormat',
                }
            })

    service.spreadsheets().batchUpdate(
        spreadsheetId=SS_ID, body={'requests': requests}).execute()

    relabel_note = f' + relabel→"{NEW_ROIC_LABEL}"' if relabel else ''
    print(f"  {sheet_name}: ✓ ROIC done ({len(data_cols)} formulas each){relabel_note}")
    return 'update'


# ── Main ────────────────────────────────────────────────────────

def main():
    dry_run = '--dry-run' in sys.argv

    service = get_service()
    company_sheets = get_company_sheets(service)
    print(f"Found {len(company_sheets)} company tabs\n")

    yoy_u = yoy_s = 0
    roic_u = roic_s = 0

    for i, (name, sid) in enumerate(company_sheets):
        if i > 0:
            time.sleep(1)

        print(f"--- {name} ---")
        r = add_yoy_section(service, name, sid, dry_run)
        if r == 'update': yoy_u += 1
        else: yoy_s += 1

        time.sleep(1)
        r = add_roic_methods(service, name, sid, dry_run)
        if r == 'update': roic_u += 1
        else: roic_s += 1

        print()

    tag = "[DRY RUN] " if dry_run else ""
    print(f"\n{tag}YoY: Updated={yoy_u}, Skipped={yoy_s}")
    print(f"{tag}ROIC: Updated={roic_u}, Skipped={roic_s}")


if __name__ == '__main__':
    main()
