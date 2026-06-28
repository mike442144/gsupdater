#!/usr/bin/env python3
"""
Build a "<name>运营数据" segments tab for HK-listed companies in a Google
Spreadsheet, sourcing 主营构成 / 分部 data from annual-report PDFs.

Data source: HKEX annual reports downloaded by
``~/Projects/tinyant/hkexnews/index.js`` into ``data/<code>/`` (e.g.
``data/02233/2024_西部水泥_年度报告.pdf``). This script parses the segment
note ("收入及分部資料") out of those PDFs with PyMuPDF.

Mirrors add_segments_section.py (A-share, eastmoney) and reuses its Google
Sheets layout engine (plan_rows / apply_tab / apply_tab_incremental) with
HK-specific metrics. Annual data only; one annual report yields its fiscal
year plus the prior comparative year.

Layout (same hierarchy as the A-share tab; year columns mirror 财务):

    A            B            C            | 2007 ... 2024 2025
    主营构成
    营业收入
    按产品
                 銷售水泥及相關產品            |  ...   7,645.6
                 ...
    按地区
                 中國市場                     |  ...   5,184.8
                 海外市場                     |  ...   3,160.2
    分部业绩
    按地区
                 中國市場                     |  ...     348.0
                 海外市場                     |  ...     892.8

HK notes disclose segment revenue and segment profit only (no segment
assets/liabilities), so the metrics are 营业收入 (external sales) and
分部业绩 (分部溢利). Values are converted from the report's 千元 into
百万元 to match the 财务 tab's millions convention.

Usage:
    python add_hk_segments_section.py --parse data/02233/2024_*.pdf   # debug parse
    python add_hk_segments_section.py --sheet-id <id> --codes 2233 --dry-run
    python add_hk_segments_section.py --sheet-id <id> --codes 2233            # appends new years
    python add_hk_segments_section.py --sheet-id <id> --codes 2233 --rebuild
"""

import sys
import os
import re
import glob
import argparse
import subprocess

import fitz  # PyMuPDF

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# Reuse the A-share module's Sheets layout engine + helpers (parameterised
# below with HK metrics). See add_segments_section.py for the layout contract.
from add_segments_section import (
    GOOGLE_TOKEN_PATH,
    YEAR_START_COL, TAB_SUFFIX, TAB_TITLE, UNIT_HEADER, LABEL_COL_WIDTH,
    get_summary_mapping, resolve_sheet, get_fiscal_years, find_tab,
    cell, col_to_letter,
    plan_rows, apply_tab, apply_tab_incremental,
)

# ── Configuration ───────────────────────────────────────────────────────────

HKEXNEWS_DIR = os.path.expanduser('~/Projects/tinyant/hkexnews')
HKEXNEWS_SCRIPT = os.path.join(HKEXNEWS_DIR, 'index.js')
HKEXNEWS_DATA = os.path.join(HKEXNEWS_DIR, 'data')

# HK metrics: (label, csv_col_unused, number-format type, pattern).
# 4-tuple shape matches add_segments_section.METRICS so plan_rows is reusable.
HK_METRICS = [
    ('营业收入', None, 'NUMBER', '#,##0.0'),
    ('分部业绩', None, 'NUMBER', '#,##0.0'),
]
HK_METRIC_FMT = {label: (ftype, fpat) for label, _c, ftype, fpat in HK_METRICS}
HK_CLASSIFICATION_ORDER = ['按产品', '按地区']

# ── Token / value helpers ───────────────────────────────────────────────────

DASH_TOKS = {'–', '—', '-', '－'}
_NUM_RE = re.compile(r'^[\d,]+$')
_PAREN_RE = re.compile(r'^\(([\d,]+)\)$')


def parse_money(tok):
    """'7,645,607' -> 7645607; '(53,713)' -> -53713; dash/nil -> None."""
    s = (tok or '').strip()
    if s in DASH_TOKS:
        return None
    m = _PAREN_RE.match(s)
    if m:
        return -int(m.group(1).replace(',', ''))
    if _NUM_RE.match(s):
        return int(s.replace(',', ''))
    return None


def is_value_tok(tok):
    s = (tok or '').strip()
    return s in DASH_TOKS or bool(_NUM_RE.match(s)) or bool(_PAREN_RE.match(s))


_CN_DIGIT = {'零': '0', '〇': '0', '一': '1', '二': '2', '兩': '2', '两': '2',
             '三': '3', '四': '4', '五': '5', '六': '6', '七': '7',
             '八': '8', '九': '9'}


def cn_year_to_arabic(text):
    """'截至二零二四年十二月三十一日止年度' -> '2024'. Passthrough for arabic."""
    m = re.search(r'((?:[零〇一二兩两三四五六七八九]\s?){4})', text)
    if m:
        digits = re.sub(r'\s', '', m.group(1))
        return ''.join(_CN_DIGIT.get(c, c) for c in digits)
    m = re.search(r'(\d{4})', text)
    return m.group(1) if m else None


def take_nums(block, start, count):
    """Collect up to `count` consecutive value tokens from `block` at `start`."""
    vals = []
    i = start
    while i < len(block) and len(vals) < count:
        if is_value_tok(block[i]):
            vals.append(parse_money(block[i]))
            i += 1
        else:
            break
    return vals


# ── PDF segment-note parsing ────────────────────────────────────────────────

# Heading -> classification, for the "simple" two-column tables (item + two
# numbers + bare total). Covers both the product-type table (2021+) and the
# geographic-revenue table some reporters use instead (2019/2020 for 02233).
SIMPLE_HEADINGS = {
    '產品及服務種類': '按产品', '產品及服務類別': '按产品',
    '产品及服务种类': '按产品', '产品及服务类别': '按产品',
    '地區市場': '按地区', '地區資料': '按地区', '地区市场': '按地区',
}
SIMPLE_STOP_KW = ('收入確認', '客戶', '客户', '政策', '履約', '截至',
                  '經營分部', '经营分部', '分部收益', '確認', '合約',
                  '業務分部', '业务分部', '可報告', '可报告')

COL_STOP = {'總計', '总计', '調整及對銷', '调整及对销', '綜合', '综合',
            '合計', '合计', '小計', '小计', '總額', '总额'}


def gather_note_lines(doc):
    """Flat list of non-empty lines spanning the '收入及分部資料' note.

    Returns None if the note heading is absent.
    """
    start = None
    for i in range(doc.page_count):
        t = doc[i].get_text()
        if '收入及分部資料' in t or '收入及分部资料' in t:
            start = i
            break
    if start is None:
        return None
    lines = []
    # The note spans a handful of pages; over-gather then let table parsers
    # stop at their own end signals.
    for i in range(start, min(start + 8, doc.page_count)):
        for ln in doc[i].get_text().split('\n'):
            s = ln.strip()
            if s:
                lines.append(s)
    return lines


def detect_unit_factor(lines):
    """千元 -> 0.001 (to millions); 百萬元 -> 1.0. Default 0.001 for HK."""
    text = '\n'.join(lines)
    if '千元' in text or '千元' in text:
        return 0.001
    if '百萬元' in text or '百万元' in text:
        return 1.0
    return 0.001


def parse_simple_table(lines, hidx, current_year, prior_year, factor, add):
    """Parse an item + (current, prior) two-number table starting at hidx.

    `add(classification, item, year, metric, raw_value)` records a value
    (raw in 千元; add() applies the factor). Bare trailing total row is
    excluded by taking only the first two numbers per item.
    """
    cls = SIMPLE_HEADINGS.get(lines[hidx]) or SIMPLE_HEADINGS.get(
        lines[hidx].replace(' ', ''))
    if cls is None:
        return
    n = len(lines)
    i = hidx + 1
    cur_item = None
    nums = []
    while i < n:
        tok = lines[i]
        if is_value_tok(tok):
            nums.append(parse_money(tok))
        else:
            stop = (any(k in tok for k in SIMPLE_STOP_KW)
                    or tok in SIMPLE_HEADINGS
                    or re.match(r'^\d+\.\s*$', tok)
                    or len(tok) > 12)
            if stop:
                break
            if cur_item is not None and nums:
                add(cls, cur_item, current_year, '营业收入', nums[0])
                if len(nums) > 1:
                    add(cls, cur_item, prior_year, '营业收入', nums[1])
            cur_item = tok
            nums = []
        i += 1
    if cur_item is not None and nums:
        add(cls, cur_item, current_year, '营业收入', nums[0])
        if len(nums) > 1:
            add(cls, cur_item, prior_year, '营业收入', nums[1])


def parse_segment_tables(lines, factor, add):
    """Parse the 經營分部 / 分部收入及業績 multi-column tables.

    Each report has up to two sub-tables (current FY + prior comparative),
    each introduced by a '截至...止年度' marker. Page headers repeat that
    marker text, so a marker qualifies only if its block (up to the next
    marker) actually contains the 外部銷售 and 分部溢利 rows.
    """
    # Region starts at the segment-results heading.
    region_start = None
    for i, ln in enumerate(lines):
        if '分部收入及業績' in ln or '分部收入及业绩' in ln:
            region_start = i
            break
    if region_start is None:
        return
    region_end = len(lines)

    markers = [i for i in range(region_start, region_end)
               if lines[i].startswith('截至') and '止年度' in lines[i]]

    for mi, m in enumerate(markers):
        block_end = markers[mi + 1] if mi + 1 < len(markers) else region_end
        block = lines[m:block_end]
        if ('外部銷售' not in block and '外部销售' not in block):
            continue
        if ('分部溢利' not in block and '分部業績' not in block
                and '分部业绩' not in block):
            continue

        fy = cn_year_to_arabic(lines[m])
        if not fy:
            continue

        # Column headers: tokens after the marker up to the units / group label.
        headers = []
        h = 1
        while h < len(block):
            t = block[h]
            if (t.startswith('截至') or t == '分部收益'
                    or t.startswith('人民幣') or t.startswith('人民币')
                    or is_value_tok(t)):
                break
            headers.append(t)
            h += 1
        seg_names = []
        for x in headers:
            if x in COL_STOP:
                break
            seg_names.append(x)
        if not seg_names:
            continue
        n_total = len(headers)

        rev_row = None
        profit_row = None
        k = 0
        while k < len(block):
            t = block[k]
            if t in ('外部銷售', '外部销售'):
                rev_row = take_nums(block, k + 1, n_total)
                k += 1 + n_total
                continue
            if t in ('分部溢利', '分部業績', '分部业绩'):
                profit_row = take_nums(block, k + 1, n_total)
                k += 1 + n_total
                continue
            k += 1

        for idx, seg in enumerate(seg_names):
            if rev_row and idx < len(rev_row):
                add('按地区', seg, fy, '营业收入', rev_row[idx])
            if profit_row and idx < len(profit_row):
                add('按地区', seg, fy, '分部业绩', profit_row[idx])


def extract_segments(pdf_path, report_year):
    """Parse one annual report PDF into the nested segments structure.

    Returns {classification: {'items': [...], 'data': {item: {year: {metric: val}}}}}
    with values converted to 百万元.
    """
    doc = fitz.open(pdf_path)
    lines = gather_note_lines(doc)
    doc.close()
    if not lines:
        return {}

    factor = detect_unit_factor(lines)
    segments = {}

    def add(cls, item, year, metric, raw):
        if raw is None or not item:
            return
        val = round(raw * factor, 1)
        block = segments.setdefault(cls, {'items': [], 'data': {}})
        yd = block['data'].setdefault(item, {}).setdefault(str(year), {})
        yd[metric] = val
        if item not in block['items']:
            block['items'].append(item)

    for idx, ln in enumerate(lines):
        if ln in SIMPLE_HEADINGS or ln.replace(' ', '') in SIMPLE_HEADINGS:
            parse_simple_table(lines, idx, report_year, report_year - 1, factor, add)

    parse_segment_tables(lines, factor, add)
    return segments


def build_segments_for_code(code_dir, code, years_wanted):
    """Merge segment data across all available annual-report PDFs for a code.

    Each report covers its FY and the prior comparative; later reports
    override earlier ones for overlapping years (re-stated comparatives).
    Returns the nested segments structure.
    """
    segments = {}

    def merge(src):
        # Newest report processed first (see reverse=True below); keep the
        # first-seen value per (item, year, metric) so a newer report's
        # restated comparative wins over the older report's original.
        for cls, block in src.items():
            dst = segments.setdefault(cls, {'items': [], 'data': {}})
            for item in block['items']:
                if item not in dst['items']:
                    dst['items'].append(item)
            for item, ydata in block['data'].items():
                ditem = dst['data'].setdefault(item, {})
                for yr, metrics in ydata.items():
                    dyr = ditem.setdefault(yr, {})
                    for metric, val in metrics.items():
                        if metric not in dyr:
                            dyr[metric] = val

    # Process newest-first so a newer report's restated comparative wins.
    pdfs = sorted(glob.glob(os.path.join(code_dir, '*_年度报告.pdf')), reverse=True)
    for pdf in pdfs:
        m = re.search(r'(\d{4})_', os.path.basename(pdf))
        if not m:
            continue
        report_year = int(m.group(1))
        try:
            parsed = extract_segments(pdf, report_year)
        except Exception as e:  # noqa: BLE001 - one bad PDF shouldn't abort the run
            print(f"  WARN: failed to parse {os.path.basename(pdf)}: {e}")
            continue
        if parsed:
            print(f"  {os.path.basename(pdf)}: "
                  + ', '.join(f"{c}:{len(b['items'])}" for c, b in parsed.items()))
            merge(parsed)
    return segments


# ── HKEX report download ────────────────────────────────────────────────────

def pad_code(code):
    """Summary stores '2233'; hkexnews + data dir use 5-digit '02233'."""
    return str(code).strip().zfill(5)


def find_code_dir(code):
    for d in (os.path.join(HKEXNEWS_DATA, pad_code(code)),
              os.path.join(HKEXNEWS_DATA, str(code).strip())):
        if os.path.isdir(d):
            return d
    return None


def ensure_reports(code, years, fetch=True):
    """Make sure annual-report PDFs for `years` exist; download missing ones.

    Returns the code's data directory (creating it via hkexnews if needed).
    """
    code_dir = find_code_dir(code)
    have = set()
    if code_dir:
        for pdf in glob.glob(os.path.join(code_dir, '*_年度报告.pdf')):
            m = re.search(r'(\d{4})_', os.path.basename(pdf))
            if m:
                have.add(int(m.group(1)))
    missing = sorted(y for y in years if y not in have)
    if not missing or not fetch:
        if missing and not fetch:
            print(f"  --no-fetch: missing years {missing} not downloaded")
        return code_dir

    if not os.path.isfile(HKEXNEWS_SCRIPT):
        print(f"  WARN: hkexnews script not found at {HKEXNEWS_SCRIPT}; "
              f"cannot fetch missing years {missing}")
        return code_dir

    yr_range = f"{min(missing)}-{max(missing)}"
    cmd = ['node', os.path.basename(HKEXNEWS_SCRIPT),
           '--codes', pad_code(code), '--annual', '--year', yr_range, '--pdf-only']
    print(f"  Fetching missing annual reports ({yr_range}) via hkexnews...")
    result = subprocess.run(cmd, cwd=HKEXNEWS_DIR, capture_output=True,
                            text=True, timeout=600)
    if result.returncode != 0:
        print(f"  WARN: hkexnews failed: {result.stderr[:300]}")
    return find_code_dir(code)


# ── Google Sheets driver ────────────────────────────────────────────────────

def get_service():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    return build('sheets', 'v4', credentials=creds)


def process_code_hk(service, spreadsheet_id, code, fin_sheet_name,
                    dry_run, rebuild=False, fetch=True):
    print(f"\n{'='*60}\n{code}\n{'='*60}")

    fin_name, fin_id, _rc, fin_cols = resolve_sheet(
        service, spreadsheet_id, fin_sheet_name)
    if fin_id is None:
        print(f"  ERROR: 财务 tab '{fin_sheet_name}' not found; skipping.")
        return
    years = get_fiscal_years(service, spreadsheet_id, fin_name, fin_cols)
    if not years:
        print("  ERROR: no year columns found on 财务 tab; skipping.")
        return

    code_dir = ensure_reports(code, [int(y) for y in years], fetch=fetch)
    if not code_dir:
        print(f"  ERROR: no annual-report PDFs for {code} "
              f"(looked in {HKEXNEWS_DATA}/{pad_code(code)}); skipping.")
        return

    segments = build_segments_for_code(code_dir, code, [int(y) for y in years])
    if not segments:
        print("  No segment data parsed from annual reports; skipping.")
        return

    base = fin_name[:-2] if fin_name.endswith('财务') else fin_name
    tab_name = base + TAB_SUFFIX
    print(f"  Mirroring years {years[0]}-{years[-1]} ({len(years)} cols) "
          f"from '{fin_name}' -> '{tab_name}'")
    for cls in HK_CLASSIFICATION_ORDER:
        if cls in segments:
            shown = [it for it in segments[cls]['items']
                     if any(y in segments[cls]['data'][it] for y in years)]
            print(f"    {cls}: {len(shown)} items ({', '.join(shown)})")

    sheet_id, _rc, _cc = find_tab(service, spreadsheet_id, tab_name)
    if sheet_id is not None and not rebuild:
        apply_tab_incremental(service, spreadsheet_id, tab_name, sheet_id,
                              _rc, _cc, segments, years, dry_run,
                              metric_fmt=HK_METRIC_FMT)
    else:
        if rebuild and sheet_id is not None:
            print(f"  --rebuild: regenerating '{tab_name}' from scratch")
        apply_tab(service, spreadsheet_id, tab_name, segments, years, dry_run,
                  metrics=HK_METRICS, classification_order=HK_CLASSIFICATION_ORDER,
                  skip_empty_metric_blocks=True)


# ── CLI ─────────────────────────────────────────────────────────────────────

def is_hk(code):
    """HKEX code: numeric, 1-5 digits (excludes 6-digit A-shares and US tickers)."""
    s = str(code).strip()
    return s.isdigit() and 1 <= len(s) <= 5


def main():
    p = argparse.ArgumentParser(description='Build HK 分部 运营数据 tab from annual reports')
    p.add_argument('--sheet-id', help='Target Google Spreadsheet ID')
    p.add_argument('--codes', help='Comma-separated HK codes (default: all HK in Summary)')
    p.add_argument('--dry-run', action='store_true')
    p.add_argument('--rebuild', action='store_true')
    p.add_argument('--no-fetch', action='store_true', help="Don't download missing PDFs")
    p.add_argument('--parse', metavar='PDF', help='Debug: parse one PDF and print structure')
    args = p.parse_args()

    if args.parse:
        m = re.search(r'(\d{4})_', os.path.basename(args.parse))
        yr = int(m.group(1)) if m else 0
        seg = extract_segments(args.parse, yr)
        print(f"Parsed {args.parse} (report year {yr}):\n")
        for cls in HK_CLASSIFICATION_ORDER:
            if cls not in seg:
                continue
            print(f"[{cls}]")
            for item in seg[cls]['items']:
                yd = seg[cls]['data'][item]
                yrs = ', '.join(f"{y}={yd[y]}" for y in sorted(yd))
                print(f"  {item}: {yrs}")
        return

    if not args.sheet_id:
        p.error('--sheet-id is required (or use --parse <pdf> for a debug parse)')

    service = get_service()
    mapping = get_summary_mapping(service, args.sheet_id)
    codes = ([c.strip() for c in args.codes.split(',') if c.strip()]
             if args.codes else list(mapping.keys()))
    hk = [c for c in codes if is_hk(c)]
    skipped = [c for c in codes if not is_hk(c)]
    if skipped:
        print(f"Skipping non-HK codes (A-share/US, unsupported): {skipped}")
    if not hk:
        print("No HK codes to process.")
        return

    for code in hk:
        if code not in mapping:
            print(f"\nSKIP {code}: not found in Summary")
            continue
        process_code_hk(service, args.sheet_id, code, mapping[code],
                        args.dry_run, args.rebuild, fetch=not args.no_fetch)


if __name__ == '__main__':
    main()
