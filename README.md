# financials-updater

Tools for syncing financial data into Google Sheets from two sources.

## Scripts

### `new_company.sh` — Full flow for creating a new company tab

Orchestrates the complete workflow to create a new company tab from scratch:
1. Creates tab structure with `create_company_tab.py` (headers, items, Key Stats formulas with YoY sub-groups)
2. Fills financial data with `update_financials.py` (IS, BS, CF values + Capital Structure sync)
3. Updates 扣非净利润 with `update_kcfjcxsyjlr.py` (A-share only)

```bash
./new_company.sh --code 600660 --name "福耀玻璃" --excel /path/to/CIQ_file.xls
```

Options:
- `--spreadsheet-id <ID>` — target a specific spreadsheet (auto-resolved from `industry_spreadsheets.json` if omitted)
- `--skip-koufei` — skip step 3 (for non-A-share companies like HK/US listings)
- `--sheet-suffix <suffix>` — customize the tab name suffix (default: `财务`)

### `update_financials.py` — CIQ Excel → Google Sheets

Reads Capital IQ Excel files (`.xls`/`.xlsx`) and writes Income Statement, Balance Sheet, and Cash Flow data into the corresponding Google Sheets. Matches rows by item name (column B), not row number, so sheet structure changes are handled safely.

Also writes Payout Ratio formulas (`DPS / Basic EPS`), copies Key Stats formulas to new columns automatically, and syncs Capital Structure Details from Excel to the Google Sheets 资本结构 tab.

**Single file:**
```bash
python update_financials.py path/to/CIQ_file.xls "公司财务"
python update_financials.py path/to/CIQ_file.xls "公司财务" --spreadsheet-id <ID>  # target specific spreadsheet
```

**Batch mode** (routes files by stock code using `industry_spreadsheets.json`):
```bash
python update_financials.py --batch ./CIQ_Financials
```

**Dry run:**
```bash
python update_financials.py --dry-run --batch ./CIQ_Financials
```

### `update_kcfjcxsyjlr.py` — Eastmoney → Google Sheets (扣非净利润)

Fetches 扣非净利润 (non-recurring net profit) for A-share companies from the eastmoney API and fills:
- Annual values in year columns (2024, 2025, ...)
- Quarterly values in quarter columns (Q1 2024, Q2 2024, ...)
- LTM formula (sum of latest 4 quarters)

```bash
python update_kcfjcxsyjlr.py                     # all A-shares from Summary
python update_kcfjcxsyjlr.py --codes 600519,000568  # specific codes
python update_kcfjcxsyjlr.py --dry-run           # preview only
python update_kcfjcxsyjlr.py --sheet-id <ID>     # target a different spreadsheet
```

### `add_segments_section.py` — Eastmoney → Google Sheets (主营构成 segments)

Builds a dedicated **`<name>运营数据`** tab per A-share company holding 主营构成 (main-business composition) data, fetched from the eastmoney API via `mainop.js`. Layout is **metric → classification → item**:
- Column A: metric (营业收入, 收入占比, 毛利率, 营业成本, 营业利润)
- Column B: classification (按行业 / 按产品 / 按地区)
- Column C: item (e.g. 茅台酒, 国内); annual values land on the item row
- Column D: `Unit` header (mirrors the 财务 tab); year columns E onward mirror the company's 财务 tab

Annual data only (`REPORT_DATE == YYYY-12-31`). The tab is created if missing and fully rebuilt on re-run (idempotent). Items with no data inside the mirrored year range (legacy labels) are dropped. A-share codes only — HK/US codes are skipped (`mainop.js` builds `.SZ/.SH` SECUCODEs). Amounts format as `#,##0`, ratios as `0.0%`; columns A–D are 80px wide with the header row and A–D frozen.

> **Tip:** Year columns mirror the 财务 tab, so for a newly released year, update the 财务 tab first, then re-run this. Re-running wipes and rebuilds the whole tab, so don't keep manual edits here.

```bash
python add_segments_section.py --sheet-id <ID>                             # all A-shares from Summary
python add_segments_section.py --sheet-id <ID> --codes 600519,000568       # specific codes
python add_segments_section.py --sheet-id <ID> --codes 600519 --dry-run    # preview only
```

`--codes` defaults to every A-share in the spreadsheet's `Summary` tab — the authoritative company list. Prefer the no-`--codes` form for rollouts; `industry_spreadsheets.json` can lag behind Summary and miss companies.

### `add_hk_segments_section.py` — HK annual reports → Google Sheets (分部 segments)

The HK counterpart of `add_segments_section.py` for A-shares. HK listings are skipped by the A-share tool (eastmoney `mainop.js` builds `.SZ/.SH` codes only), so this one sources 主营构成 / 分部 data from **annual-report PDFs** downloaded by [`~/Projects/tinyant/hkexnews/index.js`](../tinyant/hkexnews) into `data/<code>/` (e.g. `data/02233/2024_西部水泥_年度报告.pdf`).

Parses the "收入及分部資料" segment note out of each PDF with **PyMuPDF** (rule-based — no LLM), handling two layouts the reports use:
- a simple item + two-year revenue table (按產品/服務種類, or 地區市場 when the reporter has no operating segments), and
- the 經營分部 multi-column table (external sales + 分部溢利), which can span pages with repeating page-header markers.

Each annual report yields its fiscal year plus the prior comparative; reports are merged newest-first so a newer report's restated comparative wins. HK notes disclose **segment revenue and segment profit only** (no segment assets/liabilities), so the metrics are 营业收入 (external sales) and 分部业绩 (分部溢利). Values are converted from the report's 千元 into **百万元** to match the 财务 tab's millions (e.g. 02233's 2024 segment revenue 8,344,946 千元 → 8,344.9, tying to the 财务 tab's Total Revenue 8,345).

Same `<name>运营数据` tab schema as the A-share tool (metric → classification → item; year columns mirror 财务); the layout engine (`plan_rows` / `apply_tab`) is shared via the A-share module with HK-specific metrics. Item names stay in the report's original (traditional) Chinese.

```bash
python add_hk_segments_section.py --sheet-id <ID> --codes 2233 --dry-run    # preview parsed data
python add_hk_segments_section.py --sheet-id <ID> --codes 2233              # appends new years
python add_hk_segments_section.py --sheet-id <ID> --codes 2233 --rebuild    # wipe + regenerate
python add_hk_segments_section.py --parse path/to/2024_*.pdf                # debug: parse one PDF
```

`--codes` defaults to every HK code in `Summary`; pass `--no-fetch` to read only PDFs already on disk (otherwise missing years are fetched via hkexnews). Anchor keywords may need tuning per company — reports whose segment note uses a different heading/table shape (e.g. 02233's 2016–2018 reports, which have no segment-style note) yield no data for those years and are skipped.

## How it works

1. `industry_spreadsheets.json` defines industry → spreadsheet + stock codes mapping
2. Each spreadsheet has a `Summary` sheet (row 1: codes, row 2: company names) used for routing
3. Excel filenames are parsed for stock codes (e.g. `SHSE 600519` → `600519`), then routed to the correct spreadsheet and sheet
4. Columns are matched by year via a two-pointer algorithm — existing columns get updated, empty columns get backfilled, no insert/delete dimensions

## Dependencies

| Package | Used by |
|---|---|
| `google-api-python-client` | both |
| `google-auth` / `google-oauth` | both |
| `xlrd` | `update_financials.py` (read `.xls`) |
| PyMuPDF (`fitz`) | `add_hk_segments_section.py` (parse annual-report PDFs) |
| Node.js | `update_kcfjcxsyjlr.py`, `add_segments_section.py` (eastmoney scripts); `add_hk_segments_section.py` (hkexnews) |

## Configuration

Paths used by the scripts (via `os.path.expanduser('~')`):

| Setting | Value |
|---|---|
| Google token | `~/.hermes/google_token.json` |
| Eastmoney script | `~/projects/tinyant/eastmoney/index.js` |
| HKEX annual-report downloader | `~/Projects/tinyant/hkexnews/index.js` (PDFs in `…/data/<code>/`) |

Spreadsheet ID is auto-resolved from `industry_spreadsheets.json` by stock code.
Override with `--spreadsheet-id` (financials), `--sheet-id` (扣非), or `--spreadsheets` (batch financials).

### `create_company_tab.py` — Create new company tab from scratch

Builds a new company tab in Google Sheets from CIQ Excel files — no template copy needed:
1. Fills section headers (Key Stats with YoY sub-groups, IS, BS, CF) and item names in column B
2. Creates year headers (2007 to current year), LTM, and quarterly columns
3. Writes Key Stats formulas by resolving item names to row numbers
4. Adds company to the Summary sheet (cloning formulas/format from a same-exchange column)

```bash
python create_company_tab.py --code 600660 --name "福耀玻璃" --excel /path/to/CIQ_file.xls
```

**Recommended:** use `./new_company.sh` to run all steps in one command (see above).

**Manual flow** (run each step individually):
```bash
# Step 1: Create tab with structure and formulas
python create_company_tab.py --code 600660 --name "福耀玻璃" --excel file.xls

# Step 2: Fill financial data
python update_financials.py file.xls "福耀玻璃财务" --spreadsheet-id <ID>

# Step 3 (optional): Update 扣非净利润
python update_kcfjcxsyjlr.py --codes 600660 --sheet-id <ID>
```

**Formula templates** in `create_company_tab.py` include three parallel ROIC rows, a Payout Ratio % row, ROE, margins, coverage ratios, FCFF, and Basic EPS. The three ROIC methods are:

| Row | Approach | Formula |
|---|---|---|
| `ROIC (资本来源法)` | Capital-source / financing side, two-period average | `EBIT × (1 − Tax) / (Net Debt + Common Equity + Minority Interest + prior-year values) × 2` |
| `ROIC (资产法)` | Operating-asset side: total assets − excess cash − non-interest-bearing current liabilities, two-period average | `EBIT × (1 − Tax) / (avg invested capital from asset side)` |
| `ROIC (Greenblatt)` | Greenblatt/McKinsey tangible capital, pre-tax, **beginning-of-period** base (no averaging) | `EBIT / (operating working capital + net fixed assets)`, denominator uses prior column |

The `Payout Ratio %` row sits just below the ROIC block:

```
Payout Ratio % = IFERROR( N(Dividends per Share) / Basic EPS , )
```
`N()` coerces a no-dividend `'-'` to `0` (→ 0%); `IFERROR` blanks the cell when Basic EPS is zero/negative.

Formula templates support item markers to handle CIQ data quirks:
- `{?Item}` — optional: a CIQ-omitted (zero) line resolves to a bare `0`
- `{!Item}` — required but `N()`-wrapped, coercing CIQ's `'-'` text nil marker to `0` so arithmetic doesn't yield `#VALUE!`

The Key Stats EPS row is `Basic EPS` (formerly `Diluted EPS Excl. Extra Items`).

### `add_yoy_section.py` — Add YoY sub-groups to Key Stats

Restructures existing company tabs to add 盈利指标 / 同比增速 sub-groups under Key Stats. For existing tabs, it inserts the sub-header row and shifts items accordingly.

```bash
python add_yoy_section.py --spreadsheet-id <ID> --sheet "公司财务"
python add_yoy_section.py --spreadsheet-id <ID> --sheet "公司财务" --dry-run  # preview only
```

### `fix_summary_formulas.py` — Fix Summary INDIRECT row refs after a row insertion

When a row is inserted in the company tabs, items at/below it shift down and Summary INDIRECT formulas referencing those rows become stale. This script increments the row refs (rows >= `--after-row`) in Summary to match. `--after-row` is the insertion position; it defaults to `3` (the YoY `盈利指标` sub-header from `add_yoy_section.py`). Pass `--after-row 6` for the Net Income to Company row.

```bash
python fix_summary_formulas.py --spreadsheet-id <ID>                  # YoY (row 3)
python fix_summary_formulas.py --spreadsheet-id <ID> --after-row 6    # Net Income to Company
python fix_summary_formulas.py --spreadsheet-id <ID> --dry-run        # preview only
```

### `add_roic_methods.py` — Add the three ROIC methods to existing tabs

Retrofits existing company tabs with the three ROIC rows. If a tab already has a base ROIC row, it relabels it to `ROIC (资本来源法)` and inserts the 资产法 & Greenblatt rows after it. If a tab has no base ROIC, it anchors on the Key Stats `Total Revenue` row and inserts all three rows. Idempotent — re-running rewrites in place rather than inserting duplicates.

```bash
python add_roic_methods.py                          # all industries
python add_roic_methods.py --spreadsheet-id <ID>    # single spreadsheet
python add_roic_methods.py --dry-run                # preview only
```

Tabs are skipped when they have neither a base ROIC nor a `Total Revenue` row, or lack the EBIT / current-asset lines the formulas need (e.g. bank/auto-financing tabs like `易鑫财务`).

### `add_payout_ratio.py` — Add the Payout Ratio % row to existing tabs

Retrofits existing company tabs with a `Payout Ratio %` row (`IFERROR(N(Dividends per Share)/Basic EPS,)`), inserted right below the last ROIC row — the same audited-safe zone as the ROIC rollout, so no Summary fix is needed. Idempotent: a tab that already has the row gets its formula rewritten in place. The label lands in column B or C automatically, matching the anchor row's layout. Tabs without a ROIC / Total Revenue anchor are skipped.

```bash
python add_payout_ratio.py                          # all industries
python add_payout_ratio.py --spreadsheet-id <ID>    # single spreadsheet
python add_payout_ratio.py --dry-run                # preview only
```

### `wrap_keystats_refs.py` — Wrap Key Stats refs to an item in `N()`

Finds Key Stats formulas that reference a named item's row(s) and wraps those refs in `N()`, coercing CIQ's `'-'` text nil marker to `0`. Only touches the Key Stats section.

```bash
python wrap_keystats_refs.py --item "Minority Interest"                    # all industries
python wrap_keystats_refs.py --item "Effective Tax Rate %" --spreadsheet-id <ID>
python wrap_keystats_refs.py --item "Minority Interest" --dry-run
```

### `audit_roic_rollout.py` — Safety audit before a Key Stats rollout

Reports, per spreadsheet, the max `Summary` INDIRECT row reference vs the ROIC row(s) in each tab. Inserting rows after ROIC is only safe when `maxSummaryRef <= roic_row`, since Summary references company-tab rows by fixed number (text refs are NOT auto-adjusted on row insert). Also lists tabs missing a base ROIC.

```bash
python audit_roic_rollout.py
```

### `rename_eps.py` — Rename EPS row to Basic EPS on existing tabs

Renames the Key Stats item `Diluted EPS Excl. Extra Items` → `Basic EPS` and updates all formulas (Key Stats + Summary INDIRECT) to reference the Basic EPS row in the Income Statement.

```bash
python rename_eps.py                          # all industries
python rename_eps.py --spreadsheet-id <ID>    # single spreadsheet
python rename_eps.py --dry-run                # preview only
```

### `run_rollout.py` — Driver for the ROIC rollout

Runs `add_roic_methods.py`, then `wrap_keystats_refs.py` for `Minority Interest` and `Effective Tax Rate %`, over the audit-cleared "safe" industry list. Industries with divergent Summary layouts / no usable ROIC anchor (地产开发, 金融, 租赁物业, 教育) are excluded.

```bash
python run_rollout.py            # all safe industries
python run_rollout.py 互联网 食品  # specific industries
```

### `gs_rankings.py` — Fetch all ranking data from GS in one pass (recommended)

Reads each company tab **once** and extracts **six dimensions**: EV/EBIT, ROIC, Profit Quality (扣非净利润 / 净利润), FCF Ratio (自由现金流 / 公司净利润), Capex Ratio (Σ|资本开支| / Σ经营活动现金流, 近3年), and 5-year aggregate Payout Ratio (`Σ(DPS) / Σ(EPS)`). Outputs two CSVs that feed into `combined_ranking.py`.

```bash
python gs_rankings.py                                       # all rollout industries
python gs_rankings.py 互联网 食品                            # specific industries
python gs_rankings.py --rankings r.csv --payout p.csv       # custom output paths
```

### `combined_ranking.py` — Master ranking

Merges `rankings.csv` and `payout_rankings.csv` into one master ranking. The combined score is the **sum of selected ranks** (lower = better). Use `--preset` to choose which dimensions to include:

| Preset | Dimensions |
|---|---|
| `full` (default) | EV/EBIT + ROIC + Profit Quality + FCF Ratio + Capex Ratio + Payout (all 6) |
| `classic` | EV/EBIT + ROIC + Profit Quality + Payout (original 4) |
| `fcf4` | EV/EBIT + ROIC + FCF Ratio + Payout |
| `fcf5` | EV/EBIT + ROIC + FCF Ratio + Capex Ratio + Payout |

```bash
python combined_ranking.py                              # default: all 6 dims
python combined_ranking.py --preset classic              # original 4 dims
python combined_ranking.py --preset fcf4                 # no profit quality
python combined_ranking.py --preset fcf5                 # FCF + Capex, no profit quality
python combined_ranking.py --output fcf4.csv --preset fcf4  # custom output
```

Each dimension is ranked independently:
- EV/EBIT: low → high (cheaper = better)
- ROIC: high → low (more efficient = better)
- Profit Quality (扣非净利润/净利润): high → low (more stable earnings = better)
- FCF Ratio (自由现金流/公司净利润): high → low (better cash conversion = better)
- Capex Ratio (Σ|资本开支|/Σ经营活动现金流, 近3年): low → high (lower capex intensity = better)
- Payout Ratio (5-yr aggregate): high → low (more generous = better)

