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
| Node.js | `update_kcfjcxsyjlr.py` (calls eastmoney script) |

## Configuration

Paths used by the scripts (via `os.path.expanduser('~')`):

| Setting | Value |
|---|---|
| Google token | `~/.hermes/google_token.json` |
| Eastmoney script | `~/projects/tinyant/eastmoney/index.js` |
| Default spreadsheet | `1huXdbAgYR2xul5CDtOmuoCjBKGwQu69XB9_AcooRPC0` |

Override the spreadsheet with `--sheet-id` (扣非) or `--spreadsheets` (batch financials).

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

**Formula templates** in `create_company_tab.py` include ROIC (with tax rate adjustment and two-year average capital base), ROE, margins, coverage ratios, and FCFF. ROIC formula follows the pattern:
```
ROIC = EBIT × (1 - Tax Rate) / (Net Debt + Common Equity + Minority Interest + Previous Year values) × 2
```

### `add_yoy_section.py` — Add YoY sub-groups to Key Stats

Restructures existing company tabs to add 盈利指标 / 同比增速 sub-groups under Key Stats. For existing tabs, it inserts the sub-header row and shifts items accordingly.

```bash
python add_yoy_section.py --spreadsheet-id <ID> --sheet "公司财务"
python add_yoy_section.py --spreadsheet-id <ID> --sheet "公司财务" --dry-run  # preview only
```

### `fix_summary_formulas.py` — Fix Summary INDIRECT row refs after YoY insertion

When `add_yoy_section.py` inserts a sub-header row, Key Stats items shift down and Summary INDIRECT formulas referencing those rows become stale. This script increments the row refs (rows >= 3) in Summary to match.

```bash
python fix_summary_formulas.py --spreadsheet-id <ID>
python fix_summary_formulas.py --spreadsheet-id <ID> --dry-run  # preview only
```
