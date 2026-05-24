# financials-updater

Tools for syncing financial data into Google Sheets from two sources.

## Scripts

### `update_financials.py` — CIQ Excel → Google Sheets

Reads Capital IQ Excel files (`.xls`/`.xlsx`) and writes Income Statement, Balance Sheet, and Cash Flow data into the corresponding Google Sheets. Matches rows by item name (column B), not row number, so sheet structure changes are handled safely.

Also writes Payout Ratio formulas (`DPS / Basic EPS`) and copies Key Stats formulas to new columns automatically.

**Single file:**
```bash
python update_financials.py path/to/CIQ_file.xls "公司财务"
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

Paths hardcoded in the scripts:

| Setting | Value |
|---|---|
| Google token | `/home/mike/.hermes/google_token.json` |
| Eastmoney script | `/home/mike/projects/tinyant/eastmoney/index.js` |
| Default spreadsheet | `1huXdbAgYR2xul5CDtOmuoCjBKGwQu69XB9_AcooRPC0` |

Override the spreadsheet with `--sheet-id` (扣非) or `--spreadsheets` (batch financials).
