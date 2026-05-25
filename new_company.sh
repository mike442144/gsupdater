#!/usr/bin/env bash
# Organize the full flow for creating a new company tab:
# 1. Create tab structure (create_company_tab.py)
# 2. Fill financial data (update_financials.py)
# 3. Update 扣非净利润 (update_kcfjcxsyjlr.py)
#
# Usage:
#   ./new_company.sh --code 600660 --name "福耀玻璃" --excel /path/to/file.xls
#   ./new_company.sh --code 600660 --name "福耀玻璃" --excel file.xls --skip-koufei

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

CODE=""
NAME=""
EXCEL=""
SHEET_SUFFIX="财务"
SKIP_KOUFEI=false
SPREADSHEET_ID=""

usage() {
    echo "Usage: $0 --code <code> --name <name> --excel <file> [--skip-koufei] [--spreadsheet-id <id>]"
    exit 1
}

while [[ $# -gt 0 ]]; do
    case "$1" in
        --code) CODE="$2"; shift 2 ;;
        --name) NAME="$2"; shift 2 ;;
        --excel) EXCEL="$2"; shift 2 ;;
        --sheet-suffix) SHEET_SUFFIX="$2"; shift 2 ;;
        --skip-koufei) SKIP_KOUFEI=true; shift ;;
        --spreadsheet-id) SPREADSHEET_ID="$2"; shift 2 ;;
        -h|--help) usage ;;
        *) echo "Unknown arg: $1"; usage ;;
    esac
done

if [[ -z "$CODE" || -z "$NAME" || -z "$EXCEL" ]]; then
    echo "Error: --code, --name, and --excel are required"
    usage
fi

SHEET_NAME="${NAME}${SHEET_SUFFIX}"
echo "========================================"
echo "Creating new company tab"
echo "  Code: $CODE"
echo "  Name: $NAME"
echo "  Sheet: $SHEET_NAME"
echo "  Excel: $EXCEL"
echo "========================================"

# Step 1: Create tab structure with Key Stats formulas
echo ""
echo ">>> Step 1: Creating tab structure..."
python3 "$SCRIPT_DIR/create_company_tab.py" \
    --code "$CODE" \
    --name "$NAME" \
    --sheet-suffix "$SHEET_SUFFIX" \
    --excel "$EXCEL"

# Resolve spreadsheet ID if not provided
if [[ -z "$SPREADSHEET_ID" ]]; then
    SPREADSHEET_ID=$(python3 -c "
import json, sys
sys.path.insert(0, '$SCRIPT_DIR')
# extract industry from code mapping in industry_spreadsheets.json
with open('$SCRIPT_DIR/industry_spreadsheets.json') as f:
    cfg = json.load(f)
for industry, data in cfg.items():
    if '$CODE' in [str(c) for c in data.get('codes', [])]:
        print(data['spreadsheet_id'])
        sys.exit(0)
# fallback to default
print('1huXdbAgYR2xul5CDtOmuoCjBKGwQu69XB9_AcooRPC0')
")
    echo "  Resolved spreadsheet ID: $SPREADSHEET_ID"
fi

# Step 2: Fill financial data
echo ""
echo ">>> Step 2: Filling financial data..."
python3 "$SCRIPT_DIR/update_financials.py" \
    "$EXCEL" \
    "$SHEET_NAME" \
    --spreadsheet-id "$SPREADSHEET_ID"

# Step 3 (optional): Update 扣非净利润
if [[ "$SKIP_KOUFEI" == true ]]; then
    echo ""
    echo ">>> Step 3: Skipped (扣非净利润)"
else
    echo ""
    echo ">>> Step 3: Updating 扣非净利润..."
    python3 "$SCRIPT_DIR/update_kcfjcxsyjlr.py" \
        --codes "$CODE" \
        --sheet-id "$SPREADSHEET_ID"
fi

echo ""
echo "========================================"
echo "Done! Sheet: $SHEET_NAME"
echo "========================================"
