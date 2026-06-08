### `add_net_income_to_company.py` — Add Net Income to Company row to Key Stats

Adds a "Net Income to Company" row to each company tab's Key Stats section, right below the "Net Income" row.

    Net Income to Company = Net Income - Minority Interest

Minority Interest is N()-wrapped so CIQ's '-' nil marker (zero-value) coerces to 0. IFERROR blanks the cell when the calculation yields an error.

Anchor = Net Income row. Since this inserts between Net Income and Gross Margin, all Key Stats items after Net Income shift down by 1 row. This affects Summary INDIRECT formulas that reference those rows — they must be fixed in a separate pass.

```bash
python add_net_income_to_company.py --spreadsheet-id <ID>    # single spreadsheet
python add_net_income_to_company.py --spreadsheet-id <ID> --dry-run  # preview
python add_net_income_to_company.py                          # all industries
python add_net_income_to_company.py --dry-run                # preview all
```

### `fix_summary_net_income_refs.py` — Fix Summary after Net Income to Company insertion

When `add_net_income_to_company.py` inserts the "Net Income to Company" row after "Net Income", all Key Stats items that were originally after Net Income shift down by 1 row. Summary INDIRECT formulas still reference old row numbers and need +1.

Also updates the "自由现金流/净利润" ratio to use "Net Income to Company" as denominator (label → "自由现金流/公司净利润").

```bash
python fix_summary_net_income_refs.py --spreadsheet-id <ID>
python fix_summary_net_income_refs.py --spreadsheet-id <ID> --dry-run
```
