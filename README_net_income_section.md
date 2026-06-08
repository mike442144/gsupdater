### `add_net_income_to_company.py` — Add Net Income to Company row to Key Stats

Adds a "Net Income to Company" row to each company tab's Key Stats section, right below the "Net Income" row.

The Key Stats row pulls the value straight from the Income Statement's existing "Net Income to Company" item (same column):

    =IFERROR(<col><Net Income to Company row>,)

IFERROR blanks the cell when that item can't be resolved. The row is added to every company tab (with or without Minority Interest) so the Summary page can reference it uniformly.

Anchor = Net Income row. Since this inserts between Net Income and Gross Margin, all Key Stats items after Net Income shift down by 1 row. This affects Summary INDIRECT formulas that reference those rows — they must be fixed in a separate pass.

```bash
python add_net_income_to_company.py --spreadsheet-id <ID>    # single spreadsheet
python add_net_income_to_company.py --spreadsheet-id <ID> --dry-run  # preview
python add_net_income_to_company.py                          # all industries
python add_net_income_to_company.py --dry-run                # preview all
```

### `fix_summary_net_income_refs.py` — Retarget the Summary FCF ratio to Net Income to Company

Renames the Summary "自由现金流/净利润" ratio to "自由现金流/公司净利润" and repoints its denominator from Net Income (company-tab row N) to Net Income to Company (row N+1). It does **not** shift row refs — the `+1` shift for the inserted row is `fix_summary_formulas.py --after-row 6` (run it first).

```bash
python fix_summary_net_income_refs.py --spreadsheet-id <ID>
python fix_summary_net_income_refs.py --spreadsheet-id <ID> --dry-run
```

### Run order

1. `python add_net_income_to_company.py --spreadsheet-id <ID>` — insert the row in each company tab.
2. `python fix_summary_formulas.py --spreadsheet-id <ID> --after-row 6` — shift Summary INDIRECT refs (rows >= 6) for the inserted row.
3. `python fix_summary_net_income_refs.py --spreadsheet-id <ID>` — rename the ratio label and retarget its denominator.
