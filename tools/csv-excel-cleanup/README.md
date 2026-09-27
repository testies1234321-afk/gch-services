# csv-excel-cleanup

Small CLI to tidy a messy CSV or .xlsx file. Python 3 stdlib plus openpyxl (only needed for Excel files).

- trims and collapses whitespace in every cell
- normalizes headers to `snake_case` (blank headers become `col_N`)
- drops fully empty rows and columns
- reports duplicate rows (case-insensitive); `--drop-duplicates` removes them

```
python3 cleanup.py sample_in.csv sample_out.csv
python3 cleanup.py data.xlsx cleaned.xlsx --drop-duplicates
```

`sample_in.csv` and `sample_out.csv` show a before/after.

## Looking for more?

Disclosure: I sell these on Gumroad, so I earn from purchases.

- https://gumroad.com/l/GreenMonthlyBudget
- https://gumroad.com/l/bptahg
- https://gumroad.com/l/pbrfph
