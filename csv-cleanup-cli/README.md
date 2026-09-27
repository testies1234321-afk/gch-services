# csv-cleanup-cli

Free, single-file Python tool for CSV cleanup: tidy messy spreadsheet data and export a cleaned CSV and an Excel (.xlsx) sheet.

What it does: strips whitespace, drops fully empty rows and columns, removes duplicate rows, normalises headers to snake_case, converts obvious numeric and date columns (dates become ISO `YYYY-MM-DD`), and reports every change it made.

## Install

```
pip install openpyxl
curl -O https://raw.githubusercontent.com/testies1234321-afk/gch-services/main/csv-cleanup-cli/csv_cleanup.py
```

## Usage

```
python3 csv_cleanup.py messy.csv -o cleaned
```

Real output from a run on a small messy file (5 data rows, with a blank row, a blank column, a duplicate, padded headers and `1,200.50` / `03/15/2026` values):

```
read 5 data row(s) from messy.csv
 - dropped 1 fully empty row(s)
 - dropped 1 fully empty column(s)
 - removed 1 duplicate row(s)
 - renamed header 'Customer Name' -> 'customer_name'
 - renamed header 'Order Date' -> 'order_date'
 - renamed header 'Amount ($)' -> 'amount'
 - renamed header 'Notes' -> 'notes'
 - column 'order_date': converted to dates (ISO)
 - column 'amount': converted to numbers
wrote cleaned.csv and cleaned.xlsx (3 rows, 4 columns)
```

Result (`cleaned.csv`):

```
customer_name,order_date,amount,notes
Alice,2026-03-15,1200.5,ok
Bob,2026-03-16,300,
Carol,2026-03-17,45,late
```

Note: slash dates are parsed trying day-first then month-first formats in a fixed order, and a column is only converted if every non-empty value parses. Check ambiguous dates yourself.

## Disclosure: the author also sells two paid products

The tool above is free and complete. Separately, the same author sells two optional paid products on Gumroad. You do not need them to use this tool.

- CSV & Excel Data Cleanup Cheat Sheet (7 USD): https://gchservices.gumroad.com/l/bptahg
- Excel Operations Dashboard Template (12 USD): https://gchservices.gumroad.com/l/pbrfph

## Licence

MIT License

Copyright (c) 2026 testies1234321-afk

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
