# Photocard shop: sales and inventory report

A Python script that turns the export of my stock tracker into a sales report for my second-hand K-pop photocard shop on Vinted.

## Why this exists

Vinted shows completed sales only in a list that cannot be downloaded. To see how much I sell, which products earn the most and how much money is tied up in stock, I rebuilt the sales data from my own stock tracker (the website in `index.html`).

## How it works

- Every export of the tracker keeps a copy of the full inventory (member, category, quantity, list price of each card).
- When I sell a card I lower its quantity. A quantity drop between an earlier and a later copy is counted as a sale.
- Changes in August 2026 are treated as stock registration, not sales. Selling started on 1 September 2026.
- The script builds a sales log and the current inventory, runs data checks, and writes an Excel report with live formulas.

## Usage

```
pip install pandas openpyxl
python skz_sales_report.py tracker_export.html --out output_folder
```

Optional: `--costs cost_workbook.xlsx` reads purchase costs from my own cost workbook. Without it, the values in the Settings block at the top of the script are used.

The tracker export itself is not part of this repository, so the script cannot be run end to end without my data.

## Output

- `photocard_sales_report.xlsx`: report sheet (KPIs, sales by month, member and category, findings), sales log, inventory, stock trend, data checks and notes
- `sales.csv`, `inventory.csv`, `stock_trend.csv`: the same tables as CSV files

## Data checks

The script lists seven types of discrepancies for review, for example units sold with no price on record, listings deleted from the tracker, and sudden quantity drops of three or more units.

## Limitations

- Sales made before a listing was uploaded are not captured, so sales figures are a lower bound.
- Stock is updated in batches, so a sale shows on the processing date, not the order date.
- Revenue is estimated from list prices and an average discount (adjustable in the report), not from actual transaction prices.
- The monthly summary currently covers September and October 2026 and needs a small code change for later months.

Code written with the help of an AI coding assistant (Claude).
