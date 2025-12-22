# Invoice extractor

This repository contains a script to extract invoice summaries and generate all reports and charts from a Crystal Reports XML export (example: `test.xml`).

## Quick Start

Install required dependencies:

```bash
pip install openpyxl matplotlib
```

To generate all reports and charts from your XML file, run:

```bash
python3 scripts/generate_all_reports.py test.xml
```

This will create a folder named with the last date found in your invoice data (e.g., `2025-12-19`) and place all output files there:

- `invoices.xlsx` (all invoices)
- `grouped_by_truck.xlsx` (grouped summary)
- `grouped_totals.png` (bar chart: total per vehicle)
- `labor_vs_parts_pie.png` (pie chart: labor vs parts)

```

Both plotting scripts accept optional `--customer` and `--date-range` arguments to add extra info to the chart subtitle.

**Note:** The scripts expect the following columns in `grouped_by_truck.csv`:  
`vehicle`, `Unit`, `quantity of invoices`, `Parts`, `Labor`, `Discount`, `Haz Mat`, `Supplies`, `Tax`, `Total`
