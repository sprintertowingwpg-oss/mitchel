#!/usr/bin/env python3
"""
Generate all reports and charts from a Crystal Reports XML input.

Usage:
    python3 scripts/generate_all_reports.py test.xml

This will produce:
    - invoices.csv, invoices.xlsx
    - grouped_by_truck.csv, grouped_by_truck.xlsx
    - grouped_totals.png (bar chart)
    - labor_vs_parts_pie.png (pie chart)
"""

import sys
import subprocess
from pathlib import Path
import shutil
import csv
import re

def run(cmd, desc=None):
    print(f"\n[Running] {desc or ' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(result.stdout)
        print(result.stderr)
        raise SystemExit(f"Failed: {' '.join(cmd)}")
    print(result.stdout)

def get_last_invoice_date(xlsx_path):
    try:
        from openpyxl import load_workbook
    except ImportError:
        print("openpyxl is required to read XLSX files.")
        return None
    wb = load_workbook(xlsx_path, read_only=True)
    ws = wb.active
    date_col = None
    for idx, cell in enumerate(next(ws.iter_rows(min_row=1, max_row=1))):
        if str(cell.value).strip().lower() == 'date':
            date_col = idx
            break
    if date_col is None:
        return None
    dates = []
    date_re = re.compile(r'\d{1,2}/\d{1,2}/\d{2,4}')
    for row in ws.iter_rows(min_row=2):
        v = row[date_col].value
        if v and isinstance(v, str) and date_re.match(v.strip()):
            dates.append(v.strip())
    if not dates:
        return None
    # Parse dates and return the latest
    from datetime import datetime
    dt_objs = []
    for d in dates:
        for fmt in ('%m/%d/%Y', '%Y-%m-%d'):
            try:
                dt_objs.append(datetime.strptime(d, fmt))
                break
            except Exception:
                continue
    if not dt_objs:
        return None
    return max(dt_objs).strftime('%Y-%m-%d')

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 scripts/generate_all_reports.py <input.xml>")
        sys.exit(1)
    xml = sys.argv[1]
    base = Path(xml).stem


    # 1. Extract invoices and grouped report (no CSVs)
    # Find last date first (will be used for output folder)
    temp_xlsx = 'invoices.xlsx'
    temp_group_xlsx = 'grouped_by_truck.xlsx'
    run([
        sys.executable, 'scripts/extract_invoices.py', xml,
        '--group',
        '--xlsx', temp_xlsx,
        '--group-xlsx', temp_group_xlsx,
    ], desc='Extracting invoices and grouped report (XLSX only)')

    # 2. Find last date in invoices.xlsx
    last_date = get_last_invoice_date(temp_xlsx)
    if not last_date:
        last_date = 'unknown_date'
    outdir = Path(last_date)
    outdir.mkdir(exist_ok=True)

    # 3. Move XLSX files to output folder (before plotting)
    out_invoices_xlsx = outdir / 'invoices.xlsx'
    out_grouped_xlsx = outdir / 'grouped_by_truck.xlsx'
    shutil.move(temp_xlsx, out_invoices_xlsx)
    shutil.move(temp_group_xlsx, out_grouped_xlsx)

    # 4. Generate bar chart (total per vehicle)
    run([
        sys.executable, 'scripts/plot_grouped.py', str(out_grouped_xlsx), '--out', str(outdir / 'grouped_totals.png')
    ], desc='Generating bar chart (grouped_totals.png)')

    # 5. Generate pie chart (labor vs parts)
    run([
        sys.executable, 'scripts/plot_pie_labor_parts.py', str(out_grouped_xlsx), '--out', str(outdir / 'labor_vs_parts_pie.png')
    ], desc='Generating pie chart (labor_vs_parts_pie.png)')

    print(f"\nAll reports and charts generated in folder: {outdir}")

if __name__ == '__main__':
    main()