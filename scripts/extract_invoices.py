#!/usr/bin/env python3
"""Extract invoices from Crystal Reports XML to CSV and optionally XLSX.

Usage: python3 scripts/extract_invoices.py test.xml --csv invoices.csv --xlsx invoices.xlsx

Outputs a CSV (always) and an XLSX if `openpyxl` is installed.
"""
import re
import csv
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {'cr': 'urn:crystal-reports:schemas:report-detail'}


def get_text(elem, tag='FormattedValue'):
    if elem is None:
        return ''
    tv = elem.find(f'cr:{tag}', NS)
    return (tv.text or '').strip() if tv is not None and tv.text else ''


def build_parent_map(root):
    parent = {}
    for p in root.iter():
        for c in list(p):
            parent[c] = p
    return parent

def parse_vehicle_fields(vehicle_str):
    # Remove 'Vehicle: ' prefix if present
    s = vehicle_str.strip()
    if s.lower().startswith('vehicle:'):
        s = s[len('vehicle:'):].strip()
    # Split by commas
    parts = [p.strip() for p in s.split(',')]
    truck = parts[0] if parts else ''
    license = ''
    unit = ''
    if len(parts) >= 3:
        if parts[-1].replace(' ', '').isdigit():
            unit = parts[-1]
            license = parts[-2]
        else:
            license = parts[-1]
    elif len(parts) == 2:
        license = parts[-1]
    return truck, license, unit

# --- ADDED: parse_invhdr and find_ancestor_by_tag ---
def parse_invhdr(inv_text):
    """Parse invoice header string to extract invoice number and date."""
    # Example: 'Invoice: 12345 Date: 12/21/2025' or similar
    inv_num = ''
    date = ''
    if inv_text:
        m = re.search(r'(?:Invoice:?\s*)(\d+)', inv_text)
        if m:
            inv_num = m.group(1)
        # Try to match 'Date' or 'Posted On' for the date field
        m = re.search(r'(?:Date:?|Posted On:?)[\s,]*(\d{1,2}/\d{1,2}/\d{2,4})', inv_text)
        if m:
            date = m.group(1)
    return inv_num, date

def find_ancestor_by_tag(elem, parent_map, tag):
    """Climb the parent map to find the nearest ancestor with the given tag name."""
    p = parent_map.get(elem)
    while p is not None:
        if p.tag.endswith(tag):
            return p
        p = parent_map.get(p)
    return None

# --- ADDED: find_field_value ---
def find_field_value(group_elem, field_name):
    """Find the value of a field with the given FieldName in the group element."""
    if group_elem is None:
        return ''
    # Search for Field element with the given FieldName
    f = group_elem.find(f'.//cr:Field[@FieldName="{field_name}"]', NS)
    if f is not None:
        return get_text(f, 'FormattedValue') or get_text(f, 'Value')
    return ''

def extract(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    parent_map = build_parent_map(root)

    rows = []

    # Find all invoice header Fields
    inv_fields = root.findall('.//cr:Field[@FieldName="{@InvHdr}"]', NS)
    for f in inv_fields:
        inv_text = get_text(f, 'FormattedValue') or get_text(f, 'Value')
        invoice_num, date = parse_invhdr(inv_text)

        # find the enclosing Group for this invoice
        invoice_group = find_ancestor_by_tag(f, parent_map, 'Group')

        # find vehicle by climbing ancestors until a field @YmmEngLic is found
        vehicle = ''
        g = invoice_group
        while g is not None:
            v = g.find('.//cr:Field[@FieldName="{@YmmEngLic}"]', NS)
            if v is not None:
                vehicle = get_text(v, 'FormattedValue') or get_text(v, 'Value')
                break
            g = find_ancestor_by_tag(g, parent_map, 'Group')

        # Always fill Vehicle column from XML, then parse Truck, License, Unit from that value
        vehicle_value = vehicle
        truck, license, unit = parse_vehicle_fields(vehicle_value)

        def tofloat(s):
            try:
                return float(s.replace(',', '').strip()) if s else 0.0
            except Exception:
                return 0.0

        parts_val = find_field_value(invoice_group, '{@PartsTotal}')
        labor_val = find_field_value(invoice_group, '{@LaborTotal}')
        discount_val = find_field_value(invoice_group, '{@DiscountTotal}')
        hazmat_val = find_field_value(invoice_group, '{@HazMat}')
        supplies_val = find_field_value(invoice_group, '{@Supplies}')
        tax_val = find_field_value(invoice_group, '{@TaxTotal}')
        total_val = find_field_value(invoice_group, '{@Total}')

        rows.append({
            'Invoice': invoice_num,
            'Date': date,
            'Truck': truck,
            'License': license,
            'Unit': unit,
            'Parts': tofloat(parts_val),
            'Labor': tofloat(labor_val),
            'Discount': tofloat(discount_val),
            'Haz Mat': tofloat(hazmat_val),
            'Supplies': tofloat(supplies_val),
            'Tax': tofloat(tax_val),
            'Total': tofloat(total_val),
            'Vehicle': vehicle_value,
        })

    return rows

def _parse_date_key(datestr):
    # Expecting MM/DD/YYYY; fallback to minimal
    from datetime import datetime, date
    if not datestr:
        return date.min
    for fmt in ('%m/%d/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(datestr, fmt).date()
        except Exception:
            continue
    # try to extract numbers
    m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', datestr)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(1)), int(m.group(2))).date()
        except Exception:
            pass
    return date.min


def _parse_invoice_key(inv):
    if not inv:
        return 0
    m = re.search(r'(\d+)', inv)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return 0
    return 0


def write_csv(rows, out_csv):
    if not rows:
        print('No invoices found.')
        return
    # sort by vehicle, date (ascending), invoice number (ascending)
    rows.sort(key=lambda r: (
        (r.get('Vehicle') or '').lower(),
        _parse_date_key(r.get('Date')),
        _parse_invoice_key(r.get('Invoice')),
    ))

    headers = ['Invoice', 'Date', 'Truck', 'License', 'Unit', 'Parts', 'Labor', 'Discount', 'Haz Mat', 'Supplies', 'Tax', 'Total', 'Vehicle']
    with open(out_csv, 'w', newline='', encoding='utf-8') as fh:
        w = csv.DictWriter(fh, fieldnames=headers)
        w.writeheader()
        for r in rows:
            w.writerow(r)
    print(f'Wrote CSV: {out_csv}')


def write_xlsx(rows, out_xlsx):
    try:
        from openpyxl import Workbook
    except Exception:
        print('openpyxl not installed; skipping XLSX write. Install with: pip install openpyxl')
        return
    wb = Workbook()
    ws = wb.active
    # ensure same sorting as CSV
    rows.sort(key=lambda r: (
        (r.get('Vehicle') or '').lower(),
        _parse_date_key(r.get('Date')),
        _parse_invoice_key(r.get('Invoice')),
    ))

    headers = ['Invoice', 'Date', 'Truck', 'License', 'Unit', 'Parts', 'Labor', 'Discount', 'Haz Mat', 'Supplies', 'Tax', 'Total', 'Vehicle']
    ws.append(headers)
    for r in rows:
        ws.append([r[h] for h in headers])
    wb.save(out_xlsx)
    print(f'Wrote XLSX: {out_xlsx}')


def main():
    p = argparse.ArgumentParser()
    p.add_argument('xml', help='Crystal Reports XML file (test.xml)')
    p.add_argument('--xlsx', default='invoices.xlsx', help='Output XLSX file (optional)')
    p.add_argument('--group', action='store_true', help='Also produce grouped report by vehicle')
    p.add_argument('--group-xlsx', default='grouped_by_truck.xlsx', help='Grouped XLSX output')
    args = p.parse_args()

    rows = extract(args.xml)
    write_xlsx(rows, args.xlsx)
    if args.group:
        grouped = group_by_vehicle(rows)
        write_group_xlsx(grouped, args.group_xlsx)


def group_by_vehicle(rows):
    from collections import defaultdict
    groups = defaultdict(lambda: {
        'vehicle': '',
        'Unit': '',
        'quantity of invoices': 0,
        'Parts': 0.0,
        'Labor': 0.0,
        'Discount': 0.0,
        'Haz Mat': 0.0,
        'Supplies': 0.0,
        'Tax': 0.0,
        'Total': 0.0,
    })
    for r in rows:
        vehicle = (r.get('Vehicle') or '').strip() or 'Unknown'
        unit = (r.get('Unit') or '').strip()
        g = groups[vehicle]
        g['vehicle'] = vehicle
        g['Unit'] = unit
        g['quantity of invoices'] += 1
        g['Parts'] += float(r.get('Parts') or 0)
        g['Labor'] += float(r.get('Labor') or 0)
        g['Discount'] += float(r.get('Discount') or 0)
        g['Haz Mat'] += float(r.get('Haz Mat') or 0)
        g['Supplies'] += float(r.get('Supplies') or 0)
        g['Tax'] += float(r.get('Tax') or 0)
        g['Total'] += float(r.get('Total') or 0)
    # return as list, only the requested columns
    return [
        {
            'vehicle': g['vehicle'],
            'Unit': g['Unit'],
            'quantity of invoices': g['quantity of invoices'],
            'Parts': g['Parts'],
            'Labor': g['Labor'],
            'Discount': g['Discount'],
            'Haz Mat': g['Haz Mat'],
            'Supplies': g['Supplies'],
            'Tax': g['Tax'],
            'Total': g['Total'],
        }
        for g in groups.values()
    ]


def write_group_csv(groups, out_csv):
    headers = ['vehicle', 'Unit', 'quantity of invoices', 'Parts', 'Labor', 'Discount', 'Haz Mat', 'Supplies', 'Tax', 'Total']
    with open(out_csv, 'w', newline='', encoding='utf-8') as fh:
        w = csv.DictWriter(fh, fieldnames=headers)
        w.writeheader()
        # sort by 'Total' descending
        for g in sorted(groups, key=lambda x: float(x.get('Total') or 0), reverse=True):
            w.writerow(g)
    print(f'Wrote grouped CSV: {out_csv}')


def write_group_xlsx(groups, out_xlsx):
    try:
        from openpyxl import Workbook
    except Exception:
        print('openpyxl not installed; skipping grouped XLSX write.')
        return
    wb = Workbook()
    ws = wb.active
    ws.title = 'Grouped by Truck'
    headers = ['vehicle', 'Unit', 'quantity of invoices', 'Parts', 'Labor', 'Discount', 'Haz Mat', 'Supplies', 'Tax', 'Total']
    ws.append(headers)
    # sort by 'Total' descending
    for g in sorted(groups, key=lambda x: float(x.get('Total') or 0), reverse=True):
        ws.append([g[h] for h in headers])
    wb.save(out_xlsx)
    print(f'Wrote grouped XLSX: {out_xlsx}')


if __name__ == '__main__':
    main()
