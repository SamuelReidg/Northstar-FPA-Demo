#!/usr/bin/env python3
"""Emit raw_data.js — every row of the workbook's `Raw Data` sheet, for the Excel export.

FP&A (2026-09-14): the Export to Excel workbook should carry a Raw Data tab with all rows.
The exporter builds its sheets from the rendered DOM, and a year of timesheet rows cannot live inside
index.html, so the rows ship in this separate file which the page loads ONLY when the export
runs (lazy <script> injection — see the RAW DATA block at the end of index.html).

Streams the sheet XML row by row (reusing regenerate_pmo_reduced.Workbook, stdlib only) and
writes JSON incrementally, so memory stays flat — calamine's to_python on this sheet OOMs.

Re-run after every weekly refresh, alongside regenerate_pmo_reduced.py.

    python3 regenerate_raw_data.py --workbook "Utilization Report.xlsx" --out raw_data.js
"""
import argparse, datetime as dt, json, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from regenerate_pmo_reduced import Workbook, col_num, serial_to_date   # noqa: E402

# Date columns are matched by HEADER, not by letter: both of these arrive as Excel serials
# and would otherwise land in the workbook as bare numbers (46039 instead of 2026-01-17).
DATE_HEADERS = {'entry date', 'week end', 'month'}

# FP&A picked the reporting columns (2026-09-14); the sheet's other ~18 columns are lookup/flag
# helpers (Billing Type*, Exclusions, Location Conversion, Document #, Short Term PTO …) and
# tripled the workbook's size for no reporting value.  Set KEEP_COLUMNS = None for all columns.
KEEP_COLUMNS = ['Legacy ID', 'Contact name', 'Location name', 'Entry date', 'Project ID',
                'Utilized hours', 'Non-utilized hours', 'Hours', 'Utilization Rate',
                'Time type', 'Task name', 'Week End']


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--workbook', required=True)
    ap.add_argument('--out', default='raw_data.js')
    ap.add_argument('--sheet', default='Raw Data')
    a = ap.parse_args()

    bk = Workbook(a.workbook)
    if a.sheet not in bk.sheets:
        sys.exit(f'ERROR: sheet {a.sheet!r} not found: {sorted(bk.sheets)[:20]}')

    it = bk.rows(a.sheet)
    try:
        _, first = next(it)
    except StopIteration:
        sys.exit('ERROR: sheet is empty')
    letters = sorted(first, key=col_num)
    if KEEP_COLUMNS:
        want = {str(h).strip(): i for i, h in enumerate(KEEP_COLUMNS)}
        letters = sorted([c for c in letters if str(first[c]).strip() in want],
                         key=lambda c: want[str(first[c]).strip()])
        missing = set(want) - {str(first[c]).strip() for c in letters}
        if missing:
            print('WARNING: columns not found on the sheet:', sorted(missing))
    header = [first[c] for c in letters]
    width = len(letters)
    date_letters = {c for c in letters if str(first[c]).strip().lower() in DATE_HEADERS}

    n = 0
    with open(a.out, 'w', encoding='utf-8') as f:
        f.write('/* Raw Data sheet for the Excel export — regenerated weekly by '
                'regenerate_raw_data.py. Loaded lazily by index.html when Export runs. */\n')
        f.write('window.RAW_DATA_SHEET={"generated":' + json.dumps(dt.datetime.now().isoformat(timespec='seconds'))
                + ',"source_sheet":' + json.dumps(a.sheet)
                + ',"columns":' + json.dumps(header, separators=(',', ':')) + ',"rows":[')
        for _, d in it:
            if not d:
                continue
            row = []
            for c in letters:
                v = d.get(c, '')
                if v == '' or v is None:
                    row.append('')
                elif c in date_letters and isinstance(v, float):
                    row.append(serial_to_date(v).isoformat())
                elif isinstance(v, float) and v.is_integer():
                    row.append(int(v))
                else:
                    row.append(v)
            f.write(('' if n == 0 else ',') + json.dumps(row, separators=(',', ':')))
            n += 1
            if n % 25000 == 0:
                print(f'  … {n:,} rows', flush=True)
        f.write(']};\n')
    size = os.path.getsize(a.out)
    print(f'{a.out}: {n:,} rows x {width} cols, {size/1048576:.1f} MB')
    print('columns:', header)


if __name__ == '__main__':
    main()
