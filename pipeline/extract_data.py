"""Extract data from the Utilization Report Excel into the site's JSON files.

Reads SOURCE_XLSX and writes:
  - site_data/<Slug>.json   one per sheet (tab name slugified)
  - site_data/_manifest.json
  - employees.json          (special-shape file for Employee Utilization tab)
  - site_data/On_Call_Trend.json   (copy of PT_Hours.json — same data, different tab nav)

The output is identical in shape to what was already on disk so build_site.py
can render against it unchanged.
"""

import json, os, sys, datetime as _dt, re
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))


def report_date():
    """The reporting 'today'. FPA_AS_OF=YYYY-MM-DD pins it (reproducible builds of a past
    package); without it this is the real current date, exactly as before."""
    v = os.environ.get('FPA_AS_OF', '').strip()
    if v:
        return _dt.date.fromisoformat(v)
    return _dt.date.today()


def slugify(name: str) -> str:
    """Mirror the slug pattern used by the existing JSON filenames."""
    s = re.sub(r'[^\w\s]', ' ', name)        # punctuation (incl. parens, dashes) → space
    s = re.sub(r'\s+', '_', s.strip())       # collapse whitespace
    s = re.sub(r'_+', '_', s)                # collapse multiple underscores
    return s.strip('_')


def cell_to_json(v):
    """Convert an openpyxl cell value into a JSON-safe form."""
    if v is None or isinstance(v, (str, int, float, bool)):
        return v
    if isinstance(v, (_dt.datetime, _dt.date)):
        # Use the same `{'__date__': iso}` marker the rest of the code expects.
        if isinstance(v, _dt.date) and not isinstance(v, _dt.datetime):
            v = _dt.datetime(v.year, v.month, v.day)
        return {'__date__': v.isoformat()}
    # Fallback — string repr
    return str(v)


def extract_sheet(ws):
    """Return list of rows; each row is a list of JSON-safe cell values."""
    out = []
    for row in ws.iter_rows(values_only=True):
        out.append([cell_to_json(v) for v in row])
    return out


# Employee Utilization is a special-shape file — `employees.json`.
# It's pulled from the Utilization (Employee) sheet.
# Sheet structure (0-indexed col → 1-indexed key the build code expects):
#   col 0  → office name
#   col 1  → employee id
#   col 2  → "Last, First" name
#   col 3  → Reg/Temp (Full Time / Part Time)
#   col 4  → Target %
#   col 5  → YTD Actual %
#   col 6  → Deviation
#   col 8-20  → weekly Q1 values   (key '9'..'21')
#   col 22 → Q1 actual            (key '23')
#   col 25-37 → weekly Q2 values   (key '26'..'38')
#   col 39 → Q2 actual            (key '40')
#   col 42-54 → weekly Q3 values   (key '43'..'55')
#   col 56 → Q3 actual            (key '57')
#   col 59-71 → weekly Q4 values   (key '60'..'72')
#   col 73 → Q4 actual            (key '74')
# Keys in JSON output are 1-indexed (sheet-column number).
def extract_employees(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return None

    # Find date-header row (row with many datetime cells)
    header_row_idx = -1
    for i, r in enumerate(rows[:10]):
        date_cells = sum(1 for v in r if isinstance(v, (_dt.datetime, _dt.date)))
        if date_cells >= 8:
            header_row_idx = i; break
    if header_row_idx < 0: return None
    header = rows[header_row_idx]

    # Weekly columns = every col that holds a datetime; stored 1-indexed
    weekly_cols = []
    date_headers = {}
    for c, v in enumerate(header):
        if isinstance(v, (_dt.datetime, _dt.date)):
            weekly_cols.append(c + 1)
            d = v if isinstance(v, _dt.datetime) else _dt.datetime(v.year, v.month, v.day)
            date_headers[c + 1] = d.date().isoformat()

    # Quarterly columns — Q1 actual, Q2 actual, etc. Located right after each weekly cluster.
    # Detect from row above (idx header_row_idx-1) which holds 'Q1 2026' etc.
    quarter_label_row = rows[header_row_idx - 1] if header_row_idx > 0 else ()
    quarterly_cols = []
    for c, v in enumerate(quarter_label_row):
        if isinstance(v, str) and re.fullmatch(r'\s*Q\d\s*2\d{3}\s*', v.strip(), re.I):
            quarterly_cols.append(c + 1)
    if not quarterly_cols:
        # Fallback to known offsets
        quarterly_cols = [23, 40, 57, 74]

    # Fixed column positions (1-indexed → 0-indexed in code via -1)
    OFFICE_C, ID_C, NAME_C, REG_C, TARGET_C, YTD_C = 0, 1, 2, 3, 4, 5

    employees = []
    for r in rows[header_row_idx + 1:]:
        if not r or not any(v not in (None, '') for v in r): continue
        nm = r[NAME_C] if NAME_C < len(r) else None
        if not isinstance(nm, str) or ',' not in nm: continue
        office = r[OFFICE_C] if OFFICE_C < len(r) else None
        emp_id = r[ID_C] if ID_C < len(r) else None
        regtemp = r[REG_C] if REG_C < len(r) else None
        target = r[TARGET_C] if TARGET_C < len(r) else None
        ytd = r[YTD_C] if YTD_C < len(r) else None
        weeks = {}
        for one_idx in weekly_cols:
            c0 = one_idx - 1
            v = r[c0] if c0 < len(r) else None
            if v is None: continue
            weeks[str(one_idx)] = v if not isinstance(v, (_dt.datetime, _dt.date)) else None
        quarters = {}
        for one_idx in quarterly_cols:
            c0 = one_idx - 1
            v = r[c0] if c0 < len(r) else None
            if v is None: continue
            quarters[str(one_idx)] = v if not isinstance(v, (_dt.datetime, _dt.date)) else None
        employees.append({
            'office': office.strip() if isinstance(office, str) else office,
            'id': emp_id,
            'name': nm.strip(),
            'regtemp': regtemp.strip() if isinstance(regtemp, str) else regtemp,
            'target': target,
            'ytd': ytd,
            'weeks': weeks,
            'quarters': quarters,
        })

    # Sort by office then name for stable display
    employees.sort(key=lambda e: ((e['office'] or '').lower(), (e['name'] or '').lower()))

    return {
        'date_headers': {str(k): v for k, v in date_headers.items()},
        'weekly_cols': weekly_cols,
        'quarterly_cols': quarterly_cols,
        'employees': employees,
    }



def compute_mtd_billable(source_xlsx: str):
    """Read 'Raw Data' sheet and compute Month-to-Date Billable Hours.

    Summarizes col F (Utilized hours) for rows whose col D (Entry date) is in
    the current calendar month, excluding rows whose col K (Task name) contains
    "non-posting" (which covers both "Overhead Non-Posting" and the regional
    "CA - Overhead Non-Posting" variant).

    Returns dict with the MTD value and context, or None if Raw Data is missing.
    Uses python-calamine if available for ~10x faster reads; falls back to
    openpyxl read-only mode otherwise.
    """
    today = report_date()
    yr, mo = today.year, today.month
    month_name = today.strftime('%B')

    # Calamine fast path
    rows_iter = None
    try:
        from python_calamine import CalamineWorkbook  # type: ignore
        _wb = CalamineWorkbook.from_path(source_xlsx)
        if 'Raw Data' not in _wb.sheet_names:
            return None
        rows_iter = iter(_wb.get_sheet_by_name('Raw Data').to_python(skip_empty_area=True))
        next(rows_iter, None)  # skip header
    except ImportError:
        # Fallback to openpyxl (slower)
        _wb = load_workbook(source_xlsx, data_only=True, read_only=True)
        if 'Raw Data' not in _wb.sheetnames:
            _wb.close(); return None
        _ws = _wb['Raw Data']
        rows_iter = (row for row in _ws.iter_rows(min_row=2, values_only=True))

    total = 0.0
    excluded = 0.0
    rows_in_month = 0
    rows_excluded = 0
    last_entry_date = None
    for row in rows_iter:
        if not row or row[0] is None:
            continue
        entry_date = row[3] if len(row) > 3 else None
        util_hrs   = row[5] if len(row) > 5 else None
        task_name  = row[10] if len(row) > 10 else None
        if not isinstance(entry_date, (_dt.date, _dt.datetime)):
            continue
        if entry_date.year != yr or entry_date.month != mo:
            continue
        if util_hrs is None:
            continue
        try:
            h = float(util_hrs)
        except (TypeError, ValueError):
            continue
        if last_entry_date is None or entry_date > last_entry_date:
            last_entry_date = entry_date
        rows_in_month += 1
        if isinstance(task_name, str) and 'non-posting' in task_name.lower():
            excluded += h
            rows_excluded += 1
        else:
            total += h

    # Normalize last_entry_date to date for JSON
    if isinstance(last_entry_date, _dt.datetime):
        last_entry_date = last_entry_date.date()

    return {
        'mtd_billable_hours': round(total, 1),
        'excluded_hours': round(excluded, 1),
        'rows_in_month': rows_in_month,
        'rows_excluded': rows_excluded,
        'month_year': f'{yr}-{mo:02d}',
        'month_name': month_name,
        'year': yr,
        'last_entry_date': last_entry_date.isoformat() if last_entry_date else None,
        'generated_at': _dt.datetime.now().isoformat(timespec='seconds'),
    }


def main(source_xlsx: str, outputs_dir: str = HERE):
    print(f'Reading {source_xlsx}')
    wb = load_workbook(source_xlsx, data_only=True, read_only=True)
    site_data = os.path.join(outputs_dir, 'site_data')
    os.makedirs(site_data, exist_ok=True)

    manifest = []
    for sheet_name in wb.sheetnames:
        # Skip helper sheets that aren't shown in the site
        if sheet_name.startswith('Backend') or sheet_name.startswith('Raw'):
            continue
        ws = wb[sheet_name]
        rows = extract_sheet(ws)
        slug = slugify(sheet_name)
        out_path = os.path.join(site_data, f'{slug}.json')
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump({'tab': sheet_name, 'rows': rows}, f)
        manifest.append({'tab': sheet_name, 'slug': slug, 'rows': len(rows)})
        print(f'  {sheet_name:40s} → {slug}.json ({len(rows)} rows)')

    # Save manifest
    with open(os.path.join(site_data, '_manifest.json'), 'w', encoding='utf-8') as f:
        json.dump(manifest, f, indent=2)

    # On_Call_Trend tab uses PT_Hours data (separate slug)
    pt_path = os.path.join(site_data, 'PT_Hours.json')
    if os.path.exists(pt_path):
        import shutil
        shutil.copyfile(pt_path, os.path.join(site_data, 'On_Call_Trend.json'))
        print('  PT_Hours.json → On_Call_Trend.json (copy)')

    # Employees (special shape)
    emp_sheet = None
    for sn in wb.sheetnames:
        if 'utilization' in sn.lower() and 'employee' in sn.lower():
            emp_sheet = sn; break
    if emp_sheet:
        emp_data = extract_employees(wb[emp_sheet])
        if emp_data is not None:
            with open(os.path.join(outputs_dir, 'employees.json'), 'w', encoding='utf-8') as f:
                json.dump(emp_data, f)
            print(f'  {emp_sheet} → employees.json ({len(emp_data["employees"])} employees)')

    wb.close()

    # MTD Billable Hours (Raw Data → site_data/_mtd_billable.json)
    try:
        mtd = compute_mtd_billable(source_xlsx)
        if mtd is not None:
            with open(os.path.join(site_data, '_mtd_billable.json'), 'w', encoding='utf-8') as f:
                json.dump(mtd, f, indent=2)
            print(f"  MTD Billable Hours → _mtd_billable.json ({mtd['mtd_billable_hours']:,.1f} hrs, {mtd['month_name']} {mtd['year']})")
    except Exception as e:
        print(f"  [warn] MTD computation failed: {e}")

    print('Extraction complete.')


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'utilization_report.xlsx')
    main(src)
