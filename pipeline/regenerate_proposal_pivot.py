#!/usr/bin/env python3
"""Emit proposal_pivot.js — the "Proposal Hours" tab of the Excel export, built from Raw Data.

FP&A (2026-09-17): the export workbook should carry a table like the workbook's old
`Pivot - Proposal` sheet — the 8 most recent weeks of proposal time by office and employee,
plus an 8-week total — and it has to be built from `Raw Data`, because the weekly workbook
will stop shipping the `Pivot - Proposal` sheet.

What the old pivot did, reverse-engineered from the 2026-09-15 workbook and reconciled
office-by-office (see `--check` below):
  * filter  : Task name in PROPOSAL_TASKS  (the pivot's "(Multiple Items)" Task-name filter)
  * rows    : Employee Home Location  →  Contact name       (NOT `Location name`: several
              people charge proposal time to an office other than their own, and the pivot
              follows the person, not the charge)
  * values  : sum of Hours
  * columns : the last 8 `Week End` dates, then an 8-week total
The window rolls on its own — each weekly refresh drops the oldest column and adds a new one.

Two outputs, both regenerated on every weekly refresh:
  * `proposal_pivot.js`          — read by the Excel export in index.html (tiny, ~20 KB)
  * `site_data/Pivot_Proposal.json` (with --site-data) — same shape `extract_data.py` used to
    write from the now-retired sheet, so the on-screen Proposal Hours tab keeps rendering.
    Write it BEFORE build_site.py runs, i.e. straight after extract_data.py.

    python3 regenerate_proposal_pivot.py --workbook "Utilization Report.xlsx" \
        --out proposal_pivot.js --site-data site_data
    python3 regenerate_proposal_pivot.py --workbook "…" --check      # reconcile vs the old sheet

Streams the sheet XML row by row (reusing regenerate_pmo_reduced.Workbook, stdlib only):
calamine's to_python on the full-size Raw Data sheet OOMs the device.
"""
import argparse, datetime as dt, json, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from regenerate_pmo_reduced import Workbook, col_num, serial_to_date   # noqa: E402

# The pivot's Task-name filter, spelled out. Both of these are also the only tasks the
# workbook's own `Proposal` helper column ever flags, so the two definitions agree today;
# this list is the one that governs. Add a task here if the ERP grows another proposal code.
PROPOSAL_TASKS = [
    'OVHD - Labor - Proposal Writing',
    'OVHD - Production - Proposal Editing',
]

WEEKS = 8           # rolling window width
TITLE = 'Total Proposal Hours'
TOTAL_LABEL = 'Total Hours\n(Previous %d Weeks)' % WEEKS
BLANK_OFFICE = '-'  # how the old pivot labelled rows with no home office

# Employees whose office label is overridden across the dashboard — kept in sync with
# build_site.EMPLOYEE_OFFICE_OVERRIDES, which is imported when it can be. Keyed by the
# employee ID, which arrives on this sheet as `Legacy ID`.
EMPLOYEE_OFFICE_OVERRIDES = {
    'S4412': 'International (Global LLC)',
    'S4413': 'International (Global LLC)',
}
try:
    from build_site import EMPLOYEE_OFFICE_OVERRIDES as _BS_OVERRIDES   # noqa: E402
    if _BS_OVERRIDES != EMPLOYEE_OFFICE_OVERRIDES:
        print('note: taking EMPLOYEE_OFFICE_OVERRIDES from build_site.py:', _BS_OVERRIDES)
    EMPLOYEE_OFFICE_OVERRIDES = dict(_BS_OVERRIDES)
except Exception as e:                                                  # pragma: no cover
    print(f'note: using this file\'s copy of EMPLOYEE_OFFICE_OVERRIDES ({e})')

NEEDED = ['Legacy ID', 'Contact name', 'Employee Home Location', 'Hours', 'Task name', 'Week End']


def _emp_id(v):
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v or '').strip()


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _week_key(v):
    """`Week End` arrives as an Excel serial (or, rarely, already a string)."""
    if isinstance(v, (int, float)):
        return serial_to_date(float(v)).isoformat()
    return str(v or '')[:10]


def collect(bk, sheet='Raw Data'):
    """One streaming pass → {(office, employee): {week: hours}} plus the set of weeks seen."""
    it = bk.rows(sheet)
    try:
        _, first = next(it)
    except StopIteration:
        sys.exit('ERROR: sheet is empty')
    letters = {}
    for c in sorted(first, key=col_num):
        h = str(first[c]).strip()
        if h in NEEDED and h not in letters:
            letters[h] = c
    missing = [h for h in NEEDED if h not in letters]
    if missing:
        sys.exit(f'ERROR: {sheet} is missing columns: {missing}')

    tasks = set(PROPOSAL_TASKS)
    cells, weeks = {}, set()
    scanned = matched = 0
    for _, d in it:
        if not d:
            continue
        scanned += 1
        if str(d.get(letters['Task name'], '')).strip() not in tasks:
            continue
        matched += 1
        week = _week_key(d.get(letters['Week End']))
        if not week:
            continue
        office = str(d.get(letters['Employee Home Location'], '') or '').strip() or BLANK_OFFICE
        office = EMPLOYEE_OFFICE_OVERRIDES.get(_emp_id(d.get(letters['Legacy ID']))) or office
        emp = str(d.get(letters['Contact name'], '') or '').strip() or '(unnamed)'
        weeks.add(week)
        cells.setdefault((office, emp), {})[week] = \
            cells.setdefault((office, emp), {}).get(week, 0.0) + _num(d.get(letters['Hours']))
    print(f'  scanned {scanned:,} rows, {matched:,} proposal rows, {len(weeks)} distinct weeks')
    return cells, sorted(weeks)


def build(cells, all_weeks, weeks=WEEKS):
    """Shape the aggregate into the pivot's row order: office block, employees, grand total."""
    window = all_weeks[-weeks:]
    if len(window) < weeks:
        print(f'  WARNING: only {len(window)} weeks of proposal time in the sheet')

    by_office = {}
    for (office, emp), wk in cells.items():
        vals = [round(wk.get(w, 0.0), 4) for w in window]
        if not any(vals):
            continue                      # employee has nothing inside the window
        by_office.setdefault(office, []).append((emp, vals))

    out = []
    # Offices alphabetically (FP&A, 2026-09-17); the unassigned "-" block leads, as it did
    # in the old pivot. Employees inside an office run highest 8-week total first.
    for office in sorted(by_office, key=lambda o: (o != BLANK_OFFICE, o.lower())):
        members = sorted(by_office[office], key=lambda t: (-sum(t[1]), t[0].lower()))
        totals = [round(sum(m[1][i] for m in members), 4) for i in range(len(window))]
        out.append({'t': 'o', 'l': office, 'v': totals + [round(sum(totals), 4)]})
        for emp, vals in members:
            out.append({'t': 'e', 'l': emp, 'v': vals + [round(sum(vals), 4)]})

    grand = [round(sum(r['v'][i] for r in out if r['t'] == 'o'), 4) for i in range(len(window) + 1)]
    out.append({'t': 'g', 'l': TOTAL_LABEL, 'v': grand})
    return window, out


def to_aoa(window, rows):
    """The old sheet's literal layout — what `site_data/Pivot_Proposal.json` has to look like
    for build_site.render_proposal_hours_v3 (it finds the header by counting date cells)."""
    width = len(window) + 2
    pad = lambda r: r + [''] * (width - len(r))
    aoa = [pad(['Task name', '(Multiple Items)']),
           pad([TITLE]),
           pad([' ', '  ']),
           [' '] + [{'__date__': w + 'T00:00:00'} for w in window] + [TOTAL_LABEL]]
    for r in rows:
        aoa.append([r['l']] + [(v if v else '') for v in r['v']])
    return aoa


def check(bk, window, rows, sheet='Pivot - Proposal'):
    """Reconcile against the workbook's own cached pivot, while it is still shipped."""
    if sheet not in bk.sheets:
        print(f'--check: {sheet!r} is not in this workbook — nothing to reconcile against')
        return 0
    # Blank cells are simply absent from the sheet XML, so every lookup has to go through the
    # header row's column LETTERS — collapsing each row to a dense list misaligns the columns.
    old, old_weeks, date_cols, label_col = {}, [], [], None
    for _, d in bk.rows(sheet):
        if not d:
            continue
        cols = sorted(d, key=col_num)
        if not old_weeks:
            dates = [c for c in cols if isinstance(d[c], float) and 40000 < d[c] < 60000]
            if len(dates) >= 3:
                date_cols = dates
                old_weeks = [serial_to_date(d[c]).isoformat() for c in dates]
                label_col = cols[0]
            continue
        label = d.get(label_col)
        if isinstance(label, str) and label.strip():
            old[label.strip()] = [_num(d.get(c)) for c in date_cols]
    if not old_weeks:
        print('--check: could not find the header row on the cached pivot')
        return 0
    shared = [w for w in window if w in old_weeks]
    print(f'--check: cached pivot covers {old_weeks[0]}…{old_weeks[-1]}; '
          f'{len(shared)} of our {len(window)} weeks overlap')
    bad = 0
    for r in rows:
        if r['t'] == 'g' or r['l'] not in old:
            continue
        for w in shared:
            mine = r['v'][window.index(w)]
            theirs = old[r['l']][old_weeks.index(w)]
            if abs(mine - theirs) > 0.001:
                bad += 1
                print(f'  DIFF {r["l"]:<32} {w}  ours {mine:>8.2f}  cached {theirs:>8.2f}')
    missing = [r['l'] for r in rows if r['t'] != 'g' and r['l'] not in old]
    if missing:
        print(f'  {len(missing)} rows absent from the cached pivot (new since it was refreshed): '
              + ', '.join(missing[:8]) + ('…' if len(missing) > 8 else ''))
    print(f'--check: {bad} differing cell(s) over the overlapping weeks')
    return bad


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--workbook', required=True)
    ap.add_argument('--out', default='proposal_pivot.js')
    ap.add_argument('--sheet', default='Raw Data')
    ap.add_argument('--site-data', help='directory to also write Pivot_Proposal.json into')
    ap.add_argument('--weeks', type=int, default=WEEKS)
    ap.add_argument('--check', action='store_true',
                    help="reconcile against the workbook's cached Pivot - Proposal sheet")
    a = ap.parse_args()

    bk = Workbook(a.workbook)
    if a.sheet not in bk.sheets:
        sys.exit(f'ERROR: sheet {a.sheet!r} not found: {sorted(bk.sheets)[:20]}')

    cells, all_weeks = collect(bk, a.sheet)
    window, rows = build(cells, all_weeks, a.weeks)

    payload = {
        'generated': dt.datetime.now().isoformat(timespec='seconds'),
        'source_sheet': a.sheet,
        'sheet_name': 'Proposal Hours',
        'title': TITLE,
        'tasks': PROPOSAL_TASKS,
        'weeks': window,
        'total_label': TOTAL_LABEL.replace('\n', ' '),
        'rows': rows,
    }
    with open(a.out, 'w', encoding='utf-8') as f:
        f.write('/* Proposal Hours tab for the Excel export — regenerated weekly by '
                'regenerate_proposal_pivot.py from the Raw Data sheet.\n'
                '   Loaded lazily by index.html when Export runs. */\n')
        f.write('window.PROPOSAL_PIVOT_SHEET=' + json.dumps(payload, separators=(',', ':')) + ';\n')
    offices = sum(1 for r in rows if r['t'] == 'o')
    print(f'{a.out}: {window[0]}…{window[-1]}, {offices} offices, {len(rows)} rows, '
          f'{rows[-1]["v"][-1]:,.2f} hrs, {os.path.getsize(a.out)/1024:.0f} KB')

    if a.site_data:
        os.makedirs(a.site_data, exist_ok=True)
        p = os.path.join(a.site_data, 'Pivot_Proposal.json')
        with open(p, 'w', encoding='utf-8') as f:
            json.dump({'tab': 'Pivot - Proposal', 'rows': to_aoa(window, rows)}, f)
        print(f'{p}: on-screen Proposal Hours tab ({len(rows) + 4} rows)')

    if a.check:
        check(bk, window, rows)


if __name__ == '__main__':
    main()
