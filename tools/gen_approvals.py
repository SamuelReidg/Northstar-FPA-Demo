#!/usr/bin/env python3
"""Write the three Approvals Dashboard source exports for the demo package.

    python3 tools/gen_approvals.py            # -> demo_data/Weekly and Monthly Updates/9.29.2026/
                                              #    Timesheets Draft & Submitted/

Outputs (same shapes as the ERP's weekly downloads; the drifting "(n)" in each name is
normal -- build_dashboard_html.py identifies the files by content, not by name):

  Timesheets (14).csv      Time > Staff Timesheets > "Location" view. One row per timesheet that
                           is NOT yet approved (Draft / Submitted / Declined) -- approved sheets
                           drop out of this view. UTF-8 with BOM, two unlabeled trailing columns
                           carrying the view's "Print" / "Duplicate" action links. Newest period first.
  Employees (4).xlsx       Time > Employees > "Detailed Employee View". Every person ever on the
                           roster (Not Employed = True for leavers). All cells are text; blanks
                           are a single space, as the ERP writes them.
  Staff expenses (29).xlsx Expenses > Staff Expenses > "Detail" view, sheet "Staff expenses", one
                           row per expense report, newest Date Filed first.

Everything is derived from the shared world (tools/demo_world.py):
  * timesheets = one per employee per week they have W.entries for (period Sun..Sat, the week
    ending Saturday). Whether a sheet is still open is decided by its age: the in-flight week is
    mostly open, the week before partly, and a thin tail of older stragglers never got
    submitted/approved. Stragglers from before the current accounting period have their GL post
    date batched onto the period's first day (the ERP's clamp); others post on the End date.
    Leavers often leave their final week in Draft; executives run late on their own sheets.
    A few people pre-fill future weeks (planned leave / fieldwork).
  * the employee master = W.employees (supervisor, office Location ID, Legacy ID = world id,
    Title, Department ID from the office's dept code, FLSA, paygroup, email).
  * expense reports = W.expense_reports (state, filed, GL, total, last modified) with the
    "Reason for expense" text built from each report's own W.expense_lines (types and dates).
    As on the timesheet side, open reports left over from a closed accounting period carry the
    open period's first day as Date Filed / GL Posting Date (the ERP's batching clamp).
Extra fields (time of day, payment dates, which sheets are stragglers) use a seeded
random.Random(f"{W.seed}-approvals"), so the output is byte-for-byte deterministic.
"""
import csv
import datetime as dt
import os
import random
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from demo_world import world  # noqa: E402

import openpyxl  # noqa: E402
from openpyxl.styles import PatternFill, Font, Alignment  # noqa: E402

ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, 'demo_data', 'Weekly and Monthly Updates', '9.29.2026',
                   'Timesheets Draft & Submitted')
TS_NAME, EMP_NAME, EX_NAME = 'Timesheets (14).csv', 'Employees (4).xlsx', 'Staff expenses (29).xlsx'

HEADER_FILL = PatternFill('solid', fgColor='FFDDD7C1')
BLANK = ' '                              # the ERP writes empty text cells as a single space


def erp_id(e):
    """The ERP Employee ID ("E#####") used in all three exports: the world's sage_id."""
    return e['sage_id']


def save_fixed(wb, path, stamp=dt.datetime(2026, 9, 28, 7, 30)):
    """Save a workbook with pinned document timestamps and zip entry dates, so reruns are
    byte-identical (openpyxl otherwise stamps the save time into docProps/core.xml)."""
    import io, re, zipfile
    wb.properties.creator = 'ERP export'
    wb.properties.created = stamp
    buf = io.BytesIO()
    wb.save(buf)
    iso = stamp.strftime('%Y-%m-%dT%H:%M:%SZ')
    zin = zipfile.ZipFile(io.BytesIO(buf.getvalue()))
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'docProps/core.xml':
                data = re.sub(rb'(<dcterms:(?:created|modified)[^>]*>)[^<]*', lambda m: m.group(1) + iso.encode(), data)
            zi = zipfile.ZipInfo(item.filename, date_time=stamp.timetuple()[:6])
            zi.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(zi, data)


def mdY(d):
    return d.strftime('%m/%d/%Y')


def main():
    W = world()
    rnd = random.Random(f'{W.seed}-approvals')
    as_of = W.as_of                                   # 2026-09-26, last complete week
    export_day = as_of + dt.timedelta(days=2)         # views pulled Monday 09/28
    period_start = as_of.replace(day=1)               # current accounting period (09/01)
    os.makedirs(OUT, exist_ok=True)

    emps = W.employees
    by_name = {e['name']: e for e in emps}
    # ERP display quirk: some people's names come out as "Last, First - Employee". Per person,
    # so the same person carries it in both the timesheet and the expense export.
    suffixed = {str(e['id']) for e in emps if rnd.random() < 0.11}

    def disp(e):
        return e['name'] + (' - Employee' if str(e['id']) in suffixed else '')

    # ------------------------------------------------------------------ employee master
    ceo = next(e for e in emps if e['office'] == 'Corporate' and e['title'] == 'Chief Executive Officer')
    def title(e):
        return e['title']

    def paygroup(e):
        if e['intl']:
            return 'Contractor' if str(e['id']).startswith('S') else 'NH Gulf'
        return e['paygroup']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Employees'
    hdr = ['Employee ID', 'Legacy ID', 'Primary contact name', 'Title', 'First name', 'Last name',
           'Primary email address', 'Location ID', 'Department ID', 'Manager', 'Supervisor name',
           'FLSA STATUS', 'PAYGROUP', 'Salesforce ID', 'Not Employed']
    ws.append(hdr)
    for e in emps:
        o = W.office[e['office']]
        sup = by_name.get(e['supervisor'] or '')
        leaver = e['status'] != 'Active'
        # leavers from before FY2026 often have their supervisor / FLSA cleared in the master
        stale = leaver and (e['term'] is None or e['term'].year < 2026)
        flsa = BLANK if stale and rnd.random() < 0.25 else e['flsa']
        sup_name = BLANK if (stale and rnd.random() < 0.02) or not sup else sup['name']
        ws.append([erp_id(e), str(e['id']), e['name'], title(e), e['first'], e['last'], e['email'],
                   o['location_id'], 'D%d' % o['dept'], erp_id(sup) if sup_name != BLANK else BLANK,
                   sup_name, flsa, paygroup(e), BLANK, 'True' if leaver else 'False'])
    for c in ws[1]:
        c.fill = HEADER_FILL
    for col, w in zip('ABCDEFGHIJKLMNO', (12.7, 12.9, 33.6, 28.6, 21.1, 18.9, 37.3, 10.6, 13.1, 8.4,
                                          33.6, 17.4, 22.3, 12.3, 13.0)):
        ws.column_dimensions[col].width = w
    save_fixed(wb, os.path.join(OUT, EMP_NAME))

    # ------------------------------------------------------------------ timesheets
    weeks_by_emp = defaultdict(set)
    for r in W.entries:
        weeks_by_emp[r['emp']].add(r['week'])

    def open_prob(age_weeks):
        # share of sheets still open in the export, by how many weeks ago the period ended
        return {0: 0.30, 1: 0.10, 2: 0.03, 3: 0.015, 4: 0.01}.get(age_weeks, 0.0008)

    # executives are habitually late with their own sheets (they approve everyone else's)
    late = {'Chief Executive Officer', 'Vice President of Operations', 'Chief Financial Officer',
            'Controller', 'Director of Finance'}

    ts_rows = []
    for e in emps:
        wk = sorted(weeks_by_emp.get(str(e['id']), ()))
        boost = 8 if e['office'] == 'Corporate' and title(e) in late else 1
        if not wk:
            continue
        last_week = wk[-1]
        for w in wk:
            age = (as_of - w).days // 7
            left_after = e['term'] is not None and w == last_week and e['status'] != 'Active'
            if left_after and rnd.random() < 0.15:
                state = 'Draft'                                     # leaver's final sheet, never submitted
            elif rnd.random() < open_prob(age) * (boost if age >= 2 else 1):
                state = rnd.choices(['Draft', 'Submitted'], weights=[4, 5] if age == 0 else [6, 4])[0]
            elif rnd.random() < 0.006:
                state = 'Declined'                                  # sent back; the corrected copy is a new sheet
            else:
                continue                                            # approved -> not in this view
            gl = w
            if state in ('Draft', 'Submitted') and w < period_start and rnd.random() < 0.7:
                gl = period_start                                   # batched onto the open period
            ts_rows.append((e, w - dt.timedelta(days=6), w, gl, state))
        # a few active people pre-fill upcoming weeks (planned leave / fieldwork)
        if e['status'] == 'Active' and last_week == as_of:
            n_future = 1 if rnd.random() < 0.22 else 0
            if n_future and rnd.random() < 0.18:
                n_future += rnd.randint(1, 7)
            for k in range(1, n_future + 1):
                w = as_of + dt.timedelta(days=7 * k)
                ts_rows.append((e, w - dt.timedelta(days=6), w, w,
                                rnd.choices(['Draft', 'Submitted'], weights=[3, 2])[0]))

    ts_rows.sort(key=lambda t: (-t[2].toordinal(), erp_id(t[0])))
    with open(os.path.join(OUT, TS_NAME), 'w', newline='', encoding='utf-8-sig') as f:
        wr = csv.writer(f, lineterminator='\n')
        wr.writerow(['Employee ID', 'Employee name', 'Employee last name', 'Employee first name',
                     'Begin date', 'End date', 'GL post date', 'State', '', ''])
        for e, b, en, gl, st in ts_rows:
            wr.writerow([erp_id(e), disp(e), e['last'], e['first'], mdY(b), mdY(en), mdY(gl), st,
                         'Print', 'Duplicate'])

    # ------------------------------------------------------------------ staff expenses
    lines = defaultdict(list)
    for ln in W.expense_lines:
        lines[ln['er']].append(ln)

    def reason(r):
        ls = lines.get(r['er'], [])
        types = []
        for ln in ls:
            if ln['type'] not in types:
                types.append(ln['type'])
        d0 = min((ln['date'] for ln in ls), default=r['week'])
        d1 = max((ln['date'] for ln in ls), default=r['week'])
        span = d0.strftime('%m/%d/%y') + ('' if d0 == d1 else ' -> ' + d1.strftime('%m/%d/%y'))
        if not types:
            return r['reason']
        head = types[0]
        if head in ('Lodging', 'Airfare', 'Vehicle Rental', 'Equipment Rental', 'Field Supplies'):
            # card purchases are reconciled as one statement-range line
            return 'Corp Card %s - %s' % (d0.strftime('%-m/%-d/%y'), d1.strftime('%-m/%-d/%y')) \
                if d0 != d1 else 'Corp Card %s' % d0.strftime('%-m/%-d/%y')
        s = '%s - %s' % (head, span)
        if len(types) > 1 and types[1] in ('Mileage', 'Per Diem'):
            s += ' & ' + types[1]
        return s

    ex_rows = []
    for r in W.expense_reports:
        e = W.emp[r['emp']]
        st = r['state']
        filed, gl = r['filed'], r['gl']
        if st in ('Draft', 'Submitted') and gl < period_start and rnd.random() < 0.6:
            # the ERP batches open reports from a closed period onto the open period's first day
            filed = gl = period_start
        amt = -r['total'] if st == 'Reversal' else r['total']
        paid = BLANK
        lm = r['last_modified']
        if st == 'Paid':
            pd = min(export_day, gl + dt.timedelta(days=rnd.randint(4, 12)))
            paid = dt.datetime(pd.year, pd.month, pd.day)
            lm = max(lm, pd)
        age = (as_of - gl).days
        if st in ('Draft', 'Submitted', 'Declined'):
            summ, due = 'Open', (amt if st != 'Draft' or rnd.random() < 0.4 else 0)
        elif st == 'Approved' and age <= 14:
            summ, due = 'Open', amt                                # approved, awaiting the pay run
        else:
            summ, due = 'Closed', 0
        t = dt.datetime(lm.year, lm.month, lm.day, rnd.randint(6, 19), rnd.randint(0, 59), rnd.randint(0, 59))
        ex_rows.append([erp_id(e), disp(e), r['er'], reason(r), st,
                        dt.datetime(filed.year, filed.month, filed.day),
                        dt.datetime(gl.year, gl.month, gl.day),
                        int(amt) if amt == int(amt) else amt, paid,
                        int(due) if due == int(due) else due,
                        t.strftime('%m/%d/%Y %I:%M:%S %p'), summ])
    ex_rows.sort(key=lambda x: (-x[5].toordinal(), x[2]))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Staff expenses'
    ws.append(['Employee ID', 'Employee Name', 'ER #', 'Reason for expense', 'State', 'Date Filed',
               'GL Posting Date', 'Amount', 'Date Paid', 'Amount Due', 'Last Modified', 'Summary State'])
    for row in ex_rows:
        ws.append(row)
    for c in ws[1]:
        c.fill = HEADER_FILL
    fmts = {6: 'MM/DD/YYYY', 7: 'MM/DD/YYYY', 8: '$#,##0.00;([Red]$#,##0.00)',
            9: 'MM/DD/YYYY', 10: '#,##0.00;([Red]#,##0.00)'}
    for row in ws.iter_rows(min_row=2):
        for ci, fm in fmts.items():
            if row[ci - 1].value != BLANK:
                row[ci - 1].number_format = fm
    for col, w in zip('ABCDEFGHIJKL', (11.9, 33.6, 20.7, 60.0, 15.9, 10.4, 15.0, 11.6, 10.4, 11.6, 21.6, 14.3)):
        ws.column_dimensions[col].width = w
    save_fixed(wb, os.path.join(OUT, EX_NAME))

    n_open_ts = sum(1 for t in ts_rows if t[4] in ('Draft', 'Submitted'))
    print('wrote %s\n  %s: %d rows (%d open)\n  %s: %d people\n  %s: %d reports'
          % (OUT, TS_NAME, len(ts_rows), n_open_ts, EMP_NAME, len(emps), EX_NAME, len(ex_rows)))
    print('  CEO (excluded on the dashboard): %s  %s' % (ceo['name'], erp_id(ceo)))


if __name__ == '__main__':
    main()
