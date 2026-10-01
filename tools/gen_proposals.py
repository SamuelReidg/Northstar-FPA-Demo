"""Write the synthetic monthly "Proposal Hours Summary" export from the demo world.

    python3 tools/gen_proposals.py

Output: demo_data/Weekly and Monthly Updates/9.29.2026/Proposal Hours Summary/
        Proposal Hours Summary_9.2026_.xlsx

Layout mirrors the monthly workbook that pipeline/regenerate_proposal_data.py reads (same sheet
names and order, header rows, column order, label text and quirks). Everything is derived from
tools/demo_world.py:

  * Hours  = W.proposal_hours (the proposal-writing / proposal-editing timesheet rows, each tied
             to an opportunity), bucketed by the charging employee's ERP location and by month.
  * # / $  = W.opportunities, bucketed by the opportunity's office (CRM "Project Management
             Office") and its created month. Every opportunity counts (status "Any").
  * Hours per opportunity = month hours / opportunities created that month; average proposal
    size = pipeline $ / opportunities; region, PMO-total and grand-total rows are built from the
    same cells, so every roll-up is the sum of its children.

The reporting window is FY2026 year-to-date (January through the as-of month; the September
column is month-to-date through the as-of week), which is the span the demo timesheets cover.
The only seeded choices are free-text flavour (timesheet notes) and which International
opportunities were booked to the separate "Global LLC" CRM office.
"""
from __future__ import annotations

import datetime as dt
import os
import random
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from demo_world import world, LOCATION_OF, NON_POSTING, eomonth  # noqa: E402

from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Font, PatternFill, Alignment  # noqa: E402

OUT_DIR = os.path.join(ROOT, 'demo_data', 'Weekly and Monthly Updates', '9.29.2026', 'Proposal Hours Summary')
OUT = os.path.join(OUT_DIR, 'Proposal Hours Summary_9.2026_.xlsx')

PROPOSAL_TASKS = ('OVHD - Labor - Proposal Writing', 'OVHD - Production - Proposal Editing')
REGION_OF_KIND = {'regional': 'Domestic', 'service': 'Domestic', 'acquired': 'Meridian',
                  'international': 'International'}
GLOBAL_LLC_PMO = 'Global LLC International'
BOLD = Font(bold=True)
HDR_FILL = PatternFill('solid', fgColor='D9E1F2')
DATE_FMT = 'mmm-yy'


def dtt(d):
    return dt.datetime(d.year, d.month, d.day)


def month_ends(first, last):
    out, d = [], first
    while d <= last:
        out.append(eomonth(d))
        d = eomonth(d) + dt.timedelta(days=1)
    return out


def r2(x):
    return round(x + 0.0, 2)


def main():
    W = world()
    rnd = random.Random(f'{W.seed}-proposals')
    as_of = W.as_of
    MONTHS = month_ends(dt.date(as_of.year, 1, 1), as_of)          # FY YTD, month-end dates
    MKEY = {(m.year, m.month): i for i, m in enumerate(MONTHS)}
    NM = len(MONTHS)

    offices = {o['name']: o for o in W.offices}
    pmo_offices = sorted([o for o in W.offices if o['kind'] in REGION_OF_KIND], key=lambda o: o['name'])

    # CRM "Project Management Office" label for each office (as the CRM report prints it)
    def crm_label(o):
        if o['kind'] == 'regional':
            return '%s, %s' % (o['name'], o['state'])
        if o['kind'] == 'acquired':
            return o['name'].replace('Meridian ', 'Meridian - ')
        return o['name']

    # ---------------- opportunities -> # / $ by CRM office x created month -------------------
    opp = {op['id']: op for op in W.opportunities}
    crm_rows = []   # (office short name, CRM label, region)
    for o in pmo_offices:
        crm_rows.append((o['name'], crm_label(o), REGION_OF_KIND[o['kind']]))
    crm_rows.append(('International', GLOBAL_LLC_PMO, 'International'))

    def crm_of(op):
        o = offices[op['office']]
        if o['kind'] == 'international':
            # a small share of International pursuits is contracted through the separate entity
            if random.Random(f"{W.seed}-proposals-{op['id']}").random() < 0.08:
                return GLOBAL_LLC_PMO
        return crm_label(o)

    cnt = defaultdict(lambda: [0] * NM)
    amt = defaultdict(lambda: [0.0] * NM)
    pipe_cnt = defaultdict(lambda: defaultdict(int))     # full CRM history for the wide sheet
    pipe_amt = defaultdict(lambda: defaultdict(float))
    for op in W.opportunities:
        lab = crm_of(op)
        c = op['created']
        pipe_cnt[lab][(c.year, c.month)] += 1
        pipe_amt[lab][(c.year, c.month)] += op['amount']
        i = MKEY.get((c.year, c.month))
        if i is None:
            continue
        cnt[lab][i] += 1
        amt[lab][i] += op['amount']

    # ---------------- proposal hours -> by ERP location x month ------------------------------
    emp = W.emp
    hrs_loc = defaultdict(lambda: [0.0] * NM)
    prop_rows = []
    ind = defaultdict(lambda: defaultdict(lambda: [0.0] * NM))   # home dept -> employee -> months

    def home_dept(e):
        return e['office'] if e['status'] == 'Active' else 'INACTIVE %s' % e['office']

    note_tpl = ['{s} proposal {n}', '{s} - draft narrative', '{s} scope & budget', 'Proposal edits - {s}',
                '{s} - interview prep', 'Go/no-go review {n}', 'Cost proposal - {s}', '{s} RFP review']
    for r in sorted((x for x in W.entries if x['project'] == 'OVHD' and x['task'] in PROPOSAL_TASKS),
                    key=lambda x: (x['emp'], x['date'], x['doc'])):
        i = MKEY.get((r['date'].year, r['date'].month))
        if i is None:
            continue
        e = emp[r['emp']]
        loc = LOCATION_OF[e['office']]
        hrs_loc[loc][i] += r['hours']
        ind[home_dept(e)][e['name']][i] += r['hours']
        op = opp.get(r.get('opportunity'))
        num = op['id'].split('-')[1].lstrip('0') if op else None
        short = ' '.join(op['name'].split(' | ')[-1].split()[:3]) if op else 'General'
        note = rnd.choice(note_tpl).format(s=short, n=num or '')
        o = offices[e['office']]
        region = REGION_OF_KIND.get(o['kind'], 'Domestic')
        prop_rows.append([str(e['id']), e['name'], loc, dtt(r['date']), 'OVHD', 0, r['hours'], r['hours'], 0,
                          r['tt'], r['task'], note, None, int(num) if num and rnd.random() < 0.55 else None,
                          region, dtt(MONTHS[i]), home_dept(e)])

    wb = Workbook()
    wb.remove(wb.active)

    # ======================= Office Summary (Proj. Loc.) ======================================
    ws = wb.create_sheet('Office Summary (Proj. Loc.)')
    nm = NM
    c_hours = list(range(2, 2 + nm))
    c_havg = 2 + nm + 1
    c_size = list(range(c_havg + 2, c_havg + 2 + nm))
    c_savg = c_size[-1] + 2
    ws.cell(1, 1, 'Hours Spent per Opportunity (Favorable if Lower)').font = BOLD
    ws.cell(1, c_size[0], 'Average Proposal Size').font = BOLD
    ws.cell(2, 1, 'Office').font = BOLD
    for cols, avgc in ((c_hours, c_havg), (c_size, c_savg)):
        for j, c in enumerate(cols):
            x = ws.cell(2, c, dtt(MONTHS[j])); x.number_format = DATE_FMT; x.font = BOLD; x.fill = HDR_FILL
        ws.cell(2, avgc, 'Ave. per \nPeriod').font = BOLD

    def office_cnt(name):
        tot = [0] * nm
        for (short, lab, _reg) in crm_rows:
            if short == name:
                for i in range(nm):
                    tot[i] += cnt[lab][i]
        return tot

    def office_amt(name):
        tot = [0.0] * nm
        for (short, lab, _reg) in crm_rows:
            if short == name:
                for i in range(nm):
                    tot[i] += amt[lab][i]
        return tot

    def ratio(h, n):
        if n:
            return h / n
        return 0 if abs(h) < 1e-9 else '-'

    row = 3
    pmo_h, pmo_n = [0.0] * nm, [0] * nm
    for o in pmo_offices:
        loc = LOCATION_OF[o['name']]
        h, n, a = hrs_loc[loc], office_cnt(o['name']), office_amt(o['name'])
        ws.cell(row, 1, o['name'])
        for i in range(nm):
            ws.cell(row, c_hours[i], ratio(h[i], n[i]))
            ws.cell(row, c_size[i], (a[i] / n[i]) if n[i] else 0)
            pmo_h[i] += h[i]
        ws.cell(row, c_havg, ratio(sum(h), sum(n)))
        ws.cell(row, c_savg, (sum(a) / sum(n)) if sum(n) else '#DIV/0!')
        row += 1
    tot_n = [sum(cnt[lab][i] for (_s, lab, _r) in crm_rows) for i in range(nm)]
    tot_a = [sum(amt[lab][i] for (_s, lab, _r) in crm_rows) for i in range(nm)]
    dom_n = [sum(cnt[lab][i] for (_s, lab, reg) in crm_rows if reg == 'Domestic') for i in range(nm)]
    ws.cell(row, 1, 'PMO Total').font = BOLD
    for i in range(nm):
        ws.cell(row, c_hours[i], ratio(pmo_h[i], tot_n[i]))
    ws.cell(row, c_havg, ratio(sum(pmo_h), sum(tot_n)))
    row += 2
    corp_h = hrs_loc[LOCATION_OF['Corporate']]
    sm_h = hrs_loc[LOCATION_OF['Sales & Marketing']]
    for lab, h in (('Corporate*', corp_h), ('Sales and Marketing*', sm_h)):
        ws.cell(row, 1, lab)
        for i in range(nm):
            ws.cell(row, c_hours[i], ratio(h[i], dom_n[i]))
        ws.cell(row, c_havg, ratio(sum(h), sum(dom_n)))
        row += 1
    row += 1
    ws.cell(row, 1, 'Grand Total**').font = BOLD
    all_h = [pmo_h[i] + corp_h[i] + sm_h[i] for i in range(nm)]
    for i in range(nm):
        ws.cell(row, c_hours[i], ratio(all_h[i], tot_n[i]))
        ws.cell(row, c_size[i], (tot_a[i] / tot_n[i]) if tot_n[i] else '#DIV/0!')
    ws.cell(row, c_havg, ratio(sum(all_h), sum(tot_n)))
    ws.cell(row, c_savg, sum(tot_a) / sum(tot_n))
    row += 2
    for lab in ('< 8 Hours', '8 - 10 Hours', '> 10 Hours'):
        ws.cell(row, 1, lab); row += 1
    row += 1
    ws.cell(row, 1, '* Corporate and S&M are calculated as: Total Proposal Hours Spent by Corp-S&M / '
                    'Total # of Domestic Projects.'); row += 1
    ws.cell(row, 1, '** Sum of PMO Total, Corporate, and S&M do not match Grand Total due to weighting of segments. ')
    row += 3
    ws.cell(row, 1, 'Data').font = BOLD
    row += 2

    def data_block(title, first_label, rows, total_label='Grand Total'):
        nonlocal row
        ws.cell(row, 1, title).font = BOLD; row += 1
        ws.cell(row, 1, first_label).font = BOLD
        for i in range(nm):
            x = ws.cell(row, c_hours[i], dtt(MONTHS[i])); x.number_format = DATE_FMT; x.font = BOLD
        ws.cell(row, c_havg, 'Total').font = BOLD
        row += 1
        tot = [0.0] * nm
        for lab, vals in rows:
            ws.cell(row, 1, lab)
            for i in range(nm):
                ws.cell(row, c_hours[i], r2(vals[i]) if isinstance(vals[i], float) else vals[i])
                tot[i] += vals[i]
            ws.cell(row, c_havg, r2(sum(vals)) if isinstance(vals[0], float) else sum(vals))
            row += 1
        ws.cell(row, 1, total_label).font = BOLD
        for i in range(nm):
            ws.cell(row, c_hours[i], r2(tot[i]))
        ws.cell(row, c_havg, r2(sum(tot)))
        row += 2

    data_block('Hours Spent (Proposal Time)', 'Region',
               [(LOCATION_OF[o['name']], hrs_loc[LOCATION_OF[o['name']]]) for o in
                sorted(pmo_offices, key=lambda o: LOCATION_OF[o['name']])])
    data_block('Hours Spent (Non-PMO)', None, [('Corporate', corp_h), ('Sales and Marketing', sm_h)])
    row += 1
    data_block('# of Proposals', 'Region', [(lab, cnt[lab]) for (_s, lab, _r) in
                                            sorted(crm_rows, key=lambda t: t[1]) if lab != GLOBAL_LLC_PMO])
    ws.column_dimensions['A'].width = 28

    # ======================= Segmented Summary (Proj. Loc.) ===================================
    ws = wb.create_sheet('Segmented Summary (Proj. Loc.)')
    reg_h = {'Domestic': [pmo_h[i] - sum(hrs_loc[LOCATION_OF[o['name']]][i] for o in pmo_offices
                                          if REGION_OF_KIND[o['kind']] != 'Domestic') + corp_h[i] + sm_h[i]
                          for i in range(nm)]}
    dom_pmo_h = [reg_h['Domestic'][i] - corp_h[i] - sm_h[i] for i in range(nm)]
    for reg in ('Meridian', 'International'):
        reg_h[reg] = [sum(hrs_loc[LOCATION_OF[o['name']]][i] for o in pmo_offices if REGION_OF_KIND[o['kind']] == reg)
                      for i in range(nm)]
    reg_n = {reg: [sum(cnt[lab][i] for (_s, lab, rg) in crm_rows if rg == reg) for i in range(nm)]
             for reg in ('Domestic', 'Meridian', 'International')}
    reg_a = {reg: [sum(amt[lab][i] for (_s, lab, rg) in crm_rows if rg == reg) for i in range(nm)]
             for reg in ('Domestic', 'Meridian', 'International')}
    REGS = ('Domestic', 'Meridian', 'International')
    for d in (reg_h, reg_n, reg_a):
        d['Grand Total'] = [sum(d[r][i] for r in REGS) for i in range(nm)]

    def hdr(r, first, last_label):
        ws.cell(r, 1, first).font = BOLD
        for i in range(nm):
            x = ws.cell(r, 2 + i, dtt(MONTHS[i])); x.number_format = DATE_FMT; x.font = BOLD; x.fill = HDR_FILL
        ws.cell(r, 2 + nm + 1, last_label).font = BOLD

    ws.cell(1, 1, 'Hours Spent per Opportunity (Favorable if Lower)').font = BOLD
    hdr(2, 'Region', 'Ave. Period')
    seg_rows = [('Domestic Roll-Up', reg_h['Domestic'], reg_n['Domestic']),
                ('       Domestic PMO', dom_pmo_h, reg_n['Domestic']),
                ('       Corporate', corp_h, reg_n['Domestic']),
                ('       Sales and Marketing', sm_h, reg_n['Domestic']),
                ('Meridian', reg_h['Meridian'], reg_n['Meridian']),
                ('International', reg_h['International'], reg_n['International']),
                ('Grand Total', reg_h['Grand Total'], reg_n['Grand Total'])]
    r = 3
    for lab, h, n in seg_rows:
        ws.cell(r, 1, lab)
        for i in range(nm):
            ws.cell(r, 2 + i, (h[i] / n[i]) if n[i] else '#DIV/0!')
        ws.cell(r, 2 + nm + 1, sum(h) / sum(n) if sum(n) else '#DIV/0!')
        r += 1
    ws.cell(11, 1, 'Data').font = BOLD
    r = 13
    for title, d, is_money in (('Hours Spent (Proposal Time)', reg_h, False), ('# of Proposals', reg_n, False),
                               ('$ of Proposals', reg_a, True)):
        ws.cell(r, 1, title).font = BOLD
        hdr(r + 1, 'Region', 'Total')
        for k, reg in enumerate(REGS + ('Grand Total',)):
            ws.cell(r + 2 + k, 1, reg)
            for i in range(nm):
                v = d[reg][i]
                ws.cell(r + 2 + k, 2 + i, r2(v) if isinstance(v, float) else v)
            tv = sum(d[reg])
            ws.cell(r + 2 + k, 2 + nm + 1, r2(tv) if isinstance(tv, float) else tv)
        r += 7
    ws.cell(r, 1, 'Average Proposal Size by Region').font = BOLD
    hdr(r + 1, 'Region', 'Total')
    for k, reg in enumerate(REGS + ('Grand Total',)):
        ws.cell(r + 2 + k, 1, reg)
        for i in range(nm):
            n = reg_n[reg][i]
            ws.cell(r + 2 + k, 2 + i, reg_a[reg][i] / n if n else '#DIV/0!')
        ws.cell(r + 2 + k, 2 + nm, '#DIV/0!')          # spacer column carries a stray formula
        ws.cell(r + 2 + k, 2 + nm + 1, sum(reg_a[reg]) / sum(reg_n[reg]) if sum(reg_n[reg]) else '#DIV/0!')
    r += 7
    ws.cell(r, 1, 'Notes:').font = BOLD
    ws.cell(r + 1, 1, 'Corporate and S&M Hours have been attributed to Domestic.')
    ws.column_dimensions['A'].width = 30

    # ======================= Hours Summary (PMO)  (hidden) ====================================
    ws = wb.create_sheet('Hours Summary (PMO)')
    ws.sheet_state = 'hidden'
    tot = defaultdict(lambda: [0.0] * nm)
    nonpost = defaultdict(lambda: [0.0] * nm)
    nonutil = defaultdict(lambda: [0.0] * nm)
    bill = defaultdict(lambda: [0.0] * nm)
    for e_ in W.entries:
        i = MKEY.get((e_['date'].year, e_['date'].month))
        if i is None:
            continue
        loc = LOCATION_OF[emp[e_['emp']]['office']]
        tot[loc][i] += e_['hours']
        if e_['task'] == NON_POSTING:
            nonpost[loc][i] += e_['hours']
        if e_['util'] > 0:
            bill[loc][i] += e_['util']
        else:
            nonutil[loc][i] += e_['hours']
    locs = sorted(tot)
    r = 2
    for title, d, dash in (('Total Hours', tot, False), ('Overhead Non-Posting*', nonpost, True),
                           ('Non-Utilized', nonutil, False), ('Billable Hours**', bill, False)):
        ws.cell(r, 1, title).font = BOLD
        for i in range(nm):
            x = ws.cell(r, 2 + i, dt.datetime(MONTHS[i].year, MONTHS[i].month, 1)); x.number_format = DATE_FMT
        r += 1
        gt = [0.0] * nm
        for loc in locs:
            ws.cell(r, 1, loc)
            for i in range(nm):
                v = d[loc][i]
                ws.cell(r, 2 + i, '-' if dash and abs(v) < 1e-9 else r2(v))
                gt[i] += v
            r += 1
        ws.cell(r, 1, 'Grand Total').font = BOLD
        for i in range(nm):
            ws.cell(r, 2 + i, r2(gt[i]))
        r += 3
    ws.cell(r, 1, 'Hours designations follow the ERP time type and task on each timesheet row.')
    ws.cell(r + 1, 1, '*  Overhead Non-Posting = hours charged to the Overhead Non-Posting task.')
    ws.cell(r + 2, 1, '** Billable Hours = utilized hours (project time).')

    # ======================= Individual Summary (Emp. Loc.) ===================================
    ws = wb.create_sheet('Individual Summary (Emp. Loc.)')
    ws.cell(3, 1, 'Sum of Hours').font = BOLD
    ws.cell(3, 2, 'Column Labels').font = BOLD
    ws.cell(4, 1, 'Row Labels').font = BOLD
    for i in range(nm):
        x = ws.cell(4, 2 + i, dtt(MONTHS[i])); x.number_format = DATE_FMT; x.font = BOLD
    ws.cell(4, 2 + nm, 'Grand Total').font = BOLD
    r = 5
    grand = [0.0] * nm

    def put_vals(rr, vals, blank_zero):
        for i in range(nm):
            v = vals[i]
            ws.cell(rr, 2 + i, None if (blank_zero and abs(v) < 1e-9) else r2(v))
        ws.cell(rr, 2 + nm, r2(sum(vals)))

    for dept in sorted(ind, key=lambda d: (d.startswith('INACTIVE'), d)):
        people = ind[dept]
        dv = [sum(p[i] for p in people.values()) for i in range(nm)]
        ws.cell(r, 1, dept).font = BOLD
        put_vals(r, dv, False)
        r += 1
        for name in sorted(people, key=lambda n: (-sum(people[n]), n)):
            ws.cell(r, 1, name)
            put_vals(r, people[name], True)
            r += 1
        grand = [grand[i] + dv[i] for i in range(nm)]
    ws.cell(r, 1, 'Grand Total').font = BOLD
    put_vals(r, grand, False)
    ws.column_dimensions['A'].width = 30

    # ======================= Time Data -> =====================================================
    wb.create_sheet('Time Data ->')

    # ======================= Proposal Time ====================================================
    ws = wb.create_sheet('Proposal Time')
    ws.append(['Legacy ID', 'Contact name', 'Location name', 'Entry date', 'Project ID', 'Utilized hours',
               'Non-utilized hours', 'Hours', 'Utilization Rate', 'Time type', 'Task name', 'Notes',
               'Billing type', 'Opportunity Number', 'Region', 'Month', 'Employee Home Location with SME Exceptions'])
    for c in ws[1]:
        c.font = BOLD
    for pr in prop_rows:
        ws.append(pr)
    for rr in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        rr[0].number_format = 'm/d/yyyy'
    for rr in ws.iter_rows(min_row=2, min_col=16, max_col=16):
        rr[0].number_format = 'm/d/yyyy'

    # ======================= HR Roster (employee master) ======================================
    ws = wb.create_sheet('HR Roster')
    ws.append(['Employee Number', 'Full Name', 'Status', 'Work Email', 'Department Code', 'Department Name',
               'Work Location', 'Job Title', 'Manager', 'Employee Type', 'Paygroup', 'FLSA Status', 'Hire Date',
               'Rehire Date', 'Termination Date', 'Utilization category (Text)', 'Target (Text)', 'Status Type',
               'State/Province'])
    for c in ws[1]:
        c.font = BOLD
    ucat = {'field': 'Field Staff', 'senior': 'Principal Investigator/Technical Lead', 'lead': 'Office Leadership',
            'admin': 'Corporate Non-billable'}
    for e in sorted(W.employees, key=lambda e: str(e['id'])):
        o = offices[e['office']]
        active = e['status'] == 'Active'
        ws.append([str(e['id']), e['roster_name'], e['status'], e['email'] if active else ' ', str(o['dept']),
                   home_dept(e), o['work_location'], e['title'], e['supervisor'], 'Regular', e['paygroup'],
                   e['flsa'], dtt(e['hire']), None, dtt(e['term']) if e['term'] else None,
                   ucat.get(e['role'], ' '), str(int(round(e['target'] * 100))), e['ft'], o['state']])

    # ======================= SME (Exceptions) ================================================
    ws = wb.create_sheet('SME (Exceptions)')
    ws.append(['ID', 'Name', 'Job Description', 'Revised Location', None, 'Notes'])
    for c in ws[1]:
        c.font = BOLD
    sm_staff = [e for e in W.employees if e['office'] == 'Sales & Marketing'][:4]
    for k, e in enumerate(sm_staff):
        ws.append([e['id'], e['name'], 'SME', 'Sales & Marketing', None,
                   'Proposal time reported under Sales & Marketing for all periods' if k == 0 else None])

    # ======================= Proposal Data -> =================================================
    wb.create_sheet('Proposal Data ->')

    # ======================= $ Amount / # Records (CRM report) ================================
    corp = [e for e in W.employees if e['office'] == 'Sales & Marketing' and e['status'] == 'Active']
    gen_by = (corp[0]['first'] + ' ' + corp[0]['last']) if corp else 'Report Owner'
    run_at = dt.datetime(as_of.year, as_of.month, as_of.day) + dt.timedelta(days=2, hours=9, minutes=41, seconds=12)
    fy0 = as_of.year - 1

    def crm_head(ws_, measure):
        ws_.cell(2, 2, 'New Pipeline Created by Month').font = Font(bold=True, size=14)
        ws_.cell(3, 2, 'As of %s Pacific Standard Time/PST • Generated by %s' % (run_at.strftime('%Y-%m-%d %H:%M:%S'), gen_by))
        ws_.cell(6, 2, 'Filtered By').font = BOLD
        ws_.cell(7, 2, 'Territories: All')
        ws_.cell(8, 2, 'Date Field: Created Date equals Current and Previous FY (1/1/%d to 12/31/%d)' % (fy0, as_of.year))
        ws_.cell(9, 2, 'Show: All opportunities')
        ws_.cell(10, 2, 'Opportunity Status: Any')
        ws_.cell(11, 2, 'Probability: All')
        ws_.cell(13, 2, 'Created Date →').font = BOLD
        ws_.cell(14, 2, 'Project Management Office  ↑').font = BOLD
        for i, m in enumerate(MONTHS):
            # the CRM export mixes text month labels with real dates once months are re-typed
            v = m.strftime('%B %Y') if i < 3 else (m.strftime('%B') if i < 5 else dtt(m))
            x = ws_.cell(13, 4 + i, v)
            if isinstance(v, dt.datetime):
                x.number_format = DATE_FMT
            ws_.cell(14, 4 + i, measure)

    for title, src, measure in (('$ Amount', amt, 'Sum of Amount'), ('# Records', cnt, 'Record Count')):
        ws = wb.create_sheet(title)
        crm_head(ws, measure)
        rr = 15
        tots = [0] * nm
        for (short, lab, reg) in crm_rows:
            ws.cell(rr, 1, short); ws.cell(rr, 2, lab); ws.cell(rr, 3, reg)
            for i in range(nm):
                v = src[lab][i]
                ws.cell(rr, 4 + i, r2(v) if isinstance(v, float) else v)
                tots[i] += v
            if title == '# Records':
                ws.cell(rr, 4 + nm, '#N/A' if lab == GLOBAL_LLC_PMO else lab)   # lookup helper column
            rr += 1
        ws.cell(rr, 2, 'Total').font = BOLD
        for i in range(nm):
            ws.cell(rr, 4 + i, r2(tots[i]) if isinstance(tots[i], float) else tots[i])
        ws.column_dimensions['B'].width = 30

    # ======================= New Pipeline Created by Month (wide CRM report) ==================
    ws = wb.create_sheet('New Pipeline Created by Month')
    allm = sorted({k for lab in pipe_cnt for k in pipe_cnt[lab]})
    ws.cell(2, 2, 'New Pipeline Created by Month').font = Font(bold=True, size=14)
    ws.cell(3, 2, 'As of %s Pacific Standard Time/PST • Generated by %s' % (run_at.strftime('%Y-%m-%d %H:%M:%S'), gen_by))
    ws.cell(6, 2, 'Filtered By').font = BOLD
    ws.cell(7, 2, 'Territories: All')
    ws.cell(8, 2, 'Date Field: Created Date equals Current and Previous 2 FY (1/1/%d to 12/31/%d)' % (fy0 - 1, as_of.year))
    ws.cell(9, 2, 'Show: All opportunities')
    ws.cell(10, 2, 'Opportunity Status: Any')
    ws.cell(11, 2, 'Probability: All')
    ws.cell(13, 2, 'Created Date →').font = BOLD
    ws.cell(14, 2, 'Project Management Office  ↑').font = BOLD
    for j, (y, m) in enumerate(allm):
        ws.cell(13, 3 + 2 * j, dt.date(y, m, 1).strftime('%B %Y'))
        ws.cell(14, 3 + 2 * j, 'Sum of Amount'); ws.cell(14, 4 + 2 * j, 'Record Count')
    tc = 3 + 2 * len(allm)
    ws.cell(13, tc, 'Total'); ws.cell(14, tc, 'Sum of Amount'); ws.cell(14, tc + 1, 'Record Count')
    rr = 15
    gt_a, gt_c = defaultdict(float), defaultdict(int)
    for (_short, lab, _reg) in crm_rows:
        ws.cell(rr, 1, lab); ws.cell(rr, 2, lab)
        for j, k in enumerate(allm):
            ws.cell(rr, 3 + 2 * j, r2(pipe_amt[lab][k])); ws.cell(rr, 4 + 2 * j, pipe_cnt[lab][k])
            gt_a[k] += pipe_amt[lab][k]; gt_c[k] += pipe_cnt[lab][k]
        ws.cell(rr, tc, r2(sum(pipe_amt[lab].values()))); ws.cell(rr, tc + 1, sum(pipe_cnt[lab].values()))
        rr += 1
    ws.cell(rr, 1, 'Total').font = BOLD; ws.cell(rr, 2, 'Total').font = BOLD
    for j, k in enumerate(allm):
        ws.cell(rr, 3 + 2 * j, r2(gt_a[k])); ws.cell(rr, 4 + 2 * j, gt_c[k])
    ws.cell(rr, tc, r2(sum(gt_a.values()))); ws.cell(rr, tc + 1, sum(gt_c.values()))

    os.makedirs(OUT_DIR, exist_ok=True)
    wb.save(OUT)
    print('wrote', OUT)
    print('  months: %s .. %s (%d)' % (MONTHS[0], MONTHS[-1], NM))
    print('  proposal-time rows: %d   hours: %.2f' % (len(prop_rows), sum(all_h)))
    print('  opportunities in window: %d   pipeline $: %.0f' % (sum(tot_n), sum(tot_a)))
    print('  hours / opportunity (grand): %.2f' % (sum(all_h) / sum(tot_n)))


if __name__ == '__main__':
    main()
