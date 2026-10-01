#!/usr/bin/env python3
"""gen_revenue_forecast.py — synthetic source for the Revenue Forecast dashboard.

Writes, from the shared demo world (tools/demo_world.py):

  demo_data/Weekly and Monthly Updates/9.29.2026/Revenue Forecast/Revenue Forecast.xlsx
      The ERP's "OTC-Project Profitability CM to PM (In USD)" export, single tab
      'OTC-Project Profitability CM to', in the export's layout: company name in A1, report
      title in A3, header rows 4-7 (measure / period / period-date / column caption), an
      "All Projects" group row on row 8, one row per project from row 9 sorted by status then
      Project ID, then the 'WIP Reserve Project', 'Total All Projects' and 'All Locations' rows.

  demo_data/.../Revenue Forecast/revenue_forecast_entries.csv
      The PMs' forecast entries in exactly the dashboard's CSV Export/Import format
      (Project ID, Name, PM, Location, Avg Rate, 3 x [Hrs, Cost, Exp], Exp Markup %).
      The pipeline embeds it as the dashboard's starting values (see
      pipeline/regenerate_revenue_forecast_data.py).

How the numbers are derived (nothing is drawn independently):
  * Contract, revised budget, revenue ITD, labor (Wages) and non-labor (Cost of Sales) ITD come
    straight from each project in the world, so ITD gross profit = revenue - COS - wages and the
    totals row is the sum of the project rows.
  * Month-to-date (Sep 1 - as-of) and prior-month (August) columns are the slices of those ITD
    figures earned in each period: wages = hours charged in the period x each person's burdened
    cost rate (1.5x for overtime); COS = the project's expense share of those wages; revenue =
    the increase in the world's own revenue formula (fixed fee: % complete on revised budget x
    contract; T&M: hours x bill rates, capped at contract) between the period's cut-off dates.
  * 'Observed percent' is the PM-reported % complete (fixed-fee projects mostly), equal to the
    revenue % complete the ERP uses; 'Auto Calculate %' flags the ones the ERP computes itself.
  * Forecast entries: hours per month = the project's FY2026 burn rate (hours per active week
    from the timesheets) x working weeks in the month up to the project's end date (overdue
    projects get a short close-out), eased for the November/December holidays, and capped so
    expected revenue never exceeds the remaining contract value (contract - revenue ITD).
    Avg Rate = the project's recent revenue per hour (as the pipeline suggests it) or its blended
    bill rate; cost = hours x the project's average cost rate; expenses = its expense share.

Usage:  python3 tools/gen_revenue_forecast.py
Deterministic: the only randomness is random.Random(f"{W.seed}-revenue-forecast").
"""
from __future__ import annotations

import calendar
import csv
import datetime as dt
import os
import random
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from demo_world import world, COMPANY, NON_POSTING, FIRST_WEEK_END  # noqa: E402

import openpyxl  # noqa: E402
from openpyxl.styles import Font  # noqa: E402

REPO = os.path.dirname(HERE)
OUT_DIR = os.path.join(REPO, 'demo_data', 'Weekly and Monthly Updates', '9.29.2026', 'Revenue Forecast')
XLSX = os.path.join(OUT_DIR, 'Revenue Forecast.xlsx')
ENTRIES = os.path.join(OUT_DIR, 'revenue_forecast_entries.csv')
SHEET = 'OTC-Project Profitability CM to'

# Statuses the pipeline keeps (it drops Completed, On Hold, Cost Tracking, Final Billing ...).
FORECAST_STATUSES = {'In Progress', 'Final Reporting/Review', 'Lab/Curation', 'Transferred Entity'}
# Forecast quarter (must match meta.months in the pipeline script).
FC_MONTHS = [(2026, 10), (2026, 11), (2026, 12)]
HOLIDAY_EASE = {10: 0.95, 11: 0.88, 12: 0.76}    # end of field season, Thanksgiving, year-end


def _d(x):
    return dt.datetime(x.year, x.month, x.day)


def _labor_by_cutoff(W, cutoffs):
    """Per project: cumulative FY labor $, billed $ (bill rate, excl. non-posting) and hours at
    each cut-off date, plus FY totals and first/last charge weeks."""
    cum = {c: (defaultdict(float), defaultdict(float), defaultdict(float)) for c in cutoffs}
    tot_h = defaultdict(float); tot_l = defaultdict(float); tot_b = defaultdict(float)
    weeks = defaultdict(set)
    ovhd = {c: [0.0, 0.0] for c in cutoffs}; ovhd_tot = [0.0, 0.0]
    for r in W.entries:
        e = W.emp[r['emp']]
        mult = 1.5 if r['tt'] == 'OT' else 1.0
        lab = r['hours'] * e['cost_rate'] * mult
        if r['project'] == 'OVHD':
            ovhd_tot[0] += lab; ovhd_tot[1] += r['hours']
            for c in cutoffs:
                if r['date'] <= c:
                    ovhd[c][0] += lab; ovhd[c][1] += r['hours']
            continue
        pid = r['project']
        bill = r['hours'] * e['bill_rate'] if r['task'] != NON_POSTING else 0.0
        tot_h[pid] += r['hours']; tot_l[pid] += lab; tot_b[pid] += bill
        weeks[pid].add(r['week'])
        for c in cutoffs:
            if r['date'] <= c:
                L, B, H = cum[c]
                L[pid] += lab; B[pid] += bill; H[pid] += r['hours']
    return cum, (tot_l, tot_b, tot_h), weeks, ovhd, ovhd_tot


def _prior_labor(p, fy_labor):
    """Pre-FY2026 labor exactly as the world computed it."""
    rb = p['revised_budget']
    if p['status'] in ('Completed', 'Final Billing/Invoice', 'Cost Tracking on Closed Project') and fy_labor == 0:
        return rb * (1 - p['expense_share']) * min(1.3, max(0.7, 1 + p['burn_bias']))
    return p['prior_share'] * rb * (1 - p['expense_share']) * (1 + p['burn_bias'] * 0.5)


def _rev_formula(p, prior, fy_labor, billed):
    """The world's revenue rule evaluated on cumulative charges (before any status true-up)."""
    if p['billing'] == 'Fixed fee':
        cost = (prior + fy_labor) / (1 - p['expense_share'])
        pct = min(1.0, cost / p['revised_budget']) if p['revised_budget'] else 0.0
        return p['contract'] * pct
    return min(p['contract'], p['prior_share'] * p['contract'] + billed)


def otc_rows(W):
    """Project rows of the OTC-Project Profitability tab as dicts (shared semantics:
    contract, revised budget, revenue/COS/wages/GP for MTD, prior month and ITD)."""
    as_of = W.as_of
    prev_end = as_of.replace(day=1) - dt.timedelta(days=1)          # Aug 31
    prev_start_cut = prev_end.replace(day=1) - dt.timedelta(days=1)  # Jul 31
    cutoffs = [prev_start_cut, prev_end, as_of]
    cum, (tot_l, tot_b, tot_h), weeks, ovhd, ovhd_tot = _labor_by_cutoff(W, cutoffs)
    rnd = random.Random('%s-revenue-forecast' % W.seed)
    rows = []
    for p in W.projects:
        pid = p['id']; es = p['expense_share']
        fy_l = tot_l[pid]
        prior = _prior_labor(p, fy_l)
        f = {c: _rev_formula(p, prior, cum[c][0][pid], cum[c][1][pid]) for c in cutoffs}
        # monthly revenue = increase in the revenue rule; ITD = the world's revenue (incl. any
        # status true-up, which therefore sits before the two reported months)
        rev_mtd = f[as_of] - f[prev_end]
        rev_prev = f[prev_end] - f[prev_start_cut]
        wages_mtd = cum[as_of][0][pid] - cum[prev_end][0][pid]
        wages_prev = cum[prev_end][0][pid] - cum[prev_start_cut][0][pid]
        hrs_mtd = cum[as_of][2][pid] - cum[prev_end][2][pid]
        hrs_prev = cum[prev_end][2][pid] - cum[prev_start_cut][2][pid]
        ratio = es / (1 - es)
        cos_mtd, cos_prev = wages_mtd * ratio, wages_prev * ratio
        rev_itd, cos_itd, wages_itd = p['revenue'], p['expense_cost'], p['labor_cost']
        # observed % complete (PM-reported) and the ERP's auto-calculate flag
        pct = round(min(1.0, p['rev_rec']), 3)
        if p['billing'] == 'Fixed fee':
            auto = rnd.random() < 0.26
            observed = pct if (auto or rnd.random() < 0.45) else 0.0
        else:
            auto = False
            observed = pct if rnd.random() < 0.2 else 0.0
        rows.append(dict(
            id=pid, name=p['name'], billing=p['billing'], status=p['status'],
            observed=observed, auto=auto,
            contract=p['contract'], revised_budget=p['revised_budget'],
            rev_mtd=rev_mtd, cos_mtd=cos_mtd, wages_mtd=wages_mtd, hrs_mtd=hrs_mtd,
            rev_prev=rev_prev, cos_prev=cos_prev, wages_prev=wages_prev, hrs_prev=hrs_prev,
            rev_itd=rev_itd, cos_itd=cos_itd, wages_itd=wages_itd,
            location=p['location'], pm=p['pm'], start=p['start'], end=p['end'],
            fy_hours=tot_h[pid], fy_labor=fy_l, fy_billed=tot_b[pid], weeks=weeks[pid]))
    # company overhead line (all OVHD timesheet time): wages only, no revenue
    wo_mtd = ovhd[as_of][0] - ovhd[prev_end][0]; ho_mtd = ovhd[as_of][1] - ovhd[prev_end][1]
    wo_prev = ovhd[prev_end][0] - ovhd[prev_start_cut][0]; ho_prev = ovhd[prev_end][1] - ovhd[prev_start_cut][1]
    rows.append(dict(id='OVHD', name='Overhead General', billing=None, status='In Progress',
                     observed=0.0, auto=False, contract=0.0, revised_budget=0.0,
                     rev_mtd=0.0, cos_mtd=0.0, wages_mtd=wo_mtd, hrs_mtd=ho_mtd,
                     rev_prev=0.0, cos_prev=0.0, wages_prev=wo_prev, hrs_prev=ho_prev,
                     rev_itd=0.0, cos_itd=0.0, wages_itd=ovhd_tot[0],
                     location='Corporate', pm=None, start=None, end=None,
                     fy_hours=ovhd_tot[1], fy_labor=ovhd_tot[0], fy_billed=0.0, weeks=set()))
    rows.sort(key=lambda r: (r['status'], r['id']))
    return rows, prev_end


def write_xlsx(W, rows, prev_end):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = SHEET
    ws['A1'] = COMPANY
    ws['A3'] = 'OTC-Project Profitability CM to PM (In USD)'
    as_of = _d(W.as_of); pe = _d(prev_end)
    # header block: row 4 measure, row 5 period, row 6 period date, row 7 column caption
    hdr = {  # col: (row4, row5, row6 date, row7)
        8: ('Contract Value', 'Inception To Date', as_of, 'Actual'),
        9: ('Total Revised Budget', 'ITD As of', as_of, 'Actual'),
        11: ('Revenue Recognized', 'Month To Date', as_of, 'Actual MTD Revenue'),
        12: ('Cost of Sales', 'Month To Date', as_of, 'Actual MTD COS'),
        13: ('Wages', 'Month To Date', as_of, 'Actual MTD Wages'),
        14: ('Total Project Hours', 'Month To Date', as_of, 'Actual MTD Hrs'),
        15: ('Project Gross Profit', 'Month To Date', as_of, 'Actual MTD GP'),
        17: ('Revenue Recognized', 'Month Ending', pe, 'Actual'),
        18: ('Cost of Sales', 'Month Ending', pe, 'Actual'),
        19: ('Wages', 'Month Ending', pe, 'Actual'),
        20: ('Total Project Hours', 'Month Ending', pe, 'Actual'),
        21: ('Project Gross Profit', 'Month Ending', pe, 'Actual'),
        23: ('Revenue Recognized', 'ITD', as_of, 'Actual'),
        24: ('Cost of Sales', 'ITD', as_of, 'Actual'),
        25: ('Wages', 'ITD', as_of, 'Actual'),
        26: ('Project Gross Profit', 'ITD', as_of, 'Actual'),
    }
    for c, (a, b, d, cap) in hdr.items():
        ws.cell(4, c, a); ws.cell(5, c, b)
        cell = ws.cell(6, c, d); cell.number_format = 'mm-dd-yy'
        ws.cell(7, c, cap)
    for c, cap in {1: 'Project ID', 2: 'Project Name', 3: 'Project Billing type', 4: 'Project status',
                   6: 'Observed percent ...', 7: 'Summary', 27: 'Location Name',
                   28: 'Auto Calculate % ...', 29: 'Project Manager Name', 30: 'Project Begin date',
                   31: 'Project End date'}.items():
        ws.cell(7, c, cap)
    ws.cell(8, 1, 'All Projects'); ws.cell(8, 2, 'All Projects')

    def i(v):
        return int(round(v))

    def put(r, x):
        ws.cell(r, 1, x['id']); ws.cell(r, 2, x['name'])
        if x['billing']: ws.cell(r, 3, x['billing'])
        if x['status']: ws.cell(r, 4, x['status'])
        ob = x['observed']
        ws.cell(r, 6, ob if ob else 0); ws.cell(r, 7, round(ob * 100, 1) if ob else 0)
        ws.cell(r, 8, i(x['contract'])); ws.cell(r, 9, i(x['revised_budget']))
        vals = {11: x['rev_mtd'], 12: x['cos_mtd'], 13: x['wages_mtd'], 14: x['hrs_mtd'],
                17: x['rev_prev'], 18: x['cos_prev'], 19: x['wages_prev'], 20: x['hrs_prev'],
                23: x['rev_itd'], 24: x['cos_itd'], 25: x['wages_itd']}
        for c, v in vals.items():
            ws.cell(r, c, i(v))
        # GP from the rounded components so each row foots exactly
        ws.cell(r, 15, i(x['rev_mtd']) - i(x['cos_mtd']) - i(x['wages_mtd']))
        ws.cell(r, 21, i(x['rev_prev']) - i(x['cos_prev']) - i(x['wages_prev']))
        ws.cell(r, 26, i(x['rev_itd']) - i(x['cos_itd']) - i(x['wages_itd']))
        if x['location']: ws.cell(r, 27, x['location'])
        if x['auto'] is not None: ws.cell(r, 28, bool(x['auto']))
        if x['pm']: ws.cell(r, 29, x['pm'])
        for c, d in ((30, x['start']), (31, x['end'])):
            if d:
                ws.cell(r, c, _d(d)).number_format = 'mm-dd-yy'

    r = 9
    for x in rows:
        put(r, x); r += 1
    # unallocated WIP reserve (a GL revenue reserve, sized off the active book's unbilled revenue)
    unbilled = sum(p['revenue'] - p['invoiced'] for p in W.projects if p['status'] in FORECAST_STATUSES)
    reserve = -round(unbilled * 0.12, -3)
    wip = dict(id='WIP Reserve Project', name='WIP Reserve Project', billing=None, status=None, observed=0.0,
               auto=False, contract=0, revised_budget=0, rev_mtd=0, cos_mtd=0, wages_mtd=0, hrs_mtd=0,
               rev_prev=0, cos_prev=0, wages_prev=0, hrs_prev=0, rev_itd=reserve, cos_itd=0, wages_itd=0,
               location=None, pm=None, start=None, end=None)
    put(r, wip); r += 1
    # totals: sums of the rows above (columns 8..26)
    first, last = 9, r - 1
    for label, fpct in (('Total All Projects', True), ('All Locations', False)):
        ws.cell(r, 1, label); ws.cell(r, 2, label)
        if fpct:
            n = sum(1 for x in rows if x['observed'])
            avg = sum(x['observed'] for x in rows) / max(1, n)
            ws.cell(r, 6, round(avg, 4)); ws.cell(r, 7, round(avg * 100, 2))
        else:
            ws.cell(r, 7, 0)
        for c in (8, 9, 11, 12, 13, 14, 15, 17, 18, 19, 20, 21, 23, 24, 25, 26):
            ws.cell(r, c, sum(ws.cell(k, c).value or 0 for k in range(first, last + 1)))
        r += 1
    for c in ('A1', 'A3'):
        ws[c].font = Font(bold=True, size=11)
    return wb


def _workdays(a, b):
    """Mon-Fri days in [a, b]."""
    n, d = 0, a
    while d <= b:
        if d.weekday() < 5: n += 1
        d += dt.timedelta(days=1)
    return n


def forecast_entries(W, rows):
    """PM forecast entries for the active book, consistent with burn rate, end date and backlog."""
    rnd = random.Random('%s-revenue-forecast-entries' % W.seed)
    as_of = W.as_of
    fy_start = FIRST_WEEK_END
    out = []
    for x in rows:
        if x['status'] not in FORECAST_STATUSES or x['id'] == 'OVHD':
            continue
        p = W.project[x['id']]
        remaining = x['contract'] - round(x['rev_itd'])
        # burn rate: FY hours per active week (weeks from the first week the job was live in FY
        # through the as-of week, or its end date if that came first)
        first = max(fy_start, p['start'])
        last = min(as_of, p['end'] + dt.timedelta(days=45))
        act_weeks = max(1.0, (last - first).days / 7.0) if last > first else 1.0
        burn = x['fy_hours'] / act_weeks if x['fy_hours'] > 0 else 0.0
        # recent pace matters too: blend in the last two months' hours per week
        prev_start = (as_of.replace(day=1) - dt.timedelta(days=1)).replace(day=1)   # Aug 1
        recent = (x['hrs_mtd'] + x['hrs_prev']) / ((as_of - prev_start).days / 7.0)
        if burn > 0:
            burn = 0.35 * burn + 0.65 * recent          # PMs anchor on the current pace
        # rate: recent revenue per hour (the pipeline's suggestion) or the blended bill rate
        hr = x['hrs_mtd'] if x['hrs_mtd'] > 0 else x['hrs_prev']
        rv = x['rev_mtd'] if x['hrs_mtd'] > 0 else x['rev_prev']
        rate = rv / hr if hr > 0 and rv > 0 else (x['fy_billed'] / x['fy_hours'] if x['fy_hours'] > 0 else 0.0)
        cost_rate = x['fy_labor'] / x['fy_hours'] if x['fy_hours'] > 0 else 0.0
        es = p['expense_share']
        if p['billing'] == 'Fixed fee':
            # fixed-fee revenue per hour includes the expense recovery; expenses are entered
            # separately, so the labor rate is the remainder
            rate = max(0.0, rate - cost_rate * es / (1 - es))
        rate = int(round(rate / 5.0) * 5)
        mk = rnd.choice([10, 10, 12, 15]) if p['billing'] == 'Time and material' else 0
        # the window a PM would forecast: to the end date; overdue jobs get a ~5-week close-out
        eff_end = p['end'] if p['end'] > as_of else as_of + dt.timedelta(days=35)
        if p['status'] in ('Final Reporting/Review', 'Lab/Curation'):
            eff_end = min(eff_end, as_of + dt.timedelta(days=75))    # reporting tail winds down
        cap = max(0.0, remaining)
        months = []
        for (y, m) in FC_MONTHS:
            a = dt.date(y, m, 1); b = dt.date(y, m, calendar.monthrange(y, m)[1])
            wk = _workdays(max(a, as_of + dt.timedelta(days=1)), min(b, eff_end)) / 5.0 if eff_end >= a else 0.0
            h = int(round(burn * wk * HOLIDAY_EASE[m]))

            def vals(h):
                c = int(round(h * cost_rate))
                xexp = int(round(h * cost_rate * es / (1 - es) / 50.0) * 50)
                return c, xexp, h * rate + xexp * (1 + mk / 100.0)
            c, xexp, rev = vals(h)
            if rev > cap and h > 0:                         # never forecast past the backlog
                per_h = rev / h
                h = int(cap // per_h) if per_h > 0 else 0
                c, xexp, rev = vals(h)
                while h > 0 and rev > cap:                  # expense rounding can tip it over
                    h -= 1; c, xexp, rev = vals(h)
            if h <= 0:
                months.append((0, 0, 0)); continue
            cap -= rev
            months.append((h, c, xexp))
        if not any(mo[0] for mo in months):
            continue
        out.append(dict(id=x['id'], name=x['name'], pm=x['pm'], loc=x['location'], rate=rate, mk=mk, m=months))
    return out


def write_entries(entries):
    names = [calendar.month_name[m] for _, m in FC_MONTHS]
    head = ['Project ID', 'Project Name', 'PM', 'Location', 'Avg Rate']
    for n in names:
        head += ['%s Hrs' % n, '%s Cost' % n, '%s Exp' % n]
    head.append('Exp Markup %')
    with open(ENTRIES, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f, lineterminator='\n')
        w.writerow(head)
        for e in entries:
            row = [e['id'], e['name'], e['pm'], e['loc'], e['rate']]
            for h, c, x in e['m']:
                row += [h, c, x]
            row.append(e['mk'])
            w.writerow(row)


def main():
    W = world()
    os.makedirs(OUT_DIR, exist_ok=True)
    rows, prev_end = otc_rows(W)
    wb = write_xlsx(W, rows, prev_end)
    wb.save(XLSX)
    entries = forecast_entries(W, rows)
    write_entries(entries)
    act = [x for x in rows if x['status'] in FORECAST_STATUSES and x['id'] != 'OVHD']
    print('wrote', os.path.relpath(XLSX, REPO), '(%d project rows)' % len(rows))
    print('wrote', os.path.relpath(ENTRIES, REPO), '(%d forecast entries)' % len(entries))
    print('active book: n=%d contract=$%.1fM revITD=$%.1fM' % (
        len(act), sum(x['contract'] for x in act) / 1e6, sum(x['rev_itd'] for x in act) / 1e6))


if __name__ == '__main__':
    main()
