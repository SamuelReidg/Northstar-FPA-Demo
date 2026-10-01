#!/usr/bin/env python3
"""gen_ews.py -- synthetic "SR of Projects @ Risk - Automated <date>.xlsx" for the demo.

Feeds the Early Warning System dashboard, its Projects tab (Recovery Builder) and the
Task Budget Detail panel (pipeline/regenerate_ews_data.py + pipeline/regenerate_project_tasks.py).

Writes the automated workbook in the ERP's post-08/2026 export layout, CALCULATED (every cell the
pipeline reads holds a value -- the real workbook is formula-driven and must be recalculated
before the scripts can read it; this file is the already-calculated equivalent):

  'Project Details '                 (trailing space) title rows 1-3, header row 4, one row per project
                                     from row 5 -- the columns the real sheet computes with SUMIFs over
                                     the task tab and VLOOKUPs into the OTC and _Lookup tabs.
  'Budget to Actual Costs by Task'   ERP report banner rows 1-11, header row 12 ("Task Project ID"...),
                                     "All Tasks" label row 13, task rows from 14, "Total All" row,
                                     "Created on" footer. Money columns sit behind the spacer column
                                     (Contract Value G, Revised Budget I, Total Costs N, Budgeted Hours O,
                                     Actual Hours P); no Task Overrun columns (the export dropped them).
  'OTC All reports keys account in'  ERP banner rows 1-10, header row 11, "All Locations" /
                                     "All Customers" / "All Projects" subtotal rows, project rows,
                                     "Total All Projects". Location and PM arrive as IDs.
  '_Calc' (hidden)                   per-task overrun $ / hours, row-aligned with the task tab, plus
                                     the layout self-checks in column G.
  'How to Refresh'                   the operator notes tab.
  '_Lookup' (hidden)                 Location ID -> Location Name, PM ID -> Project Manager Name.

Everything is derived from tools/demo_world.py -- nothing is drawn independently:
  * project list, status, billing, PM, office/location, contract and budgets come from W.projects;
  * task lines are the project's p['tasks'] plus an "OVHD-Non-Posting - Overhead Non-Posting" line
    when time was charged to non-posting in W.entries;
  * task COSTS split p['cost_to_date'] by the labor $ actually charged to each task in W.entries
    (hours x each person's burdened cost rate, OT x1.5); pre-FY2026 labor and expenses are spread
    by the task plan -- so task costs sum exactly to the project's cost to date;
  * task BUDGETS split p['revised_budget'] (and original budget / contract) by a task plan built from
    the hours charged to each task, perturbed by a seeded estimating error -- they sum exactly to the
    project budget; non-posting lines carry no budget (their whole cost is overrun, the ERP rule);
  * ACTUAL HOURS = FY2026 hours charged per task (W.entries) + pre-FY2026 hours implied by the
    project's prior labor at its blended cost rate; BUDGETED HOURS = the task's labor budget at a
    planned blended rate (a seeded estimate of the actual staff mix); some projects budget in $ only;
  * overruns follow the _Calc rule exactly, so cost / hours %, the >5% flag, % complete and the EWS
    tiers are all consequences of those numbers.

usage: python3 tools/gen_ews.py [out.xlsx]
"""
import os, sys, math, random, datetime as dt
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from demo_world import world, NON_POSTING, COMPANY_SHORT  # noqa: E402

import openpyxl  # noqa: E402
from openpyxl.styles import Font, PatternFill, Alignment  # noqa: E402

ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, 'demo_data', 'Weekly and Monthly Updates', '9.29.2026', 'Early Warning System',
                   'SR of Projects @ Risk - Automated 9.29.26.xlsx')

PD, T2, OTC, CALC, HOWTO, LOOK = ('Project Details ', 'Budget to Actual Costs by Task',
                                  'OTC All reports keys account in', '_Calc', 'How to Refresh', '_Lookup')
OVHD_NP_LINE = 'OVHD-Non-Posting - Overhead Non-Posting'
T2_HDR_ROW, OTC_HDR_ROW = 12, 11
MONTH_END = dt.date(2026, 9, 30)

# typical share of effort by task type (used for the plan of projects with no FY2026 time yet,
# and as the floor for tasks nobody has charged to so far)
TASK_TYPE_W = {'Prefield': 0.7, "Fieldwork": 3.0, 'Reporting': 1.8, 'Project Management': 0.9,
               'Background Research': 0.6, 'Monitoring': 2.5, 'Lab Analysis': 1.2, 'Site Recordation': 1.2,
               'GIS & Mapping': 0.9, 'Draft Report': 1.6, 'Final Report': 0.8, 'Agency Consultation': 0.5,
               'Curation': 0.6, 'Data Recovery Excavation': 3.0, 'Testing': 1.8, 'Survey': 2.5,
               'Desktop Review': 0.5}


def first_last(name):
    """'Last, First' -> 'First Last' (how the ERP's PM name reads)."""
    if ', ' in name:
        last, first = name.split(', ', 1)
        return '%s %s' % (first, last)
    return name


def split_exact(total, weights, nd=2):
    """Split `total` by weights, rounded to nd decimals, residual on the largest share (sums exactly)."""
    s = sum(weights)
    if not weights:
        return []
    if s <= 0:
        weights = [1.0] * len(weights); s = float(len(weights))
    parts = [round(total * w / s, nd) for w in weights]
    k = max(range(len(parts)), key=lambda i: weights[i])
    parts[k] = round(total - sum(parts) + parts[k], nd)
    return parts


def build_rows(W):
    rnd = random.Random('%s-ews' % W.seed)
    emp_by_name = {e['name']: e for e in W.employees}

    # ---- roll the timesheets up by project x task --------------------------------------------
    t_hours = defaultdict(lambda: defaultdict(float))     # pid -> task -> hours (FY2026)
    t_labor = defaultdict(lambda: defaultdict(float))     # pid -> task -> labor $ (hours x cost rate)
    sep_hours = defaultdict(lambda: [0.0, 0.0])           # pid -> [posting, non-posting] hours in Sept
    office_rate = defaultdict(lambda: [0.0, 0.0])
    for r in W.entries:
        pid = r['project']
        if pid == 'OVHD':
            continue
        e = W.emp[r['emp']]
        mult = 1.5 if r['tt'] == 'OT' else 1.0
        t_hours[pid][r['task']] += r['hours']
        t_labor[pid][r['task']] += r['hours'] * e['cost_rate'] * mult
        office_rate[e['office']][0] += r['hours'] * e['cost_rate'] * mult
        office_rate[e['office']][1] += r['hours']
        if r['date'].month == 9:
            sep_hours[pid][1 if r['task'] == NON_POSTING else 0] += r['hours']
    orate = {o: (v[0] / v[1] if v[1] else 60.0) for o, v in office_rate.items()}

    projects = sorted(W.projects, key=lambda p: p['id'])
    task_rows, proj_rows = [], []
    code = [1000]
    for p in projects:
        pid, rb = p['id'], p['revised_budget']
        hrs, lab = t_hours.get(pid, {}), t_labor.get(pid, {})
        fy_hours = sum(hrs.values())
        fy_labor = sum(lab.values())
        rate = fy_labor / fy_hours if fy_hours > 0 else orate.get(p['office'], 60.0)
        tasks = list(p['tasks'])

        # task plan: what each task was estimated to take (hours share, seeded estimating error)
        tot_h = sum(max(0.0, hrs.get(t, 0.0)) for t in tasks)
        tw = [TASK_TYPE_W.get(t, 1.0) for t in tasks]
        tw = [w / sum(tw) for w in tw]
        if tot_h > 0:
            sigma = 0.55 + 0.45 * rnd.random()
            plan = [max(0.0, hrs.get(t, 0.0)) / tot_h * min(2.2, max(0.4, math.exp(rnd.gauss(-0.03, sigma)))) + 0.05 * w
                    for t, w in zip(tasks, tw)]
        else:
            plan = [w * math.exp(rnd.gauss(0, 0.2)) for w in tw]
        s = sum(plan); plan = [x / s for x in plan]

        # costs: FY labor by task from the timesheets; prior-year labor + expenses spread by plan/labor
        prior_labor = max(0.0, p['labor_cost'] - fy_labor)
        prior_t = split_exact(prior_labor, plan) if tasks else []
        labor_t = [lab.get(t, 0.0) + pr for t, pr in zip(tasks, prior_t)]
        np_labor = lab.get(NON_POSTING, 0.0)
        np_hours = hrs.get(NON_POSTING, 0.0)
        has_np = abs(np_hours) > 0
        labor_all = labor_t + ([np_labor] if has_np else [])
        # the non-posting line carries labor only; expenses follow the billable task labor
        exp_t = split_exact(p['expense_cost'], [max(0.0, x) for x in labor_t] or [1.0]) + ([0.0] if has_np else [])
        cost_t = [a + b for a, b in zip(labor_all, exp_t)]
        # force the lines to sum exactly to the project cost to date (cent rounding)
        if cost_t:
            k = max(range(len(tasks)), key=lambda i: cost_t[i])
            cost_t[k] += round(p['cost_to_date'] - sum(cost_t), 2)
            labor_all[k] = cost_t[k] - exp_t[k]

        # hours: FY hours by task + prior hours implied by prior labor at the blended rate
        # (pre-FY2026 hours at the task's own mix of staff -- field tasks run cheaper than reporting)
        act_h = [hrs.get(t, 0.0) + pr / (rate * math.exp(rnd.gauss(0.0, 0.10))) for t, pr in zip(tasks, prior_t)] \
            + ([np_hours] if has_np else [])

        # budgets by plan; non-posting line has none
        rev_t = split_exact(rb, plan) + ([0.0] if has_np else [])
        org_t = split_exact(p['orig_budget'], plan) + ([0.0] if has_np else [])
        con_t = split_exact(p['contract'], plan, 0) + ([0.0] if has_np else [])
        hours_budgeted = rnd.random() > 0.18                      # some projects budget in $ only
        plan_rate = rate * math.exp(rnd.gauss(0.15, 0.22))       # planned staff mix (bids assume a richer, faster crew)
        labor_budget = rb * (1 - p['expense_share'])
        # each task's hours are estimated separately (crew size / duration assumptions differ by task)
        bud_h = ([round(labor_budget * w / plan_rate * math.exp(rnd.gauss(0.0, 0.12))) for w in plan]
                 if hours_budgeted else [0] * len(tasks)) \
            + ([0] if has_np else [])

        status = {'In Progress': 'In progress', 'On Hold': 'On hold'}.get(p['status'], 'Completed')
        names = tasks + ([None] if has_np else [])
        lines = []
        for i, t in enumerate(names):
            code[0] += rnd.randint(1, 7)
            label = OVHD_NP_LINE if t is None else 'NHT-%06d - %s' % (code[0], t)
            tstat = status
            if status == 'In progress' and t is not None and rev_t[i] > 0 and cost_t[i] >= rev_t[i] * 0.97 \
                    and rnd.random() < 0.5:
                tstat = 'Completed'
            N = round(cost_t[i], 2)
            K = round(exp_t[i], 2)
            L = round(N - K, 2)
            I = round(rev_t[i], 2)
            O = int(bud_h[i]); P = round(act_h[i], 2)
            # the _Calc rule, verbatim
            if I > 0:
                ovr = max(0.0, N - I)
            elif t is None:
                ovr = N
            else:
                ovr = 0.0
            if O > 0:
                ovh = max(0.0, P - O)
            elif t is None:
                ovh = P
            else:
                ovh = 0.0
            lines.append(dict(label='    ' + label, pid=pid, status=tstat, G=con_t[i], H=round(org_t[i], 2),
                              I=I, K=K, L=L, M=0, N=N, O=O, P=P, ovr=round(ovr, 2), ovh=round(ovh, 2)))
        task_rows.extend(lines)

        tot = lambda k: sum(l[k] for l in lines)
        H_ = tot('I'); J_ = tot('ovr'); N_ = tot('O'); P_ = tot('ovh')
        pm = emp_by_name.get(p['pm'])
        proj_rows.append(dict(
            p=p, pm_id=pm['sage_id'] if pm else '', pm_name=first_last(p['pm']),
            contract=tot('G'), revised=H_, cost=tot('N'), ovr=J_,
            ovr_pct=(J_ / H_ if H_ > 0 else 0), gt5=(1 if (H_ > 0 and J_ / H_ > 0.05) else 0),
            pct_complete=(0 if H_ == 0 else (tot('N') - J_) / H_),
            bud_h=N_, act_h=tot('P'), ovr_h=P_, ovr_h_pct=(P_ / N_ if N_ > 0 else 0),
            sep=sep_hours.get(pid, [0.0, 0.0]), exp=sum(l['K'] for l in lines)))
    return task_rows, proj_rows


def recent_billing(W):
    """Bill-rate value of time charged in the last 30 / 60 days, per project (drives AR and cash)."""
    b30, b60 = defaultdict(float), defaultdict(float)
    for r in W.entries:
        if r['project'] == 'OVHD' or r['task'] == NON_POSTING:
            continue
        age = (W.as_of - r['date']).days
        v = r['hours'] * W.emp[r['emp']]['bill_rate']
        if age <= 30: b30[r['project']] += v
        if age <= 60: b60[r['project']] += v
    return b30, b60


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else OUT
    W = world()
    rnd = random.Random('%s-ews-otc' % W.seed)
    task_rows, proj_rows = build_rows(W)
    asof = W.as_of
    asof_s = asof.strftime('%m/%d/%Y')
    asof_long = asof.strftime('%B %-d, %Y')
    bold = Font(bold=True)
    hdr_fill = PatternFill('solid', fgColor='DCE6F1')
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ------------------------------------------------------------------ Project Details
    ws = wb.create_sheet(PD)
    n_over = sum(r['gt5'] for r in proj_rows)
    ws['A1'] = 'Projects with Task Overruns >5%'; ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = 'As of ' + asof_long
    ws['A3'] = 'Layout check OK  |  %s projects  |  %s over 5%%' % (format(len(proj_rows), ','), format(n_over, ','))
    ws['A3'].font = Font(bold=True, color='2E7D32')
    heads = ['Project ID', 'Project Name', 'Billing Type', 'Project Status', 'Location Name ', 'Project Manager Name',
             'Contract Value', 'Revised Budget ', 'Total Cost ITD', 'Total Task Overrun ', 'Task Overrun %', '> 5%',
             'Cost % Complete', 'Budgeted Hours', 'Actual Hours', 'Task Overrun Hours', 'Hours Overrun %']
    for j, h in enumerate(heads, 1):
        c = ws.cell(4, j, h); c.font = bold; c.fill = hdr_fill; c.alignment = Alignment(wrap_text=True)
    # Project Details lists projects in first-appearance order on the task tab (_Calc idx)
    for i, r in enumerate(proj_rows):
        p = r['p']
        row = [p['id'], '    ' + p['name'], p['billing'], p['status'], p['location'], r['pm_name'],
               round(r['contract'], 2), round(r['revised'], 2), round(r['cost'], 2), round(r['ovr'], 2),
               r['ovr_pct'], r['gt5'], r['pct_complete'], r['bud_h'], round(r['act_h'], 2), round(r['ovr_h'], 2),
               r['ovr_h_pct']]
        for j, v in enumerate(row, 1):
            c = ws.cell(5 + i, j, v)
            if j in (7, 8, 9, 10): c.number_format = '#,##0.00'
            elif j in (11, 13, 17): c.number_format = '0.0%'
            elif j in (14, 15, 16): c.number_format = '#,##0.0'
    ws.freeze_panes = 'C5'
    for col, w in zip('ABCDEFGHIJKLMNOPQ', [14, 46, 16, 22, 20, 22, 14, 14, 14, 14, 11, 6, 11, 12, 12, 12, 11]):
        ws.column_dimensions[col].width = w

    # ------------------------------------------------------------------ Budget to Actual Costs by Task
    ws = wb.create_sheet(T2)
    banner = [[COMPANY_SHORT], ['Budget to Actual Costs by Task (In USD)'], ['As of ' + asof_long],
              ['Reporting Book:', 'ACCRUAL'], ['As of Date:', asof_s], ['Location:', COMPANY_SHORT], [' ']]
    for i, row in enumerate(banner, 1):
        for j, v in enumerate(row, 1):
            ws.cell(i, j, v)
    ws['A1'].font = Font(bold=True, size=13)
    B = ' '
    r8 = [B, B, None, None, None, B, 'Bookings', 'Original Budget', 'Total Revised Budget', B, 'Total Costs',
          None, None, None, 'Budgeted Project Hours', 'Total Hours', B]
    r9 = [B, B, None, None, None, B, B, B, B, B, 'Cost of Sales', 'Wages', 'Remaining Commit', 'Total Costs', B, B, B]
    r10 = [B, B, None, None, None, B] + ['Inception To Date'] * 3 + [B] + ['Inception To Date'] * 6 + [B]
    r11 = [B, B, None, None, None, B] + [asof_s] * 3 + [B] + [asof_s] * 6 + [B]
    r12 = [B, 'Task Project ID', 'Task status', 'Location Name', 'Project Manager Name', B, 'Contract Value',
           'Original Budget', 'Revised Budget', B, 'Actual', 'Actual', 'Actual', 'Actual', 'Budgeted Hours',
           'Actual Hours', B]
    r13 = ['  All Tasks'] + ['  '] * 16
    for i, row in zip(range(8, 14), [r8, r9, r10, r11, r12, r13]):
        for j, v in enumerate(row, 1):
            if v is not None:
                c = ws.cell(i, j, v)
                if i in (8, 9, 12): c.font = bold
    for i, t in enumerate(task_rows):
        rr = 14 + i
        vals = [t['label'], t['pid'], t['status'], B, B, B, t['G'], t['H'], t['I'], B, t['K'], t['L'], t['M'],
                t['N'], t['O'], t['P'], B]
        for j, v in enumerate(vals, 1):
            c = ws.cell(rr, j, v)
            if j in (7, 8, 9, 11, 12, 13, 14): c.number_format = '#,##0.00'
    last = 14 + len(task_rows)
    tot = lambda k: round(sum(t[k] for t in task_rows), 2)
    for j, v in enumerate(['  Total All ', B, B, B, B, B, tot('G'), tot('H'), tot('I'), B, tot('K'), tot('L'), 0,
                           tot('N'), tot('O'), tot('P'), B], 1):
        c = ws.cell(last, j, v); c.font = bold
        if j >= 7 and isinstance(v, (int, float)): c.number_format = '#,##0.00'
    ws.cell(last + 1, 1, ' '); ws.cell(last + 2, 1, ' ')
    ws.cell(last + 3, 1, 'Created on : %s 06:42:18 AM' % (asof + dt.timedelta(days=3)).strftime('%m/%d/%Y'))
    ws.column_dimensions['A'].width = 52; ws.column_dimensions['B'].width = 14
    t2_rows = last + 3

    # ------------------------------------------------------------------ OTC All reports keys account in
    ws = wb.create_sheet(OTC)
    b30, b60 = recent_billing(W)
    for i, row in enumerate([[COMPANY_SHORT], ['OTC All reports keys account information (In USD)'],
                             ['Reporting Book:', 'ACCRUAL'], ['As of Date:', asof_s], ['Location:', COMPANY_SHORT],
                             [' ']], 1):
        for j, v in enumerate(row, 1):
            ws.cell(i, j, v)
    ws['A1'].font = Font(bold=True, size=13)
    me = MONTH_END.strftime('%m/%d/%Y')
    o7 = [B, B, B, None, None, None, None, B, 'Contract Value', 'Invoiced Amount', 'Total Revised Budget',
          'Accounts Receivable', 'Cash and Cash Equivalents', B, 'Revenues', 'Cost of Sales', 'Wages', B,
          'Total Delayed Revenue', None, None, None, 'OTC Cycle Walkforward Validation', B, 'Total Billable Hours',
          'Total Non-Billable Hours', 'Budgeted Hours', B]
    o8 = [B, B, B, None, None, None, None] + [B] * 11 + ['Unbilled Receivable', 'WIP', 'Deferred Revenue',
                                                         'Total Delayed Revenue'] + [B] * 6
    o9 = [B, B, B, None, None, None, None, B, 'ITD As of', 'ITD As of', 'ITD As of', 'ITD As of', 'Month To Date', B,
          'ITD As of', 'ITD As of', 'Month To Date', B, 'ITD As of', 'ITD As of', 'ITD As of', 'ITD As of',
          'ITD As of', B, 'Month Ending', 'Month Ending', 'Month Ending', B]
    o10 = [B, B, B, None, None, None, None, B] + [asof_s] * 5 + [B] + [asof_s] * 3 + [B] + [asof_s] * 5 + [B] + \
        [me] * 3 + [B]
    o11 = ['Project ID', B, 'Billing type', 'Project status', 'Location ID', 'Project manager ID', 'Customer Name',
           '% Complete', 'Actual', 'Actual', 'Actual', 'Actual', 'Actual', B, 'Actual', 'Actual', 'Actual', B,
           'Actual', 'Actual', 'Actual', 'Actual', 'Actual', B, 'Actual', 'Actual', 'Actual', 'Project Customer ID',
           'Salesforce ID', 'Auto Calculate % ...']
    for i, row in zip(range(7, 12), [o7, o8, o9, o10, o11]):
        for j, v in enumerate(row, 1):
            if v is not None:
                c = ws.cell(i, j, v)
                if i in (7, 8, 11): c.font = bold
    cust_id = {c: 10000 + 37 * k + rnd.randint(0, 30) for k, c in enumerate(W.clients)}
    alnum = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
    otc_rows = []
    for r in proj_rows:
        p = r['p']
        rev, inv = p['revenue'], p['invoiced']
        ar = round(min(inv, b60.get(p['id'], 0.0) * rnd.uniform(0.25, 0.7)), 2)
        cash = round(min(max(0.0, inv - ar), b60.get(p['id'], 0.0) * rnd.uniform(0.15, 0.45)), 2)
        wages_mtd = 0.0                     # filled from the September timesheets below
        wip, dfr = round(max(0.0, rev - inv), 2), round(max(0.0, inv - rev), 2)
        pct = 1.0 if p['status'] in ('Completed', 'Final Billing/Invoice') else \
            (round(min(1.0, p['budget_spent']), 4) if p['billing'] == 'Fixed fee' else 0)
        crm = '006' + ''.join(rnd.choice(alnum) for _ in range(15))
        otc_rows.append([p['id'], '    ' + p['name'], p['billing'], p['status'], W.office[p['office']]['location_id'],
                         r['pm_id'], p['client'], pct, p['contract'], inv, p['revised_budget'], ar, cash, B, rev,
                         round(r['exp'], 2), wages_mtd, B, 0, wip, dfr, round(wip + dfr, 2), 0, B,
                         round(r['sep'][0], 2), round(r['sep'][1], 2), r['bud_h'], str(cust_id[p['client']]), crm,
                         p['billing'] == 'Fixed fee' and rnd.random() < 0.3])
    # Wages MTD: September labor $ per project from the timesheets
    sep_wages = defaultdict(float)
    for e_ in W.entries:
        if e_['project'] != 'OVHD' and e_['date'].month == 9:
            sep_wages[e_['project']] += e_['hours'] * W.emp[e_['emp']]['cost_rate'] * (1.5 if e_['tt'] == 'OT' else 1.0)
    for row in otc_rows:
        row[16] = round(sep_wages.get(row[0], 0.0), 2)
    numcols = [8, 9, 10, 11, 12, 14, 15, 16, 18, 19, 20, 21, 22, 24, 25, 26]
    sums = {j: round(sum(row[j] for row in otc_rows), 2) for j in numcols}
    sub = lambda label: [B, label, B, B, B, B, B, 0] + [sums.get(j, B) for j in range(8, 27)] + [B, B, B]
    for i, row in enumerate([sub('  All Locations'), sub('  All Customers'), [B + ' ', '  All Projects'] + ['  '] * 28]):
        for j, v in enumerate(row, 1):
            ws.cell(12 + i, j, v)
    for i, row in enumerate(otc_rows):
        for j, v in enumerate(row, 1):
            c = ws.cell(15 + i, j, v)
            if j in (9, 10, 11, 12, 13, 15, 16, 17, 19, 20, 21, 22, 23): c.number_format = '#,##0.00'
            elif j == 8: c.number_format = '0.00%'
    endr = 15 + len(otc_rows)
    for j, v in enumerate(sub('  Total All Projects'), 1):
        c = ws.cell(endr, j, v); c.font = bold
    ws.cell(endr + 1, 1, ' '); ws.cell(endr + 2, 1, ' ')
    ws.column_dimensions['B'].width = 46; ws.column_dimensions['G'].width = 30

    # ------------------------------------------------------------------ _Calc (hidden, calculated)
    ws = wb.create_sheet(CALC)
    for j, v in enumerate(['TaskProjectID', 'isFirst', 'idx', 'TaskOverrun$', 'TaskOverrunHrs'], 1):
        ws.cell(1, j, v)
    seen, idx = set(), 0
    for i, t in enumerate(task_rows):
        rr = 14 + i                    # row-aligned with the task tab
        first = 0 if t['pid'] in seen else 1
        seen.add(t['pid'])
        if first: idx += 1
        ws.cell(rr, 1, t['pid']); ws.cell(rr, 2, first)
        if first: ws.cell(rr, 3, idx)
        ws.cell(rr, 4, t['ovr']); ws.cell(rr, 5, t['ovh'])
    for rr in list(range(2, 14)) + list(range(14 + len(task_rows), t2_rows + 1)):
        ws.cell(rr, 2, 0)
    checks = [(1, T2_HDR_ROW, 'T2 header row'), (2, OTC_HDR_ROW, 'OTC header row'),
              (4, 0, 'T2 col B = Task Project ID'), (5, 0, 'T2 col G = Contract Value'),
              (6, 0, 'T2 col I = Revised Budget'), (7, 0, 'T2 col N = Total Costs'),
              (8, 0, 'T2 col O = Budgeted Hours'), (9, 0, 'T2 col P = Actual Hours'),
              (10, 0, 'OTC col A = Project ID'), (11, 0, 'OTC col C = Billing type'),
              (12, 0, 'OTC col D = Project status'), (13, 0, 'OTC col E = Location'),
              (14, 0, 'OTC col F = Project manager'), (16, 0, 'FAILED CHECKS'), (17, idx, 'project count')]
    for r_, v, lab in checks:
        ws.cell(r_, 7, v); ws.cell(r_, 8, lab)
    ws.sheet_state = 'hidden'

    # ------------------------------------------------------------------ How to Refresh
    ws = wb.create_sheet(HOWTO)
    n_ovr_lines = sum(1 for t in task_rows if t['ovr'] != 0)
    locs = sorted({(o['location_id'], o['location']) for o in W.offices})
    pms = sorted({(r['pm_id'], r['pm_name']) for r in proj_rows if r['pm_id']})
    lines = [None, 'Weekly Refresh — How It Works',
             'Project Details is driven entirely by formulas: refresh it by pasting the two new ERP exports into their tabs; nothing is typed by hand.',
             None, 'Steps each week',
             '1.  Run the "Budget to Actual Costs by Task" report in the ERP and copy the whole sheet.',
             '2.  Paste it at cell A1 of this workbook\'s tab of the same name. Always paste into the existing tab; replacing the tab renames it and breaks every reference to it.',
             '3.  Repeat with the "OTC All reports keys account in" report and its tab.',
             '4.  Go to Project Details (press F9 if the numbers have not updated) and read the status line in A3 before using anything.',
             None, 'The status line (Project Details cell A3)',
             'When the layout checks in _Calc all pass, A3 shows "Layout check OK" with the project count and the number of projects over 5%. If the ERP moves a column, A3 switches to "EXPORT LAYOUT CHANGED" with the number of failed checks: stop and fix the column mapping instead of using the figures.',
             None, 'How this file handles the current export layout',
             '•  Header and data rows sit below a report banner on both tabs, and subtotal rows ("All Tasks", "All Locations", "All Customers", "All Projects") appear. The formulas find the header row by name and ignore label and total lines, so row drift is harmless.',
             '•  A spacer column sits before the money columns on the Budget-to-Actual tab (Contract Value G, Revised Budget I, Total Costs N, Budgeted Hours O, Actual Hours P).',
             '•  The export does not supply "Task Overrun $" or "Task Overrun Hours". Both are calculated per task in the hidden _Calc tab: when a task has a revised budget, overrun = Total Cost minus Revised Budget, floored at zero; when it has no revised budget the overrun is zero, except on "OVHD-Non-Posting - Overhead Non-Posting" lines, where the full cost carries as overrun (including credits, which carry as negatives). This export has %s task lines with a non-zero overrun.' % format(n_ovr_lines, ','),
             '•  The OTC export returns Location ID and Project Manager ID instead of names. A hidden "_Lookup" tab translates them back to names (%d locations, %d managers). Any ID not in that list displays as the raw ID — add a row to _Lookup to name it.' % (len(locs), len(pms)),
             None, 'Housekeeping',
             '•  When a new export is shorter than the last one, clear the old rows left below the pasted data on both tabs.',
             '•  Project Details is all formulas — never overwrite a cell with a typed value. Filtering the sheet is fine; sorting it is not.',
             '•  The report depends on the hidden _Calc (per-task overrun) and _Lookup (ID to name) tabs — do not delete or rename them.',
             '•  Capacity: ~16,000 task rows and ~6,000 projects. Current data is %s task rows and %s projects.' % (format(len(task_rows), ','), format(len(proj_rows), ',')),
             None, 'What each computed column does',
             'Contract Value, Revised Budget, Total Cost ITD, Budgeted Hours and Actual Hours are summed across the project’s task lines. Total Task Overrun and Task Overrun Hours are summed from the per-task overrun rule above. Task Overrun % = Overrun ÷ Revised Budget. "> 5%" = 1 when that is above 5%. Cost % Complete = (Total Cost − Overrun) ÷ Revised Budget. Hours Overrun % = Overrun Hours ÷ Budgeted Hours.']
    for i, v in enumerate(lines, 1):
        if v is not None:
            c = ws.cell(i, 1, v)
            if i in (2, 5, 11, 14, 20, 26): c.font = bold
    ws.column_dimensions['A'].width = 140

    # ------------------------------------------------------------------ _Lookup (hidden)
    ws = wb.create_sheet(LOOK)
    for j, v in enumerate(['Location ID', 'Location Name', None, 'PM ID', 'Project Manager Name', None,
                           'Built from the %s export (ID -> name). Add rows here for any new location or manager ID; '
                           'unmapped IDs display as the raw ID on Project Details.' % asof.strftime('%-m.%-d.%y')], 1):
        if v is not None:
            ws.cell(1, j, v).font = bold
    for i, (lid, lname) in enumerate(locs, 2):
        ws.cell(i, 1, lid); ws.cell(i, 2, lname)
    for i, (pmid, pmn) in enumerate(pms, 2):
        ws.cell(i, 4, pmid); ws.cell(i, 5, pmn)
    ws.sheet_state = 'hidden'

    stamp = dt.datetime(asof.year, asof.month, asof.day) + dt.timedelta(days=3, hours=6, minutes=42)
    wb.properties.creator = COMPANY_SHORT + ' ERP'
    wb.properties.created = wb.properties.modified = stamp
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)

    ip = [r for r in proj_rows if r['p']['status'] == 'In Progress']
    print('wrote', out)
    print('projects:', len(proj_rows), '| task rows:', len(task_rows), '| In Progress:', len(ip),
          '| over 5% (all statuses):', n_over)
    assert abs(sum(t['N'] for t in task_rows) - sum(p['cost_to_date'] for p in W.projects)) < 1.0
    assert abs(sum(t['I'] for t in task_rows) - sum(p['revised_budget'] for p in W.projects)) < 1.0
    print('checks: task costs sum to project cost to date; task budgets sum to revised budgets -- OK')


if __name__ == '__main__':
    main()
