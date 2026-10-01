"""Location Scorecards data — writes site/drivers_data.js and the page's embedded snapshot.

The Location Scorecards page (site/location-scorecards.html) is an April-2026 "company
dashboard" snapshot per PMO plus drill-down detail. In production its driver file
(`drivers_data.js`, window.CH_DRIVERS) came from a CRM/ERP workbook with five tabs
(Bookings, T12 Win Rate, project profitability, monthly revenue by project, backlog); this
tool derives the same file from the demo world instead:

  bookings   Closed Won opportunities with a close date in the snapshot month
  winrate    Closed Won + Closed Lost opportunities closed in the trailing 12 months
  pipeline   opportunities open at month end (created <= month end < close), with the
             stage they were in at that point of their life-cycle
  backlog    active projects: contract - revenue recognised through month end (> 0 only)
  revenue    project revenue for the snapshot month and YTD, recognised from the timesheets
             (T&M: charged hours x bill rates; fixed fee: cost-to-cost % complete x contract)

It also rewrites the page's embedded snapshot block (between the DRIVERS-SNAPSHOT markers):
per-PMO KPI rows, FY budget targets and Jan..snapshot monthly revenue / GP / NI actuals,
all rolled up from the same project and timesheet figures so the drill-downs tie to the KPI
cards. Budget targets are a capacity plan (staff on the books at 1 Jan x planned utilisation
x working hours x bill rates), so actual-vs-target variances come from real utilisation,
time off and revenue recognition, not from noise.

    python3 tools/gen_location_drivers.py                     # site/drivers_data.js + page block
    python3 tools/gen_location_drivers.py --util-stub work/location/util_data.js
                                                              # + a stand-in util_data.js for testing

util_data.js itself is produced by the utilization pipeline (pipeline/sync_util_data.py);
the --util-stub writer only exists so the page can be tested before that pipeline has run.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import random
import sys
from collections import defaultdict, OrderedDict

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from demo_world import world, HOLIDAYS, NON_POSTING, season  # noqa: E402

YEAR = 2026
SNAP_MONTH = 4                                   # the page is an April snapshot (index 3, key "apr")
SNAP_END = dt.date(YEAR, SNAP_MONTH, 30)
T12_START = dt.date(YEAR - 1, SNAP_MONTH + 1, 1)  # trailing twelve months: May 2025 .. Apr 2026
MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

SITE_JS = os.path.join(ROOT, 'site', 'drivers_data.js')
SITE_HTML = os.path.join(ROOT, 'site', 'location-scorecards.html')
BEGIN_MARK = '// <<DRIVERS-SNAPSHOT:BEGIN>>'
END_MARK = '// <<DRIVERS-SNAPSHOT:END>>'

# Scorecard (CRM/ERP) names differ from the utilization site for one PMO: the International
# office is booked under its ERP location "NH Gulf (USD)" and shows as "NH Gulf" here.
SCORECARD_NAME = {'International': 'NH Gulf'}
FIELD_KINDS = ('regional', 'service', 'acquired', 'international')
# PMOs without a published FY budget: a small island office and a service line stood up
# after the budget was set; the Southwest acquisition closed after budgeting, so it has a
# placeholder (all-null) revenue row and no GP / NI rows.
NO_TARGETS = {'Island', 'Geospatial'}
NULL_REVENUE_TARGETS = {'Meridian Southwest'}
# Backlog excludes these statuses (the full production list, most of which the demo never uses).
BACKLOG_EXCLUDED = ['Legacy', 'Completed', 'On Hold', 'Cancelled', 'Work Hold', 'Regulatory Work Stop',
                    'Cost Tracking on Closed Project', 'Parent']
PLANNED_REALIZATION = 0.93        # budget assumes 93% of planned billable time converts to revenue
PLANNED_PTO = 0.05                # budgeted time off, share of full-time working hours (plus paid holidays)
ADMIN_KINDS = ('admin',)

LEAD_SOURCES = [('[Sales] Existing Customer Account', 64), ('[Sales] Employee Referral', 7),
                ('[Marketing] Direct Traffic', 6), ('[Sales] Partner', 5), ('[Sales] Account-Based Outreach', 3),
                ('[RFP] RFP/Public Notice (Direct, No Portal)', 3), ('[Sales] External Referral', 3),
                ('[RFP] Bidding System', 3), ('Unknown', 2), ('[Marketing] Advertisement', 1),
                ('[Marketing] Inbound Call', 1), ('[RFP] Bidnet', 1), ("[Marketing] Approved Vendor's List", 1)]
NEXT_STEPS = ['Follow up with client on scope and schedule', 'Submit proposal', 'Awaiting client review of proposal',
              'Confirm fee with PM and submit', 'Send revised cost estimate', 'Schedule kickoff call with client',
              'Awaiting RFP release', 'Negotiate rates and terms', 'Awaiting signed contract',
              'Check in on award decision', 'Prepare SOQ', 'Coordinate subconsultant quotes',
              'Waiting on agency permit timeline', 'Set up MSA / vendor registration']
PROJECT_TYPE = {'Class III Survey': 'Survey', 'Cultural Resources Monitoring': 'Monitoring',
                'Phase I Survey': 'Survey', 'Phase II Testing': 'Testing', 'Data Recovery': 'Data Recovery',
                'Historic Resources Evaluation': 'Consulting', 'Paleontological Monitoring': 'Monitoring',
                'Architectural History Survey': 'Survey', 'Section 106 Consultation': 'Consulting',
                'Cultural Resources Assessment': 'Cultural Consulting', 'Archaeological Testing': 'Testing',
                'Built Environment Survey': 'Survey', 'Maritime Remote Sensing Survey': 'Remote Sensing',
                'GIS Mapping': 'Reporting', 'Curation': 'Curation', 'Environmental Review Support': 'Environmental'}


def first_last(name):
    """ERP 'Last, First' -> CRM 'First Last'."""
    if ',' in name:
        last, first = [s.strip() for s in name.split(',', 1)]
        return '%s %s' % (first, last)
    return name


def workdays(y, m):
    d, n = dt.date(y, m, 1), 0
    while d.month == m:
        if d.weekday() < 5 and d not in HOLIDAYS:
            n += 1
        d += dt.timedelta(days=1)
    return n


def r2(x):
    return round(x, 2)


def build(W):
    rnd = random.Random('%s-location-drivers' % W.seed)
    offices = [o for o in W.offices if o['kind'] in FIELD_KINDS]
    pmo_of = {o['name']: SCORECARD_NAME.get(o['name'], o['name']) for o in W.offices}
    kind_of = {o['name']: o['kind'] for o in W.offices}

    # ── timesheet roll-ups by month ─────────────────────────────────────────
    labor = defaultdict(lambda: [0.0] * 12)        # project -> direct labor $ by month
    billed = defaultdict(lambda: [0.0] * 12)       # project -> hours x bill rate (excl. non-posting)
    last_month = {}                                # project -> last month with charged hours
    ind_cost = defaultdict(lambda: [0.0] * 12)     # office -> indirect (OVHD) labor $ by month
    ga_cost = [0.0] * 12                           # admin offices' non-project labor $
    util_h = defaultdict(lambda: [0.0] * 12)       # office -> utilized hrs (billable staff)
    tot_h = defaultdict(lambda: [0.0] * 12)        # office -> total hrs (billable staff)
    nonpost = defaultdict(lambda: [0.0] * 12)      # office -> Overhead Non-Posting hrs (exact task)
    bill_h = defaultdict(lambda: [0.0] * 12)       # ERP location -> utilized hrs excl. non-posting tasks
    for r in W.entries:
        # FY2026 starts with the week ending 3 Jan: its late-December days count toward January
        # in the P&L; the billable-hours series (like sync_util_data.py) keeps calendar 2026 only.
        cal = r['date'].year == YEAR
        m = r['date'].month - 1 if cal else 0
        e = W.emp[r['emp']]
        off = e['office']
        if r['project'] == 'OVHD':
            c = r['hours'] * e['cost_rate']
            if kind_of[off] in ADMIN_KINDS:
                ga_cost[m] += c
            else:
                ind_cost[off][m] += c
        else:
            mult = 1.5 if r['tt'] == 'OT' else 1.0
            labor[r['project']][m] += r['hours'] * e['cost_rate'] * mult
            if r['task'] != NON_POSTING:
                billed[r['project']][m] += r['hours'] * e['bill_rate']
            last_month[r['project']] = max(last_month.get(r['project'], 0), m)
        if e['billable'] == 'Billable':
            util_h[off][m] += r['util']
            tot_h[off][m] += r['hours']
        if r['task'] == NON_POSTING:
            nonpost[off][m] += r['hours']
        if cal and r['util'] > 0 and r['task'] not in (NON_POSTING, 'CA - Overhead Non-Posting'):
            bill_h[W.office[off]['location']][m] += r['util']

    # ── monthly revenue per project (mirrors demo_world._finalize_projects) ─
    rev_m, prior_rev, cost_m = {}, {}, {}
    for p in W.projects:
        rb, es = p['revised_budget'], p['expense_share']
        fy = labor.get(p['id'], [0.0] * 12)
        prior_labor = p['prior_share'] * rb * (1 - es) * (1 + p['burn_bias'] * 0.5)
        closed = p['status'] in ('Completed', 'Final Billing/Invoice', 'Cost Tracking on Closed Project')
        if closed and sum(fy) == 0:
            prior_labor = rb * (1 - es) * min(1.3, max(0.7, 1 + p['burn_bias']))
        gross = 1 / (1 - es)                       # labor -> labor + expenses
        cost_m[p['id']] = [x * gross for x in fy]
        cum = []
        if p['billing'] == 'Fixed fee':
            c = prior_labor * gross
            pr = p['contract'] * min(1.0, c / rb) if rb else 0.0
            for x in fy:
                c += x * gross
                cum.append(p['contract'] * min(1.0, c / rb) if rb else 0.0)
            if p['status'] in ('Completed', 'Final Billing/Invoice'):
                if sum(fy) == 0:
                    pr = p['contract']
                    cum = [pr] * 12
                else:                                  # final billing trues up to the full fee
                    lm = last_month[p['id']]
                    for i in range(lm, 12):
                        cum[i] = p['contract']
        else:
            pr = p['prior_share'] * p['contract']
            b = 0.0
            for x in billed.get(p['id'], [0.0] * 12):
                b += x
                cum.append(min(p['contract'], pr + b))
        prior_rev[p['id']] = pr
        rev_m[p['id']] = [cum[0] - pr] + [cum[i] - cum[i - 1] for i in range(1, 12)]
        # the world rounds % complete to 4 dp; book that rounding residual in the last active
        # month so the roll-up reproduces each project's revenue to the cent
        resid = p['revenue'] - (pr + sum(rev_m[p['id']]))
        if p['id'] in last_month:
            rev_m[p['id']][last_month[p['id']]] += resid
        else:
            prior_rev[p['id']] = pr = pr + resid
        assert abs(pr + sum(rev_m[p['id']]) - p['revenue']) < 1.0, (p['id'], pr + sum(rev_m[p['id']]), p['revenue'])

    # ── per-office monthly P&L actuals ──────────────────────────────────────
    rev_o = defaultdict(lambda: [0.0] * 12)
    dcost_o = defaultdict(lambda: [0.0] * 12)
    lab_o = defaultdict(lambda: [0.0] * 12)
    for p in W.projects:
        for i in range(12):
            rev_o[p['office']][i] += rev_m[p['id']][i]
            dcost_o[p['office']][i] += cost_m[p['id']][i]
            lab_o[p['office']][i] += labor.get(p['id'], [0.0] * 12)[i]
    gp_o = {o['name']: [rev_o[o['name']][i] - dcost_o[o['name']][i] for i in range(12)] for o in offices}
    # corporate G&A is allocated on each office's share of direct labor that month
    ni_o = {}
    for o in offices:
        n = o['name']
        row = []
        for i in range(12):
            tot_lab = sum(lab_o[x['name']][i] for x in offices)
            alloc = ga_cost[i] * (lab_o[n][i] / tot_lab) if tot_lab else 0.0
            row.append(gp_o[n][i] - ind_cost[n][i] - alloc)
        ni_o[n] = row

    # ── budget (capacity plan) ──────────────────────────────────────────────
    jan1 = dt.date(YEAR, 1, 1)
    plan_bias = {o['name']: rnd.gauss(0.0, 0.05) for o in offices}
    budget_ga = [0.0] * 12
    for e in W.employees:
        if W.office[e['office']]['kind'] in ADMIN_KINDS and e['status'] == 'Active' and e['hire'] <= jan1:
            for i in range(12):
                budget_ga[i] += 8 * workdays(YEAR, i + 1) * e['cost_rate']
    plan = {}
    for o in offices:
        n = o['name']
        staff = [e for e in W.employees if e['office'] == n and e['hire'] <= jan1
                 and (e['term'] is None or e['term'] >= jan1)]
        # expense gross-up at plan: the office's average project expense share (live jobs)
        live = [p for p in W.projects if p['office'] == n and p['end'] >= jan1] or \
            [p for p in W.projects if p['office'] == n]
        es = sum(p['expense_share'] for p in live) / len(live)
        rev, dcost, ind, dlab, ut = [], [], [], [], []
        for i in range(12):
            sea = season(o['kind'], dt.date(YEAR, i + 1, 15))
            hrs_fte = 8 * workdays(YEAR, i + 1)
            hol = 8 * sum(1 for d in HOLIDAYS if d.year == YEAR and d.month == i + 1)
            r_ = d_ = x_ = 0.0
            for e in staff:
                if e['ft'] == 'Full Time':
                    h, paid_off = hrs_fte * (1 - PLANNED_PTO), hrs_fte * PLANNED_PTO + hol
                else:
                    h, paid_off = hrs_fte * 0.55 * (1 + 3 * sea), 0.0
                u = max(0.0, min(0.97, e['target'] + sea)) if e['target'] > 0 else 0.0
                r_ += h * u * e['bill_rate'] * PLANNED_REALIZATION
                d_ += h * u * e['cost_rate']
                x_ += (h * (1 - u) + paid_off) * e['cost_rate']
            rev.append(r_ * (1 + plan_bias[n])); dlab.append(d_); dcost.append(d_ / (1 - es)); ind.append(x_)
            ut.append(round(o['target'] + round(sea / 2, 2), 2) if o['target'] else None)
        plan[n] = dict(rev=rev, gp=[rev[i] - dcost[i] for i in range(12)], ind=ind, dlab=dlab, util=ut)
    # budgeted G&A is allocated on planned direct labor, as the actuals are
    for i in range(12):
        tot = sum(plan[n]['dlab'][i] for n in plan)
        for n in plan:
            plan[n].setdefault('ni', [None] * 12)[i] = \
                plan[n]['gp'][i] - plan[n]['ind'][i] - budget_ga[i] * plan[n]['dlab'][i] / tot

    # ── CRM: bookings / win rate / pipeline / proposals ─────────────────────
    lead = {}
    for op in W.opportunities:
        orng = random.Random('%s-location-drivers-%s' % (W.seed, op['id']))
        lead[op['id']] = orng.choices([s for s, _ in LEAD_SOURCES], weights=[w for _, w in LEAD_SOURCES])[0]
    opps = sorted(W.opportunities, key=lambda x: (x['close'], x['id']))
    bookings = [dict(pmo=pmo_of[op['office']], account=op['client'], opportunity=op['name'], stage='Closed Won',
                     amount=float(op['amount']), close_date=op['close'].isoformat(), owner=first_last(op['owner']),
                     lead_source=lead[op['id']])
                for op in opps if op['stage'] == 'Closed Won' and op['close'].year == YEAR and op['close'].month == SNAP_MONTH]
    winrate = [dict(pmo=pmo_of[op['office']], account=op['client'], opportunity=op['name'], stage=op['stage'],
                    amount=float(op['amount']), close_date=op['close'].isoformat(),
                    fiscal_period='Q%d-%d' % ((op['close'].month - 1) // 3 + 1, op['close'].year))
               for op in opps if op['stage'] in ('Closed Won', 'Closed Lost') and T12_START <= op['close'] <= SNAP_END]
    prob = {'Prospecting': 0.10, 'Qualification': 0.20, 'Proposal Development': 0.40, 'Proposal Submitted': 0.50,
            'Negotiation': 0.75, 'Follow-Up': 0.50}
    pipeline = []
    for op in sorted(W.opportunities, key=lambda x: (pmo_of[x['office']], x['created'], x['id'])):
        if not (op['created'] <= SNAP_END < op['close']):
            continue
        if op['stage'] in prob:
            stage = op['stage']
        else:                                      # closed later: stage it was in at month end
            f = (SNAP_END - op['created']).days / max(1, (op['close'] - op['created']).days)
            stage = ('Prospecting' if f < 0.15 else 'Qualification' if f < 0.3 else 'Proposal Development' if f < 0.5
                     else 'Proposal Submitted' if f < 0.75 else 'Follow-Up' if f < 0.9 else 'Negotiation')
        orng = random.Random('%s-location-drivers-ns-%s' % (W.seed, op['id']))
        pipeline.append(dict(pmo=pmo_of[op['office']], account=op['client'], opportunity=op['name'], stage=stage,
                             amount=float(op['amount']), weighted=r2(op['amount'] * prob[stage]),
                             probability=prob[stage], close_date=op['close'].isoformat(),
                             owner=first_last(op['owner']),
                             next_step='' if orng.random() < 0.14 else orng.choice(NEXT_STEPS),
                             lead_source=lead[op['id']]))

    # ── backlog at month end ────────────────────────────────────────────────
    backlog = []
    snap_i = SNAP_MONTH - 1
    for p in sorted(W.projects, key=lambda x: (x['id'])):
        if p['start'] > SNAP_END:
            continue
        status = p['status']
        if status in ('Completed', 'Final Billing/Invoice', 'Cost Tracking on Closed Project', 'Transferred Entity') \
                and last_month.get(p['id'], -1) > snap_i:
            status = 'In Progress'                 # still being worked after the snapshot
        if status in BACKLOG_EXCLUDED:
            continue
        rec = prior_rev[p['id']] + sum(rev_m[p['id']][:snap_i + 1])
        bl = p['contract'] - rec
        if bl <= 0.5:
            continue
        backlog.append(dict(pmo=pmo_of[p['office']], project_id=p['id'], project_name=p['name'],
                            billing_type=p['billing'], status=status, contract_value=float(p['contract']),
                            revenue_recognized=r2(rec), backlog=r2(bl), pm=first_last(p['pm'])))
    backlog.sort(key=lambda r: (r['project_id']))
    backlog_totals = OrderedDict()
    for r in backlog:
        backlog_totals[r['pmo']] = r2(backlog_totals.get(r['pmo'], 0.0) + r['backlog'])

    # ── revenue by project (snapshot month + YTD) ───────────────────────────
    revenue = []
    for p in sorted(W.projects, key=lambda x: (pmo_of[x['office']], x['id'])):
        apr = rev_m[p['id']][snap_i]
        ytd = sum(rev_m[p['id']][:snap_i + 1])
        if abs(ytd) < 0.005 and abs(apr) < 0.005:
            continue
        prng = random.Random('%s-location-drivers-pt-%s' % (W.seed, p['id']))
        ptype = PROJECT_TYPE.get(p['svc'], 'Consulting')
        if ptype == 'Survey' and prng.random() < 0.08:
            ptype = 'survey'                       # free-text field: entry quirks survive in the ERP
        revenue.append(dict(pmo=pmo_of[p['office']], project=p['name'], dimension_id=p['id'], customer=p['client'],
                            project_type=ptype, sector='' if prng.random() < 0.16 else p['sector'],
                            apr=r2(apr), ytd=r2(ytd)))

    # ── per-PMO snapshot rows ───────────────────────────────────────────────
    def total(rows, pmo, key):
        return sum(r[key] for r in rows if r['pmo'] == pmo)

    snap_rows = []
    util_act = {}
    for o in sorted(offices, key=lambda x: pmo_of[x['name']].lower()):
        n, loc = o['name'], pmo_of[o['name']]
        won = total([r for r in winrate if r['stage'] == 'Closed Won'], loc, 'amount')
        lost = total([r for r in winrate if r['stage'] == 'Closed Lost'], loc, 'amount')
        bk = [r for r in bookings if r['pmo'] == loc]
        props = sum(op['amount'] for op in W.opportunities
                    if op['office'] == n and op['created'].year == YEAR and op['created'].month == SNAP_MONTH)
        ut = [round(util_h[n][i] / tot_h[n][i], 4) if tot_h[n][i] else None for i in range(snap_i + 1)]
        util_act[loc] = ut
        snap_rows.append(OrderedDict(
            loc=loc, proposals=r2(props), pipeline=r2(total(pipeline, loc, 'amount')),
            wr=round(won / (won + lost), 4) if won + lost else None,
            bookings=r2(sum(r['amount'] for r in bk)) if bk else None,
            backlog=round(backlog_totals.get(loc, 0.0)),
            bhrs=r2(bill_h[o['location']][snap_i]),
            rev=r2(rev_o[n][snap_i]), gp=r2(gp_o[n][snap_i]), ni=r2(ni_o[n][snap_i]),
            costs=round(lab_o[n][snap_i]), util=ut[snap_i], ovhd=r2(nonpost[n][snap_i])))

    drivers = OrderedDict(bookings=bookings, winrate=winrate, backlog=backlog, revenue=revenue,
                          backlog_totals=backlog_totals, pipeline=pipeline)
    targets = OrderedDict()
    for sec in ('revenue', 'gp', 'ni', 'util'):
        t = OrderedDict()
        for o in sorted(offices, key=lambda x: pmo_of[x['name']].lower()):
            n, loc = o['name'], pmo_of[o['name']]
            if n in NO_TARGETS:
                continue
            if n in NULL_REVENUE_TARGETS:
                if sec == 'revenue':
                    t[loc] = [None] * 12
                elif sec == 'util':
                    t[loc] = plan[n]['util']
                continue
            vals = plan[n]['rev' if sec == 'revenue' else sec]
            t[loc] = [round(v) for v in vals] if sec != 'util' else vals
        targets[sec] = t
    acts = {k: OrderedDict((pmo_of[o['name']], [round(v) for v in src[o['name']][:snap_i + 1]])
                           for o in sorted(offices, key=lambda x: pmo_of[x['name']].lower()))
            for k, src in (('rev', rev_o), ('gp', gp_o), ('ni', ni_o))}
    extra = dict(as_of=W.as_of, util_h=util_h, tot_h=tot_h, nonpost=nonpost, bill_h=bill_h, plan=plan, offices=W.offices)
    return drivers, snap_rows, targets, util_act, acts, extra


# ── writers ──────────────────────────────────────────────────────────────────
def js_num(v):
    if v is None:
        return 'null'
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return repr(v)


def write_drivers_js(drivers, path):
    js = ('// Auto-generated by tools/gen_location_drivers.py. Sources: demo CRM/ERP (Bookings, T12 Win Rate, '
          'Project Profitability, Monthly Revenue by Project, Backlog).\n'
          '// Backlog excludes statuses: %s. Negative backlog excluded.\n'
          '// PMO names as the scorecard shows them (International -> NH Gulf alias applied).\n'
          'window.CH_DRIVERS = %s;\n') % (', '.join(BACKLOG_EXCLUDED),
                                         json.dumps(drivers, ensure_ascii=False, separators=(',', ':')))
    with open(path, 'w', encoding='utf-8') as f:
        f.write(js)


def arr(vals):
    return '[' + ','.join(js_num(v) for v in vals) + ']'


def snapshot_block(rows, targets, util_act, acts):
    w = max(len(r['loc']) for r in rows) + 3
    out = [BEGIN_MARK,
           '// Generated by tools/gen_location_drivers.py from the demo world; do not edit by hand.',
           '// ── April Dashboard Snapshot ────────────────────────────────────────────────',
           '// One row per PMO (field offices, service lines, acquisitions and the International',
           '// entity, which the CRM/ERP books as "NH Gulf").',
           'const APRIL = [']
    for r in rows:
        parts = ['%s:%s' % (k, js_num(r[k])) for k in
                 ('proposals', 'pipeline', 'wr', 'bookings', 'backlog', 'bhrs', 'rev', 'gp', 'ni', 'costs', 'util', 'ovhd')]
        out.append('  {loc:%s' % ("'%s'," % r['loc']).ljust(w + 1) + ', '.join(parts) + '},')
    out += ['];', '',
            '// ── Annual Targets (Jan–Dec 2026) ────────────────────────────────────────────',
            '// FY budget by PMO. PMOs without a published budget have no entry (the scorecard',
            '// shows actuals only); the mid-year acquisition carries a placeholder revenue row.',
            'let TARGETS = {']
    for sec, t in targets.items():
        out.append('  %s:{' % sec)
        for loc, vals in t.items():
            out.append("    %s%s," % (("'%s':" % loc).ljust(w + 1), arr(vals)))
        out.append('  },')
    out += ['};', '',
            '// ── Billable hours by month (overridden by util_data.js when present) ────────',
            'let BILLABLE_HRS = {};   // {office: [Jan…Dec]}  — populated from util_data.js', '',
            '// ── Monthly Actuals Jan–Apr 2026 (overridden by util_data.js when present) ───',
            'let UTIL_ACT = {']
    for loc, vals in util_act.items():
        out.append("  %s%s," % (("'%s':" % loc).ljust(w + 1), arr(vals)))
    out += ['};', '',
            '// ── Monthly Revenue / GP / NI Actuals (Jan–Apr 2026) ────────────────────────',
            '// Rolled up from project revenue and timesheet costs per PMO (same source as the drill-downs).']
    for name, key in (('REV_ACT', 'rev'), ('GP_ACT', 'gp'), ('NI_ACT', 'ni')):
        out.append('const %s = {' % name)
        for loc, vals in acts[key].items():
            out.append("  %s%s," % (("'%s':" % loc).ljust(w + 1), arr(vals)))
        out += ['};', '']
    out.append(END_MARK)
    return '\n'.join(out)


def patch_page(path, block):
    with open(path, encoding='utf-8') as f:
        html = f.read()
    i, j = html.find(BEGIN_MARK), html.find(END_MARK)
    if i < 0 or j < 0:
        raise SystemExit('snapshot markers not found in %s' % path)
    html = html[:i] + block + html[j + len(END_MARK):]
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html)


def write_util_stub(extra, path):
    """Stand-in for pipeline/sync_util_data.py output (window.CH_UTIL_DATA), same schema:
    generated, source_file, util_targets / util_actuals {office: 12 x fraction|None} (utilization-site
    office names, the hidden-on-monthly office omitted), nonpost_hrs {office: [Jan..May, YTD total]},
    billable_hrs {ERP location: 12 x hours} (utilized hours excluding Overhead Non-Posting tasks)."""
    from demo_world import HIDDEN_ON_MONTHLY
    last_m = extra['as_of'].month - 2          # last complete month before the as-of date (Aug)
    ut, ua, npost, bh = OrderedDict(), OrderedDict(), OrderedDict(), OrderedDict()
    for o in sorted(extra['offices'], key=lambda x: x['name']):
        n = o['name']
        if o['kind'] not in FIELD_KINDS or n.lower() in HIDDEN_ON_MONTHLY:
            continue
        ut[n] = extra['plan'][n]['util']
        ua[n] = [round(extra['util_h'][n][i] / extra['tot_h'][n][i], 2) if i <= last_m and extra['tot_h'][n][i] else None
                 for i in range(12)]
        np_ = [int(round(extra['nonpost'][n][i])) for i in range(5)]
        npost[n] = np_ + [int(round(sum(extra['nonpost'][n][:last_m + 1])))]
    for loc, vals in sorted(extra['bill_h'].items()):
        bh[loc.replace(' (USD)', '')] = [round(v) for v in vals]
    data = OrderedDict(generated=dt.datetime.now(dt.timezone.utc).isoformat(), source_file='index.html',
                       util_targets=ut, util_actuals=ua, nonpost_hrs=npost, billable_hrs=bh)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write('// Auto-generated by sync_util_data.py — do not edit manually.\n'
                '// Re-run sync_util_data.py to refresh from index.html\n'
                'window.CH_UTIL_DATA = ' + json.dumps(data, indent=2, ensure_ascii=False) + ';\n')


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('--out', default=SITE_JS)
    ap.add_argument('--page', default=SITE_HTML)
    ap.add_argument('--no-page', action='store_true', help='do not rewrite the page snapshot block')
    ap.add_argument('--util-stub', help='also write a stand-in util_data.js here (testing only)')
    a = ap.parse_args()
    W = world()
    drivers, rows, targets, util_act, acts, extra = build(W)
    write_drivers_js(drivers, a.out)
    if not a.no_page:
        patch_page(a.page, snapshot_block(rows, targets, util_act, acts))
    if a.util_stub:
        write_util_stub(extra, a.util_stub)
    # summary
    tot = lambda k: sum((r[k] or 0) for r in rows)
    print('drivers_data.js: bookings %d, winrate %d, backlog %d, revenue %d, pipeline %d rows, %d PMOs'
          % (len(drivers['bookings']), len(drivers['winrate']), len(drivers['backlog']), len(drivers['revenue']),
             len(drivers['pipeline']), len(rows)))
    print('April: revenue %.0f  GP %.0f  NI %.0f  bookings %.0f  backlog %.0f  pipeline %.0f'
          % (tot('rev'), tot('gp'), tot('ni'), tot('bookings'), tot('backlog'), tot('pipeline')))
    # ties: KPI rows == drill-down sums
    for r in rows:
        rv = sum(x['apr'] for x in drivers['revenue'] if x['pmo'] == r['loc'])
        assert abs(rv - r['rev']) < 1, (r['loc'], rv, r['rev'])
        assert abs(drivers['backlog_totals'].get(r['loc'], 0) - r['backlog']) < 1
    print('checks passed: revenue and backlog drill-downs tie to the KPI rows')


if __name__ == '__main__':
    main()
