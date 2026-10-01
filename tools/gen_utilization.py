#!/usr/bin/env python3
"""gen_utilization.py — synthetic source for the main Utilization Site (index.html hub).

Writes, from the shared demo world (tools/demo_world.py):

  demo_data/Weekly and Monthly Updates/9.29.2026/Utilization/Utilization Report.xlsx
      The weekly utilization workbook in the export's layout (sheet names, header rows,
      column positions, labels and the export's quirks), with cached VALUES only:
        Executive Summary, PMO (Monthly), PMO (Weekly), Utilization (Employee),
        Specialized Reporting ->, Summary - OT, Summary - Non-Posting, Summary - PTO,
        Pivot - Fleet, Non-Targeted, PT Hours, OTC-Project Profitability CM to, Raw Data,
        Employee Dashboard, Backend ->, Proj. Sector, Data, Current Employee Roster
      (no Pivot - Proposal: the pipeline rebuilds it from Raw Data.)

  demo_data/.../Utilization/Projects.csv
      The ERP project master (Project ID, name, type, customer, category, status, sector).

  demo_data/.../Utilization/Retired Tabs - Utilization Report 8.15.2026.xlsx
      The last export that still carried the two retired tabs, 'Summary - OVHD (FT Employee)'
      and 'Recovery Planning' (as of 2026-08-15). The site carries those two tabs forward; the
      pipeline renders them from this file once (see the refresh notes), after which merge.py
      keeps them unchanged, exactly as it does for the live site.

How the numbers are derived (nothing is drawn independently):
  * Raw Data is one row per world timesheet line (W.entries) with the export's 30 columns;
    the helper columns (Proposal, Vacation, Required Training, Exclusions, Short Term PTO,
    Billing Type*, Month, Employee Target …) follow the workbook's own column rules.
  * Data is the workbook's per-employee engine: for every week / month / quarter / YTD block,
    Utld | NonUtld | Total | Unadj. % | Exclu. | Adj. %, computed exactly as its SUMIFS
    formulas do from Raw Data (entry date in (previous block date, block date]; weekly and
    stub Totals carry the 40-hour salaried cap; Exclu. = the Vacation helper column), with
    the eight company summary rows on top.
  * Every other sheet is an aggregation of Raw Data + Data + the roster, mirroring the
    formulas of the real workbook: PMO (Weekly) = SUMIFS over Data blocks by department
    for Billable employees (Company / Full Time / Part Time), PMO (Monthly) the same over
    the month blocks, Utilization (Employee) the per-employee weekly Unadj. % (N/A in weeks
    the person was not on the active roster), Summary - OT / Non-Posting by project
    location and month, Employee Dashboard the per-week hour sums by Week End, etc. So every
    office figure is the sum of its employees and every company figure the sum of offices.
  * Monthly office targets are the plan: each office's hours-weighted average of its billable
    employees' targets (future months weighted by the current active roster).
  * The OTC-Project Profitability tab reuses gen_revenue_forecast.otc_rows(W), so the backlog
    here equals the Revenue Forecast dashboard's.

Usage:  python3 tools/gen_utilization.py
Deterministic: the only randomness is random.Random(f"{W.seed}-utilization").
"""
from __future__ import annotations

import calendar
import csv
import datetime as dt
import hashlib
import os
import random
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from demo_world import (world, COMPANY, NON_POSTING, FIRST_WEEK_END, LOCATION_OF,  # noqa: E402
                        INTL_ENTITY_LABEL, TASK_TRAINING)
from gen_revenue_forecast import otc_rows  # noqa: E402

import openpyxl  # noqa: E402
from openpyxl.cell import WriteOnlyCell  # noqa: E402

REPO = os.path.dirname(HERE)
OUT_DIR = os.path.join(REPO, 'demo_data', 'Weekly and Monthly Updates', '9.29.2026', 'Utilization')
XLSX = os.path.join(OUT_DIR, 'Utilization Report.xlsx')
RETIRED_XLSX = os.path.join(OUT_DIR, 'Retired Tabs - Utilization Report 8.15.2026.xlsx')
PROJECTS_CSV = os.path.join(OUT_DIR, 'Projects.csv')

FY = 2026
WEEKS = []                                      # 52 week-ending Saturdays of FY2026
_d = FIRST_WEEK_END
while _d.year == FY or _d == FIRST_WEEK_END:
    WEEKS.append(_d); _d += dt.timedelta(days=7)
WEEKS = WEEKS[:52]
WEEK_STUB = FIRST_WEEK_END - dt.timedelta(days=7)          # 2025-12-27 (Data!M9)
PERIOD_STUB = dt.date(FY - 1, 12, 28)                       # month/quarter stub (=M9+1)
MONTH_ENDS = [dt.date(FY, m, calendar.monthrange(FY, m)[1]) for m in range(1, 13)]
QUARTER_ENDS = [MONTH_ENDS[2], MONTH_ENDS[5], MONTH_ENDS[8], MONTH_ENDS[11]]
MONTH_NAMES = [calendar.month_name[m] for m in range(1, 13)]
MON3 = [calendar.month_abbr[m] for m in range(1, 13)]

PTO_TT = {'HOL', 'PTO', 'Vac', 'Sick', 'MAT'}          # Raw Data 'Vacation' helper
SHORT_PTO_TT = {'HOL', 'PTO', 'Vac', 'Sick'}           # 'Short Term PTO (Value)'
EXTENDED_TT = {'STD', 'MAT'}                            # Summary - PTO 'Extended PTO'
PROPOSAL_TASKS = {'OVHD - Labor - Proposal Writing', 'OVHD - Production - Proposal Editing'}
FLEET_TASKS = {'OVHD - Labor - Local Fleet Management', 'OVHD - Labor - National Fleet Management'}
CONTRACTOR = 'Contractor'                   # Data!D for contractor-platform staff (IDs S…/E…)
ADMIN_DEPTS = {'Corporate', 'Sales & Marketing', 'Archive Services'}

RAW_HEADER = ['Legacy ID', 'Contact name', 'Location name', 'Entry date', 'Project ID', 'Utilized hours',
              'Non-utilized hours', 'Hours', 'Utilization Rate', 'Time type', 'Task name', 'Notes',
              'Billing type', 'Billing Type*', 'Proposal', 'Vacation', 'Required Training', 'Exclusions',
              'Week End', 'PT / FT', 'Employee Home Location', 'Location Conversion', 'Employee Target',
              'Document #', 'Hourly / Salary', 'Month', 'Supervisor', 'Short Term PTO (Value)', 'Role',
              'Status']

NOTES = {
    'Fieldwork': ['Field day', 'field work', 'Field survey', 'fieldwork'],
    'Survey': ['Survey', 'pedestrian survey', 'survey coverage'],
    'Monitoring': ['Monitoring', 'construction monitoring', 'monitoring - grading'],
    'Report': ['Reporting', 'report edits', 'draft report', 'report figures'],
    'Research': ['background research', 'records search', 'research'],
    'Desktop': ['desktop review', 'records review'],
    'GIS': ['GIS', 'mapping', 'figures'],
    'Lab': ['lab', 'artifact analysis', 'catalog'],
    'Curation': ['curation packaging', 'curation'],
    'Project Management': ['PM', 'project management', 'billing review', 'client call'],
    'Consultation': ['agency call', 'consultation letters'],
    'Prefield': ['pre-field planning', 'field prep', 'safety plan'],
    'Recordation': ['site forms', 'site recordation'],
    'Testing': ['test units', 'shovel testing'],
    'Excavation': ['excavation', 'unit excavation'],
    'Non-Posting': ['budget review', 'invoice prep', 'client follow up', 'project setup'],
    'Vacation': ['VAC', 'Vacation'],
    'PTO': ['PTO', 'VAC'],
    'Holiday': ['Holiday', 'VAC'],
    'Sick': ['Sick'],
    'Parental': ['Leave'],
    'Training': ['required training', 'safety training', 'training module'],
    'Meetings': ['staff meeting', 'weekly meeting', '1:1 with supervisor', 'team meeting'],
    'Proposal': ['proposal', 'proposal writing', 'cost estimate', 'scope & budget'],
    'Business Development': ['client outreach', 'BD call', 'conference follow up'],
    'Marketing': ['marketing', 'qualifications package'],
    'Fleet': ['vehicle maintenance', 'fleet scheduling', 'mileage logs'],
    'Administrative': ['admin', 'Admin', 'email', 'timesheets', 'Admin tasks'],
}


def note_for(task, rnd):
    for key, opts in NOTES.items():
        if key.lower() in task.lower():
            return rnd.choice(opts)
    return rnd.choice(['.', 'admin', 'misc', 'coordination'])


def eomonth(d):
    return dt.date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])


def eid_value(e):
    """Legacy ID as the export stores it: numeric IDs as numbers, platform IDs as text."""
    return e['id'] if isinstance(e['id'], int) else str(e['id'])


def is_platform(e):
    s = str(e['id'])
    return s[:1] in ('S', 'E')


def num(x):
    if isinstance(x, float) and x.is_integer():
        return int(x)
    return x


def spaced(s):
    return ' '.join(s)


def excel_style_package(path):
    """Re-save the package the way Excel does, as in the real export: text in a shared-strings
    table (openpyxl's write-only mode writes inline strings), workbook relationships with Id
    first and targets relative to xl/. Stream readers keyed on Excel's layout
    (regenerate_pmo_reduced.Workbook) depend on both."""
    import re
    import shutil
    import tempfile
    import zipfile
    sst, sst_idx = [], {}
    cell_re = re.compile(r'<c r="([A-Z]+[0-9]+)"([^>]*?) t="inlineStr"><is><t(?: xml:space="preserve")?>(.*?)</t></is></c>', re.S)
    empty_re = re.compile(r'<c r="([A-Z]+[0-9]+)"([^>]*?) t="inlineStr"><is><t(?: xml:space="preserve")?/></is></c>')

    def to_shared(m):
        txt = m.group(3)
        i = sst_idx.get(txt)
        if i is None:
            i = sst_idx[txt] = len(sst)
            sst.append(txt)
        return '<c r="%s"%s t="s"><v>%d</v></c>' % (m.group(1), m.group(2), i)
    tmp = tempfile.mktemp(suffix='.xlsx', dir=os.path.dirname(path))
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
                t = data.decode('utf-8')
                t = empty_re.sub(lambda m: '<c r="%s"%s/>' % (m.group(1), m.group(2)), t)
                t = cell_re.sub(to_shared, t)
                # Excel records the used range; readers (openpyxl read-only) pad rows to it
                if '<dimension' not in t:
                    cols = set(re.findall(r'<c r="([A-Z]+)[0-9]+"', t))
                    rows_ = re.findall(r'<row r="([0-9]+)"', t)
                    if cols and rows_:
                        def cn(x):
                            n = 0
                            for ch in x:
                                n = n * 26 + ord(ch) - 64
                            return n
                        last_col = max(cols, key=cn)
                        t = t.replace('<sheetViews>', '<dimension ref="A1:%s%s"/><sheetViews>' % (last_col, rows_[-1]), 1)
                data = t.encode('utf-8')
            elif item.filename == 'xl/_rels/workbook.xml.rels':
                t = data.decode('utf-8')

                def fix(m):
                    attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
                    tgt = attrs['Target']
                    if tgt.startswith('/xl/'):
                        tgt = tgt[len('/xl/'):]
                    return '<Relationship Id="%s" Type="%s" Target="%s"/>' % (attrs['Id'], attrs['Type'], tgt)
                t = re.sub(r'<Relationship ([^>]*?)/>', fix, t)
                if 'sharedStrings' not in t:
                    n = len(re.findall(r'<Relationship ', t)) + 1
                    t = t.replace('</Relationships>', '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
                                  'officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
                                  '</Relationships>' % n)
                data = t.encode('utf-8')
            elif item.filename == '[Content_Types].xml':
                t = data.decode('utf-8')
                if 'sharedStrings' not in t:
                    t = t.replace('</Types>', '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
                                  'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>')
                data = t.encode('utf-8')
            zout.writestr(item, data)
        if 'xl/sharedStrings.xml' not in names:
            body = ''.join('<si><t xml:space="preserve">%s</t></si>' % x if (x[:1].isspace() or x[-1:].isspace() or '\n' in x)
                           else '<si><t>%s</t></si>' % x for x in sst)
            zout.writestr('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                          '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">%s</sst>'
                          % (len(sst), len(sst), body))
    shutil.move(tmp, path)


# ─────────────────────────────────────────────────────────────────────────────────────────────
class Model:
    """Everything derived from the world that more than one sheet needs."""

    def __init__(self, W):
        self.W = W
        self.as_of = W.as_of
        self.rnd = random.Random('%s-utilization' % W.seed)
        self.emps = sorted(W.employees, key=lambda e: (isinstance(e['id'], str), str(e['id']) if isinstance(e['id'], str) else e['id']))
        self.emp = {str(e['id']): e for e in W.employees}
        self.office = W.office
        # per-employee day series: date -> [utld, nonutld, vacation(P), exclusions(R), ot]
        self.day0 = WEEK_STUB - dt.timedelta(days=7)
        self.ndays = (MONTH_ENDS[-1] - self.day0).days + 1
        self.cum = {}
        self.raw = []                      # Raw Data rows (lists, 30 cols)
        self._build_raw()
        self.last_week_idx = max(i for i, w in enumerate(WEEKS) if w <= self.as_of)
        self.past_weeks = WEEKS[:self.last_week_idx + 1]
        self._build_data_blocks()

    # ── Raw Data ─────────────────────────────────────────────────────────────────────────────
    def location_of_entry(self, r, e):
        if r['project'] != 'OVHD':
            return self.W.project[r['project']]['location']
        return LOCATION_OF[e['office']]

    def _build_raw(self):
        W = self.W
        rnd = self.rnd
        daily = defaultdict(lambda: [0.0] * 5)
        order = sorted(W.entries, key=lambda r: (isinstance(self.emp[r['emp']]['id'], str),
                                                 str(self.emp[r['emp']]['id']) if isinstance(self.emp[r['emp']]['id'], str) else self.emp[r['emp']]['id'],
                                                 r['date'], r['doc']))
        self.entries = order
        for r in order:
            e = self.emp[r['emp']]
            h, u = r['hours'], r['util']
            n = round(h - u, 4)
            task = r['task']
            tt = r['tt']
            proj = r['project']
            billing = '' if proj == 'OVHD' else W.project[proj]['billing']
            prop = h if task in PROPOSAL_TASKS else 0.0
            vac = h if tt in PTO_TT else 0.0
            train = h if task == TASK_TRAINING[0] else 0.0
            excl = prop + vac + train
            stp = h if tt in SHORT_PTO_TT else 0.0
            loc = self.location_of_entry(r, e)
            rate = 100.0 if (u != 0 and 'non-posting' not in task.lower()) else 0.0
            self.raw.append([
                eid_value(e), e['name'], loc, r['date'], proj, u, n, h, rate, tt, task,
                note_for(task, rnd), billing, billing or 'Overhead', prop, vac, train, excl,
                r['week'], e['ft'], e['office'], loc, e['target'], r['doc'], 0.0, eomonth(r['date']),
                e['supervisor'] or '', stp, e['title'], e['status']])
            d = daily[(r['emp'], r['date'])]
            d[0] += u; d[1] += n; d[2] += vac; d[3] += excl
            if tt == 'OT':
                d[4] += h
        # cumulative arrays per employee (index 0 = day0)
        per = defaultdict(dict)
        for (k, day), v in daily.items():
            per[k][(day - self.day0).days] = v
        for e in self.emps:
            k = str(e['id'])
            arr = [[0.0] * 5 for _ in range(self.ndays + 1)]
            acc = [0.0] * 5
            days = per.get(k, {})
            for i in range(self.ndays):
                v = days.get(i)
                if v:
                    acc = [acc[j] + v[j] for j in range(5)]
                arr[i + 1] = acc
            self.cum[k] = arr

    def sums(self, k, lo, hi):
        """[utld, nonutld, vac, excl, ot] for entry dates in (lo, hi]."""
        arr = self.cum[k]
        a = min(max((lo - self.day0).days + 1, 0), self.ndays)
        b = min(max((hi - self.day0).days + 1, 0), self.ndays)
        return [arr[b][j] - arr[a][j] for j in range(5)]

    def ytd(self, k):
        arr = self.cum[k]
        return arr[self.ndays]

    # ── Data blocks ──────────────────────────────────────────────────────────────────────────
    @staticmethod
    def block_values(e, s, cap):
        """Utld, NonUtld, Total, Unadj %, Exclu, Adj % exactly as the Data formulas compute them."""
        u, n = s[0], s[1]
        tot_raw = u + n
        if cap:
            total = 40.0 if (tot_raw != 0 and e['flsa'] == 'Salary Exempt') else tot_raw
        else:
            total = tot_raw
        exclu = s[2]
        if is_platform(e):
            unadj = (u / total) if total else 'N/A'
        elif e['flsa'] == 'Hourly Non Exempt' and tot_raw == 0:
            unadj = 'N/A'
        elif tot_raw == 0:
            unadj = 'N/T'
        else:
            unadj = (u / total) if total else 0
        if total == exclu and exclu > 0:
            adj = 'N/A'
        else:
            adj = (u / (total - exclu)) if (total - exclu) else 0
        return [u, n, total, unadj, exclu, adj]

    def _build_data_blocks(self):
        self.blocks = {}   # key -> list of (label, date, cap, lo, hi)
        wk = []
        prev = WEEK_STUB
        for w in WEEKS:
            wk.append((w, prev, w)); prev = w
        self.week_windows = wk
        mo = []
        prev = PERIOD_STUB
        for m in MONTH_ENDS:
            mo.append((m, prev, m)); prev = m
        self.month_windows = mo
        qu = []
        prev = PERIOD_STUB
        for q in QUARTER_ENDS:
            qu.append((q, prev, q)); prev = q
        self.quarter_windows = qu
        # per-employee values
        self.dv = {}
        for e in self.emps:
            k = str(e['id'])
            v = {'W': [self.block_values(e, self.sums(k, lo, hi), True) for _, lo, hi in wk],
                 'M': [self.block_values(e, self.sums(k, lo, hi), False) for _, lo, hi in mo],
                 'Q': [self.block_values(e, self.sums(k, lo, hi), False) for _, lo, hi in qu],
                 'MS': self.block_values(e, self.sums(k, PERIOD_STUB - dt.timedelta(days=31), PERIOD_STUB), True),
                 'QS': self.block_values(e, self.sums(k, PERIOD_STUB - dt.timedelta(days=92), PERIOD_STUB), True),
                 'Y': self.block_values(e, self.ytd(k), False)}
            self.dv[k] = v

    # ── roster helpers ───────────────────────────────────────────────────────────────────────
    def active_in_week(self, e, w):
        start = w - dt.timedelta(days=6)
        if e['hire'] > w:
            return False
        if e['term'] and e['term'] < start:
            return False
        return True

    def discipline(self, e):
        if e['target'] <= 0:
            return '-'
        d = e['office']
        if 'Meridian' in d:
            return 'Meridian'
        if 'Geospatial' in d:
            return 'Geospatial'
        if 'International' in d:
            return 'International'
        return 'Domestic'

    def data_status(self, e):
        return e['status']

    def position(self, e):
        return CONTRACTOR if is_platform(e) else e['title']


# ─────────────────────────────────────────────────────────────────────────────────────────────
def pct_or(v):
    return v


def ratio(a, b, na='N/A'):
    return (a / b) if b else na


def main():
    W = world()
    M = Model(W)
    os.makedirs(OUT_DIR, exist_ok=True)
    rnd = random.Random('%s-utilization-sheets' % W.seed)
    as_of = W.as_of
    LW = M.last_week_idx

    offices_all = [o['name'] for o in W.offices]
    field_offices = sorted(o['name'] for o in W.offices if o['kind'] != 'admin')
    company_offices = sorted(o for o in offices_all if o != 'Archive Services')      # PMO (Weekly) Company
    ftpt_offices = sorted(offices_all)                                               # FT / PT tables
    monthly_offices = field_offices                                                  # PMO (Monthly)

    emps = M.emps
    billable = [e for e in emps if e['billable'] == 'Billable']
    util_emps = sorted([e for e in emps if e['status'] == 'Active' and e['billable'] == 'Billable'],
                       key=lambda e: (e['office'].lower(), isinstance(e['id'], str), str(e['id']) if isinstance(e['id'], str) else e['id']))

    # ── office aggregates over Data blocks (Billable employees; scope FT/PT) ───────────────
    def agg(kind, idx, office=None, ftpt=None, pred=None):
        u = t = 0.0
        for e in billable:
            if office is not None and e['office'] != office:
                continue
            if ftpt is not None and e['ft'] != ftpt:
                continue
            if pred is not None and not pred(e):
                continue
            b = M.dv[str(e['id'])][kind]
            b = b[idx] if idx is not None else b
            u += b[0]; t += b[2]
        return u, t

    def util_pct(kind, idx, office=None, ftpt=None):
        u, t = agg(kind, idx, office, ftpt)
        return (u / t) if t else 'N/A'

    # monthly plan targets: hours-weighted billable targets (future: active roster weights)
    def plan_target(month_i, office=None):
        num_ = den = 0.0
        for e in billable:
            if office is not None and e['office'] != office:
                continue
            if MONTH_ENDS[month_i] <= eomonth(as_of):
                wgt = M.dv[str(e['id'])]['M'][month_i][2]
            else:
                wgt = (40.0 if e['ft'] == 'Full Time' else 20.0) if e['status'] == 'Active' else 0.0
            num_ += e['target'] * wgt; den += wgt
        return (num_ / den) if den else None

    def plan_target_ytd(office=None, pred=None):
        num_ = den = 0.0
        for e in billable:
            if office is not None and e['office'] != office:
                continue
            if pred is not None and not pred(e):
                continue
            wgt = M.dv[str(e['id'])]['Y'][2]
            num_ += e['target'] * wgt; den += wgt
        return (num_ / den) if den else None

    wb = openpyxl.Workbook(write_only=True)

    def sheet(name):
        return wb.create_sheet(name)

    def write_grid(ws, grid, ncols=None):
        """grid: {row(1-based): {col(0-based): value}} -> sequential appends."""
        if not grid:
            return
        last = max(grid)
        for r in range(1, last + 1):
            row = grid.get(r)
            if not row:
                ws.append([])
                continue
            width = max(row) + 1
            out = [None] * width
            for c, v in row.items():
                out[c] = v
            ws.append(out)

    # ═════════════════════════════ PMO (Monthly) (computed first; others reference it) ═══════
    months_done = [i for i, m in enumerate(MONTH_ENDS) if agg('M', i)[1] > 0]
    mon = {}                 # office -> {'t': [12], 'a': [12], 'ytd_t', 'ytd_a'}
    for o in monthly_offices + ['Total']:
        oo = None if o == 'Total' else o
        t = [plan_target(i, oo) for i in range(12)]
        a = [util_pct('M', i, oo) for i in range(12)]
        yt = plan_target_ytd(oo)
        ya = util_pct('Y', None, oo)
        mon[o] = {'t': t, 'a': a, 'yt': yt, 'ya': ya}
    g = {}
    g[2] = {}
    for i in range(12):
        g[2][1 + 3 * i] = MONTH_NAMES[i]
    g[2][37] = 'Total'
    g[3] = {0: 'PMO'}
    for i in range(13):
        g[3][1 + 3 * i] = 'Target'; g[3][2 + 3 * i] = 'Actual'; g[3][3 + 3 * i] = 'Δ'
    r = 4
    for o in monthly_offices + ['Total']:
        row = {0: o}
        d = mon[o]
        for i in range(12):
            tv, av = d['t'][i], d['a'][i]
            row[1 + 3 * i] = tv if tv is not None else '-'
            row[2 + 3 * i] = av
            if isinstance(av, float) and tv is not None:
                row[3 + 3 * i] = av - tv
        row[37] = d['yt'] if d['yt'] is not None else '-'
        row[38] = d['ya']
        if isinstance(d['ya'], float) and d['yt'] is not None:
            row[39] = d['ya'] - d['yt']
        g[r] = row
        r += 1
    total_row = r - 1
    below = [o for o in monthly_offices if isinstance(mon[o]['ya'], float) and mon[o]['yt'] is not None and mon[o]['ya'] < mon[o]['yt']]
    ya_t = mon['Total']['ya']; yt_t = mon['Total']['yt']
    g[r] = {38: ya_t * 100 if isinstance(ya_t, float) else None, 39: ((ya_t - yt_t) * -100) if isinstance(ya_t, float) else None}
    g[r + 1] = {0: '> = 10%', 39: len(below)}
    g[r + 2] = {0: '> = 0%'}
    g[r + 3] = {0: '< 0%'}
    g[r + 5] = {0: '* This does not intrinsically align with weekly results due to weekly treatment of 40 hours '
                   'base assumptions and hours worked over/under 40 base. As well as leave.'}
    monthly_grid = g
    month_target_ytd = {o: mon[o]['yt'] for o in mon}

    # ═════════════════════════════ PMO (Weekly) ═══════════════════════════════════════════════
    g = {}
    g[2] = {0: 'Headcound and weekly utilization.'}
    g[3] = {0: ' PMO Utilization Performance and Employee count per office through the week ending %s. '
               % ('%s %d' % (calendar.month_name[as_of.month], as_of.day))}
    g[4] = {108: 'YTD', 115: 'Wkly +', 116: 'Wkly -', 117: 'YTD +', 118: 'YTD -'}

    def hdr_row(a, b):
        row = {0: a, 1: b}
        for i, w in enumerate(WEEKS):
            row[2 + i] = w; row[55 + i] = w
        row[108] = 'Actual'; row[109] = 'Target '; row[110] = 'Var. '
        return row

    def headcount(office, ftpt, week_i, billable_only):
        n = 0
        for e in emps:
            if e['office'] != office or e['ft'] != ftpt:
                continue
            if billable_only and e['billable'] != 'Billable':
                continue
            if M.dv[str(e['id'])]['W'][week_i][2] > 0:
                n += 1
        return n

    def ft_target(office):
        vals = [e['target'] for e in util_emps if e['office'] == office and e['ft'] == 'Full Time']
        return (sum(vals) / len(vals)) if vals else '#DIV/0!'

    weekly_company_util = {}
    weekly_ytd = {}
    weekly_target = {}
    g[5] = hdr_row('Company', 'PMO')
    flags = []
    r = 6
    for o in company_offices:
        row = {0: 'Company' + o, 1: o}
        for i in range(52):
            row[2 + i] = headcount(o, 'Full Time', i, True) + headcount(o, 'Part Time', i, False)
            row[55 + i] = util_pct('W', i, o)
        weekly_company_util[o] = [row[55 + i] for i in range(52)]
        ya = util_pct('Y', None, o)
        tgt = month_target_ytd.get(o)
        row[108] = ya
        row[109] = tgt if tgt is not None else '-'
        row[110] = (ya - tgt) if (isinstance(ya, float) and tgt is not None) else 'N/A'
        weekly_ytd[o] = ya; weekly_target[o] = tgt
        if tgt is None:
            f = (0, 0, 0, 0)
        else:
            lw = row[55 + LW]
            f = ((1 if isinstance(lw, float) and lw >= tgt else 0), (1 if isinstance(lw, float) and lw < tgt else 0),
                 (1 if isinstance(ya, float) and ya >= tgt else 0), (1 if isinstance(ya, float) and ya < tgt else 0))
        flags.append(f)
        for j in range(4):
            row[115 + j] = f[j]
        g[r] = row
        r += 1
    for j in range(4):
        g[5][115 + j] = sum(f[j] for f in flags)
    tot = {1: 'Total'}
    for i in range(52):
        tot[2 + i] = sum(g[x][2 + i] for x in range(6, r))
        tot[55 + i] = util_pct('W', i)
    ya = util_pct('Y', None)
    tot[108] = ya; tot[109] = yt_t; tot[110] = (ya - yt_t) if isinstance(ya, float) else 'N/A'
    g[r] = tot
    company_total_row = r
    weekly_company_util['__total__'] = [tot[55 + i] for i in range(52)]
    below_target_count = g[5][118]

    def ftpt_table(start, ftpt):
        g[start - 1] = {0: ' ', 108: 'YTD'}
        g[start] = hdr_row(ftpt, 'Employee Count by Office')
        rr = start + 1
        for o in ftpt_offices:
            row = {1: o}
            for i in range(52):
                row[2 + i] = headcount(o, ftpt, i, ftpt == 'Full Time')
                row[55 + i] = util_pct('W', i, o, ftpt)
            ya_ = util_pct('Y', None, o, ftpt)
            tg = ft_target(o) if ftpt == 'Full Time' else 0.95
            row[108] = ya_
            row[109] = tg
            row[110] = (ya_ - tg) if (isinstance(ya_, float) and isinstance(tg, float)) else 'N/A'
            g[rr] = row
            rr += 1
        row = {1: 'Total'}
        for i in range(52):
            row[2 + i] = sum(g[x][2 + i] for x in range(start + 1, rr))
            row[55 + i] = util_pct('W', i, None, ftpt)
        ya_ = util_pct('Y', None, None, ftpt)
        if ftpt == 'Full Time':
            vals = [e['target'] for e in util_emps if e['ft'] == 'Full Time']
            tg = sum(vals) / len(vals)
        else:
            tg = 0.95
        row[108] = ya_; row[109] = tg; row[110] = (ya_ - tg) if isinstance(ya_, float) else 'N/A'
        g[rr] = row
        return rr
    end_ft = ftpt_table(39, 'Full Time')
    ftpt_table(end_ft + 3, 'Part Time')
    weekly_grid = g

    # ═════════════════════════════ Utilization (Employee) ═════════════════════════════════════
    # week columns (0-based): Q1 8..20, Q2 25..37, Q3 42..54, Q4 59..71; quarter cols 22/39/56/73
    QWEEK_START = [8, 25, 42, 59]
    QCOL = [22, 39, 56, 73]
    week_col = {}
    for q in range(4):
        for j in range(13):
            week_col[q * 13 + j] = QWEEK_START[q] + j
    WEEKNUM_LABELS = list(range(1, 14)) + [14, 14, 14] + list(range(15, 25)) + [27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39] + [39, 39] + list(range(40, 51))

    def emp_week_values(e):
        k = str(e['id'])
        out = []
        for i in range(52):
            w = WEEKS[i]
            val = M.dv[k]['W'][i][3]
            if e['status'] == 'Active' and is_platform(e):
                out.append(val)
            elif M.active_in_week(e, w) and w <= as_of:
                out.append(val)
            else:
                out.append('N/A')
        return out

    def avg_numeric(vals):
        nums = [v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
        return (sum(nums) / len(nums)) if nums else 'N/A'

    emp_rows = {}
    for e in emps:
        k = str(e['id'])
        wv = emp_week_values(e)
        qv = []
        for q in range(4):
            if e['ft'] == 'Full Time':
                qv.append(avg_numeric(wv[q * 13:(q + 1) * 13]))
            else:
                v = M.dv[k]['Q'][q][3]
                qv.append(v if isinstance(v, float) else 'N/A')
        if e['ft'] == 'Full Time':
            ytdv = avg_numeric(wv)
        else:
            v = M.dv[k]['Y'][3]
            ytdv = v if isinstance(v, (int, float)) else 'N/A'
        emp_rows[k] = {'w': wv, 'q': qv, 'ytd': ytdv}

    def util_sheet_header(g, shift, title_extra=None):
        g[2] = {0 + shift: '. ', 3 + shift: '%s Utilization Report' % COMPANY, 15 + shift: ' ', 45 + shift: ' '}
        g[3] = {1 + shift: ' ', 3 + shift: 'as of %s' % as_of.strftime('%m/%d/%Y')}
        g[4] = {3 + shift: 'Unadjusted Utilization - Billable Employees'}
        for q in range(4):
            g[4][QCOL[q] + shift] = 'Q%d %d' % (q + 1, FY)
        g[5] = {2 + shift: ' ', 5 + shift: 'Year-To-Date'}
        g[6] = {3 + shift: 'Reg/Temp', 4 + shift: 'Target %', 5 + shift: 'Actual %', 6 + shift: 'Deviation'}
        for i in range(52):
            g[5][week_col[i] + shift] = WEEKS[i]
            g[6][week_col[i] + shift] = WEEKNUM_LABELS[i]
        for q in range(4):
            g[6][QCOL[q] + shift] = 'Actual %%  Q%d' % (q + 1)
            g[6][QCOL[q] + 1 + shift] = 'Deviation Q%d' % (q + 1)

    def put_emp_values(row, e, shift):
        k = str(e['id'])
        er = emp_rows[k]
        tgt = e['target']
        row[5 + shift] = er['ytd']
        row[6 + shift] = (er['ytd'] - tgt) if isinstance(er['ytd'], float) else 'N/A'
        for i in range(52):
            row[week_col[i] + shift] = er['w'][i]
        for q in range(4):
            qv = er['q'][q]
            row[QCOL[q] + shift] = qv
            row[QCOL[q] + 1 + shift] = (qv - tgt) if isinstance(qv, float) else 'N/A'

    g = {}
    util_sheet_header(g, 0)
    # company rows = Data summary rows 1 (hourly), 3 (salary), 4 (total billable)
    def summary_unadj(kind, idx, pred):
        u = t = 0.0
        for e in emps:
            if not pred(e):
                continue
            b = M.dv[str(e['id'])][kind]
            b = b[idx] if idx is not None else b
            u += b[0]; t += b[2]
        return (u / t) if t else 0
    preds = {
        'Hourly ': lambda e: e['billable'] == 'Billable' and e['flsa'] == 'Hourly Non Exempt',
        'Salary': lambda e: e['billable'] == 'Billable' and e['flsa'] == 'Salary Exempt',
        '': lambda e: e['billable'] == 'Billable',
    }
    g[8] = {2: 'COMPANY - BILLABLE EMPLOYEES'}
    for rr, lab in ((9, 'Hourly '), (10, 'Salary'), (11, '')):
        row = {2: lab} if lab else {}
        pr = preds[lab]
        row[5] = summary_unadj('Y', None, pr)
        for i in range(52):
            row[week_col[i]] = summary_unadj('W', i, pr)
        for q in range(4):
            row[QCOL[q]] = summary_unadj('Q', q, pr)
        g[rr] = row
    g[13] = {2: 'COMPANY - BILLABLE EMPLOYEES (COST)'}
    g[14] = {2: 'Hourly '}
    g[15] = {2: 'Salary'}
    r = 18
    for e in util_emps:
        row = {0: e['office'], 1: eid_value(e), 2: e['roster_name'], 3: e['ft'], 4: e['target']}
        put_emp_values(row, e, 0)
        g[r] = row
        r += 1
    # formula rows with no employee behind them (the template's spare rows)
    n_pad = 60
    for _ in range(n_pad):
        row = {0: 0.0, 1: '', 2: 0.0, 3: 0.0, 4: 0.0, 5: 'N/A', 6: 'N/A'}
        for i in range(52):
            row[week_col[i]] = 'N/A'
        for q in range(4):
            row[QCOL[q]] = 'N/A'; row[QCOL[q] + 1] = 'N/A'
        g[r] = row
        r += 1
    # office summary block (fixed position in the template)
    r = max(r + 20, 700)
    for o in monthly_offices:
        row = {0: o}
        tgt = weekly_target.get(o)
        ya = weekly_ytd.get(o, util_pct('Y', None, o))
        row[4] = tgt if tgt is not None else '-'
        row[5] = ya
        row[6] = (ya - tgt) if (isinstance(ya, float) and tgt is not None) else 'N/A'
        wu = weekly_company_util.get(o) or [util_pct('W', i, o) for i in range(52)]
        for i in range(52):
            row[week_col[i]] = wu[i]
        for q in range(4):
            qv = util_pct('Q', q, o)
            row[QCOL[q]] = qv
            row[QCOL[q] + 1] = (qv - tgt) if (isinstance(qv, float) and tgt is not None) else 'N/A'
        g[r] = row
        r += 1
    util_grid = g

    # ═════════════════════════════ Non-Targeted ═══════════════════════════════════════════════
    g = {}
    util_sheet_header(g, 1)
    g[2][3] = ' ( 0 to +10% )'
    g[3][3] = ' ( +10% & HIGHER ) '
    g[4][3] = 'MISSING TIMESHEET'
    g[5][3] = ' '
    nt = [e for e in emps if e['status'] == 'Active' and e['billable'] != 'Billable']

    def nt_group(e):
        t = e['title'].lower()
        if any(k in t for k in ('chief', 'vice president', 'director', 'controller')):
            return 'EXECUTIVE TEAM'
        if any(k in t for k in ('billing', 'accountant', 'payroll')):
            return 'BILLING TEAM'
        return 'OTHER STAFF'
    r = 8
    for grp in ('EXECUTIVE TEAM', 'BILLING TEAM', 'OTHER STAFF'):
        members = sorted([e for e in nt if nt_group(e) == grp],
                         key=lambda e: (isinstance(e['id'], str), str(e['id']) if isinstance(e['id'], str) else e['id']))
        if not members:
            continue
        g[r] = {3: grp}
        r += 1
        for e in members:
            row = {0: '-', 1: e['status'], 2: eid_value(e), 3: e['roster_name'], 4: e['ft'], 5: 0}
            put_emp_values(row, e, 1)
            g[r] = row
            r += 1
        r += 1
    for _ in range(12):
        g[r] = {3: 0.0, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0}
        r += 1
    r += 10
    g[r] = {4: 'HOURS', 7: 'N/A'}
    for q in range(4):
        g[r][QCOL[q] + 1] = 'N/A'; g[r][QCOL[q] + 2] = 'N/A'
    nontarget_grid = g

    # ═════════════════════════════ Executive Summary ══════════════════════════════════════════
    comp_actual = mon['Total']['a']
    comp_target = mon['Total']['t']
    # prior-year (FY25) monthly actuals: the FY26 plan less the gap the business closed this year
    prnd = random.Random('%s-utilization-prior-year' % W.seed)
    prior = []
    for i in range(12):
        base = comp_target[i] if comp_target[i] is not None else comp_target[months_done[-1]]
        prior.append(round(base - prnd.uniform(0.070, 0.095), 4))
    last_m = months_done[-1]
    streak = 0
    for i in range(last_m, 0, -1):
        a, b = comp_actual[i], comp_actual[i - 1]
        if isinstance(a, float) and isinstance(b, float) and a > b:
            streak += 1
        else:
            break
    fy_plan = [t for t in comp_target if t is not None]
    fy_target = sum(fy_plan) / len(fy_plan)
    gap_pts = (fy_target - ya_t) * 100 if isinstance(ya_t, float) else 0
    words = {0: 'Holding steady', 1: 'Up one month', 2: 'Climbing two straight months', 3: 'Climbing three straight months',
             4: 'Climbing four straight months', 5: 'Climbing five straight months'}
    lead = words.get(streak, 'Climbing %d straight months' % streak)
    subtitle = ('%s. Still %.1f points behind the %.1f%% target.' % (lead, gap_pts, fy_target * 100)) if gap_pts > 0 \
        else ('%s. %.1f points ahead of the %.1f%% target.' % (lead, -gap_pts, fy_target * 100))
    ahead = sorted([o for o in company_offices if isinstance(weekly_ytd.get(o), float) and weekly_target.get(o)
                    and weekly_ytd[o] >= weekly_target[o]], key=lambda o: -(weekly_ytd[o] - weekly_target[o]))
    if ahead:
        names = ahead[:4]
        bright = 'Bright spots — %s %s at or above YTD target.' % (
            ', '.join(names[:-1]) + (' and ' if len(names) > 1 else '') + names[-1], 'are' if len(names) > 1 else 'is')
    else:
        bright = 'Bright spots — no office is at its YTD target yet; the closest are listed on the PMO tab.'
    g = {}
    title = 'E X E C U T I V E  S U M M A R Y    ·    F P & A    ·    %s\nCompany performance.' % (
        '  '.join(spaced(x) for x in as_of.strftime('%B %-d, %Y').upper().split(' ')))
    g[2] = {0: title}
    g[3] = {0: subtitle}
    g[6] = {7: 'YTD Utilization', 8: ya_t}
    g[7] = {7: 'FY26 Target', 8: round(fy_target, 3)}
    g[8] = {7: 'Variance', 8: (ya_t - round(fy_target, 3)) if isinstance(ya_t, float) else None}
    g[9] = {7: 'Below Target', 8: below_target_count}
    g[10] = {7: 'Prior Year'}
    g[15] = {1: 'M O N T H LY   T R E N D    ·    YTD', 4: 'THE BIG FOUR'}
    g[16] = {1: 'Actual vs target, FY26', 4: 'Utilization vs target, YTD'}
    g[33] = {1: bright}
    first_of_month = as_of.replace(day=1)
    g[54] = {0: as_of}
    g[55] = {0: dt.date(FY, 1, 1)}
    g[56] = {0: first_of_month}
    g[57] = {0: first_of_month - dt.timedelta(days=1)}
    g[58] = {0: str(FY)}
    g[59] = {0: as_of - dt.timedelta(days=6)}
    g[60] = {0: eomonth(as_of)}
    g[62] = {1: 'Actual', 2: 'Target', 3: 'Prior Year', 4: 'Current YOY Difference'}
    for i in range(12):
        a = comp_actual[i]
        row = {0: MON3[i], 2: comp_target[i], 3: prior[i]}
        if isinstance(a, float):
            row[1] = a
        else:
            row[1] = '' if i == 0 else '#N/A'
        if i == 0:
            row[4] = (ya_t - prior[as_of.month - 1]) * 100 if isinstance(ya_t, float) else None
        g[63 + i] = row
    g[76] = {0: 'Dicipline', 1: 'Actual', 2: 'Helper', 3: 'Target *Not Graphed Manual'}
    for j, disc in enumerate(('Domestic', 'International', 'Meridian', 'Geospatial')):
        u = t = 0.0
        for e in billable:
            if M.discipline(e) != disc:
                continue
            b = M.dv[str(e['id'])]['Y']
            u += b[0]; t += b[2]
        act = (u / t - 0.075) if t else 'N/A'
        tg = plan_target_ytd(pred=lambda e, d=disc: M.discipline(e) == d)
        g[77 + j] = {0: disc, 1: act, 2: 0.075, 3: tg}
    exec_grid = g

    # ═════════════════════════════ Summary - OT / Non-Posting (by project location) ════════════
    loc_label = {}
    for o in W.offices:
        loc = LOCATION_OF[o['name']]
        loc_label[loc] = 'International' if o['name'] == 'International' else loc
    active_locs = sorted([l for l, o in ((LOCATION_OF[o['name']], o) for o in W.offices) if o['ft'] + o['pt'] > 0],
                         key=lambda l: loc_label[l].lower())
    historical = sorted([LOCATION_OF[o['name']] for o in W.offices if o['ft'] + o['pt'] == 0])
    # aggregate Raw Data by (location, month)
    ot = defaultdict(lambda: defaultdict(float))       # loc -> key -> hours
    np_ = defaultdict(lambda: [0.0] * 12)
    util_by_loc = defaultdict(lambda: [0.0] * 12)
    tot_by_loc = defaultdict(float)
    util_tot_loc = defaultdict(float)
    for row in M.raw:
        loc, d, h, u, tt, task = row[2], row[3], row[7], row[5], row[9], row[10]
        # January's SUMIFS have no lower date bound (OT) or a Dec-1 bound (non-posting), so the
        # Dec 29-31 days of the first FY week land in January
        mi = d.month - 1 if d.year == FY else 0
        tot_by_loc[loc] += h
        util_tot_loc[loc] += u
        util_by_loc[loc][mi] += u
        if tt == 'OT':
            bt = row[13]
            key = {'Time and material': 'TM', 'Fixed fee': 'FF', 'Overhead': 'OH'}[bt]
            ot[loc][('ALL', mi)] += h
            ot[loc][(key, mi)] += h
            ot[loc][('ALL', 'T')] += h
            ot[loc][(key, 'T')] += h
        if task == NON_POSTING and d > dt.date(FY - 1, 12, 1):
            np_[loc][mi] += h

    g = {}
    g[1] = {0: COMPANY}
    g[2] = {0: 'Summary of Overtime by Project Location'}
    g[4] = {1: 'Overtime Summary: Total', 20: 'T&M Projects', 37: 'FF Projects', 54: 'Overhead'}
    g[6] = {0: 'Office Name'}
    anchors = [(1, 'ALL'), (21, 'TM'), (38, 'FF'), (55, 'OH')]
    for a, _ in anchors:
        if a > 1:
            g[6][a - 1] = 'Office Name'
        for i in range(12):
            g[6][a + i] = MONTH_ENDS[i]
        g[6][a + 12] = 'Total'
        g[6][a + 15] = '% of  \nTotal Hours'
    g[6][15] = 'Total Hours'
    g[6][18] = 'OT ChecL'
    g[6][71] = 'OT Breakout \nCheck'

    def ot_row(loc):
        lab = loc_label.get(loc, loc)
        row = {0: lab}
        for a, key in anchors:
            if a > 1:
                row[a - 1] = lab
            for i in range(12):
                row[a + i] = ot[loc][(key, i)]
            row[a + 12] = ot[loc][(key, 'T')]
            row[a + 15] = (ot[loc][(key, 'T')] / tot_by_loc[loc]) if tot_by_loc[loc] else '#DIV/0!'
        row[15] = tot_by_loc[loc]
        row[18] = ot[loc][('ALL', 'T')] - sum(ot[loc][('ALL', i)] for i in range(11))
        row[71] = ot[loc][('OH', 'T')] + ot[loc][('FF', 'T')] + ot[loc][('TM', 'T')] - ot[loc][('ALL', 'T')]
        return row
    r = 7
    for loc in active_locs:
        g[r] = ot_row(loc); r += 1
    row = {0: 'Total'}
    for c in list(range(1, 14)) + [15]:
        row[c] = sum(g[x].get(c, 0) or 0 for x in range(7, r) if isinstance(g[x].get(c, 0), (int, float)))
    row[16] = (row[13] / row[15]) if row[15] else 0
    g[r] = row
    r += 3
    g[r] = {0: 'Historical Offices', 18: 'OT ChecL', 20: 'Historical Offices', 37: 'Historical Offices',
            54: 'Historical Offices', 71: 'OT Check'}
    r += 1
    for loc in historical:
        g[r] = ot_row(loc); r += 1
    ot_grid = g

    g = {}
    g[1] = {0: COMPANY}
    g[2] = {0: 'Summary of OVHD Non-Posting Hours by Project Location', 33: dt.date(FY - 1, 12, 1)}
    g[4] = {0: 'Total OVHD Non-Posting Hours', 16: '% of Total Hours'}
    g[6] = {0: 'Office Name', 14: 'Total Hours', 16: 'Office Name', 30: '% of  \nTotal Hours'}
    for i in range(12):
        g[6][1 + i] = MONTH_ENDS[i]
        g[6][17 + i] = MONTH_ENDS[i]
    r = 7
    for loc in active_locs:
        lab = loc_label.get(loc, loc)
        row = {0: lab, 16: lab}
        for i in range(12):
            row[1 + i] = np_[loc][i]
            ub = util_by_loc[loc][i]
            row[17 + i] = (np_[loc][i] / ub) if ub else '#DIV/0!'
        row[14] = sum(np_[loc])
        row[30] = (row[14] / util_tot_loc[loc]) if util_tot_loc[loc] else '#DIV/0!'
        g[r] = row
        r += 1
    row = {0: 'Total', 16: 'Total'}
    for i in range(12):
        row[1 + i] = sum(np_[l][i] for l in active_locs)
        ub = sum(util_by_loc[l][i] for l in active_locs)
        row[17 + i] = (row[1 + i] / ub) if ub else '#DIV/0!'
    row[14] = sum(row[1 + i] for i in range(12))
    ut = sum(util_tot_loc[l] for l in active_locs)
    row[30] = row[14] / ut if ut else 0
    g[r] = row
    g[r + 2] = {9: sum(row[1 + i] for i in months_done) / max(1, len(months_done))}
    g[r + 3] = {0: 0}
    g[r + 4] = {0: '1-3%'}
    g[r + 5] = {0: 'above 3%'}
    np_grid = g

    # ═════════════════════════════ Summary - PTO ══════════════════════════════════════════════
    stp = defaultdict(float); ext = defaultdict(float)
    for row in M.raw:
        k = str(row[0])
        stp[k] += row[27]
        if row[9] in EXTENDED_TT:
            ext[k] += row[7]
    g = {}
    g[1] = {0: ' ', 12: 'Short Term PTO: Vacation, PTO, Sick, Jury Duty, Holiday.'}
    g[2] = {8: 'YTD %d' % FY, 12: 'Extended PTO: Short Term Disability, Maternity & Paternity Leave. '}
    g[3] = {0: 'ID #', 1: 'PMO', 2: 'Employee Name', 3: 'Type', 4: 'Hourly / Salary', 5: 'Utilization\n(YTD  Actual)',
            6: 'Utilization\n(Target)', 7: 'Target\nVariance', 8: 'Short Term\n PTO', 9: 'Extended\n PTO'}
    r = 4
    for e in sorted(util_emps, key=lambda e: (-stp[str(e['id'])], str(e['id']))):
        k = str(e['id'])
        er = emp_rows[k]
        g[r] = {0: eid_value(e), 1: e['office'], 2: e['roster_name'], 3: e['ft'], 4: e['flsa'], 5: er['ytd'],
                6: e['target'], 7: (er['ytd'] - e['target']) if isinstance(er['ytd'], float) else 'N/A',
                8: stp[k], 9: ext[k]}
        r += 1
    pto_grid = g

    # ═════════════════════════════ Pivot - Fleet ══════════════════════════════════════════════
    fleet = defaultdict(lambda: [0.0] * 12)
    for row in M.raw:
        if row[10] in FLEET_TASKS and row[3].year == FY:
            fleet[str(row[0])][row[3].month - 1] += row[7]
    fleet_months = [i for i in range(12) if any(v[i] for v in fleet.values())]
    fleet_months = list(range(0, max(fleet_months) + 1)) if fleet_months else []
    g = {}
    g[1] = {0: 'Task name', 1: '(Multiple Items)'}
    g[2] = {0: 'Fleet Time Charged YTD'}
    g[3] = {0: 'Sum of Hours', 4: 'Month'}
    g[4] = {0: 'Contact name', 1: 'Employee Home Location', 2: 'Role', 3: 'Supervisor'}
    for j, i in enumerate(fleet_months):
        g[4][4 + j] = MONTH_ENDS[i]
    tc = 4 + len(fleet_months)
    g[4][tc] = 'Total Hours\n(YTD)'
    r = 5
    for k in sorted(fleet, key=lambda k: (-sum(fleet[k]), M.emp[k]['name'])):
        e = M.emp[k]
        row = {0: e['name'], 1: e['office'], 2: e['title'], 3: e['supervisor'] or ''}
        for j, i in enumerate(fleet_months):
            if fleet[k][i]:
                row[4 + j] = fleet[k][i]
        row[tc] = sum(fleet[k])
        g[r] = row
        r += 1
    row = {0: 'Total Hours\n(YTD)'}
    for j, i in enumerate(fleet_months):
        row[4 + j] = sum(v[i] for v in fleet.values())
    row[tc] = sum(sum(v) for v in fleet.values())
    g[r] = row
    fleet_grid = g

    # ═════════════════════════════ PT Hours (+ on-call trend) ═════════════════════════════════
    pt_emps = sorted([e for e in emps if e['status'] == 'Active' and e['billable'] == 'Billable' and e['ft'] == 'Part Time'],
                     key=lambda e: (isinstance(e['id'], str), str(e['id']) if isinstance(e['id'], str) else e['id']))
    pt_weeks = list(range(0, LW + 1))
    last_entry = {}
    for row in M.raw:
        k = str(row[0])
        if k not in last_entry or row[3] > last_entry[k]:
            last_entry[k] = row[3]
    g = {}
    hdr = {0: 'EE', 1: 'Status', 2: 'Name', 3: 'Department', 4: 'Employee\nState', 5: 'Position', 6: 'Supervisor ',
           7: 'Target', 8: 'Hire/Rehire\nDate', 9: 'FT/PT', 10: 'Date of Last Time Submission (Hard)',
           11: 'Date of Last Time Submission ', 12: 'Hourly/Salary'}
    for c in range(13, 25):
        hdr[c] = 'Placeholder'
    for j, i in enumerate(pt_weeks):
        hdr[25 + j] = WEEKS[i]
    c0 = 25 + len(pt_weeks)
    hdr[c0] = '% Utilized\nPrior Week'; hdr[c0 + 1] = '4 Weeks'; hdr[c0 + 2] = '8 Weeks'; hdr[c0 + 3] = '12 Weeks'
    hdr[c0 + 4] = 'Prior Week'
    g[1] = hdr
    hard_cut = dt.date(FY, 3, 20)
    r = 2
    for e in pt_emps:
        k = str(e['id'])
        tots = [M.dv[k]['W'][i][2] for i in pt_weeks]
        le = last_entry.get(k)
        hard = min(le, hard_cut) if le else hard_cut
        row = {0: eid_value(e), 1: e['status'], 2: e['roster_name'], 3: e['office'],
               4: '-' if e['hire'] >= dt.date(FY, 1, 1) or is_platform(e) else M.office[e['office']]['state'],
               5: M.position(e), 6: e['supervisor'] or '', 7: e['target'],
               8: CONTRACTOR if is_platform(e) else e['hire'], 9: e['ft'], 10: hard,
               11: le if le else hard, 12: e['flsa']}
        for j, v in enumerate(tots):
            row[25 + j] = v
        adj = M.dv[k]['W'][LW][5]
        row[c0] = adj if isinstance(adj, float) else 0
        row[c0 + 1] = sum(tots[-4:]); row[c0 + 2] = sum(tots[-8:]); row[c0 + 3] = sum(tots[-12:]); row[c0 + 4] = tots[-1]
        g[r] = row
        r += 1
    r += 1
    g[r] = {9: 'Utilization of Available on Call Trend'}
    rows_oc = {'Total Hours': {}, '# of On Call Utilized': {}, '# of On Call Available': {}, '# Utilized / Available': {},
               'Utilized PT to Total Active Staff ': {}, '% PT vs. Total Hrs.': {}}
    for j, i in enumerate(pt_weeks):
        w = WEEKS[i]
        pts = [e for e in emps if e['ft'] == 'Part Time' and e['billable'] == 'Billable' and M.active_in_week(e, w)]
        th = sum(M.dv[str(e['id'])]['W'][i][2] for e in pts)
        nu = sum(1 for e in pts if M.dv[str(e['id'])]['W'][i][2] > 0)
        na = len(pts)
        allact = sum(1 for e in emps if M.active_in_week(e, w))
        allh = sum(M.dv[str(e['id'])]['W'][i][2] for e in emps)
        rows_oc['Total Hours'][25 + j] = th
        rows_oc['# of On Call Utilized'][25 + j] = nu
        rows_oc['# of On Call Available'][25 + j] = na
        rows_oc['# Utilized / Available'][25 + j] = nu / na if na else 0
        rows_oc['Utilized PT to Total Active Staff '][25 + j] = nu / allact if allact else 0
        rows_oc['% PT vs. Total Hrs.'][25 + j] = th / allh if allh else 0
    r += 1
    for lab in ('Total Hours', '# of On Call Utilized', '# of On Call Available', '# Utilized / Available'):
        g[r] = dict(rows_oc[lab]); g[r][9] = lab; r += 1
    r += 1
    for lab in ('Utilized PT to Total Active Staff ', '% PT vs. Total Hrs.'):
        g[r] = dict(rows_oc[lab]); g[r][9] = lab; r += 1
    pt_grid = g

    # ═════════════════════════════ Proj. Sector / Projects.csv ════════════════════════════════
    ptype = {'Class III Survey': 'Survey', 'Cultural Resources Monitoring': 'Monitoring', 'Phase I Survey': 'Survey',
             'Phase II Testing': 'Testing', 'Data Recovery': 'Data Recovery', 'Historic Resources Evaluation': 'Evaluation',
             'Paleontological Monitoring': 'Monitoring', 'Architectural History Survey': 'Survey',
             'Section 106 Consultation': 'Consulting', 'Cultural Resources Assessment': 'Cultural Consulting',
             'Archaeological Testing': 'Testing', 'Built Environment Survey': 'Survey',
             'Maritime Remote Sensing Survey': 'Survey', 'GIS Mapping': 'GIS', 'Curation': 'Curation',
             'Environmental Review Support': 'Environmental Consulting'}
    cust_id = {c: 10500 + i * 3 for i, c in enumerate(W.clients)}

    def proj_sector(p):
        h = int(hashlib.md5(p['id'].encode()).hexdigest(), 16)
        # the ERP's older jobs (pre-2024 legacy imports) often carry no sector
        if p['id'].startswith('23-') and h % 4 == 0:
            return ''
        return p['sector']
    proj_rows = []
    for p in sorted(W.projects, key=lambda p: p['id']):
        proj_rows.append([p['id'], p['name'], ptype.get(p['svc'], 'Consulting'), cust_id[p['client']], 'Contract',
                          p['status'], proj_sector(p), p['client']])
    with open(PROJECTS_CSV, 'w', newline='', encoding='utf-8-sig') as f:
        wcsv = csv.writer(f)
        wcsv.writerow(['Project ID', 'Project name', 'Project type', 'Customer ID', 'Project category',
                       'Project status', 'Project Sector', 'Customer Name'])
        for pr in proj_rows:
            wcsv.writerow(pr)
    # the workbook's own Proj. Sector tab only carries the current-year project list
    sector_rows = [pr for pr in proj_rows if pr[0].startswith('26-')]
    sector_lookup = {pr[0]: pr for pr in sector_rows}

    # ═════════════════════════════ Employee Dashboard ═════════════════════════════════════════
    ed_rows = []
    by_emp_week = defaultdict(lambda: defaultdict(lambda: [0.0] * 5))
    tasks_by_emp = defaultdict(lambda: defaultdict(float))
    proj_by_emp = defaultdict(lambda: defaultdict(float))
    ot_by_emp = defaultdict(float)
    week_idx = {w: i for i, w in enumerate(WEEKS)}
    for row in M.raw:
        k = str(row[0])
        wi = week_idx.get(row[18])
        u, n, h, task = row[5], row[6], row[7], row[10]
        if wi is not None:
            b = by_emp_week[k][wi]
            b[0] += u; b[1] += n; b[2] += h
            if task.lower() != 'overhead non-posting':
                b[3] += u
            else:
                b[4] += u
        tasks_by_emp[k][task] += h
        proj_by_emp[k][row[4]] += h
        if row[9] == 'OT':
            ot_by_emp[k] += h
    g = {}
    h1 = {0: 'Employee ID', 1: 'Utilized Hours per week', 53: 'Non-Hours per Week', 105: 'Total Hours per Week',
          157: 'Utilization', 158: 'Hours of Admin', 159: 'Hours of PTO', 160: 'Hours of S&M', 161: 'Hours of Overtime',
          162: 'Top Tasks by # of Hours', 178: 'Top six Projects\nBy Total Hours', 184: 'Top six Projects\nNames',
          190: 'Top six Projects\nSectors', 196: 'Utilized\nBillable', 248: 'Utilized\nUnbillable', 300: 'Admin',
          352: 'PTO', 404: 'S&M'}
    h2 = {}
    for base in (1, 53, 105, 196, 248, 300, 352, 404):
        for i in range(52):
            h2[base + i] = WEEKS[i]
    for c in (157, 158, 159, 160, 161):
        h2[c] = 'YTD'
    for kk in range(8):
        h2[162 + 2 * kk] = '#%d Name' % (kk + 1); h2[163 + 2 * kk] = '#%d Hours' % (kk + 1)
    for kk in range(6):
        h2[178 + kk] = '#%d' % (kk + 1); h2[184 + kk] = '#%d' % (kk + 1); h2[190 + kk] = '#%d' % (kk + 1)
    g[1] = h1; g[2] = h2
    r = 3
    for e in util_emps:
        k = str(e['id'])
        row = {0: eid_value(e)}
        bw = by_emp_week[k]
        for i in range(52):
            b = bw.get(i, [0.0] * 5)
            row[1 + i] = b[0]; row[53 + i] = b[1]; row[105 + i] = b[2]
            row[196 + i] = b[3]; row[248 + i] = b[4]
            row[300 + i] = 0; row[352 + i] = 0; row[404 + i] = 0
        row[157] = emp_rows[k]['ytd']
        row[158] = 0; row[159] = 0; row[160] = 0
        row[161] = ot_by_emp[k]
        tk = sorted(tasks_by_emp[k].items(), key=lambda x: (-x[1], x[0]))[:8]
        if not tk:
            row[162] = 'No time Found'
        for j, (tn, th) in enumerate(tk):
            row[162 + 2 * j] = tn; row[163 + 2 * j] = th
        pj = sorted(proj_by_emp[k].items(), key=lambda x: (-x[1], x[0]))[:6]
        for j, (pid, _) in enumerate(pj):
            row[178 + j] = pid
            if pid == 'OVHD':
                row[184 + j] = '#N/A'; row[190 + j] = '#N/A'
            else:
                sp = sector_lookup.get(pid)
                row[184 + j] = sp[1] if sp else '#N/A'
                row[190 + j] = (sp[6] or 0) if sp else '#N/A'
        g[r] = row
        r += 1
    # UNIQUE() over the Utilization (Employee) ID column also returns its blank as a 0 row
    blank = {0: 0}
    for base in (1, 53, 105, 196, 248, 300, 352, 404):
        for i in range(52):
            blank[base + i] = 0
    for c in (158, 159, 160, 161):
        blank[c] = 0
    g[r] = blank
    ed_grid = g

    # ═════════════════════════════ Data ═══════════════════════════════════════════════════════
    WEEK_COL0 = 18
    YTD_COL = 330
    Q_COLS = [336, 342, 348, 354, 360]
    M_COLS = [366 + 6 * i for i in range(13)]
    g = {}
    # summary rows 1-8
    groups = [(1, 'Billable', 'Hourly Non Exempt'), (2, 'Billable', 'Hourly Exempt'), (3, 'Billable', 'Salary Exempt'),
              (5, 'Non-Billable', 'Salary Exempt'), (6, 'Non-Billable', 'Hourly Non Exempt')]
    all_blocks = ([('W', i, WEEK_COL0 + 6 * i) for i in range(52)] + [('Y', None, YTD_COL)] + [('QS', None, Q_COLS[0])]
                  + [('Q', q, Q_COLS[q + 1]) for q in range(4)] + [('MS', None, M_COLS[0])]
                  + [('M', m, M_COLS[m + 1]) for m in range(12)])

    def grp_sum(kind, idx, bill, flsa):
        s = [0.0, 0.0, 0.0, 0.0]
        for e in emps:
            if e['billable'] != bill:
                continue
            if flsa is not None and e['flsa'] != flsa:
                continue
            b = M.dv[str(e['id'])][kind]
            b = b[idx] if idx is not None else b
            s[0] += b[0]; s[1] += b[1]; s[2] += b[2]; s[3] += b[4]
        return s

    def put_sum(row, col, s):
        row[col] = s[0]; row[col + 1] = s[1]; row[col + 2] = s[2]
        row[col + 3] = (s[0] / s[2]) if s[2] else 0
        row[col + 4] = s[3]
        row[col + 5] = (s[0] / (s[2] - s[3])) if (s[2] - s[3]) else 0

    for rr in range(1, 9):
        g[rr] = {}
    g[1].update({10: 'Billable Employees', 11: 'Hourly Non Exempt'})
    g[2].update({11: 'Hourly Exempt'})
    g[3].update({11: 'Salary Exempt'})
    g[4].update({10: 'Total Utilization, Billable'})
    g[5].update({10: 'Non-Billable Employees', 11: 'Salary Exempt'})
    g[6].update({11: 'Hourly Non Exempt'})
    g[7].update({10: 'Total Utilization, Non-billable'})
    g[8].update({10: 'Total Utilization, ALL'})
    for rr in (1, 2, 3):
        for c in range(12, 16):
            g[rr][c] = 0
    for kind, idx, col in all_blocks:
        sums_ = {}
        for rr, bill, flsa in groups:
            s = grp_sum(kind, idx, bill, flsa)
            sums_[rr] = s
            put_sum(g[rr], col, s)
        s4 = [sums_[1][j] + sums_[2][j] + sums_[3][j] for j in range(4)]
        s7 = [sums_[5][j] + sums_[6][j] for j in range(4)]
        s8 = [s4[j] + s7[j] for j in range(4)]
        put_sum(g[4], col, s4); put_sum(g[7], col, s7); put_sum(g[8], col, s8)
    g[9] = {11: dt.date(2023, 11, 25), 12: WEEK_STUB}
    for i in range(52):
        g[9][WEEK_COL0 + 6 * i] = WEEKS[i]
    g[9][YTD_COL] = 'YTD (Full Year Calculation)'
    g[9][Q_COLS[0]] = PERIOD_STUB
    for q in range(4):
        g[9][Q_COLS[q + 1]] = QUARTER_ENDS[q]
    g[9][M_COLS[0]] = PERIOD_STUB
    for m in range(12):
        g[9][M_COLS[m + 1]] = MONTH_ENDS[m]
    hdr = {0: 'EE', 1: 'Status', 2: 'Name', 3: 'Position', 4: 'Department', 5: 'Service Line', 6: 'Dicipline',
           7: 'Supervisor ', 8: 'Target', 9: 'FT/PT', 10: 'Hourly/Salary', 11: 'Billable/Non', 12: 'Utld', 13: 'NonUtld',
           14: 'Total', 15: 'Exclu.'}
    for kind, idx, col in all_blocks:
        for j, lab in enumerate(('Utld', 'NonUtld', 'Total', 'Unadj. %', 'Exclu.', 'Adj. %')):
            hdr[col + j] = lab
    g[10] = hdr
    r = 11
    for e in emps:
        k = str(e['id'])
        row = {0: eid_value(e), 1: M.data_status(e), 2: e['roster_name'], 3: M.position(e), 4: e['office'], 5: '-',
               6: M.discipline(e), 7: e['supervisor'] or '', 8: e['target'], 9: e['ft'], 10: e['flsa'],
               11: e['billable'], 12: 0, 13: 0, 14: 0, 15: 0}
        dv = M.dv[k]
        for kind, idx, col in all_blocks:
            b = dv[kind][idx] if idx is not None else dv[kind]
            for j in range(6):
                row[col + j] = b[j]
        g[r] = row
        r += 1
    data_grid = g

    # ═════════════════════════════ Current Employee Roster ════════════════════════════════════
    ucat = {'field': 'Field/Lab/Temporary', 'senior': 'Principal Investigator/Technical Lead',
            'lead': 'Office Leadership', 'admin': 'Corporate Non-billable'}
    roster = [e for e in sorted(W.employees, key=lambda e: str(e['id'])) if not is_platform(e)]
    g = {1: {0: 'Employee Number', 1: 'Full Name', 2: 'Status', 3: 'Work Email', 4: 'Department Code', 5: 'Department Name',
             6: 'Work Location', 7: 'Job Title', 8: 'Manager', 9: 'Employee Type', 10: 'Paygroup', 11: 'FLSA Status',
             12: 'Hire Date', 13: 'Rehire Date', 14: 'Termination Date', 15: 'Utilization category (Text)',
             16: 'Target (Text)', 17: 'Status Type', 18: 'State/Province'}}
    r = 2
    for e in roster:
        o = M.office[e['office']]
        active = e['status'] == 'Active'
        g[r] = {0: e['id'], 1: e['roster_name'], 2: e['status'], 3: e['email'] if active else ' ', 4: o['dept'],
                5: e['office'] if active else 'INACTIVE %s' % e['office'], 6: o['work_location'], 7: e['title'],
                8: e['supervisor'] or '', 9: 'Regular' if e['ft'] == 'Full Time' else 'Temporary', 10: e['paygroup'],
                11: e['flsa'] if active else ' ', 12: e['hire'], 14: e['term'],
                15: ucat.get(e['role'], ' ') if active else ' ', 16: int(round(e['target'] * 100)) if active else ' ',
                17: e['ft'], 18: o['state']}
        r += 1
    roster_grid = g

    # ═════════════════════════════ OTC-Project Profitability CM to ════════════════════════════
    orows, prev_end = otc_rows(W)
    g = {1: {0: 'CONS-OTC-Project Profitability CM to PM (In USD)'}}
    hdr_m = {7: ('Contract Value', 'Inception To Date', as_of, 'Actual'),
             8: ('Total Revised Budget', 'ITD As of', as_of, 'Actual'),
             10: ('Revenue Recognized', 'Month To Date', as_of, 'Actual MTD Revenue'),
             11: ('Cost of Sales', 'Month To Date', as_of, 'Actual MTD COS'),
             12: ('Wages', 'Month To Date', as_of, 'Actual MTD Wages'),
             13: ('Total Project Hours', 'Month To Date', as_of, 'Actual MTD Hrs'),
             14: ('Project Gross Profit', 'Month To Date', as_of, 'Actual MTD GP'),
             16: ('Revenue Recognized', 'Month Ending', prev_end, 'Actual'),
             17: ('Cost of Sales', 'Month Ending', prev_end, 'Actual'),
             18: ('Wages', 'Month Ending', prev_end, 'Actual'),
             19: ('Total Project Hours', 'Month Ending', prev_end, 'Actual'),
             20: ('Project Gross Profit', 'Month Ending', prev_end, 'Actual'),
             22: ('Revenue Recognized', 'ITD', as_of, 'Actual'),
             23: ('Cost of Sales', 'ITD', as_of, 'Actual'),
             24: ('Wages', 'ITD', as_of, 'Actual'),
             25: ('Project Gross Profit', 'ITD', as_of, 'Actual')}
    g[2] = {}; g[3] = {}; g[4] = {}; g[5] = {}
    for c, (a, b, d, cap) in hdr_m.items():
        g[2][c] = a; g[3][c] = b; g[4][c] = d; g[5][c] = cap
    g[5].update({0: 'Project ID', 1: 'Project Name', 2: 'Project Billing type', 3: 'Project status',
                 5: 'Observed percent ...', 6: 'Summary', 26: 'Location Name', 27: 'Auto Calculate % ...',
                 28: 'Project Manager Name'})
    g[6] = {0: 'All Projects', 1: 'All Projects'}

    def i_(v):
        return int(round(v))
    r = 7
    tots = defaultdict(float)
    for x in orows:
        ob = x['observed']
        row = {0: x['id'], 1: x['name'], 5: ob if ob else 0, 6: round(ob * 100, 1) if ob else 0,
               7: i_(x['contract']), 8: i_(x['revised_budget'])}
        if x['billing']:
            row[2] = x['billing']
        if x['status']:
            row[3] = x['status']
        vals = {10: x['rev_mtd'], 11: x['cos_mtd'], 12: x['wages_mtd'], 13: x['hrs_mtd'],
                16: x['rev_prev'], 17: x['cos_prev'], 18: x['wages_prev'], 19: x['hrs_prev'],
                22: x['rev_itd'], 23: x['cos_itd'], 24: x['wages_itd']}
        for c, v in vals.items():
            row[c] = i_(v)
        row[14] = i_(x['rev_mtd']) - i_(x['cos_mtd']) - i_(x['wages_mtd'])
        row[20] = i_(x['rev_prev']) - i_(x['cos_prev']) - i_(x['wages_prev'])
        row[25] = i_(x['rev_itd']) - i_(x['cos_itd']) - i_(x['wages_itd'])
        if x['location']:
            row[26] = x['location']
        row[27] = bool(x['auto'])
        if x['pm']:
            pm = x['pm']
            if ',' in pm:
                last, first = [s.strip() for s in pm.split(',', 1)]
                pm = '%s %s' % (first, last)
            row[28] = pm
        for c in list(range(5, 26)):
            if isinstance(row.get(c), (int, float)) and c != 5:
                tots[c] += row[c]
        tots[5] += row[5]
        g[r] = row
        r += 1
    tr = {0: 'Total All Projects', 1: 'Total All Projects'}
    for c, v in tots.items():
        tr[c] = v
    g[r] = tr
    g[r + 1] = dict(tr); g[r + 1][0] = 'All Locations'; g[r + 1][1] = 'All Locations'; g[r + 1].pop(5, None)
    otc_grid = g

    # ═════════════════════════════ write the workbook ═════════════════════════════════════════
    order = [('Executive Summary', exec_grid), ('PMO (Monthly)', monthly_grid), ('PMO (Weekly)', weekly_grid),
             ('Utilization (Employee)', util_grid), ('Specialized Reporting ->', {1: {0: ' '}}),
             ('Summary - OT', ot_grid), ('Summary - Non-Posting', np_grid), ('Summary - PTO', pto_grid),
             ('Pivot - Fleet', fleet_grid), ('Non-Targeted', nontarget_grid), ('PT Hours', pt_grid),
             ('OTC-Project Profitability CM to', otc_grid), ('Raw Data', None), ('Employee Dashboard', ed_grid),
             ('Backend ->', {}), ('Proj. Sector', None), ('Data', data_grid), ('Current Employee Roster', roster_grid)]
    for name, grid in order:
        ws = sheet(name)
        if name == 'Raw Data':
            ws.append(RAW_HEADER)
            for row in M.raw:
                ws.append([num(v) if isinstance(v, float) else v for v in row])
        elif name == 'Proj. Sector':
            ws.append(['Project ID', 'Project name', 'Project type', 'Customer ID', 'Project category',
                       'Project status', 'Project Sector', 'Customer Name'])
            for pr in sector_rows:
                ws.append(pr)
        else:
            write_grid(ws, grid)
    wb.save(XLSX)
    excel_style_package(XLSX)

    # ═════════════════════════════ retired tabs (snapshot as of the last export carrying them) ═
    write_retired(W, M)

    # ── headline numbers ──────────────────────────────────────────────────────────────────────
    backlog = sum(max(0.0, x['contract'] - x['rev_itd']) for x in orows
                  if str(x['status']).lower() in ('in progress', 'transferred entity') and x['id'] != 'OVHD')
    print('Wrote %s' % XLSX)
    print('  Raw Data rows: %d   employees in Data: %d   active billable: %d' % (len(M.raw), len(emps), len(util_emps)))
    print('  Company YTD utilization (billable): %.1f%%   FY plan %.1f%%   offices below YTD target: %d'
          % (ya_t * 100, fy_target * 100, below_target_count))
    print('  Positive backlog (In Progress + Transferred Entity): $%s' % format(round(backlog), ','))
    print('Wrote %s (%d projects)' % (PROJECTS_CSV, len(proj_rows)))
    print('Wrote %s' % RETIRED_XLSX)


# ─────────────────────────────────────────────────────────────────────────────────────────────
RETIRED_AS_OF = dt.date(2026, 8, 15)
PTO_TASKS = {'OVHD - Labor - Holiday', 'OVHD - Labor - Vacation (Salaried)', 'OVHD - Labor - PTO (Hourly Employee)',
             'OVHD - Labor - Sick (Salaried/Temps)', 'OVHD - Labor - Parental Leave'}
SM_TASKS = {'OVHD - Labor - Business Development', 'OVHD - Labor - Marketing', 'OVHD - Labor - Proposal Writing',
            'OVHD - Production - Proposal Editing'}


def write_retired(W, M):
    """'Summary - OVHD (FT Employee)' and 'Recovery Planning' as the last export carried them."""
    cut = RETIRED_AS_OF
    fte = [e for e in M.emps if e['status'] == 'Active' and e['billable'] == 'Billable' and e['ft'] == 'Full Time'
           and e['hire'] <= cut]
    cat = defaultdict(lambda: [0.0] * 5)          # billable, unbillable(non-posting), admin, pto, s&m
    for row in M.raw:
        if row[3] > cut:
            continue
        k = str(row[0])
        u, h, task = row[5], row[7], row[10]
        if u and u > 0:
            if task == NON_POSTING:
                cat[k][1] += u
            else:
                cat[k][0] += u
            continue
        if task in PTO_TASKS:
            cat[k][3] += h
        elif task in SM_TASKS:
            cat[k][4] += h
        elif task.startswith('OVHD'):
            cat[k][2] += h
    weeks_done = [w for w in WEEKS if w <= cut]
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet('Summary - OVHD (FT Employee)')
    ws.append([COMPANY])
    ws.append(['Overhead Summary — Full-Time Employees', None, None, None, None, 'YTD %d (through %s)' % (FY, cut.strftime('%m/%d/%Y'))])
    ws.append([])
    ws.append(['PMO', 'Employee Name', 'Position', 'Target', 'Target Variance', 'Utilized Billable', 'Utilized Unbillable',
               'Admin', 'PTO', 'S&M', None, 'Utilized Billable %', 'Utilized Unbillable %', 'Admin %', 'PTO %', 'S&M %'])
    for e in sorted(fte, key=lambda e: (e['office'], e['name'])):
        k = str(e['id'])
        c = cat[k]
        tot = sum(c)
        if tot <= 0:
            continue
        pct = [v / tot for v in c]
        ws.append([e['office'], e['roster_name'], e['title'], e['target'], pct[0] - e['target']] + c + [None] + pct)
    ws2 = wb.create_sheet('Recovery Planning')
    # custom scenario parameters live in col T (rows 2-3)
    ws2.append([COMPANY])
    row2 = [None] * 20; row2[18] = 'Custom Weeks'; row2[19] = 24
    row3 = [None] * 20; row3[18] = 'Custom % Total'; row3[19] = 0.95
    ws2.append(row2); ws2.append(row3)
    ws2.append([])
    ws2.append(['Office', 'ID', 'Employee Name', 'Hire Date', 'Weeks Available', 'YTD Actual', 'Target',
                '8 Week\n+10% (Required)', '8 Week\n+15% (Required)', '8 Week\n+20% (Required)', None,
                '12 Week\n+10% (Required)', '12 Week\n+15% (Required)', '12 Week\n+20% (Required)', None, None, None, None,
                'Custom\n(Required %)'])
    wd = len(weeks_done)
    for e in sorted(fte, key=lambda e: (e['office'], e['name'])):
        k = str(e['id'])
        vals = [v for w, v in zip(WEEKS, [M.dv[k]['W'][i][3] for i in range(52)]) if w <= cut and M.active_in_week(e, w)]
        nums = [v for v in vals if isinstance(v, float)]
        if not nums:
            continue
        ytd = sum(nums) / len(nums)
        if ytd >= e['target']:
            continue
        n = len(nums)

        def req(add, weeks):
            return ((ytd + add) * (n + weeks) - ytd * n) / weeks
        cw, cp = 24, 0.95
        ws2.append([e['office'], eid_value(e), e['roster_name'], e['hire'], n, ytd, e['target'],
                    req(0.10, 8), req(0.15, 8), req(0.20, 8), None, req(0.10, 12), req(0.15, 12), req(0.20, 12),
                    None, None, None, None, (cp * (n + cw) - ytd * n) / cw])
    wb.save(RETIRED_XLSX)
    excel_style_package(RETIRED_XLSX)


if __name__ == '__main__':
    main()
