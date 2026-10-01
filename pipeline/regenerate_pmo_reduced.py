#!/usr/bin/env python3
"""Compute PMO 'reduced' (billable-adjusted) weekly utilization and patch it into index.html.

WHAT THIS IS
------------
The main Utilization Report counts *all* utilized hours in the numerator.  FP&A's
secondary model ("2026 Utilization Report - V7.xlsx") also carries a reduced view that
strips Overhead Non-Posting hours out of utilized and treats them as non-utilized:

    Data!'Billable Hours'  =  Utld  -  'Non-Posting Hours'
    Non-Posting Hours      =  SUM('Raw Data'!Hours) where Task name = "Overhead Non-Posting"
                              for that employee, in that week
    reduced %              =  SUM(Billable Hours) / SUM(Total)
                              over the PMO's employees with Billable/Non = "Billable"

The weekly export has NO such columns, so this script derives them from 'Raw Data' +
'Data' every week.  Verified against V7 on 2026-08-21: the per-employee-week Non-Posting
values reproduce V7's own hidden column for every weekly block and every employee on YTD,
using an EXACT task-name match (V7's SUMIFS is exact, so regional variants such
as "CA - Overhead Non-Posting" are deliberately NOT stripped; --loose-task-match
includes them, which will NOT tie to V7).

WHAT IT WRITES
--------------
  site_data/PMO_Reduced.json   the payload (handy for eyeballing)
  index.html                   <script id="pmoReducedData"> payload swapped in, plus
                               (first run only) the export-wrapper script

The dashboard is untouched.  Two sheets of the exported workbook gain rows:

  PMO Utilization Weekly   one '<Office> - Billable' row under each PMO, on all three
                           tables (Company / Full Time / Part Time)
  Employee Utilization     two rows at the foot of each office's employee block --
                           '<Office> PMO Total - Utilized' and '... - Billable' --
                           carrying the same PMO numbers across all 52 week columns plus
                           Target, YTD, the four quarter columns and their deviations.
                           This mirrors V7's 'Utilization_Employee (Company)' rows 27/28,
                           which are themselves XLOOKUPs into the PMO weekly figures.

SELF-CHECK
----------
Before writing anything, three checks run and any failure aborts the write -- an export
layout shift otherwise produces quiet nonsense rather than an error:

  weekly     our UNADJUSTED percentages vs the cached values on the workbook's own
             'PMO (Weekly)' tab, cell for cell (every office x week x scope)
  monthly    same, against 'PMO (Monthly)' (every office x month)
  quarterly  each quarter block's Utld and Total must equal the sum of its three month
             blocks -- an internal identity between two independently detected series,
             so a mis-detected block series cannot slip through (office x scope x quarter)

Validated once by hand on 2026-08-21: the company-level Q1 unadjusted figure reproduces
'Utilization (Employee)'!W11 exactly, and the per-office Q1 figures reproduce V7's
'Utilization_Employee (Company)' row-27 quarter column for every office except those with a
known roster-snapshot difference between the two files.

USAGE
-----
  python3 regenerate_pmo_reduced.py
  python3 regenerate_pmo_reduced.py --workbook X.xlsx --index index.html
  python3 regenerate_pmo_reduced.py --check-only
"""
import argparse, datetime as dt, json, os, re, sys, zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
EPOCH = dt.date(1899, 12, 30)
NP_TASK = 'overhead non-posting'
HERE = os.path.dirname(os.path.abspath(__file__))


def col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def serial_to_date(v):
    return EPOCH + dt.timedelta(days=int(v))


class Workbook:
    """Minimal streaming reader for cached values. No openpyxl/calamine needed."""

    def __init__(self, path):
        self.path = path
        self.z = zipfile.ZipFile(path)
        wb = self.z.read('xl/workbook.xml').decode('utf-8')
        rels = self.z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        rmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        self.sheets = {}
        for name, rid in re.findall(r'<sheet[^>]*?name="([^"]*)"[^>]*?r:id="([^"]+)"', wb):
            name = (name.replace('&amp;', '&').replace('&lt;', '<')
                        .replace('&gt;', '>').replace('&quot;', '"').replace('&apos;', "'"))
            self.sheets[name] = 'xl/' + rmap[rid].lstrip('/')
        self.ss = []
        if 'xl/sharedStrings.xml' in self.z.namelist():
            with self.z.open('xl/sharedStrings.xml') as f:
                for _, el in ET.iterparse(f, events=('end',)):
                    if el.tag == NS + 'si':
                        self.ss.append(''.join(t.text or '' for t in el.iter(NS + 't')))
                        el.clear()

    def rows(self, sheet, max_row=None):
        with self.z.open(self.sheets[sheet]) as f:
            for _, el in ET.iterparse(f, events=('end',)):
                if el.tag != NS + 'row':
                    continue
                r = int(el.get('r'))
                d = {}
                for c in el.findall(NS + 'c'):
                    v = c.find(NS + 'v')
                    if v is None or v.text is None:
                        continue
                    col = ''.join(ch for ch in c.get('r') if ch.isalpha())
                    val = v.text
                    if c.get('t') == 's':
                        val = self.ss[int(val)]
                    elif c.get('t') in (None, 'n'):
                        try:
                            val = float(val)
                        except ValueError:
                            pass
                    d[col] = val
                el.clear()
                yield r, d
                if max_row and r >= max_row:
                    break

    def raw_stream(self, sheet):
        """Yield raw <row> byte blobs for the (very large) Raw Data sheet."""
        buf = b''
        with self.z.open(self.sheets[sheet]) as f:
            while True:
                chunk = f.read(1 << 23)
                if not chunk:
                    break
                buf += chunk
                parts = buf.split(b'</row>')
                buf = parts.pop()
                for p in parts:
                    yield p


def norm_id(v):
    if v is None:
        return None
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else str(v)
    s = str(v).strip()
    try:
        f = float(s)
        return str(int(f)) if f == int(f) else str(f)
    except ValueError:
        return s


# ------------------------------------------------------- Overhead Non-Posting hours
CELL_RE = re.compile(
    rb'<c r="([A-Z]+)\d+"([^>]*)>(?:<f[^>]*(?:/>|>.*?</f>))?(?:<v>([^<]*)</v>)?', re.S)


def nonposting_hours(bk, loose=False, raw_sheet='Raw Data'):
    """{(employee_id, entry_date_serial): hours} for Overhead Non-Posting rows.

    'Raw Data' columns used: A = Legacy ID, D = Entry date, H = Hours, K = Task name.
    """
    idxs = set()
    for i, s in enumerate(bk.ss):
        sl = s.strip().lower()
        if sl == NP_TASK or (loose and 'non-posting' in sl and len(sl) < 60):
            idxs.add(i)
    if not idxs:
        raise SystemExit("ABORT: no 'Overhead Non-Posting' task name in the workbook's shared "
                         "strings - has the ERP renamed the task?")
    needles = [('>%d</v>' % i).encode() for i in idxs]
    out = defaultdict(float)
    rows = hits = 0
    for row in bk.raw_stream(raw_sheet):
        rows += 1
        if not any(n in row for n in needles):
            continue
        emp = hrs = task = edate = None
        for m in CELL_RE.finditer(row):
            col = m.group(1).decode()
            attrs, val = m.group(2), m.group(3)
            if val is None:
                continue
            if col == 'A':
                emp = bk.ss[int(val)] if b't="s"' in attrs else val.decode()
            elif col == 'D':
                try:
                    edate = int(float(val))
                except ValueError:
                    pass
            elif col == 'H':
                try:
                    hrs = float(val)
                except ValueError:
                    pass
            elif col == 'K' and b't="s"' in attrs:
                try:
                    task = int(val)
                except ValueError:
                    task = None
        if task in idxs and emp is not None and hrs is not None and edate is not None:
            hits += 1
            out[(norm_id(emp), edate)] += hrs
    return dict(out), {'raw_rows': rows, 'nonposting_rows': hits,
                       'nonposting_hours': round(sum(out.values()), 2),
                       'task_match': 'loose' if loose else 'exact'}


# ------------------------------------------------------------------- Data sheet
def load_data(bk, sheet='Data'):
    """Return (blocks, employees, diagnostics).

    blocks: ordered {'start','utld','total','date','ytd'} derived from the 'Utld' labels
            on the header row and the dates on the row directly above it.
    """
    hdr_dates, hdr_labels, header, emps = {}, {}, None, []
    date_row = label_row = None
    for r, d in bk.rows(sheet):
        if header is None:
            labels = [v for v in d.values() if isinstance(v, str)]
            if 'EE' in labels and 'Department' in labels and 'Utld' in labels:
                header, label_row = d, r
                continue
            if sum(1 for v in d.values() if isinstance(v, float) and 40000 < v < 60000) >= 8:
                hdr_dates = {c: serial_to_date(v) for c, v in d.items()
                             if isinstance(v, float) and 40000 < v < 60000}
                hdr_labels = {c: v for c, v in d.items() if isinstance(v, str)}
                date_row = r
            continue
        if d.get('A') is None:
            continue
        emps.append(d)
    if header is None:
        raise SystemExit("ABORT: could not find the 'Data' header row (needs EE / Department / Utld).")
    if date_row is None or date_row != label_row - 1:
        raise SystemExit('ABORT: the date header row above Data row %s is missing or moved.' % label_row)

    starts = sorted([c for c, v in header.items() if isinstance(v, str) and v.strip() == 'Utld'],
                    key=col_num)
    blocks = []
    for c in starts:
        n = col_num(c)
        tot = header.get(col_letters(n + 2), '')
        if not isinstance(tot, str) or tot.strip() != 'Total':
            continue
        lbl = hdr_labels.get(c, '')
        blocks.append({'start': c, 'utld': c, 'total': col_letters(n + 2),
                       'date': hdr_dates.get(c),
                       'ytd': isinstance(lbl, str) and lbl.strip().upper().startswith('YTD')})
    return blocks, emps, {'header_row': label_row, 'employees': len(emps), 'blocks': len(blocks)}


def weekly_blocks(blocks):
    """The longest run of consecutive 7-day-apart dated blocks."""
    dated = [b for b in blocks if b['date'] and not b['ytd']]
    runs, cur = [], []
    for b in dated:
        if cur and (b['date'] - cur[-1]['date']).days == 7:
            cur.append(b)
        else:
            if len(cur) > len(runs):
                runs = cur
            cur = [b]
    if len(cur) > len(runs):
        runs = cur
    return runs


def _gap_kind(days):
    if 6 <= days <= 8:
        return 'weekly'
    if 26 <= days <= 35:
        return 'monthly'
    if 85 <= days <= 95:
        return 'quarterly'
    return None


def classify_series(blocks):
    """Split the dated period blocks into weekly / monthly / quarterly series.

    The Data sheet lays out several parallel series of 'Utld | NonUtld | Total | …' blocks —
    52 weekly, 12 monthly, 4 quarterly, plus a YTD block — and their column positions have
    moved before now, so nothing here is hardcoded.  Each *gap* between consecutive dated
    blocks is classified by length (7 / ~30 / ~91 days) and runs of same-class gaps become a
    series; the jump back in time between one series and the next classifies as None and so
    breaks the run.  The longest run of each class wins.
    """
    dated = sorted([b for b in blocks if b['date'] and not b['ytd']],
                   key=lambda b: col_num(b['start']))
    runs = []
    cur, cur_kind = [], None
    for i in range(len(dated) - 1):
        kind = _gap_kind((dated[i + 1]['date'] - dated[i]['date']).days)
        if kind and kind == cur_kind:
            cur.append(dated[i + 1])
        else:
            if cur_kind and len(cur) > 1:
                runs.append((cur_kind, cur))
            cur, cur_kind = ([dated[i], dated[i + 1]], kind) if kind else ([], None)
    if cur_kind and len(cur) > 1:
        runs.append((cur_kind, cur))

    out = {'weekly': [], 'monthly': [], 'quarterly': []}
    for kind, run in runs:
        if len(run) > len(out[kind]):
            out[kind] = run
    for kind in ('weekly', 'monthly', 'quarterly'):
        out[kind] = sorted(out[kind], key=lambda b: b['date'])
    ytds = [b for b in blocks if b['ytd']]
    out['ytd'] = ytds[0] if ytds else None
    return out


def period_ranges(run, kind):
    """[(block, lower_bound_exclusive, upper_bound_inclusive)] for one series.

    Data's own SUMIFS use `> previous block's date` / `<= this block's date`, so the first
    block of a run is a prior-period stub that only ever serves as a lower bound.
    """
    out = []
    for i, b in enumerate(run):
        if i == 0:
            continue                    # stub: 2025-12-27 weekly / 2025-12-28 monthly+quarterly
        out.append((b, run[i - 1]['date'], b['date']))
    if kind == 'weekly' and run:
        # weekly stubs are exactly 7 days back, so the first real week is safe to include
        first = run[0]
        out.insert(0, (first, first['date'] - dt.timedelta(days=7), first['date']))
    return out


def quarter_label(d):
    return 'Q%d' % ((d.month - 1) // 3 + 1)


# ------------------------------------------------------------------ aggregation
SCOPES = (('PMO', None), ('Full Time', 'Full Time'), ('Part Time', 'Part Time'))
TOTAL_KEY = '__TOTAL__'


def compute(bk, npm, blocks, emps):
    """(agg, series) where agg[(scope, office)][period_key] = [utld, total, nonposting]."""
    series = classify_series(blocks)
    weeks = period_ranges(series['weekly'], 'weekly')
    months = period_ranges(series['monthly'], 'monthly')
    quarters = period_ranges(series['quarterly'], 'quarterly')
    ytd = series['ytd']
    if len(weeks) < 40:
        raise SystemExit('ABORT: only %d weekly blocks found in Data (expected ~52).' % len(weeks))
    if len(quarters) < 4:
        raise SystemExit('ABORT: only %d quarterly blocks found in Data (expected 4).' % len(quarters))
    if ytd is None:
        raise SystemExit('ABORT: no YTD block found in Data.')

    np_by_emp = defaultdict(list)
    for (e, d), v in npm.items():
        np_by_emp[e].append((serial_to_date(d), v))

    every = ([('W', b, lo, hi) for b, lo, hi in weeks]
             + [('M', b, lo, hi) for b, lo, hi in months]
             + [('Q', b, lo, hi) for b, lo, hi in quarters])
    agg = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0, 0.0]))
    offices = set()
    for e in emps:
        office, billnon, ftpt = e.get('E'), e.get('L'), e.get('J')
        if not isinstance(office, str) or not isinstance(billnon, str):
            continue
        if billnon.strip() != 'Billable':
            continue
        office = office.strip()
        offices.add(office)
        nplist = np_by_emp.get(norm_id(e.get('A')), [])
        np_total = sum(v for _, v in nplist)
        scopes = [s for s, f in SCOPES
                  if f is None or (isinstance(ftpt, str) and ftpt.strip() == f)]
        for scope in scopes:
            for key in ((scope, office), (scope, TOTAL_KEY)):
                bucket = agg[key]
                for tag, b, lo, hi in every:
                    u, t = e.get(b['utld']), e.get(b['total'])
                    cell = bucket[tag + ':' + b['start']]
                    cell[0] += u if isinstance(u, float) else 0.0
                    cell[1] += t if isinstance(t, float) else 0.0
                    cell[2] += sum(v for d, v in nplist if lo < d <= hi)
                u, t = e.get(ytd['utld']), e.get(ytd['total'])
                cell = bucket['YTD']
                cell[0] += u if isinstance(u, float) else 0.0
                cell[1] += t if isinstance(t, float) else 0.0
                cell[2] += np_total
    series['weeks'] = weeks
    series['months'] = months
    series['quarters'] = quarters
    series['offices'] = sorted(offices)
    return agg, series


def pct(cell, reduced):
    u, t, n = cell
    if not t:
        return None
    return ((u - n) if reduced else u) / t


# ------------------------------------------------------------------- PMO targets
SCOPE_OF_LABEL = {'company': 'PMO', 'full time': 'Full Time', 'part time': 'Part Time'}


def read_targets(bk, sheet='PMO (Weekly)'):
    """{(scope, office): target} from the Target column of the workbook's own PMO (Weekly)."""
    rows = list(bk.rows(sheet, max_row=200))
    tgt_cols = []
    for r, d in rows:
        cand = [c for c, v in d.items() if isinstance(v, str) and v.strip().lower() == 'target']
        if cand:
            tgt_cols = sorted(cand, key=col_num)
            break
    if not tgt_cols:
        return {}
    tcol = tgt_cols[-1]
    out, scope = {}, None
    for r, d in rows:
        a, b = d.get('A'), d.get('B')
        if isinstance(a, str) and a.strip().lower() in SCOPE_OF_LABEL:
            scope = SCOPE_OF_LABEL[a.strip().lower()]
            continue
        if isinstance(a, str) and a.startswith('Company'):
            scope = 'PMO'
        if scope is None or not isinstance(b, str) or not b.strip():
            continue
        v = d.get(tcol)
        if isinstance(v, float):
            key = TOTAL_KEY if b.strip().lower() == 'total' else b.strip()
            out[(scope, key)] = v
    return out


# ---------------------------------------------------------------------- self-check
def _header_week_cols(rows, sheet):
    """The utilisation block's {column: date} on a PMO sheet header row.

    Both PMO tabs put a headcount block of week dates first and the utilisation block
    second, so the LAST contiguous run of date columns is the one we want.
    """
    for r, d in rows:
        dates = {c: v for c, v in d.items() if isinstance(v, float) and 40000 < v < 60000}
        if len(dates) < 20:
            continue
        cols = sorted(dates, key=col_num)
        runs, cur = [], []
        for c in cols:
            if cur and col_num(c) == col_num(cur[-1]) + 1:
                cur.append(c)
            else:
                if cur:
                    runs.append(cur)
                cur = [c]
        if cur:
            runs.append(cur)
        if len(runs) >= 2:
            return {c: serial_to_date(dates[c]) for c in runs[-1]}
        return {c: serial_to_date(dates[c]) for c in runs[0]}
    return {}


def _walk_scoped(rows):
    """Yield (scope, office_key, row_dict) for the Company / Full Time / Part Time sections."""
    scope = None
    for r, d in rows:
        a, b = d.get('A'), d.get('B')
        if isinstance(a, str) and a.strip().lower() in SCOPE_OF_LABEL:
            scope = SCOPE_OF_LABEL[a.strip().lower()]
            continue
        if isinstance(a, str) and a.startswith('Company'):
            scope = 'PMO'
        if scope is None or not isinstance(b, str) or not b.strip():
            continue
        if b.strip().lower() == 'employee count by office':
            continue
        key = TOTAL_KEY if b.strip().lower() == 'total' else b.strip()
        yield scope, key, d


def check_weekly(bk, agg, weeks, sheet='PMO (Weekly)'):
    rows = list(bk.rows(sheet, max_row=200))
    hdr = _header_week_cols(rows, sheet)
    if not hdr:
        return {'name': 'weekly vs PMO (Weekly)', 'status': 'skipped', 'checked': 0,
                'matched': 0, 'diffs': [], 'reason': 'utilisation header not found'}
    by_date = {hi: b for b, lo, hi in weeks}
    checked = matched = 0
    diffs = []
    for scope, key, d in _walk_scoped(rows):
        bucket = agg.get((scope, key))
        if not bucket:
            continue
        for c, wdate in hdr.items():
            want, b = d.get(c), by_date.get(wdate)
            if not isinstance(want, float) or b is None:
                continue
            got = pct(bucket['W:' + b['start']], reduced=False)
            if got is None:
                continue
            checked += 1
            if abs(got - want) < 1e-9:
                matched += 1
            elif len(diffs) < 10:
                diffs.append({'scope': scope, 'office': key, 'period': wdate.isoformat(),
                              'workbook': want, 'computed': got})
    return {'name': 'weekly vs PMO (Weekly)',
            'status': 'ok' if checked and matched == checked else 'MISMATCH',
            'checked': checked, 'matched': matched, 'diffs': diffs}


def check_monthly(bk, agg, months, sheet='PMO (Monthly)'):
    """PMO (Monthly) is Target/Actual/Δ triples per month, then a final YTD triple."""
    rows = list(bk.rows(sheet, max_row=80))
    hdr_row = None
    for r, d in rows:
        if any(isinstance(v, str) and v.strip() == 'Actual' for v in d.values()):
            hdr_row = (r, d); break
    if not hdr_row:
        return {'name': 'monthly vs PMO (Monthly)', 'status': 'skipped', 'checked': 0,
                'matched': 0, 'diffs': [], 'reason': 'no Actual columns found'}
    hr, hd = hdr_row
    acts = sorted([c for c, v in hd.items() if isinstance(v, str) and v.strip() == 'Actual'],
                  key=col_num)[:len(months)]
    checked = matched = 0
    diffs = []
    for r, d in rows:
        if r <= hr:
            continue
        name = d.get('A')
        if not isinstance(name, str) or not name.strip():
            continue
        key = TOTAL_KEY if name.strip().lower() == 'total' else name.strip()
        bucket = agg.get(('PMO', key))
        if not bucket:
            continue
        for i, c in enumerate(acts):
            want = d.get(c)
            if not isinstance(want, float):
                continue
            got = pct(bucket['M:' + months[i][0]['start']], reduced=False)
            if got is None:
                continue
            checked += 1
            if abs(got - want) < 1e-9:
                matched += 1
            elif len(diffs) < 10:
                diffs.append({'scope': 'PMO', 'office': key,
                              'period': months[i][0]['date'].isoformat(),
                              'workbook': want, 'computed': got})
    return {'name': 'monthly vs PMO (Monthly)',
            'status': 'ok' if checked and matched == checked else 'MISMATCH',
            'checked': checked, 'matched': matched, 'diffs': diffs}


def check_quarterly(agg, months, quarters):
    """Each quarter's Utld/Total must equal the sum of the month blocks inside it."""
    checked = matched = 0
    diffs = []
    for qb, qlo, qhi in quarters:
        inside = [m for m, lo, hi in months if qlo < hi <= qhi]
        if len(inside) != 3:
            return {'name': 'quarterly == sum(monthly)', 'status': 'skipped', 'checked': 0,
                    'matched': 0, 'diffs': [],
                    'reason': 'quarter %s covers %d month blocks, expected 3'
                              % (qhi.isoformat(), len(inside))}
        for (scope, key), bucket in agg.items():
            q = bucket['Q:' + qb['start']]
            su = sum(bucket['M:' + m['start']][0] for m in inside)
            st = sum(bucket['M:' + m['start']][1] for m in inside)
            checked += 1
            if abs(q[0] - su) < 1e-6 and abs(q[1] - st) < 1e-6:
                matched += 1
            elif len(diffs) < 10:
                diffs.append({'scope': scope, 'office': key, 'period': qhi.isoformat(),
                              'workbook': round(su, 2), 'computed': round(q[0], 2)})
    return {'name': 'quarterly == sum(monthly)',
            'status': 'ok' if checked and matched == checked else 'MISMATCH',
            'checked': checked, 'matched': matched, 'diffs': diffs}


# ------------------------------------------------------------------------ payload
def build_payload(agg, series, targets, diag, source, loose):
    weeks, quarters = series['weeks'], series['quarters']
    week_meta = [{'iso': hi.isoformat(), 'label': hi.strftime('%b %-d')} for b, lo, hi in weeks]
    quarter_meta = [{'label': quarter_label(hi), 'iso': hi.isoformat()} for b, lo, hi in quarters]
    scopes = {}
    for scope, _ in SCOPES:
        entry = {}
        for office in list(series['offices']) + [TOTAL_KEY]:
            bucket = agg.get((scope, office))
            if not bucket:
                continue
            wk, wk_u = {}, {}
            for b, lo, hi in weeks:
                cell = bucket['W:' + b['start']]
                r, u = pct(cell, True), pct(cell, False)
                if r is not None:
                    wk[hi.isoformat()] = round(r, 10)
                if u is not None:
                    wk_u[hi.isoformat()] = round(u, 10)
            qs = {}
            for b, lo, hi in quarters:
                cell = bucket['Q:' + b['start']]
                r, u = pct(cell, True), pct(cell, False)
                if r is None and u is None:
                    continue
                qs[quarter_label(hi)] = {
                    'reduced': None if r is None else round(r, 10),
                    'unadjusted': None if u is None else round(u, 10),
                }
            y = pct(bucket['YTD'], True)
            yu = pct(bucket['YTD'], False)
            if not wk and y is None:
                continue
            name = 'Total' if office == TOTAL_KEY else office
            entry[name] = {
                'weeks': wk,
                'weeks_unadjusted': wk_u,
                'quarters': qs,
                'ytd': None if y is None else round(y, 10),
                'ytd_unadjusted': None if yu is None else round(yu, 10),
                'target': targets.get((scope, office)),
                'nonposting_hours_ytd': round(bucket['YTD'][2], 2),
            }
        scopes[scope] = entry
    asof = None
    tot = agg.get(('PMO', TOTAL_KEY), {})
    for b, lo, hi in weeks:
        if tot.get('W:' + b['start'], [0, 0, 0])[1]:
            asof = hi
    return {
        'generated_at': dt.datetime.now().isoformat(timespec='seconds'),
        'source': os.path.basename(source),
        'as_of': asof.isoformat() if asof else None,
        'rule': ('reduced pct = (utilized hours - Overhead Non-Posting hours) / total hours, '
                 'over Billable employees; task-name match: '
                 + ('loose (any *non-posting*)' if loose else 'exact "Overhead Non-Posting"')),
        'row_label_suffix': ' — Billable',
        'employee_row_labels': {'unadjusted': 'PMO Total — Utilized',
                                'reduced': 'PMO Total — Billable'},
        'weeks': week_meta,
        'quarters': quarter_meta,
        'scopes': scopes,
        'diagnostics': diag,
    }


# -------------------------------------------------------------------- index.html
PAYLOAD_RE = re.compile(
    r'(<script id="pmoReducedData" type="application/json">)(.*?)(</script>)', re.S)
WRAPPER_MARK = 'PMO reduced (billable-adjusted) rows for the Excel export'

WRAPPER_JS = r'''
<script id="pmoReducedData" type="application/json">__PAYLOAD__</script>
<script>
/* PMO reduced (billable-adjusted) rows for the Excel export -- EXPORT ONLY.
   Wraps window.exportReport.  Immediately before the workbook is built we add rows to two
   sections, let the existing DOM-driven exporter pick them up, then remove them again:

     PMO_Weekly            one '<Office> - Billable' row under each PMO row (all 3 tables)
     Utilization_Employee  two rows at the foot of each office's employee block,
                           '<Office> PMO Total - Utilized' and '... - Billable'

   The dashboard itself is never changed.  Data comes from the pmoReducedData payload above,
   refreshed weekly by regenerate_pmo_reduced.py. */
(function () {
  var DATA = null;
  try {
    var el = document.getElementById('pmoReducedData');
    if (el) DATA = JSON.parse(el.textContent || el.innerText || 'null');
  } catch (e) { DATA = null; }
  if (!DATA || !DATA.scopes) return;

  var SUFFIX = DATA.row_label_suffix || ' -- Billable';
  var EMP_LABELS = DATA.employee_row_labels ||
      { unadjusted: 'PMO Total -- Utilized', reduced: 'PMO Total -- Billable' };
  var SCOPE_ORDER = ['PMO', 'Full Time', 'Part Time'];
  var TINT_UNADJ = '#F0EBE4';
  var TINT_RED = '#FBEEDC';
  var labelToIso = {};
  (DATA.weeks || []).forEach(function (w) { labelToIso[String(w.label).toUpperCase()] = w.iso; });  /* keys upper-cased: on the ACTIVE tab the header cells render uppercase via CSS     text-transform and innerText reflects that, so a case-sensitive match silently     skipped the visible table (fixed 2026-09-11) */

  function pct0(v) { return (v * 100).toFixed(0) + '%'; }
  function pct1(v) { return (v * 100).toFixed(1) + '%'; }
  function parsePct(txt) {
    var m = (txt || '').replace(/,/g, '').match(/(-?\d+(?:\.\d+)?)\s*%/);
    return m ? parseFloat(m[1]) / 100 : null;
  }
  function textOf(cell) {
    if (!cell) return '';
    return (cell.innerText || cell.textContent || '').replace(/\s+/g, ' ').trim();
  }
  function pill(td, val, target, decimals) {
    /* exact value + tenths format for the Excel export (FP&A, 2026-09-14); the pill text below
       stays rounded the way the table shows it */
    td.dataset.val = val;
    td.dataset.fmt = '0.0%';
    var span = document.createElement('span');
    span.className = 'pill ' + (target != null && val >= target ? 'good' : 'bad');
    span.textContent = decimals ? pct1(val) : pct0(val);
    td.appendChild(span);
  }
  function delta(td, val, target) {
    if (val == null || target == null) return;
    var d = (val - target) * 100;
    var s = document.createElement('span');
    s.className = 'delta ' + (d >= 0 ? 'pos' : 'neg');
    /* match the employee grid's own glyphs: '+1.9' / U+2212 '0.9', so the column stays text */
    s.textContent = d >= 0 ? '+' + d.toFixed(1) : '\u2212' + Math.abs(d).toFixed(1);
    td.appendChild(s);
  }

  /* ---------------- PMO Utilization Weekly: one paired row per PMO ---------------- */
  function classifyWeekly(head) {
    var kinds = [];
    for (var c = 0; c < head.cells.length; c++) {
      var h = textOf(head.cells[c]);
      var isoH = labelToIso[h.toUpperCase()];
      if (isoH) kinds.push({ k: 'week', iso: isoH });
      else if (/^ytd/i.test(h)) kinds.push({ k: 'ytd' });
      else if (/^target/i.test(h)) kinds.push({ k: 'target' });
      else if (/^(variance|var\.|Δ)/i.test(h)) kinds.push({ k: 'var' });
      else kinds.push({ k: c === 0 ? 'label' : 'skip' });
    }
    return kinds;
  }

  function injectPmoWeekly() {
    var section = document.getElementById('PMO_Weekly');
    if (!section) return [];
    var added = [];
    Array.prototype.slice.call(section.querySelectorAll('table')).forEach(function (table, ti) {
      var scope = DATA.scopes[SCOPE_ORDER[ti]] || DATA.scopes[SCOPE_ORDER[0]];
      var head = table.tHead && table.tHead.rows[0];
      if (!scope || !head || !table.tBodies[0]) return;
      var kinds = classifyWeekly(head);
      if (!kinds.some(function (k) { return k.k === 'week'; })) return;
      Array.prototype.slice.call(table.tBodies[0].rows).forEach(function (row) {
        if (row.getAttribute('data-reduced-row')) return;
        var name = textOf(row.cells[0]);
        var rec = scope[name];
        if (!rec) return;
        var target = null;
        for (var i = 0; i < kinds.length && i < row.cells.length; i++) {
          if (kinds[i].k === 'target') target = parsePct(textOf(row.cells[i]));
        }
        if (target == null && rec.target != null) target = rec.target;
        var tr = document.createElement('tr');
        tr.setAttribute('data-reduced-row', '1');
        tr.className = row.className;
        for (var j = 0; j < kinds.length; j++) {
          var src = row.cells[j];
          var td = document.createElement('td');
          td.className = src ? src.className : '';
          var kind = kinds[j].k, val = null;
          if (kind === 'label') {
            td.textContent = name + SUFFIX;
            td.style.fontStyle = 'italic';
          } else if (kind === 'week') {
            val = rec.weeks ? rec.weeks[kinds[j].iso] : null;
            if (val != null) pill(td, val, target, false);
          } else if (kind === 'ytd') {
            val = rec.ytd;
            if (val != null) pill(td, val, target, true);
          } else if (kind === 'target') {
            td.textContent = target == null ? '' : pct1(target);
            if (target != null) { td.dataset.val = target; td.dataset.fmt = '0.0%'; }
          } else if (kind === 'var') {
            if (rec.ytd != null && target != null) {
              var dv = (rec.ytd - target) * 100;
              var sd = document.createElement('span');
              sd.className = 'delta ' + (dv >= 0 ? 'pos' : 'neg');
              sd.textContent = (dv >= 0 ? '+' : '') + dv.toFixed(1) + 'pts';
              td.appendChild(sd);
            }
          }
          td.style.backgroundColor = TINT_RED;
          tr.appendChild(td);
        }
        if (row.parentNode) row.parentNode.insertBefore(tr, row.nextSibling);
        added.push(tr);
      });
    });
    return added;
  }
'''

WRAPPER_JS += r'''
  /* --------- Employee Utilization: PMO Utilized + Billable at the foot of each office ------ */
  function classifyEmployee(table) {
    var head = table.tHead;
    if (!head || head.rows.length < 1) return null;
    var top = head.rows[0];
    var weekRow = head.rows.length > 1 ? head.rows[1] : head.rows[0];
    var kinds = [];
    var lead = 0;
    for (var c = 0; c < top.cells.length; c++) {
      var cell = top.cells[c];
      if ((cell.rowSpan || 1) < 2 && head.rows.length > 1) continue;   // quarter group headers
      var h = textOf(cell).toLowerCase();
      var k = 'skip';
      if (/office/.test(h)) k = 'office';
      else if (/employee|name/.test(h)) k = 'name';
      else if (/type/.test(h)) k = 'type';
      else if (/target/.test(h)) k = 'target';
      else if (/^ytd/.test(h)) k = 'ytd';
      else if (h === 'δ' || /^var/.test(h)) k = 'ytddelta';
      kinds.push({ k: k });
      lead++;
    }
    if (head.rows.length > 1) {
      var lastQ = null;
      for (var w = 0; w < weekRow.cells.length; w++) {
        var t = textOf(weekRow.cells[w]);
        var isoT = labelToIso[t.toUpperCase()];
        if (isoT) kinds.push({ k: 'week', iso: isoT });
        else if (/^Q\d$/i.test(t)) { lastQ = t.toUpperCase(); kinds.push({ k: 'qtotal', q: lastQ }); }
        else if (t === 'Δ' || /^var/i.test(t)) kinds.push({ k: 'qdelta', q: lastQ });
        else kinds.push({ k: 'skip' });
      }
    }
    return { kinds: kinds, lead: lead };
  }

  function buildEmployeeTotalRow(office, rec, kinds, template, reduced) {
    var target = rec.target;
    var tr = document.createElement('tr');
    tr.setAttribute('data-reduced-row', '1');
    tr.className = template ? template.className : '';
    var tint = reduced ? TINT_RED : TINT_UNADJ;
    for (var j = 0; j < kinds.length; j++) {
      var src = template ? template.cells[j] : null;
      var td = document.createElement('td');
      td.className = src ? src.className : '';
      var kind = kinds[j].k, val = null;
      if (kind === 'office') {
        td.textContent = office;
      } else if (kind === 'name') {
        td.textContent = reduced ? EMP_LABELS.reduced : EMP_LABELS.unadjusted;
      } else if (kind === 'type') {
        td.textContent = '';
      } else if (kind === 'target') {
        td.textContent = target == null ? '' : pct1(target);
      } else if (kind === 'ytd') {
        val = reduced ? rec.ytd : rec.ytd_unadjusted;
        if (val != null) pill(td, val, target, true);
      } else if (kind === 'ytddelta') {
        delta(td, reduced ? rec.ytd : rec.ytd_unadjusted, target);
      } else if (kind === 'week') {
        var src2 = reduced ? rec.weeks : rec.weeks_unadjusted;
        val = src2 ? src2[kinds[j].iso] : null;
        if (val != null) pill(td, val, target, false);
      } else if (kind === 'qtotal') {
        var q = rec.quarters && kinds[j].q ? rec.quarters[kinds[j].q] : null;
        val = q ? (reduced ? q.reduced : q.unadjusted) : null;
        if (val != null) pill(td, val, target, false);
      } else if (kind === 'qdelta') {
        var q2 = rec.quarters && kinds[j].q ? rec.quarters[kinds[j].q] : null;
        delta(td, q2 ? (reduced ? q2.reduced : q2.unadjusted) : null, target);
      }
      td.style.backgroundColor = tint;
      td.style.fontWeight = '600';
      tr.appendChild(td);
    }
    return tr;
  }

  function injectEmployee() {
    var section = document.getElementById('Utilization_Employee');
    var table = section && section.querySelector('table');
    if (!table || !table.tBodies[0]) return [];
    var info = classifyEmployee(table);
    if (!info || !info.kinds.some(function (k) { return k.k === 'week'; })) return [];
    var scope = DATA.scopes[SCOPE_ORDER[0]];
    if (!scope) return [];

    // Group the employee rows by office, in document order.
    var groups = [], byOffice = {};
    Array.prototype.slice.call(table.tBodies[0].rows).forEach(function (row) {
      if (row.getAttribute('data-reduced-row')) return;
      var office = row.getAttribute('data-office');
      if (!office) return;
      if (!byOffice[office]) { byOffice[office] = { office: office, last: row, rows: 0 }; groups.push(byOffice[office]); }
      byOffice[office].last = row;
      byOffice[office].rows++;
    });

    var added = [];
    groups.forEach(function (g) {
      var rec = scope[g.office];
      if (!rec) return;
      var unadj = buildEmployeeTotalRow(g.office, rec, info.kinds, g.last, false);
      var red = buildEmployeeTotalRow(g.office, rec, info.kinds, g.last, true);
      var parent = g.last.parentNode;
      if (!parent) return;
      parent.insertBefore(unadj, g.last.nextSibling);
      parent.insertBefore(red, unadj.nextSibling);
      added.push(unadj, red);
    });
    return added;
  }

  function strip() {
    Array.prototype.slice.call(document.querySelectorAll('tr[data-reduced-row]'))
      .forEach(function (tr) { if (tr.parentNode) tr.parentNode.removeChild(tr); });
  }

  var orig = window.exportReport;
  if (typeof orig !== 'function') return;
  window.exportReport = function () {
    var added = [];
    try { added = added.concat(injectPmoWeekly()); }
    catch (e) { console.warn('reduced-row inject (PMO weekly) failed', e); }
    try { added = added.concat(injectEmployee()); }
    catch (e) { console.warn('reduced-row inject (employee) failed', e); }
    try { return orig.apply(this, arguments); }
    finally { setTimeout(strip, added.length ? 1500 : 0); }
  };
})();
</script>
'''


def patch_index(index_path, payload_json, check_only=False):
    with open(index_path, encoding='utf-8') as f:
        html = f.read()
    if PAYLOAD_RE.search(html):
        new, n = PAYLOAD_RE.subn(lambda m: m.group(1) + payload_json + m.group(3), html, count=1)
        assert n == 1
        action = 'payload refreshed'
        if 'injectEmployee' not in html:
            # An older wrapper is installed: replace the whole two-script block, payload
            # included.  The block runs from the payload <script> to the </script> that closes
            # the wrapper (the mark sits INSIDE the wrapper, so the first </script> after it is
            # the one we want -- do not hop to the next).
            i = html.find('<script id="pmoReducedData"')
            mk = html.find(WRAPPER_MARK, i if i >= 0 else 0)
            k = html.find('</script>', mk) if mk > 0 else -1
            if i >= 0 and mk > i and k > mk:
                end = k + len('</script>')
                new = html[:i] + WRAPPER_JS.replace('__PAYLOAD__', payload_json) + html[end:]
                action = 'wrapper UPGRADED (employee rows added) + payload refreshed'
            else:
                print('WARNING: an older wrapper is installed but its bounds could not be '
                      'located, so only the payload was refreshed. The Employee Utilization '
                      'rows will NOT appear. Restore a clean index.html and rerun.')
    else:
        if WRAPPER_MARK in html:
            raise SystemExit('ABORT: the export wrapper is present but the payload script is not - '
                             'index.html was hand-edited; fix that before rerunning.')
        i = html.rfind('</body>')
        if i < 0:
            raise SystemExit('ABORT: no </body> in index.html.')
        new = html[:i] + WRAPPER_JS.replace('__PAYLOAD__', payload_json) + '\n' + html[i:]
        action = 'wrapper + payload installed'
    if check_only:
        return action, len(new), None
    bak = index_path + '.bak.' + dt.datetime.now().strftime('%Y%m%d-%H%M%S')
    with open(bak, 'w', encoding='utf-8') as f:
        f.write(html)
    tmp = index_path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        f.write(new)
    os.replace(tmp, index_path)
    return action, len(new), bak


# --------------------------------------------------------------------------- main
def find_workbook():
    for c in (os.path.join(HERE, 'Utilization Report.xlsx'),
              os.path.join(HERE, 'Weekly and Monthly Updates', 'Utilization',
                           'Utilization Report.xlsx'),
              os.path.join(os.getcwd(), 'Utilization Report.xlsx')):
        if os.path.exists(c):
            return c
    raise SystemExit('ABORT: could not find "Utilization Report.xlsx". Pass --workbook.')


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument('--workbook')
    ap.add_argument('--index', default=os.path.join(HERE, 'index.html'))
    ap.add_argument('--out-json')
    ap.add_argument('--check-only', action='store_true')
    ap.add_argument('--loose-task-match', action='store_true',
                    help='also strip "CA - Overhead Non-Posting" variants (will NOT tie to V7)')
    a = ap.parse_args(argv)

    src = a.workbook or find_workbook()
    print('Workbook: %s' % src)
    bk = Workbook(src)
    for need in ('Data', 'Raw Data', 'PMO (Weekly)'):
        if need not in bk.sheets:
            raise SystemExit('ABORT: sheet %r missing from the workbook.' % need)

    npm, np_diag = nonposting_hours(bk, loose=a.loose_task_match)
    print('Raw Data: %(raw_rows)s rows, %(nonposting_rows)s Overhead Non-Posting rows, '
          '%(nonposting_hours)s hrs (%(task_match)s match)' % np_diag)

    blocks, emps, d_diag = load_data(bk)
    print('Data: header row %(header_row)s, %(employees)s employee rows, %(blocks)s period blocks'
          % d_diag)

    agg, series = compute(bk, npm, blocks, emps)
    print('Series: %d weeks (%s .. %s), %d months, %d quarters   Offices: %d'
          % (len(series['weeks']), series['weeks'][0][2], series['weeks'][-1][2],
             len(series['months']), len(series['quarters']), len(series['offices'])))

    targets = read_targets(bk)
    print('Targets read from PMO (Weekly): %d office/scope pairs' % len(targets))

    checks = [check_weekly(bk, agg, series['weeks']),
              check_monthly(bk, agg, series['months']),
              check_quarterly(agg, series['months'], series['quarters'])]
    failed = False
    for c in checks:
        if c['status'] == 'skipped':
            print('Self-check SKIPPED  %-28s %s' % (c['name'], c.get('reason', '')))
        else:
            print('Self-check %-10s %-28s %d/%d cells'
                  % (c['status'], c['name'], c['matched'], c['checked']))
        if c['status'] == 'MISMATCH':
            failed = True
            for d in c['diffs']:
                print('   %(scope)s / %(office)s / %(period)s: workbook %(workbook)r vs computed %(computed)r' % d)
    if failed:
        raise SystemExit('ABORT: the unadjusted figures no longer reproduce the workbook. The '
                         'export layout probably changed - see CLAUDE.md before overriding.')

    diag = dict(np_diag)
    diag.update(d_diag)
    diag['selfchecks'] = [{k: c[k] for k in ('name', 'status', 'checked', 'matched')} for c in checks]
    payload = build_payload(agg, series, targets, diag, src, a.loose_task_match)
    pj = json.dumps(payload, separators=(',', ':'), ensure_ascii=False)

    out_json = a.out_json or os.path.join(HERE, 'site_data', 'PMO_Reduced.json')
    if not a.check_only:
        os.makedirs(os.path.dirname(out_json), exist_ok=True)
        with open(out_json, 'w', encoding='utf-8') as f:
            json.dump(payload, f, indent=1, ensure_ascii=False)
        print('Wrote %s (%.1f KB payload)' % (out_json, len(pj) / 1024))

    action, size, bak = patch_index(a.index, pj, check_only=a.check_only)
    print('index.html: %s -> %s bytes%s'
          % (action, size, '' if not bak else '  (backup %s)' % os.path.basename(bak)))
    if a.check_only:
        print('--check-only: nothing was written.')

    pmo = payload['scopes'].get('PMO', {})
    print('\nPMO   YTD unadj. ->  reduced   Q1..Q4 reduced            non-posting hrs YTD')
    for name in sorted(pmo, key=lambda n: (n == 'Total', n)):
        r = pmo[name]
        if r['ytd'] is None:
            continue
        qs = ' '.join(('%5.1f%%' % (r['quarters'][q]['reduced'] * 100))
                      if r['quarters'].get(q) and r['quarters'][q]['reduced'] is not None else '    - '
                      for q in ('Q1', 'Q2', 'Q3', 'Q4'))
        print('  %-22s %6.1f%% -> %6.1f%%   %s   %9.1f'
              % (name, (r['ytd_unadjusted'] or 0) * 100, r['ytd'] * 100, qs,
                 r['nonposting_hours_ytd']))
    return 0


if __name__ == '__main__':
    sys.exit(main())
