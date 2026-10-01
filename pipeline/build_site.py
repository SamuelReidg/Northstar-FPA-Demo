"""Build a polished single-file HTML site from extracted xlsx data."""
import json, html as html_lib, os, datetime
from pathlib import Path

DATA_DIR = Path('site_data')
OUT_FILE = 'Northstar Heritage Utilization Site.html'

TABS = [
    {'name': 'Executive Summary',           'slug': 'Executive_Summary',        'group': 'Overview',  'renderer': 'exec_summary'},
    {'name': 'PMO Dashboard',               'slug': 'PMO_Dashboard',            'group': 'Overview',  'renderer': 'pmo_dashboard'},
    {'name': 'Employee Dashboard',          'slug': 'Employee_Dashboard',       'group': 'Overview',  'renderer': 'employee_dashboard'},
    {'name': 'PMO Utilization: Monthly',    'slug': 'PMO_Monthly',              'group': 'Summaries', 'renderer': 'pmo_period'},
    {'name': 'PMO Utilization: Weekly',     'slug': 'PMO_Weekly',               'group': 'Summaries', 'renderer': 'pmo_weekly'},
    {'name': 'Employee Utilization',        'slug': 'Utilization_Employee',     'group': 'Summaries', 'renderer': 'bubble_dashboard'},
    {'name': 'Overtime',                    'slug': 'Summary_OT',               'group': 'Specialized Reporting', 'renderer': 'overtime'},
    {'name': 'Non-Posting Hours',           'slug': 'Summary_Non_Posting',      'group': 'Specialized Reporting', 'renderer': 'non_posting'},
    {'name': 'Overhead',                    'slug': 'Summary_OVHD_FT_Employee', 'group': 'Specialized Reporting', 'renderer': 'overhead'},
    {'name': 'Paid Time Off',               'slug': 'Summary_PTO',              'group': 'Specialized Reporting', 'renderer': 'pto'},
    {'name': 'Recovery Planning',           'slug': 'Recovery_Planning',        'group': 'Specialized Reporting', 'renderer': 'recovery'},
    {'name': 'Proposal Hours',              'slug': 'Pivot_Proposal',           'group': 'Specialized Reporting', 'renderer': 'proposal_hours'},
    {'name': 'Fleet Management',            'slug': 'Pivot_Fleet',              'group': 'Specialized Reporting', 'renderer': 'pivot_hours'},
    {'name': 'Non-Targeted Employees',      'slug': 'Non_Targeted',             'group': 'Specialized Reporting', 'renderer': 'non_targeted'},
    {'name': 'Part Time Employee Activity', 'slug': 'PT_Hours',                 'group': 'Specialized Reporting', 'renderer': 'pt_activity'},
    {'name': 'On Call Utilization Trend',   'slug': 'On_Call_Trend',            'group': 'Specialized Reporting', 'renderer': 'on_call_trend'},
]

# ---------- helpers ----------
def esc(s):
    return html_lib.escape(str(s)) if s is not None else ''

def pn(v):
    """Parse number or None."""
    if v is None: return None
    if isinstance(v, bool): return None
    if isinstance(v, (int, float)): return float(v)
    if isinstance(v, str):
        s = v.strip().rstrip('%').replace(',', '').replace('$', '')
        if s in ('', 'N/A', '-', '#N/A', '#REF!', '#DIV/0!', '#VALUE!', '#NAME?', 'N/T'): return None
        try: return float(s)
        except: return None
    return None

def is_date(v):
    return isinstance(v, dict) and '__date__' in v

def to_date(v):
    if not is_date(v): return None
    try:
        return datetime.datetime.fromisoformat(v['__date__'])
    except: return None

def fmt_date(v, short=False):
    d = to_date(v)
    if not d: return ''
    if short: return d.strftime('%b %-d')
    return d.strftime('%m/%d/%Y')

def fmt_month(v):
    d = to_date(v)
    if not d: return ''
    return d.strftime("%b '%y")

def fmt_pct(v, decimals=0):
    n = pn(v)
    if n is None:
        if v in (None, ''): return ''
        return esc(str(v))
    return f'{n*100:.{decimals}f}%'

def fmt_num(v, decimals=1):
    n = pn(v)
    if n is None:
        if v in (None, ''): return ''
        return esc(str(v))
    if n == int(n): return f'{int(n):,}'
    return f'{n:,.{decimals}f}'

def fmt_int(v):
    n = pn(v)
    if n is None:
        if v in (None, ''): return ''
        return esc(str(v))
    return f'{int(n):,}'

def fmt_money(v):
    n = pn(v)
    if n is None:
        if v in (None, ''): return ''
        return esc(str(v))
    sign = '-' if n < 0 else ''
    return f'{sign}${abs(n):,.0f}'

def pill(value, kind, *, formatter=fmt_pct):
    """Return pill HTML. kind is 'good', 'bad', 'na', 'neutral'."""
    txt = formatter(value)
    if not txt: return ''
    return f'<span class="pill {kind}">{txt}</span>'

def pill_cmp(value, target, *, formatter=fmt_pct, allow_neutral=False):
    """Pill based on comparison to target."""
    v = pn(value); t = pn(target)
    if v is None: return pill(value, 'na', formatter=formatter) if value not in (None, '') else ''
    if t is None: return pill(value, 'neutral' if allow_neutral else 'na', formatter=formatter)
    return pill(value, 'good' if v >= t else 'bad', formatter=formatter)

# Source sheets retired from the export (2026-07/08): `Summary - OVHD (FT Employee)`
# and `Recovery Planning`. extract_data.py no longer emits their JSON, and this build
# used to die on the missing file, forcing whoever ran the refresh to hunt down
# placeholder copies. The merge keeps the LIVE section for both tabs anyway
# (KEEP_ENTIRE in merge.py), so an empty placeholder here is thrown away — it exists
# only so the build can finish. Any OTHER missing tab still fails loudly, as it should.
RETIRED_SLUGS = {'Summary_OVHD_FT_Employee', 'Recovery_Planning'}

def load_tab(slug):
    path = DATA_DIR / f'{slug}.json'
    if slug in RETIRED_SLUGS and not path.exists():
        print(f'NOTE: {path} missing (retired source sheet) -> empty placeholder. '
              'The merge keeps the live section for this tab, so this is expected.')
        return {'rows': []}
    with open(path) as f:
        return json.load(f)

def find_row_with(rows, keyword, start=0):
    """Find first row index where any cell contains the keyword (case-insensitive)."""
    kw = keyword.lower()
    for i in range(start, len(rows)):
        for v in rows[i]:
            if isinstance(v, str) and kw in v.lower():
                return i
    return -1

def non_empty(row):
    return [v for v in row if v not in (None, '')]


# ---------- renderers ----------

def render_bubble_dashboard(_data):
    """The Utilization (Employee) bubble dashboard."""
    with open('employees.json') as f:
        edata = json.load(f)
    employees = edata['employees']
    date_headers = edata['date_headers']
    WEEKLY_Q1 = list(range(9, 22)); WEEKLY_Q2 = list(range(26, 39))
    WEEKLY_Q3 = list(range(43, 56)); WEEKLY_Q4 = list(range(60, 73))
    QUARTERLY = [23, 40, 57, 74]
    quarters_def = [('Q1', WEEKLY_Q1, 23), ('Q2', WEEKLY_Q2, 40), ('Q3', WEEKLY_Q3, 57), ('Q4', WEEKLY_Q4, 74)]
    week_labels = {c: fmt_date({'__date__': date_headers[str(c)]+ 'T00:00:00'}, short=True) for c in WEEKLY_Q1+WEEKLY_Q2+WEEKLY_Q3+WEEKLY_Q4}
    offices = sorted(set(e['office'] for e in employees if e['office']))
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Employee Utilization</h2>')
    p.append(f'<p class="page-sub">Weekly performance for {len(employees)} billable employees across {len(offices)} offices, pilled green when at or above each employee\'s individual target.</p>')
    p.append('</div>')
    p.append('<div class="toolbar">')
    p.append('<input type="text" id="bubbleSearch" class="input" placeholder="Search employee or ID…" />')
    p.append('<select id="bubbleOffice" class="input"><option value="">All offices</option>')
    for o in offices: p.append(f'<option value="{esc(o)}">{esc(o)}</option>')
    p.append('</select>')
    p.append('<select id="bubbleStatus" class="input"><option value="">All status</option>')
    p.append('<option value="bad">Has below-target weeks</option>')
    p.append('<option value="good">Has at/above-target weeks</option>')
    p.append('</select>')
    p.append('<span class="legend">')
    p.append('<span class="pill good">≥ target</span>')
    p.append('<span class="pill bad">below target</span>')
    p.append('<span class="pill na">N/A</span>')
    p.append('</span>')
    p.append('</div>')
    p.append(f'<div class="result-count"><span id="bubbleCount">{len(employees)}</span> of {len(employees)} employees</div>')
    p.append('<div class="scroll-card"><table class="bubble-grid"><thead>')
    p.append('<tr class="row-quarter">')
    p.append('<th class="col-office" rowspan="2">Office</th>')
    p.append('<th class="col-name" rowspan="2">Employee</th>')
    p.append('<th class="col-target" rowspan="2">Target</th>')
    p.append('<th class="ytd-sep" rowspan="2">YTD</th>')
    for q_label, weeks, _ in quarters_def:
        span = len(weeks) + 1
        p.append(f'<th class="quarter-sep" colspan="{span}">{q_label} 2026</th>')
    p.append('</tr><tr class="row-week">')
    for q_label, weeks, _ in quarters_def:
        for i, c in enumerate(weeks):
            cls = 'quarter-sep' if i == 0 else ''
            p.append(f'<th class="{cls}">{week_labels[c]}</th>')
        p.append(f'<th>{q_label} Tot</th>')
    p.append('</tr></thead><tbody>')
    current_office = None
    for emp in employees:
        office = emp['office'] or ''
        if office != current_office:
            current_office = office
            total_cols = 4 + sum(len(w)+1 for _, w, _ in quarters_def)
            p.append(f'<tr class="office-group" data-office-group="{esc(office)}"><td colspan="{total_cols}">{esc(office.upper())}</td></tr>')
        name = emp['name'] or ''; emp_id = emp['id'] or ''
        target = emp['target']
        has_bad = False; has_good = False
        for c in WEEKLY_Q1+WEEKLY_Q2+WEEKLY_Q3+WEEKLY_Q4+QUARTERLY:
            val = emp['weeks'].get(str(c), emp['quarters'].get(str(c)))
            v = pn(val); t = pn(target)
            if v is not None and t is not None:
                if v >= t: has_good = True
                else: has_bad = True
        ytd_pill = pill_cmp(emp['ytd'], target)
        p.append(f'<tr class="emp-row" data-office="{esc(office)}" data-name="{esc(str(name).lower())}" data-id="{esc(str(emp_id).lower())}" data-has-bad="{int(has_bad)}" data-has-good="{int(has_good)}">')
        p.append(f'<td class="col-office">{esc(office)}</td>')
        p.append(f'<td class="col-name">{esc(name)}<br><span class="emp-id">{esc(str(emp_id))}</span></td>')
        t_disp = fmt_pct(target) if pn(target) is not None else '—'
        p.append(f'<td class="col-target">{t_disp}</td>')
        p.append(f'<td class="cell ytd-sep">{ytd_pill}</td>')
        for q_label, weeks, q_col in quarters_def:
            for i, c in enumerate(weeks):
                cls_cell = 'cell' + (' quarter-sep' if i == 0 else '')
                val = emp['weeks'].get(str(c))
                p.append(f'<td class="{cls_cell}">{pill_cmp(val, target)}</td>')
            val = emp['quarters'].get(str(q_col))
            p.append(f'<td class="cell">{pill_cmp(val, target)}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def _exec_compute_below_target():
    """Count offices below target using latest available month from PMO Monthly."""
    try:
        with open(DATA_DIR / 'PMO_Monthly.json') as f:
            pmo = json.load(f)
        rows = pmo['rows']
        if len(rows) < 4: return None, None
        period_row = rows[1]
        period_blocks = []
        c = 1
        while c < len(period_row):
            v = period_row[c]
            if v not in (None, '') and isinstance(v, (str, dict)):
                period_blocks.append(c); c += 3
            else: c += 1
        # Find latest period with actuals
        latest_block = None
        for col_idx in period_blocks:
            has_any = False
            for i in range(3, len(rows)):
                r = rows[i]
                if col_idx+1 < len(r) and pn(r[col_idx+1]) is not None:
                    has_any = True; break
            if has_any: latest_block = col_idx
        if latest_block is None: return None, None
        below = 0; total = 0
        for i in range(3, len(rows)):
            r = rows[i]
            office = r[0] if r else ''
            if not office or not isinstance(office, str): continue
            if 'total' in office.lower() or 'company' in office.lower(): continue
            t_val = pn(r[latest_block] if latest_block < len(r) else None)
            a_val = pn(r[latest_block+1] if latest_block+1 < len(r) else None)
            if t_val is None or a_val is None: continue
            total += 1
            if a_val < t_val: below += 1
        return below, total
    except Exception:
        return None, None


def _exec_line_chart(monthly, title_left='Actual', title_right='Target'):
    """SVG line chart: Actual (black) + Target (gray) over 12 months, with value rows."""
    W, H = 600, 320
    M = {'top': 18, 'right': 24, 'bottom': 78, 'left': 48}
    IW = W - M['left'] - M['right']
    IH = H - M['top'] - M['bottom']
    all_vals = []
    for m in monthly:
        for k in ('actual', 'target'):
            n = pn(m[k])
            if n is not None: all_vals.append(n)
    if not all_vals: return ''
    y_min = max(0, (min(all_vals) // 0.05) * 0.05 - 0.05)
    y_max = (max(all_vals) // 0.05) * 0.05 + 0.05
    if y_max - y_min < 0.20: y_max = y_min + 0.25
    n_months = max(12, len(monthly))
    def xf(i): return M['left'] + IW * i / (n_months - 1)
    def yf(v): return M['top'] + IH * (1 - (v - y_min) / (y_max - y_min))
    s = [f'<svg viewBox="0 0 {W} {H}" class="exec-line-chart" preserveAspectRatio="xMidYMid meet">']
    # Y gridlines
    step = 0.05
    v = y_min
    while v <= y_max + 0.001:
        y = yf(v)
        s.append(f'<line x1="{M["left"]}" y1="{y:.1f}" x2="{M["left"]+IW}" y2="{y:.1f}" stroke="#eef2f7" stroke-width="1" />')
        s.append(f'<text x="{M["left"]-8}" y="{y:.1f}" font-size="10" fill="#94a3b8" text-anchor="end" dominant-baseline="middle">{int(round(v*100))}%</text>')
        v += step
    # Target line (gray, all months)
    tpts = []
    for i, m in enumerate(monthly):
        n = pn(m['target'])
        if n is not None: tpts.append((xf(i), yf(n)))
    if len(tpts) >= 2:
        d = 'M ' + ' L '.join(f'{x:.1f} {y:.1f}' for x, y in tpts)
        s.append(f'<path d="{d}" stroke="#9ca3af" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" />')
    # Actual line (black, only where actual exists)
    apts = []
    for i, m in enumerate(monthly):
        n = pn(m['actual'])
        if n is not None: apts.append((xf(i), yf(n)))
    if len(apts) >= 2:
        d = 'M ' + ' L '.join(f'{x:.1f} {y:.1f}' for x, y in apts)
        s.append(f'<path d="{d}" stroke="#0f172a" stroke-width="2.5" fill="none" stroke-linecap="round" stroke-linejoin="round" />')
    # Month labels
    label_y = M['top'] + IH + 16
    for i, m in enumerate(monthly[:12]):
        x = xf(i)
        s.append(f'<text x="{x:.1f}" y="{label_y}" font-size="10" fill="#64748b" text-anchor="middle">{esc(m["month"])}</text>')
    # Value rows: Actual then Target
    row_a = label_y + 22
    row_t = row_a + 16
    s.append(f'<rect x="0" y="{row_a-12}" width="{M["left"]-6}" height="14" fill="#0f172a"/>')
    s.append(f'<text x="{M["left"]-10}" y="{row_a-1}" font-size="9" fill="#ffffff" text-anchor="end" font-weight="700" letter-spacing="0.5">ACTUAL</text>')
    s.append(f'<text x="{M["left"]-10}" y="{row_t-1}" font-size="9" fill="#94a3b8" text-anchor="end" font-weight="700" letter-spacing="0.5">TARGET</text>')
    for i, m in enumerate(monthly[:12]):
        x = xf(i)
        a = pn(m['actual']); t = pn(m['target'])
        if a is not None:
            s.append(f'<text x="{x:.1f}" y="{row_a-1}" font-size="10" fill="#0f172a" font-weight="600" text-anchor="middle" font-variant-numeric="tabular-nums">{int(round(a*100))}%</text>')
        if t is not None:
            s.append(f'<text x="{x:.1f}" y="{row_t-1}" font-size="10" fill="#94a3b8" text-anchor="middle" font-variant-numeric="tabular-nums">{int(round(t*100))}%</text>')
    s.append('</svg>')
    return ''.join(s)


def _exec_bar_chart(disciplines):
    """SVG horizontal bar chart: dark bars with target tick marks."""
    n = len(disciplines)
    BH = 44
    LBL = 110
    PAD_R = 30
    W = 520
    CHART_W = W - LBL - PAD_R
    AX_H = 26
    H = n * BH + AX_H + 8
    scale_max = 1.1
    s = [f'<svg viewBox="0 0 {W} {H}" class="exec-bar-chart" preserveAspectRatio="xMidYMid meet">']
    for i, d in enumerate(disciplines):
        y_top = i * BH
        y_mid = y_top + BH/2
        a = pn(d['actual']); t = pn(d['target'])
        s.append(f'<text x="{LBL-12}" y="{y_mid}" font-size="13" fill="#0f172a" text-anchor="end" dominant-baseline="middle">{esc(d["name"])}</text>')
        bar_y = y_top + 8
        bar_h = BH - 16
        # Background bar
        s.append(f'<rect x="{LBL}" y="{bar_y}" width="{CHART_W}" height="{bar_h}" rx="{bar_h/2:.1f}" fill="#f1f5f9" />')
        # Actual bar
        if a is not None:
            bw = CHART_W * min(scale_max, a) / scale_max
            s.append(f'<rect x="{LBL}" y="{bar_y}" width="{bw:.1f}" height="{bar_h}" rx="{bar_h/2:.1f}" fill="#0f172a" />')
        # Target tick
        if t is not None:
            tx = LBL + CHART_W * (t / scale_max)
            s.append(f'<line x1="{tx:.1f}" y1="{y_top+4}" x2="{tx:.1f}" y2="{y_top+BH-4}" stroke="#0f172a" stroke-width="2" />')
    # X axis ticks
    base_y = n * BH + 4
    for pct in [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0, 1.1]:
        x = LBL + CHART_W * (pct / scale_max)
        s.append(f'<text x="{x:.1f}" y="{base_y+14}" font-size="10" fill="#94a3b8" text-anchor="middle">{int(pct*100)}%</text>')
    s.append('</svg>')
    return ''.join(s)


def render_exec_summary(data):
    """Hero-style exec summary with breadcrumb, big title, 6 KPI cards, line+bar charts."""
    rows = data['rows']
    # Extract subtitle (3rd row first cell, e.g. 'Climbing five straight months. Still 6.4 points behind the 80.2% target.')
    subtitle = ''
    for r in rows[1:5]:
        for v in r:
            if isinstance(v, str) and v.strip() and 'EXECUTIVE' not in v.upper() and 'E X E C U T I V E' not in v:
                subtitle = v.strip(); break
        if subtitle: break

    # Bright spots line
    bright = ''
    for r in rows[25:50]:
        for v in r:
            if isinstance(v, str) and 'bright spot' in v.lower():
                bright = v.strip(); break
        if bright: break

    # Find report "as of" date
    as_of = None
    for r in rows[50:65]:
        for v in r:
            if is_date(v):
                d = to_date(v)
                if d and d.year >= 2025 and d.year <= 2027:
                    as_of = d; break
        if as_of: break

    # Monthly trend
    monthly = []
    trend_start = find_row_with(rows, 'Actual', start=50)
    if trend_start != -1:
        for i in range(trend_start + 1, min(trend_start + 14, len(rows))):
            r = rows[i]
            if not r or not r[0]: continue
            if isinstance(r[0], str) and len(r[0]) <= 4 and r[0].strip():
                monthly.append({
                    'month': r[0],
                    'actual': r[1] if len(r) > 1 else None,
                    'target': r[2] if len(r) > 2 else None,
                    'prior': r[3] if len(r) > 3 else None,
                })

    # Disciplines
    disc_start = find_row_with(rows, 'Dicipline', start=60)
    if disc_start == -1: disc_start = find_row_with(rows, 'Discipline', start=60)
    disciplines = []
    if disc_start != -1:
        for i in range(disc_start + 1, len(rows)):
            r = rows[i]
            if not r or not r[0] or not isinstance(r[0], str): continue
            disciplines.append({'name': r[0], 'actual': r[1] if len(r) > 1 else None, 'target': r[3] if len(r) > 3 else None})

    # KPIs
    ytd_actual = None; ytd_target = None; latest_idx = -1
    for i, m in enumerate(monthly):
        if pn(m['actual']) is not None:
            ytd_actual = pn(m['actual']); ytd_target = pn(m['target']); latest_idx = i
    prior_actual = pn(monthly[latest_idx]['prior']) if latest_idx >= 0 else None
    yoy_diff = (ytd_actual - prior_actual) if (ytd_actual is not None and prior_actual is not None) else None
    vs_target = (ytd_actual - ytd_target) if (ytd_actual is not None and ytd_target is not None) else None
    # FY26 target = average target across all 12 months (weighted plan)
    fy26_targets = [pn(m['target']) for m in monthly if pn(m['target']) is not None]
    fy26_target = sum(fy26_targets)/len(fy26_targets) if fy26_targets else ytd_target
    # Below target count
    below_n, total_offices = _exec_compute_below_target()

    as_of_str = as_of.strftime('%B %-d, %Y').upper() if as_of else 'FY26'
    breadcrumb = f'EXECUTIVE SUMMARY · FP&amp;A · {as_of_str}'

    p = []
    p.append('<div class="exec-page">')
    # Breadcrumb + display title + subtitle
    p.append(f'<div class="exec-eyebrow">{breadcrumb}</div>')
    p.append('<h1 class="exec-display">Company performance.</h1>')
    if subtitle:
        p.append(f'<p class="exec-subtitle">{esc(subtitle)}</p>')

    # KPI Grid (6 cards)
    p.append('<div class="exec-kpi-grid">')
    # 1. YTD Util
    if ytd_actual is not None:
        p.append('<div class="ek"><div class="ek-eyebrow">YTD UTIL.</div>')
        p.append(f'<div class="ek-num">{ytd_actual*100:.1f}<span class="ek-unit">%</span></div>')
        p.append('<div class="ek-foot">Company-wide</div></div>')
    # 2. FY26 Target
    if fy26_target is not None:
        p.append('<div class="ek"><div class="ek-eyebrow">FY26 TARGET</div>')
        p.append(f'<div class="ek-num">{fy26_target*100:.1f}<span class="ek-unit">%</span></div>')
        p.append('<div class="ek-foot">Weighted plan</div></div>')
    # 3. Variance
    if vs_target is not None:
        sign = '−' if vs_target < 0 else '+'
        status_cls = 'bad' if vs_target < 0 else 'good'
        status_lbl = 'Behind' if vs_target < 0 else 'Ahead'
        p.append('<div class="ek"><div class="ek-eyebrow">VARIANCE</div>')
        p.append(f'<div class="ek-num">{sign}{abs(vs_target)*100:.1f}<span class="ek-unit">pts</span></div>')
        p.append('<div class="ek-foot">YTD vs FY26 plan</div>')
        p.append(f'<div class="ek-pill"><span class="pill {status_cls}">{status_lbl}</span></div></div>')
    # 4. Prior Year
    if prior_actual is not None:
        p.append('<div class="ek"><div class="ek-eyebrow">PRIOR YEAR</div>')
        p.append(f'<div class="ek-num">{prior_actual*100:.1f}<span class="ek-unit">%</span></div>')
        p.append('<div class="ek-foot">FY25 same period</div></div>')
    # 5. YOY Change
    if yoy_diff is not None:
        sign = '+' if yoy_diff >= 0 else '−'
        status_cls = 'good' if yoy_diff >= 0 else 'bad'
        status_lbl = 'Up' if yoy_diff >= 0 else 'Down'
        p.append('<div class="ek"><div class="ek-eyebrow">YOY CHANGE</div>')
        p.append(f'<div class="ek-num">{sign}{abs(yoy_diff)*100:.1f}<span class="ek-unit">pts</span></div>')
        p.append('<div class="ek-foot">vs prior year</div>')
        p.append(f'<div class="ek-pill"><span class="pill {status_cls}">{status_lbl}</span></div></div>')
    # 6. Below target
    if below_n is not None and total_offices:
        p.append('<div class="ek"><div class="ek-eyebrow">BELOW TGT</div>')
        p.append(f'<div class="ek-num">{below_n}<span class="ek-unit">/{total_offices}</span></div>')
        p.append(f'<div class="ek-foot">Of {total_offices} active offices</div></div>')
    p.append('</div>')

    # Two charts side by side
    p.append('<div class="exec-charts">')
    if monthly:
        p.append('<div class="exec-chart-block">')
        p.append('<div class="exec-section-eyebrow">MONTHLY TREND · YTD</div>')
        p.append('<div class="exec-section-sub">Actual vs target, FY26</div>')
        p.append(_exec_line_chart(monthly))
        p.append('</div>')
    if disciplines:
        p.append('<div class="exec-chart-block">')
        p.append('<div class="exec-section-eyebrow">THE BIG FOUR</div>')
        p.append('<div class="exec-section-sub">Utilization vs target, YTD</div>')
        p.append(_exec_bar_chart(disciplines))
        p.append('</div>')
    p.append('</div>')

    if bright:
        p.append(f'<div class="exec-bright">{esc(bright)}</div>')

    p.append('</div>')
    return ''.join(p)


def render_target_margin(data):
    """Two side-by-side cards: office target margin/util, and 2025 monthly utilization."""
    rows = data['rows']
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Office Targets &amp; 2025 Performance</h2>')
    p.append('<p class="page-sub">Per-office target margin and target utilization, with historical 2024 actuals and monthly 2025 utilization.</p>')
    p.append('</div>')
    p.append('<div class="two-col">')
    # Left card: Targets
    p.append('<div class="card">')
    p.append('<div class="card-head"><h3>Office Targets</h3><span class="card-sub">Target margin & utilization, with 2024 actual</span></div>')
    p.append('<div class="card-body"><table class="data-table"><thead><tr>')
    p.append('<th>Office</th><th class="num">Target Margin</th><th class="num">Target Utilization</th><th class="num">2024 Actual</th><th class="num">Variance</th>')
    p.append('</tr></thead><tbody>')
    for i in range(2, len(rows)):
        r = rows[i]
        if not r: continue
        office = r[0] if len(r) > 0 else ''
        if not office or office in ('·',) or not isinstance(office, str): continue
        if office.startswith('*'): continue
        tm = r[1] if len(r) > 1 else None
        tu = r[2] if len(r) > 2 else None
        act = r[4] if len(r) > 4 else None
        var = r[6] if len(r) > 6 else None
        tm_d = fmt_pct(tm, 0) if pn(tm) is not None else '—'
        tu_d = fmt_pct(tu, 0) if pn(tu) is not None else '—'
        act_d = fmt_pct(act, 1) if pn(act) is not None else '—'
        var_n = pn(var)
        if var_n is None:
            var_d = '—'
        else:
            sgn = '+' if var_n >= 0 else ''
            cls = 'pos' if var_n >= 0 else 'neg'
            var_d = f'<span class="delta {cls}">{sgn}{var_n*100:.1f}pts</span>'
        p.append(f'<tr><td>{esc(office)}</td><td class="num">{tm_d}</td><td class="num">{tu_d}</td><td class="num">{act_d}</td><td class="num">{var_d}</td></tr>')
    p.append('</tbody></table></div></div>')
    # Right card: 2025 monthly utilization
    header_row = rows[1] if len(rows) > 1 else []
    date_cols = []
    for c in range(14, len(header_row)):
        if is_date(header_row[c]):
            date_cols.append((c, header_row[c]))
    p.append('<div class="card">')
    p.append('<div class="card-head"><h3>2025 Monthly Utilization</h3><span class="card-sub">Office utilization rate by month</span></div>')
    p.append('<div class="scroll-card-inner"><table class="data-table"><thead><tr>')
    p.append('<th class="sticky-l">Office</th>')
    p.append('<th class="num">2025 Total</th>')
    for c, dv in date_cols:
        p.append(f'<th class="num">{fmt_month(dv)}</th>')
    p.append('</tr></thead><tbody>')
    for i in range(2, len(rows)):
        r = rows[i]
        if not r or len(r) < 11: continue
        office = r[10] if len(r) > 10 else ''
        if not office or office in ('·',) or not isinstance(office, str): continue
        total = r[13] if len(r) > 13 else None
        if pn(total) is None: continue
        total_d = fmt_pct(total, 1)
        is_total_row = office == 'Total'
        cls = ' class="total-row"' if is_total_row else ''
        p.append(f'<tr{cls}><td class="sticky-l">{esc(office)}</td><td class="num">{total_d}</td>')
        for c, _ in date_cols:
            v = r[c] if c < len(r) else None
            p.append(f'<td class="num">{fmt_pct(v, 1) if pn(v) is not None else ""}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div></div>')
    p.append('</div>')
    return ''.join(p)
def render_summary_monthly(data):
    """Office × month employee count and utilization."""
    rows = data['rows']
    if len(rows) < 6: return '<p class="empty">No data.</p>'
    header_row = rows[4]
    date_cols = [(c, header_row[c]) for c in range(len(header_row)) if is_date(header_row[c])]
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Monthly Summary by Office</h2>')
    p.append('<p class="page-sub">Headcount per office by month.</p>')
    p.append('</div>')
    p.append('<div class="card">')
    p.append('<div class="card-head"><h3>Headcount by Office</h3><span class="card-sub">Active employees by month</span></div>')
    p.append('<div class="scroll-card-inner"><table class="data-table"><thead><tr>')
    p.append('<th class="sticky-l">Office</th>')
    for c, dv in date_cols:
        p.append(f'<th class="num">{fmt_month(dv)}</th>')
    p.append('</tr></thead><tbody>')
    for i in range(5, len(rows)):
        r = rows[i]
        if not r: continue
        office = r[1] if len(r) > 1 else ''
        if not office or not isinstance(office, str): continue
        is_total = 'total' in office.lower() or 'company' in office.lower()
        cls = ' class="total-row"' if is_total else ''
        p.append(f'<tr{cls}><td class="sticky-l">{esc(office)}</td>')
        for c, _ in date_cols:
            v = r[c] if c < len(r) else None
            n = pn(v)
            disp = fmt_int(v) if n is not None and n == int(n) and abs(n) < 1000 else (fmt_pct(v, 1) if n is not None else '')
            p.append(f'<td class="num">{disp}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div></div>')
    return ''.join(p)


def render_pmo_period(data):
    """PMO Monthly/Weekly: Office × period × (Target, Actual, Δ) with pill on Actual."""
    rows = data['rows']
    if len(rows) < 4: return '<p class="empty">No data.</p>'
    period_row = rows[1]
    period_blocks = []
    c = 1
    while c < len(period_row):
        v = period_row[c]
        if v not in (None, '') and isinstance(v, (str, dict)):
            period_blocks.append((c, v))
            c += 3
        else:
            c += 1
    is_weekly = 'Weekly' in (data.get('tab') or '') or len(period_blocks) > 15
    p = []
    p.append('<div class="page-intro">')
    p.append(f'<h2>{"Weekly" if is_weekly else "Monthly"} PMO Performance</h2>')
    p.append('<p class="page-sub">Office-level target vs actual utilization. Actuals are pilled green when the office met or beat target for the period.</p>')
    p.append('</div>')
    p.append('<div class="toolbar"><span class="legend"><span class="pill good">≥ target</span><span class="pill bad">below target</span></span></div>')
    p.append('<div class="scroll-card"><table class="pmo-grid"><thead>')
    p.append('<tr class="hdr-period">')
    p.append('<th class="col-office" rowspan="2">Office</th>')
    for col_idx, val in period_blocks:
        if is_date(val):
            lbl = fmt_date(val, short=True) if is_weekly else fmt_month(val)
        else:
            lbl = esc(str(val))
        p.append(f'<th colspan="3" class="period-sep">{lbl}</th>')
    p.append('</tr>')
    p.append('<tr class="hdr-sub">')
    for _ in period_blocks:
        p.append('<th class="period-sep">Tgt</th><th>Act</th><th>Δ</th>')
    p.append('</tr></thead><tbody>')
    for i in range(3, len(rows)):
        r = rows[i]
        if not r: continue
        office = r[0] if len(r) > 0 else ''
        if not office or not isinstance(office, str): continue
        is_total = 'total' in office.lower() or 'company' in office.lower() or office.lower().strip() in ('grand total',)
        cls = ' class="total-row"' if is_total else ''
        p.append(f'<tr{cls}><td class="col-office">{esc(office)}</td>')
        for col_idx, _ in period_blocks:
            t_val = r[col_idx] if col_idx < len(r) else None
            a_val = r[col_idx+1] if col_idx+1 < len(r) else None
            d_val = r[col_idx+2] if col_idx+2 < len(r) else None
            t_disp = fmt_pct(t_val, 0) if pn(t_val) is not None else '<span class="muted">—</span>'
            a_pill = pill_cmp(a_val, t_val)
            d_n = pn(d_val)
            if d_n is None:
                d_disp = '<span class="muted">—</span>'
            else:
                d_disp = f'<span class="delta {"pos" if d_n >= 0 else "neg"}">{"+" if d_n>=0 else ""}{d_n*100:.0f}</span>'
            a_cell = a_pill if a_pill else '<span class="muted">—</span>'
            p.append(f'<td class="num period-sep">{t_disp}</td><td class="num">{a_cell}</td><td class="num">{d_disp}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def render_office_hours(data):
    """Summary - OT / Non-Posting: Office × month × hours with heatmap intensity."""
    rows = data['rows']
    header_row_idx = -1
    for i, r in enumerate(rows[:10]):
        date_count = sum(1 for v in r if is_date(v))
        if date_count >= 3:
            header_row_idx = i; break
    if header_row_idx == -1: return render_premium_generic(data)
    header_row = rows[header_row_idx]
    date_cols = [(c, v) for c, v in enumerate(header_row) if is_date(v)]
    all_vals = []
    for i in range(header_row_idx+1, len(rows)):
        for c, _ in date_cols:
            n = pn(rows[i][c] if c < len(rows[i]) else None)
            if n is not None and n > 0: all_vals.append(n)
    max_val = max(all_vals) if all_vals else 1
    tab_name = data.get('tab', '')
    intro_title = 'Overtime' if 'OT' in tab_name else ('Non-Posting Overhead' if 'Non-Posting' in tab_name else 'Hours')
    p = []
    p.append('<div class="page-intro">')
    p.append(f'<h2>{intro_title} Hours by Office</h2>')
    p.append('<p class="page-sub">Monthly total hours per office. Darker cells indicate higher concentrations.</p>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="heat-grid"><thead><tr>')
    p.append('<th class="col-office sticky-l">Office</th>')
    for c, dv in date_cols:
        p.append(f'<th class="num">{fmt_month(dv)}</th>')
    p.append('<th class="num">Total</th>')
    p.append('</tr></thead><tbody>')
    for i in range(header_row_idx+1, len(rows)):
        r = rows[i]
        if not r: continue
        office = r[0] if len(r) > 0 else ''
        if not office or not isinstance(office, str): continue
        row_total = 0; row_has_data = False
        cells = []
        for c, _ in date_cols:
            n = pn(r[c] if c < len(r) else None)
            if n is None: cells.append(('', None)); continue
            row_has_data = True; row_total += n
            cells.append((fmt_num(n, 1), n))
        is_total_row = 'total' in office.lower()
        cls = ' class="total-row"' if is_total_row else ''
        p.append(f'<tr{cls}><td class="col-office sticky-l">{esc(office)}</td>')
        for txt, n in cells:
            if n is None or n == 0:
                p.append('<td class="num zero">·</td>')
            else:
                intensity = min(1.0, n / max_val) if max_val > 0 else 0
                p.append(f'<td class="num heat" style="--intensity:{intensity:.3f}">{txt}</td>')
        p.append(f'<td class="num total-cell">{fmt_num(row_total, 1) if row_has_data else ""}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def render_employee_pill(data):
    """Goal Planning / OVHD Employee / PTO: employee tables, pill where target/actual exist."""
    rows = data['rows']
    header_idx = -1
    for i, r in enumerate(rows[:8]):
        if not r: continue
        text = ' '.join(str(v).lower() for v in r if v)
        if any(k in text for k in ['employee name', 'office', 'pmo']) and any(k in text for k in ['target', 'ytd', 'utiliz']):
            header_idx = i; break
    if header_idx == -1:
        return render_premium_generic(data)
    headers = rows[header_idx]
    cols = [{'idx': c, 'name': str(h).replace('\n',' ').strip()} for c, h in enumerate(headers) if h not in (None, '')]
    target_col = None; ytd_col = None
    for c in cols:
        nl = c['name'].lower()
        if target_col is None and (nl == 'target %' or nl == 'target' or ('target' in nl and 'variance' not in nl and ('%' in c['name'] or nl == 'target'))):
            target_col = c['idx']
        if ytd_col is None and (('ytd' in nl and 'actual' in nl) or 'actual %' in nl or 'utilization (ytd' in nl or 'ytd actual' in nl):
            ytd_col = c['idx']
    p = []
    p.append('<div class="page-intro">')
    p.append(f'<h2>{esc(data.get("tab", ""))}</h2>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="emp-pill-grid"><thead><tr>')
    for c in cols:
        p.append(f'<th>{esc(c["name"])}</th>')
    p.append('</tr></thead><tbody>')
    for i in range(header_idx+1, len(rows)):
        r = rows[i]
        if not r: continue
        if not any(v not in (None,'') for v in r): continue
        first_val = r[0] if r else None
        is_section = (sum(1 for v in r if v not in (None, '')) <= 2) and isinstance(first_val, str) and len(first_val) > 5 and not pn(first_val) and not is_date(first_val)
        if is_section:
            p.append(f'<tr class="section-row"><td colspan="{len(cols)}">{esc(first_val)}</td></tr>')
            continue
        p.append('<tr>')
        for c in cols:
            v = r[c['idx']] if c['idx'] < len(r) else None
            nl = c['name'].lower()
            cls_cell = ''
            if c['idx'] == ytd_col and target_col is not None:
                target_val = r[target_col] if target_col < len(r) else None
                disp = pill_cmp(v, target_val)
                cls_cell = ' class="num"'
            elif 'target' in nl and ('%' in c['name'] or nl in ('target','target %')) and 'variance' not in nl:
                disp = fmt_pct(v, 0) if pn(v) is not None else (esc(str(v)) if v not in (None,'') else '')
                cls_cell = ' class="num"'
            elif 'variance' in nl:
                n = pn(v)
                if n is None: disp = esc(str(v)) if v not in (None,'') else ''
                else: disp = f'<span class="delta {"pos" if n>=0 else "neg"}">{"+" if n>=0 else ""}{n*100:.0f}pts</span>'
                cls_cell = ' class="num"'
            elif is_date(v):
                disp = fmt_date(v); cls_cell = ' class="date"'
            elif pn(v) is not None and 0 < abs(pn(v)) < 2 and pn(v) != int(pn(v)) and 'hour' not in nl and 'pto' not in nl:
                disp = fmt_pct(v, 1); cls_cell = ' class="num"'
            elif pn(v) is not None:
                disp = fmt_num(v, 1); cls_cell = ' class="num"'
            else:
                disp = esc(str(v)) if v not in (None, '') else ''
            p.append(f'<td{cls_cell}>{disp}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def render_pivot_hours(data):
    """Pivot - Proposal / Fleet: rows × month × hours with heatmap."""
    rows = data['rows']
    header_idx = -1
    for i, r in enumerate(rows[:8]):
        if not r: continue
        date_count = sum(1 for v in r if is_date(v))
        if date_count >= 3:
            header_idx = i; break
    if header_idx == -1:
        return render_premium_generic(data)
    headers = rows[header_idx]
    label_cols = []; date_cols = []; total_col = None
    for c, h in enumerate(headers):
        if is_date(h):
            date_cols.append((c, h))
        elif h not in (None, ''):
            h_str = str(h).replace('\n',' ').strip()
            if 'total' in h_str.lower():
                total_col = (c, h_str)
            else:
                label_cols.append((c, h_str))
    p = []
    p.append('<div class="page-intro">')
    p.append(f'<h2>{esc(data.get("tab", ""))}</h2>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="data-table"><thead><tr>')
    for c, name in label_cols:
        p.append(f'<th class="sticky-l">{esc(name)}</th>')
    for c, dv in date_cols:
        p.append(f'<th class="num">{fmt_month(dv)}</th>')
    if total_col:
        p.append(f'<th class="num">Total</th>')
    p.append('</tr></thead><tbody>')
    all_vals = []
    for i in range(header_idx+1, len(rows)):
        for c, _ in date_cols:
            n = pn(rows[i][c] if c < len(rows[i]) else None)
            if n is not None and n > 0: all_vals.append(n)
    max_val = max(all_vals) if all_vals else 1
    for i in range(header_idx+1, len(rows)):
        r = rows[i]
        if not r: continue
        if not any(v not in (None,'') for v in r): continue
        non_empty_count = sum(1 for v in r if v not in (None, ''))
        first_val = r[0] if r else None
        is_section = non_empty_count == 1 and isinstance(first_val, str) and len(first_val) > 3
        total_cols = len(label_cols) + len(date_cols) + (1 if total_col else 0)
        if is_section:
            p.append(f'<tr class="section-row"><td colspan="{total_cols}">{esc(first_val)}</td></tr>')
            continue
        p.append('<tr>')
        for c, _ in label_cols:
            v = r[c] if c < len(r) else None
            p.append(f'<td class="sticky-l">{esc(str(v)) if v not in (None,"") else ""}</td>')
        for c, _ in date_cols:
            n = pn(r[c] if c < len(r) else None)
            if n is None or n == 0:
                p.append('<td class="num zero">·</td>')
            else:
                intensity = min(1.0, n / max_val) if max_val > 0 else 0
                p.append(f'<td class="num heat" style="--intensity:{intensity:.3f}">{fmt_num(n, 1)}</td>')
        if total_col:
            v = r[total_col[0]] if total_col[0] < len(r) else None
            p.append(f'<td class="num total-cell">{fmt_num(v, 1)}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def render_pivot_tree(data):
    """Pivot - Project by Task: hierarchical tree."""
    rows = data['rows']
    header_idx = -1
    for i, r in enumerate(rows[:5]):
        if r and any('Row Labels' in str(v) if v else False for v in r):
            header_idx = i; break
    if header_idx == -1: header_idx = 3
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Project Hours by Task</h2>')
    p.append('<p class="page-sub">Hierarchical breakdown of utilized and non-utilized hours by project and task.</p>')
    p.append('</div>')
    p.append('<div class="toolbar"><input type="text" id="treeSearch" class="input" placeholder="Search project or task…" /></div>')
    p.append('<div class="scroll-card"><table class="data-table tree-table"><thead><tr>')
    p.append('<th class="sticky-l">Project / Task</th><th class="num">Utilized hrs</th><th class="num">Non-utilized hrs</th><th class="num">Total hrs</th>')
    p.append('</tr></thead><tbody>')
    for i in range(header_idx+1, len(rows)):
        r = rows[i]
        if not r: continue
        label = r[0] if r else ''
        if label in (None, '') and not any(v not in (None,'') for v in r[1:] if len(r)>1): continue
        util = r[1] if len(r) > 1 else None
        nonutil = r[2] if len(r) > 2 else None
        tot = r[3] if len(r) > 3 else None
        label_str = str(label) if label else ''
        is_top = bool(label_str) and len(label_str) >= 2 and label_str[:2].isdigit() and '-' in label_str
        is_total = 'total' in label_str.lower() and 'grand' in label_str.lower()
        cls = 'project-row' if is_top else ('total-row' if is_total else 'task-row')
        p.append(f'<tr class="{cls}" data-label="{esc(label_str.lower())}">')
        indent = '' if is_top else '<span class="indent">↳</span> '
        p.append(f'<td class="sticky-l">{indent}{esc(label_str)}</td>')
        p.append(f'<td class="num">{fmt_num(util, 2)}</td>')
        p.append(f'<td class="num">{fmt_num(nonutil, 2)}</td>')
        p.append(f'<td class="num total-cell">{fmt_num(tot, 2)}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('<script>(function(){var s=document.getElementById("treeSearch");if(!s) return;var rows=document.querySelectorAll(".tree-table tbody tr");s.addEventListener("input",function(){var q=s.value.toLowerCase().trim();rows.forEach(function(r){r.classList.toggle("hidden",q && !(r.dataset.label||"").includes(q));});});})();</script>')
    return ''.join(p)


def render_employee_roster(data):
    """PT Hours: PT staff roster with search."""
    rows = data['rows']
    if not rows: return '<p class="empty">No data.</p>'
    headers = rows[0]
    label_cols = [(c, str(h).replace('\n',' ').strip()) for c, h in enumerate(headers) if h not in (None, '')]
    target_col = None
    for c, h in label_cols:
        if h.lower() == 'target': target_col = c; break
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Part-Time Employee Roster</h2>')
    p.append(f'<p class="page-sub">{len(rows)-1} part-time staff with targets and recent timesheet activity.</p>')
    p.append('</div>')
    p.append('<div class="toolbar"><input type="text" id="rosterSearch" class="input" placeholder="Search name, ID, office, position…" /></div>')
    p.append('<div class="scroll-card"><table class="data-table" id="rosterTable"><thead><tr>')
    for c, h in label_cols:
        p.append(f'<th>{esc(h)}</th>')
    p.append('</tr></thead><tbody>')
    for i in range(1, len(rows)):
        r = rows[i]
        if not r or not any(v not in (None,'') for v in r): continue
        rtext = ' '.join(str(v) for v in r if v not in (None,'')).lower()
        p.append(f'<tr data-search="{esc(rtext)}">')
        for c, h in label_cols:
            v = r[c] if c < len(r) else None
            nl = h.lower()
            cls_cell = ''
            if is_date(v):
                disp = fmt_date(v); cls_cell = ' class="date"'
            elif c == target_col:
                disp = fmt_pct(v, 0) if pn(v) is not None else (esc(str(v)) if v not in (None,'') else '')
                cls_cell = ' class="num"'
            elif pn(v) is not None and 0 < abs(pn(v)) < 2 and pn(v) != int(pn(v)):
                disp = fmt_pct(v, 1); cls_cell = ' class="num"'
            elif pn(v) is not None:
                disp = fmt_num(v, 1); cls_cell = ' class="num"'
            else:
                disp = esc(str(v)) if v not in (None, '') else ''
            p.append(f'<td{cls_cell}>{disp}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('<script>(function(){var s=document.getElementById("rosterSearch");if(!s) return;var rows=document.querySelectorAll("#rosterTable tbody tr");s.addEventListener("input",function(){var q=s.value.toLowerCase().trim();rows.forEach(function(r){r.classList.toggle("hidden",q && !(r.dataset.search||"").includes(q));});});})();</script>')
    return ''.join(p)


def render_ovhd_hours(data):
    """Summary - OVHD (Hours): grouped headers (Present Week / Month / YTD)."""
    rows = data['rows']
    if len(rows) < 5: return render_premium_generic(data)
    group_row = rows[2]
    head_row = rows[3]
    cols = []
    for c, h in enumerate(head_row):
        if h in (None, ''): continue
        cols.append({'idx': c, 'name': str(h).replace('\n', ' ').strip(), 'group': ''})
    group_label = ''
    for c, gv in enumerate(group_row):
        if gv not in (None, '') and isinstance(gv, str):
            group_label = gv
        for col in cols:
            if col['idx'] == c:
                col['group'] = group_label
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Overhead Hours by Office</h2>')
    p.append('<p class="page-sub">Hour allocation buckets: Billable vs Unbillable utilization, plus Admin / PTO / S&amp;M categories — across Present Week, Present Month, and YTD.</p>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="data-table"><thead>')
    p.append('<tr class="hdr-group">')
    i = 0
    while i < len(cols):
        g = cols[i]['group']; j = i
        while j < len(cols) and cols[j]['group'] == g:
            j += 1
        p.append(f'<th colspan="{j-i}">{esc(g) if g else ""}</th>')
        i = j
    p.append('</tr><tr>')
    for col in cols:
        cls = 'sticky-l' if col['idx'] == 0 else ('' if col['idx'] <= 1 else 'num')
        p.append(f'<th class="{cls}">{esc(col["name"])}</th>')
    p.append('</tr></thead><tbody>')
    for i in range(4, len(rows)):
        r = rows[i]
        if not r: continue
        first = r[0] if r else None
        if not first: continue
        is_total = isinstance(first, str) and 'total' in first.lower()
        cls = ' class="total-row"' if is_total else ''
        p.append(f'<tr{cls}>')
        for col in cols:
            v = r[col['idx']] if col['idx'] < len(r) else None
            nl = col['name'].lower()
            if col['idx'] == 0:
                p.append(f'<td class="sticky-l">{esc(str(v)) if v not in (None,"") else ""}</td>')
            elif col['idx'] == 1:
                p.append(f'<td>{esc(str(v)) if v not in (None,"") else ""}</td>')
            else:
                n = pn(v)
                if n is None: disp = esc(str(v)) if v not in (None, '') else ''
                elif ('billable' in nl or 'unbillable' in nl) and 'ytd' in col['group'].lower() and abs(n) < 2:
                    disp = fmt_pct(n, 1)
                else:
                    disp = fmt_num(n, 1)
                p.append(f'<td class="num">{disp}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def render_premium_generic(data):
    """Polished generic table for anything without a custom renderer."""
    rows = data['rows']
    if not rows: return '<p class="empty">No data.</p>'
    header_idx = 0
    while header_idx < len(rows) and not any(v not in (None, '') for v in rows[header_idx]):
        header_idx += 1
    if header_idx >= len(rows): return '<p class="empty">No data.</p>'
    headers = rows[header_idx]
    max_cols = max(len(r) for r in rows[header_idx:])
    col_modes = []
    for c in range(max_cols):
        vals = [rows[i][c] for i in range(header_idx+1, len(rows)) if c < len(rows[i])]
        nums = [pn(v) for v in vals if pn(v) is not None]
        h_str = str(headers[c]).lower() if c < len(headers) and headers[c] else ''
        if not nums:
            col_modes.append('text')
        elif any(k in h_str for k in ['%', 'util', 'target', 'actual', 'variance', 'margin']):
            col_modes.append('pct')
        elif all(n == int(n) and abs(n) < 1e6 for n in nums) and not any(k in h_str for k in ['hour', 'fee']):
            col_modes.append('int')
        elif all(-2 <= n <= 2 for n in nums):
            col_modes.append('pct')
        else:
            col_modes.append('num')
    p = []
    p.append('<div class="page-intro">')
    p.append(f'<h2>{esc(data.get("tab", ""))}</h2>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="data-table"><thead><tr>')
    for c in range(max_cols):
        h = headers[c] if c < len(headers) else ''
        h_disp = esc(str(h).replace('\n', ' ').strip()) if h not in (None, '') else ''
        p.append(f'<th class="{"num" if col_modes[c] in ("num","int","pct") else ""}">{h_disp}</th>')
    p.append('</tr></thead><tbody>')
    for i in range(header_idx+1, len(rows)):
        r = rows[i]
        if not any(v not in (None, '') for v in r): continue
        non_empty_count = sum(1 for v in r if v not in (None, ''))
        first_val = r[0] if r else None
        is_section = non_empty_count <= 2 and isinstance(first_val, str) and len(first_val.strip()) >= 3 and not pn(first_val)
        if is_section:
            p.append(f'<tr class="section-row"><td colspan="{max_cols}">{esc(first_val)}</td></tr>')
            continue
        p.append('<tr>')
        for c in range(max_cols):
            v = r[c] if c < len(r) else None
            mode = col_modes[c]
            if v in (None, ''):
                p.append('<td></td>'); continue
            if is_date(v):
                p.append(f'<td class="date">{fmt_date(v)}</td>'); continue
            if mode == 'pct':
                p.append(f'<td class="num">{fmt_pct(v, 1) if pn(v) is not None else esc(str(v))}</td>')
            elif mode == 'int':
                p.append(f'<td class="num">{fmt_int(v) if pn(v) is not None else esc(str(v))}</td>')
            elif mode == 'num':
                p.append(f'<td class="num">{fmt_num(v, 2) if pn(v) is not None else esc(str(v))}</td>')
            else:
                p.append(f'<td>{esc(str(v))}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def render_pmo_weekly(data):
    """PMO Weekly: multiple stacked sections (headcount, total util, billable util).
    Row 5 is the date header. Section breaks occur where col 0 is empty."""
    rows = data['rows']
    if len(rows) < 6: return render_premium_generic(data)
    # Find header row (first row with 5+ dates)
    header_idx = -1
    for i, r in enumerate(rows[:8]):
        date_count = sum(1 for v in r if is_date(v))
        if date_count >= 5: header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    header = rows[header_idx]
    date_cols = [(c, v) for c, v in enumerate(header) if is_date(v)]
    # Detect section breaks: rows where col 0 has a label like "TOTAL" or where data resets
    # We'll segment by detecting when col 0 is empty and col 1 has an office (= continuation)
    # vs when row contains a label-only marker
    sections = []  # list of (title, [row indices])
    cur_title = 'Headcount by Office'
    cur_rows = []
    section_titles_seen = 0
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r: continue
        col0 = r[0] if len(r) > 0 else ''
        col1 = r[1] if len(r) > 1 else ''
        # Section break? rows where most cells are empty but there's text
        non_empty = sum(1 for v in r if v not in (None, ''))
        if non_empty <= 2 and isinstance(col0, str) and len(col0) > 3 and not pn(col0):
            # save current
            if cur_rows: sections.append((cur_title, cur_rows))
            cur_title = col0; cur_rows = []; continue
        # Office row: col 1 is name. Some rows have col0 = "CompanyBig Sky", col1 = "Big Sky"
        if col1 and isinstance(col1, str) and not pn(col1):
            cur_rows.append(i)
        elif col0 and isinstance(col0, str) and not pn(col0) and not col0.lower().startswith('company'):
            cur_rows.append(i)
    if cur_rows: sections.append((cur_title, cur_rows))
    # Heuristic: assign meaningful titles based on order
    auto_titles = ['Headcount by Office', 'Total Utilization', 'Billable Utilization', 'Non-Billable Utilization', 'Section 5', 'Section 6', 'Section 7']
    for idx, (t, rs) in enumerate(sections):
        if idx < len(auto_titles) and (t == 'Headcount by Office' or 'company' in t.lower()):
            sections[idx] = (auto_titles[idx], rs)

    # Compute global max for heatmap (per section actually — better)
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Weekly PMO Performance</h2>')
    p.append('<p class="page-sub">Stacked weekly views by office: headcount, total utilization, billable utilization, and more.</p>')
    p.append('</div>')

    for sec_idx, (title, idxs) in enumerate(sections):
        if not idxs: continue
        # Determine if values look like utilization (decimals < 2) or headcount (ints)
        sample = [pn(rows[i][c] if c < len(rows[i]) else None) for i in idxs for c, _ in date_cols[:10]]
        sample = [s for s in sample if s is not None]
        is_pct_section = bool(sample) and sum(1 for s in sample if 0 <= s <= 2 and s != int(s)) > len(sample) * 0.3
        # Build pills if pct, else heatmap
        all_vals = [pn(rows[i][c] if c < len(rows[i]) else None) for i in idxs for c, _ in date_cols]
        all_vals = [v for v in all_vals if v is not None and v > 0]
        max_val = max(all_vals) if all_vals else 1
        p.append('<div class="card">')
        p.append(f'<div class="card-head"><h3>{esc(title)}</h3><span class="card-sub">Office by week · {len(idxs)} rows · {len(date_cols)} weeks</span></div>')
        p.append('<div class="scroll-card-inner"><table class="data-table"><thead><tr>')
        p.append('<th class="sticky-l">Office</th>')
        for c, dv in date_cols:
            p.append(f'<th class="num">{fmt_date(dv, short=True)}</th>')
        p.append('</tr></thead><tbody>')
        for ri in idxs:
            r = rows[ri]
            office_name = r[1] if len(r) > 1 and r[1] else (r[0] if r else '')
            if isinstance(office_name, str) and office_name.lower().startswith('company'):
                office_name = office_name[7:] if len(office_name) > 7 else office_name
            is_total = isinstance(office_name, str) and 'total' in office_name.lower()
            cls = ' class="total-row"' if is_total else ''
            p.append(f'<tr{cls}><td class="sticky-l">{esc(office_name)}</td>')
            for c, _ in date_cols:
                v = r[c] if c < len(r) else None
                n = pn(v)
                if n is None:
                    p.append('<td class="num zero">·</td>')
                elif is_pct_section:
                    # Color by intensity: closer to 1 = darker green
                    intensity = min(1.0, max(0.0, n))
                    p.append(f'<td class="num heat" style="--intensity:{intensity:.3f}">{fmt_pct(n, 0)}</td>')
                elif n == 0:
                    p.append('<td class="num zero">·</td>')
                else:
                    intensity = min(1.0, n / max_val) if max_val > 0 else 0
                    p.append(f'<td class="num heat" style="--intensity:{intensity:.3f}">{fmt_int(n) if n == int(n) else fmt_num(n, 1)}</td>')
            p.append('</tr>')
        p.append('</tbody></table></div></div>')
    return ''.join(p)


RENDERERS = {
    'bubble_dashboard': render_bubble_dashboard,
    'exec_summary': render_exec_summary,
    'target_margin': render_target_margin,
    'summary_monthly': render_summary_monthly,
    'pmo_period': render_pmo_period,
    'pmo_weekly': render_pmo_weekly,
    'office_hours': render_office_hours,
    'employee_pill': render_employee_pill,
    'pivot_hours': render_pivot_hours,
    'pivot_tree': render_pivot_tree,
    'employee_roster': render_employee_roster,
    'ovhd_hours': render_ovhd_hours,
    'premium_generic': render_premium_generic,
}

# =============================================================================
# Updated renderers for the edits round
# =============================================================================

import datetime as _dt

# Report as of the build date (auto-advances each refresh). FPA_AS_OF=YYYY-MM-DD pins it so a
# past weekly package rebuilds exactly as it did that week; without it, behaviour is unchanged.
_FPA_AS_OF = os.environ.get('FPA_AS_OF', '').strip()
TODAY_REF = (_dt.datetime.strptime(_FPA_AS_OF, '%Y-%m-%d') if _FPA_AS_OF
             else _dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0))


def render_bubble_dashboard_v2(_data):
    """Employee Utilization (revised).
    - Replace All Status with Employee Type + Above/Below Target YTD filters
    - Toggle to hide/show weekly cells (collapsed view: Office, Employee, Target, YTD, Quarterly only)
    - YTD and Quarterly Δ shown as small triangle delta indicator
    - Hide weeks that haven't occurred yet
    """
    with open('employees.json') as f:
        edata = json.load(f)
    employees = edata['employees']
    date_headers = edata['date_headers']
    WEEKLY_Q1 = list(range(9, 22)); WEEKLY_Q2 = list(range(26, 39))
    WEEKLY_Q3 = list(range(43, 56)); WEEKLY_Q4 = list(range(60, 73))
    QUARTERLY = [23, 40, 57, 74]
    quarters_def = [('Q1', WEEKLY_Q1, 23), ('Q2', WEEKLY_Q2, 40), ('Q3', WEEKLY_Q3, 57), ('Q4', WEEKLY_Q4, 74)]

    # Compute which weeks are in the past
    week_dt = {}
    for c in WEEKLY_Q1 + WEEKLY_Q2 + WEEKLY_Q3 + WEEKLY_Q4:
        try:
            week_dt[c] = _dt.datetime.fromisoformat(date_headers[str(c)])
        except Exception:
            week_dt[c] = None
    def week_is_past(c):
        d = week_dt.get(c)
        return d is not None and d <= TODAY_REF

    week_labels = {c: fmt_date({'__date__': date_headers[str(c)] + 'T00:00:00'}, short=True) for c in WEEKLY_Q1+WEEKLY_Q2+WEEKLY_Q3+WEEKLY_Q4}
    offices = sorted(set(e['office'] for e in employees if e['office']))

    def delta_html(actual, target):
        v = pn(actual); t = pn(target)
        if v is None or t is None: return '<span class="delta-muted">—</span>'
        d = v - t
        if d > 0.0005: return f'<span class="delta pos">▲ {abs(d)*100:.1f}</span>'
        if d < -0.0005: return f'<span class="delta neg">▼ {abs(d)*100:.1f}</span>'
        return '<span class="delta-muted">●</span>'

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Employee Utilization</h2>')
    p.append(f'<p class="page-sub">Weekly performance for {len(employees)} billable employees across {len(offices)} offices — pills are green when the value meets or beats each employee\'s individual target.</p>')
    p.append('</div>')
    p.append('<div class="toolbar">')
    p.append('<input type="text" id="bubbleSearch" class="input" placeholder="Search employee or ID…" />')
    p.append('<select id="bubbleOffice" class="input"><option value="">All offices</option>')
    for o in offices: p.append(f'<option value="{esc(o)}">{esc(o)}</option>')
    p.append('</select>')
    p.append('<select id="bubbleType" class="input"><option value="">All employees</option><option value="Full Time">Full Time</option><option value="Part Time">Part Time</option></select>')
    p.append('<select id="bubbleYtd" class="input"><option value="">YTD vs target</option><option value="good">At or above target</option><option value="bad">Below target</option></select>')
    p.append('<label class="toggle"><input type="checkbox" id="bubbleCompact" /> <span>Compact view</span></label>')
    p.append('<span class="legend"><span class="pill good">≥ target</span><span class="pill bad">below target</span><span class="pill na">N/A</span></span>')
    p.append('</div>')
    p.append(f'<div class="result-count"><span id="bubbleCount">{len(employees)}</span> of {len(employees)} employees</div>')
    p.append('<div class="scroll-card"><table class="bubble-grid"><thead>')
    p.append('<tr class="row-quarter">')
    p.append('<th class="col-office" rowspan="2">Office</th>')
    p.append('<th class="col-name" rowspan="2">Employee</th>')
    p.append('<th class="col-target" rowspan="2">Target</th>')
    p.append('<th class="ytd-sep" rowspan="2">YTD</th>')
    p.append('<th class="ytd-sep col-delta" rowspan="2">YTD Δ</th>')
    for q_label, weeks, _ in quarters_def:
        span = len(weeks) + 2  # weeks + Q total + Q delta
        p.append(f'<th class="quarter-sep" colspan="{span}">{q_label} 2026</th>')
    p.append('</tr><tr class="row-week">')
    for q_label, weeks, _ in quarters_def:
        for i, c in enumerate(weeks):
            cls = ['week-cell']
            if i == 0: cls.append('quarter-sep')
            if not week_is_past(c): cls.append('week-future')
            p.append(f'<th class="{" ".join(cls)}" data-col="{c}">{week_labels[c]}</th>')
        p.append(f'<th class="q-total">{q_label}</th>')
        p.append(f'<th class="q-delta col-delta">{q_label} Δ</th>')
    p.append('</tr></thead><tbody>')

    current_office = None
    for emp in employees:
        office = emp['office'] or ''
        if office != current_office:
            current_office = office
            total_cols = 5 + sum(len(w)+2 for _, w, _ in quarters_def)
            p.append(f'<tr class="office-group" data-office-group="{esc(office)}"><td colspan="{total_cols}">{esc(office.upper())}</td></tr>')
        name = emp['name'] or ''
        emp_id = emp['id'] or ''
        target = emp['target']
        regtemp = (emp.get('regtemp') or '').strip()
        v_ytd = pn(emp['ytd']); t_v = pn(target)
        ytd_class = 'na' if v_ytd is None or t_v is None else ('good' if v_ytd >= t_v else 'bad')
        ytd_pill = pill_cmp(emp['ytd'], target)
        ytd_delta = delta_html(emp['ytd'], target)
        p.append(f'<tr class="emp-row" data-office="{esc(office)}" data-name="{esc(str(name).lower())}" data-id="{esc(str(emp_id).lower())}" data-type="{esc(regtemp)}" data-ytd="{ytd_class}">')
        p.append(f'<td class="col-office">{esc(office)}</td>')
        p.append(f'<td class="col-name">{esc(name)}<br><span class="emp-id">{esc(str(emp_id))}</span></td>')
        t_disp = fmt_pct(target) if pn(target) is not None else '—'
        p.append(f'<td class="col-target">{t_disp}</td>')
        p.append(f'<td class="cell ytd-sep">{ytd_pill}</td>')
        p.append(f'<td class="cell ytd-sep col-delta">{ytd_delta}</td>')
        for q_label, weeks, q_col in quarters_def:
            for i, c in enumerate(weeks):
                cls = ['cell', 'week-cell']
                if i == 0: cls.append('quarter-sep')
                if not week_is_past(c): cls.append('week-future')
                val = emp['weeks'].get(str(c))
                p.append(f'<td class="{" ".join(cls)}" data-col="{c}">{pill_cmp(val, target)}</td>')
            qval = emp['quarters'].get(str(q_col))
            p.append(f'<td class="cell q-total">{pill_cmp(qval, target)}</td>')
            p.append(f'<td class="cell q-delta col-delta">{delta_html(qval, target)}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('''<script>(function(){
var s=document.getElementById("bubbleSearch"),o=document.getElementById("bubbleOffice"),t=document.getElementById("bubbleType"),y=document.getElementById("bubbleYtd"),c=document.getElementById("bubbleCompact"),cnt=document.getElementById("bubbleCount");
if(!s) return;
var rows=document.querySelectorAll("tr.emp-row"),grps=document.querySelectorAll("tr.office-group"),wks=document.querySelectorAll("th.week-cell,td.week-cell");
function apply(){var q=s.value.trim().toLowerCase(),off=o.value,tp=t.value,yy=y.value,vis=0,seen=new Set();
rows.forEach(function(r){var ok=true;if(q&&!((r.dataset.name||"").includes(q)||(r.dataset.id||"").includes(q)))ok=false;
if(off&&r.dataset.office!==off)ok=false;
if(tp&&r.dataset.type!==tp)ok=false;
if(yy==="good"&&r.dataset.ytd!=="good")ok=false;
if(yy==="bad"&&r.dataset.ytd!=="bad")ok=false;
r.classList.toggle("hidden",!ok);if(ok){vis++;seen.add(r.dataset.office);}});
grps.forEach(function(g){g.classList.toggle("hidden",!seen.has(g.dataset.officeGroup));});
if(cnt) cnt.textContent=vis;}
function compact(){var on=c.checked; wks.forEach(function(w){w.classList.toggle("collapsed",on);});}
s.addEventListener("input",apply);o.addEventListener("change",apply);t.addEventListener("change",apply);y.addEventListener("change",apply);
c.addEventListener("change",compact);
})();</script>''')
    return ''.join(p)


def render_pmo_period_v2(data):
    """PMO Utilization Monthly with narrower PMO column."""
    rows = data['rows']
    if len(rows) < 4: return '<p class="empty">No data.</p>'
    period_row = rows[1]
    period_blocks = []
    c = 1
    while c < len(period_row):
        v = period_row[c]
        if v not in (None, '') and isinstance(v, (str, dict)):
            period_blocks.append((c, v)); c += 3
        else: c += 1
    is_weekly = 'Weekly' in (data.get('tab') or '')
    p = []
    p.append('<div class="page-intro">')
    p.append(f'<h2>{"Weekly" if is_weekly else "Monthly"} PMO Utilization</h2>')
    p.append('<p class="page-sub">Office-level target vs actual utilization. Actuals pill green when the office met or beat target for the period.</p>')
    p.append('</div>')
    p.append('<div class="toolbar"><span class="legend"><span class="pill good">≥ target</span><span class="pill bad">below target</span></span></div>')
    p.append('<div class="scroll-card"><table class="pmo-grid pmo-narrow"><thead>')
    p.append('<tr class="hdr-period">')
    p.append('<th class="col-office" rowspan="2">PMO</th>')
    for col_idx, val in period_blocks:
        if is_date(val):
            lbl = fmt_date(val, short=True) if is_weekly else fmt_month(val)
        else:
            lbl = esc(str(val))
        p.append(f'<th colspan="3" class="period-sep">{lbl}</th>')
    p.append('</tr><tr class="hdr-sub">')
    for _ in period_blocks:
        p.append('<th class="period-sep">Tgt</th><th>Act</th><th>Δ</th>')
    p.append('</tr></thead><tbody>')
    for i in range(3, len(rows)):
        r = rows[i]
        if not r: continue
        office = r[0] if len(r) > 0 else ''
        if not office or not isinstance(office, str): continue
        is_total = 'total' in office.lower() or 'company' in office.lower()
        cls = ' class="total-row"' if is_total else ''
        p.append(f'<tr{cls}><td class="col-office">{esc(office)}</td>')
        for col_idx, _ in period_blocks:
            t_val = r[col_idx] if col_idx < len(r) else None
            a_val = r[col_idx+1] if col_idx+1 < len(r) else None
            d_val = r[col_idx+2] if col_idx+2 < len(r) else None
            t_disp = fmt_pct(t_val, 0) if pn(t_val) is not None else '<span class="muted">—</span>'
            a_pill = pill_cmp(a_val, t_val)
            d_n = pn(d_val)
            if d_n is None:
                d_disp = '<span class="muted">—</span>'
            else:
                d_disp = f'<span class="delta {"pos" if d_n >= 0 else "neg"}">{"+" if d_n>=0 else ""}{d_n*100:.0f}</span>'
            a_cell = a_pill if a_pill else '<span class="muted">—</span>'
            p.append(f'<td class="num period-sep">{t_disp}</td><td class="num">{a_cell}</td><td class="num">{d_disp}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def render_pmo_weekly_v2(data):
    """PMO Utilization Weekly:
    - Most recent 10 weeks only
    - Hide the dormant Archive Services office and Sales & Marketing
    - Apply bubble pills based on per-row target (use Total Utilization rows: target is row's target; for utilization sections compare to row target)
    - Banners for Company / Full Time / Part Time
    - Vertical (long) layout like Employee Utilization
    """
    rows = data['rows']
    if len(rows) < 6: return render_premium_generic(data)
    # Find header row with most dates
    header_idx = -1
    for i, r in enumerate(rows[:10]):
        if sum(1 for v in r if is_date(v)) >= 5:
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    header = rows[header_idx]
    all_date_cols = [(c, v) for c, v in enumerate(header) if is_date(v)]
    # Most recent 10 weeks that are at or before TODAY_REF
    past_dates = [(c, v) for (c, v) in all_date_cols if to_date(v) and to_date(v) <= TODAY_REF]
    if not past_dates:
        date_cols = all_date_cols[-10:]
    else:
        date_cols = past_dates[-10:]

    # Detect sections by scanning rows for section markers (single-cell rows or specific labels)
    # Sections in PMO Weekly: "Headcount and weekly...", then probably nothing else. Section breaks are blank rows.
    sections = []
    cur_title = 'Total Utilization'
    section_labels = ['Total Utilization', 'Billable Utilization', 'Non-Billable Utilization', 'Other']
    cur_rows = []
    section_idx = 0
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r:
            continue
        non_empty = sum(1 for v in r if v not in (None, ''))
        col0 = r[0] if len(r) > 0 else ''
        col1 = r[1] if len(r) > 1 else ''
        # Section divider heuristic: only 1 cell with text and it doesn't look like an office row
        if non_empty <= 2 and isinstance(col0, str) and col0 and not col1 and not pn(col0):
            if cur_rows:
                sections.append((cur_title, cur_rows)); cur_rows = []
                section_idx += 1
                cur_title = section_labels[section_idx] if section_idx < len(section_labels) else col0
            continue
        # Office row check
        if isinstance(col1, str) and col1 and not pn(col1):
            cur_rows.append(i)
        elif isinstance(col0, str) and col0 and not col0.lower().startswith('company'):
            cur_rows.append(i)
    if cur_rows: sections.append((cur_title, cur_rows))

    # Filter out the dormant Archive Services office and Sales & Marketing rows
    hidden = {'archive services', 'sales and marketing', 'sales & marketing', 's & m', 's&m'}
    def office_of(r):
        n = r[1] if len(r) > 1 and r[1] else (r[0] if r else '')
        if isinstance(n, str) and n.lower().startswith('company'):
            n = n[7:]
        return str(n)

    # Per-office targets (try to read from row's Target % if present, else None)
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Weekly PMO Utilization</h2>')
    p.append(f'<p class="page-sub">Most recent {len(date_cols)} weeks of utilization by office, with headcount, total, billable, and non-billable views.</p>')
    p.append('</div>')

    sec_emojis = {'Total Utilization':'Total','Billable Utilization':'Billable','Non-Billable Utilization':'Non-Billable'}
    for sec_idx, (title, idxs) in enumerate(sections):
        clean_idxs = [i for i in idxs if office_of(rows[i]).strip().lower() not in hidden]
        if not clean_idxs: continue
        # Determine if pct (utilization) or int (headcount)
        sample = [pn(rows[i][c]) for i in clean_idxs for c, _ in date_cols]
        sample = [s for s in sample if s is not None]
        is_pct = bool(sample) and sum(1 for s in sample if 0 <= s <= 1.5 and s != int(s)) > len(sample) * 0.3
        section_label = sec_emojis.get(title, title) if title else f'Section {sec_idx+1}'
        if sec_idx == 0 and not is_pct: section_label = 'Headcount by Office'

        p.append('<div class="card">')
        p.append(f'<div class="card-head"><h3>{esc(section_label)}</h3><span class="card-sub">{len(clean_idxs)} rows × {len(date_cols)} weeks</span></div>')
        p.append('<div class="scroll-card-inner"><table class="data-table"><thead><tr>')
        p.append('<th class="sticky-l">Office</th>')
        for c, dv in date_cols:
            p.append(f'<th class="num">{fmt_date(dv, short=True)}</th>')
        p.append('</tr></thead><tbody>')
        # Banners by index in section: Company total row, FT block, PT block
        # Heuristic: first row is Company. Find FT/PT splits if labels exist (e.g., "Company - FT", "Company - PT")
        for j, ri in enumerate(clean_idxs):
            r = rows[ri]
            office = office_of(r)
            # Banner row?
            ol = office.lower()
            is_banner = j == 0 or any(k in ol for k in (' - ft', ' - pt', 'full time', 'part time', 'company total'))
            if is_banner and (j == 0 or '- ft' in ol or '- pt' in ol or 'full time' in ol or 'part time' in ol):
                pass
            # Render row
            cls = ''
            if 'total' in ol or ol.startswith('company'):
                cls = ' class="total-row"'
            p.append(f'<tr{cls}><td class="sticky-l">{esc(office)}</td>')
            for c, _ in date_cols:
                v = r[c] if c < len(r) else None
                n = pn(v)
                if n is None:
                    p.append('<td class="num zero">·</td>')
                elif is_pct:
                    # Bubble: compare to common target (no per-row target available reliably here) — color by 80% threshold
                    band = 'good' if n >= 0.80 else ('mid' if n >= 0.70 else 'bad')
                    pill_cls = 'good' if band == 'good' else ('neutral' if band == 'mid' else 'bad')
                    p.append(f'<td class="cell"><span class="pill {pill_cls}">{n*100:.0f}%</span></td>')
                elif n == 0:
                    p.append('<td class="num zero">·</td>')
                else:
                    p.append(f'<td class="num">{fmt_int(n) if n == int(n) else fmt_num(n, 1)}</td>')
            p.append('</tr>')
        p.append('</tbody></table></div></div>')
    return ''.join(p)
def _ot_pill(value, total):
    """OT-specific bubble: red >1% of total, yellow 0-1%, green 0%."""
    v = pn(value)
    if v is None or v == 0:
        return '<span class="pill good">0%</span>' if v == 0 else ''
    t = pn(total)
    if t is None or t == 0:
        return f'<span class="pill na">{fmt_num(v,1)}</span>'
    pct = v / t
    if pct > 0.01: cls = 'bad'
    elif pct > 0: cls = 'warn'
    else: cls = 'good'
    return f'<span class="pill {cls}">{pct*100:.1f}%</span>'


def render_overtime(data):
    """Overtime: vertical sections (Total, T&M, FF, Overhead) by office × month with bubble pills."""
    rows = data['rows']
    # Find sections: row containing 'Overtime Summary' followed by date header row
    # Multiple section labels: 'Overtime Summary: Total', 'T&M Projects', 'FF Projects', 'Overhead'
    section_keys = ['Overtime Summary: Total', 'T&M Projects', 'FF Projects', 'Overhead']
    section_starts = []
    for i, r in enumerate(rows):
        for v in r:
            if isinstance(v, str):
                s = v.strip()
                for key in section_keys:
                    if key.lower() in s.lower():
                        section_starts.append((s, i)); break
    # Also find date header rows (row 6 typically). Use first one as canonical.
    header_idx = -1
    for i, r in enumerate(rows[:10]):
        if sum(1 for v in r if is_date(v)) >= 3:
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    header = rows[header_idx]
    date_cols = [(c, v) for c, v in enumerate(header) if is_date(v)]

    # Find 'Historical Offices' row to know where to stop
    historic_idx = find_row_with(rows, 'historical', start=header_idx + 1)

    # Group data rows (office rows) between header_idx+1 and historic_idx (or end)
    end_idx = historic_idx if historic_idx > 0 else len(rows)
    data_rows = []
    for i in range(header_idx + 1, end_idx):
        r = rows[i]
        if not r: continue
        office = r[0] if r else ''
        if not office or not isinstance(office, str): continue
        if office.lower().startswith('historical'): break
        # skip blank rows / total label row
        if office.lower() == 'total': continue
        data_rows.append((i, office, r))

    # Compute totals across all data rows per month for percentage base
    monthly_totals = {}
    for c, _ in date_cols:
        s = 0
        for _, _, r in data_rows:
            v = pn(r[c] if c < len(r) else None)
            if v is not None: s += v
        monthly_totals[c] = s
    grand_total = sum(monthly_totals.values())

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Overtime Summary</h2>')
    p.append('<p class="page-sub">Total overtime hours by office and month. Pills color-coded by share of total: green for zero, yellow for sliver, red above 1%.</p>')
    p.append('</div>')

    # Render single vertical section (T&M / FF / Overhead splits aren't separable from this data shape;
    # show Total as the primary view)
    p.append('<div class="card">')
    p.append('<div class="card-head"><h3>Overtime Summary: Total</h3><span class="card-sub">All overtime hours by office</span></div>')
    p.append('<div class="scroll-card-inner"><table class="data-table"><thead><tr>')
    p.append('<th class="sticky-l">Office</th>')
    for c, dv in date_cols:
        p.append(f'<th class="num">{fmt_month(dv)}</th>')
    p.append('<th class="num">Total</th>')
    p.append('</tr></thead><tbody>')
    for _, office, r in data_rows:
        row_total = 0
        cells = []
        for c, _ in date_cols:
            n = pn(r[c] if c < len(r) else None)
            if n is None: cells.append((c, None)); continue
            row_total += n
            cells.append((c, n))
        # Pill cells using office_total as base for % share within row
        p.append(f'<tr><td class="sticky-l">{esc(office)}</td>')
        for c, n in cells:
            if n is None:
                p.append('<td class="num zero">·</td>')
            elif n == 0:
                p.append('<td class="cell"><span class="pill good">0</span></td>')
            else:
                share = n / row_total if row_total > 0 else 0
                if share > 0.01: cls = 'bad'
                elif share > 0: cls = 'warn'
                else: cls = 'good'
                p.append(f'<td class="cell"><span class="pill {cls}">{fmt_num(n,1)}</span></td>')
        p.append(f'<td class="num total-cell">{fmt_num(row_total,1)}</td></tr>')
    p.append('</tbody></table></div></div>')
    return ''.join(p)


def render_non_posting(data):
    """Non-Posting Hours: cleaned up with % bubble (red>3%, yellow 1-3%, green 0%)."""
    rows = data['rows']
    header_idx = -1
    for i, r in enumerate(rows[:10]):
        if sum(1 for v in r if is_date(v)) >= 3:
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    header = rows[header_idx]
    all_date_cols = [(c, v) for c, v in enumerate(header) if is_date(v)]

    # Collect data rows (office rows), stopping at 'Total' or end
    data_rows = []
    total_row = None
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r: continue
        office = r[0] if r else ''
        if not office or not isinstance(office, str): continue
        if office.lower() == 'total':
            total_row = r; break
        data_rows.append((office, r))

    # Filter date columns: only those where at least one office has a non-zero value
    used_cols = []
    for c, dv in all_date_cols:
        for office, r in data_rows:
            if pn(r[c] if c < len(r) else None) not in (None, 0):
                used_cols.append((c, dv)); break
    if not used_cols: used_cols = all_date_cols  # fallback

    # Compute each office's row total and overall total
    office_totals = {}
    for office, r in data_rows:
        s = 0
        for c, _ in used_cols:
            n = pn(r[c] if c < len(r) else None)
            if n is not None: s += n
        office_totals[office] = s
    grand_total = sum(office_totals.values())

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Non-Posting Hours by Office</h2>')
    p.append('<p class="page-sub">Monthly non-posting overhead hours per office. Pills show share of total — green at zero, yellow under 1%, red above 3%.</p>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="data-table"><thead><tr>')
    p.append('<th class="sticky-l">Office</th>')
    for c, dv in used_cols:
        p.append(f'<th class="num">{fmt_month(dv)}</th>')
    p.append('<th class="num">Total hrs</th><th class="num">% of total</th>')
    p.append('</tr></thead><tbody>')
    for office, r in data_rows:
        rt = office_totals.get(office, 0)
        pct = (rt / grand_total) if grand_total > 0 else 0
        p.append(f'<tr><td class="sticky-l">{esc(office)}</td>')
        for c, _ in used_cols:
            n = pn(r[c] if c < len(r) else None)
            if n is None or n == 0:
                p.append('<td class="num zero">·</td>')
            else:
                p.append(f'<td class="num">{int(round(n)):,}</td>')
        p.append(f'<td class="num total-cell">{int(round(rt)):,}</td>')
        if pct == 0: pcls = 'good'
        elif pct > 0.03: pcls = 'bad'
        elif pct > 0.01: pcls = 'warn'
        else: pcls = 'good'
        p.append(f'<td class="num"><span class="pill {pcls}">{pct*100:.1f}%</span></td>')
        p.append('</tr>')
    if total_row:
        p.append('<tr class="total-row"><td class="sticky-l">Total</td>')
        for c, _ in used_cols:
            n = pn(total_row[c] if c < len(total_row) else None)
            p.append(f'<td class="num">{int(round(n)) if n is not None else "":,}</td>' if n is not None else '<td></td>')
        p.append(f'<td class="num total-cell">{int(round(grand_total)):,}</td><td class="num">100.0%</td></tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def render_overhead(data):
    """Overhead (was the OVHD employee summary sheet): show only excel cols B-P, with F-J as 'YTD 2026 #' and L-P as 'YTD 2026 %' sections."""
    rows = data['rows']
    # Find header row at row 3 (PMO, Employee Name, Position, Target, Target Variance, then YTD blocks)
    header_idx = -1
    for i, r in enumerate(rows[:6]):
        if r and any(isinstance(v, str) and 'employee' in v.lower() for v in r):
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    headers = rows[header_idx]
    # B-P => indices 1-15 (0-indexed); F-J => 5-9; L-P => 11-15
    # Build column list with two sections
    col_specs = []
    # Left: B-E
    for c in range(1, 5):  # B,C,D,E
        if c < len(headers) and headers[c] not in (None, ''):
            col_specs.append({'idx': c, 'name': str(headers[c]).replace('\n',' ').strip(), 'sec': ''})
    # Middle: F-J (YTD 2026 #)
    for c in range(5, 10):
        if c < len(headers) and headers[c] not in (None, ''):
            col_specs.append({'idx': c, 'name': str(headers[c]).replace('\n',' ').strip(), 'sec': 'YTD 2026 #'})
    # Right: L-P (YTD 2026 %)
    for c in range(11, 16):
        if c < len(headers) and headers[c] not in (None, ''):
            col_specs.append({'idx': c, 'name': str(headers[c]).replace('\n',' ').strip(), 'sec': 'YTD 2026 %'})
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Overhead by Employee</h2>')
    p.append('<p class="page-sub">YTD overhead allocation per employee: raw hours buckets on the left, percentage allocation on the right.</p>')
    p.append('</div>')
    p.append('<div class="toolbar"><input type="text" id="ovhdSearch" class="input" placeholder="Search employee, PMO, or position…" /></div>')
    p.append('<div class="scroll-card"><table class="data-table"><thead>')
    # Section header row
    p.append('<tr class="hdr-group">')
    last_sec = None; span = 0; group_start_i = 0
    i = 0
    while i < len(col_specs):
        sec = col_specs[i]['sec']
        j = i
        while j < len(col_specs) and col_specs[j]['sec'] == sec:
            j += 1
        if sec:
            sec_cls = "sec-" + sec.split()[0].lower()
            p.append(f'<th colspan="{j-i}" class="{sec_cls}">{esc(sec)}</th>')
        else:
            p.append(f'<th colspan="{j-i}"></th>')
        i = j
    p.append('</tr><tr>')
    for col in col_specs:
        nl = col['name'].lower()
        is_num = col['sec'] != '' or 'target' in nl or 'variance' in nl
        cls = 'num' if is_num else 'sticky-l'
        p.append(f'<th class="{cls}">{esc(col["name"])}</th>')
    p.append('</tr></thead><tbody>')
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r or not any(v not in (None, '') for v in r): continue
        rtext = ' '.join(str(v) for v in r if v).lower()
        p.append(f'<tr data-search="{esc(rtext)}">')
        for col in col_specs:
            v = r[col['idx']] if col['idx'] < len(r) else None
            nl = col['name'].lower()
            if col['sec'] == 'YTD 2026 %':
                disp = fmt_pct(v, 1) if pn(v) is not None else (esc(str(v)) if v not in (None, '') else '')
                p.append(f'<td class="num">{disp}</td>')
            elif col['sec'] == 'YTD 2026 #':
                disp = fmt_num(v, 1) if pn(v) is not None else (esc(str(v)) if v not in (None, '') else '')
                p.append(f'<td class="num">{disp}</td>')
            elif 'target' in nl and 'variance' not in nl:
                disp = fmt_pct(v, 0) if pn(v) is not None else (esc(str(v)) if v not in (None, '') else '')
                p.append(f'<td class="num">{disp}</td>')
            elif 'variance' in nl:
                n = pn(v)
                if n is None: disp = esc(str(v)) if v not in (None, '') else ''
                else: disp = f'<span class="delta {"pos" if n>=0 else "neg"}">{"+" if n>=0 else ""}{n*100:.1f}pts</span>'
                p.append(f'<td class="num">{disp}</td>')
            else:
                disp = esc(str(v)) if v not in (None, '') else ''
                p.append(f'<td>{disp}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('<script>(function(){var s=document.getElementById("ovhdSearch");if(!s) return;var rows=document.querySelectorAll(".scroll-card tbody tr");s.addEventListener("input",function(){var q=s.value.toLowerCase().trim();rows.forEach(function(r){r.classList.toggle("hidden",q && !(r.dataset.search||"").includes(q));});});})();</script>')
    return ''.join(p)


def render_pto(data):
    """Paid Time Off: hide ID column, tighter target/variance columns."""
    rows = data['rows']
    header_idx = -1
    for i, r in enumerate(rows[:8]):
        if r and any(isinstance(v, str) and 'employee name' in v.lower() for v in r):
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    headers = rows[header_idx]
    # All columns except ID (col 0 is ID#)
    col_specs = []
    for c, h in enumerate(headers):
        if c == 0: continue  # skip ID #
        if h in (None, ''): continue
        col_specs.append({'idx': c, 'name': str(h).replace('\n', ' ').strip()})
    # Find target & ytd columns for pill
    target_col = None; ytd_col = None; var_col = None
    for cs in col_specs:
        nl = cs['name'].lower()
        if target_col is None and ('target' in nl and 'variance' not in nl): target_col = cs['idx']
        if ytd_col is None and ('ytd' in nl and 'actual' in nl): ytd_col = cs['idx']
        if var_col is None and 'variance' in nl: var_col = cs['idx']
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Paid Time Off</h2>')
    p.append('<p class="page-sub">Short-term and extended PTO balances per employee, with utilization vs target.</p>')
    p.append('</div>')
    p.append('<div class="toolbar"><input type="text" id="ptoSearch" class="input" placeholder="Search employee, office…" /></div>')
    p.append('<div class="scroll-card"><table class="data-table pto-table"><thead><tr>')
    for cs in col_specs:
        nl = cs['name'].lower()
        # Allow wrapping on long target/variance headers
        wrap_cls = ' wrap-narrow' if any(k in nl for k in ('utilization', 'target', 'variance')) else ''
        is_num = any(k in nl for k in ('utiliz', 'target', 'variance', 'pto', 'hour'))
        p.append(f'<th class="{"num" if is_num else ""}{wrap_cls}">{esc(cs["name"])}</th>')
    p.append('</tr></thead><tbody>')
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r or not any(v not in (None, '') for v in r): continue
        # Skip placeholder rows where most cells are 0 and look like garbage
        first_real = r[1] if len(r) > 1 else None
        if isinstance(first_real, str) and first_real == '0': continue
        rtext = ' '.join(str(v) for v in r if v not in (None,'')).lower()
        p.append(f'<tr data-search="{esc(rtext)}">')
        for cs in col_specs:
            v = r[cs['idx']] if cs['idx'] < len(r) else None
            nl = cs['name'].lower()
            if cs['idx'] == ytd_col and target_col is not None:
                target_val = r[target_col] if target_col < len(r) else None
                p.append(f'<td class="num">{pill_cmp(v, target_val)}</td>')
            elif cs['idx'] == target_col:
                p.append(f'<td class="num">{fmt_pct(v, 0) if pn(v) is not None else esc(str(v))}</td>')
            elif cs['idx'] == var_col:
                n = pn(v)
                if n is None: disp = ''
                else: disp = f'<span class="delta {"pos" if n>=0 else "neg"}">{"+" if n>=0 else ""}{n*100:.1f}pts</span>'
                p.append(f'<td class="num">{disp}</td>')
            elif 'pto' in nl or 'hour' in nl:
                _pn = pn(v)
                if _pn is None:
                    p.append(f'<td class="num">{esc(str(v)) if v not in (None,"") else ""}</td>')
                else:
                    p.append(f'<td class="num">{_pn:,.1f}</td>')
            elif pn(v) is not None and 0 < abs(pn(v)) < 2 and pn(v) != int(pn(v)):
                p.append(f'<td class="num">{fmt_pct(v, 1)}</td>')
            elif pn(v) is not None:
                p.append(f'<td class="num">{fmt_num(v, 1)}</td>')
            elif is_date(v):
                p.append(f'<td class="date">{fmt_date(v)}</td>')
            else:
                p.append(f'<td>{esc(str(v)) if v not in (None,"") else ""}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('<script>(function(){var s=document.getElementById("ptoSearch");if(!s) return;var rows=document.querySelectorAll(".pto-table tbody tr");s.addEventListener("input",function(){var q=s.value.toLowerCase().trim();rows.forEach(function(r){r.classList.toggle("hidden",q && !(r.dataset.search||"").includes(q));});});})();</script>')
    return ''.join(p)


def render_goal_planning(data):
    """Goal Planning: only Excel cols A, C, F, G, H, J, K. Office filter + name search."""
    rows = data['rows']
    header_idx = -1
    for i, r in enumerate(rows[:8]):
        if r and any(isinstance(v, str) and 'employee name' in v.lower() for v in r):
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    headers = rows[header_idx]
    # Excel cols A=0, C=2, F=5, G=6, H=7, J=9, K=10
    keep_idx = [0, 2, 5, 6, 7, 9, 10]
    col_specs = []
    for c in keep_idx:
        if c < len(headers) and headers[c] not in (None, ''):
            col_specs.append({'idx': c, 'name': str(headers[c]).replace('\n', ' ').strip()})
    target_col = None; ytd_col = None
    for cs in col_specs:
        nl = cs['name'].lower()
        if target_col is None and nl == 'target %': target_col = cs['idx']
        if ytd_col is None and 'ytd actual' in nl: ytd_col = cs['idx']
    # Build office list
    offices = sorted({r[0] for i, r in enumerate(rows[header_idx+1:]) if r and isinstance(r[0], str) and r[0]})
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Goal Planning</h2>')
    p.append('<p class="page-sub">Year-end goals per employee, with additional hours needed and weeks remaining to make target.</p>')
    p.append('</div>')
    p.append('<div class="toolbar">')
    p.append('<input type="text" id="goalSearch" class="input" placeholder="Search employee name…" />')
    p.append('<select id="goalOffice" class="input"><option value="">All offices</option>')
    for o in offices: p.append(f'<option value="{esc(o)}">{esc(o)}</option>')
    p.append('</select></div>')
    p.append('<div class="scroll-card"><table class="data-table goal-table"><thead><tr>')
    for cs in col_specs:
        nl = cs['name'].lower()
        wrap_cls = ' wrap-narrow' if 'additional hours' in nl or 'weeks remaining' in nl else ''
        is_num = any(k in nl for k in ('target', 'actual', 'hours', 'weeks', '%'))
        p.append(f'<th class="{"num" if is_num else ""}{wrap_cls}">{esc(cs["name"])}</th>')
    p.append('</tr></thead><tbody>')
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r or not any(v not in (None, '') for v in r): continue
        office = r[0] if r else ''
        name_val = r[2] if len(r) > 2 else ''
        p.append(f'<tr data-office="{esc(office)}" data-name="{esc(str(name_val).lower())}">')
        for cs in col_specs:
            v = r[cs['idx']] if cs['idx'] < len(r) else None
            nl = cs['name'].lower()
            if cs['idx'] == ytd_col and target_col is not None:
                target_val = r[target_col] if target_col < len(r) else None
                p.append(f'<td class="num">{pill_cmp(v, target_val)}</td>')
            elif 'target' in nl and '%' in cs['name']:
                p.append(f'<td class="num">{fmt_pct(v, 0) if pn(v) is not None else ""}</td>')
            elif 'additional hours' in nl or 'weeks remaining' in nl:
                p.append(f'<td class="num">{fmt_int(v) if pn(v) is not None else ""}</td>')
            elif is_date(v):
                p.append(f'<td class="date">{fmt_date(v)}</td>')
            else:
                p.append(f'<td>{esc(str(v)) if v not in (None,"") else ""}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('<script>(function(){var s=document.getElementById("goalSearch"),o=document.getElementById("goalOffice");if(!s) return;var rows=document.querySelectorAll(".goal-table tbody tr");function ap(){var q=s.value.toLowerCase().trim(),of=o.value;rows.forEach(function(r){var ok=true;if(q&&!(r.dataset.name||"").includes(q))ok=false;if(of&&r.dataset.office!==of)ok=false;r.classList.toggle("hidden",!ok);});}s.addEventListener("input",ap);o.addEventListener("change",ap);})();</script>')
    return ''.join(p)
def render_proposal_hours(data):
    """Proposal Hours: color-code office-level rows; show dates as Week + Month."""
    rows = data['rows']
    header_idx = -1
    for i, r in enumerate(rows[:10]):
        if r and sum(1 for v in r if is_date(v)) >= 3:
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    headers = rows[header_idx]
    label_cols = []; date_cols = []; total_col = None
    for c, h in enumerate(headers):
        if is_date(h): date_cols.append((c, h))
        elif h not in (None, ''):
            h_str = str(h).replace('\n', ' ').strip()
            if 'total' in h_str.lower(): total_col = (c, h_str)
            else: label_cols.append((c, h_str))
    # Date label: "Week N - Mon"
    def week_month(d):
        try:
            dt = _dt.datetime.fromisoformat(d['__date__'])
            wk = ((dt - _dt.datetime(dt.year, 1, 1)).days // 7) + 1
            return f'Wk {wk} · {dt.strftime("%b")}'
        except: return ''
    # Identify office vs name rows. Office rows are e.g. 'Big Sky' (matches a known set).
    # Heuristic: a row is an "office" row if its first cell matches a known office name pattern (Title Case city)
    # Or simpler: if the first cell looks like a top-level location (alphabetic, no comma).
    KNOWN_OFFICES = {'-','International (Global LLC)','Big Sky','Bluegrass','Built Environment','Cascade','Chesapeake','Coastal Bend',
                     'Corporate','Front Range','Geospatial','Great Lakes','Gulf Coast','High Desert','Hill Country','International',
                     'Island','Lakeshore','Maritime','Meridian Central','Meridian Southwest','Paleontology','Piedmont','Prairie',
                     'Red Rock','Sierra','Archive Services'}
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Proposal Hours</h2>')
    p.append('<p class="page-sub">Time charged to proposals, grouped by office.</p>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="data-table prop-table"><thead><tr>')
    for c, name in label_cols:
        p.append(f'<th class="sticky-l">{esc(name)}</th>')
    for c, dv in date_cols:
        p.append(f'<th class="num">{week_month(dv)}</th>')
    if total_col: p.append(f'<th class="num">Total</th>')
    p.append('</tr></thead><tbody>')
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r: continue
        if not any(v not in (None, '') for v in r): continue
        first = r[0] if r else ''
        is_office = isinstance(first, str) and first in KNOWN_OFFICES
        cls = ' class="office-level"' if is_office else ''
        p.append(f'<tr{cls}>')
        for c, _ in label_cols:
            v = r[c] if c < len(r) else None
            p.append(f'<td class="sticky-l">{esc(str(v)) if v not in (None,"") else ""}</td>')
        for c, _ in date_cols:
            n = pn(r[c] if c < len(r) else None)
            if n is None or n == 0:
                p.append('<td class="num zero">·</td>')
            else:
                p.append(f'<td class="num">{fmt_num(n, 1)}</td>')
        if total_col:
            v = r[total_col[0]] if total_col[0] < len(r) else None
            p.append(f'<td class="num total-cell">{fmt_num(v, 1)}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def render_non_targeted(data):
    """Non-Targeted Employees — mirror Employee Utilization style.
    Source-sheet column shape:
      col 0: status flag ('-')
      col 1: 'Active' / employee status
      col 2: employee id
      col 3: 'Last, First' name  OR  section banner string (e.g. 'REGIONAL PRINCIPAL')
      col 4: Reg/Temp
      col 5: Target %  (often 0)
      col 6: Actual %  (YTD)
      col 7: Deviation
      col 9+ : weekly values
    """
    rows = data['rows']
    if len(rows) < 8: return render_premium_generic(data)

    # Locate the date-header row (the one with many datetime cells)
    date_row_idx = -1
    for i, r in enumerate(rows[:10]):
        if r and sum(1 for v in r if is_date(v)) >= 5:
            date_row_idx = i; break
    if date_row_idx < 0: return render_premium_generic(data)
    date_row = rows[date_row_idx]
    date_cols = [(c, v) for c, v in enumerate(date_row) if is_date(v)]
    # All weeks, not the first 20 (FP&A, 2026-09-17).  The 20-column cap froze this tab at
    # mid-May while every other tab rolled forward with the weekly refresh.  Future weeks are
    # dimmed exactly like the Employee Utilization grid; the table opens on the latest real week.
    show_cols = date_cols
    week_future = [not (to_date(v) and to_date(v) <= TODAY_REF) for _, v in show_cols]

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Non-Targeted Employees</h2>')
    p.append('<p class="page-sub">Employees without a utilization target — section grouped by role. Pills follow the same color rules as Employee Utilization.</p>')
    p.append('</div>')
    p.append('<div class="toolbar"><input type="text" id="ntSearch" class="input" placeholder="Search employee, ID, or role…" /></div>')
    # NOT scroll-flat: 52 week columns need the grid's own horizontal scroller, and with it the
    # sticky identity columns and sticky header — the same treatment Employee Utilization gets.
    p.append('<div class="scroll-card nt-scroll"><table class="bubble-grid nt-grid"><thead><tr>')
    p.append('<th class="col-name">Employee</th>')
    p.append('<th class="col-type">Type</th>')
    p.append('<th class="col-target">Target</th>')
    p.append('<th class="ytd-sep">YTD</th>')
    for i, (c, dv) in enumerate(show_cols):
        cls = 'num week-cell' + (' week-future' if week_future[i] else '')
        p.append(f'<th class="{cls}" data-col="{c}">{fmt_date(dv, short=True)}</th>')
    p.append('</tr></thead><tbody>')

    def green_pill(v):
        """Non-targeted employees have no target — always render the % in a green pill."""
        n = pn(v)
        if n is None:
            return '<span class="muted">—</span>'
        if isinstance(v, str) and v.startswith('#'):
            return '<span class="muted">—</span>'
        return f'<span class="pill good">{n*100:.0f}%</span>'

    total_cols = 4 + len(show_cols)
    current_section = ''
    for i in range(date_row_idx + 1, len(rows)):
        r = rows[i]
        if not r: continue
        col3 = r[3] if len(r) > 3 else None
        col2 = r[2] if len(r) > 2 else None
        col1 = r[1] if len(r) > 1 else None
        col4 = r[4] if len(r) > 4 else None
        # Section banner: col 3 is a non-empty string and not a "Last, First" name
        if isinstance(col3, str) and col3.strip() and ',' not in col3 and not isinstance(col2, (int, float)):
            sec = col3.strip()
            # skip purely numeric header rows (col labels '1','2','3' ...)
            if sec and not sec.isdigit():
                current_section = sec
                p.append(f'<tr class="office-group" data-section="{esc(sec)}"><td colspan="{total_cols}"><span class="nt-band-label">{esc(sec.upper())}</span></td></tr>')
            continue
        # Employee row: col 3 must be "Last, First"
        if not isinstance(col3, str) or ',' not in col3:
            continue
        name = col3.strip()
        emp_id = col2
        regtemp = col4 if isinstance(col4, str) else ''
        # FT/PT short label
        type_lbl = ''
        rl = regtemp.lower()
        if 'full' in rl: type_lbl = 'FT'
        elif 'part' in rl: type_lbl = 'PT'
        target = r[5] if len(r) > 5 else None
        ytd = r[6] if len(r) > 6 else None
        rtext = f'{name} {emp_id} {current_section} {regtemp}'.lower()
        p.append(f'<tr class="emp-row" data-search="{esc(rtext)}" data-section="{esc(current_section)}">')
        p.append(f'<td class="col-name">{esc(name)}<br><span class="emp-id">{esc(str(emp_id))}</span></td>')
        p.append(f'<td class="col-type">{esc(type_lbl)}</td>')
        t_disp = fmt_pct(target) if pn(target) is not None and pn(target) != 0 else '<span class="muted">—</span>'
        p.append(f'<td class="col-target">{t_disp}</td>')
        p.append(f'<td class="cell ytd-sep">{green_pill(ytd)}</td>')
        for i, (c, _) in enumerate(show_cols):
            v = r[c] if c < len(r) else None
            cls = 'cell week-future' if week_future[i] else 'cell'
            p.append(f'<td class="{cls}">{green_pill(v) if v not in (None,"") else ""}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('<script>(function(){var s=document.getElementById("ntSearch");if(!s) return;var rows=document.querySelectorAll(".nt-grid tbody tr.emp-row"),grps=document.querySelectorAll(".nt-grid tbody tr.office-group");function ap(){var q=(s.value||"").toLowerCase().trim(),seen=new Set();rows.forEach(function(r){var ok=!q||(r.dataset.search||"").includes(q);r.classList.toggle("hidden",!ok);if(ok)seen.add(r.dataset.section);});grps.forEach(function(g){g.classList.toggle("hidden",q&&!seen.has(g.dataset.section));});}s.addEventListener("input",ap);})();</script>')
    # Open scrolled to the latest real week (the tab starts hidden, so re-run on activation).
    p.append('<script>(function(){var t=document.querySelector(".nt-grid");if(!t)return;var box=t.closest(".scroll-card");if(!box)return;function jump(){var w=t.querySelectorAll("thead th.week-cell:not(.week-future)");if(!w.length)return;var r=w[w.length-1].getBoundingClientRect(),b=box.getBoundingClientRect();var d=(r.right-b.right)+16;if(d>0)box.scrollLeft+=d;}function pin(){var row=t.querySelector("tbody tr.emp-row");if(!row||row.children.length<4)return;var c=row.children,w=function(i){return c[i].getBoundingClientRect().width;};var a=w(0),b2=a+w(1),c3=b2+w(2);t.style.setProperty("--nt-l2",a+"px");t.style.setProperty("--nt-l3",b2+"px");t.style.setProperty("--nt-l4",c3+"px");}function go(){pin();jump();}var sec=document.getElementById("Non_Targeted");if(sec){new MutationObserver(function(){if(sec.classList.contains("active"))setTimeout(go,30);}).observe(sec,{attributes:true,attributeFilter:["class"]});}window.addEventListener("resize",pin);setTimeout(go,80);})();</script>')
    return ''.join(p)


def render_pt_activity(data):
    """Part Time Employee Activity: Name, Department, Position, Date of Last Time Submission,
    most recent 8 weeks (weekly dates), then 4 Weeks, 8 Weeks, 12 Weeks, Prior Week rollups.
    Filters: Hire Date, Date of Last Time Submission."""
    rows = data['rows']
    if not rows: return '<p class="empty">No data.</p>'
    headers = rows[0]
    # Identify columns
    name_c = dept_c = pos_c = hire_c = lts_c = pw_c = w4_c = w8_c = w12_c = None
    date_cols = []
    for c, h in enumerate(headers):
        if h in (None, ''): continue
        if is_date(h):
            date_cols.append((c, h)); continue
        nl = str(h).replace('\n',' ').strip().lower()
        if 'name' in nl and name_c is None: name_c = c
        elif 'department' in nl and dept_c is None: dept_c = c
        elif 'position' in nl and pos_c is None: pos_c = c
        elif 'hire' in nl and hire_c is None: hire_c = c
        elif 'last time' in nl and lts_c is None: lts_c = c
        elif nl == '4 weeks': w4_c = c
        elif nl == '8 weeks': w8_c = c
        elif nl == '12 weeks': w12_c = c
        elif nl == 'prior week': pw_c = c
    # Most recent 8 weekly date cols that are past
    past_dates = [(c, v) for c, v in date_cols if to_date(v) and to_date(v) <= TODAY_REF]
    recent_weeks = past_dates[-8:] if past_dates else date_cols[-8:]

    # Stop processing rows once we hit the "Utilization of Available on Call Trend" section
    body_end = len(rows)
    for i in range(1, len(rows)):
        for v in rows[i]:
            if isinstance(v, str) and 'utilization of available' in v.lower():
                body_end = i; break
        if body_end < len(rows): break

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Part-Time Employee Activity</h2>')
    p.append('<p class="page-sub">Recent timesheet activity for part-time staff. Pills show weekly utilization percentages; rollups summarize the last 4, 8, and 12 weeks plus prior week.</p>')
    p.append('</div>')
    p.append('<div class="toolbar">')
    p.append('<input type="text" id="ptSearch" class="input" placeholder="Search name, department, position…" />')
    p.append('<select id="ptHire" class="input"><option value="">Any hire date</option><option value="recent">Hired in last 12 months</option><option value="1to3">1–3 years ago</option><option value="3plus">3+ years ago</option></select>')
    p.append('<select id="ptLts" class="input"><option value="">Any last-submission</option><option value="6mo">No time in last 6 months</option><option value="1yr">No time in last 12 months</option><option value="never">No time recorded ever</option></select>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="data-table pt-table"><thead><tr>')
    p.append('<th class="sticky-l">Name</th>')
    p.append('<th>Department</th><th>Position</th><th>Last Submission</th>')
    for c, dv in recent_weeks:
        p.append(f'<th class="num">{fmt_date(dv, short=True)}</th>')
    p.append('<th class="num">Prior Wk</th><th class="num">4 Wks</th><th class="num">8 Wks</th><th class="num">12 Wks</th>')
    p.append('</tr></thead><tbody>')
    today_dt = TODAY_REF
    for i in range(1, body_end):
        r = rows[i]
        if not r or not any(v not in (None, '') for v in r): continue
        if name_c is None or not r[name_c]: continue
        name = str(r[name_c]) if name_c < len(r) else ''
        if ',' not in name: continue
        dept = r[dept_c] if dept_c is not None and dept_c < len(r) else ''
        pos = r[pos_c] if pos_c is not None and pos_c < len(r) else ''
        hire = r[hire_c] if hire_c is not None and hire_c < len(r) else None
        lts = r[lts_c] if lts_c is not None and lts_c < len(r) else None
        # Compute hire-age and lts-age categories
        hire_cat = ''
        if is_date(hire):
            d = to_date(hire)
            if d:
                years = (today_dt - d).days / 365.25
                if years <= 1: hire_cat = 'recent'
                elif years <= 3: hire_cat = '1to3'
                else: hire_cat = '3plus'
        lts_cat = ''
        if is_date(lts):
            d = to_date(lts)
            if d:
                days = (today_dt - d).days
                if days > 365: lts_cat = '1yr'
                elif days > 180: lts_cat = '6mo'
        else:
            lts_cat = 'never'
        rtext = f'{name} {dept} {pos}'.lower()
        p.append(f'<tr data-search="{esc(rtext)}" data-hire="{hire_cat}" data-lts="{lts_cat}">')
        p.append(f'<td class="sticky-l">{esc(name)}</td>')
        p.append(f'<td>{esc(str(dept)) if dept else ""}</td>')
        p.append(f'<td>{esc(str(pos)) if pos else ""}</td>')
        p.append(f'<td class="date">{fmt_date(lts) if is_date(lts) else ""}</td>')
        for c, _ in recent_weeks:
            v = r[c] if c < len(r) else None
            n = pn(v)
            if n is None:
                p.append('<td class="num zero">·</td>')
            else:
                # Bubble: any utilization > 0 is green-ish, 0 muted
                if n > 0:
                    p.append(f'<td class="cell"><span class="pill good">{n*100:.0f}%</span></td>')
                else:
                    p.append('<td class="num zero">0</td>')
        for col in (pw_c, w4_c, w8_c, w12_c):
            if col is not None and col < len(r):
                v = r[col]
                n = pn(v)
                disp = (f'{n*100:.0f}%' if (n is not None and abs(n) <= 2) else fmt_num(v, 1)) if n is not None else ''
                p.append(f'<td class="num">{disp}</td>')
            else:
                p.append('<td></td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('''<script>(function(){var s=document.getElementById("ptSearch"),h=document.getElementById("ptHire"),l=document.getElementById("ptLts");if(!s) return;var rows=document.querySelectorAll(".pt-table tbody tr");function ap(){var q=s.value.toLowerCase().trim(),hv=h.value,lv=l.value;rows.forEach(function(r){var ok=true;if(q&&!(r.dataset.search||"").includes(q))ok=false;if(hv&&r.dataset.hire!==hv)ok=false;if(lv&&r.dataset.lts!==lv)ok=false;r.classList.toggle("hidden",!ok);});}s.addEventListener("input",ap);h.addEventListener("change",ap);l.addEventListener("change",ap);})();</script>''')
    return ''.join(p)


def render_on_call_trend(data):
    """Secondary tab: 'Utilization of Available on Call Trend' section from PT Hours."""
    rows = data['rows']
    # Find section start
    sec_start = -1
    for i, r in enumerate(rows):
        for v in r:
            if isinstance(v, str) and 'utilization of available' in v.lower():
                sec_start = i; break
        if sec_start != -1: break
    if sec_start == -1: return '<p class="empty">On Call Trend section not found.</p>'

    # The section has: header row with dates, then rows for "# of On Call Utilized", "# of On Call Available", "# Utilized / Available"
    # Header dates are in row sec_start+1 or so
    # Find row with dates
    hdr_idx = -1
    for i in range(sec_start, min(sec_start + 8, len(rows))):
        if sum(1 for v in rows[i] if is_date(v)) >= 3:
            hdr_idx = i; break
    if hdr_idx == -1: return '<p class="empty">On Call Trend headers not found.</p>'
    header = rows[hdr_idx]
    date_cols = [(c, v) for c, v in enumerate(header) if is_date(v)]
    # Data rows after header
    data_rows = []
    for i in range(hdr_idx + 1, min(hdr_idx + 8, len(rows))):
        r = rows[i]
        if not r: continue
        label = r[0] if r else (r[1] if len(r) > 1 else '')
        if not label:
            # Try col 2
            label = r[2] if len(r) > 2 else ''
        if isinstance(label, str) and label.strip():
            data_rows.append((label, r))

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>On Call Utilization Trend</h2>')
    p.append('<p class="page-sub">Weekly trend of part-time on-call staff that were utilized versus available.</p>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="data-table"><thead><tr>')
    p.append('<th class="sticky-l">Metric</th>')
    for c, dv in date_cols:
        p.append(f'<th class="num">{fmt_date(dv, short=True)}</th>')
    p.append('</tr></thead><tbody>')
    for label, r in data_rows:
        is_pct = '/' in label  # "# Utilized / Available" → percentage
        p.append(f'<tr><td class="sticky-l">{esc(label)}</td>')
        for c, _ in date_cols:
            v = r[c] if c < len(r) else None
            n = pn(v)
            if n is None:
                p.append('<td class="num zero">·</td>')
            elif is_pct:
                p.append(f'<td class="num">{fmt_pct(n, 1)}</td>')
            else:
                p.append(f'<td class="num">{fmt_int(n)}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


# Update RENDERERS dict with new and revised renderers
RENDERERS = {
    'bubble_dashboard': render_bubble_dashboard_v2,
    'exec_summary': render_exec_summary,
    'pmo_period': render_pmo_period_v2,
    'pmo_weekly': render_pmo_weekly_v2,
    'overtime': render_overtime,
    'non_posting': render_non_posting,
    'overhead': render_overhead,
    'pto': render_pto,
    'goal_planning': render_goal_planning,
    'proposal_hours': render_proposal_hours,
    'pivot_hours': render_pivot_hours,
    'non_targeted': render_non_targeted,
    'pt_activity': render_pt_activity,
    'on_call_trend': render_on_call_trend,
    'premium_generic': render_premium_generic,
}
# =============================================================================
# Round 2 edits — updated renderers
# =============================================================================

# Employees whose office label is overridden on the Employee Utilization tab and the
# Employee Dashboard. 2026-09-15, FP&A: the International office carries two people on
# the Global LLC entity, who need to read separately from the rest of the
# International roster. Keyed by employee ID (string); delete an entry to restore.
# Display only -- the PMO Weekly/Monthly tabs still show one combined International row,
# since those come straight from the source workbook.
EMPLOYEE_OFFICE_OVERRIDES = {
    'S4412': 'International (Global LLC)',
    'S4413': 'International (Global LLC)',
}


def _emp_id_str(eid):
    if isinstance(eid, (int, float)):
        try: return str(int(eid))
        except Exception: pass
    return str(eid or '').strip()


def apply_office_overrides(employees):
    """Re-label the offices in EMPLOYEE_OFFICE_OVERRIDES, then regroup.

    The bubble grid emits an office-group header whenever the office changes, so the
    list has to stay grouped. The sort is stable and keyed on office alone, which keeps
    every office's existing internal order and drops a renamed employee into the right
    alphabetical group.
    """
    touched = 0
    for e in employees:
        new = EMPLOYEE_OFFICE_OVERRIDES.get(_emp_id_str(e.get('id')))
        if new and e.get('office') != new:
            e['office'] = new
            touched += 1
    if touched:
        employees.sort(key=lambda e: (e.get('office') or '').lower())
    return employees


def render_bubble_dashboard_v3(_data):
    """Employee Utilization (round 2):
    - "Δ" label instead of "YTD Δ"
    - No triangle indicators in delta cells (just signed pts)
    - Sticky quarterly headers preserved in Quarterly View
    - "Quarterly View" toggle instead of "Compact view"
    - New filter: 'Submitting time YTD' (excludes employees with N/A YTD)
    """
    with open('employees.json') as f:
        edata = json.load(f)
    employees = apply_office_overrides(edata['employees'])
    date_headers = edata['date_headers']
    WEEKLY_Q1 = list(range(9, 22)); WEEKLY_Q2 = list(range(26, 39))
    WEEKLY_Q3 = list(range(43, 56)); WEEKLY_Q4 = list(range(60, 73))
    QUARTERLY = [23, 40, 57, 74]
    quarters_def = [('Q1', WEEKLY_Q1, 23), ('Q2', WEEKLY_Q2, 40), ('Q3', WEEKLY_Q3, 57), ('Q4', WEEKLY_Q4, 74)]
    week_dt = {c: to_date({'__date__': date_headers[str(c)] + 'T00:00:00'}) for c in WEEKLY_Q1+WEEKLY_Q2+WEEKLY_Q3+WEEKLY_Q4}
    def week_is_past(c):
        d = week_dt.get(c)
        return d is not None and d <= TODAY_REF
    week_labels = {c: fmt_date({'__date__': date_headers[str(c)] + 'T00:00:00'}, short=True) for c in WEEKLY_Q1+WEEKLY_Q2+WEEKLY_Q3+WEEKLY_Q4}
    offices = sorted(set(e['office'] for e in employees if e['office']))

    def delta_text(actual, target):
        v = pn(actual); t = pn(target)
        if v is None or t is None: return '<span class="delta-muted">—</span>'
        d = v - t
        if abs(d) < 0.0005: return '<span class="delta-muted">·</span>'
        sgn = '+' if d > 0 else '−'
        cls = 'pos' if d > 0 else 'neg'
        return f'<span class="delta {cls}">{sgn}{abs(d)*100:.1f}</span>'

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Employee Utilization</h2>')
    p.append(f'<p class="page-sub">Weekly performance for {len(employees)} billable employees across {len(offices)} offices, pilled green when at or above each employee\'s individual target.</p>')
    p.append('</div>')
    p.append('<div class="toolbar">')
    p.append('<input type="text" id="bubbleSearch" class="input" placeholder="Search employee or ID…" />')
    p.append('<select id="bubbleOffice" class="input"><option value="">All offices</option>')
    for o in offices: p.append(f'<option value="{esc(o)}">{esc(o)}</option>')
    p.append('</select>')
    p.append('<select id="bubbleType" class="input"><option value="">All employees</option><option value="Full Time">Full Time</option><option value="Part Time">Part Time</option></select>')
    p.append('<select id="bubbleYtd" class="input"><option value="">YTD vs target</option><option value="good">At or above target</option><option value="bad">Below target</option><option value="hasdata">Submitting time YTD</option></select>')
    p.append('<label class="toggle"><input type="checkbox" id="bubbleCompact" /> <span>Quarterly View</span></label>')
    p.append('<span class="legend"><span class="pill good">≥ target</span><span class="pill bad">below target</span><span class="pill na">N/A</span></span>')
    p.append('</div>')
    p.append(f'<div class="result-count"><span id="bubbleCount">{len(employees)}</span> of {len(employees)} employees</div>')
    p.append('<div class="scroll-card"><table class="bubble-grid"><thead>')
    p.append('<tr class="row-quarter">')
    p.append('<th class="col-office" rowspan="2">Office</th>')
    p.append('<th class="col-name" rowspan="2">Employee</th>')
    p.append('<th class="col-type" rowspan="2">Type</th>')
    p.append('<th class="col-target" rowspan="2">Target</th>')
    p.append('<th class="ytd-sep" rowspan="2">YTD</th>')
    p.append('<th class="ytd-sep col-delta" rowspan="2">Δ</th>')
    for q_label, weeks, _ in quarters_def:
        span = len(weeks) + 2
        p.append(f'<th class="quarter-sep q-hdr" colspan="{span}">{q_label} 2026</th>')
    p.append('</tr><tr class="row-week">')
    for q_label, weeks, _ in quarters_def:
        for i, c in enumerate(weeks):
            cls = ['week-cell']
            if i == 0: cls.append('quarter-sep')
            if not week_is_past(c): cls.append('week-future')
            p.append(f'<th class="{" ".join(cls)}" data-col="{c}">{week_labels[c]}</th>')
        p.append(f'<th class="q-total">{q_label}</th>')
        p.append(f'<th class="q-delta col-delta">Δ</th>')
    p.append('</tr></thead><tbody>')
    current_office = None
    for emp in employees:
        office = emp['office'] or ''
        if office != current_office:
            current_office = office
            total_cols = 6 + sum(len(w)+2 for _, w, _ in quarters_def)
            p.append(f'<tr class="office-group" data-office-group="{esc(office)}"><td colspan="{total_cols}">{esc(office.upper())}</td></tr>')
        name = emp['name'] or ''
        emp_id = emp['id'] or ''
        target = emp['target']
        regtemp = (emp.get('regtemp') or '').strip()
        # Normalize FT/PT label
        type_label = ''
        rl = regtemp.lower()
        if 'full' in rl: type_label = 'FT'
        elif 'part' in rl: type_label = 'PT'
        else: type_label = regtemp[:2].upper() if regtemp else ''
        v_ytd = pn(emp['ytd']); t_v = pn(target)
        ytd_class = 'na' if v_ytd is None or t_v is None else ('good' if v_ytd >= t_v else 'bad')
        has_ytd = '1' if v_ytd is not None else '0'
        ytd_pill = pill_cmp(emp['ytd'], target)
        ytd_delta = delta_text(emp['ytd'], target)
        p.append(f'<tr class="emp-row" data-office="{esc(office)}" data-name="{esc(str(name).lower())}" data-id="{esc(str(emp_id).lower())}" data-type="{esc(regtemp)}" data-ytd="{ytd_class}" data-hasytd="{has_ytd}">')
        p.append(f'<td class="col-office">{esc(office)}</td>')
        p.append(f'<td class="col-name"><a href="#Employee_Dashboard" class="emp-name-link" data-emp-name="{esc(name)}">{esc(name)}</a><br><span class="emp-id">{esc(str(emp_id))}</span></td>')
        p.append(f'<td class="col-type">{esc(type_label)}</td>')
        t_disp = fmt_pct(target) if pn(target) is not None else '—'
        p.append(f'<td class="col-target">{t_disp}</td>')
        p.append(f'<td class="cell ytd-sep">{ytd_pill}</td>')
        p.append(f'<td class="cell ytd-sep col-delta">{ytd_delta}</td>')
        for q_label, weeks, q_col in quarters_def:
            for i, c in enumerate(weeks):
                cls = ['cell', 'week-cell']
                if i == 0: cls.append('quarter-sep')
                if not week_is_past(c): cls.append('week-future')
                val = emp['weeks'].get(str(c))
                p.append(f'<td class="{" ".join(cls)}" data-col="{c}">{pill_cmp(val, target)}</td>')
            qval = emp['quarters'].get(str(q_col))
            p.append(f'<td class="cell q-total">{pill_cmp(qval, target)}</td>')
            p.append(f'<td class="cell q-delta col-delta">{delta_text(qval, target)}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('''<script>(function(){
var s=document.getElementById("bubbleSearch"),o=document.getElementById("bubbleOffice"),t=document.getElementById("bubbleType"),y=document.getElementById("bubbleYtd"),c=document.getElementById("bubbleCompact"),cnt=document.getElementById("bubbleCount");
if(!s) return;
var rows=document.querySelectorAll("tr.emp-row"),grps=document.querySelectorAll("tr.office-group"),wkCells=document.querySelectorAll("th.week-cell,td.week-cell"),qHdrs=document.querySelectorAll(".q-hdr");
function apply(){var q=s.value.trim().toLowerCase(),off=o.value,tp=t.value,yy=y.value,vis=0,seen=new Set();
rows.forEach(function(r){var ok=true;
if(q&&!((r.dataset.name||"").includes(q)||(r.dataset.id||"").includes(q)))ok=false;
if(off&&r.dataset.office!==off)ok=false;
if(tp&&r.dataset.type!==tp)ok=false;
if(yy==="good"&&r.dataset.ytd!=="good")ok=false;
if(yy==="bad"&&r.dataset.ytd!=="bad")ok=false;
if(yy==="hasdata"&&r.dataset.hasytd!=="1")ok=false;
r.classList.toggle("hidden",!ok);if(ok){vis++;seen.add(r.dataset.office);}});
grps.forEach(function(g){g.classList.toggle("hidden",!seen.has(g.dataset.officeGroup));});
if(cnt) cnt.textContent=vis;}
function compact(){var on=c.checked;
wkCells.forEach(function(w){w.classList.toggle("collapsed",on);});
qHdrs.forEach(function(h){
  if(on){ if(!h.dataset.origCol) h.dataset.origCol = h.getAttribute("colspan"); h.setAttribute("colspan","2"); }
  else { if(h.dataset.origCol) h.setAttribute("colspan", h.dataset.origCol); }
  h.classList.toggle("q-hdr-compact",on);
});
document.querySelector("table.bubble-grid").classList.toggle("compact-mode",on);}
s.addEventListener("input",apply);o.addEventListener("change",apply);t.addEventListener("change",apply);y.addEventListener("change",apply);
c.addEventListener("change",compact);
})();</script>''')
    return ''.join(p)


PMO_MONTHLY_HIDE_OFFICES = {'lakeshore'}


def render_pmo_period_v3(data):
    """PMO Monthly (round 2):
    - Hide threshold rows (>=10%, >=0%, <0%) near bottom
    - Future months collapsed by default with a toggle button
    """
    rows = data['rows']
    if len(rows) < 4: return '<p class="empty">No data.</p>'
    period_row = rows[1]
    period_blocks = []
    c = 1
    while c < len(period_row):
        v = period_row[c]
        if v not in (None, '') and isinstance(v, (str, dict)):
            period_blocks.append((c, v)); c += 3
        else: c += 1
    is_weekly = 'Weekly' in (data.get('tab') or '')
    # Determine which periods are in the future.
    # Periods may be date objects, month names ("January", "February", ...),
    # or other strings like "Total". Compare to TODAY_REF when possible.
    month_names = {
        'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
        'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'jun': 6, 'jul': 7, 'aug': 8,
        'sep': 9, 'sept': 9, 'oct': 10, 'nov': 11, 'dec': 12,
    }
    def is_future(val):
        if is_date(val):
            d = to_date(val)
            return bool(d and d.replace(day=28) > TODAY_REF)
        if isinstance(val, str):
            m = month_names.get(val.strip().lower())
            if m is not None:
                # Treat as current calendar year — future when month > current month
                return m > TODAY_REF.month
        return False
    # Preserve original column order; mark each block as future or not
    ordered_blocks = [(c, v, is_future(v)) for c, v in period_blocks]

    threshold_markers = ('>=10', '>=0', '<0', '> = 10', '> = 0', '< 0')
    def is_threshold(office_text):
        if not office_text: return False
        ol = office_text.lower().replace(' ', '')
        for m in threshold_markers:
            if m.replace(' ', '') in ol: return True
        return False

    p = []
    p.append('<div class="page-intro">')
    p.append(f'<h2>{"Weekly" if is_weekly else "Monthly"} PMO Utilization</h2>')
    p.append('<p class="page-sub">Office-level target vs actual utilization. Actuals pill green when the office met or beat target for the period.</p>')
    p.append('</div>')
    p.append('<div class="toolbar">')
    # 'Show future months' toggle removed at FP&A's request 2026-09-14 -- future
    # months stay collapsed via the table's hide-future class.
    p.append('<span class="legend"><span class="pill good">≥ target</span><span class="pill bad">below target</span></span>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="pmo-grid pmo-narrow hide-future"><thead>')
    p.append('<tr class="hdr-period">')
    p.append('<th class="col-office" rowspan="2">PMO</th>')
    for col_idx, val, fut in ordered_blocks:
        lbl = fmt_month(val) if is_date(val) else esc(str(val))
        fcls = ' future-period' if fut else ''
        p.append(f'<th colspan="3" class="period-sep{fcls}" data-col="{col_idx}">{lbl}</th>')
    p.append('</tr><tr class="hdr-sub">')
    for col_idx, _, fut in ordered_blocks:
        fcls = ' future-period' if fut else ''
        p.append(f'<th class="period-sep{fcls}" data-col="{col_idx}">Tgt</th><th class="{fcls.strip()}" data-col="{col_idx}">Act</th><th class="{fcls.strip()}" data-col="{col_idx}">Δ</th>')
    p.append('</tr></thead><tbody>')
    for i in range(3, len(rows)):
        r = rows[i]
        if not r: continue
        office = r[0] if len(r) > 0 else ''
        if not office or not isinstance(office, str): continue
        if is_threshold(office): continue
        # Offices FP&A asked to keep off the Monthly tab (2026-09-14: Lakeshore).
        # Display only -- the Total/Company row still includes them, since it
        # comes from the source workbook.
        if office.strip().lower() in PMO_MONTHLY_HIDE_OFFICES: continue
        is_total = 'total' in office.lower() or 'company' in office.lower()
        cls = ' class="total-row"' if is_total else ''
        p.append(f'<tr{cls}><td class="col-office">{esc(office)}</td>')
        def render_block(col_idx):
            t_val = r[col_idx] if col_idx < len(r) else None
            a_val = r[col_idx+1] if col_idx+1 < len(r) else None
            d_val = r[col_idx+2] if col_idx+2 < len(r) else None
            t_disp = fmt_pct(t_val, 0) if pn(t_val) is not None else '<span class="muted">—</span>'
            a_pill = pill_cmp(a_val, t_val) or '<span class="muted">—</span>'
            d_n = pn(d_val)
            d_disp = '<span class="muted">—</span>' if d_n is None else f'<span class="delta {"pos" if d_n >= 0 else "neg"}">{"+" if d_n>=0 else ""}{d_n*100:.0f}</span>'
            return t_disp, a_pill, d_disp
        for col_idx, _, fut in ordered_blocks:
            t_disp, a_pill, d_disp = render_block(col_idx)
            fcls = ' future-period' if fut else ''
            p.append(f'<td class="num period-sep{fcls}" data-col="{col_idx}">{t_disp}</td><td class="num{fcls}" data-col="{col_idx}">{a_pill}</td><td class="num{fcls}" data-col="{col_idx}">{d_disp}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def render_pmo_weekly_v3(data):
    """PMO Weekly (round 2):
    - Change "total" -> "Company"; trim 25 rows × 10 weeks below
    - Change Billable -> Full Time; trim 26 rows × 10 weeks below
    - Change Non-Billable -> Part Time; trim 26 rows × 10 weeks below
    - Use weeks BEFORE current week (so May 16 excluded when today is May 18)
    - No blue pills (use good/bad/warn only)
    - Add YTD Actual, Target and Variance columns for each table
    """
    rows = data['rows']
    if len(rows) < 6: return render_premium_generic(data)
    header_idx = -1
    for i, r in enumerate(rows[:10]):
        if sum(1 for v in r if is_date(v)) >= 5:
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    header = rows[header_idx]
    all_date_cols = [(c, v) for c, v in enumerate(header) if is_date(v)]
    # Show weeks strictly before current week's end-of-week (Saturday).
    # Use weeks where (week_date + 6 days) < TODAY_REF, i.e. week is fully complete.
    # Simpler: show weeks where date < (TODAY_REF - 6 days)
    # Weeks are week-ENDING Saturdays, so a week is fully complete once its date has passed.
    cutoff = TODAY_REF + _dt.timedelta(days=1)  # include weeks ending on/before today; excludes the in-progress week
    past_dates = [(c, v) for c, v in all_date_cols if to_date(v) and to_date(v) < cutoff]
    date_cols = past_dates[-10:] if len(past_dates) >= 10 else past_dates

    # Detect sections by walking rows. PMO Weekly stacks 4 sections by office count: Total, Billable, Non-Billable, Other.
    # Section divider rows have most cells empty.
    sections = []
    cur_rows = []
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r:
            if cur_rows: sections.append(cur_rows); cur_rows = []
            continue
        non_empty = sum(1 for v in r if v not in (None, ''))
        col0 = r[0] if len(r) > 0 else ''
        col1 = r[1] if len(r) > 1 else ''
        if non_empty <= 2 and isinstance(col0, str) and col0 and not col1 and not pn(col0):
            if cur_rows: sections.append(cur_rows); cur_rows = []
            continue
        if isinstance(col1, str) and col1 and not pn(col1):
            cur_rows.append(i)
        elif isinstance(col0, str) and col0 and not col0.lower().startswith('company'):
            cur_rows.append(i)
    if cur_rows: sections.append(cur_rows)

    # Rename map
    section_titles = ['Company', 'Full Time', 'Part Time']
    # Filter the dormant Archive Services office and Sales & Marketing
    hidden = {'archive services', 'sales and marketing', 'sales & marketing', 's & m', 's&m'}
    def office_of(r):
        n = r[1] if len(r) > 1 and r[1] else (r[0] if r else '')
        if isinstance(n, str) and n.lower().startswith('company'):
            n = n[7:] if len(n) > 7 else n
        return str(n).strip()

    # Compute YTD per row: sum of utilization values across all past_dates for utilization sections;
    # for headcount section show last week's value as "Latest"
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Weekly PMO Utilization</h2>')
    p.append(f'<p class="page-sub">Most recent {len(date_cols)} completed weeks of utilization by office, with full-time and part-time breakdowns.</p>')
    p.append('</div>')

    # YTD Actual / Target / Variance from cols DE/DF/DG (0-indexed 108/109/110)
    YTD_ACTUAL_COL = 108
    YTD_TARGET_COL = 109
    YTD_VAR_COL = 110

    for sec_idx, idxs in enumerate(sections[:3]):
        title = section_titles[sec_idx] if sec_idx < len(section_titles) else f'Section {sec_idx+1}'
        clean_idxs = [i for i in idxs if office_of(rows[i]).lower() not in hidden]
        if not clean_idxs: continue

        # Determine if pct (utilization) - based on sample values
        sample = [pn(rows[i][c]) for i in clean_idxs for c, _ in date_cols[:5]]
        sample = [s for s in sample if s is not None]
        is_pct = bool(sample) and sum(1 for s in sample if 0 <= s <= 1.5 and (s != int(s) or s == 0)) > len(sample) * 0.3

        p.append('<div class="card">')
        p.append(f'<div class="card-head"><h3>{esc(title)}</h3><span class="card-sub">{len(clean_idxs)} rows · last {len(date_cols)} weeks</span></div>')
        p.append('<div class="scroll-card-inner"><table class="data-table pmo-week-table"><thead><tr>')
        p.append('<th class="sticky-l">Office</th>')
        for c, dv in date_cols:
            p.append(f'<th class="num center">{fmt_date(dv, short=True)}</th>')
        if is_pct:
            p.append('<th class="num center">YTD Actual</th><th class="num center">Target</th><th class="num center">Variance</th>')
        else:
            p.append('<th class="num center">Latest</th>')
        p.append('</tr></thead><tbody>')
        for ri in clean_idxs:
            r = rows[ri]
            office = office_of(r)
            # Pull YTD Actual / Target / Variance from cols DE/DF/DG
            ytd_actual_raw = r[YTD_ACTUAL_COL] if len(r) > YTD_ACTUAL_COL else None
            ytd_target_raw = r[YTD_TARGET_COL] if len(r) > YTD_TARGET_COL else None
            ytd_var_raw = r[YTD_VAR_COL] if len(r) > YTD_VAR_COL else None
            ytd_avg = pn(ytd_actual_raw)
            target = pn(ytd_target_raw)
            variance = pn(ytd_var_raw)
            ol = office.lower()
            cls = ' class="total-row"' if 'total' in ol or ol == 'company' else ''
            p.append(f'<tr{cls}><td class="sticky-l">{esc(office)}</td>')
            for c, _ in date_cols:
                v = r[c] if c < len(r) else None
                n = pn(v)
                if n is None:
                    p.append('<td class="num center zero">·</td>')
                elif is_pct:
                    # colour each week against THIS office's target (the Target column two cells
                    # to the right), not a flat 80% -- a flat threshold made e.g. an office at 77%
                    # read red against its own 75.0% target.  Fixed 2026-09-14.
                    tcmp_wk = target if target is not None else 0.80
                    # compare what the cell SHOWS (whole %) against the shown target, so a cell
                    # reading 75% against a 75.0% target is never red on a hidden decimal
                    cls_p = 'good' if round(n * 100) >= round(tcmp_wk * 100, 1) else 'bad'
                    # data-val = the unrounded figure for the Excel export (FP&A, 2026-09-14:
                    # tenths in the workbook); the on-screen pill stays a whole percent.
                    p.append(f'<td class="cell center" data-val="{n:.6f}" data-fmt="0.0%">'
                             f'<span class="pill {cls_p}">{n*100:.0f}%</span></td>')
                elif n == 0:
                    p.append('<td class="num center zero">·</td>')
                else:
                    p.append(f'<td class="num center">{fmt_int(n) if n == int(n) else fmt_num(n, 1)}</td>')
            # YTD / Target / Variance columns (only for utilization sections) — sourced from DE/DF/DG
            if is_pct:
                y_n = ytd_avg
                t_n = target
                if y_n is None:
                    p.append('<td class="num center zero">·</td>')
                else:
                    tcmp = t_n if t_n is not None else 0.80
                    ycls = 'good' if y_n >= tcmp else 'bad'
                    p.append(f'<td class="num center" data-val="{y_n:.6f}" data-fmt="0.0%">'
                             f'<span class="pill {ycls}">{y_n*100:.1f}%</span></td>')
                if t_n is None:
                    p.append('<td class="num center zero">·</td>')
                else:
                    p.append(f'<td class="num center" data-val="{t_n:.6f}" data-fmt="0.0%">{t_n*100:.1f}%</td>')
                v_n = variance if variance is not None else ((y_n - t_n) if (y_n is not None and t_n is not None) else None)
                if v_n is None:
                    p.append('<td class="num center zero">·</td>')
                else:
                    var_cls = 'pos' if v_n >= 0 else 'neg'
                    var_sgn = '+' if v_n >= 0 else ''
                    p.append(f'<td class="num center"><span class="delta {var_cls}">{var_sgn}{v_n*100:.1f}pts</span></td>')
            else:
                latest = None
                for c, _ in reversed(past_dates):
                    n = pn(r[c] if c < len(r) else None)
                    if n is not None: latest = n; break
                p.append(f'<td class="num center total-cell">{fmt_int(latest) if latest is not None and latest == int(latest) else (fmt_num(latest, 1) if latest is not None else "")}</td>')
            p.append('</tr>')
        p.append('</tbody></table></div></div>')
    return ''.join(p)
def render_overtime_v2(data):
    """Overtime (round 2): 4 sections split by anchor columns B=1, U=20, AL=37, BC=54.
    Each section: header row at row 4 cell at the col, with date columns following.
    """
    rows = data['rows']
    if len(rows) < 7: return render_premium_generic(data)
    section_anchors = [
        ('Overtime Summary: Total', 1),
        ('T&M Projects', 20),
        ('FF Projects', 37),
        ('Overhead', 54),
    ]
    # Header row with dates is row 6 (index 5)
    hdr_row_idx = None
    for i, r in enumerate(rows[:10]):
        if sum(1 for v in r if is_date(v)) >= 3:
            hdr_row_idx = i; break
    if hdr_row_idx is None: return render_premium_generic(data)
    hdr = rows[hdr_row_idx]
    # For each anchor, find the date columns starting at that anchor and continuing for ~13 months
    sections = []
    for title, anchor in section_anchors:
        date_cols = []
        for c in range(anchor, min(anchor + 14, len(hdr))):
            if is_date(hdr[c]):
                date_cols.append((c, hdr[c]))
        if not date_cols: continue
        # Find Total column right after dates
        total_col = None
        for c in range(date_cols[-1][0] + 1, min(date_cols[-1][0] + 4, len(hdr))):
            if hdr[c] not in (None, '') and isinstance(hdr[c], str) and 'total' in str(hdr[c]).lower():
                total_col = c; break
        sections.append((title, anchor, date_cols, total_col))

    # Find data rows (office rows). Office name is in col 0.
    historic_idx = find_row_with(rows, 'historical', start=hdr_row_idx + 1)
    end_idx = historic_idx if historic_idx > 0 else len(rows)
    data_rows = []
    for i in range(hdr_row_idx + 1, end_idx):
        r = rows[i]
        if not r: continue
        office = r[0] if r else ''
        if not office or not isinstance(office, str): continue
        if office.lower().startswith('historical'): break
        if office.lower() == 'total': continue
        data_rows.append((office, r))

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Overtime Summary</h2>')
    p.append('<p class="page-sub ot-sub">Total overtime hours by office and month, split by project type. Pills color-coded by share of row total: green at zero, yellow between 0% and 1%, red above 1%.</p>')
    p.append('</div>')

    for sec_title, anchor, date_cols, total_col in sections:
        p.append('<div class="card">')
        p.append(f'<div class="card-head"><h3>{esc(sec_title)}</h3></div>')
        p.append('<div class="scroll-card-inner scroll-flat"><table class="data-table ot-table"><thead><tr>')
        p.append('<th class="sticky-l">Office</th>')
        for c, dv in date_cols:
            p.append(f'<th class="num center">{fmt_month(dv)}</th>')
        p.append('<th class="num right">Total</th>')
        p.append('</tr></thead><tbody>')
        for office, r in data_rows:
            row_total = 0
            cells = []
            for c, _ in date_cols:
                n = pn(r[c] if c < len(r) else None)
                if n is None: cells.append(None); continue
                row_total += n
                cells.append(n)
            p.append(f'<tr><td class="sticky-l">{esc(office)}</td>')
            for n in cells:
                if n is None:
                    p.append('<td class="num center zero">·</td>')
                elif n == 0:
                    p.append('<td class="cell center"><span class="pill good">0.0</span></td>')
                else:
                    share = n / row_total if row_total > 0 else 0
                    if share > 0.01: cls_p = 'bad'
                    elif share > 0: cls_p = 'warn'
                    else: cls_p = 'good'
                    p.append(f'<td class="cell center"><span class="pill {cls_p}">{n:.1f}</span></td>')
            # Total — 1-decimal, right-aligned
            tv = pn(r[total_col]) if total_col is not None and total_col < len(r) else row_total
            tv = tv if tv is not None else row_total
            p.append(f'<td class="num right total-cell">{tv:.1f}</td></tr>')
        p.append('</tbody></table></div></div>')
    return ''.join(p)


def render_goal_planning_v2(data):
    """Goal Planning (round 2): narrower employee name col (3/8 of current),
    centered numeric headers."""
    rows = data['rows']
    header_idx = -1
    for i, r in enumerate(rows[:8]):
        if r and any(isinstance(v, str) and 'employee name' in v.lower() for v in r):
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    headers = rows[header_idx]
    keep_idx = [0, 2, 5, 6, 7, 9, 10]
    col_specs = []
    for c in keep_idx:
        if c < len(headers) and headers[c] not in (None, ''):
            col_specs.append({'idx': c, 'name': str(headers[c]).replace('\n', ' ').strip()})
    target_col = None; ytd_col = None
    for cs in col_specs:
        nl = cs['name'].lower()
        if target_col is None and nl == 'target %': target_col = cs['idx']
        if ytd_col is None and 'ytd actual' in nl: ytd_col = cs['idx']
    offices = sorted({r[0] for r in rows[header_idx+1:] if r and isinstance(r[0], str) and r[0]})
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Goal Planning</h2>')
    p.append('<p class="page-sub">Year-end goals per employee, with additional hours needed and weeks remaining.</p>')
    p.append('</div>')
    p.append('<div class="toolbar">')
    p.append('<input type="text" id="goalSearch" class="input" placeholder="Search employee name…" />')
    p.append('<select id="goalOffice" class="input"><option value="">All offices</option>')
    for o in offices: p.append(f'<option value="{esc(o)}">{esc(o)}</option>')
    p.append('</select></div>')
    p.append('<div class="scroll-card"><table class="data-table goal-table"><thead><tr>')
    for cs in col_specs:
        nl = cs['name'].lower()
        is_num = any(k in nl for k in ('target', 'actual', 'hours', 'weeks', '%'))
        is_name = 'employee name' in nl
        cls_parts = []
        if is_num: cls_parts.append('num center')
        if is_name: cls_parts.append('col-name-narrow')
        if any(k in nl for k in ('additional hours', 'weeks remaining')):
            cls_parts.append('wrap-narrow')
        cls = ' '.join(cls_parts)
        p.append(f'<th class="{cls}">{esc(cs["name"])}</th>')
    p.append('</tr></thead><tbody>')
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r or not any(v not in (None, '') for v in r): continue
        office = r[0] if r else ''
        name_val = r[2] if len(r) > 2 else ''
        p.append(f'<tr data-office="{esc(office)}" data-name="{esc(str(name_val).lower())}">')
        for cs in col_specs:
            v = r[cs['idx']] if cs['idx'] < len(r) else None
            nl = cs['name'].lower()
            if cs['idx'] == ytd_col and target_col is not None:
                target_val = r[target_col] if target_col < len(r) else None
                p.append(f'<td class="num center">{pill_cmp(v, target_val)}</td>')
            elif 'target' in nl and '%' in cs['name']:
                p.append(f'<td class="num center">{fmt_pct(v, 0) if pn(v) is not None else ""}</td>')
            elif 'additional hours' in nl or 'weeks remaining' in nl:
                p.append(f'<td class="num center">{fmt_int(v) if pn(v) is not None else ""}</td>')
            elif is_date(v):
                p.append(f'<td class="date">{fmt_date(v)}</td>')
            elif 'employee name' in nl:
                p.append(f'<td class="col-name-narrow">{esc(str(v)) if v not in (None,"") else ""}</td>')
            else:
                p.append(f'<td>{esc(str(v)) if v not in (None,"") else ""}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('<script>(function(){var s=document.getElementById("goalSearch"),o=document.getElementById("goalOffice");if(!s) return;var rows=document.querySelectorAll(".goal-table tbody tr");function ap(){var q=s.value.toLowerCase().trim(),of=o.value;rows.forEach(function(r){var ok=true;if(q&&!(r.dataset.name||"").includes(q))ok=false;if(of&&r.dataset.office!==of)ok=false;r.classList.toggle("hidden",!ok);});}s.addEventListener("input",ap);o.addEventListener("change",ap);})();</script>')
    return ''.join(p)


def render_proposal_hours_v2(data):
    """Proposal Hours (round 2): only office-row coloring (no value coloring),
    remove row borders, format dates as 'Feb 28'."""
    rows = data['rows']
    header_idx = -1
    for i, r in enumerate(rows[:10]):
        if r and sum(1 for v in r if is_date(v)) >= 3:
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    headers = rows[header_idx]
    label_cols = []; date_cols = []; total_col = None
    for c, h in enumerate(headers):
        if is_date(h): date_cols.append((c, h))
        elif h not in (None, ''):
            h_str = str(h).replace('\n', ' ').strip()
            if 'total' in h_str.lower(): total_col = (c, h_str)
            else: label_cols.append((c, h_str))
    KNOWN_OFFICES = {'-','International (Global LLC)','Big Sky','Bluegrass','Built Environment','Cascade','Chesapeake','Coastal Bend',
                     'Corporate','Front Range','Geospatial','Great Lakes','Gulf Coast','High Desert','Hill Country','International',
                     'Island','Lakeshore','Maritime','Meridian Central','Meridian Southwest','Paleontology','Piedmont','Prairie',
                     'Red Rock','Sierra','Archive Services'}
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Proposal Hours</h2>')
    p.append('<p class="page-sub">Time charged to proposals, grouped by office.</p>')
    p.append('</div>')
    p.append('<div class="scroll-card scroll-flat"><table class="data-table prop-table no-borders"><thead><tr>')
    for c, name in label_cols:
        p.append(f'<th class="sticky-l">{esc(name)}</th>')
    for c, dv in date_cols:
        p.append(f'<th class="num">{fmt_date(dv, short=True)}</th>')
    if total_col: p.append(f'<th class="num">Total</th>')
    p.append('</tr></thead><tbody>')
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r: continue
        if not any(v not in (None, '') for v in r): continue
        first = r[0] if r else ''
        is_office = isinstance(first, str) and first in KNOWN_OFFICES
        cls = ' class="office-level"' if is_office else ''
        p.append(f'<tr{cls}>')
        for c, _ in label_cols:
            v = r[c] if c < len(r) else None
            p.append(f'<td class="sticky-l">{esc(str(v)) if v not in (None,"") else ""}</td>')
        for c, _ in date_cols:
            n = pn(r[c] if c < len(r) else None)
            if n is None or n == 0:
                p.append('<td class="num zero">·</td>')
            else:
                p.append(f'<td class="num">{fmt_num(n, 1)}</td>')
        if total_col:
            v = r[total_col[0]] if total_col[0] < len(r) else None
            p.append(f'<td class="num total-cell">{fmt_num(v, 1)}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def render_pivot_hours_v2(data):
    """Fleet Management (round 2): remove borders, all values 1-decimal."""
    rows = data['rows']
    header_idx = -1
    for i, r in enumerate(rows[:8]):
        if not r: continue
        date_count = sum(1 for v in r if is_date(v))
        if date_count >= 3:
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    headers = rows[header_idx]
    label_cols = []; date_cols = []; total_col = None
    for c, h in enumerate(headers):
        if is_date(h): date_cols.append((c, h))
        elif h not in (None, ''):
            h_str = str(h).replace('\n', ' ').strip()
            if 'total' in h_str.lower(): total_col = (c, h_str)
            else: label_cols.append((c, h_str))
    p = []
    p.append('<div class="page-intro">')
    p.append(f'<h2>{esc(data.get("tab", ""))}</h2>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="data-table no-borders"><thead><tr>')
    for c, name in label_cols:
        p.append(f'<th class="sticky-l">{esc(name)}</th>')
    for c, dv in date_cols:
        p.append(f'<th class="num">{fmt_month(dv)}</th>')
    if total_col: p.append(f'<th class="num">Total</th>')
    p.append('</tr></thead><tbody>')
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r: continue
        if not any(v not in (None, '') for v in r): continue
        first_val = r[0] if r else None
        is_section = (sum(1 for v in r if v not in (None, '')) == 1)
        if is_section and isinstance(first_val, str):
            total_w = len(label_cols) + len(date_cols) + (1 if total_col else 0)
            p.append(f'<tr class="section-row"><td colspan="{total_w}">{esc(first_val)}</td></tr>')
            continue
        p.append('<tr>')
        for c, _ in label_cols:
            v = r[c] if c < len(r) else None
            p.append(f'<td class="sticky-l">{esc(str(v)) if v not in (None,"") else ""}</td>')
        for c, _ in date_cols:
            n = pn(r[c] if c < len(r) else None)
            if n is None or n == 0:
                p.append('<td class="num zero">·</td>')
            else:
                # Always show 1 decimal
                p.append(f'<td class="num">{n:,.1f}</td>')
        if total_col:
            v = r[total_col[0]] if total_col[0] < len(r) else None
            n = pn(v)
            disp = f'{n:,.1f}' if n is not None else ''
            p.append(f'<td class="num total-cell">{disp}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


def render_pt_activity_v2(data):
    """Part-Time Employee Activity (round 2):
    - Weekly dates show as numbers with 1 decimal place (no pills/percentages)
    - Filter logic for 'No time in last X' includes 'never recorded'
    """
    rows = data['rows']
    if not rows: return '<p class="empty">No data.</p>'
    headers = rows[0]
    name_c = dept_c = pos_c = hire_c = lts_c = pw_c = w4_c = w8_c = w12_c = None
    date_cols = []
    for c, h in enumerate(headers):
        if h in (None, ''): continue
        if is_date(h):
            date_cols.append((c, h)); continue
        nl = str(h).replace('\n', ' ').strip().lower()
        if 'name' in nl and name_c is None: name_c = c
        elif 'department' in nl and dept_c is None: dept_c = c
        elif 'position' in nl and pos_c is None: pos_c = c
        elif 'hire' in nl and hire_c is None: hire_c = c
        elif 'last time' in nl and lts_c is None: lts_c = c
        elif nl == '4 weeks': w4_c = c
        elif nl == '8 weeks': w8_c = c
        elif nl == '12 weeks': w12_c = c
        elif nl == 'prior week': pw_c = c
    past_dates = [(c, v) for c, v in date_cols if to_date(v) and to_date(v) <= TODAY_REF]
    recent_weeks = past_dates[-8:] if past_dates else date_cols[-8:]
    body_end = len(rows)
    for i in range(1, len(rows)):
        for v in rows[i]:
            if isinstance(v, str) and 'utilization of available' in v.lower():
                body_end = i; break
        if body_end < len(rows): break
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Part-Time Employee Activity</h2>')
    p.append('<p class="page-sub">Recent timesheet activity for part-time staff. Weekly hours and rollups for the last 4, 8, and 12 weeks plus prior week.</p>')
    p.append('</div>')
    p.append('<div class="toolbar">')
    p.append('<input type="text" id="ptSearch" class="input" placeholder="Search name, department, position…" />')
    p.append('<select id="ptHire" class="input"><option value="">Any hire date</option><option value="recent">Hired in last 12 months</option><option value="1to3">1–3 years ago</option><option value="3plus">3+ years ago</option></select>')
    p.append('<select id="ptLts" class="input"><option value="">Any last-submission</option><option value="6mo">No time in last 6 months (incl. never)</option><option value="1yr">No time in last 12 months (incl. never)</option><option value="never">No time recorded ever</option></select>')
    p.append('<button type="button" id="ptExport" class="btn-toggle" title="Download current table as a CSV file (opens in Excel)">Export to Excel</button>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="data-table pt-table"><thead><tr>')
    p.append('<th class="sticky-l pt-name">Name</th><th class="pt-dept">Department</th><th class="pt-pos">Position</th><th class="pt-lts">Last Submission</th>')
    for c, dv in recent_weeks:
        p.append(f'<th class="num pt-wk">{fmt_date(dv, short=True)}</th>')
    p.append('<th class="num pt-rollup pt-rollup-first">Prior<br>Wk</th><th class="num pt-rollup">4 Wks</th><th class="num pt-rollup">8 Wks</th><th class="num pt-rollup">12 Wks</th>')
    p.append('</tr></thead><tbody>')
    today_dt = TODAY_REF
    for i in range(1, body_end):
        r = rows[i]
        if not r or not any(v not in (None, '') for v in r): continue
        if name_c is None or not r[name_c]: continue
        name = str(r[name_c]) if name_c < len(r) else ''
        if ',' not in name: continue
        dept = r[dept_c] if dept_c is not None and dept_c < len(r) else ''
        pos = r[pos_c] if pos_c is not None and pos_c < len(r) else ''
        hire = r[hire_c] if hire_c is not None and hire_c < len(r) else None
        lts = r[lts_c] if lts_c is not None and lts_c < len(r) else None
        hire_cat = ''
        if is_date(hire):
            d = to_date(hire)
            if d:
                years = (today_dt - d).days / 365.25
                if years <= 1: hire_cat = 'recent'
                elif years <= 3: hire_cat = '1to3'
                else: hire_cat = '3plus'
        # lts_cat: include 'never' in 6mo and 1yr filters
        lts_cat = []
        if is_date(lts):
            d = to_date(lts)
            if d:
                days = (today_dt - d).days
                if days > 180: lts_cat.append('6mo')
                if days > 365: lts_cat.append('1yr')
        else:
            lts_cat.extend(['6mo', '1yr', 'never'])
        rtext = f'{name} {dept} {pos}'.lower()
        p.append(f'<tr data-search="{esc(rtext)}" data-hire="{hire_cat}" data-lts="{" ".join(lts_cat)}">')
        p.append(f'<td class="sticky-l pt-name">{esc(name)}</td>')
        p.append(f'<td class="pt-dept">{esc(str(dept)) if dept else ""}</td>')
        p.append(f'<td class="pt-pos">{esc(str(pos)) if pos else ""}</td>')
        p.append(f'<td class="date pt-lts">{fmt_date(lts) if is_date(lts) else ""}</td>')
        for c, _ in recent_weeks:
            n = pn(r[c] if c < len(r) else None)
            if n is None:
                p.append('<td class="num pt-wk zero">·</td>')
            else:
                if 0 <= n <= 2 and n != int(n):
                    p.append(f'<td class="num pt-wk">{n*100:.1f}</td>')
                else:
                    p.append(f'<td class="num pt-wk">{n:,.1f}</td>')
        for k, col in enumerate((pw_c, w4_c, w8_c, w12_c)):
            rcls = 'pt-rollup pt-rollup-first' if k == 0 else 'pt-rollup'
            if col is not None and col < len(r):
                n = pn(r[col])
                if n is None:
                    p.append(f'<td class="num zero {rcls}">·</td>')
                elif 0 <= n <= 2 and n != int(n):
                    p.append(f'<td class="num {rcls}">{n*100:.1f}%</td>')
                else:
                    p.append(f'<td class="num {rcls}">{n:,.1f}</td>')
            else:
                p.append(f'<td class="{rcls}"></td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('''<script>(function(){
var s=document.getElementById("ptSearch"),h=document.getElementById("ptHire"),l=document.getElementById("ptLts");
var ex=document.getElementById("ptExport");
if(!s) return;
var rows=document.querySelectorAll(".pt-table tbody tr");
function ap(){var q=s.value.toLowerCase().trim(),hv=h.value,lv=l.value;rows.forEach(function(r){var ok=true;if(q&&!(r.dataset.search||"").includes(q))ok=false;if(hv&&r.dataset.hire!==hv)ok=false;if(lv){var cats=(r.dataset.lts||"").split(" ");if(!cats.includes(lv)) ok=false;}r.classList.toggle("hidden",!ok);});}
s.addEventListener("input",ap);h.addEventListener("change",ap);l.addEventListener("change",ap);
function csvEscape(v){v=String(v==null?"":v).replace(/\\s+/g," ").trim();if(/[",\\n]/.test(v))return '"'+v.replace(/"/g,'""')+'"';return v;}
function exportCsv(){
  var tbl=document.querySelector(".pt-table");
  if(!tbl) return;
  var lines=[];
  var hdr=Array.prototype.map.call(tbl.querySelectorAll("thead th"),function(th){return csvEscape(th.textContent);}).join(",");
  lines.push(hdr);
  Array.prototype.forEach.call(tbl.querySelectorAll("tbody tr"),function(tr){
    if(tr.classList.contains("hidden")) return;
    var cells=Array.prototype.map.call(tr.querySelectorAll("td"),function(td){return csvEscape(td.textContent.replace(/·/g,""));});
    lines.push(cells.join(","));
  });
  var csv="﻿"+lines.join("\\r\\n");
  var blob=new Blob([csv],{type:"text/csv;charset=utf-8;"});
  var a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download="part_time_activity.csv";
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function(){URL.revokeObjectURL(a.href);},1000);
}
if(ex) ex.addEventListener("click",exportCsv);
})();</script>''')
    return ''.join(p)


def render_on_call_trend_v2(data):
    """Load PT_Hours.json data and render only the 'Utilization of Available on Call Trend' section."""
    # Always load from PT_Hours since data file is now On_Call_Trend.json (copy of PT_Hours.json)
    return render_on_call_trend(data)


def render_non_posting_v2(data):
    """Non-Posting v2: renders the hour-based block (cols 1-12 + col 14)
    AND the percentage-based block (col 17 onward) as pilled percentages."""
    rows = data['rows']
    if not rows: return '<p class="empty">No data.</p>'
    # The header row has TWO sets of date columns. Find them.
    header_idx = -1
    for i, r in enumerate(rows[:10]):
        if sum(1 for v in r if is_date(v)) >= 3:
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    header = rows[header_idx]
    # Hours section: cols 0..15  Percentage section: cols 16..end
    hours_dates = [(c, v) for c, v in enumerate(header) if is_date(v) and c < 16]
    pct_dates = [(c, v) for c, v in enumerate(header) if is_date(v) and c >= 16]
    # Total cols
    hours_total_col = None
    pct_total_col = None
    for c, v in enumerate(header):
        if isinstance(v, str) and 'total' in v.lower() and c < 16: hours_total_col = c
        if isinstance(v, str) and 'total' in v.lower() and c >= 16: pct_total_col = c

    # Collect data rows
    data_rows = []
    total_row = None
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r: continue
        office = r[0] if r else ''
        if not office or not isinstance(office, str): continue
        if office.lower() == 'total':
            total_row = r; break
        data_rows.append((office, r))

    # Filter date columns to those with at least one non-zero value in either section
    def used(dates):
        out = []
        for c, dv in dates:
            for office, r in data_rows:
                v = pn(r[c] if c < len(r) else None)
                if v not in (None, 0): out.append((c, dv)); break
        return out or dates
    used_hours = used(hours_dates)
    used_pct = used(pct_dates)

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Non-Posting Hours by Office</h2>')
    p.append('<p class="page-sub">Non-posting share — green at zero, yellow under 3%, red above 3%.</p>')
    p.append('</div>')

    # --- Percentage table (TOP) ---
    p.append('<div class="card">')
    p.append('<div class="card-head"><h3>Share of Total Hours</h3><span class="card-sub">Non-posting share — green at zero, yellow under 3%, red above 3%</span></div>')
    p.append('<div class="scroll-card-inner scroll-flat"><table class="data-table np-pct-table"><thead><tr>')
    p.append('<th class="sticky-l">Office</th>')
    for c, dv in used_pct:
        p.append(f'<th class="num center">{fmt_month(dv)}</th>')
    p.append('<th class="num center">% of total</th></tr></thead><tbody>')
    def pct_pill(v):
        if v is None: return '<td class="num center zero">·</td>'
        if v == 0: cls_p = 'good'
        elif v > 0.03: cls_p = 'bad'
        elif v > 0: cls_p = 'warn'
        else: cls_p = 'good'
        return f'<td class="cell center"><span class="pill {cls_p}">{v*100:.1f}%</span></td>'
    for office, r in data_rows:
        p.append(f'<tr><td class="sticky-l">{esc(office)}</td>')
        for c, _ in used_pct:
            raw = r[c] if c < len(r) else None
            v = pn(raw)
            if isinstance(raw, str) and raw.startswith('#'): v = None
            p.append(pct_pill(v))
        tot = pn(r[pct_total_col]) if pct_total_col is not None and pct_total_col < len(r) else None
        p.append(pct_pill(tot))
        p.append('</tr>')
    if total_row:
        p.append('<tr class="total-row"><td class="sticky-l">Total</td>')
        for c, _ in used_pct:
            v = pn(total_row[c] if c < len(total_row) else None)
            p.append(pct_pill(v))
        tot = pn(total_row[pct_total_col]) if pct_total_col is not None and pct_total_col < len(total_row) else None
        p.append(pct_pill(tot))
        p.append('</tr>')
    p.append('</tbody></table></div></div>')

    # --- Hours table (BOTTOM) ---
    p.append('<div class="card">')
    p.append('<div class="card-head"><h3>Hours</h3><span class="card-sub">Non-posting hours per month per office</span></div>')
    p.append('<div class="scroll-card-inner scroll-flat"><table class="data-table"><thead><tr>')
    p.append('<th class="sticky-l">Office</th>')
    for c, dv in used_hours:
        p.append(f'<th class="num center">{fmt_month(dv)}</th>')
    p.append('<th class="num center">Total hrs</th></tr></thead><tbody>')
    for office, r in data_rows:
        rt = pn(r[hours_total_col]) if hours_total_col is not None and hours_total_col < len(r) else None
        if rt is None:
            rt = sum((pn(r[c]) or 0) for c, _ in used_hours)
        p.append(f'<tr><td class="sticky-l">{esc(office)}</td>')
        for c, _ in used_hours:
            n = pn(r[c] if c < len(r) else None)
            if n is None or n == 0:
                p.append('<td class="num center zero">·</td>')
            else:
                p.append(f'<td class="num center">{int(round(n)):,}</td>')
        p.append(f'<td class="num center total-cell">{int(round(rt)):,}</td></tr>')
    if total_row:
        p.append('<tr class="total-row"><td class="sticky-l">Total</td>')
        gt = 0
        for c, _ in used_hours:
            n = pn(total_row[c] if c < len(total_row) else None)
            if n is not None:
                p.append(f'<td class="num center">{int(round(n)):,}</td>')
                gt += n
            else:
                p.append('<td></td>')
        tt = pn(total_row[hours_total_col]) if hours_total_col is not None else gt
        p.append(f'<td class="num center total-cell">{int(round(tt if tt is not None else gt)):,}</td></tr>')
    p.append('</tbody></table></div></div>')
    return ''.join(p)


# Update RENDERERS with new versions
RENDERERS['bubble_dashboard'] = render_bubble_dashboard_v3
RENDERERS['pmo_period'] = render_pmo_period_v3
RENDERERS['pmo_weekly'] = render_pmo_weekly_v3
RENDERERS['overtime'] = render_overtime_v2
RENDERERS['goal_planning'] = render_goal_planning_v2
RENDERERS['proposal_hours'] = render_proposal_hours_v2
RENDERERS['pivot_hours'] = render_pivot_hours_v2
RENDERERS['pt_activity'] = render_pt_activity_v2
RENDERERS['on_call_trend'] = render_on_call_trend_v2
RENDERERS['non_posting'] = render_non_posting_v2


def render_overhead_v2(data):
    """Overhead (Employee) — full-width view matching the screenshot:
    PMO | Employee Name | Position | Target | Target Variance |
    YTD 2026: Utilized Billable | Utilized Unbillable | Admin | PTO | S&M (hours)
    YTD 2026: same five (%) — % rows are pill-colored
    Target Variance is pill-colored (red below, green at/above target).
    """
    rows = data['rows']
    if not rows: return '<p class="empty">No data.</p>'
    # Find header row (col 0 == 'PMO', col 1 == 'Employee Name')
    header_idx = -1
    for i, r in enumerate(rows[:10]):
        if r and len(r) > 4 and isinstance(r[0], str) and r[0].strip() == 'PMO':
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    headers = rows[header_idx]
    # Column indexes per screenshot/data:
    PMO_C, NAME_C, POS_C, TGT_C, VAR_C = 0, 1, 2, 3, 4
    HRS_COLS = [5, 6, 7, 8, 9]   # Utilized Billable / Unbillable / Admin / PTO / S&M
    PCT_COLS = [11, 12, 13, 14, 15]
    CAT_NAMES = ['Utilized Billable', 'Utilized Unbillable', 'Admin', 'PTO', 'S&M']

    # Office list for filter (unique non-empty PMO codes — but here PMO is integer ID; use Position grouping)
    # We won't filter by office in this tab; we'll just provide a name search.

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Overhead — Employee Detail</h2>')
    p.append('<p class="page-sub">Per-employee year-to-date utilization breakdown for full-time staff only. Variance pill: red if below target, green if at/above target.</p>')
    p.append('</div>')
    p.append('<div class="toolbar">')
    p.append('<input type="text" id="ovhdSearch" class="input" placeholder="Search employee, position…" />')
    p.append('<label class="toggle"><input type="checkbox" id="ovhdBelow" /> <span>Below target only</span></label>')
    p.append('<label class="toggle"><span>Total hrs &gt;</span> <input type="number" id="ovhdTotalMin" class="input small" placeholder="0" min="0" step="1" style="width:72px" /></label>')
    p.append('<span class="legend" id="ovhdCount"></span>')
    p.append('</div>')

    p.append('<div class="scroll-card"><table class="data-table ovhd-table"><thead>')
    # Grouped header — thick dividers right of Target Variance and between Hours/% sections.
    p.append('<tr class="hdr-group">')
    p.append('<th class="sticky-l" rowspan="2">PMO</th>')
    p.append('<th rowspan="2">Employee Name</th>')
    p.append('<th rowspan="2">Position</th>')
    p.append('<th class="num center" rowspan="2">Target</th>')
    p.append('<th class="num center" rowspan="2">Target Δ</th>')
    p.append(f'<th class="center hours-section-start" colspan="{len(HRS_COLS)}">YTD 2026 · Hours</th>')
    p.append(f'<th class="center pct-section-start" colspan="{len(PCT_COLS)}">YTD 2026 · % of Total</th>')
    p.append('</tr><tr>')
    for j, n in enumerate(CAT_NAMES):
        bcls = ' hours-section-start' if j == 0 else ''
        p.append(f'<th class="num center wrap-narrow hours-col{bcls}">{esc(n)}</th>')
    for j, n in enumerate(CAT_NAMES):
        bcls = ' pct-section-start' if j == 0 else ''
        p.append(f'<th class="num center wrap-narrow{bcls}">{esc(n)}</th>')
    p.append('</tr></thead><tbody>')

    # Index of Utilized Billable column inside PCT_COLS — first entry
    BILLABLE_PCT_IDX = 0  # PCT_COLS[0] is Utilized Billable %

    def pct_pill_cat(v, is_billable=False, tgt=None):
        """% pill color by share — for Utilized Billable: compare against target (green if ≥);
        for other categories: green at 0, yellow under 1%, red above 1%."""
        if v is None: return '<td class="num center zero">·</td>'
        if is_billable:
            tcmp = tgt if tgt is not None else 0.80
            cls = 'good' if v >= tcmp else 'bad'
        else:
            if v == 0: cls = 'good'
            elif v > 0.01: cls = 'bad'
            elif v > 0: cls = 'warn'
            else: cls = 'good'
        return f'<td class="cell center"><span class="pill {cls}">{v*100:.0f}%</span></td>'

    def var_pill(v):
        if v is None: return '<td class="num center zero">·</td>'
        cls = 'bad' if v < 0 else 'good'
        sign = '+' if v >= 0 else '−'
        return f'<td class="cell center"><span class="pill {cls}">{sign}{abs(v)*100:.0f}%</span></td>'

    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r or not any(v not in (None, '') for v in r): continue
        nm = r[NAME_C] if NAME_C < len(r) else None
        if not nm or not isinstance(nm, str) or not nm.strip(): continue
        if ',' not in nm: continue  # employee names typically "Last, First"
        pmo = r[PMO_C] if PMO_C < len(r) else ''
        pos = r[POS_C] if POS_C < len(r) else ''
        tgt = pn(r[TGT_C] if TGT_C < len(r) else None)
        var = pn(r[VAR_C] if VAR_C < len(r) else None)
        below = 1 if (var is not None and var < 0) else 0
        rtext = f'{nm} {pos} {pmo}'.lower()
        _row_total_hours = 0.0
        for _c in HRS_COLS:
            _vv = pn(r[_c] if _c < len(r) else None)
            if _vv is not None: _row_total_hours += _vv
        p.append(f'<tr data-search="{esc(rtext)}" data-below="{below}" data-total="{_row_total_hours:.2f}">')
        p.append(f'<td class="sticky-l">{esc(str(pmo)) if pmo not in (None,"") else ""}</td>')
        p.append(f'<td>{esc(nm)}</td>')
        p.append(f'<td>{esc(str(pos)) if pos else ""}</td>')
        p.append(f'<td class="num center">{fmt_pct(tgt, 0) if tgt is not None else ""}</td>')
        p.append(var_pill(var))
        # Hours — rounded whole numbers, right-justified, equal-width columns; first cell gets thick left border
        for j, c in enumerate(HRS_COLS):
            v = pn(r[c] if c < len(r) else None)
            bcls = ' hours-section-start' if j == 0 else ''
            if v is None or v == 0:
                p.append(f'<td class="num right zero hours-col{bcls}">·</td>')
            else:
                p.append(f'<td class="num right hours-col{bcls}">{int(round(v)):,}</td>')
        # Percentages — thick border at the first % cell
        for j, c in enumerate(PCT_COLS):
            raw = r[c] if c < len(r) else None
            v = pn(raw)
            if isinstance(raw, str) and raw.startswith('#'): v = None
            cell = pct_pill_cat(v, is_billable=(j == BILLABLE_PCT_IDX), tgt=tgt)
            if j == 0:
                cell = cell.replace('class="cell center"', 'class="cell center pct-section-start"').replace('class="num center zero"', 'class="num center zero pct-section-start"')
            p.append(cell)
        p.append('</tr>')

    p.append('</tbody></table></div>')
    p.append('''<script>(function(){
var s=document.getElementById("ovhdSearch"),b=document.getElementById("ovhdBelow"),tm=document.getElementById("ovhdTotalMin"),c=document.getElementById("ovhdCount");
if(!s) return;
var rs=document.querySelectorAll(".ovhd-table tbody tr");
function ap(){
  var q=(s.value||"").toLowerCase().trim(),bo=b&&b.checked,n=0;
  var mn = (tm && tm.value!=="") ? parseFloat(tm.value) : null;
  rs.forEach(function(r){
    var ok=true;
    if(q && !(r.dataset.search||"").includes(q)) ok=false;
    if(bo && r.dataset.below!=="1") ok=false;
    if(mn!==null && !isNaN(mn) && parseFloat(r.dataset.total||"0") <= mn) ok=false;
    r.classList.toggle("hidden",!ok);
    if(ok) n++;
  });
  if(c) c.textContent=n+" employees";
}
s.addEventListener("input",ap);
if(b) b.addEventListener("change",ap);
if(tm) tm.addEventListener("input",ap);
ap();
})();</script>''')
    return ''.join(p)


RENDERERS['overhead'] = render_overhead_v2


def render_pto_v2(data):
    """PTO v2: hide ID column, add filters for not-meeting-target and PTO hours threshold."""
    rows = data['rows']
    header_idx = -1
    for i, r in enumerate(rows[:8]):
        if r and any(isinstance(v, str) and 'employee name' in v.lower() for v in r):
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    headers = rows[header_idx]
    col_specs = []
    for c, h in enumerate(headers):
        if c == 0: continue
        if h in (None, ''): continue
        col_specs.append({'idx': c, 'name': str(h).replace('\n', ' ').strip()})
    target_col = None; ytd_col = None; var_col = None
    short_pto_col = None; ext_pto_col = None
    for cs in col_specs:
        nl = cs['name'].lower()
        if target_col is None and ('target' in nl and 'variance' not in nl): target_col = cs['idx']
        if ytd_col is None and ('ytd' in nl and 'actual' in nl): ytd_col = cs['idx']
        if var_col is None and 'variance' in nl: var_col = cs['idx']
        if short_pto_col is None and 'short' in nl and 'pto' in nl: short_pto_col = cs['idx']
        if ext_pto_col is None and 'extended' in nl and 'pto' in nl: ext_pto_col = cs['idx']

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Paid Time Off</h2>')
    p.append('<p class="page-sub">Short-term and extended PTO balances per employee, with utilization vs target.</p>')
    p.append('</div>')
    p.append('<div class="toolbar pto-toolbar">')
    p.append('<input type="text" id="ptoSearch" class="input" placeholder="Search employee, office…" />')
    p.append('<label class="toggle"><input type="checkbox" id="ptoBelow" /> <span>Below target only</span></label>')
    p.append('<label class="toggle"><span>PTO hrs &gt;</span> <input type="number" id="ptoMin" class="input small" placeholder="0" min="0" step="1" style="width:72px" /></label>')
    p.append('<span class="legend" id="ptoCount"></span>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="data-table pto-table"><thead><tr>')
    for cs in col_specs:
        nl = cs['name'].lower()
        wrap_cls = ' wrap-narrow' if any(k in nl for k in ('utilization', 'target', 'variance')) else ''
        is_num = any(k in nl for k in ('utiliz', 'target', 'variance', 'pto', 'hour'))
        p.append(f'<th class="{"num" if is_num else ""}{wrap_cls}">{esc(cs["name"])}</th>')
    p.append('</tr></thead><tbody>')
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r or not any(v not in (None, '') for v in r): continue
        first_real = r[1] if len(r) > 1 else None
        if isinstance(first_real, str) and first_real == '0': continue
        rtext = ' '.join(str(v) for v in r if v not in (None,'')).lower()
        # data attrs for filtering
        ytd_v = pn(r[ytd_col] if ytd_col is not None and ytd_col < len(r) else None)
        tgt_v = pn(r[target_col] if target_col is not None and target_col < len(r) else None)
        below = 1 if (ytd_v is not None and tgt_v is not None and ytd_v < tgt_v) else 0
        short_pto = pn(r[short_pto_col] if short_pto_col is not None and short_pto_col < len(r) else None) or 0
        ext_pto = pn(r[ext_pto_col] if ext_pto_col is not None and ext_pto_col < len(r) else None) or 0
        pto_total = short_pto + ext_pto
        p.append(f'<tr data-search="{esc(rtext)}" data-below="{below}" data-pto="{pto_total:.2f}">')
        for cs in col_specs:
            v = r[cs['idx']] if cs['idx'] < len(r) else None
            nl = cs['name'].lower()
            if cs['idx'] == ytd_col and target_col is not None:
                target_val = r[target_col] if target_col < len(r) else None
                p.append(f'<td class="num">{pill_cmp(v, target_val)}</td>')
            elif cs['idx'] == target_col:
                p.append(f'<td class="num">{fmt_pct(v, 0) if pn(v) is not None else esc(str(v))}</td>')
            elif cs['idx'] == var_col:
                n = pn(v)
                if n is None: disp = ''
                else: disp = f'<span class="delta {"pos" if n>=0 else "neg"}">{"+" if n>=0 else ""}{n*100:.0f}pts</span>'
                p.append(f'<td class="num">{disp}</td>')
            elif 'pto' in nl or 'hour' in nl:
                _pn = pn(v)
                if _pn is None:
                    p.append(f'<td class="num">{esc(str(v)) if v not in (None,"") else ""}</td>')
                else:
                    p.append(f'<td class="num">{_pn:,.1f}</td>')
            elif pn(v) is not None and 0 < abs(pn(v)) < 2 and pn(v) != int(pn(v)):
                p.append(f'<td class="num">{fmt_pct(v, 1)}</td>')
            elif pn(v) is not None:
                p.append(f'<td class="num">{fmt_num(v, 1)}</td>')
            elif is_date(v):
                p.append(f'<td class="date">{fmt_date(v)}</td>')
            else:
                p.append(f'<td>{esc(str(v)) if v not in (None,"") else ""}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('''<script>(function(){
var s=document.getElementById("ptoSearch");
var below=document.getElementById("ptoBelow");
var minIn=document.getElementById("ptoMin");
var count=document.getElementById("ptoCount");
if(!s) return;
var rows=document.querySelectorAll(".pto-table tbody tr");
function apply(){
  var q=(s.value||"").toLowerCase().trim();
  var bo=below && below.checked;
  var mn=minIn && minIn.value!==""?parseFloat(minIn.value):null;
  var n=0;
  rows.forEach(function(r){
    var ok=true;
    if(q && !(r.dataset.search||"").includes(q)) ok=false;
    if(bo && r.dataset.below!=="1") ok=false;
    if(mn!==null && !isNaN(mn) && parseFloat(r.dataset.pto||"0")<=mn) ok=false;
    r.classList.toggle("hidden",!ok);
    if(ok) n++;
  });
  if(count) count.textContent=n+" employees";
}
[s,below,minIn].forEach(function(el){if(el)el.addEventListener("input",apply);if(el && el.tagName==="INPUT" && el.type==="checkbox")el.addEventListener("change",apply);});
apply();
})();</script>''')
    return ''.join(p)


RENDERERS['pto'] = render_pto_v2


def render_proposal_hours_v3(data):
    """Proposal Hours v3: darker office row highlight, PMO filter, min hours filter."""
    rows = data['rows']
    header_idx = -1
    for i, r in enumerate(rows[:10]):
        if r and sum(1 for v in r if is_date(v)) >= 3:
            header_idx = i; break
    if header_idx == -1: return render_premium_generic(data)
    headers = rows[header_idx]
    label_cols = []; date_cols = []; total_col = None
    for c, h in enumerate(headers):
        if is_date(h): date_cols.append((c, h))
        elif h not in (None, ''):
            h_str = str(h).replace('\n', ' ').strip()
            if 'total' in h_str.lower(): total_col = (c, h_str)
            else: label_cols.append((c, h_str))
    KNOWN_OFFICES = {'-','International (Global LLC)','Big Sky','Bluegrass','Built Environment','Cascade','Chesapeake','Coastal Bend',
                     'Corporate','Front Range','Geospatial','Great Lakes','Gulf Coast','High Desert','Hill Country','International',
                     'Island','Lakeshore','Maritime','Meridian Central','Meridian Southwest','Paleontology','Piedmont','Prairie',
                     'Red Rock','Sierra','Archive Services','Sales & Marketing'}
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Proposal Hours</h2>')
    p.append('<p class="page-sub">Time charged to proposals, grouped by office.</p>')
    p.append('</div>')

    # Find offices that exist in data
    offices_present = []
    seen = set()
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r: continue
        first = r[0] if r else ''
        if isinstance(first, str) and first in KNOWN_OFFICES and first not in seen:
            offices_present.append(first); seen.add(first)
    offices_present.sort()

    p.append('<div class="toolbar">')
    p.append('<select id="propOffice" class="input"><option value="">All offices</option>')
    for o in offices_present:
        p.append(f'<option value="{esc(o)}">{esc(o)}</option>')
    p.append('</select>')
    p.append('<label class="toggle"><span>Hours &gt;</span> <input type="number" id="propMin" class="input small" placeholder="0" min="0" step="1" style="width:72px" /></label>')
    p.append('<span class="legend" id="propCount"></span>')
    p.append('</div>')

    p.append('<div class="scroll-card scroll-flat"><table class="data-table prop-table no-borders"><thead><tr>')
    for c, name in label_cols:
        p.append(f'<th class="sticky-l">{esc(name)}</th>')
    for c, dv in date_cols:
        p.append(f'<th class="num">{fmt_date(dv, short=True)}</th>')
    if total_col: p.append(f'<th class="num">Total</th>')
    p.append('</tr></thead><tbody>')

    current_office = ''
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r: continue
        if not any(v not in (None, '') for v in r): continue
        first = r[0] if r else ''
        is_office = isinstance(first, str) and first in KNOWN_OFFICES
        if is_office: current_office = first
        # Compute row hour total
        row_sum = 0
        for c, _ in date_cols:
            n = pn(r[c] if c < len(r) else None)
            if n is not None: row_sum += n
        if total_col:
            tv = pn(r[total_col[0]] if total_col[0] < len(r) else None)
            if tv is not None: row_sum = tv
        cls = ' office-level' if is_office else ''
        p.append(f'<tr class="prop-row{cls}" data-office="{esc(current_office)}" data-hours="{row_sum:.2f}" data-isoffice="{1 if is_office else 0}">')
        for c, _ in label_cols:
            v = r[c] if c < len(r) else None
            p.append(f'<td class="sticky-l">{esc(str(v)) if v not in (None,"") else ""}</td>')
        for c, _ in date_cols:
            n = pn(r[c] if c < len(r) else None)
            if n is None or n == 0:
                p.append('<td class="num zero">·</td>')
            else:
                p.append(f'<td class="num">{fmt_num(n, 1)}</td>')
        if total_col:
            v = r[total_col[0]] if total_col[0] < len(r) else None
            p.append(f'<td class="num total-cell">{fmt_num(v, 1)}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('''<script>(function(){
var off=document.getElementById("propOffice");
var minIn=document.getElementById("propMin");
var count=document.getElementById("propCount");
var rows=document.querySelectorAll(".prop-table tbody tr.prop-row");
function apply(){
  var selOff=off?off.value:"";
  var mn=minIn && minIn.value!==""?parseFloat(minIn.value):null;
  var n=0;
  rows.forEach(function(r){
    var ok=true;
    if(selOff && r.dataset.office!==selOff) ok=false;
    if(mn!==null && !isNaN(mn) && parseFloat(r.dataset.hours||"0")<=mn) ok=false;
    r.classList.toggle("hidden",!ok);
    if(ok) n++;
  });
  if(count) count.textContent=n+" rows";
}
if(off) off.addEventListener("change",apply);
if(minIn) minIn.addEventListener("input",apply);
apply();
})();</script>''')
    return ''.join(p)


RENDERERS['proposal_hours'] = render_proposal_hours_v3


def render_recovery(data):
    """Recovery Planning — per-employee path back to target with 8-week and 12-week scenarios.
    Header in row 5: Office | ID | Employee Name | Hire Date | % Weeks Available | YTD Actual % | Target % | [Column1] |
                     +10/+15/+20 % (8-Week) | [Column2] | +10/+15/+20 % (12-Week) ...
    """
    rows = data['rows']
    if len(rows) < 8: return render_premium_generic(data)
    # Find header row
    header_idx = -1
    for i, r in enumerate(rows[:10]):
        if r and any(isinstance(v, str) and 'employee name' in v.lower() for v in r):
            header_idx = i; break
    if header_idx < 0: return render_premium_generic(data)
    h = rows[header_idx]
    # Map columns
    OFFICE_C = 0; ID_C = 1; NAME_C = 2; HIRE_C = 3
    AVAIL_C = 4; YTD_C = 5; TGT_C = 6
    # Detect 8-week and 12-week +X% target columns by header label
    eight_cols = []
    twelve_cols = []
    for c, v in enumerate(h):
        if isinstance(v, str):
            lab = v.replace('\n', ' ').strip()
            if '8 Week' in lab and '+' in lab:
                eight_cols.append((c, lab))
            elif '12 Week' in lab and '+' in lab:
                twelve_cols.append((c, lab))
    # Limit to 3 each (assumed +10/+15/+20)
    eight_cols = eight_cols[:3]
    twelve_cols = twelve_cols[:3]

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Recovery Planning</h2>')
    p.append('<p class="page-sub">For each employee below target, the weekly utilization required over the next 8 and 12 weeks to recover by +10 / +15 / +20 percentage points above their current YTD.</p>')
    p.append('</div>')

    # Find office list and toolbar
    offices = []
    seen_off = set()
    employee_rows = []
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r: continue
        nm = r[NAME_C] if NAME_C < len(r) else None
        if not isinstance(nm, str) or ',' not in nm: continue
        off = r[OFFICE_C] if OFFICE_C < len(r) else ''
        if isinstance(off, str) and off.strip() and off not in seen_off:
            seen_off.add(off); offices.append(off)
        employee_rows.append((i, r))
    offices.sort()

    p.append('<div class="toolbar">')
    p.append('<input type="text" id="rcSearch" class="input" placeholder="Search employee, ID…" />')
    p.append('<select id="rcOffice" class="input"><option value="">All offices</option>')
    for o in offices:
        p.append(f'<option value="{esc(o)}">{esc(o)}</option>')
    p.append('</select>')
    p.append('<label class="toggle"><input type="checkbox" id="rcBelow" checked /> <span>Below target only</span></label>')
    p.append('<span id="rcCount" class="legend"></span>')
    p.append('</div>')

    p.append('<div class="scroll-card scroll-flat"><table class="data-table rc-table"><thead>')
    p.append('<tr class="hdr-group">')
    p.append('<th class="sticky-l" rowspan="2">Office</th>')
    p.append('<th rowspan="2">Employee</th>')
    p.append('<th class="num center" rowspan="2">Target</th>')
    p.append('<th class="num center" rowspan="2">YTD Actual</th>')
    p.append('<th class="num center" rowspan="2">Δ</th>')
    if eight_cols:
        p.append(f'<th class="center rc-sec-start" colspan="{len(eight_cols)}">8-Week Recovery</th>')
    if twelve_cols:
        p.append(f'<th class="center rc-sec-start" colspan="{len(twelve_cols)}">12-Week Recovery</th>')
    p.append('</tr><tr>')
    for j, (_, lab) in enumerate(eight_cols):
        # Extract the +N% label
        bcls = ' rc-sec-start' if j == 0 else ''
        short = lab.split('(')[0].strip()  # "+ 10% Target"
        p.append(f'<th class="num center wrap-narrow{bcls}">{esc(short)}</th>')
    for j, (_, lab) in enumerate(twelve_cols):
        bcls = ' rc-sec-start' if j == 0 else ''
        short = lab.split('(')[0].strip()
        p.append(f'<th class="num center wrap-narrow{bcls}">{esc(short)}</th>')
    p.append('</tr></thead><tbody>')

    def rec_pill(v, ytd, target):
        """Color the recovery cell: green if value <= 100% (achievable), yellow if 100-115%, red above 115%."""
        n = pn(v)
        if n is None: return '<td class="num center zero">·</td>'
        if n >= 1.0: cls = 'bad'
        elif n >= 0.85: cls = 'warn'
        else: cls = 'good'
        return f'<td class="cell center"><span class="pill {cls}">{n*100:.0f}%</span></td>'

    for orig_i, r in employee_rows:
        office = r[OFFICE_C] if OFFICE_C < len(r) else ''
        emp_id = r[ID_C] if ID_C < len(r) else ''
        nm = r[NAME_C] if NAME_C < len(r) else ''
        ytd = pn(r[YTD_C] if YTD_C < len(r) else None)
        tgt = pn(r[TGT_C] if TGT_C < len(r) else None)
        gap = (ytd - tgt) if (ytd is not None and tgt is not None) else None
        below = 1 if (gap is not None and gap < 0) else 0
        rtext = f'{nm} {emp_id} {office}'.lower()
        p.append(f'<tr class="rc-row" data-search="{esc(rtext)}" data-office="{esc(office)}" data-below="{below}">')
        p.append(f'<td class="sticky-l">{esc(office)}</td>')
        p.append(f'<td>{esc(nm)}<br><span class="emp-id">{esc(str(emp_id))}</span></td>')
        p.append(f'<td class="num center">{fmt_pct(tgt,0) if tgt is not None else ""}</td>')
        # YTD pill: green if at/above target, red if below
        if ytd is None:
            p.append('<td class="num center zero">·</td>')
        else:
            ycls = 'good' if (tgt is not None and ytd >= tgt) else 'bad'
            p.append(f'<td class="cell center"><span class="pill {ycls}">{ytd*100:.0f}%</span></td>')
        # Δ
        if gap is None:
            p.append('<td class="num center zero">·</td>')
        else:
            sgn = '+' if gap >= 0 else '−'
            dcls = 'pos' if gap >= 0 else 'neg'
            p.append(f'<td class="num center"><span class="delta {dcls}">{sgn}{abs(gap)*100:.1f}</span></td>')
        # Recovery columns
        for j, (c, _) in enumerate(eight_cols):
            cell = rec_pill(r[c] if c < len(r) else None, ytd, tgt)
            if j == 0:
                cell = cell.replace('class="cell center"', 'class="cell center rc-sec-start"').replace('class="num center zero"', 'class="num center zero rc-sec-start"')
            p.append(cell)
        for j, (c, _) in enumerate(twelve_cols):
            cell = rec_pill(r[c] if c < len(r) else None, ytd, tgt)
            if j == 0:
                cell = cell.replace('class="cell center"', 'class="cell center rc-sec-start"').replace('class="num center zero"', 'class="num center zero rc-sec-start"')
            p.append(cell)
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('''<script>(function(){
var s=document.getElementById("rcSearch"),o=document.getElementById("rcOffice"),b=document.getElementById("rcBelow"),c=document.getElementById("rcCount");
if(!s) return;
var rows=document.querySelectorAll(".rc-table tbody tr.rc-row");
function ap(){var q=(s.value||"").toLowerCase().trim(),of=o.value,bo=b&&b.checked,n=0;
rows.forEach(function(r){var ok=true;
if(q&&!(r.dataset.search||"").includes(q))ok=false;
if(of&&r.dataset.office!==of)ok=false;
if(bo&&r.dataset.below!=="1")ok=false;
r.classList.toggle("hidden",!ok); if(ok)n++;});
if(c) c.textContent=n+" employees";}
[s,o,b].forEach(function(el){if(el)el.addEventListener("input",ap);if(el && el.tagName==="SELECT")el.addEventListener("change",ap);if(el && el.type==="checkbox")el.addEventListener("change",ap);});
ap();
})();</script>''')
    return ''.join(p)


def render_recovery_v2(data):
    """Recovery Planning — v2.

    Changes vs v1:
    - 8-Week Recovery section hidden (per request, only 12-week shown).
    - New "Custom" recovery column sourced from source-sheet col S (idx 18).
      The custom-input parameters (Custom Weeks, Custom % Total) live in
      rows 1-2 col T of the source sheet and are surfaced in the page intro
      so the user knows the scenario being shown.
    - Removed "Below target only" toggle.
    - Pill coloring is binary red/green only (no yellow/warn).
    """
    rows = data['rows']
    if len(rows) < 8: return render_premium_generic(data)

    # Find header row
    header_idx = -1
    for i, r in enumerate(rows[:10]):
        if r and any(isinstance(v, str) and 'employee name' in v.lower() for v in r):
            header_idx = i; break
    if header_idx < 0: return render_premium_generic(data)
    h = rows[header_idx]

    OFFICE_C = 0; ID_C = 1; NAME_C = 2; HIRE_C = 3
    AVAIL_C = 4; YTD_C = 5; TGT_C = 6
    CUSTOM_C = 18

    # 12-week target columns (3 expected)
    twelve_cols = []
    for c, v in enumerate(h):
        if isinstance(v, str):
            lab = v.replace('\n', ' ').strip()
            if '12 Week' in lab and '+' in lab:
                twelve_cols.append((c, lab))
    twelve_cols = twelve_cols[:3]

    # Custom input parameters from rows above the header (row 1 / row 2, col T = idx 19)
    custom_weeks = None
    custom_pct = None
    if len(rows) > 1 and len(rows[1]) > 19:
        cw = pn(rows[1][19])
        if cw is not None: custom_weeks = cw
    if len(rows) > 2 and len(rows[2]) > 19:
        cp = pn(rows[2][19])
        if cp is not None: custom_pct = cp

    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>Recovery Planning</h2>')
    intro = ('For each employee, the weekly utilization required over the next 12 weeks '
             'to recover by +10 / +15 / +20 percentage points above their current YTD.')
    if custom_weeks is not None and custom_pct is not None:
        intro += (f' The <strong>Custom</strong> column shows the rate required to hit '
                  f'{custom_pct*100:.0f}% over {custom_weeks:.0f} weeks '
                  f'(configurable in cols S/T of the source sheet).')
    elif custom_weeks is not None or custom_pct is not None:
        intro += ' The <strong>Custom</strong> column reflects the parameters configured in cols S/T of the source sheet.'
    p.append(f'<p class="page-sub">{intro}</p>')
    p.append('</div>')

    # Toolbar (no Below-target toggle)
    offices = []
    seen_off = set()
    employee_rows = []
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if not r: continue
        nm = r[NAME_C] if NAME_C < len(r) else None
        if not isinstance(nm, str) or ',' not in nm: continue
        off = r[OFFICE_C] if OFFICE_C < len(r) else ''
        if isinstance(off, str) and off.strip() and off not in seen_off:
            seen_off.add(off); offices.append(off)
        employee_rows.append((i, r))
    offices.sort()

    p.append('<div class="toolbar">')
    p.append('<input type="text" id="rcSearch" class="input" placeholder="Search employee, ID…" />')
    p.append('<select id="rcOffice" class="input"><option value="">All offices</option>')
    for o in offices:
        p.append(f'<option value="{esc(o)}">{esc(o)}</option>')
    p.append('</select>')
    p.append('<span id="rcCount" class="legend"></span>')
    p.append('</div>')

    p.append('<div class="scroll-card scroll-flat"><table class="data-table rc-table"><thead>')
    p.append('<tr class="hdr-group">')
    p.append('<th class="sticky-l" rowspan="2">Office</th>')
    p.append('<th rowspan="2">Employee</th>')
    p.append('<th class="num center" rowspan="2">Target</th>')
    p.append('<th class="num center" rowspan="2">YTD Actual</th>')
    p.append('<th class="num center" rowspan="2">Δ</th>')
    if twelve_cols:
        p.append(f'<th class="center rc-sec-start" colspan="{len(twelve_cols)}">12-Week Recovery</th>')
    # Custom section header — show parameters if available
    cust_label = 'Custom'
    if custom_weeks is not None and custom_pct is not None:
        cust_label = f'Custom ({custom_pct*100:.0f}% / {custom_weeks:.0f} wks)'
    p.append(f'<th class="center rc-sec-start" colspan="1">{esc(cust_label)}</th>')
    p.append('</tr><tr>')
    for j, (_, lab) in enumerate(twelve_cols):
        bcls = ' rc-sec-start' if j == 0 else ''
        short = lab.split('(')[0].strip()
        p.append(f'<th class="num center wrap-narrow{bcls}">{esc(short)}</th>')
    p.append('<th class="num center wrap-narrow rc-sec-start">Required %</th>')
    p.append('</tr></thead><tbody>')

    def rec_pill(v):
        """Binary red/green pill — green when achievable (< 100%), red otherwise."""
        n = pn(v)
        if n is None: return '<td class="num center zero">·</td>'
        cls = 'bad' if n >= 1.0 else 'good'
        return f'<td class="cell center"><span class="pill {cls}">{n*100:.0f}%</span></td>'

    for orig_i, r in employee_rows:
        office = r[OFFICE_C] if OFFICE_C < len(r) else ''
        emp_id = r[ID_C] if ID_C < len(r) else ''
        nm = r[NAME_C] if NAME_C < len(r) else ''
        ytd = pn(r[YTD_C] if YTD_C < len(r) else None)
        tgt = pn(r[TGT_C] if TGT_C < len(r) else None)
        gap = (ytd - tgt) if (ytd is not None and tgt is not None) else None
        rtext = f'{nm} {emp_id} {office}'.lower()
        p.append(f'<tr class="rc-row" data-search="{esc(rtext)}" data-office="{esc(office)}">')
        p.append(f'<td class="sticky-l">{esc(office)}</td>')
        p.append(f'<td>{esc(nm)}<br><span class="emp-id">{esc(str(emp_id))}</span></td>')
        p.append(f'<td class="num center">{fmt_pct(tgt,0) if tgt is not None else ""}</td>')
        if ytd is None:
            p.append('<td class="num center zero">·</td>')
        else:
            ycls = 'good' if (tgt is not None and ytd >= tgt) else 'bad'
            p.append(f'<td class="cell center"><span class="pill {ycls}">{ytd*100:.0f}%</span></td>')
        if gap is None:
            p.append('<td class="num center zero">·</td>')
        else:
            sgn = '+' if gap >= 0 else '−'
            dcls = 'pos' if gap >= 0 else 'neg'
            p.append(f'<td class="num center"><span class="delta {dcls}">{sgn}{abs(gap)*100:.1f}</span></td>')
        for j, (c, _) in enumerate(twelve_cols):
            cell = rec_pill(r[c] if c < len(r) else None)
            if j == 0:
                cell = cell.replace('class="cell center"', 'class="cell center rc-sec-start"').replace('class="num center zero"', 'class="num center zero rc-sec-start"')
            p.append(cell)
        # Custom column (from col S = idx 18)
        cust_cell = rec_pill(r[CUSTOM_C] if CUSTOM_C < len(r) else None)
        cust_cell = cust_cell.replace('class="cell center"', 'class="cell center rc-sec-start"').replace('class="num center zero"', 'class="num center zero rc-sec-start"')
        p.append(cust_cell)
        p.append('</tr>')
    p.append('</tbody></table></div>')
    p.append('''<script>(function(){
var s=document.getElementById("rcSearch"),o=document.getElementById("rcOffice"),c=document.getElementById("rcCount");
if(!s) return;
var rows=document.querySelectorAll(".rc-table tbody tr.rc-row");
function ap(){var q=(s.value||"").toLowerCase().trim(),of=o.value,n=0;
rows.forEach(function(r){var ok=true;
if(q&&!(r.dataset.search||"").includes(q))ok=false;
if(of&&r.dataset.office!==of)ok=false;
r.classList.toggle("hidden",!ok); if(ok)n++;});
if(c) c.textContent=n+" employees";}
[s,o].forEach(function(el){if(el)el.addEventListener("input",ap);if(el && el.tagName==="SELECT")el.addEventListener("change",ap);});
ap();
})();</script>''')
    return ''.join(p)


RENDERERS['recovery'] = render_recovery_v2


# Sector → emoji icon (used by Employee Dashboard top projects card)
SECTOR_ICONS = {
    'federal government': '🏛️',
    'state/local government': '🏢',
    'oil and gas': '🛢️',
    'generation and transmission': '⚡',
    'renewable energy': '🌱',
    'mining': '⛏️',
    'transportation': '🚧',
    'utilities': '💡',
    'water': '💧',
    'cultural resources': '🪶',
    'private': '🏗️',
    'tribal': '🪶',
    'cell tower': '📡',
    'overhead': '🛠️',
}
def sector_icon(s):
    if not s or not isinstance(s, str): return '📁'
    return SECTOR_ICONS.get(s.lower().strip(), '📁')


JS_ED2_RENDER = """<script>(function(){
var raw = document.getElementById('ed2Data').textContent;
var data = JSON.parse(raw);
var emps = data.employees, dates = data.date_iso;
var picker = document.getElementById('ed2EmpPicker');
var startSel = document.createElement('select'); /* date range controls removed from UI; kept in-memory (full range) */
var endSel = document.createElement('select');

function fmt(d){if(!d) return '';var t=d.split('-');return t[1]+'/'+t[2];}
emps.forEach(function(e,i){
  var opt = document.createElement('option');
  opt.value = i;
  opt.textContent = e.name + (e.office ? '  '+e.office : '');
  picker.appendChild(opt);
});
dates.forEach(function(d,i){
  if(!d) return;
  var o1=document.createElement('option'); o1.value=i; o1.textContent=fmt(d);
  var o2=document.createElement('option'); o2.value=i; o2.textContent=fmt(d);
  startSel.appendChild(o1); endSel.appendChild(o2);
});
startSel.value = 0;
var _lastIdx = (typeof data.last_week_idx === 'number' && data.last_week_idx >= 0) ? data.last_week_idx : (dates.length-1);
endSel.value = _lastIdx;

function fmtH(n){if(n==null) return '-'; return (Math.round(n*10)/10).toLocaleString();}
function escHtml(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}

var COL_BILL='#3b6bc0', COL_UNBILL='#7c5fd0', COL_ADMIN='#6db5e8', COL_PTO='#4ec394', COL_SM='#f5a35e';
var SPARK_COLOR={util:'#4f7cc8', nonutil:'#7eb8d6', total:'#5fc28e'};

function smoothPath(coords){
  if(!coords || coords.length<2) return '';
  var pts = coords.map(function(c){var p=c.split(' ');return {x:parseFloat(p[0]),y:parseFloat(p[1])};});
  if(pts.length===2) return 'M '+pts[0].x.toFixed(1)+' '+pts[0].y.toFixed(1)+' L '+pts[1].x.toFixed(1)+' '+pts[1].y.toFixed(1);
  var t = 0.2;
  var d = 'M '+pts[0].x.toFixed(1)+' '+pts[0].y.toFixed(1);
  for(var i=0;i<pts.length-1;i++){
    var p0 = pts[Math.max(0,i-1)], p1 = pts[i], p2 = pts[i+1], p3 = pts[Math.min(pts.length-1,i+2)];
    var c1x = p1.x + (p2.x - p0.x)*t;
    var c1y = p1.y + (p2.y - p0.y)*t;
    var c2x = p2.x - (p3.x - p1.x)*t;
    var c2y = p2.y - (p3.y - p1.y)*t;
    d += ' C '+c1x.toFixed(1)+' '+c1y.toFixed(1)+' '+c2x.toFixed(1)+' '+c2y.toFixed(1)+' '+p2.x.toFixed(1)+' '+p2.y.toFixed(1);
  }
  return d;
}

function spark(svgEl, arr, color, lo, hi){
  var W=320, H=64, pad=4;
  var n = hi-lo+1; if(n<=0){svgEl.innerHTML=''; return;}
  var slice = arr.slice(lo, hi+1);
  var maxV = Math.max.apply(null, slice.concat([0.001]));
  var pts = [];
  for(var i=0;i<n;i++){
    var x = pad + (W-pad*2)*i/(n-1||1);
    var y = H - pad - (H-pad*2)*(slice[i]/maxV);
    pts.push(x.toFixed(1)+' '+y.toFixed(1));
  }
  var area = 'M '+pad+' '+(H-pad)+' L '+pts.join(' L ')+' L '+(W-pad)+' '+(H-pad)+' Z';
  var line = 'M '+pts.join(' L ');
  svgEl.innerHTML = '<path d=\"'+area+'\" fill=\"'+color+'\" opacity=\"0.18\"/>'
                  + '<path d=\"'+line+'\" stroke=\"'+color+'\" stroke-width=\"2\" fill=\"none\" stroke-linecap=\"round\" stroke-linejoin=\"round\"/>';
}

function donut(slices){
  var total = slices.reduce(function(a,s){return a+(s.v||0);},0);
  var svg = document.getElementById('ed2Donut');
  var R = 95, r=58, cx=120, cy=120;
  var g = ['<g transform=\"translate('+cx+','+cy+')\">'];
  if(total > 0){
    var ang0 = -Math.PI/2;
    slices.forEach(function(s){
      var pct = (s.v||0)/total;
      if(pct<=0) return;
      var ang1 = ang0 + pct*Math.PI*2;
      var x0=Math.cos(ang0)*R, y0=Math.sin(ang0)*R;
      var x1=Math.cos(ang1)*R, y1=Math.sin(ang1)*R;
      var ix0=Math.cos(ang0)*r, iy0=Math.sin(ang0)*r;
      var ix1=Math.cos(ang1)*r, iy1=Math.sin(ang1)*r;
      var big = pct>0.5?1:0;
      var d = 'M '+x0+' '+y0+' A '+R+' '+R+' 0 '+big+' 1 '+x1+' '+y1+' L '+ix1+' '+iy1+' A '+r+' '+r+' 0 '+big+' 0 '+ix0+' '+iy0+' Z';
      g.push('<path d=\"'+d+'\" fill=\"'+s.c+'\"/>');
      ang0 = ang1;
    });
    g.push('<text x=\"0\" y=\"-4\" text-anchor=\"middle\" font-size=\"26\" font-weight=\"700\" fill=\"#0f172a\">'+total.toFixed(0)+'</text>');
    g.push('<text x=\"0\" y=\"18\" text-anchor=\"middle\" font-size=\"10\" fill=\"#64748b\" letter-spacing=\"2\">HOURS</text>');
  } else {
    g.push('<circle cx=\"0\" cy=\"0\" r=\"'+R+'\" fill=\"#eef2ff\"/>');
    g.push('<text x=\"0\" y=\"6\" text-anchor=\"middle\" font-size=\"14\" fill=\"#94a3b8\">No data</text>');
  }
  g.push('</g>');
  svg.innerHTML = g.join('');
  var leg = slices.map(function(s){
    var pct = total>0?((s.v||0)/total*100).toFixed(0):0;
    return '<div class=\"ed2-leg\"><span class=\"ed2-leg-dot\" style=\"background:'+s.c+'\"></span><span class=\"ed2-leg-name\">'+escHtml(s.n)+'</span><span class=\"ed2-leg-pct\">'+pct+'%</span></div>';
  }).join('');
  document.getElementById('ed2DonutLegend').innerHTML = leg;
}

function trendChart(emp, lo, hi){
  var W=1200, H=240, M={t:24,r:60,b:36,l:60};
  var IW=W-M.l-M.r, IH=H-M.t-M.b;
  var n = hi-lo+1;
  if(n<=0){document.getElementById('ed2TrendChart').innerHTML=''; return;}
  var totals = [], maxV = 0;
  for(var i=lo;i<=hi;i++){
    var v = emp.total_wk[i]||0;
    totals.push(v); if(v>maxV) maxV=v;
  }
  if(maxV<10) maxV=10;
  maxV = Math.ceil(maxV/10)*10;
  function xc(i){ return M.l + IW*(i+0.5)/n; }
  function yP(p){ return M.t + IH*(1 - p/1.0); }
  var bw = (IW/n) * 0.6;
  var svg = ['<g>'];
  for(var gg=0;gg<=4;gg++){
    var y = M.t + IH*gg/4, v = maxV - maxV*gg/4;
    svg.push('<line x1=\"'+M.l+'\" y1=\"'+y.toFixed(1)+'\" x2=\"'+(M.l+IW)+'\" y2=\"'+y.toFixed(1)+'\" stroke=\"#dbe5f5\" stroke-width=\"0.5\" stroke-dasharray=\"3 3\"/>');
    svg.push('<text x=\"'+(M.l-8)+'\" y=\"'+y.toFixed(1)+'\" text-anchor=\"end\" dominant-baseline=\"middle\" font-size=\"10\" fill=\"#94a3b8\">'+v.toFixed(0)+'</text>');
  }
  for(var gg=0;gg<=4;gg++){
    var y=M.t+IH*gg/4, p=100-25*gg;
    svg.push('<text x=\"'+(M.l+IW+8)+'\" y=\"'+y.toFixed(1)+'\" text-anchor=\"start\" dominant-baseline=\"middle\" font-size=\"10\" fill=\"#94a3b8\">'+p+'%</text>');
  }
  for(var i=0;i<n;i+=4){
    var d = dates[lo+i] || '';
    svg.push('<text x=\"'+xc(i).toFixed(1)+'\" y=\"'+(M.t+IH+18)+'\" text-anchor=\"middle\" font-size=\"10\" fill=\"#64748b\">'+(d?d.slice(5):'')+'</text>');
  }
  for(var i=0;i<n;i++){
    var x = xc(i)-bw/2, y0 = M.t+IH;
    var segs = [
      {v:emp.w_bill[lo+i]||0,   c:COL_BILL},
      {v:emp.w_unbill[lo+i]||0, c:COL_UNBILL},
      {v:emp.w_admin[lo+i]||0,  c:COL_ADMIN},
      {v:emp.w_pto[lo+i]||0,    c:COL_PTO},
      {v:emp.w_sm[lo+i]||0,     c:COL_SM}
    ];
    segs.forEach(function(s){
      var h = (s.v/maxV)*IH;
      if(h>0){
        svg.push('<rect x=\"'+x.toFixed(1)+'\" y=\"'+(y0-h).toFixed(1)+'\" width=\"'+bw.toFixed(1)+'\" height=\"'+h.toFixed(1)+'\" fill=\"'+s.c+'\" opacity=\"0.92\"/>');
        y0 -= h;
      }
    });
  }
  // Green Utilized% line — sourced from per-week util % on Utilization (Employee) tab
  var pts = [], dots = [];
  var prevValid = false, curRun = [];
  var upw = emp.util_pct_wk || [];
  function flushRun(){
    if(curRun.length>1) svg.push('<path d=\"'+smoothPath(curRun)+'\" stroke=\"#0e9488\" stroke-width=\"2.5\" fill=\"none\" stroke-linecap=\"round\" stroke-linejoin=\"round\"/>');
    curRun = [];
  }
  for(var i=0;i<n;i++){
    var v = upw[lo+i];
    if(typeof v === 'number' && isFinite(v)){
      var coord = xc(i).toFixed(1)+' '+yP(v).toFixed(1);
      curRun.push(coord);
      dots.push(coord);
    } else {
      flushRun();
    }
  }
  flushRun();
  dots.forEach(function(coord){
    var p = coord.split(' ');
    svg.push('<circle cx=\"'+p[0]+'\" cy=\"'+p[1]+'\" r=\"3.5\" fill=\"#0e9488\" stroke=\"#fff\" stroke-width=\"1.5\"/>');
  });
  var _tgt = (typeof emp.target === 'number' && isFinite(emp.target) && emp.target > 0) ? emp.target : 0.8;
  var ty = yP(_tgt);
  svg.push('<line x1=\"'+M.l+'\" y1=\"'+ty.toFixed(1)+'\" x2=\"'+(M.l+IW)+'\" y2=\"'+ty.toFixed(1)+'\" stroke=\"#94a3b8\" stroke-width=\"1\" stroke-dasharray=\"6 4\"/>');
  svg.push('<text x=\"'+(M.l+IW-4)+'\" y=\"'+(ty-4).toFixed(1)+'\" text-anchor=\"end\" font-size=\"10\" fill=\"#94a3b8\" font-weight=\"600\">Target ('+(_tgt*100).toFixed(0)+'%)</text>');
  var lx=M.l+8, ly=14;
  [['Billable',COL_BILL],['Non-Posting',COL_UNBILL],['Admin',COL_ADMIN],['PTO',COL_PTO],['S&M',COL_SM]].forEach(function(l,i){
    var x=lx+i*100;
    svg.push('<rect x=\"'+x+'\" y=\"'+(ly-9)+'\" width=\"11\" height=\"9\" fill=\"'+l[1]+'\"/>');
    svg.push('<text x=\"'+(x+15)+'\" y=\"'+ly+'\" font-size=\"10\" fill=\"#0f172a\">'+l[0]+'</text>');
  });
  svg.push('<line x1=\"'+(lx+5*100)+'\" y1=\"'+(ly-4)+'\" x2=\"'+(lx+5*100+18)+'\" y2=\"'+(ly-4)+'\" stroke=\"#0e9488\" stroke-width=\"2.5\"/>');
  svg.push('<text x=\"'+(lx+5*100+22)+'\" y=\"'+ly+'\" font-size=\"10\" fill=\"#0f172a\">Utilized %</text>');
  svg.push('</g>');
  document.getElementById('ed2TrendChart').innerHTML = svg.join('');
}

function tasksList(emp){
  var maxH = 0;
  emp.tasks.forEach(function(t){ if(t.hours>maxH) maxH=t.hours; });
  if(maxH<1) maxH=1;
  var html = emp.tasks.slice(0,8).map(function(t){
    var pct = (t.hours/maxH*100).toFixed(0);
    var nm = t.name.length>60 ? t.name.slice(0,59)+'...' : t.name;
    return '<div class=\"ed2-task\"><div class=\"ed2-task-row\"><span class=\"ed2-task-name\" title=\"'+escHtml(t.name)+'\">'+escHtml(nm)+'</span><span class=\"ed2-task-hrs\">'+t.hours.toFixed(1)+' hrs</span></div><div class=\"ed2-task-bar-bg\"><div class=\"ed2-task-bar-fill\" style=\"width:'+pct+'%\"></div></div></div>';
  }).join('');
  document.getElementById('ed2TasksList').innerHTML = html || '<div class=\"ed2-empty\">No task data</div>';
}

function projects(emp){
  var icons = {'federal government':'🏛️','state/local government':'🏢','oil and gas':'🛢️','generation and transmission':'⚡','renewable energy':'🌱','mining':'⛏️','transportation':'🚧','utilities':'💡','water':'💧','cultural resources':'🪶','private':'🏗️','tribal':'🪶','cell tower':'📡','overhead':'🛠️'};
  var html = emp.projects.slice(0,6).map(function(pr){
    var sl = (pr.sector||'').toLowerCase().trim();
    var icon = icons[sl] || '📁';
    var nm = pr.name.length>70 ? pr.name.slice(0,69)+'...' : pr.name;
    return '<div class=\"ed2-proj\"><div class=\"ed2-proj-icon\">'+icon+'</div><div class=\"ed2-proj-info\"><div class=\"ed2-proj-name\" title=\"'+escHtml(pr.name)+'\">'+escHtml(nm)+'</div><div class=\"ed2-proj-sector\">'+escHtml(pr.sector||'-')+'</div></div></div>';
  }).join('');
  document.getElementById('ed2Projects').innerHTML = html || '<div class=\"ed2-empty\">No project data</div>';
}

function render(){
  var emp = emps[+picker.value];
  if(!emp) return;
  var lo = parseInt(startSel.value), hi = parseInt(endSel.value);
  if(hi<lo){ var t=lo; lo=hi; hi=t; }
  document.getElementById('ed2Title').textContent = emp.name + ': Utilization Analysis';
  document.getElementById('ed2KpiYtd').innerHTML = (emp.ytd_util==null?'-':(emp.ytd_util*100).toFixed(1)+'<span class=\"ed2-k-unit\">%</span>');
  document.getElementById('ed2KpiAdmin').innerHTML = fmtH(emp.admin);
  document.getElementById('ed2KpiPto').innerHTML = fmtH(emp.pto);
  document.getElementById('ed2KpiOt').innerHTML = fmtH(emp.ot);
  function rangeSum(arr){ var s=0; for(var i=lo;i<=hi;i++) s += arr[i]||0; return s; }
  document.getElementById('ed2Mini_util').textContent = fmtH(rangeSum(emp.util_wk));
  document.getElementById('ed2Mini_nonutil').textContent = fmtH(rangeSum(emp.nonutil_wk));
  document.getElementById('ed2Mini_total').textContent = fmtH(rangeSum(emp.total_wk));
  spark(document.getElementById('ed2Spark_util'), emp.util_wk, SPARK_COLOR.util, lo, hi);
  spark(document.getElementById('ed2Spark_nonutil'), emp.nonutil_wk, SPARK_COLOR.nonutil, lo, hi);
  spark(document.getElementById('ed2Spark_total'), emp.total_wk, SPARK_COLOR.total, lo, hi);
  donut([
    {n:'Billable',     v:rangeSum(emp.w_bill),   c:COL_BILL},
    {n:'Non-Posting', v:rangeSum(emp.w_unbill), c:COL_UNBILL},
    {n:'Admin',        v:rangeSum(emp.w_admin),  c:COL_ADMIN},
    {n:'PTO',          v:rangeSum(emp.w_pto),    c:COL_PTO},
    {n:'S&M',          v:rangeSum(emp.w_sm),     c:COL_SM}
  ]);
  trendChart(emp, lo, hi);
  tasksList(emp);
  projects(emp);
  var d1 = dates[lo]||'', d2 = dates[hi]||'';
  document.getElementById('ed2TrendRange').textContent = (d1 && d2) ? (d1.slice(5) + ' - ' + d2.slice(5)) : '';
}
[picker, startSel, endSel].forEach(function(el){ el.addEventListener('change', render); });
if(emps.length){ picker.value = 0; render(); }
})();</script>"""


JS_ED3_RENDER = """<script>(function(){
var raw = document.getElementById('ed3Data').textContent;
var data = JSON.parse(raw);
var emps = data.employees, dates = data.date_iso;
var picker = document.getElementById('ed3EmpPicker');
var startSel = document.createElement('select'); /* date range controls removed from UI; kept in-memory (full range) */
var endSel = document.createElement('select');

function fmt(d){if(!d) return '';var t=d.split('-');return t[1]+'/'+t[2];}
emps.forEach(function(e,i){
  var opt = document.createElement('option');
  opt.value = i;
  opt.textContent = e.name + (e.office ? '  '+e.office : '');
  picker.appendChild(opt);
});
dates.forEach(function(d,i){
  if(!d) return;
  var o1=document.createElement('option'); o1.value=i; o1.textContent=fmt(d);
  var o2=document.createElement('option'); o2.value=i; o2.textContent=fmt(d);
  startSel.appendChild(o1); endSel.appendChild(o2);
});
startSel.value = 0;
var _lastIdx = (typeof data.last_week_idx === 'number' && data.last_week_idx >= 0) ? data.last_week_idx : (dates.length-1);
endSel.value = _lastIdx;

function fmtH(n){if(n==null) return '-'; return (Math.round(n*10)/10).toLocaleString();}
function escHtml(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}

var COL_BILL='#3b6bc0', COL_UNBILL='#7c5fd0', COL_ADMIN='#6db5e8', COL_PTO='#4ec394', COL_SM='#f5a35e';
var SPARK_COLOR={util:'#4f7cc8', nonutil:'#7eb8d6', total:'#5fc28e'};

function smoothPath(coords){
  if(!coords || coords.length<2) return '';
  var pts = coords.map(function(c){var p=c.split(' ');return {x:parseFloat(p[0]),y:parseFloat(p[1])};});
  if(pts.length===2) return 'M '+pts[0].x.toFixed(1)+' '+pts[0].y.toFixed(1)+' L '+pts[1].x.toFixed(1)+' '+pts[1].y.toFixed(1);
  var t = 0.2;
  var d = 'M '+pts[0].x.toFixed(1)+' '+pts[0].y.toFixed(1);
  for(var i=0;i<pts.length-1;i++){
    var p0 = pts[Math.max(0,i-1)], p1 = pts[i], p2 = pts[i+1], p3 = pts[Math.min(pts.length-1,i+2)];
    var c1x = p1.x + (p2.x - p0.x)*t;
    var c1y = p1.y + (p2.y - p0.y)*t;
    var c2x = p2.x - (p3.x - p1.x)*t;
    var c2y = p2.y - (p3.y - p1.y)*t;
    d += ' C '+c1x.toFixed(1)+' '+c1y.toFixed(1)+' '+c2x.toFixed(1)+' '+c2y.toFixed(1)+' '+p2.x.toFixed(1)+' '+p2.y.toFixed(1);
  }
  return d;
}

function spark(svgEl, arr, color, lo, hi){
  var W=320, H=64, pad=4;
  var n = hi-lo+1; if(n<=0){svgEl.innerHTML=''; return;}
  var slice = arr.slice(lo, hi+1);
  var maxV = Math.max.apply(null, slice.concat([0.001]));
  var pts = [];
  for(var i=0;i<n;i++){
    var x = pad + (W-pad*2)*i/(n-1||1);
    var y = H - pad - (H-pad*2)*(slice[i]/maxV);
    pts.push(x.toFixed(1)+' '+y.toFixed(1));
  }
  var area = 'M '+pad+' '+(H-pad)+' L '+pts.join(' L ')+' L '+(W-pad)+' '+(H-pad)+' Z';
  var line = 'M '+pts.join(' L ');
  svgEl.innerHTML = '<path d=\"'+area+'\" fill=\"'+color+'\" opacity=\"0.18\"/>'
                  + '<path d=\"'+line+'\" stroke=\"'+color+'\" stroke-width=\"2\" fill=\"none\" stroke-linecap=\"round\" stroke-linejoin=\"round\"/>';
}

function donut(slices){
  var total = slices.reduce(function(a,s){return a+(s.v||0);},0);
  var svg = document.getElementById('ed3Donut');
  var R = 95, r=58, cx=120, cy=120;
  var g = ['<g transform=\"translate('+cx+','+cy+')\">'];
  if(total > 0){
    var ang0 = -Math.PI/2;
    slices.forEach(function(s){
      var pct = (s.v||0)/total;
      if(pct<=0) return;
      var ang1 = ang0 + pct*Math.PI*2;
      var x0=Math.cos(ang0)*R, y0=Math.sin(ang0)*R;
      var x1=Math.cos(ang1)*R, y1=Math.sin(ang1)*R;
      var ix0=Math.cos(ang0)*r, iy0=Math.sin(ang0)*r;
      var ix1=Math.cos(ang1)*r, iy1=Math.sin(ang1)*r;
      var big = pct>0.5?1:0;
      var d = 'M '+x0+' '+y0+' A '+R+' '+R+' 0 '+big+' 1 '+x1+' '+y1+' L '+ix1+' '+iy1+' A '+r+' '+r+' 0 '+big+' 0 '+ix0+' '+iy0+' Z';
      g.push('<path d=\"'+d+'\" fill=\"'+s.c+'\"/>');
      ang0 = ang1;
    });
    g.push('<text x=\"0\" y=\"-4\" text-anchor=\"middle\" font-size=\"26\" font-weight=\"700\" fill=\"#0f172a\">'+total.toFixed(0)+'</text>');
    g.push('<text x=\"0\" y=\"18\" text-anchor=\"middle\" font-size=\"10\" fill=\"#64748b\" letter-spacing=\"2\">HOURS</text>');
  } else {
    g.push('<circle cx=\"0\" cy=\"0\" r=\"'+R+'\" fill=\"#eef2ff\"/>');
    g.push('<text x=\"0\" y=\"6\" text-anchor=\"middle\" font-size=\"14\" fill=\"#94a3b8\">No data</text>');
  }
  g.push('</g>');
  svg.innerHTML = g.join('');
  var leg = slices.map(function(s){
    var pct = total>0?((s.v||0)/total*100).toFixed(0):0;
    return '<div class=\"ed3-leg\"><span class=\"ed3-leg-dot\" style=\"background:'+s.c+'\"></span><span class=\"ed3-leg-name\">'+escHtml(s.n)+'</span><span class=\"ed3-leg-pct\">'+pct+'%</span></div>';
  }).join('');
  document.getElementById('ed3DonutLegend').innerHTML = leg;
}

function trendChart(emp, lo, hi){
  var W=1200, H=240, M={t:24,r:60,b:36,l:60};
  var IW=W-M.l-M.r, IH=H-M.t-M.b;
  var n = hi-lo+1;
  if(n<=0){document.getElementById('ed3TrendChart').innerHTML=''; return;}
  var totals = [], maxV = 0;
  for(var i=lo;i<=hi;i++){
    var v = emp.total_wk[i]||0;
    totals.push(v); if(v>maxV) maxV=v;
  }
  if(maxV<10) maxV=10;
  maxV = Math.ceil(maxV/10)*10;
  function xc(i){ return M.l + IW*(i+0.5)/n; }
  function yP(p){ return M.t + IH*(1 - p/1.0); }
  var bw = (IW/n) * 0.6;
  var svg = ['<g>'];
  for(var gg=0;gg<=4;gg++){
    var y = M.t + IH*gg/4, v = maxV - maxV*gg/4;
    svg.push('<line x1=\"'+M.l+'\" y1=\"'+y.toFixed(1)+'\" x2=\"'+(M.l+IW)+'\" y2=\"'+y.toFixed(1)+'\" stroke=\"#dbe5f5\" stroke-width=\"0.5\" stroke-dasharray=\"3 3\"/>');
    svg.push('<text x=\"'+(M.l-8)+'\" y=\"'+y.toFixed(1)+'\" text-anchor=\"end\" dominant-baseline=\"middle\" font-size=\"10\" fill=\"#94a3b8\">'+v.toFixed(0)+'</text>');
  }
  for(var gg=0;gg<=4;gg++){
    var y=M.t+IH*gg/4, p=100-25*gg;
    svg.push('<text x=\"'+(M.l+IW+8)+'\" y=\"'+y.toFixed(1)+'\" text-anchor=\"start\" dominant-baseline=\"middle\" font-size=\"10\" fill=\"#94a3b8\">'+p+'%</text>');
  }
  for(var i=0;i<n;i+=4){
    var d = dates[lo+i] || '';
    svg.push('<text x=\"'+xc(i).toFixed(1)+'\" y=\"'+(M.t+IH+18)+'\" text-anchor=\"middle\" font-size=\"10\" fill=\"#64748b\">'+(d?d.slice(5):'')+'</text>');
  }
  for(var i=0;i<n;i++){
    var x = xc(i)-bw/2, y0 = M.t+IH;
    var segs = [
      {v:emp.w_bill[lo+i]||0,   c:COL_BILL},
      {v:emp.w_unbill[lo+i]||0, c:COL_UNBILL},
      {v:emp.w_admin[lo+i]||0,  c:COL_ADMIN},
      {v:emp.w_pto[lo+i]||0,    c:COL_PTO},
      {v:emp.w_sm[lo+i]||0,     c:COL_SM}
    ];
    segs.forEach(function(s){
      var h = (s.v/maxV)*IH;
      if(h>0){
        svg.push('<rect x=\"'+x.toFixed(1)+'\" y=\"'+(y0-h).toFixed(1)+'\" width=\"'+bw.toFixed(1)+'\" height=\"'+h.toFixed(1)+'\" fill=\"'+s.c+'\" opacity=\"0.92\"/>');
        y0 -= h;
      }
    });
  }
  // Green Utilized% line — sourced from util_pct_wk (PMO Weekly tab)
  var dots = [];
  var curRun = [];
  var upw = emp.util_pct_wk || [];
  function flushRun(){
    if(curRun.length>1) svg.push('<path d=\"'+smoothPath(curRun)+'\" stroke=\"#0e9488\" stroke-width=\"2.5\" fill=\"none\" stroke-linecap=\"round\" stroke-linejoin=\"round\"/>');
    curRun = [];
  }
  for(var i=0;i<n;i++){
    var v = upw[lo+i];
    if(typeof v === 'number' && isFinite(v)){
      var coord = xc(i).toFixed(1)+' '+yP(v).toFixed(1);
      curRun.push(coord);
      dots.push(coord);
    } else {
      flushRun();
    }
  }
  flushRun();
  dots.forEach(function(coord){
    var p = coord.split(' ');
    svg.push('<circle cx=\"'+p[0]+'\" cy=\"'+p[1]+'\" r=\"3.5\" fill=\"#0e9488\" stroke=\"#fff\" stroke-width=\"1.5\"/>');
  });
  if (typeof emp.target === 'number' && isFinite(emp.target) && emp.target > 0) {
    var _tgt = emp.target;
    var ty = yP(_tgt);
    svg.push('<line x1=\"'+M.l+'\" y1=\"'+ty.toFixed(1)+'\" x2=\"'+(M.l+IW)+'\" y2=\"'+ty.toFixed(1)+'\" stroke=\"#94a3b8\" stroke-width=\"1\" stroke-dasharray=\"6 4\"/>');
    svg.push('<text x=\"'+(M.l+IW-4)+'\" y=\"'+(ty-4).toFixed(1)+'\" text-anchor=\"end\" font-size=\"10\" fill=\"#94a3b8\" font-weight=\"600\">Target ('+(_tgt*100).toFixed(0)+'%)</text>');
  }
  var lx=M.l+8, ly=14;
  [['Billable',COL_BILL],['Non-Posting',COL_UNBILL],['Admin',COL_ADMIN],['PTO',COL_PTO],['S&M',COL_SM]].forEach(function(l,i){
    var x=lx+i*100;
    svg.push('<rect x=\"'+x+'\" y=\"'+(ly-9)+'\" width=\"11\" height=\"9\" fill=\"'+l[1]+'\"/>');
    svg.push('<text x=\"'+(x+15)+'\" y=\"'+ly+'\" font-size=\"10\" fill=\"#0f172a\">'+l[0]+'</text>');
  });
  svg.push('<line x1=\"'+(lx+5*100)+'\" y1=\"'+(ly-4)+'\" x2=\"'+(lx+5*100+18)+'\" y2=\"'+(ly-4)+'\" stroke=\"#0e9488\" stroke-width=\"2.5\"/>');
  svg.push('<text x=\"'+(lx+5*100+22)+'\" y=\"'+ly+'\" font-size=\"10\" fill=\"#0f172a\">Utilized %</text>');
  svg.push('</g>');
  document.getElementById('ed3TrendChart').innerHTML = svg.join('');
}

function tasksList(emp){
  var maxH = 0;
  emp.tasks.forEach(function(t){ if(t.hours>maxH) maxH=t.hours; });
  if(maxH<1) maxH=1;
  var html = emp.tasks.slice(0,8).map(function(t){
    var pct = (t.hours/maxH*100).toFixed(0);
    var nm = t.name.length>60 ? t.name.slice(0,59)+'...' : t.name;
    return '<div class=\"ed3-task\"><div class=\"ed3-task-row\"><span class=\"ed3-task-name\" title=\"'+escHtml(t.name)+'\">'+escHtml(nm)+'</span><span class=\"ed3-task-hrs\">'+t.hours.toFixed(1)+' hrs</span></div><div class=\"ed3-task-bar-bg\"><div class=\"ed3-task-bar-fill\" style=\"width:'+pct+'%\"></div></div></div>';
  }).join('');
  document.getElementById('ed3TasksList').innerHTML = html || '<div class=\"ed3-empty\">No task data</div>';
}

function projects(emp){
  var icons = {'federal government':'🏛️','state/local government':'🏢','oil and gas':'🛢️','generation and transmission':'⚡','renewable energy':'🌱','mining':'⛏️','transportation':'🚧','utilities':'💡','water':'💧','cultural resources':'🪶','private':'🏗️','tribal':'🪶','cell tower':'📡','overhead':'🛠️'};
  var html = emp.projects.slice(0,6).map(function(pr){
    var sl = (pr.sector||'').toLowerCase().trim();
    var icon = icons[sl] || '📁';
    var nm = pr.name.length>70 ? pr.name.slice(0,69)+'...' : pr.name;
    return '<div class=\"ed3-proj\"><div class=\"ed3-proj-icon\">'+icon+'</div><div class=\"ed3-proj-info\"><div class=\"ed3-proj-name\" title=\"'+escHtml(pr.name)+'\">'+escHtml(nm)+'</div><div class=\"ed3-proj-sector\">'+escHtml(pr.sector||'-')+'</div></div></div>';
  }).join('');
  document.getElementById('ed3Projects').innerHTML = html || '<div class=\"ed3-empty\">No project data</div>';
}

function render(){
  var emp = emps[+picker.value];
  if(!emp) return;
  var lo = parseInt(startSel.value), hi = parseInt(endSel.value);
  if(hi<lo){ var t=lo; lo=hi; hi=t; }
  document.getElementById('ed3Title').textContent = emp.name + ': Utilization Analysis...';
  document.getElementById('ed3KpiYtd').innerHTML = (emp.ytd_util==null?'-':(emp.ytd_util*100).toFixed(1)+'<span class=\"ed3-k-unit\">%</span>');
  document.getElementById('ed3KpiAdmin').innerHTML = fmtH(emp.admin);
  document.getElementById('ed3KpiPto').innerHTML = fmtH(emp.pto);
  document.getElementById('ed3KpiOt').innerHTML = fmtH(emp.ot);
  function rangeSum(arr){ var s=0; for(var i=lo;i<=hi;i++) s += arr[i]||0; return s; }
  document.getElementById('ed3Mini_util').textContent = fmtH(rangeSum(emp.util_wk));
  document.getElementById('ed3Mini_nonutil').textContent = fmtH(rangeSum(emp.nonutil_wk));
  document.getElementById('ed3Mini_total').textContent = fmtH(rangeSum(emp.total_wk));
  spark(document.getElementById('ed3Spark_util'), emp.util_wk, SPARK_COLOR.util, lo, hi);
  spark(document.getElementById('ed3Spark_nonutil'), emp.nonutil_wk, SPARK_COLOR.nonutil, lo, hi);
  spark(document.getElementById('ed3Spark_total'), emp.total_wk, SPARK_COLOR.total, lo, hi);
  donut([
    {n:'Billable',     v:rangeSum(emp.w_bill),   c:COL_BILL},
    {n:'Non-Posting', v:rangeSum(emp.w_unbill), c:COL_UNBILL},
    {n:'Admin',        v:rangeSum(emp.w_admin),  c:COL_ADMIN},
    {n:'PTO',          v:rangeSum(emp.w_pto),    c:COL_PTO},
    {n:'S&M',          v:rangeSum(emp.w_sm),     c:COL_SM}
  ]);
  trendChart(emp, lo, hi);
  tasksList(emp);
  projects(emp);
  var d1 = dates[lo]||'', d2 = dates[hi]||'';
  document.getElementById('ed3TrendRange').textContent = (d1 && d2) ? (d1.slice(5) + ' - ' + d2.slice(5)) : '';
}
[picker, startSel, endSel].forEach(function(el){ el.addEventListener('change', render); });
if(emps.length){ picker.value = 0; render(); }
})();</script>"""


def render_employee_dashboard(data):
    """Per-employee dashboard - visual layout modeled on the attached design."""
    rows = data['rows']
    if len(rows) < 4:
        return '<p class="empty">No data.</p>'

    row1 = rows[1]
    date_iso = []
    for c in range(1, 53):
        v = row1[c] if c < len(row1) else None
        if isinstance(v, dict) and '__date__' in v:
            date_iso.append(v['__date__'][:10])
        elif isinstance(v, str):
            date_iso.append(v)
        else:
            date_iso.append('')

    name_by_id = {}
    office_by_id = {}
    target_by_id = {}  # eid_str -> target_pct
    util_by_id = {}   # eid_str -> {date_iso: util_pct}
    date_headers_map = {}  # col_str -> date_iso
    try:
        with open('employees.json') as f:
            emp_master = json.load(f)
        date_headers_map = emp_master.get('date_headers', {}) or {}
        for e in emp_master.get('employees', []):
            eid = e.get('id')
            if eid is None: continue
            eid_s = str(int(eid)) if isinstance(eid, (int, float)) else str(eid)
            name_by_id[eid_s] = e.get('name') or ''
            office_by_id[eid_s] = EMPLOYEE_OFFICE_OVERRIDES.get(eid_s) or e.get('office') or ''
            tgt = e.get('target')
            target_by_id[eid_s] = tgt if isinstance(tgt, (int, float)) else None
            wk = e.get('weeks') or {}
            d2u = {}
            for col_str, pct in wk.items():
                d = date_headers_map.get(col_str)
                if d and isinstance(pct, (int, float)):
                    d2u[d] = pct
            util_by_id[eid_s] = d2u
    except Exception:
        pass

    employees = []
    for i in range(2, len(rows)):
        r = rows[i]
        if not r: continue
        eid = r[0] if r else None
        if eid in (None, ''): continue
        def pull(s, e):
            out = []
            for c in range(s, e+1):
                v = r[c] if c < len(r) else None
                n = pn(v)
                out.append(n if n is not None else 0.0)
            return out
        util_wk    = pull(1, 52)
        nonutil_wk = pull(53, 104)
        total_wk   = pull(105, 156)
        ytd_util = pn(r[157] if len(r) > 157 else None)
        admin = pn(r[158] if len(r) > 158 else None) or 0
        pto   = pn(r[159] if len(r) > 159 else None) or 0
        sm    = pn(r[160] if len(r) > 160 else None) or 0
        ot    = pn(r[161] if len(r) > 161 else None) or 0
        tasks = []
        for k in range(8):
            nm = r[162 + k*2] if 162+k*2 < len(r) else None
            hr = pn(r[163 + k*2] if 163+k*2 < len(r) else None)
            if nm and isinstance(nm, str) and nm.strip() and hr is not None and hr > 0:
                tasks.append({'name': nm.strip(), 'hours': hr})
        projects = []
        for k in range(6):
            nm = r[184 + k] if 184+k < len(r) else None
            sec = r[190 + k] if 190+k < len(r) else None
            if nm and isinstance(nm, str) and nm.strip():
                projects.append({'name': nm.strip(), 'sector': sec if isinstance(sec, str) else ''})
        billable   = pn(r[456] if len(r) > 456 else None) or 0
        unbillable = pn(r[457] if len(r) > 457 else None) or 0
        def pull_w(s, e):
            return [pn(r[c] if c < len(r) else None) or 0.0 for c in range(s, e+1)]
        w_bill   = pull_w(196, 247)
        w_unbill = pull_w(248, 299)
        w_admin  = pull_w(300, 351)
        w_pto    = pull_w(352, 403)
        w_sm     = pull_w(404, 455)
        eid_s = str(int(eid)) if isinstance(eid, (int, float)) else str(eid)
        # Per-week utilization % (aligned to dashboard's date_iso list); null where missing.
        emp_d2u = util_by_id.get(eid_s, {})
        util_pct_wk = [emp_d2u.get(d) for d in date_iso]
        employees.append({
            'id': eid_s,
            'name': name_by_id.get(eid_s, 'ID '+eid_s),
            'office': office_by_id.get(eid_s, ''),
            'target': target_by_id.get(eid_s),
            'ytd_util': ytd_util, 'admin': admin, 'pto': pto, 'sm': sm, 'ot': ot,
            'util_wk': util_wk, 'nonutil_wk': nonutil_wk, 'total_wk': total_wk,
            'tasks': tasks, 'projects': projects,
            'billable': billable, 'unbillable': unbillable,
            'w_bill': w_bill, 'w_unbill': w_unbill,
            'w_admin': w_admin, 'w_pto': w_pto, 'w_sm': w_sm,
            'util_pct_wk': util_pct_wk,
        })
    employees.sort(key=lambda e: e['name'].lower())

    # Last week index with any data, across all employees (utilized hrs > 0 OR util_pct present)
    last_week_idx = -1
    for w in range(52):
        any_data = False
        for emp in employees:
            uw = emp.get('util_wk') or []
            if w < len(uw) and isinstance(uw[w], (int, float)) and uw[w] > 0:
                any_data = True; break
            up = emp.get('util_pct_wk') or []
            if w < len(up) and isinstance(up[w], (int, float)):
                any_data = True; break
        if any_data:
            last_week_idx = w
    if last_week_idx < 0:
        last_week_idx = len(date_iso) - 1

    payload = json.dumps({'date_iso': date_iso, 'employees': employees, 'last_week_idx': last_week_idx})

    p = []
    p.append('<div class="ed2-app">')
    p.append('<div class="ed2-hero">')
    p.append('<div class="ed2-hero-title-wrap">')
    p.append('<div class="ed2-hero-icon">&#128202;</div>')
    p.append('<h2 class="ed2-hero-title" id="ed2Title">Employee: Utilization Analysis</h2>')
    p.append('</div>')
    p.append('<div class="ed2-hero-filters">')
    p.append('<label class="ed2-filter"><span>Employee</span><select id="ed2EmpPicker" class="ed2-select"></select></label>')
    p.append('</div></div>')

    p.append('<div class="ed2-main">')
    p.append('<div class="ed2-col-left">')
    for key, lbl in [('util','Utilized Hrs / Wk'),('nonutil','Non-Utilized Hrs / Wk'),('total','Total Hrs / Wk')]:
        p.append('<div class="ed2-mini" data-mini="'+key+'">')
        p.append('<div class="ed2-mini-head"><span class="ed2-mini-icon">&bull;</span><span class="ed2-mini-label">'+lbl+'</span><span class="ed2-mini-val" id="ed2Mini_'+key+'">&mdash;</span></div>')
        p.append('<svg class="ed2-mini-spark" id="ed2Spark_'+key+'" viewBox="0 0 320 64" preserveAspectRatio="none"></svg>')
        p.append('</div>')
    p.append('</div>')

    p.append('<div class="ed2-col-mid">')
    p.append('<div class="ed2-kpis">')
    p.append('<div class="ed2-kpi ed2-k-blue1"><div class="ed2-k-label">YTD<br/>Utilization</div><div class="ed2-k-val" id="ed2KpiYtd">&mdash;</div></div>')
    p.append('<div class="ed2-kpi ed2-k-blue2"><div class="ed2-k-label">Admin<br/>Hours</div><div class="ed2-k-val" id="ed2KpiAdmin">&mdash;</div></div>')
    p.append('<div class="ed2-kpi ed2-k-green"><div class="ed2-k-label">PTO<br/>Hours</div><div class="ed2-k-val" id="ed2KpiPto">&mdash;</div></div>')
    p.append('<div class="ed2-kpi ed2-k-orange"><div class="ed2-k-label">Overtime<br/>Hours</div><div class="ed2-k-val" id="ed2KpiOt">&mdash;</div></div>')
    p.append('</div>')

    p.append('<div class="ed2-mid-row">')
    p.append('<div class="ed2-panel ed2-allocation">')
    p.append('<div class="ed2-panel-head"><span class="ed2-panel-icon">&#128202;</span><span class="ed2-panel-title">Hours Allocation</span></div>')
    p.append('<div class="ed2-allocation-body">')
    p.append('<svg id="ed2Donut" viewBox="0 0 240 240" class="ed2-donut"></svg>')
    p.append('<div id="ed2DonutLegend" class="ed2-donut-legend"></div>')
    p.append('</div></div>')
    p.append('<div class="ed2-panel ed2-projects-panel">')
    p.append('<div class="ed2-panel-head"><span class="ed2-panel-icon">&#127919;</span><span class="ed2-panel-title">Top 6 Projects</span></div>')
    p.append('<div id="ed2Projects" class="ed2-projects-grid"></div>')
    p.append('</div>')
    p.append('</div>')
    p.append('</div>')

    p.append('<div class="ed2-col-right">')
    p.append('<div class="ed2-panel ed2-tasks">')
    p.append('<div class="ed2-panel-head"><span class="ed2-panel-icon">&#128203;</span><span class="ed2-panel-title">Top Tasks by Hours</span></div>')
    p.append('<div id="ed2TasksList" class="ed2-tasks-list"></div>')
    p.append('</div></div>')
    p.append('</div>')

    p.append('<div class="ed2-trend">')
    p.append('<div class="ed2-panel-head"><span class="ed2-panel-icon">&#128200;</span><span class="ed2-panel-title">Weekly Utilization Trend</span><span id="ed2TrendRange" class="ed2-trend-range"></span></div>')
    p.append('<svg id="ed2TrendChart" viewBox="0 0 1200 240" preserveAspectRatio="xMidYMid meet" class="ed2-trend-chart"></svg>')
    p.append('</div>')

    p.append('<div class="ed2-caveat">* Weekly aggregates may not exactly match YTD Utilization elsewhere on the site due to weekly FT/PT status and target weighting.</div>')
    p.append('</div>')

    p.append('<script id="ed2Data" type="application/json">' + payload + '</script>')
    p.append(JS_ED2_RENDER)
    return ''.join(p)


RENDERERS['employee_dashboard'] = render_employee_dashboard


def render_pmo_dashboard(data):
    """Per-PMO (office/region/discipline) dashboard.

    Mirrors the Employee Dashboard layout but reads from the PMO Dashboard sheet,
    where row[0] is the entity name (e.g. 'Big Sky') instead of an employee id.
    """
    rows = data['rows']
    if len(rows) < 4:
        return '<p class="empty">No data.</p>'

    row1 = rows[1]
    date_iso = []
    for c in range(1, 53):
        v = row1[c] if c < len(row1) else None
        if isinstance(v, dict) and '__date__' in v:
            date_iso.append(v['__date__'][:10])
        elif isinstance(v, str):
            date_iso.append(v)
        else:
            date_iso.append('')

    employees = []
    for i in range(2, len(rows)):
        r = rows[i]
        if not r: continue
        nm = r[0] if r else None
        if nm in (None, ''): continue
        def pull(s, e):
            out = []
            for c in range(s, e+1):
                v = r[c] if c < len(r) else None
                n = pn(v)
                out.append(n if n is not None else 0.0)
            return out
        util_wk    = pull(1, 52)
        nonutil_wk = pull(53, 104)
        total_wk   = pull(105, 156)
        ytd_util = pn(r[157] if len(r) > 157 else None)
        admin = pn(r[158] if len(r) > 158 else None) or 0
        pto   = pn(r[159] if len(r) > 159 else None) or 0
        sm    = pn(r[160] if len(r) > 160 else None) or 0
        ot    = pn(r[161] if len(r) > 161 else None) or 0
        tasks = []
        for k in range(8):
            tnm = r[162 + k*2] if 162+k*2 < len(r) else None
            thr = pn(r[163 + k*2] if 163+k*2 < len(r) else None)
            if tnm and isinstance(tnm, str) and tnm.strip() and thr is not None and thr > 0:
                tasks.append({'name': tnm.strip(), 'hours': thr})
        projects = []
        for k in range(6):
            pnm = r[184 + k] if 184+k < len(r) else None
            sec = r[190 + k] if 190+k < len(r) else None
            if pnm and isinstance(pnm, str) and pnm.strip():
                projects.append({'name': pnm.strip(), 'sector': sec if isinstance(sec, str) else ''})
        billable   = pn(r[456] if len(r) > 456 else None) or 0
        unbillable = pn(r[457] if len(r) > 457 else None) or 0
        def pull_w(s, e):
            return [pn(r[c] if c < len(r) else None) or 0.0 for c in range(s, e+1)]
        w_bill   = pull_w(196, 247)
        w_unbill = pull_w(248, 299)
        w_admin  = pull_w(300, 351)
        w_pto    = pull_w(352, 403)
        w_sm     = pull_w(404, 455)
        # Per-week util % computed locally from PMO-level util / total hours.
        util_pct_wk = []
        for j in range(52):
            t = total_wk[j] if j < len(total_wk) else 0
            u = util_wk[j] if j < len(util_wk) else 0
            util_pct_wk.append((u/t) if (isinstance(t, (int, float)) and t > 0) else None)
        employees.append({
            'id': str(nm),
            'name': str(nm),
            'office': '',
            'target': None,  # no per-PMO target wired yet
            'ytd_util': ytd_util, 'admin': admin, 'pto': pto, 'sm': sm, 'ot': ot,
            'util_wk': util_wk, 'nonutil_wk': nonutil_wk, 'total_wk': total_wk,
            'tasks': tasks, 'projects': projects,
            'billable': billable, 'unbillable': unbillable,
            'w_bill': w_bill, 'w_unbill': w_unbill,
            'w_admin': w_admin, 'w_pto': w_pto, 'w_sm': w_sm,
            'util_pct_wk': util_pct_wk,
        })
    employees.sort(key=lambda e: e['name'].lower())

    last_week_idx = -1
    for w in range(52):
        for emp in employees:
            uw = emp.get('util_wk') or []
            if w < len(uw) and isinstance(uw[w], (int, float)) and uw[w] > 0:
                last_week_idx = w; break
            up = emp.get('util_pct_wk') or []
            if w < len(up) and isinstance(up[w], (int, float)):
                last_week_idx = w; break
    if last_week_idx < 0:
        last_week_idx = len(date_iso) - 1

    payload = json.dumps({'date_iso': date_iso, 'employees': employees, 'last_week_idx': last_week_idx})

    p = []
    p.append('<div class="ed3-app">')
    p.append('<div class="ed3-hero">')
    p.append('<div class="ed3-hero-title-wrap">')
    p.append('<div class="ed3-hero-icon">&#128202;</div>')
    p.append('<h2 class="ed3-hero-title" id="ed3Title">PMO: Utilization Analysis...</h2>')
    p.append('</div>')
    p.append('<div class="ed3-hero-filters">')
    p.append('<label class="ed3-filter"><span>PMO</span><select id="ed3EmpPicker" class="ed3-select"></select></label>')
    p.append('</div></div>')

    p.append('<div class="ed3-main">')
    p.append('<div class="ed3-col-left">')
    for key, lbl in [('util','Utilized Hrs / Wk'),('nonutil','Non-Utilized Hrs / Wk'),('total','Total Hrs / Wk')]:
        p.append('<div class="ed3-mini" data-mini="'+key+'">')
        p.append('<div class="ed3-mini-head"><span class="ed3-mini-icon">&bull;</span><span class="ed3-mini-label">'+lbl+'</span><span class="ed3-mini-val" id="ed3Mini_'+key+'">&mdash;</span></div>')
        p.append('<svg class="ed3-mini-spark" id="ed3Spark_'+key+'" viewBox="0 0 320 64" preserveAspectRatio="none"></svg>')
        p.append('</div>')
    p.append('</div>')

    p.append('<div class="ed3-col-mid">')
    p.append('<div class="ed3-kpis">')
    p.append('<div class="ed3-kpi ed3-k-blue1"><div class="ed3-k-label">YTD<br/>Utilization</div><div class="ed3-k-val" id="ed3KpiYtd">&mdash;</div></div>')
    p.append('<div class="ed3-kpi ed3-k-blue2"><div class="ed3-k-label">Admin<br/>Hours</div><div class="ed3-k-val" id="ed3KpiAdmin">&mdash;</div></div>')
    p.append('<div class="ed3-kpi ed3-k-green"><div class="ed3-k-label">PTO<br/>Hours</div><div class="ed3-k-val" id="ed3KpiPto">&mdash;</div></div>')
    p.append('<div class="ed3-kpi ed3-k-orange"><div class="ed3-k-label">Overtime<br/>Hours</div><div class="ed3-k-val" id="ed3KpiOt">&mdash;</div></div>')
    p.append('</div>')

    p.append('<div class="ed3-mid-row">')
    p.append('<div class="ed3-panel ed3-allocation">')
    p.append('<div class="ed3-panel-head"><span class="ed3-panel-icon">&#128202;</span><span class="ed3-panel-title">Hours Allocation</span></div>')
    p.append('<div class="ed3-allocation-body">')
    p.append('<svg id="ed3Donut" viewBox="0 0 240 240" class="ed3-donut"></svg>')
    p.append('<div id="ed3DonutLegend" class="ed3-donut-legend"></div>')
    p.append('</div></div>')
    p.append('<div class="ed3-panel ed3-projects-panel">')
    p.append('<div class="ed3-panel-head"><span class="ed3-panel-icon">&#127919;</span><span class="ed3-panel-title">Top 6 Projects</span></div>')
    p.append('<div id="ed3Projects" class="ed3-projects-grid"></div>')
    p.append('</div>')
    p.append('</div>')
    p.append('</div>')

    p.append('<div class="ed3-col-right">')
    p.append('<div class="ed3-panel ed3-tasks">')
    p.append('<div class="ed3-panel-head"><span class="ed3-panel-icon">&#128203;</span><span class="ed3-panel-title">Top Tasks by Hours</span></div>')
    p.append('<div id="ed3TasksList" class="ed3-tasks-list"></div>')
    p.append('</div></div>')
    p.append('</div>')

    p.append('<div class="ed3-trend">')
    p.append('<div class="ed3-panel-head"><span class="ed3-panel-icon">&#128200;</span><span class="ed3-panel-title">Weekly Utilization Trend</span><span id="ed3TrendRange" class="ed3-trend-range"></span></div>')
    p.append('<svg id="ed3TrendChart" viewBox="0 0 1200 240" preserveAspectRatio="xMidYMid meet" class="ed3-trend-chart"></svg>')
    p.append('</div>')

    p.append('<div class="ed3-caveat">* Weekly utilization here = utilized hours / total hours for the PMO, computed from the PMO Dashboard sheet.</div>')
    p.append('</div>')

    p.append('<script id="ed3Data" type="application/json">' + payload + '</script>')
    p.append(JS_ED3_RENDER)
    return ''.join(p)


def render_pmo_dashboard_v2(data):
    """Per-PMO dashboard, v2.

    Same layout as v1, but the weekly utilization trend line is now sourced
    from the PMO (Weekly) tab (PMO_Weekly.json) instead of computed locally
    from the PMO Dashboard sheet. The PMO (Weekly) tab is the source of truth
    for office-level weekly utilization %, so the trend line matches the
    PMO Utilization: Weekly tab exactly.

    Data source for the trend: row whose col[1] matches the PMO name,
    columns 55-106 (52 weekly utilization % values). The headcount block
    sits at columns 2-53 of the same sheet, which we ignore here.
    """
    # Build {pmo_name_lower: [52 weekly util pct]} from PMO_Weekly.json
    pmo_weekly_pct = {}
    try:
        with open('site_data/PMO_Weekly.json') as _f:
            _wkly = json.load(_f)
        _wrows = _wkly.get('rows', [])
        # Utilization block columns: 55..106 inclusive (52 weeks)
        UTIL_START, UTIL_END = 55, 106
        # Walk rows 5..30 (data rows for offices; row 31 is Total — skip it).
        for _ri in range(5, min(31, len(_wrows))):
            _r = _wrows[_ri]
            if not _r or len(_r) < 2: continue
            _nm = _r[1]
            if not isinstance(_nm, str) or not _nm.strip() or _nm.strip().lower() == 'total':
                continue
            _series = []
            for _c in range(UTIL_START, UTIL_END + 1):
                _v = _r[_c] if _c < len(_r) else None
                _series.append(pn(_v))  # pn() returns None for 'N/A', empty, etc.
            pmo_weekly_pct[_nm.strip().lower()] = _series
    except (FileNotFoundError, json.JSONDecodeError, KeyError):
        pmo_weekly_pct = {}

    rows = data['rows']
    if len(rows) < 4:
        return '<p class="empty">No data.</p>'

    row1 = rows[1]
    date_iso = []
    for c in range(1, 53):
        v = row1[c] if c < len(row1) else None
        if isinstance(v, dict) and '__date__' in v:
            date_iso.append(v['__date__'][:10])
        elif isinstance(v, str):
            date_iso.append(v)
        else:
            date_iso.append('')

    employees = []
    for i in range(2, len(rows)):
        r = rows[i]
        if not r: continue
        nm = r[0] if r else None
        if nm in (None, ''): continue
        def pull(s, e):
            out = []
            for c in range(s, e+1):
                v = r[c] if c < len(r) else None
                n = pn(v)
                out.append(n if n is not None else 0.0)
            return out
        util_wk    = pull(1, 52)
        nonutil_wk = pull(53, 104)
        total_wk   = pull(105, 156)
        ytd_util = pn(r[157] if len(r) > 157 else None)
        admin = pn(r[158] if len(r) > 158 else None) or 0
        pto   = pn(r[159] if len(r) > 159 else None) or 0
        sm    = pn(r[160] if len(r) > 160 else None) or 0
        ot    = pn(r[161] if len(r) > 161 else None) or 0
        tasks = []
        for k in range(8):
            tnm = r[162 + k*2] if 162+k*2 < len(r) else None
            thr = pn(r[163 + k*2] if 163+k*2 < len(r) else None)
            if tnm and isinstance(tnm, str) and tnm.strip() and thr is not None and thr > 0:
                tasks.append({'name': tnm.strip(), 'hours': thr})
        projects = []
        for k in range(6):
            pnm = r[184 + k] if 184+k < len(r) else None
            sec = r[190 + k] if 190+k < len(r) else None
            if pnm and isinstance(pnm, str) and pnm.strip():
                projects.append({'name': pnm.strip(), 'sector': sec if isinstance(sec, str) else ''})
        billable   = pn(r[456] if len(r) > 456 else None) or 0
        unbillable = pn(r[457] if len(r) > 457 else None) or 0
        def pull_w(s, e):
            return [pn(r[c] if c < len(r) else None) or 0.0 for c in range(s, e+1)]
        w_bill   = pull_w(196, 247)
        w_unbill = pull_w(248, 299)
        w_admin  = pull_w(300, 351)
        w_pto    = pull_w(352, 403)
        w_sm     = pull_w(404, 455)

        # --- v2 change: util_pct_wk now comes from PMO Weekly tab when available ---
        key = str(nm).strip().lower()
        weekly_series = pmo_weekly_pct.get(key)
        if weekly_series is not None and len(weekly_series) >= 52:
            util_pct_wk = list(weekly_series[:52])
        else:
            # Fallback to the v1 local computation if the PMO is missing
            # from the PMO Weekly sheet (defensive — shouldn't happen for
            # standard offices).
            util_pct_wk = []
            for j in range(52):
                t = total_wk[j] if j < len(total_wk) else 0
                u = util_wk[j] if j < len(util_wk) else 0
                util_pct_wk.append((u/t) if (isinstance(t, (int, float)) and t > 0) else None)

        employees.append({
            'id': str(nm),
            'name': str(nm),
            'office': '',
            'target': None,
            'ytd_util': ytd_util, 'admin': admin, 'pto': pto, 'sm': sm, 'ot': ot,
            'util_wk': util_wk, 'nonutil_wk': nonutil_wk, 'total_wk': total_wk,
            'tasks': tasks, 'projects': projects,
            'billable': billable, 'unbillable': unbillable,
            'w_bill': w_bill, 'w_unbill': w_unbill,
            'w_admin': w_admin, 'w_pto': w_pto, 'w_sm': w_sm,
            'util_pct_wk': util_pct_wk,
        })
    employees.sort(key=lambda e: e['name'].lower())

    last_week_idx = -1
    for w in range(52):
        for emp in employees:
            uw = emp.get('util_wk') or []
            if w < len(uw) and isinstance(uw[w], (int, float)) and uw[w] > 0:
                last_week_idx = w; break
            up = emp.get('util_pct_wk') or []
            if w < len(up) and isinstance(up[w], (int, float)):
                last_week_idx = w; break
    if last_week_idx < 0:
        last_week_idx = len(date_iso) - 1

    payload = json.dumps({'date_iso': date_iso, 'employees': employees, 'last_week_idx': last_week_idx})

    p = []
    p.append('<div class="ed3-app">')
    p.append('<div class="ed3-hero">')
    p.append('<div class="ed3-hero-title-wrap">')
    p.append('<div class="ed3-hero-icon">&#128202;</div>')
    p.append('<h2 class="ed3-hero-title" id="ed3Title">PMO: Utilization Analysis...</h2>')
    p.append('</div>')
    p.append('<div class="ed3-hero-filters">')
    p.append('<label class="ed3-filter"><span>PMO</span><select id="ed3EmpPicker" class="ed3-select"></select></label>')
    p.append('</div></div>')

    p.append('<div class="ed3-main">')
    p.append('<div class="ed3-col-left">')
    for key, lbl in [('util','Utilized Hrs / Wk'),('nonutil','Non-Utilized Hrs / Wk'),('total','Total Hrs / Wk')]:
        p.append('<div class="ed3-mini" data-mini="'+key+'">')
        p.append('<div class="ed3-mini-head"><span class="ed3-mini-icon">&bull;</span><span class="ed3-mini-label">'+lbl+'</span><span class="ed3-mini-val" id="ed3Mini_'+key+'">&mdash;</span></div>')
        p.append('<svg class="ed3-mini-spark" id="ed3Spark_'+key+'" viewBox="0 0 320 64" preserveAspectRatio="none"></svg>')
        p.append('</div>')
    p.append('</div>')

    p.append('<div class="ed3-col-mid">')
    p.append('<div class="ed3-kpis">')
    p.append('<div class="ed3-kpi ed3-k-blue1"><div class="ed3-k-label">YTD<br/>Utilization</div><div class="ed3-k-val" id="ed3KpiYtd">&mdash;</div></div>')
    p.append('<div class="ed3-kpi ed3-k-blue2"><div class="ed3-k-label">Admin<br/>Hours</div><div class="ed3-k-val" id="ed3KpiAdmin">&mdash;</div></div>')
    p.append('<div class="ed3-kpi ed3-k-green"><div class="ed3-k-label">PTO<br/>Hours</div><div class="ed3-k-val" id="ed3KpiPto">&mdash;</div></div>')
    p.append('<div class="ed3-kpi ed3-k-orange"><div class="ed3-k-label">Overtime<br/>Hours</div><div class="ed3-k-val" id="ed3KpiOt">&mdash;</div></div>')
    p.append('</div>')

    p.append('<div class="ed3-mid-row">')
    p.append('<div class="ed3-panel ed3-allocation">')
    p.append('<div class="ed3-panel-head"><span class="ed3-panel-icon">&#128202;</span><span class="ed3-panel-title">Hours Allocation</span></div>')
    p.append('<div class="ed3-allocation-body">')
    p.append('<svg id="ed3Donut" viewBox="0 0 240 240" class="ed3-donut"></svg>')
    p.append('<div id="ed3DonutLegend" class="ed3-donut-legend"></div>')
    p.append('</div></div>')
    p.append('<div class="ed3-panel ed3-projects-panel">')
    p.append('<div class="ed3-panel-head"><span class="ed3-panel-icon">&#127919;</span><span class="ed3-panel-title">Top 6 Projects</span></div>')
    p.append('<div id="ed3Projects" class="ed3-projects-grid"></div>')
    p.append('</div>')
    p.append('</div>')
    p.append('</div>')

    p.append('<div class="ed3-col-right">')
    p.append('<div class="ed3-panel ed3-tasks">')
    p.append('<div class="ed3-panel-head"><span class="ed3-panel-icon">&#128203;</span><span class="ed3-panel-title">Top Tasks by Hours</span></div>')
    p.append('<div id="ed3TasksList" class="ed3-tasks-list"></div>')
    p.append('</div></div>')
    p.append('</div>')

    p.append('<div class="ed3-trend">')
    p.append('<div class="ed3-panel-head"><span class="ed3-panel-icon">&#128200;</span><span class="ed3-panel-title">Weekly Utilization Trend</span><span id="ed3TrendRange" class="ed3-trend-range"></span></div>')
    p.append('<svg id="ed3TrendChart" viewBox="0 0 1200 240" preserveAspectRatio="xMidYMid meet" class="ed3-trend-chart"></svg>')
    p.append('</div>')

    p.append('</div>')

    p.append('<div class="ed3-caveat">* Weekly utilization trend sourced from the PMO Utilization: Weekly tab (cols 55-106 of each PMO row).</div>')
    p.append('<script id="ed3Data" type="application/json">' + payload + '</script>')
    p.append(JS_ED3_RENDER)
    return ''.join(p)


RENDERERS['pmo_dashboard'] = render_pmo_dashboard_v2


def render_on_call_trend_v3(data):
    """Render the on-call section using PT_Hours row 1 (main headers) for date columns."""
    rows = data['rows']
    if not rows: return '<p class="empty">No data.</p>'
    headers = rows[0]
    # All date columns from main headers
    date_cols = [(c, v) for c, v in enumerate(headers) if is_date(v)]
    # Find section start
    sec_start = -1
    for i, r in enumerate(rows):
        for v in r:
            if isinstance(v, str) and 'utilization of available' in v.lower():
                sec_start = i; break
        if sec_start != -1: break
    if sec_start == -1 or not date_cols:
        return '<p class="empty">On Call Trend section not found.</p>'
    # Data rows: scan from sec_start onward, looking for rows with label in col 9 and date values
    data_rows = []
    for i in range(sec_start + 1, len(rows)):
        r = rows[i]
        if not r: continue
        # Get label - typically in col 9, sometimes col 0
        label = None
        for cand in (9, 0, 1):
            if cand < len(r) and isinstance(r[cand], str) and r[cand].strip() and not is_date(r[cand]):
                label = r[cand]; break
        if not label: continue
        # Verify row has numeric data in date cols
        has_data = False
        for c, _ in date_cols:
            if pn(r[c] if c < len(r) else None) is not None:
                has_data = True; break
        if has_data:
            data_rows.append((label.strip(), r))
    if not data_rows:
        return '<p class="empty">On Call Trend section has no data rows.</p>'
    # Show only weeks with data
    used_cols = []
    for c, dv in date_cols:
        for _, r in data_rows:
            if pn(r[c] if c < len(r) else None) is not None:
                used_cols.append((c, dv)); break
    if not used_cols: used_cols = date_cols
    # Recent weeks only (last 20 with data)
    used_cols = used_cols[-20:]
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>On Call Utilization Trend</h2>')
    p.append('<p class="page-sub">Weekly trend of part-time on-call staff utilized vs. available.</p>')
    p.append('</div>')
    p.append('<div class="scroll-card"><table class="data-table"><thead><tr>')
    p.append('<th class="sticky-l">Metric</th>')
    for c, dv in used_cols:
        p.append(f'<th class="num">{fmt_date(dv, short=True)}</th>')
    p.append('</tr></thead><tbody>')
    for label, r in data_rows:
        is_pct = ('%' in label) or ('/' in label) or ('vs.' in label.lower())
        p.append(f'<tr><td class="sticky-l">{esc(label)}</td>')
        for c, _ in used_cols:
            v = r[c] if c < len(r) else None
            n = pn(v)
            if n is None:
                p.append('<td class="num zero">·</td>')
            elif is_pct:
                p.append(f'<td class="num">{fmt_pct(n, 1)}</td>')
            else:
                p.append(f'<td class="num">{fmt_num(n, 1) if n != int(n) else fmt_int(n)}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


RENDERERS['on_call_trend'] = render_on_call_trend_v3


def _smooth(arr, window=3):
    """Centered moving average; preserves None for missing values at edges."""
    n = len(arr)
    out = [None] * n
    for i in range(n):
        lo = max(0, i - window // 2)
        hi = min(n, i + window // 2 + 1)
        chunk = [v for v in arr[lo:hi] if v is not None]
        if chunk: out[i] = sum(chunk) / len(chunk)
    return out


def render_on_call_trend_v4(data):
    """On Call Utilization Trend — combo chart with bars (Total Hours), smoothed utilization %,
    and smoothed PT-vs-Total % line. Includes a table below the chart."""
    rows = data['rows']
    if not rows: return '<p class="empty">No data.</p>'
    headers = rows[0]
    date_cols = [(c, v) for c, v in enumerate(headers) if is_date(v)]
    sec_start = -1
    for i, r in enumerate(rows):
        for v in r:
            if isinstance(v, str) and 'utilization of available' in v.lower():
                sec_start = i; break
        if sec_start != -1: break
    if sec_start == -1 or not date_cols:
        return '<p class="empty">On Call Trend section not found.</p>'

    def find_row(keyword):
        kw = keyword.lower()
        for i in range(sec_start + 1, len(rows)):
            r = rows[i]
            if not r: continue
            for c in (9, 0, 1):
                if c < len(r) and isinstance(r[c], str) and kw in r[c].lower():
                    return r
        return None

    tot_row = find_row('total hours')
    util_n_row = find_row('# of on call utilized')
    avail_n_row = find_row('# of on call available')
    util_pct_row = find_row('# utilized / available')
    pt_pct_row = find_row('% pt vs')

    # Restrict to weeks with data in the total hours row
    weeks = []
    for c, dv in date_cols:
        if tot_row and pn(tot_row[c] if c < len(tot_row) else None) is not None:
            weeks.append((c, dv))
    if not weeks: weeks = date_cols

    def series(r):
        if not r: return [None] * len(weeks)
        out = []
        for c, _ in weeks:
            out.append(pn(r[c] if c < len(r) else None))
        return out

    tot_series = series(tot_row)
    util_pct_series = series(util_pct_row)
    pt_pct_series = series(pt_pct_row)

    util_pct_smooth = _smooth(util_pct_series, window=5)
    pt_pct_smooth = _smooth(pt_pct_series, window=5)

    # Peak total hours week
    peak_i = None; peak_v = None
    for i, v in enumerate(tot_series):
        if v is not None and (peak_v is None or v > peak_v):
            peak_v = v; peak_i = i

    # Northstar Heritage brand palette
    C_EARTH = '#382e30'   # Earth Brown — Total Hours bars
    C_OCHRE = '#ad8c63'   # Ochre — smoothed util line
    C_CLAY  = '#734d45'   # Clay — accent dots
    C_OCEAN = '#61706e'   # Ocean — smoothed PT % line
    C_MOSS  = '#585444'   # Moss — gridlines
    C_GRAY  = '#d8d8d1'   # Warm Gray — light bars / borders

    # SVG chart — wider with tighter side margins; fits without scrolling
    W, H = 1200, 340
    M = {'top': 52, 'right': 46, 'bottom': 50, 'left': 44}
    IW = W - M['left'] - M['right']
    IH = H - M['top'] - M['bottom']
    n_w = len(weeks)
    # X positions: center of each week's slot
    def xc(i): return M['left'] + IW * (i + 0.5) / max(1, n_w)
    bar_w = max(6, (IW / max(1, n_w)) * 0.62)

    # Left axis: Total Hours
    valid_tot = [v for v in tot_series if v is not None]
    y_max_left = (max(valid_tot) // 500 + 1) * 500 if valid_tot else 3000
    y_min_left = 0
    def yL(v): return M['top'] + IH * (1 - (v - y_min_left) / (y_max_left - y_min_left))

    # Right axis: Percentages (0% to ~30%)
    pct_values = [v for v in (util_pct_smooth + pt_pct_smooth) if v is not None]
    y_max_right = max(0.3, (max(pct_values) * 1.2) if pct_values else 0.3)
    y_max_right = max(0.05, ((y_max_right * 100) // 5 + 1) * 0.05)  # snap to 5%
    def yR(v): return M['top'] + IH * (1 - v / y_max_right)

    s = [f'<svg viewBox="0 0 {W} {H}" class="oc-chart" preserveAspectRatio="xMidYMid meet">']
    # Title
    s.append(f'<text x="{W/2}" y="22" font-size="14" font-weight="700" fill="{C_EARTH}" text-anchor="middle">On-Call Utilization &amp; Staffing Efficiency</text>')

    # Left axis labels only (no gridlines)
    step_L = 500
    v = 0
    while v <= y_max_left + 0.001:
        y = yL(v)
        s.append(f'<text x="{M["left"]-8}" y="{y:.1f}" font-size="10" fill="{C_MOSS}" text-anchor="end" dominant-baseline="middle">{int(v):,}</text>')
        v += step_L

    # Right axis labels (no gridlines)
    step_R = 0.05
    v = 0
    while v <= y_max_right + 0.001:
        y = yR(v)
        s.append(f'<text x="{M["left"]+IW+8}" y="{y:.1f}" font-size="10" fill="{C_MOSS}" text-anchor="start" dominant-baseline="middle">{int(round(v*100))}%</text>')
        v += step_R

    # Axis titles — placed well above chart so they don't overlap top tick label
    s.append(f'<text x="{M["left"]-8}" y="{M["top"]-18}" font-size="10" font-weight="600" fill="{C_MOSS}" text-anchor="end">Hours</text>')
    s.append(f'<text x="{M["left"]+IW+8}" y="{M["top"]-18}" font-size="10" font-weight="600" fill="{C_MOSS}" text-anchor="start">%</text>')

    # Bars: Total Hours — Earth Brown
    for i, v in enumerate(tot_series):
        if v is None: continue
        bx = xc(i) - bar_w / 2
        by = yL(v)
        bh = (M['top'] + IH) - by
        s.append(f'<rect x="{bx:.1f}" y="{by:.1f}" width="{bar_w:.1f}" height="{max(0,bh):.1f}" fill="{C_EARTH}" opacity="0.85" rx="2" />')

    # Smoothed Utilized/Available % — Ochre solid line
    apts = [(xc(i), yR(v)) for i, v in enumerate(util_pct_smooth) if v is not None]
    if len(apts) >= 2:
        d_ = 'M ' + ' L '.join(f'{x:.1f} {y:.1f}' for x, y in apts)
        s.append(f'<path d="{d_}" stroke="{C_OCHRE}" stroke-width="2.5" fill="none" stroke-linecap="round" stroke-linejoin="round" />')
    # Raw util dots — Clay
    for i, v in enumerate(util_pct_series):
        if v is None: continue
        x = xc(i); y = yR(v)
        s.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="2.4" fill="{C_CLAY}" />')

    # Smoothed PT % — Ocean dashed line
    bpts = [(xc(i), yR(v)) for i, v in enumerate(pt_pct_smooth) if v is not None]
    if len(bpts) >= 2:
        d_ = 'M ' + ' L '.join(f'{x:.1f} {y:.1f}' for x, y in bpts)
        s.append(f'<path d="{d_}" stroke="{C_OCEAN}" stroke-width="2" fill="none" stroke-dasharray="5 4" stroke-linecap="round" />')

    # X axis labels (week dates), rotated
    for i, (c, dv) in enumerate(weeks):
        x = xc(i)
        d = to_date(dv)
        lbl = d.strftime('%b %-d') if d else ''
        s.append(f'<text x="{x:.1f}" y="{M["top"]+IH+16}" font-size="9" fill="{C_MOSS}" text-anchor="end" transform="rotate(-45, {x:.1f}, {M["top"]+IH+16})">{esc(lbl)}</text>')

    # Peak week annotation — text above the indicator line, no overlap
    if peak_i is not None and peak_v is not None:
        x = xc(peak_i); y = yL(peak_v)
        # Indicator line: short tick from above the bar tip
        s.append(f'<line x1="{x:.1f}" y1="{y-6:.1f}" x2="{x:.1f}" y2="{y-18:.1f}" stroke="{C_EARTH}" stroke-width="1.5" />')
        # Text stacked ABOVE the line (no overlap)
        s.append(f'<text x="{x:.1f}" y="{y-36:.1f}" font-size="11" font-weight="700" fill="{C_EARTH}" text-anchor="middle">Peak Week</text>')
        s.append(f'<text x="{x:.1f}" y="{y-22:.1f}" font-size="10" fill="{C_EARTH}" text-anchor="middle">{peak_v:,.0f} hrs</text>')

    # Legend
    lx = M['left'] + 12; ly = M['top'] - 12
    s.append(f'<rect x="{lx}" y="{ly-10}" width="14" height="10" fill="{C_EARTH}" opacity="0.85" rx="2" />')
    s.append(f'<text x="{lx+20}" y="{ly-1}" font-size="11" fill="{C_EARTH}">Total Hours (bars)</text>')
    s.append(f'<line x1="{lx+170}" y1="{ly-5}" x2="{lx+200}" y2="{ly-5}" stroke="{C_OCHRE}" stroke-width="2.5" />')
    s.append(f'<text x="{lx+206}" y="{ly-1}" font-size="11" fill="{C_EARTH}">Smoothed Utilized / Available %</text>')
    s.append(f'<line x1="{lx+440}" y1="{ly-5}" x2="{lx+470}" y2="{ly-5}" stroke="{C_OCEAN}" stroke-width="2" stroke-dasharray="5 4" />')
    s.append(f'<text x="{lx+476}" y="{ly-1}" font-size="11" fill="{C_EARTH}">Smoothed PT % of Total Hours</text>')

    s.append('</svg>')

    # Build content
    p = []
    p.append('<div class="page-intro">')
    p.append('<h2>On Call Utilization Trend</h2>')
    # period label follows the data (it was frozen at "Jan–May")
    _span = [to_date(dv) for _, dv in weeks if to_date(dv)]
    _span_lbl = (f'{_span[0].strftime("%b")}–{_span[-1].strftime("%b")}' if _span else 'YTD')
    p.append(f'<p class="page-sub">Smoothed weekly utilization trends with staffing volume ({_span_lbl}).</p>')
    p.append('</div>')
    p.append('<div class="oc-chart-card">')
    p.append(''.join(s))
    p.append('</div>')

    # Underneath: a compact data table
    used_cols = weeks
    p.append('<div class="scroll-card"><table class="data-table"><thead><tr>')
    p.append('<th class="sticky-l">Metric</th>')
    for c, dv in used_cols:
        p.append(f'<th class="num">{fmt_date(dv, short=True)}</th>')
    p.append('</tr></thead><tbody>')
    # (label, row, is_pct, force_int) — Total Hours is shown as whole numbers
    metric_rows = [
        ('Total Hours', tot_row, False, True),
        ('# On Call Utilized', util_n_row, False, True),
        ('# On Call Available', avail_n_row, False, True),
        ('Utilized / Available', util_pct_row, True, False),
        ('% PT vs. Total Hrs.', pt_pct_row, True, False),
    ]
    for label, r, is_pct, force_int in metric_rows:
        if r is None: continue
        p.append(f'<tr><td class="sticky-l">{esc(label)}</td>')
        for c, _ in used_cols:
            v = r[c] if c < len(r) else None
            n = pn(v)
            if n is None: p.append('<td class="num zero">·</td>')
            elif is_pct: p.append(f'<td class="num">{n*100:.1f}%</td>')
            elif force_int: p.append(f'<td class="num">{fmt_int(round(n))}</td>')
            else: p.append(f'<td class="num">{fmt_num(n, 1) if n != int(n) else fmt_int(n)}</td>')
        p.append('</tr>')
    p.append('</tbody></table></div>')
    return ''.join(p)


RENDERERS['on_call_trend'] = render_on_call_trend_v4
# Round 2 CSS additions and updated exec summary chart layout



def _build_weekly_company_trend():
    """Aggregate weekly company-wide utilization from employees.json.
    Returns a list of {week_label, date, actual} entries — one per week with data.
    """
    try:
        with open('employees.json') as f:
            d = json.load(f)
    except Exception:
        return []
    employees = d.get('employees', [])
    date_headers = d.get('date_headers', {})  # 1-indexed col -> 'YYYY-MM-DD'
    weekly_cols = d.get('weekly_cols', [])
    if not employees or not weekly_cols:
        return []
    out = []
    for col in weekly_cols:
        col_s = str(col)
        d_str = date_headers.get(col_s)
        if not d_str: continue
        # Determine if week is in the past (only show past weeks on the chart)
        try:
            d_obj = _dt.datetime.strptime(d_str, '%Y-%m-%d')
        except Exception:
            d_obj = None
        if d_obj and d_obj > TODAY_REF: continue
        vals = []
        for e in employees:
            v = e.get('weeks', {}).get(col_s)
            n = pn(v)
            if n is not None: vals.append(n)
        if not vals: continue
        avg = sum(vals) / len(vals)
        out.append({'date': d_str, 'actual': avg})
    return out


def _exec_weekly_line_chart(weekly):
    """Bar+line chart: bars for total active employees (proxied via count),
    or just a clean utilization line over time with 80% target reference."""
    if not weekly: return ''
    W, H = 920, 300
    M = {'top': 32, 'right': 22, 'bottom': 60, 'left': 60}
    IW = W - M['left'] - M['right']
    IH = H - M['top'] - M['bottom']
    n = len(weekly)
    actuals = [w['actual'] for w in weekly]
    y_min = max(0.0, (min(actuals) // 0.05) * 0.05 - 0.05)
    y_max = min(1.0, (max(actuals) // 0.05) * 0.05 + 0.10)
    if y_max - y_min < 0.20: y_max = y_min + 0.30
    def xf(i): return M['left'] + IW * i / max(1, n - 1)
    def yf(v): return M['top'] + IH * (1 - (v - y_min) / (y_max - y_min))
    s = [f'<svg viewBox="0 0 {W} {H}" class="exec-line-chart" preserveAspectRatio="xMidYMid meet">']
    # gridlines
    step = 0.05
    v = y_min
    while v <= y_max + 0.001:
        y = yf(v)
        s.append(f'<line x1="{M["left"]}" y1="{y:.1f}" x2="{M["left"]+IW}" y2="{y:.1f}" stroke="#eef2f7" stroke-width="1" />')
        s.append(f'<text x="{M["left"]-8}" y="{y:.1f}" font-size="10" fill="#94a3b8" text-anchor="end" dominant-baseline="middle">{int(round(v*100))}%</text>')
        v += step
    # 80% target reference line
    if y_min <= 0.80 <= y_max:
        ty = yf(0.80)
        s.append(f'<line x1="{M["left"]}" y1="{ty:.1f}" x2="{M["left"]+IW}" y2="{ty:.1f}" stroke="#ef4444" stroke-width="1.5" stroke-dasharray="6 4" />')
        s.append(f'<text x="{M["left"]+IW-4}" y="{ty-4:.1f}" font-size="10" fill="#ef4444" font-weight="700" text-anchor="end">80% target</text>')
    # actual line — Earth Brown
    pts = []
    for i, w in enumerate(weekly):
        pts.append((xf(i), yf(w['actual'])))
    if len(pts) >= 2:
        d = 'M ' + ' L '.join(f'{x:.1f} {y:.1f}' for x, y in pts)
        s.append(f'<path d="{d}" stroke="#382e30" stroke-width="2.5" fill="none" stroke-linecap="round" stroke-linejoin="round" />')
    # dots
    for x, y in pts:
        s.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3" fill="#382e30" />')
    # X labels — every ~4 weeks
    every = max(1, n // 12)
    for i, w in enumerate(weekly):
        if i % every != 0 and i != n-1: continue
        x = xf(i)
        try:
            d_obj = _dt.datetime.strptime(w['date'], '%Y-%m-%d')
            lbl = d_obj.strftime('%b %d')
        except Exception:
            lbl = w['date'][5:]
        s.append(f'<text x="{x:.1f}" y="{M["top"]+IH+18}" font-size="9" fill="#64748b" text-anchor="middle">{esc(lbl)}</text>')
    # Latest value annotation
    if pts:
        lx, ly = pts[-1]
        latest = weekly[-1]['actual']
        s.append(f'<circle cx="{lx:.1f}" cy="{ly:.1f}" r="5" fill="#ad8c63" stroke="#fff" stroke-width="2" />')
        s.append(f'<text x="{lx-8:.1f}" y="{ly-10:.1f}" font-size="11" font-weight="700" fill="#382e30" text-anchor="end">{latest*100:.1f}%</text>')
    s.append('</svg>')
    return ''.join(s)


def _exec_line_chart_v2(monthly):
    """Updated chart: wider left margin so Actual/Target value-row labels don't crowd Jan column.
    Also slightly shorter so the page fits without scrolling."""
    W, H = 640, 290
    M = {'top': 14, 'right': 16, 'bottom': 78, 'left': 78}
    IW = W - M['left'] - M['right']
    IH = H - M['top'] - M['bottom']
    all_vals = []
    for m in monthly:
        for k in ('actual', 'target'):
            n = pn(m[k])
            if n is not None: all_vals.append(n)
    if not all_vals: return ''
    y_min = max(0, (min(all_vals) // 0.05) * 0.05 - 0.05)
    y_max = (max(all_vals) // 0.05) * 0.05 + 0.05
    if y_max - y_min < 0.20: y_max = y_min + 0.25
    n_months = max(12, len(monthly))
    def xf(i): return M['left'] + IW * i / (n_months - 1)
    def yf(v): return M['top'] + IH * (1 - (v - y_min) / (y_max - y_min))
    s = [f'<svg viewBox="0 0 {W} {H}" class="exec-line-chart" preserveAspectRatio="xMidYMid meet">']
    step = 0.05
    v = y_min
    while v <= y_max + 0.001:
        y = yf(v)
        s.append(f'<line x1="{M["left"]}" y1="{y:.1f}" x2="{M["left"]+IW}" y2="{y:.1f}" stroke="#eef2f7" stroke-width="1" />')
        s.append(f'<text x="{M["left"]-8}" y="{y:.1f}" font-size="10" fill="#94a3b8" text-anchor="end" dominant-baseline="middle">{int(round(v*100))}%</text>')
        v += step
    tpts = []
    for i, m in enumerate(monthly):
        n = pn(m['target'])
        if n is not None: tpts.append((xf(i), yf(n)))
    if len(tpts) >= 2:
        d = 'M ' + ' L '.join(f'{x:.1f} {y:.1f}' for x, y in tpts)
        s.append(f'<path d="{d}" stroke="#9ca3af" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" />')
    apts = []
    for i, m in enumerate(monthly):
        n = pn(m['actual'])
        if n is not None: apts.append((xf(i), yf(n)))
    if len(apts) >= 2:
        d = 'M ' + ' L '.join(f'{x:.1f} {y:.1f}' for x, y in apts)
        s.append(f'<path d="{d}" stroke="#0f172a" stroke-width="2.5" fill="none" stroke-linecap="round" stroke-linejoin="round" />')
    label_y = M['top'] + IH + 16
    for i, m in enumerate(monthly[:12]):
        x = xf(i)
        s.append(f'<text x="{x:.1f}" y="{label_y}" font-size="10" fill="#64748b" text-anchor="middle">{esc(m["month"])}</text>')
    row_a = label_y + 22
    row_t = row_a + 16
    # Place labels well inside the wider left margin
    s.append(f'<text x="{M["left"]-12}" y="{row_a-1}" font-size="10" fill="#0f172a" text-anchor="end" font-weight="700" letter-spacing="1">Actual</text>')
    s.append(f'<text x="{M["left"]-12}" y="{row_t-1}" font-size="10" fill="#9ca3af" text-anchor="end" font-weight="700" letter-spacing="1">Target</text>')
    for i, m in enumerate(monthly[:12]):
        x = xf(i)
        a = pn(m['actual']); t = pn(m['target'])
        if a is not None:
            s.append(f'<text x="{x:.1f}" y="{row_a-1}" font-size="10" fill="#0f172a" font-weight="600" text-anchor="middle" font-variant-numeric="tabular-nums">{int(round(a*100))}%</text>')
        if t is not None:
            s.append(f'<text x="{x:.1f}" y="{row_t-1}" font-size="10" fill="#9ca3af" text-anchor="middle" font-variant-numeric="tabular-nums">{int(round(t*100))}%</text>')
    s.append('</svg>')
    return ''.join(s)


def _exec_bar_chart_v2(disciplines):
    """Bar chart sized to match the line chart's overall height. X axis 0-100%."""
    n = len(disciplines)
    LBL = 110
    PAD_R = 26
    W = 520
    CHART_W = W - LBL - PAD_R
    AX_H = 22
    TARGET_H = 290
    BH = max(38, (TARGET_H - AX_H - 4) // max(1, n))
    H = n * BH + AX_H + 4
    scale_max = 1.0  # cap at 100%
    BAR_BG = '#efe8d6'   # warm cream track
    BAR_FG = '#382e30'   # Earth Brown — matches Bright Spots
    s = [f'<svg viewBox="0 0 {W} {H}" class="exec-bar-chart" preserveAspectRatio="xMidYMid meet">']
    for i, d in enumerate(disciplines):
        y_top = i * BH
        y_mid = y_top + BH/2
        a = pn(d['actual']); t = pn(d['target'])
        s.append(f'<text x="{LBL-12}" y="{y_mid}" font-size="13" fill="#0f172a" text-anchor="end" dominant-baseline="middle">{esc(d["name"])}</text>')
        bar_y = y_top + 7
        bar_h = BH - 14
        s.append(f'<rect x="{LBL}" y="{bar_y}" width="{CHART_W}" height="{bar_h}" rx="{bar_h/2:.1f}" fill="{BAR_BG}" />')
        if a is not None:
            bw = CHART_W * min(scale_max, a) / scale_max
            s.append(f'<rect x="{LBL}" y="{bar_y}" width="{bw:.1f}" height="{bar_h}" rx="{bar_h/2:.1f}" fill="{BAR_FG}" />')
        if t is not None:
            tx = LBL + CHART_W * (min(scale_max, t) / scale_max)
            s.append(f'<line x1="{tx:.1f}" y1="{y_top+3}" x2="{tx:.1f}" y2="{y_top+BH-3}" stroke="{BAR_FG}" stroke-width="2" />')
    base_y = n * BH + 2
    for pct in [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]:
        x = LBL + CHART_W * (pct / scale_max)
        s.append(f'<text x="{x:.1f}" y="{base_y+12}" font-size="10" fill="#94a3b8" text-anchor="middle">{int(pct*100)}%</text>')
    s.append('</svg>')
    return ''.join(s)


def render_exec_summary_v2(data):
    """Round 2: chart label fix + tightened layout so Bright Spots fits without scrolling."""
    rows = data['rows']
    subtitle = ''
    for r in rows[1:5]:
        for v in r:
            if isinstance(v, str) and v.strip() and 'EXECUTIVE' not in v.upper() and 'E X E C U T I V E' not in v:
                subtitle = v.strip(); break
        if subtitle: break
    bright = ''
    for r in rows[25:50]:
        for v in r:
            if isinstance(v, str) and 'bright spot' in v.lower():
                bright = v.strip(); break
        if bright: break
    as_of = None
    for r in rows[50:65]:
        for v in r:
            if is_date(v):
                d = to_date(v)
                if d and d.year >= 2025 and d.year <= 2027:
                    as_of = d; break
        if as_of: break
    monthly = []
    trend_start = find_row_with(rows, 'Actual', start=50)
    if trend_start != -1:
        for i in range(trend_start + 1, min(trend_start + 14, len(rows))):
            r = rows[i]
            if not r or not r[0]: continue
            if isinstance(r[0], str) and len(r[0]) <= 4 and r[0].strip():
                monthly.append({'month': r[0], 'actual': r[1] if len(r) > 1 else None, 'target': r[2] if len(r) > 2 else None, 'prior': r[3] if len(r) > 3 else None})
    disc_start = find_row_with(rows, 'Dicipline', start=60)
    if disc_start == -1: disc_start = find_row_with(rows, 'Discipline', start=60)
    disciplines = []
    if disc_start != -1:
        for i in range(disc_start + 1, len(rows)):
            r = rows[i]
            if not r or not r[0] or not isinstance(r[0], str): continue
            disciplines.append({'name': r[0], 'actual': r[1] if len(r) > 1 else None, 'target': r[3] if len(r) > 3 else None})

    ytd_actual = None; ytd_target = None; latest_idx = -1
    for i, m in enumerate(monthly):
        if pn(m['actual']) is not None:
            ytd_actual = pn(m['actual']); ytd_target = pn(m['target']); latest_idx = i
    prior_actual = pn(monthly[latest_idx]['prior']) if latest_idx >= 0 else None
    yoy_diff = (ytd_actual - prior_actual) if (ytd_actual is not None and prior_actual is not None) else None
    vs_target = (ytd_actual - ytd_target) if (ytd_actual is not None and ytd_target is not None) else None
    fy26_targets = [pn(m['target']) for m in monthly if pn(m['target']) is not None]
    fy26_target = sum(fy26_targets)/len(fy26_targets) if fy26_targets else ytd_target
    below_n, total_offices = _exec_compute_below_target()
    as_of_str = as_of.strftime('%B %-d, %Y').upper() if as_of else 'FY26'
    breadcrumb = f'EXECUTIVE SUMMARY · FP&amp;A · {as_of_str}'

    p = []
    p.append('<div class="exec-page exec-tight">')
    p.append(f'<div class="exec-eyebrow">{breadcrumb}</div>')
    p.append('<h1 class="exec-display">Company performance.</h1>')
    if subtitle:
        p.append(f'<p class="exec-subtitle">{esc(subtitle)}</p>')
    p.append('<div class="exec-kpi-grid">')
    if ytd_actual is not None:
        p.append('<div class="ek"><div class="ek-eyebrow">YTD UTIL.</div>')
        p.append(f'<div class="ek-num">{ytd_actual*100:.1f}<span class="ek-unit">%</span></div>')
        p.append('<div class="ek-foot">Company-wide</div></div>')
    if fy26_target is not None:
        p.append('<div class="ek"><div class="ek-eyebrow">FY26 TARGET</div>')
        p.append(f'<div class="ek-num">{fy26_target*100:.1f}<span class="ek-unit">%</span></div>')
        p.append('<div class="ek-foot">Weighted plan</div></div>')
    if vs_target is not None:
        sign = '−' if vs_target < 0 else '+'
        status_cls = 'bad' if vs_target < 0 else 'good'
        status_lbl = 'Behind' if vs_target < 0 else 'Ahead'
        p.append('<div class="ek"><div class="ek-eyebrow">VARIANCE</div>')
        p.append(f'<div class="ek-num">{sign}{abs(vs_target)*100:.1f}<span class="ek-unit">pts</span></div>')
        p.append('<div class="ek-foot">YTD vs FY26 plan</div>')
        p.append(f'<div class="ek-pill"><span class="pill {status_cls}">{status_lbl}</span></div></div>')
    if prior_actual is not None:
        p.append('<div class="ek"><div class="ek-eyebrow">PRIOR YEAR</div>')
        p.append(f'<div class="ek-num">{prior_actual*100:.1f}<span class="ek-unit">%</span></div>')
        p.append('<div class="ek-foot">FY25 same period</div></div>')
    if yoy_diff is not None:
        sign = '+' if yoy_diff >= 0 else '−'
        status_cls = 'good' if yoy_diff >= 0 else 'bad'
        status_lbl = 'Up' if yoy_diff >= 0 else 'Down'
        p.append('<div class="ek"><div class="ek-eyebrow">YOY CHANGE</div>')
        p.append(f'<div class="ek-num">{sign}{abs(yoy_diff)*100:.1f}<span class="ek-unit">pts</span></div>')
        p.append('<div class="ek-foot">vs prior year</div>')
        p.append(f'<div class="ek-pill"><span class="pill {status_cls}">{status_lbl}</span></div></div>')
    if below_n is not None and total_offices:
        p.append('<div class="ek"><div class="ek-eyebrow">BELOW TGT</div>')
        p.append(f'<div class="ek-num">{below_n}<span class="ek-unit">/{total_offices}</span></div>')
        p.append(f'<div class="ek-foot">Of {total_offices} active offices</div></div>')
    p.append('</div>')
    p.append('<div class="exec-charts">')
    if monthly:
        p.append('<div class="exec-chart-block">')
        p.append('<div class="exec-section-eyebrow">MONTHLY TREND · YTD</div>')
        p.append('<div class="exec-section-sub">Actual vs target, FY26</div>')
        p.append(_exec_line_chart_v2(monthly))
        p.append('</div>')
    if disciplines:
        p.append('<div class="exec-chart-block">')
        p.append('<div class="exec-section-eyebrow">THE BIG FOUR</div>')
        p.append('<div class="exec-section-sub">Utilization vs target, YTD</div>')
        p.append(_exec_bar_chart_v2(disciplines))
        p.append('</div>')
    p.append('</div>')
    if bright:
        # Strip leading "Bright spots — " if present, since we render the eyebrow separately
        body = bright
        for prefix in ('Bright spots —', 'Bright spots -', 'Bright spots:'):
            if body.lower().startswith(prefix.lower()):
                body = body[len(prefix):].lstrip()
                break
        p.append('<div class="exec-bright">')
        p.append('<div class="exec-bright-eyebrow">Bright Spots</div>')
        p.append(f'<div class="exec-bright-body">{esc(body)}</div>')
        p.append('</div>')
    p.append('</div>')
    return ''.join(p)


# Additional CSS to tighten exec layout


def render_exec_summary_v3(data):
    """Round 3: same as v2 but adds an MTD Billable Hours KPI card to the right
    of "Below Tgt", sourced from site_data/_mtd_billable.json (produced by
    extract_data.py from the workbook's Raw Data sheet).
    """
    rows = data['rows']
    subtitle = ''
    for r in rows[1:5]:
        for v in r:
            if isinstance(v, str) and v.strip() and 'EXECUTIVE' not in v.upper() and 'E X E C U T I V E' not in v:
                subtitle = v.strip(); break
        if subtitle: break
    bright = ''
    for r in rows[25:50]:
        for v in r:
            if isinstance(v, str) and 'bright spot' in v.lower():
                bright = v.strip(); break
        if bright: break
    as_of = None
    for r in rows[50:65]:
        for v in r:
            if is_date(v):
                d = to_date(v)
                if d and d.year >= 2025 and d.year <= 2027:
                    as_of = d; break
        if as_of: break
    monthly = []
    trend_start = find_row_with(rows, 'Actual', start=50)
    if trend_start != -1:
        for i in range(trend_start + 1, min(trend_start + 14, len(rows))):
            r = rows[i]
            if not r or not r[0]: continue
            if isinstance(r[0], str) and len(r[0]) <= 4 and r[0].strip():
                monthly.append({'month': r[0], 'actual': r[1] if len(r) > 1 else None, 'target': r[2] if len(r) > 2 else None, 'prior': r[3] if len(r) > 3 else None})
    disc_start = find_row_with(rows, 'Dicipline', start=60)
    if disc_start == -1: disc_start = find_row_with(rows, 'Discipline', start=60)
    disciplines = []
    if disc_start != -1:
        for i in range(disc_start + 1, len(rows)):
            r = rows[i]
            if not r or not r[0] or not isinstance(r[0], str): continue
            disciplines.append({'name': r[0], 'actual': r[1] if len(r) > 1 else None, 'target': r[3] if len(r) > 3 else None})

    ytd_actual = None; ytd_target = None; latest_idx = -1
    for i, m in enumerate(monthly):
        if pn(m['actual']) is not None:
            ytd_actual = pn(m['actual']); ytd_target = pn(m['target']); latest_idx = i
    prior_actual = pn(monthly[latest_idx]['prior']) if latest_idx >= 0 else None
    yoy_diff = (ytd_actual - prior_actual) if (ytd_actual is not None and prior_actual is not None) else None
    vs_target = (ytd_actual - ytd_target) if (ytd_actual is not None and ytd_target is not None) else None
    fy26_targets = [pn(m['target']) for m in monthly if pn(m['target']) is not None]
    fy26_target = sum(fy26_targets)/len(fy26_targets) if fy26_targets else ytd_target
    below_n, total_offices = _exec_compute_below_target()

    # v3 addition: load MTD Billable Hours from extract_data output
    mtd_info = None
    try:
        _mtd_path = DATA_DIR / '_mtd_billable.json'
        if _mtd_path.exists():
            with open(_mtd_path) as _f:
                mtd_info = json.load(_f)
    except Exception:
        mtd_info = None

    as_of_str = as_of.strftime('%B %-d, %Y').upper() if as_of else 'FY26'
    breadcrumb = f'EXECUTIVE SUMMARY · FP&amp;A · {as_of_str}'

    p = []
    p.append('<div class="exec-page exec-tight">')
    p.append(f'<div class="exec-eyebrow">{breadcrumb}</div>')
    p.append('<h1 class="exec-display">Company performance.</h1>')
    if subtitle:
        p.append(f'<p class="exec-subtitle">{esc(subtitle)}</p>')
    p.append('<div class="exec-kpi-grid">')
    if ytd_actual is not None:
        p.append('<div class="ek"><div class="ek-eyebrow">YTD UTIL.</div>')
        p.append(f'<div class="ek-num">{ytd_actual*100:.1f}<span class="ek-unit">%</span></div>')
        p.append('<div class="ek-foot">Company-wide</div></div>')
    if fy26_target is not None:
        p.append('<div class="ek"><div class="ek-eyebrow">FY26 TARGET</div>')
        p.append(f'<div class="ek-num">{fy26_target*100:.1f}<span class="ek-unit">%</span></div>')
        p.append('<div class="ek-foot">Weighted plan</div></div>')
    if vs_target is not None:
        sign = '−' if vs_target < 0 else '+'
        status_cls = 'bad' if vs_target < 0 else 'good'
        status_lbl = 'Behind' if vs_target < 0 else 'Ahead'
        p.append('<div class="ek"><div class="ek-eyebrow">VARIANCE</div>')
        p.append(f'<div class="ek-num">{sign}{abs(vs_target)*100:.1f}<span class="ek-unit">pts</span></div>')
        p.append('<div class="ek-foot">YTD vs FY26 plan</div>')
        p.append(f'<div class="ek-pill"><span class="pill {status_cls}">{status_lbl}</span></div></div>')
    if prior_actual is not None:
        p.append('<div class="ek"><div class="ek-eyebrow">PRIOR YEAR</div>')
        p.append(f'<div class="ek-num">{prior_actual*100:.1f}<span class="ek-unit">%</span></div>')
        p.append('<div class="ek-foot">FY25 same period</div></div>')
    if yoy_diff is not None:
        sign = '+' if yoy_diff >= 0 else '−'
        status_cls = 'good' if yoy_diff >= 0 else 'bad'
        status_lbl = 'Up' if yoy_diff >= 0 else 'Down'
        p.append('<div class="ek"><div class="ek-eyebrow">YOY CHANGE</div>')
        p.append(f'<div class="ek-num">{sign}{abs(yoy_diff)*100:.1f}<span class="ek-unit">pts</span></div>')
        p.append('<div class="ek-foot">vs prior year</div>')
        p.append(f'<div class="ek-pill"><span class="pill {status_cls}">{status_lbl}</span></div></div>')
    if below_n is not None and total_offices:
        p.append('<div class="ek"><div class="ek-eyebrow">BELOW TGT</div>')
        p.append(f'<div class="ek-num">{below_n}<span class="ek-unit">/{total_offices}</span></div>')
        p.append(f'<div class="ek-foot">Of {total_offices} active offices</div></div>')

    # MTD Billable Hours card — to the right of Below Tgt
    if mtd_info and mtd_info.get('mtd_billable_hours') is not None:
        _mtd_hrs = float(mtd_info['mtd_billable_hours'])
        _month = mtd_info.get('month_name') or 'This month'
        _yr = mtd_info.get('year') or TODAY_REF.year
        _mo_num = None
        _my = mtd_info.get('month_year')
        if isinstance(_my, str) and '-' in _my:
            try: _mo_num = int(_my.split('-')[1])
            except ValueError: _mo_num = None
        if _mo_num is None:
            _mo_num = TODAY_REF.month
        # Workday counts (Mon-Fri) for the month
        import calendar as _cal
        _days_in_month = _cal.monthrange(_yr, _mo_num)[1]
        _last_raw = mtd_info.get('last_entry_date')
        # "Through" day = latest entry date day if available, else min(today day, days in month)
        _through_day = None
        if isinstance(_last_raw, str) and len(_last_raw) >= 10:
            try: _through_day = int(_last_raw[8:10])
            except ValueError: _through_day = None
        if _through_day is None:
            _through_day = min(TODAY_REF.day, _days_in_month)
        def _count_workdays(year, month, start_day, end_day):
            import datetime as _d
            n = 0
            for d in range(start_day, end_day + 1):
                if _d.date(year, month, d).weekday() < 5: n += 1
            return n
        _workdays_total = _count_workdays(_yr, _mo_num, 1, _days_in_month)
        _workdays_to_date = _count_workdays(_yr, _mo_num, 1, _through_day)
        _forecast_hrs = None
        if _workdays_to_date > 0:
            _forecast_hrs = _mtd_hrs * _workdays_total / _workdays_to_date
        _foot = f'{_month} 1 – {_through_day} · billable only'
        p.append('<div class="ek"><div class="ek-eyebrow">MTD BILLABLE</div>')
        p.append(f'<div class="ek-num">{_mtd_hrs:,.0f}<span class="ek-unit">hrs</span></div>')
        if _forecast_hrs is not None:
            p.append(f'<div class="ek-foot">Monthly Forecast: {_forecast_hrs:,.0f} hrs</div>')
        p.append(f'<div class="ek-foot">{esc(_foot)}</div></div>')

    p.append('</div>')
    p.append('<div class="exec-charts">')
    if monthly:
        p.append('<div class="exec-chart-block">')
        p.append('<div class="exec-section-eyebrow">MONTHLY TREND · YTD</div>')
        p.append('<div class="exec-section-sub">Actual vs target, FY26</div>')
        p.append(_exec_line_chart_v2(monthly))
        p.append('</div>')
    if disciplines:
        p.append('<div class="exec-chart-block">')
        p.append('<div class="exec-section-eyebrow">THE BIG FOUR</div>')
        p.append('<div class="exec-section-sub">Utilization vs target, YTD</div>')
        p.append(_exec_bar_chart_v2(disciplines))
        p.append('</div>')
    p.append('</div>')
    if bright:
        body = bright
        for prefix in ('Bright spots —', 'Bright spots -', 'Bright spots:'):
            if body.lower().startswith(prefix.lower()):
                body = body[len(prefix):].lstrip()
                break
        p.append('<div class="exec-bright">')
        p.append('<div class="exec-bright-eyebrow">Bright Spots</div>')
        p.append(f'<div class="exec-bright-body">{esc(body)}</div>')
        p.append('</div>')
    p.append('</div>')
    return ''.join(p)


# Swap exec summary renderer to v3
RENDERERS['exec_summary'] = render_exec_summary_v3

CSS = r"""
:root {
  --good-bg: #d1f4dd;
  --good-text: #14532d;
  --good-strong: #16a34a;
  --bad-bg: #fee2e4;
  --bad-text: #991b1b;
  --bad-strong: #dc2626;
  --na-bg: #f1f5f9;
  --na-text: #94a3b8;
  --neutral-bg: #e0e7ff;
  --neutral-text: #3730a3;
  --side-bg: #382e30;
  --side-text: #e2e8f0;
  --side-dim: #c9b89a;
  --side-active-bg: #4a3f41;
  --side-active-text: #fff;
  --accent: #ad8c63;
  --bg: #f6f7fa;
  --surface: #ffffff;
  --border: #e5e7eb;
  --text: #0f172a;
  --text-muted: #475569;
  --text-dim: #94a3b8;
  --heat-base: 16, 185, 129;
}
* { box-sizing: border-box; }
html, body { height: 100%; margin: 0; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", system-ui, "Segoe UI", "Helvetica Neue", Arial, sans-serif;
  color: var(--text);
  background: var(--bg);
  -webkit-font-smoothing: antialiased;
}
.app { display: flex; min-height: 100vh; }
.sidebar {
  width: 272px;
  background: var(--side-bg);
  color: var(--side-text);
  display: flex;
  flex-direction: column;
  position: sticky;
  top: 0;
  height: 100vh;
  overflow-y: auto;
  flex-shrink: 0;
}
.brand { padding: 24px 24px 18px; border-bottom: 1px solid #4a3f41; }
.brand-name { font-size: 15px; font-weight: 700; letter-spacing: 0.3px; }
.brand-sub { font-size: 12px; color: var(--side-dim); margin-top: 3px; }
.nav { flex: 1; padding: 12px 0 24px; }
.nav-section {
  padding: 16px 24px 6px;
  font-size: 10px;
  letter-spacing: 1.5px;
  text-transform: uppercase;
  color: var(--side-dim);
  font-weight: 600;
}
.nav-link {
  display: block;
  padding: 8px 24px;
  color: var(--side-text);
  text-decoration: none;
  font-size: 13px;
  border-left: 3px solid transparent;
}
.nav-link:hover { background: var(--side-active-bg); color: #fff; }
.nav-link.active {
  background: var(--side-active-bg);
  border-left-color: var(--accent);
  color: var(--side-active-text);
  font-weight: 600;
}
.sidebar-foot { padding: 14px 24px; font-size: 11px; color: var(--side-dim); border-top: 1px solid #4a3f41; }
.main { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.topbar {
  background: var(--surface);
  border-bottom: 1px solid var(--border);
  padding: 16px 32px;
  position: sticky;
  top: 0;
  z-index: 40;
}
.topbar h1 { font-size: 16px; margin: 0; font-weight: 600; color: var(--text); letter-spacing: 0.2px; }
.tab-host { padding: 24px 32px 80px; }
.tab-content { display: none; }
.tab-content.active { display: block; }
.page-intro { margin: 0 0 18px; }
.page-intro h2 { font-size: 22px; margin: 0; font-weight: 600; letter-spacing: -0.2px; }
.page-sub { font-size: 13px; color: var(--text-muted); margin: 4px 0 0; max-width: 720px; line-height: 1.5; }
.card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 10px;
  overflow: hidden;
  box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
  margin-bottom: 18px;
}
.card-head {
  padding: 14px 18px 12px;
  border-bottom: 1px solid var(--border);
  background: linear-gradient(180deg, #fafbfc, var(--surface));
}
.card-head h3 { font-size: 14px; margin: 0; font-weight: 600; letter-spacing: 0.1px; }
.card-sub { font-size: 12px; color: var(--text-muted); margin-top: 2px; display: block; }
.card-body { padding: 14px 18px; overflow: auto; }
.two-col { display: grid; grid-template-columns: 1fr 1.4fr; gap: 18px; }
@media (max-width: 1200px) { .two-col { grid-template-columns: 1fr; } }
.kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; margin-bottom: 18px; }
.kpi {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 16px 18px;
  box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
}
.kpi-label { font-size: 11px; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; }
.kpi-value { font-size: 28px; font-weight: 700; margin-top: 6px; font-variant-numeric: tabular-nums; letter-spacing: -0.5px; }
.kpi-value.pos { color: var(--good-strong); }
.kpi-value.neg { color: var(--bad-strong); }
.scroll-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 10px;
  overflow: auto;
  max-height: calc(100vh - 220px);
  box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
}
.scroll-card-inner { overflow: auto; max-height: 60vh; }
/* Single-scrollbar variant for Overtime, Non-Posting Hours, Proposal Hours,
   Recovery Planning, and Non-Targeted Employees — disables inner scroll so
   only the main page scrollbar appears. */
.scroll-card.scroll-flat,
.scroll-card-inner.scroll-flat { overflow: visible !important; max-height: none !important; }

.data-table, .pmo-grid, .bubble-grid, .heat-grid, .emp-pill-grid, .trend-table, .tree-table {
  border-collapse: separate;
  border-spacing: 0;
  font-size: 12px;
  white-space: nowrap;
  width: 100%;
}
.data-table th, .data-table td,
.pmo-grid th, .pmo-grid td,
.bubble-grid th, .bubble-grid td,
.heat-grid th, .heat-grid td,
.emp-pill-grid th, .emp-pill-grid td,
.trend-table th, .trend-table td,
.tree-table th, .tree-table td {
  padding: 7px 12px;
  border-bottom: 1px solid var(--border);
  text-align: left;
  vertical-align: middle;
  background: var(--surface);
}
.data-table thead th,
.pmo-grid thead th,
.bubble-grid thead th,
.heat-grid thead th,
.emp-pill-grid thead th,
.trend-table thead th,
.tree-table thead th {
  background: #fafbfc;
  color: var(--text-muted);
  font-weight: 600;
  font-size: 11px;
  letter-spacing: 0.3px;
  text-transform: uppercase;
  position: sticky;
  top: 0;
  z-index: 5;
  border-bottom: 1px solid #cbd5e1;
}
.data-table tbody tr:nth-child(2n) td { background: #fbfcfd; }
.data-table tbody tr:hover td, .pmo-grid tbody tr:hover td, .heat-grid tbody tr:hover td, .emp-pill-grid tbody tr:hover td { background: #f3f5f7; }
.data-table td.num, .pmo-grid td.num, .bubble-grid td.num, .heat-grid td.num, .emp-pill-grid td.num, .trend-table td.num, .tree-table td.num,
.data-table th.num, .pmo-grid th.num, .bubble-grid th.num, .heat-grid th.num, .emp-pill-grid th.num, .trend-table th.num, .tree-table th.num {
  text-align: right; font-variant-numeric: tabular-nums;
}
.data-table td.date, .emp-pill-grid td.date { color: var(--text-muted); }
.muted { color: var(--text-dim); }
.delta { font-weight: 600; font-variant-numeric: tabular-nums; font-size: 11px; }
.delta.pos { color: var(--good-strong); }
.delta.neg { color: var(--bad-strong); }
tr.section-row td {
  background: #f1f5f9 !important;
  font-weight: 600;
  color: var(--text);
  font-size: 12px;
  letter-spacing: 0.3px;
}
tr.total-row td {
  background: #f8fafc !important;
  font-weight: 700;
  color: var(--text);
  border-top: 1px solid #cbd5e1;
}
.pill {
  display: inline-block;
  min-width: 50px;
  padding: 3px 10px;
  border-radius: 999px;
  font-weight: 600;
  font-size: 11px;
  text-align: center;
  letter-spacing: 0.2px;
  white-space: nowrap;
  font-variant-numeric: tabular-nums;
}
.pill.good { background: var(--good-bg); color: var(--good-text); }
.pill.bad { background: var(--bad-bg); color: var(--bad-text); }
.pill.na { background: var(--na-bg); color: var(--na-text); font-weight: 500; }
.pill.neutral { background: var(--neutral-bg); color: var(--neutral-text); }
.toolbar { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; margin-bottom: 14px; }
.input {
  padding: 7px 12px;
  border: 1px solid #cbd5e1;
  border-radius: 7px;
  font-size: 13px;
  background: var(--surface);
  outline: none;
  color: var(--text);
  font-family: inherit;
}
.input:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(173,140,99,0.25); }
.legend { margin-left: auto; display: flex; gap: 6px; align-items: center; }
.result-count { font-size: 12px; color: var(--text-muted); margin-bottom: 10px; }
.bubble-grid thead tr.row-quarter th { top: 0; height: 30px; font-size: 12px; text-transform: uppercase; }
.bubble-grid thead tr.row-week th { top: 30px; height: 26px; font-size: 11px; color: var(--text-dim); font-weight: 500; text-transform: none; letter-spacing: 0; }
.bubble-grid .col-office, .bubble-grid .col-name, .bubble-grid .col-target {
  position: sticky; background: var(--surface); z-index: 4; text-align: left;
}
.bubble-grid thead .col-office, .bubble-grid thead .col-name, .bubble-grid thead .col-target {
  background: #fafbfc; z-index: 15;
}
.bubble-grid .col-office { left: 0; min-width: 130px; max-width: 130px; padding-left: 16px; }
.bubble-grid .col-name { left: 130px; min-width: 200px; max-width: 200px; font-weight: 500; }
.bubble-grid .col-target { left: 330px; min-width: 65px; max-width: 65px; text-align: center; color: var(--text-muted); }
.quarter-sep { border-left: 2px solid #cbd5e1; }
.ytd-sep { border-left: 2px solid #cbd5e1; border-right: 2px solid #cbd5e1; }
.emp-id { color: var(--text-dim); font-size: 10px; font-weight: 400; }
.cell { text-align: center; padding: 4px 6px; }
.bubble-grid tr.office-group td {
  background: var(--side-bg) !important;
  color: #f1f5f9 !important;
  font-weight: 600;
  font-size: 11px;
  letter-spacing: 0.8px;
  text-transform: uppercase;
  padding: 8px 16px;
  position: sticky; left: 0;
  z-index: 3;
}
.bubble-grid tbody tr:not(.office-group):hover td { background: #f3f5f7; }
.bubble-grid tbody tr:not(.office-group):hover .col-office,
.bubble-grid tbody tr:not(.office-group):hover .col-name,
.bubble-grid tbody tr:not(.office-group):hover .col-target { background: #f3f5f7; }
.pmo-grid .col-office {
  position: sticky; left: 0; background: var(--surface); z-index: 4;
  min-width: 200px; padding-left: 16px;
}
.pmo-grid thead .col-office { background: #fafbfc; z-index: 15; }
.pmo-grid thead tr.hdr-period th { font-size: 11px; text-transform: uppercase; letter-spacing: 0.4px; }
.pmo-grid thead tr.hdr-sub th { top: 36px; font-size: 10px; text-transform: uppercase; letter-spacing: 0.3px; color: var(--text-dim); font-weight: 600; }
.pmo-grid .period-sep { border-left: 2px solid #e5e7eb; }
.pmo-grid td { font-size: 11px; padding: 5px 8px; }
.pmo-grid tbody tr:hover td { background: #f3f5f7; }
.heat-grid .col-office, .heat-grid .sticky-l {
  position: sticky; left: 0; background: var(--surface); z-index: 4;
  min-width: 180px; padding-left: 16px;
}
.heat-grid thead .col-office, .heat-grid thead .sticky-l { background: #fafbfc; z-index: 15; }
td.heat {
  background: rgba(var(--heat-base), calc(var(--intensity) * 0.55));
  color: var(--text);
  font-weight: 500;
}
td.zero { color: var(--text-dim); text-align: center; }
td.total-cell { font-weight: 700; background: #f8fafc; border-left: 2px solid #cbd5e1; }
td.sticky-l, th.sticky-l {
  position: sticky; left: 0; background: var(--surface); z-index: 4;
  font-weight: 500;
}
.data-table thead th.sticky-l { background: #fafbfc; z-index: 15; }
.discipline-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; padding: 14px 18px; }
.disc-card {
  background: #fafbfc;
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 14px;
}
.disc-name { font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-muted); font-weight: 600; }
.disc-value { font-size: 26px; font-weight: 700; font-variant-numeric: tabular-nums; margin: 6px 0 2px; }
.disc-value.good { color: var(--good-strong); }
.disc-value.bad { color: var(--bad-strong); }
.disc-target { font-size: 11px; color: var(--text-muted); margin-bottom: 8px; }
.bar-track { height: 6px; background: #e5e7eb; border-radius: 999px; overflow: hidden; }
.bar-fill { height: 100%; background: var(--text-dim); border-radius: 999px; }
.bar-fill.good { background: var(--good-strong); }
.bar-fill.bad { background: var(--bad-strong); }
.tree-table .project-row td { font-weight: 600; background: #f8fafc; }
.tree-table .task-row td.sticky-l { padding-left: 32px; color: var(--text-muted); font-size: 12px; }
.tree-table .indent { color: var(--text-dim); margin-right: 4px; }
.hidden { display: none !important; }
.empty { color: var(--text-muted); padding: 40px; font-style: italic; text-align: center; }
/* Clickable employee names in Employee Utilization → Employee Dashboard */
.col-name a.emp-name-link { color: #382e30; text-decoration: none; cursor: pointer; border-bottom: 1px dotted transparent; }
.col-name a.emp-name-link:hover { color: #ad8c63; border-bottom-color: #ad8c63; }
.col-name a.emp-name-link:focus { outline: 2px solid #ad8c63; outline-offset: 2px; }
"""

JS = r"""
(function(){
  var links = document.querySelectorAll('.nav-link');
  var sections = document.querySelectorAll('.tab-content');
  var title = document.getElementById('tabTitle');
  function activate(slug, scroll) {
    links.forEach(function(l){ l.classList.toggle('active', l.dataset.tab === slug); });
    sections.forEach(function(s){ s.classList.toggle('active', s.id === slug); });
    var sec = document.getElementById(slug);
    if (sec) { title.textContent = sec.dataset.title || slug; if (scroll) window.scrollTo(0, 0); }
  }
  links.forEach(function(l){
    l.addEventListener('click', function(e){
      e.preventDefault();
      var slug = l.dataset.tab;
      history.replaceState(null, '', '#' + slug);
      activate(slug, true);
    });
  });
  var initial = (location.hash || '#').slice(1);
  if (initial && document.getElementById(initial)) activate(initial, false);
  else if (links.length) activate(links[0].dataset.tab, false);
  var search = document.getElementById('bubbleSearch');
  var officeF = document.getElementById('bubbleOffice');
  var statusF = document.getElementById('bubbleStatus');
  var countEl = document.getElementById('bubbleCount');
  if (search && officeF && statusF) {
    var rows = document.querySelectorAll('tr.emp-row');
    var groups = document.querySelectorAll('tr.office-group');
    function apply() {
      var q = search.value.trim().toLowerCase();
      var off = officeF.value;
      var stat = statusF.value;
      var visible = 0;
      var seen = new Set();
      rows.forEach(function(r){
        var show = true;
        if (q && !((r.dataset.name||'').includes(q) || (r.dataset.id||'').includes(q))) show = false;
        if (off && r.dataset.office !== off) show = false;
        if (stat === 'bad' && r.dataset.hasBad !== '1') show = false;
        if (stat === 'good' && r.dataset.hasGood !== '1') show = false;
        r.classList.toggle('hidden', !show);
        if (show) { visible++; seen.add(r.dataset.office); }
      });
      groups.forEach(function(g){ g.classList.toggle('hidden', !seen.has(g.dataset.officeGroup)); });
      if (countEl) countEl.textContent = visible;
    }
    search.addEventListener('input', apply);
    officeF.addEventListener('change', apply);
    statusF.addEventListener('change', apply);
  }
})();

// Clickable employee names in Employee Utilization → Employee Dashboard for that person
(function(){
  document.addEventListener('click', function(e){
    var link = e.target.closest && e.target.closest('a.emp-name-link');
    if (!link) return;
    e.preventDefault();
    var name = link.getAttribute('data-emp-name') || link.textContent || '';
    name = name.trim();
    if (!name) return;

    // Switch to the Employee Dashboard tab by clicking the matching nav link
    var navLink = document.querySelector('.nav-link[data-tab="Employee_Dashboard"]');
    if (navLink) navLink.click();

    // Select the matching employee in the dashboard picker.
    // Picker option text format is "<name>  <office>" (two spaces) or just "<name>".
    var picker = document.getElementById('ed2EmpPicker');
    if (!picker) return;

    function trySelect(){
      var target = name.toLowerCase();
      var found = -1;
      for (var i = 0; i < picker.options.length; i++){
        var t = (picker.options[i].textContent || '').toLowerCase().trim();
        if (t === target || t.indexOf(target + '  ') === 0 || t.indexOf(target + ' ') === 0){
          found = i;
          break;
        }
      }
      if (found < 0){
        for (var j = 0; j < picker.options.length; j++){
          var t2 = (picker.options[j].textContent || '').toLowerCase().trim();
          if (t2.indexOf(target) === 0){ found = j; break; }
        }
      }
      if (found >= 0){
        picker.value = picker.options[found].value;
        var evt;
        try { evt = new Event('change', { bubbles: true }); }
        catch(err){ evt = document.createEvent('Event'); evt.initEvent('change', true, true); }
        picker.dispatchEvent(evt);
        return true;
      }
      return false;
    }

    if (!trySelect()){
      var tries = 0;
      var iv = setInterval(function(){
        tries++;
        if (trySelect() || tries > 20) clearInterval(iv);
      }, 50);
    }
  });
})();

"""

HTML_SHELL = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Northstar Heritage — Utilization Report</title>
<style>{css}</style>
</head>
<body>
<div class="app">
  <aside class="sidebar">
    <div class="brand">
      <img class="brand-logo" alt="Northstar Heritage" src="data:image/png;base64,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" />
      <div class="brand-sub">FP&amp;A Utilization Report · FY26</div>
    </div>
    <nav class="nav">{nav}</nav>
    <div class="sidebar-foot">Generated {today}</div>
  </aside>
  <main class="main">
    <header class="topbar">
      <h1 id="tabTitle">Executive Summary</h1>
    </header>
    <div class="tab-host">{sections}</div>
  </main>
</div>
<script>{js}</script>
</body>
</html>"""

def build():
    sections_html = []
    nav_html = []
    current_group = None
    for i, t in enumerate(TABS):
        if t['group'] != current_group:
            current_group = t['group']
            nav_html.append(f'<div class="nav-section">{esc(current_group)}</div>')
        active = ' active' if i == 0 else ''
        nav_html.append(f'<a href="#{t["slug"]}" class="nav-link{active}" data-tab="{t["slug"]}">{esc(t["name"])}</a>')
        data = load_tab(t['slug'])
        data['tab'] = t['name']
        renderer = RENDERERS.get(t['renderer'], render_premium_generic)
        try:
            body = renderer(data)
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            body = f'<div class="empty">Error rendering this tab: {esc(str(e))}<pre style="text-align:left;font-size:11px;">{esc(tb)}</pre></div>'
        sections_html.append(f'<section id="{t["slug"]}" class="tab-content{active}" data-title="{esc(t["name"])}">{body}</section>')
    today = datetime.date.today().strftime('%B %-d, %Y')
    html = HTML_SHELL.format(
        css=CSS, js=JS, nav=''.join(nav_html), sections=''.join(sections_html), today=today,
    )
    with open(OUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Wrote {OUT_FILE}: {os.path.getsize(OUT_FILE):,} bytes')

# Additional CSS for the redesigned Executive Summary (appended after main CSS)
EXEC_CSS = """
.exec-page { max-width: 1280px; }
.exec-eyebrow {
  font-size: 11px;
  font-weight: 700;
  color: #475569;
  letter-spacing: 3px;
  text-transform: uppercase;
  margin-bottom: 14px;
}
.exec-display {
  font-size: 64px;
  font-weight: 800;
  letter-spacing: -2px;
  margin: 0 0 10px;
  color: #0f172a;
  line-height: 1.0;
}
.exec-subtitle {
  font-size: 14px;
  color: #475569;
  margin: 0 0 28px;
  max-width: 760px;
}
.exec-kpi-grid {
  display: grid;
  grid-template-columns: repeat(7, minmax(0, 1fr));
  gap: 10px;
  margin-bottom: 36px;
}
@media (max-width: 1100px) { .exec-kpi-grid { grid-template-columns: repeat(4, minmax(0, 1fr)); } }
@media (max-width: 700px)  { .exec-kpi-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
.ek {
  background: #fbfbfc;
  border: 1px solid #e5e7eb;
  border-radius: 16px;
  padding: 18px 20px 20px;
  box-shadow: 0 2px 4px rgba(15,23,42,0.04), inset 0 1px 0 rgba(255,255,255,0.8);
  display: flex;
  flex-direction: column;
  min-height: 168px;
}
.ek-eyebrow {
  font-size: 10px;
  font-weight: 700;
  color: #475569;
  letter-spacing: 2.5px;
  text-transform: uppercase;
}
.ek-num {
  font-size: 52px;
  font-weight: 800;
  color: #0f172a;
  letter-spacing: -2px;
  font-variant-numeric: tabular-nums;
  margin: 14px 0 12px;
  line-height: 1.0;
  display: flex;
  align-items: baseline;
  gap: 4px;
}
.ek-unit {
  font-size: 17px;
  font-weight: 600;
  color: #94a3b8;
  letter-spacing: 0;
}
.ek-foot {
  font-size: 12px;
  color: #64748b;
  margin-top: auto;
}
.ek-pill {
  margin-top: 10px;
}

.exec-charts {
  display: grid;
  grid-template-columns: 1.15fr 1fr;
  gap: 36px;
  margin-bottom: 28px;
}
@media (max-width: 1000px) { .exec-charts { grid-template-columns: 1fr; } }
.exec-chart-block {}
.exec-section-eyebrow {
  font-size: 11px;
  font-weight: 700;
  color: #475569;
  letter-spacing: 2.5px;
  text-transform: uppercase;
  margin-bottom: 4px;
}
.exec-section-sub {
  font-size: 13px;
  color: #64748b;
  margin-bottom: 12px;
}
.exec-line-chart, .exec-bar-chart {
  width: 100%;
  height: auto;
  display: block;
}
.exec-bright {
  font-size: 13px;
  color: #334155;
  font-weight: 500;
  padding: 16px 0 0;
  border-top: 1px solid #e5e7eb;
  margin-top: 16px;
}
"""

EXEC_CSS = EXEC_CSS + """
.brand-logo { width: 100%; height: auto; display: block; margin-bottom: 4px; }
.brand { padding: 18px 20px 14px; }
.exec-eyebrow, .ek-eyebrow, .exec-section-eyebrow, .nav-section { color: #ad8c63 !important; }
"""
# Inject EXEC_CSS into CSS variable
CSS = CSS + EXEC_CSS

# Additional CSS for the edits round - appended into CSS variable
EDITS_CSS = """
/* Brand color overrides for office banners */
:root {
  --warn-bg: #fef3c7;
  --warn-text: #92400e;
  --good-bg: #d1f4dd;
  --good-text: #14532d;
}

.pill.warn { background: var(--warn-bg); color: var(--warn-text); }

/* Toggle */
.toggle {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  color: var(--text);
  cursor: pointer;
  padding: 6px 10px;
  border: 1px solid #cbd5e1;
  border-radius: 7px;
  background: var(--surface);
}
.toggle input { accent-color: #ad8c63; }

/* Delta indicators (triangle) */
.delta { font-weight: 700; font-variant-numeric: tabular-nums; font-size: 11px; }
.delta.pos { color: #16a34a; }
.delta.neg { color: #dc2626; }
.delta-muted { color: #cbd5e1; font-size: 11px; }
.col-delta { min-width: 56px; max-width: 56px; text-align: center; }

/* Future week dimming + collapse */
.bubble-grid .week-future { background: #fafafa !important; opacity: 0.6; }
.bubble-grid .week-future .pill.na { background: transparent; color: #cbd5e1; }
.week-cell.collapsed { display: none !important; }
.q-total { background: #fafbfc; border-left: 2px solid #cbd5e1; }
.q-delta { background: #fafbfc; }

/* PMO narrower first column */
.pmo-narrow .col-office { min-width: 110px; max-width: 110px; padding-left: 12px; font-size: 11px; }

/* Wrap narrow header */
th.wrap-narrow { white-space: normal !important; max-width: 110px; line-height: 1.2; }

/* Overhead section headers */
.hdr-group th { background: #f3f4f6 !important; font-size: 10px; text-transform: uppercase; letter-spacing: 1.5px; color: #475569; }

/* Proposal Hours: color-coded office rows */
tr.office-level td {
  background: #f5f1e8 !important;
  font-weight: 700;
  color: #382e30;
  border-top: 1px solid #d4c4a5;
}
tr.office-level td.sticky-l {
  background: #f5f1e8 !important;
  padding-left: 12px;
}

/* Non-Targeted grid uses same bubble styling */
.nt-grid .col-office { min-width: 120px; max-width: 120px; padding-left: 12px; }
.nt-grid .col-name { min-width: 200px; max-width: 200px; }
/* Non-Targeted has no Office column, so the base .bubble-grid sticky offsets (which assume
   one) leave a 130px hole and let Type scroll under Target.  Re-pin the four identity
   columns for this grid only — needed since it went to all 52 weeks (2026-09-17). */
.nt-grid .col-name { left: 0; padding-left: 16px; }
.nt-grid .col-type {
  position: sticky; left: var(--nt-l2, 200px); min-width: 46px; max-width: 56px;
  background: var(--surface); z-index: 4; text-align: center;
}
.nt-grid .col-target { left: var(--nt-l3, 246px); }
.nt-grid .ytd-sep {
  position: sticky; left: var(--nt-l4, 311px); min-width: 60px; max-width: 60px;
  background: var(--surface); z-index: 4;
}
/* Section banners: the cell spans every column, so the label span does the sticking. */
.nt-grid tbody tr.office-group td { padding-left: 0; }
.nt-grid .nt-band-label { position: sticky; left: 0; display: inline-block; padding-left: 14px; }
.nt-grid thead .col-type, .nt-grid thead .ytd-sep { background: #fafbfc; z-index: 15; }
.nt-grid tbody tr:not(.office-group):hover .col-type,
.nt-grid tbody tr:not(.office-group):hover .ytd-sep { background: #f3f5f7; }
.nt-scroll { max-height: calc(100vh - 200px); }
/* Non-Targeted Employees: section banners styled like Proposal Hours office banners */
.nt-grid tr.office-group td {
  background: #382e30 !important;
  color: #c4ad7a !important;
  font-weight: 700;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  font-size: 11px;
  padding: 10px 14px;
  border-top: 1px solid #2b2225;
  border-bottom: 1px solid #2b2225;
  border-left: 3px solid #ad8c63;
}


/* PT activity */
.pt-table th.sticky-l, .pt-table td.sticky-l { min-width: 180px; max-width: 180px; }
.pt-table td.cell { padding: 4px 6px; text-align: center; }

/* PTO tighter columns */
.pto-table th.wrap-narrow { max-width: 90px; }
.pto-table { width: auto; }
.pto-table th, .pto-table td { padding: 6px 10px; }
.pto-table th { font-weight: 600; }
/* Overhead: equal-width category columns and wrappable headers */
.ovhd-table { table-layout: fixed; width: 100%; }
.ovhd-table th.wrap-narrow { white-space: normal; word-wrap: break-word; min-width: 70px; }
.ovhd-table th.hours-col, .ovhd-table th.wrap-narrow, .ovhd-table td.hours-col { width: calc((100% - 360px) / 10); }
.ovhd-table th.sticky-l, .ovhd-table th[rowspan="2"] { width: auto; }


/* Goal planning narrow columns */
.goal-table th.wrap-narrow { max-width: 95px; }

/* Office group banner uses Earth Brown */
.bubble-grid tr.office-group td,
.nt-grid tr.office-group td {
  background: #382e30 !important;
  color: #f5f1e8 !important;
}
"""
CSS = CSS + EDITS_CSS


ROUND2_CSS = """
/* Toggle button */
.btn-toggle {
  padding: 7px 12px;
  border: 1px solid #cbd5e1;
  border-radius: 7px;
  background: #ffffff;
  font-size: 13px;
  color: #382e30;
  cursor: pointer;
  font-family: inherit;
  font-weight: 500;
}
.btn-toggle:hover { background: #f5f1e8; border-color: #ad8c63; }

/* Centered cells/headers for Goal Planning numeric cols */
.goal-table th.center, .goal-table td.center { text-align: center !important; }
.goal-table th.col-name-narrow, .goal-table td.col-name-narrow {
  min-width: 75px; max-width: 95px;
  white-space: normal;
  line-height: 1.25;
}
.goal-table th { vertical-align: middle; }

/* No-borders class for Proposal/Fleet tables */
.no-borders th, .no-borders td { border-bottom: none !important; }
.no-borders thead tr th { border-bottom: 1px solid #cbd5e1 !important; }
.no-borders tbody tr:nth-child(2n) td { background: transparent; }
.no-borders tr.office-level td { border-top: none; }

/* Future month collapse: pmo */
.pmo-grid .future-period.hidden { display: none !important; }

/* Bubble grid compact (Quarterly View): hide weekly cells; keep Q totals visible */
.bubble-grid .week-cell.collapsed { display: none !important; }
/* When in compact mode, the colspan on q-hdr needs to shrink to 2 (Q-total + Q-delta) */
.bubble-grid .q-hdr-compact { /* nothing visual changes; CSS only — the colspan is the trick */ }
/* In compact mode, force quarterly header colspan via attribute selector won't change colspan, so use JS or visual hack */
.bubble-grid.compact-mode .q-hdr { /* will be reset via JS */ }

/* Office banner has its own first-column sticky offset; ensure it stays */
.bubble-grid tr.office-group td { z-index: 2; }

/* When compact, q-total and q-delta should be sticky to keep them visible */
.bubble-grid .q-total, .bubble-grid .q-delta {
  background: #fafbfc;
}
.bubble-grid thead .q-total, .bubble-grid thead .q-delta {
  background: #fafbfc;
}

/* Delta cells (no triangles): make subtle, mono */
.delta-muted { color: #c5c5c5; font-size: 11px; font-variant-numeric: tabular-nums; }
.delta { font-weight: 600; font-variant-numeric: tabular-nums; font-size: 11px; letter-spacing: 0.2px; }
.delta.pos { color: #16a34a; }
.delta.neg { color: #dc2626; }

/* PT activity weekly columns alignment */
.pt-table td.num { font-variant-numeric: tabular-nums; }

/* Section-row text */
tr.section-row td { background: #f5f1e8 !important; color: #382e30; }

/* Round 2.1: center PMO Monthly headers */
.pmo-grid thead th { text-align: center !important; }
.pmo-grid thead tr.hdr-period th, .pmo-grid thead tr.hdr-sub th { text-align: center !important; }
"""
CSS = CSS + ROUND2_CSS

ROUND2_EXEC_CSS = """
.exec-tight .exec-display { font-size: 52px; margin-bottom: 4px; }
.exec-tight .exec-subtitle { margin-bottom: 20px; }
.exec-tight .exec-kpi-grid { gap: 12px; margin-bottom: 16px; }
.exec-tight .ek { padding: 12px 12px 14px; min-width: 0; }
.exec-tight .ek-eyebrow { font-size: 9px; letter-spacing: 1px; }
.exec-tight .ek-num { font-size: 32px; margin: 6px 0 4px; }
.exec-tight .ek-foot { font-size: 10px; }

.exec-tight .ek { padding: 14px 16px 14px; min-height: 148px; }
.exec-tight .ek-num { font-size: 44px; margin: 10px 0 8px; }
.exec-tight .exec-bright {
  padding: 12px 14px;
  margin: 0 0 16px;
  background: #f5f1e8;
  border: 1px solid #e2d8c2;
  border-radius: 8px;
  border-top: none;
  font-size: 13px;
  font-weight: 500;
  color: #382e30;
}
.exec-tight .exec-charts { gap: 28px; margin-bottom: 12px; }
.exec-tight .exec-line-chart, .exec-tight .exec-bar-chart { max-height: 280px; }
"""
CSS = CSS + ROUND2_EXEC_CSS


NEW_CSS = """
/* Hide the duplicate tab-name h2 in .page-intro (title bar already shows it) */
.page-intro h2 { display: none !important; }
.page-intro .page-sub { white-space: nowrap !important; overflow: visible !important; max-width: none !important; margin-top: 4px; }
.card-head .card-sub { display: none !important; }

/* center / right utility classes */
.data-table th.center, .data-table td.center { text-align: center !important; }
.data-table th.right, .data-table td.right { text-align: right !important; }

/* PMO Monthly: center all data cells too */
.pmo-grid thead th { text-align: center !important; }
.pmo-grid tbody td, .pmo-grid tbody td.num { text-align: center !important; }
.pmo-grid tbody td.col-office { text-align: left !important; }
.pmo-grid tbody td .delta, .pmo-grid tbody td .pill { display: inline-block; }
.pmo-grid .future-period.hidden { display: none !important; }
.pmo-grid.hide-future .future-period { display: none !important; }

/* Bright spots banner (Earth Brown) */
.exec-tight .exec-bright {
  padding: 18px 24px; margin: 16px 0; background: #382e30;
  border: 1px solid #2b2225; border-radius: 14px;
  font-size: 14px; color: #efe8d6; line-height: 1.55;
}
.exec-tight .exec-bright .exec-bright-eyebrow {
  font-size: 10.5px; font-weight: 700; letter-spacing: 0.18em;
  text-transform: uppercase; color: #c4ad7a; margin-bottom: 8px;
}
.exec-tight .exec-bright .exec-bright-body { color: #efe8d6; }
.exec-display { font-size: 64px; font-weight: 800; margin: 0 0 24px; padding-bottom: 6px; line-height: 1.05; }
.exec-subtitle { margin: 8px 0 28px; }

/* Proposal Hours: darker office banners (Earth Brown) */
tr.office-level td {
  background: #382e30 !important;
  font-weight: 600; color: #efe8d6;
  border-top: 1px solid #2b2225; border-bottom: 1px solid #2b2225;
  padding-top: 10px; padding-bottom: 10px;
}
tr.office-level td.sticky-l {
  background: #382e30 !important; color: #c4ad7a;
  padding-left: 14px; font-weight: 700;
  letter-spacing: 0.04em; text-transform: uppercase;
  font-size: 11px; border-left: 3px solid #ad8c63;
}

/* On Call Trend chart */
.oc-chart-card { background: #ffffff; border: 1px solid #d8d8d1; border-radius: 12px; padding: 12px 18px; overflow: hidden; }
.oc-chart { width: 100%; height: auto; display: block; max-height: 420px; }

/* Overtime sub no wrap; overhead column widths */
.ot-sub { white-space: normal !important; overflow: visible !important; padding-bottom: 0 !important; }
.ovhd-table th.pct-section-start, .ovhd-table td.pct-section-start,
.ovhd-table th.hours-section-start, .ovhd-table td.hours-section-start {
  border-left: 3px solid #382e30 !important;
}
.ovhd-table th.hours-col, .ovhd-table td.hours-col { width: 60px; min-width: 60px; max-width: 60px; }
.ovhd-table thead tr:first-child > th:nth-child(1), .ovhd-table tbody td.sticky-l { display: none !important; }
.ovhd-table thead tr:first-child > th:nth-child(3), .ovhd-table tbody tr > td:nth-child(3) {
  max-width: 140px; white-space: nowrap; overflow: hidden;
  text-overflow: ellipsis; font-size: 11px;
}

/* PT Activity column widths */
.pt-table th.pt-name, .pt-table td.pt-name { min-width: 140px; max-width: 140px; }
.pt-table th.pt-dept, .pt-table td.pt-dept { min-width: 90px; max-width: 90px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-size: 11px; }
.pt-table th.pt-pos, .pt-table td.pt-pos { min-width: 110px; max-width: 110px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-size: 11px; }
.pt-table th.pt-lts, .pt-table td.pt-lts { min-width: 86px; max-width: 86px; font-size: 11px; }
.pt-table th.pt-wk, .pt-table td.pt-wk { width: 50px; min-width: 50px; max-width: 50px; text-align: center !important; }
.pt-table th.pt-rollup, .pt-table td.pt-rollup { width: 50px; min-width: 50px; max-width: 50px; text-align: center !important; white-space: normal; line-height: 1.15; }
.pt-table th.pt-rollup-first, .pt-table td.pt-rollup-first { border-left: 2px solid #382e30 !important; }

/* Employee Util Type column */
.bubble-grid .col-type { min-width: 46px; max-width: 46px; text-align: center; font-size: 11px; font-weight: 600; color: #475569; }
.bubble-grid .ytd-sep.col-delta { border-left: 2px solid #cbd5e1 !important; border-right: 2px solid #cbd5e1 !important; }
.bubble-grid .q-delta.col-delta { border-right: 2px solid #cbd5e1 !important; }

/* Recovery Planning dividers */
.rc-table th.rc-sec-start, .rc-table td.rc-sec-start { border-left: 3px solid #382e30 !important; }

/* Employee Dashboard */
.ed-wrap { display: flex; flex-direction: column; gap: 14px; }
.ed-header { display: flex; align-items: center; gap: 16px; margin-bottom: 4px; }
.ed-picker { font-size: 13px; padding: 8px 12px; min-width: 280px; }
.ed-title { font-size: 22px; font-weight: 600; color: #0f172a; margin: 0; letter-spacing: -0.01em; }
.ed-kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; }
.ed-kpi { padding: 18px 20px; border-radius: 14px; color: #fff; box-shadow: 0 1px 2px rgba(0,0,0,0.06); }
.ed-kpi-blue  { background: linear-gradient(135deg, #382e30 0%, #4a3d40 100%); }
.ed-kpi-ochre { background: linear-gradient(135deg, #ad8c63 0%, #c4a679 100%); }
.ed-kpi-clay  { background: linear-gradient(135deg, #734d45 0%, #8b5e54 100%); }
.ed-kpi-moss  { background: linear-gradient(135deg, #585444 0%, #6c6753 100%); }
.ed-kpi-label { font-size: 10px; font-weight: 600; letter-spacing: 0.15em; text-transform: uppercase; opacity: 0.85; margin-bottom: 8px; }
.ed-kpi-val { font-size: 32px; font-weight: 700; letter-spacing: -0.02em; line-height: 1; }
.ed-kpi-val .ed-unit { font-size: 13px; font-weight: 400; opacity: 0.7; margin-left: 6px; }
.ed-mid { display: grid; grid-template-columns: 2fr 1fr; gap: 14px; }
.ed-card { background: #ffffff; border: 1px solid #d8d8d1; border-radius: 14px; padding: 16px 18px; }
.ed-card-eyebrow { font-size: 10px; font-weight: 700; letter-spacing: 0.15em; text-transform: uppercase; color: #ad8c63; margin-bottom: 10px; }
.ed-chart { width: 100%; height: auto; display: block; }
.ed-donut-legend { margin-top: 8px; display: flex; flex-direction: column; gap: 4px; }
.ed-legend-item { display: flex; align-items: center; gap: 8px; font-size: 12px; color: #0f172a; }
.ed-swatch { width: 12px; height: 12px; border-radius: 3px; }
.ed-legend-name { flex: 1; }
.ed-legend-val { color: #64748b; font-variant-numeric: tabular-nums; }
.ed-projects { display: grid; grid-template-columns: 1fr; gap: 8px; }
.ed-project { display: flex; align-items: center; gap: 12px; padding: 10px 12px; background: #fafaf7; border-radius: 10px; border: 1px solid #ecebe6; }
.ed-project-icon { font-size: 22px; width: 36px; height: 36px; display: flex; align-items: center; justify-content: center; background: #fff; border-radius: 8px; border: 1px solid #ecebe6; }
.ed-project-info { flex: 1; min-width: 0; }
.ed-project-name { font-size: 12px; color: #0f172a; font-weight: 500; line-height: 1.3; }
.ed-project-sector { font-size: 11px; color: #64748b; margin-top: 2px; }
.ed-empty { color: #94a3b8; font-size: 12px; font-style: italic; padding: 20px; text-align: center; }
.ed-caveat { font-size: 11px; color: #64748b; font-style: italic; padding: 4px 0; }

/* Employee Dashboard v2 - matches design template */
.ed2-app { background: linear-gradient(180deg, #e8f3ff 0%, #f5faff 60%, #fff 100%); margin: -32px -40px; padding: 28px 40px 48px; min-height: 100%; }
.ed2-hero { background: linear-gradient(135deg, #d2e7ff 0%, #e6f0ff 100%); border-radius: 18px; padding: 20px 28px; margin-bottom: 22px; display: flex; align-items: center; justify-content: space-between; gap: 24px; box-shadow: 0 1px 4px rgba(79,124,200,0.1); }
.ed2-hero-title-wrap { display: flex; align-items: center; gap: 14px; }
.ed2-hero-icon { font-size: 28px; }
.ed2-hero-title { margin: 0; color: #2f5fb9; font-weight: 700; font-size: 28px; letter-spacing: -0.01em; }
.ed2-hero-filters { display: flex; gap: 18px; align-items: center; }
.ed2-filter { display: flex; flex-direction: column; gap: 4px; font-size: 11px; color: #5a78aa; font-weight: 500; }
.ed2-filter span { letter-spacing: 0.05em; }
.ed2-select { background: rgba(255,255,255,0.7); border: 1px solid #b3d0f0; border-radius: 8px; padding: 6px 10px; font-size: 12px; color: #0f172a; font-family: inherit; min-width: 130px; }
.ed2-select:focus { outline: 2px solid #4f7cc8; border-color: #4f7cc8; }
.ed2-main { display: grid; grid-template-columns: 1fr 2fr 1.05fr; gap: 16px; margin-bottom: 18px; align-items: stretch; }
.ed2-col-left, .ed2-col-mid, .ed2-col-right { display: flex; flex-direction: column; gap: 16px; }
.ed2-mid-row { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
.ed2-mid-row > .ed2-panel { margin: 0; }
.ed2-col-right .ed2-tasks { flex: 1 1 auto; display: flex; flex-direction: column; }
.ed2-col-right .ed2-tasks-list { flex: 1; }
.ed2-mini { background: #fff; border-radius: 14px; padding: 16px 18px; box-shadow: 0 2px 6px rgba(15,23,42,0.05); border: 1px solid #e7eef8; }
.ed2-mini-head { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }
.ed2-mini-icon { color: #4f7cc8; font-size: 14px; }
.ed2-mini-label { font-size: 13px; color: #475569; flex: 1; }
.ed2-mini-val { font-size: 22px; font-weight: 700; color: #0f172a; font-variant-numeric: tabular-nums; }
.ed2-mini-spark { width: 100%; height: 40px; display: block; }
.ed2-kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }
.ed2-kpi { padding: 18px 18px; border-radius: 16px; color: #fff; box-shadow: 0 4px 14px rgba(79,124,200,0.18); position: relative; overflow: hidden; }
.ed2-k-blue1 { background: linear-gradient(135deg, #3b6bc0 0%, #5588d4 100%); }
.ed2-k-blue2 { background: linear-gradient(135deg, #6db5e8 0%, #8acdf5 100%); }
.ed2-k-green { background: linear-gradient(135deg, #4ec394 0%, #6bd6a9 100%); }
.ed2-k-orange { background: linear-gradient(135deg, #566374 0%, #7c8ba1 100%); }
.ed2-k-label { font-size: 11px; opacity: 0.92; font-weight: 600; line-height: 1.25; }
.ed2-k-val { font-size: 30px; font-weight: 800; margin-top: 12px; letter-spacing: -0.02em; line-height: 1; font-variant-numeric: tabular-nums; }
.ed2-k-unit { font-size: 14px; font-weight: 500; opacity: 0.85; margin-left: 4px; }
.ed2-panel { background: #fff; border-radius: 14px; padding: 16px 18px; box-shadow: 0 2px 6px rgba(15,23,42,0.05); border: 1px solid #e7eef8; }
.ed2-panel-head { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; }
.ed2-panel-icon { font-size: 14px; color: #4f7cc8; }
.ed2-panel-title { font-size: 13px; font-weight: 600; color: #2f5fb9; }
.ed2-trend-range { margin-left: auto; font-size: 11px; color: #64748b; font-weight: 500; }
.ed2-allocation-body { display: flex; align-items: center; gap: 14px; }
.ed2-donut { flex-shrink: 0; width: 170px; height: 170px; }
.ed2-donut-legend { flex: 1; display: flex; flex-direction: column; gap: 6px; min-width: 0; }
.ed2-leg { display: grid; grid-template-columns: 12px 1fr auto; align-items: center; gap: 8px; font-size: 12px; color: #0f172a; }
.ed2-leg-dot { width: 12px; height: 12px; border-radius: 3px; }
.ed2-leg-name { color: #475569; }
.ed2-leg-pct { font-weight: 700; color: #0f172a; font-variant-numeric: tabular-nums; }
.ed2-tasks-list { display: flex; flex-direction: column; gap: 8px; flex: 1; justify-content: space-between; min-height: 0; }
.ed2-task-row { display: flex; justify-content: space-between; gap: 8px; font-size: 12px; color: #0f172a; margin-bottom: 4px; }
.ed2-task-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ed2-task-hrs { color: #4f7cc8; font-weight: 600; font-variant-numeric: tabular-nums; flex-shrink: 0; }
.ed2-task-bar-bg { width: 100%; height: 6px; background: #eef2ff; border-radius: 3px; overflow: hidden; }
.ed2-task-bar-fill { height: 100%; background: linear-gradient(90deg, #4f7cc8 0%, #7eb8d6 100%); border-radius: 3px; }
.ed2-empty { color: #94a3b8; font-size: 12px; font-style: italic; padding: 16px; text-align: center; }
.ed2-trend { background: #fff; border-radius: 14px; padding: 12px 20px 10px; box-shadow: 0 2px 6px rgba(15,23,42,0.05); border: 1px solid #e7eef8; margin-bottom: 12px; }
.ed2-trend-chart { width: 100%; height: auto; display: block; max-height: 270px; }
.ed2-projects-panel { background: #fff; border-radius: 14px; padding: 16px 18px; box-shadow: 0 2px 6px rgba(15,23,42,0.05); border: 1px solid #e7eef8; }
.ed2-projects-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px 12px; }
.ed2-proj { display: flex; align-items: center; gap: 10px; padding: 8px 10px; background: #f8fbff; border-radius: 10px; border: 1px solid #e7eef8; min-width: 0; }
.ed2-proj-icon { font-size: 18px; width: 34px; height: 34px; display: flex; align-items: center; justify-content: center; background: #fff; border-radius: 8px; border: 1px solid #e7eef8; flex-shrink: 0; }
.ed2-proj-info { flex: 1; min-width: 0; }
.ed2-proj-name { font-size: 11px; color: #0f172a; font-weight: 600; line-height: 1.25; overflow: hidden; text-overflow: ellipsis; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
.ed2-proj-sector { font-size: 10px; color: #64748b; margin-top: 2px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ed2-caveat { font-size: 11px; color: #64748b; font-style: italic; padding: 4px 0; }

/* Employee Dashboard: hide the duplicate "Employee Dashboard" banner (topbar) on this tab only */
.main:has(> .tab-host > #Employee_Dashboard.active) > .topbar { display: none !important; }
.main:has(> .tab-host > #Employee_Dashboard.active) > .tab-host { padding-top: 0 !important; }
#Employee_Dashboard .ed2-app { margin-top: 0 !important; }

/* PMO Dashboard (clone of Employee Dashboard v2) - matches design template */
.ed3-app { background: linear-gradient(180deg, #e8f3ff 0%, #f5faff 60%, #fff 100%); margin: -32px -40px; padding: 28px 40px 48px; min-height: 100%; }
.ed3-hero { background: linear-gradient(135deg, #d2e7ff 0%, #e6f0ff 100%); border-radius: 18px; padding: 20px 28px; margin-bottom: 22px; display: flex; align-items: center; justify-content: space-between; gap: 24px; box-shadow: 0 1px 4px rgba(79,124,200,0.1); }
.ed3-hero-title-wrap { display: flex; align-items: center; gap: 14px; }
.ed3-hero-icon { font-size: 28px; }
.ed3-hero-title { margin: 0; color: #2f5fb9; font-weight: 700; font-size: 28px; letter-spacing: -0.01em; }
.ed3-hero-filters { display: flex; gap: 18px; align-items: center; }
.ed3-filter { display: flex; flex-direction: column; gap: 4px; font-size: 11px; color: #5a78aa; font-weight: 500; }
.ed3-filter span { letter-spacing: 0.05em; }
.ed3-select { background: rgba(255,255,255,0.7); border: 1px solid #b3d0f0; border-radius: 8px; padding: 6px 10px; font-size: 12px; color: #0f172a; font-family: inherit; min-width: 130px; }
.ed3-select:focus { outline: 2px solid #4f7cc8; border-color: #4f7cc8; }
.ed3-main { display: grid; grid-template-columns: 1fr 2fr 1.05fr; gap: 16px; margin-bottom: 18px; align-items: stretch; }
.ed3-col-left, .ed3-col-mid, .ed3-col-right { display: flex; flex-direction: column; gap: 16px; }
.ed3-mid-row { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
.ed3-mid-row > .ed3-panel { margin: 0; }
.ed3-col-right .ed3-tasks { flex: 1 1 auto; display: flex; flex-direction: column; }
.ed3-col-right .ed3-tasks-list { flex: 1; }
.ed3-mini { background: #fff; border-radius: 14px; padding: 16px 18px; box-shadow: 0 2px 6px rgba(15,23,42,0.05); border: 1px solid #e7eef8; }
.ed3-mini-head { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }
.ed3-mini-icon { color: #4f7cc8; font-size: 14px; }
.ed3-mini-label { font-size: 13px; color: #475569; flex: 1; }
.ed3-mini-val { font-size: 22px; font-weight: 700; color: #0f172a; font-variant-numeric: tabular-nums; }
.ed3-mini-spark { width: 100%; height: 40px; display: block; }
.ed3-kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }
.ed3-kpi { padding: 18px 18px; border-radius: 16px; color: #fff; box-shadow: 0 4px 14px rgba(79,124,200,0.18); position: relative; overflow: hidden; }
.ed3-k-blue1 { background: linear-gradient(135deg, #3b6bc0 0%, #5588d4 100%); }
.ed3-k-blue2 { background: linear-gradient(135deg, #6db5e8 0%, #8acdf5 100%); }
.ed3-k-green { background: linear-gradient(135deg, #4ec394 0%, #6bd6a9 100%); }
.ed3-k-orange { background: linear-gradient(135deg, #566374 0%, #7c8ba1 100%); }
.ed3-k-label { font-size: 11px; opacity: 0.92; font-weight: 600; line-height: 1.25; }
.ed3-k-val { font-size: 30px; font-weight: 800; margin-top: 12px; letter-spacing: -0.02em; line-height: 1; font-variant-numeric: tabular-nums; }
.ed3-k-unit { font-size: 14px; font-weight: 500; opacity: 0.85; margin-left: 4px; }
.ed3-panel { background: #fff; border-radius: 14px; padding: 16px 18px; box-shadow: 0 2px 6px rgba(15,23,42,0.05); border: 1px solid #e7eef8; }
.ed3-panel-head { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; }
.ed3-panel-icon { font-size: 14px; color: #4f7cc8; }
.ed3-panel-title { font-size: 13px; font-weight: 600; color: #2f5fb9; }
.ed3-trend-range { margin-left: auto; font-size: 11px; color: #64748b; font-weight: 500; }
.ed3-allocation-body { display: flex; align-items: center; gap: 14px; }
.ed3-donut { flex-shrink: 0; width: 170px; height: 170px; }
.ed3-donut-legend { flex: 1; display: flex; flex-direction: column; gap: 6px; min-width: 0; }
.ed3-leg { display: grid; grid-template-columns: 12px 1fr auto; align-items: center; gap: 8px; font-size: 12px; color: #0f172a; }
.ed3-leg-dot { width: 12px; height: 12px; border-radius: 3px; }
.ed3-leg-name { color: #475569; }
.ed3-leg-pct { font-weight: 700; color: #0f172a; font-variant-numeric: tabular-nums; }
.ed3-tasks-list { display: flex; flex-direction: column; gap: 8px; flex: 1; justify-content: space-between; min-height: 0; }
.ed3-task-row { display: flex; justify-content: space-between; gap: 8px; font-size: 12px; color: #0f172a; margin-bottom: 4px; }
.ed3-task-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ed3-task-hrs { color: #4f7cc8; font-weight: 600; font-variant-numeric: tabular-nums; flex-shrink: 0; }
.ed3-task-bar-bg { width: 100%; height: 6px; background: #eef2ff; border-radius: 3px; overflow: hidden; }
.ed3-task-bar-fill { height: 100%; background: linear-gradient(90deg, #4f7cc8 0%, #7eb8d6 100%); border-radius: 3px; }
.ed3-empty { color: #94a3b8; font-size: 12px; font-style: italic; padding: 16px; text-align: center; }
.ed3-trend { background: #fff; border-radius: 14px; padding: 12px 20px 10px; box-shadow: 0 2px 6px rgba(15,23,42,0.05); border: 1px solid #e7eef8; margin-bottom: 12px; }
.ed3-trend-chart { width: 100%; height: auto; display: block; max-height: 270px; }
.ed3-projects-panel { background: #fff; border-radius: 14px; padding: 16px 18px; box-shadow: 0 2px 6px rgba(15,23,42,0.05); border: 1px solid #e7eef8; }
.ed3-projects-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px 12px; }
.ed3-proj { display: flex; align-items: center; gap: 10px; padding: 8px 10px; background: #f8fbff; border-radius: 10px; border: 1px solid #e7eef8; min-width: 0; }
.ed3-proj-icon { font-size: 18px; width: 34px; height: 34px; display: flex; align-items: center; justify-content: center; background: #fff; border-radius: 8px; border: 1px solid #e7eef8; flex-shrink: 0; }
.ed3-proj-info { flex: 1; min-width: 0; }
.ed3-proj-name { font-size: 11px; color: #0f172a; font-weight: 600; line-height: 1.25; overflow: hidden; text-overflow: ellipsis; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
.ed3-proj-sector { font-size: 10px; color: #64748b; margin-top: 2px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ed3-caveat { font-size: 11px; color: #64748b; font-style: italic; padding: 4px 0; }

"""
CSS = CSS + NEW_CSS

if __name__ == '__main__':
    build()
