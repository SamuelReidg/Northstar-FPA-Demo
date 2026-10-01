#!/usr/bin/env python3
"""
Weekly generator for the Approvals Dashboard.

Usage:  python build_dashboard_html.py <Timesheets.csv> <Employees.csv> <Staff expenses.xlsx>
        (the three files may be passed in any order -- they are identified by content)

Reads the three weekly exports, resolves supervisor / PMO from the employee master,
computes all dashboard aggregates, and writes a self-contained approvals-dashboard.html
with the data embedded. No Excel formulas, macros, or pivot refresh required.
"""
import json, datetime, sys, os

HERE = os.path.dirname(os.path.abspath(__file__))
exec(open(os.path.join(HERE, 'build_dashboards.py')).read().split("if __name__")[0])

if len(sys.argv) < 4:
    raise SystemExit(
        'usage: python build_dashboard_html.py <Timesheets.csv> <Employees.csv> '
        '<Staff expenses.xlsx>\n'
        '       The three files can be given in any order.\n'
        "       This week's exports live in:\n"
        '       demo_data/Weekly and Monthly Updates/<M.D.YYYY>/'
        'Timesheets Draft & Submitted/\n'
        '       dashboard_template.html must sit next to this script.')

TS_FILE, EMP_FILE, EX_FILE = resolve_sources(*sys.argv[1:4])


def dstr(d):
    return d.strftime('%Y-%m-%d') if isinstance(d, datetime.datetime) else None


# ---- People to exclude from this dashboard entirely (never shown, never counted).
# Match is case-insensitive against supervisor AND employee name, in "Last, First"
# form, after the ERP's " - Employee" suffix is stripped.
#   spence, regina   -- Chief Executive Officer. Every corporate / sales / admin person
#                       reports to her, so excluding her also keeps their records off the
#                       page. Before 2026-08-20 the same people were kept out by the
#                       Target != 0 rule, which the new sources can no longer support --
#                       they carry no target column. This name is its replacement.
#   joyce, brett     -- Controller
#   vaughan, holly   -- Director of Finance
# Both finance names are belt and braces: they report to the CEO today, but matching the
# name in EITHER column keeps them off the page if their reporting line ever changes.
# Demo package 9.29.2026: the rule removes 6 timesheets (corporate staff, all reporting
# to the CEO) and 0 expense reports.
EXCLUDE_NAMES = {'spence, regina', 'joyce, brett', 'vaughan, holly'}

# Belt and braces: exclude by Employee ID too, so a respelling of a name in either
# export can't put someone back on the page. The CEO has no supervisor of her own,
# so before names were matched on BOTH columns her records resurfaced under
# "(Unassigned)" instead of disappearing. Matches the coworker's independent
# refresh script, which excludes the same ID -- keep the two in step.
EXCLUDE_EMPLOYEE_IDS = {'E31223'}    # Spence, Regina (CEO)

def _norm(n):
    return (clean_name(n) or '').lower()

def _excluded(row):
    return str(row.get('emp_id') or '').strip() in EXCLUDE_EMPLOYEE_IDS or \
           _norm(row.get('supervisor')) in EXCLUDE_NAMES or \
           _norm(row.get('emp_name')) in EXCLUDE_NAMES


def build_data():
    emps = build_employees(EMP_FILE)
    # Expenses first: its Last Modified is the only true timestamp in either export,
    # so it sets the as_of clock that both sides are aged against.
    ex = build_expenses(EX_FILE, emps)
    ts = build_timesheets(TS_FILE, emps, as_of=ex.get('as_of'))
    as_of = ex.get('as_of') or ts.get('as_of')
    as_of_d = as_of
    as_of = datetime.datetime(as_of.year, as_of.month, as_of.day) if as_of else None

    ex['kept'] = [x for x in ex['kept'] if not _excluded(x)]
    ts['kept'] = [x for x in ts['kept'] if not _excluded(x)]

    exp_by_sup = {}
    for x in ex['kept']:
        s = x['supervisor'] or '(Unassigned)'
        d = exp_by_sup.setdefault(s, dict(draft_n=0, sub_n=0, draft_amt=0.0, sub_amt=0.0))
        amt = x['amount'] if isinstance(x['amount'], (int, float)) else 0
        if x['state'] == 'Draft':
            d['draft_n'] += 1; d['draft_amt'] += amt
        else:
            d['sub_n'] += 1; d['sub_amt'] += amt

    def exp_row(x):
        return dict(er=x['er'], employee=x['emp_name'], supervisor=x['supervisor'] or '(Unassigned)',
                    state=x['state'], amount=round(x['amount'], 2) if isinstance(x['amount'], (int, float)) else 0,
                    date_filed=dstr(x['date_filed']), gl=dstr(x['gl']), reason=x['reason'],
                    last_modified=dstr(x.get('last_modified')),
                    location=x.get('location'),
                    not_employed=bool(x['not_employed']) if x['not_employed'] is not None else False)

    ts_by_sup = {}
    for x in ts['kept']:
        s = x['supervisor'] or '(Unassigned)'
        d = ts_by_sup.setdefault(s, dict(draft_n=0, sub_n=0))
        if x['state'] == 'Draft':
            d['draft_n'] += 1
        else:
            d['sub_n'] += 1

    def ts_row(x):
        return dict(employee=x['emp_name'], supervisor=x['supervisor'] or '(Unassigned)',
                    state=x['state'], begin=dstr(x['begin']), end=dstr(x['end']), gl=dstr(x['gl']),
                    location=x.get('location'))

    def avg(lst):
        lst = [a for a in lst if a is not None]
        return round(sum(lst) / len(lst), 1) if lst else None

    people = set([str(x['emp_id']) for x in ex['kept']] + [str(x['emp_id']) for x in ts['kept']])

    return dict(
        as_of=dstr(as_of),
        # Everything on the page has been outstanding MORE THAN min_days_outstanding
        # days: dated before outstanding_before. See build_dashboards.py.
        min_days_outstanding=MIN_DAYS_OUTSTANDING,
        outstanding_before=str(ex.get('cutoff') or ts.get('cutoff') or ''),
        age_basis=dict(timesheets='/'.join(AGE_KEYS_TS), expenses='/'.join(AGE_KEYS_EX)),
        expenses=dict(
            draft=sum(1 for x in ex['kept'] if x['state'] == 'Draft'),
            submitted=sum(1 for x in ex['kept'] if x['state'] == 'Submitted'),
            total=len(ex['kept']),
            total_amount=round(sum(x['amount'] for x in ex['kept'] if isinstance(x['amount'], (int, float))), 2),
            # Age uses the SAME field the filter uses, so the KPI can never
            # contradict the row set the way it did on 08-18.
            avg_age=avg([days_outstanding(x, as_of_d, AGE_KEYS_EX) for x in ex['kept']]),
            by_supervisor=[dict(supervisor=s, draft_n=d['draft_n'], sub_n=d['sub_n'],
                                draft_amt=round(d['draft_amt'], 2), sub_amt=round(d['sub_amt'], 2),
                                total_n=d['draft_n'] + d['sub_n'],
                                total_amt=round(d['draft_amt'] + d['sub_amt'], 2))
                           for s, d in sorted(exp_by_sup.items(), key=lambda kv: -(kv[1]['draft_n'] + kv[1]['sub_n']))],
            details=[exp_row(x) for x in ex['kept']],
        ),
        timesheets=dict(
            draft=sum(1 for x in ts['kept'] if x['state'] == 'Draft'),
            submitted=sum(1 for x in ts['kept'] if x['state'] == 'Submitted'),
            total=len(ts['kept']),
            avg_age=avg([days_outstanding(x, as_of_d, AGE_KEYS_TS) for x in ts['kept']]),
            by_supervisor=[dict(supervisor=s, draft_n=d['draft_n'], sub_n=d['sub_n'],
                                total_n=d['draft_n'] + d['sub_n'])
                           for s, d in sorted(ts_by_sup.items(), key=lambda kv: -(kv[1]['draft_n'] + kv[1]['sub_n']))],
            details=[ts_row(x) for x in ts['kept']],
        ),
        employees_affected=len(people),
        # Rows whose Employee ID is missing from the employee master would show with no
        # supervisor and no PMO. Surfaced here so a broken master can't pass silently.
        unresolved_employees=sorted({str(x['emp_id']) for x in ex['kept'] + ts['kept']
                                     if not x.get('in_master')}),
    )


data = build_data()

_tpl = os.path.join(HERE, 'dashboard_template.html')
if not os.path.exists(_tpl):
    raise SystemExit('ERROR: dashboard_template.html not found next to this script (%s).\n'
                     '       Copy all three build files together: build_dashboard_html.py, '
                     'build_dashboards.py, dashboard_template.html.' % HERE)
TEMPLATE = open(_tpl).read()
html = TEMPLATE.replace('/*__DATA__*/null', json.dumps(data))
out = os.path.join(HERE, 'approvals-dashboard.html')
with open(out, 'w') as f:
    f.write(html)
print('Wrote', out, '(%.0f KB)' % (len(html) / 1024))
print('as_of', data['as_of'], '| EXP', data['expenses']['draft'], data['expenses']['submitted'],
      data['expenses']['total'], '| TS', data['timesheets']['draft'], data['timesheets']['submitted'],
      data['timesheets']['total'], '| affected', data['employees_affected'])
if data['unresolved_employees']:
    print('WARNING: Employee ID(s) missing from the employee master (no supervisor/PMO):',
          ', '.join(data['unresolved_employees'][:10]))
