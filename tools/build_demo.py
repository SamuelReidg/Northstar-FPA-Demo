#!/usr/bin/env python3
"""Rebuild the whole Northstar Heritage demo from scratch.

    python3 tools/build_demo.py                 # regenerate sources + run every refresh pipeline
    python3 tools/build_demo.py --rebuild-world # also rebuild the synthetic world from its seed
    python3 tools/build_demo.py --only budget   # one dashboard (utilization|budget|approvals|ews|pm_hours|proposals|revenue|location)

Stage 1 writes the synthetic source exports into demo_data/ (the weekly package a finance team
would drop). Stage 2 runs each dashboard's real refresh pipeline from pipeline/ in a scratch
folder (work/<name>/), exactly as the production workflow does, and copies the outputs into site/.
All pipelines run with FPA_AS_OF=2026-09-26 so the demo is reproducible whatever today's date is.
"""
import argparse
import os
import shutil
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS, PIPE, SITE = (os.path.join(ROOT, d) for d in ('tools', 'pipeline', 'site'))
PKG = os.path.join(ROOT, 'demo_data', 'Weekly and Monthly Updates', '9.29.2026')
PREV = os.path.join(ROOT, 'demo_data', 'Weekly and Monthly Updates', '9.22.2026')
ENV = dict(os.environ, FPA_AS_OF='2026-09-26', PYTHONHASHSEED='0')
PY = sys.executable


def run(args, cwd=ROOT):
    t = time.time()
    print('  $', ' '.join(a if ' ' not in a else repr(a) for a in args), flush=True)
    subprocess.run(args, cwd=cwd, env=ENV, check=True)
    print('    (%.1fs)' % (time.time() - t), flush=True)


def work(name, scripts=(), files=()):
    d = os.path.join(ROOT, 'work', name)
    shutil.rmtree(d, ignore_errors=True)
    os.makedirs(d)
    for s in scripts:
        shutil.copy(os.path.join(PIPE, s), d)
    for f in files:
        shutil.copy(f, d)
    return d


def to_site(d, *names):
    for n in names:
        shutil.copy(os.path.join(d, n), os.path.join(SITE, n))


def utilization():
    run([PY, 'tools/gen_utilization.py'])
    u = os.path.join(PKG, 'Utilization')
    d = work('util', ['extract_data.py', 'regenerate_proposal_pivot.py', 'build_pmo_dashboard.py', 'build_site.py',
                      'merge.py', 'apply_categories.py', 'sync_util_data.py', 'regenerate_pmo_extra.py',
                      'regenerate_emp_extra.py', 'regenerate_pmo_reduced.py', 'regenerate_raw_data.py',
                      'check_export_freshness.py', 'logo_b64.txt'],
             [os.path.join(u, 'Utilization Report.xlsx'), os.path.join(u, 'Projects.csv')])
    shutil.copy(os.path.join(SITE, 'index.html'), os.path.join(d, 'LIVE_index.html'))   # merge base = current design
    X = 'Utilization Report.xlsx'
    run([PY, 'extract_data.py', X], d)
    run([PY, 'regenerate_proposal_pivot.py', '--workbook', X, '--out', 'proposal_pivot.js', '--site-data', 'site_data'], d)
    run([PY, 'build_pmo_dashboard.py'], d)
    run([PY, 'build_site.py'], d)
    run([PY, 'merge.py'], d)
    run([PY, 'apply_categories.py', X, 'MERGED.html', 'MERGED.html'], d)
    shutil.copy(os.path.join(d, 'MERGED.html'), os.path.join(d, 'index.html'))
    run([PY, 'sync_util_data.py'], d)
    run([PY, 'regenerate_pmo_extra.py', X, 'util_data.js', 'ed3_extra.js'], d)
    run([PY, 'regenerate_emp_extra.py', 'MERGED.html', X, 'ed3_extra.js'], d)
    run([PY, 'regenerate_pmo_reduced.py', '--workbook', X, '--index', 'MERGED.html'], d)
    run([PY, 'regenerate_raw_data.py', '--workbook', X, '--out', 'raw_data.js'], d)
    shutil.copy(os.path.join(d, 'MERGED.html'), os.path.join(SITE, 'index.html'))
    to_site(d, 'ed3_extra.js', 'util_data.js', 'raw_data.js', 'proposal_pivot.js')
    run([PY, os.path.join(d, 'check_export_freshness.py'), '--dir', SITE])


def budget():
    run([PY, 'tools/gen_budget.py'])
    run([PY, 'tools/gen_budget.py', '--check'])
    d = work('budget', ['build_overrun_dashboard.py', 'regenerate_budget_tabs.py', 'weekly_summary.py', 'build_summary.js'])
    run([PY, 'build_overrun_dashboard.py',
         os.path.join(PKG, 'Budget Watchtower', 'FP&A & QC Project Report (41).xlsx'),
         os.path.join(PREV, 'Budget Watchtower', 'FP&A & QC Project Report (40).xlsx')], d)
    to_site(d, 'budget-overrun-summary.html', 'budget-watchtower-workbook.xlsx')


def approvals():
    run([PY, 'tools/gen_approvals.py'])
    d = work('approvals', ['build_dashboards.py', 'build_dashboard_html.py', 'dashboard_template.html'])
    src = os.path.join(PKG, 'Timesheets Draft & Submitted')
    run([PY, 'build_dashboard_html.py'] + [os.path.join(src, f) for f in sorted(os.listdir(src))], d)
    to_site(d, 'approvals-dashboard.html')


def ews():
    run([PY, 'tools/gen_ews.py'])
    x = os.path.join(PKG, 'Early Warning System', 'SR of Projects @ Risk - Automated 9.29.26.xlsx')
    d = work('ews', ['regenerate_ews_data.py', 'regenerate_project_tasks.py'], [x])
    run([PY, 'regenerate_ews_data.py', os.path.basename(x)], d)
    run([PY, 'regenerate_project_tasks.py', os.path.basename(x), 'project_tasks.js'], d)
    run([PY, 'tools/apply_ews_blocks.py', d, SITE])
    to_site(d, 'project_tasks.js')


def pm_hours():
    run([PY, 'tools/gen_pm_hours.py'])
    d = work('pm_hours', ['regenerate_pm_hours_data.py'], [os.path.join(PKG, 'PM Hours & Expenses', 'CM PM Report.xlsx')])
    run([PY, 'regenerate_pm_hours_data.py', 'CM PM Report.xlsx'], d)
    to_site(d, 'pm_hours_data.js')


def proposals():
    run([PY, 'tools/gen_proposals.py'])
    d = work('proposals', ['regenerate_proposal_data.py'])
    run([PY, 'regenerate_proposal_data.py', os.path.join(PKG, 'Proposal Hours Summary', 'Proposal Hours Summary_9.2026_.xlsx')], d)
    to_site(d, 'proposal_data.js')


def revenue():
    run([PY, 'tools/gen_revenue_forecast.py'])
    r = os.path.join(PKG, 'Revenue Forecast')
    d = work('revenue', ['regenerate_revenue_forecast_data.py'])
    run([PY, 'regenerate_revenue_forecast_data.py', os.path.join(r, 'Revenue Forecast.xlsx'), 'revenue_forecast_data.js',
         os.path.join(r, 'revenue_forecast_entries.csv')], d)
    to_site(d, 'revenue_forecast_data.js')


def location():
    run([PY, 'tools/gen_location_drivers.py'])


STEPS = {'utilization': utilization, 'budget': budget, 'approvals': approvals, 'ews': ews, 'pm_hours': pm_hours,
         'proposals': proposals, 'revenue': revenue, 'location': location}


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('--rebuild-world', action='store_true')
    ap.add_argument('--only', choices=sorted(STEPS), action='append')
    a = ap.parse_args()
    if a.rebuild_world or not os.path.exists(os.path.join(TOOLS, '.cache', 'world.pkl')):
        run([PY, 'tools/demo_world.py'])
    # revenue before utilization (the hub's backlog tab reuses the revenue export rows)
    order = ['revenue', 'utilization', 'budget', 'approvals', 'ews', 'pm_hours', 'proposals', 'location']
    for name in order:
        if a.only and name not in a.only:
            continue
        print('\n== %s' % name, flush=True)
        STEPS[name]()
    print('\nDone. Open site/index.html in a browser (or serve the folder: python3 -m http.server -d site).')


if __name__ == '__main__':
    main()
