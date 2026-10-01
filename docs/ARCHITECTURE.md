# Architecture

## Design goals

- **No server, no database.** Every dashboard is a self-contained HTML page that opens from a file
  share or any static host. Data ships as generated `*.js` files (`window.X = {...}`) or JSON
  embedded in the page.
- **Exports in, pages out.** Finance drops a dated package of ERP/CRM exports each week; a Python
  pipeline per dashboard turns them into the page's data. Nothing is typed in by hand.
- **Design is separate from data.** The hub's look was built by hand over time; the pipeline never
  overwrites it (see *Merge refresh*).
- **Fail loudly.** Export layouts drift (columns move, sheets get renamed, a tab stops refreshing).
  Every pipeline locates columns by header, sniffs files by content, and runs checks that abort
  rather than write quietly wrong numbers.

## Components

```
demo_data/Weekly and Monthly Updates/<M.D.YYYY>/
    Utilization/            Utilization Report.xlsx, Projects.csv
    Budget Watchtower/      FP&A & QC Project Report (NN).xlsx  (+ E3000 companion, in SAR)
    Timesheets Draft & Submitted/   Timesheets <n>.csv, Employees <n>.xlsx, Staff expenses <n>.xlsx
    Early Warning System/   SR of Projects @ Risk - Automated <date>.xlsx
    PM Hours & Expenses/    CM PM Report.xlsx
    Proposal Hours Summary/ Proposal Hours Summary_<month>.xlsx
    Revenue Forecast/       Revenue Forecast.xlsx

pipeline/                   → site/
  extract_data.py …          index.html, ed3_extra.js, util_data.js, raw_data.js, proposal_pivot.js
  build_overrun_dashboard.py budget-overrun-summary.html, budget-watchtower-workbook.xlsx
  build_dashboard_html.py    approvals-dashboard.html
  regenerate_ews_data.py …   early-warning-system.html (data blocks), project-recovery.html, project_tasks.js
  regenerate_pm_hours_data.py   pm_hours_data.js
  regenerate_proposal_data.py   proposal_data.js
  regenerate_revenue_forecast_data.py   revenue_forecast_data.js
  (tools/gen_location_drivers.py)       drivers_data.js — see note below
```

## Merge refresh (the hub)

`build_site.py` renders every tab from the workbook, but into an older page template. Instead of
shipping that, `merge.py` takes the **current live** `index.html` and:

1. replaces the body of each refreshable `<section id="…">` tab with the freshly rendered one;
2. for the PMO and Employee dashboards (whose layout is hand-built), keeps the live section and swaps
   only the embedded JSON payloads (`ed3Data`, `ed2Data`);
3. keeps retired tabs whose source sheets no longer exist;
4. leaves everything outside the sections (homepage, sidebar, head styles/scripts, the Excel
   exporter and its wrappers) byte-for-byte.

Post-merge scripts then patch the payloads: `apply_categories.py` (Admin/PTO/S&M split and Top-6
projects from raw timesheets), `regenerate_emp_extra.py` (role- and trend-aware observations for every
employee), `regenerate_pmo_extra.py` (office backlog panels, positive-backlog basis with a hard
reconciliation), `regenerate_pmo_reduced.py` (billable-adjusted rows for the export).

## Excel export

The hub's **Export to Excel** builds a styled workbook in the browser from the rendered tables
(`xlsx-js-style`). Three layers wrap the exporter:

- reduced (billable-adjusted) rows injected into two sheets just before the workbook is built,
  then removed — the on-screen tables never change;
- the on-screen Billable rows toggle on PMO Weekly;
- an async gate that lazy-loads `raw_data.js` (Raw Data sheet, ~80k rows) and `proposal_pivot.js`
  only when an export runs, and a freshness check that warns if either lags the page.

`data-val` / `data-fmt` attributes carry exact values so the workbook gets full precision while the
page shows rounded pills.

## Guards worth knowing

| Guard | Where | What it stops |
|---|---|---|
| Cell-for-cell self-check of weekly/monthly utilization and quarterly = Σ months | `regenerate_pmo_reduced.py` | writing reduced figures when the workbook layout moved |
| Backlog rows must tie to headline | `regenerate_pmo_extra.py` | a truncated project list under a correct total |
| Side-car freshness | `check_export_freshness.py` + in-page banner | an export silently shipping last week's raw data |
| Loader hard-fails on missing/non-numeric REV REC % / Budget Spent % | `regenerate_budget_tabs.py` | recomputing system values that set the 70%/100% boundaries |
| Files identified by content, columns by header | approvals, budget, proposals | drifting export names and column order |
| As-of date check | budget | rebuilding from an older export than the live page |

## Location Scorecard note

In production its `drivers_data.js` came from a separately maintained workbook; the demo builds it
with `tools/gen_location_drivers.py`, which also regenerates the page's embedded monthly snapshot.
The page also reads `util_data.js` from the utilization pipeline.
