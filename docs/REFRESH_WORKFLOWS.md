# Refresh workflows

Each dashboard has its own source export and pipeline. `tools/build_demo.py` runs all of them in
scratch folders (`work/<name>/`) exactly as below; this page documents what each step does and what
to check. All steps assume `FPA_AS_OF=2026-09-26` (omit it in production to use today's date).

## Utilization Report hub — weekly

Input: `Utilization/Utilization Report.xlsx` (+ `Projects.csv` for project sectors). Copy the scripts
and the workbook into a working folder and the **current** `site/index.html` in as `LIVE_index.html`.

| # | Step | Purpose |
|---|---|---|
| 1 | `extract_data.py "Utilization Report.xlsx"` | every sheet → `site_data/*.json`, employee grid, month-to-date billable hours |
| 2 | `regenerate_proposal_pivot.py --workbook … --out proposal_pivot.js --site-data site_data` | rebuilds the proposal-hours pivot from Raw Data (grouped by the person's home office) — must run before step 4 |
| 3 | `build_pmo_dashboard.py` | office-level dashboard JSON from the Employee Dashboard sheet |
| 4 | `build_site.py` | renders all 16 tabs (data donor only) |
| 5 | `merge.py` | splices fresh tab bodies + dashboard payloads into the live design |
| 6 | `apply_categories.py … MERGED.html MERGED.html` | Admin / PTO / S&M split and Top-6 projects from raw timesheets |
| 7 | `cp MERGED.html index.html && sync_util_data.py` | `util_data.js` for the Location Scorecard (reads `index.html` beside it) |
| 8 | `regenerate_pmo_extra.py … util_data.js ed3_extra.js` | backlog + key observations per PMO; **must print `reconciliation: OK`** |
| 9 | `regenerate_emp_extra.py MERGED.html … ed3_extra.js` | observations for every employee, embedded in the page |
| 10 | `regenerate_pmo_reduced.py --workbook … --index MERGED.html` | billable-adjusted export rows; **must print three `Self-check ok` lines** |
| 11 | `regenerate_raw_data.py --workbook … --out raw_data.js` | Raw Data sheet for the export |
| 12 | deploy `MERGED.html`→`index.html`, `ed3_extra.js`, `util_data.js`, `raw_data.js`, `proposal_pivot.js` | |
| 13 | `check_export_freshness.py --dir site` | **must print `OK`** — catches skipped steps 2/11 |

Business rules kept from production: non-posting hours are matched exactly (`Overhead Non-Posting`),
so regional variants are not stripped; quarter columns come from the workbook's own quarterly blocks,
not a sum of weeks; negative non-posting hours (timesheet transfers) are left as-is; a small office can
be kept off the monthly table with `PMO_MONTHLY_HIDE_OFFICES` (a generator constant, because the Excel
export reads the rendered tables); two International staff are shown under "International (Global
LLC)" via `EMPLOYEE_OFFICE_OVERRIDES`.

## Budget Watchtower — weekly

Inputs: this week's `FP&A & QC Project Report (NN).xlsx`, its `…E3000.xlsx` companion (a second entity
reporting in SAR, same "As of" date), and last week's report for deltas.

```
python3 build_overrun_dashboard.py "<this week>.xlsx" "<last week>.xlsx"
```

Rows that count: active statuses only (In Progress, Final Reporting/Review, Lab/Curation, Cost
Tracking on Closed Project, Transferred Entity); a commercial basis (contract > 0 and a budget); OVHD
rows excluded. REV REC % and Budget Spent % are read, never recomputed (except merged companion rows);
the loader hard-fails if they're missing. The companion file is merged by project ID without the
`-Int` suffix at 3.75 SAR/USD. Regression: `python3 tools/gen_budget.py --check`.

## Approvals — weekly

```
python3 build_dashboard_html.py "Timesheets <n>.csv" "Employees <n>.xlsx" "Staff expenses <n>.xlsx"   # any order
```

Keeps Draft/Submitted items more than 10 days past their GL date, as of the export's own timestamp.
Files are recognised by content, columns by header. `EXCLUDE_NAMES` / `EXCLUDE_EMPLOYEE_IDS` keep
executives off the page; unmapped location IDs show as raw IDs rather than disappearing.

## Early Warning System + Projects — monthly

```
python3 regenerate_ews_data.py "SR of Projects @ Risk - Automated <date>.xlsx"
python3 regenerate_project_tasks.py "<same>.xlsx" project_tasks.js
python3 tools/apply_ews_blocks.py work/ews site        # splices the data blocks into the pages
```

The source workbook is formula-driven in production and must be calculated (opened and saved) before
the scripts read it. Tiers: Critical > 25% over, Warning 10–25%, Watch 5–10% (cost or hours).

## PM Hours & Expenses — monthly

```
python3 regenerate_pm_hours_data.py "CM PM Report.xlsx"      # --deploy for a capped, smaller file
```

Reads the Timesheet Data and Expense Data sheets. The export covers one region's PMs ("CM" =
Central & Mountain, 11 PMOs, about half the company's project hours; see `CM_OFFICES` in
`tools/gen_pm_hours.py`), which keeps `pm_hours_data.js` near 2 MB so the page opens quickly. To move to a new period set `FPA_AS_OF` (or
`CURRENT_MONTH`) and mark the month closed with `PARTIAL_LAST_MONTH = False`.

## Proposal Command Center — monthly

```
python3 regenerate_proposal_data.py "Proposal Hours Summary_<month>.xlsx"
```

Reads the office, segmented, $ amount, record-count and individual summaries; month columns are
detected from the header.

## Revenue Forecast — monthly

```
python3 regenerate_revenue_forecast_data.py "Revenue Forecast.xlsx" revenue_forecast_data.js [entries.csv]
```

The optional CSV seeds starting forecast values (same format as the page's Export); viewers' own edits
in the browser take precedence. Forecast months roll forward each quarter.

## Location Scorecard — monthly

`python3 tools/gen_location_drivers.py` writes `drivers_data.js` and the page's snapshot block; the
page also reads `util_data.js` from the hub pipeline.
