# The synthetic data model

`tools/demo_world.py` builds a single fictional company, **Northstar Heritage** (cultural- and
environmental-resource consulting), from `SEED = 20260926`. All source exports are written from this
one world by `tools/gen_*.py`, so every dashboard reports the same people, projects and dollars.

No figure is drawn independently; each is either an input assumption or a consequence of others.

## Entities

| Entity | Count (≈) | Key attributes |
|---|---|---|
| Offices / PMOs | 26 | regional, service-line, acquired ("Meridian"), international, admin; FT/PT headcount; target utilization |
| Employees | 730 (348 active) | role (field / senior / lead / admin), FT/PT, FLSA, target, supervisor, hire/term dates, cost and bill rates |
| Clients | 160 | generic fictional organisations |
| Projects | 1,200 | office, PM, client, sector, billing type, contract, target margin, original/revised budget, tasks, dates, status |
| Timesheet entries | 81,000 | daily rows by employee × project/overhead task × time type |
| Opportunities | 2,100 | office, owner, stage, amount, created/close dates, proposal hours |
| Expense reports / lines | 3,000 / 7,500 | per-diem, lodging, mileage… charged to worked projects; Draft→Submitted→Approved→Paid |

Names come from a fixed fictional pool (`tools/name_pool.py`); places, clients and project titles
are assembled from generic word lists.

## Relationships

**Utilization.** Weekly utilization for a person = their target − a small realism gap + a
persistent person effect + a persistent office effect with an AR(1) drift + a field-season curve
(regional offices peak June–July, dip in winter and around holidays) + demand (offices with thin
backlog run lower) + noise.

**Time off & calendar.** Holidays follow the calendar; PTO probability roughly doubles in July/August
and late December; sick time rises in winter; training is front-loaded in Q1. Part-time field staff
work mostly in the field season.

**Growth.** About 9% of today's field staff joined during the year, weighted to the spring ramp, so
headcount climbs through FY2026.

**Hours → projects.** Utilized hours go to the office's live projects in proportion to remaining
labor budget and only inside each project's date window. A slice is logged as *Overhead Non-Posting*,
hourly field staff pick up overtime in busy weeks, and a few negative transfer entries exist (as in
real timesheets).

**Project economics.**

- *Original budget* = contract × (1 − target margin); margins centre around 31% for T&M and 27% for
  fixed fee, varying by sector.
- *Change orders* (more likely on long jobs) raise the revised budget and contract together.
- *Cost to date* = prior-year labor + FY labor (hours × each person's burdened cost rate, 1.5× for
  OT) + expenses (a per-project share).
- *Overruns* come from a project burn tendency plus a persistent **PM effect** and occasional scope
  creep, so a minority of jobs and some PMs run over, which is what finance teams see.
- *Revenue*: fixed fee = % complete (cost / revised budget, capped at 100%) × contract; T&M = billed
  hours × bill rates, capped at contract.
- *Invoiced* lags revenue by recent work (30–60 days unbilled). *Backlog* = contract − revenue.

**Pipeline.** Win rates differ by office; proposal-writing hours from the timesheets are attached to
the opportunities open at the time, which is what the Proposal Command Center measures.

**Totals.** Every office, PMO, region and company figure is a sum of the rows beneath it; the
pipelines' reconciliation checks confirm it.

## Headline shape (as of 26 Sep 2026)

- Company billable utilization YTD ≈ 76% against a weighted plan of 80%; 74% after stripping non-posting hours.
- Active book ≈ $184M contract, ≈ $83M positive backlog.
- Median active project has used ~60% of budget; about a third are past 70%.
- Weighted overrun ≈ 1.8% on original budget; about 18% of projects are over their original budget.

## Rebuilding

`python3 tools/build_demo.py --rebuild-world` regenerates the world (cached in `tools/.cache/`), every
export and every page. Output is deterministic apart from "generated at" timestamps.
