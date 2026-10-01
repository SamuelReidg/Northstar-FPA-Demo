"""
Regenerate pm_hours_data.js from the source PM Hours & Expenses report.

Usage:
    python regenerate_pm_hours_data.py [<source.xlsx>] [--deploy]

If no path is provided, defaults to looking for the most recent
"*CM PM Report*.xlsx" in the same folder.

Modes:
    (default)   Uncapped output (~2.5 MB). Every project/employee/category
                ships. Best when running locally with Python.
    --deploy    Generous caps applied (~600-800 KB). Used when Claude
                writes the file directly to the shared drive in one shot.
                Covers every current PM with significant headroom.

This script writes pm_hours_data.js next to itself so the
pm-hours-and-expenses.html dashboard refreshes automatically.

Required packages: openpyxl
    pip install openpyxl
"""

from __future__ import annotations
import json
import os
import sys
import calendar
import re
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("ERROR: openpyxl is required. Run:  pip install openpyxl")

# ----------------------------------------------------------------------
# Config — update if Excel column positions change
# ----------------------------------------------------------------------
CURRENT_YEAR  = 2026
CURRENT_MONTH = 9    # September (the 9.29.2026 package export runs through 09/26)

# Set True while the newest month is still in progress: its label gets a "*" so a
# part-month does not read as a collapse in hours. Flip to False once it closes.
PARTIAL_LAST_MONTH = True

# Reproducible builds: FPA_AS_OF=YYYY-MM-DD overrides the period above (year/month
# from the date; the month is partial unless the date is its last day). Unset =
# the hand-maintained constants above, exactly as before.
_AS_OF_ENV = os.environ.get("FPA_AS_OF", "").strip()
if _AS_OF_ENV:
    _as_of = datetime.strptime(_AS_OF_ENV, "%Y-%m-%d")
    CURRENT_YEAR, CURRENT_MONTH = _as_of.year, _as_of.month
    PARTIAL_LAST_MONTH = _as_of.day < calendar.monthrange(_as_of.year, _as_of.month)[1]

# MONTHS and every monthly series are DERIVED from CURRENT_MONTH -- do not hand-edit.
# (Before 2026-08-18 the sparkline was hardcoded to range(1, 6) = Jan..May and
# silently disagreed with the labels whenever CURRENT_MONTH moved.)
_MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTHS = _MONTH_ABBR[:CURRENT_MONTH]
if PARTIAL_LAST_MONTH and MONTHS:
    MONTHS = MONTHS[:-1] + [MONTHS[-1] + "*"]

# KPI strip captions -- derived, never hand-edited (they read "May" until 2026-08-18).
_LAST_DAY = calendar.monthrange(CURRENT_YEAR, CURRENT_MONTH)[1]
_CUR_ABBR = _MONTH_ABBR[CURRENT_MONTH - 1]
_PERIOD_THIS = (f"This Month ({_CUR_ABBR} 1 - {_LAST_DAY}, partial)"
                if PARTIAL_LAST_MONTH else
                f"This Month ({_CUR_ABBR} 1 - {_LAST_DAY})")
_PERIOD_YTD = f"Year to Date (Jan - {_CUR_ABBR})"

# Timesheet column indices (0-based)
TS_ENTRY_DATE  = 3   # D
TS_PROJECT_ID  = 4   # E
TS_BILLABLE    = 5   # F  (Utilized hours)
TS_HOURS       = 7   # H  (Total Hours)
TS_TASK_NAME   = 10  # K
TS_CONTACT     = 1   # B  (employee, "Last, First")
# Older file also had:
TS_PM_NAME     = 21  # V  (Project Manager Name — only in older file)
TS_PROJ_NAME   = 22  # W  (Project name — only in older file)

# Expenses column indices (0-based)
EX_TX_DATE     = 0   # A
EX_GL_DATE     = 1   # B  (we filter on this)
EX_EMPLOYEE    = 5   # F  (Employee Name)
EX_AMOUNT      = 6   # G
EX_CATEGORY    = 11  # L  (GL Account Name)
EX_PM_NAME     = 17  # R
EX_PROJECT_ID  = 18  # S
EX_PROJ_NAME   = 19  # T

# ----------------------------------------------------------------------
# Caps
# ----------------------------------------------------------------------
# Default mode = uncapped: every project / employee / category combination
# ships, so smaller PMs and growing portfolios stay fully represented.
# Local file becomes ~2.5 MB at current data volume.
#
# Deploy mode (--deploy flag) = generous caps tuned so the file fits in a
# single Claude tool call (~600-800 KB) while still covering every current
# PM with at least 6x headroom. Use this when Claude is writing the file
# directly to the shared drive. Bump caps if a PM ever outgrows them.

DEPLOY_MODE = "--deploy" in sys.argv

if DEPLOY_MODE:
    # Generous caps — tuned to fit one-shot deploy. Size them from the
    # largest PM's hours-PE / expense-PE row counts and the biggest data
    # array (usually hours_by_project_employee_ytd), leaving headroom for
    # portfolio growth.
    TOP_HOURS_BY_PROJ   = 300
    TOP_BILL_BY_PROJ    = 300
    TOP_PROJ_CATEGORY   = 500
    TOP_EMP_FLAT        = 800
    TOP_PE_FLAT         = 1500
    TOP_PER_PHASE       = None    # phase entries are tiny (~10)
    TOP_PER_CATEGORY    = None    # category entries are tiny (~20)
    TOP_MONTHLY_PIDS    = 1500
else:
    # Uncapped — set any of these to an int to re-enable that cap.
    TOP_HOURS_BY_PROJ   = None
    TOP_BILL_BY_PROJ    = None
    TOP_PROJ_CATEGORY   = None
    TOP_EMP_FLAT        = None
    TOP_PE_FLAT         = None
    TOP_PER_PHASE       = None
    TOP_PER_CATEGORY    = None
    TOP_MONTHLY_PIDS    = None

def _cap(lst, n):
    """Apply a cap if n is an int, else return list unchanged."""
    return lst if n is None else lst[:n]


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
EXCEL_EPOCH = datetime(1899, 12, 30)

def to_date(v):
    if v is None: return None
    if isinstance(v, datetime): return v
    if isinstance(v, (int, float)):
        try: return EXCEL_EPOCH + timedelta(days=float(v))
        except Exception: return None
    if isinstance(v, str):
        try: return datetime.fromisoformat(v)
        except Exception: return None
    return None

def f(v):
    try: return float(v) if v is not None else 0.0
    except Exception: return 0.0

def s(v):
    return str(v).strip() if v is not None else ""

def phase_of(task: str) -> str:
    t = task.lower()
    if not t: return "Unspecified"
    if "ovhd" in t or "overhead" in t:
        if "proposal" in t: return "Proposal Writing"
        if "training" in t: return "Training"
        if "vacation" in t or "pto" in t: return "PTO / Vacation"
        if "meeting" in t: return "Meetings"
        if "business development" in t or "bd" in t: return "Business Dev."
        if "administrative" in t or "admin" in t: return "Administrative"
        if "non-posting" in t: return "Non-Posting"
        return "Overhead Other"
    if ("fieldwork" in t or t.startswith("field ")
        or "field investigation" in t or "field inventory" in t
        or "maritime field" in t):
        return "Fieldwork"
    if "survey" in t:        return "Survey"
    if "lab " in t or "laboratory" in t: return "Lab Analysis"
    if "report" in t:        return "Reporting"
    if "monitoring" in t:    return "Monitoring"
    if "literature" in t or "research" in t or "background" in t:
        return "Research"
    if "data entry" in t or "gis" in t or "mapping" in t:
        return "GIS / Data"
    if "travel" in t:   return "Travel"
    if "curation" in t: return "Curation"
    return "Other"

def category_of(name: str) -> str:
    cl = name.lower()
    if not cl: return "Uncategorized"
    if "per diem" in cl: return "Per Diem"
    if "hotel" in cl or "lodging" in cl: return "Lodging"
    if "mileage" in cl: return "Mileage"
    if "travel" in cl: return "Travel"
    if "meal" in cl: return "Meals"
    if "materials" in cl or "supplies" in cl: return "Materials & Supplies"
    if "permit" in cl or "registration" in cl: return "Permits & Reg."
    if "consulting" in cl: return "Consulting"
    if "postage" in cl or "delivery" in cl: return "Postage & Delivery"
    if "software" in cl: return "Software"
    if "telephone" in cl or "internet" in cl: return "Telecom"
    if "rent" in cl: return "Rent"
    if "marketing" in cl: return "Marketing"
    if "dues" in cl or "subscription" in cl: return "Dues & Subs."
    if "conference" in cl or "event" in cl: return "Conferences"
    return "Other"


# ----------------------------------------------------------------------
# Excel extraction
# ----------------------------------------------------------------------
def read_xlsx(path: Path):
    wb = load_workbook(path, data_only=True, read_only=True)
    sheets = wb.sheetnames

    if "Timesheet Data" not in sheets:
        sys.exit(f"ERROR: '{path.name}' has no 'Timesheet Data' sheet")
    EXP_ALIASES = ("Expenses Data", "Expense Data", "Expenses", "Expense")
    exp_tab = next((n for n in EXP_ALIASES if n in sheets), None)
    if exp_tab is None:
        sys.exit(f"ERROR: '{path.name}' has no expenses sheet "
                 f"(tried {EXP_ALIASES}); tabs present: {sheets}")

    ts_rows = []
    for i, row in enumerate(wb["Timesheet Data"].iter_rows(values_only=True)):
        if i == 0: continue
        d = to_date(row[TS_ENTRY_DATE])
        if not d: continue
        ts_rows.append({
            "y":     d.year,
            "m":     d.month,
            "dd":    d.strftime("%m%d"),    # MMDD — compact date for transaction drill
            "pid":   s(row[TS_PROJECT_ID]),
            "util":  f(row[TS_BILLABLE]),
            "hrs":   f(row[TS_HOURS]),
            "task":  s(row[TS_TASK_NAME]),
            "emp":   s(row[TS_CONTACT]),
            "pm":    s(row[TS_PM_NAME]) if len(row) > TS_PM_NAME   else "",
            "pname": s(row[TS_PROJ_NAME]) if len(row) > TS_PROJ_NAME else "",
        })

    ex_rows = []
    for i, row in enumerate(wb[exp_tab].iter_rows(values_only=True)):
        if i == 0: continue
        gld = to_date(row[EX_GL_DATE])
        if not gld: continue
        ex_rows.append({
            "y":     gld.year,
            "m":     gld.month,
            "dd":    gld.strftime("%m%d"),
            "amt":   f(row[EX_AMOUNT]),
            "emp":   s(row[EX_EMPLOYEE]),
            "cat":   s(row[EX_CATEGORY]),
            "pm":    s(row[EX_PM_NAME]),
            "pid":   s(row[EX_PROJECT_ID]),
            "pname": s(row[EX_PROJ_NAME]),
        })

    return ts_rows, ex_rows


# ----------------------------------------------------------------------
# Build all aggregates
# ----------------------------------------------------------------------
def build_data(ts_rows, ex_rows):
    # Resolve pid -> (pm, name) using whatever has it
    pid_pm   = {}
    pid_name = {}
    for r in ts_rows + ex_rows:
        pid = r.get("pid", "")
        if not pid: continue
        if r.get("pm") and pid not in pid_pm: pid_pm[pid] = r["pm"]
        if r.get("pname") and pid not in pid_name: pid_name[pid] = r["pname"]

    def proj_name(pid: str) -> str:
        return (pid_name.get(pid) or pid)[:80]

    CURRENT = (CURRENT_YEAR, CURRENT_MONTH)

    # ---- KPIs ----
    pm_hours_m  = sum(r["hrs"]  for r in ts_rows if (r["y"], r["m"]) == CURRENT)
    pm_hours_y  = sum(r["hrs"]  for r in ts_rows if r["y"] == CURRENT_YEAR)
    bill_m      = sum(r["util"] for r in ts_rows if (r["y"], r["m"]) == CURRENT)
    bill_y      = sum(r["util"] for r in ts_rows if r["y"] == CURRENT_YEAR)
    exp_m       = sum(r["amt"]  for r in ex_rows if (r["y"], r["m"]) == CURRENT)
    exp_y       = sum(r["amt"]  for r in ex_rows if r["y"] == CURRENT_YEAR)

    # ---- Sparkline (org-wide monthly Jan..CURRENT_MONTH, billable + expense) ----
    _MRANGE = range(1, CURRENT_MONTH + 1)
    bill_series = [round(sum(r["util"] for r in ts_rows
                              if r["y"] == CURRENT_YEAR and r["m"] == m))
                    for m in _MRANGE]
    exp_series  = [round(sum(r["amt"] for r in ex_rows
                              if r["y"] == CURRENT_YEAR and r["m"] == m))
                    for m in _MRANGE]

    # ---- Phases (this month + YTD) ----
    def build_phases(filter_fn):
        totals = defaultdict(float)
        per_proj = defaultdict(lambda: defaultdict(float))
        for r in ts_rows:
            if not filter_fn(r): continue
            ph = phase_of(r["task"])
            totals[ph] += r["hrs"]
            if r["pid"]: per_proj[ph][r["pid"]] += r["hrs"]
        out = []
        for ph, h in sorted(totals.items(), key=lambda x: -x[1]):
            if h <= 0: continue
            projs = _cap(sorted(per_proj[ph].items(), key=lambda x: -x[1]), TOP_PER_PHASE)
            out.append({
                "phase":   ph,
                "hours":   round(h, 1),
                "projects":[{"name": proj_name(pid)[:60],
                              "hours": round(hv, 1),
                              "pm":    pid_pm.get(pid, ""),
                              "pid":   pid}
                            for pid, hv in projs]
            })
        return out

    phases_this_month = build_phases(lambda r: (r["y"], r["m"]) == CURRENT)
    phases_ytd        = build_phases(lambda r: r["y"] == CURRENT_YEAR)

    # ---- Categories (this month + YTD) ----
    def build_categories(filter_fn):
        totals = defaultdict(float)
        per_proj = defaultdict(lambda: defaultdict(float))
        for r in ex_rows:
            if not filter_fn(r): continue
            c = category_of(r["cat"])
            totals[c] += r["amt"]
            if r["pid"]: per_proj[c][r["pid"]] += r["amt"]
        out = []
        for c, a in sorted(totals.items(), key=lambda x: -x[1]):
            if a <= 0: continue
            projs = _cap(sorted(per_proj[c].items(), key=lambda x: -x[1]), TOP_PER_CATEGORY)
            out.append({
                "category": c,
                "amount":   round(a, 2),
                "projects":[{"name": proj_name(pid)[:60],
                              "amount": round(av, 2),
                              "pm":     pid_pm.get(pid, ""),
                              "pid":    pid}
                            for pid, av in projs]
            })
        return out

    categories_this_month = build_categories(lambda r: (r["y"], r["m"]) == CURRENT)
    categories_ytd        = build_categories(lambda r: r["y"] == CURRENT_YEAR)

    # ---- Top hours by project (current month + per-month series) ----
    hours_by_pid_m  = defaultdict(float)
    hours_by_pid_y  = defaultdict(float)
    bill_by_pid_y   = defaultdict(float)
    hours_by_pid_per_month = defaultdict(lambda: [0.0]*CURRENT_MONTH)
    bill_by_pid_per_month  = defaultdict(lambda: [0.0]*CURRENT_MONTH)
    for r in ts_rows:
        if r["y"] != CURRENT_YEAR or not r["pid"]: continue
        hours_by_pid_y[r["pid"]] += r["hrs"]
        bill_by_pid_y[r["pid"]]  += r["util"]
        if r["m"] == CURRENT_MONTH:
            hours_by_pid_m[r["pid"]] += r["hrs"]
        if 1 <= r["m"] <= CURRENT_MONTH:
            hours_by_pid_per_month[r["pid"]][r["m"]-1] += r["hrs"]
            bill_by_pid_per_month [r["pid"]][r["m"]-1] += r["util"]

    top_hours = _cap([{"name": proj_name(pid), "hours": round(h, 1),
                        "pm": pid_pm.get(pid, ""), "pid": pid}
                       for pid, h in sorted(hours_by_pid_m.items(), key=lambda x:-x[1])
                       if h > 0], TOP_HOURS_BY_PROJ)

    top_bill = _cap([{"name": proj_name(pid), "hours": round(h, 1),
                       "pm": pid_pm.get(pid, ""), "pid": pid}
                      for pid, h in sorted(bill_by_pid_y.items(), key=lambda x:-x[1])
                      if h > 0], TOP_BILL_BY_PROJ)

    # ---- Top expenses by project (current month + YTD) ----
    exp_by_pid_m = defaultdict(float)
    exp_by_pid_y = defaultdict(float)
    proj_cat_y   = defaultdict(lambda: defaultdict(float))
    proj_cat_m   = defaultdict(lambda: defaultdict(float))
    for r in ex_rows:
        if r["y"] != CURRENT_YEAR: continue
        pid = r["pid"] or "(no-pid)"
        exp_by_pid_y[pid] += r["amt"]
        cat = category_of(r["cat"])
        proj_cat_y[pid][cat] += r["amt"]
        if r["m"] == CURRENT_MONTH:
            exp_by_pid_m[pid] += r["amt"]
            proj_cat_m[pid][cat] += r["amt"]

    top_exp_m = _cap([{"name": proj_name(pid), "amount": round(a, 2),
                        "pm": pid_pm.get(pid, ""), "pid": pid}
                       for pid, a in sorted(exp_by_pid_m.items(), key=lambda x:-x[1])
                       if a > 0], TOP_HOURS_BY_PROJ)
    top_exp_y = _cap([{"name": proj_name(pid), "amount": round(a, 2),
                        "pm": pid_pm.get(pid, ""), "pid": pid}
                       for pid, a in sorted(exp_by_pid_y.items(), key=lambda x:-x[1])
                       if a > 0], TOP_HOURS_BY_PROJ)

    def proj_category_breakdown(per_pid_per_cat, top_pids):
        out = []
        for pid in top_pids:
            cats = sorted(per_pid_per_cat[pid].items(), key=lambda x:-x[1])  # no cap on categories per project
            out.append({
                "pid":     pid,
                "name":    proj_name(pid)[:70],
                "pm":      pid_pm.get(pid, ""),
                "amount":  round(sum(per_pid_per_cat[pid].values()), 2),
                "categories":[{"category": c, "amount": round(a, 2)}
                               for c, a in cats if a != 0]
            })
        return out

    top_pids_y = _cap([p for p, _ in
                       sorted(exp_by_pid_y.items(), key=lambda x:-x[1])
                       if _ > 0], TOP_PROJ_CATEGORY)
    top_pids_m = _cap([p for p, _ in
                       sorted(exp_by_pid_m.items(), key=lambda x:-x[1])
                       if _ > 0], TOP_PROJ_CATEGORY)
    exp_proj_cat_y = proj_category_breakdown(proj_cat_y, top_pids_y)
    exp_proj_cat_m = proj_category_breakdown(proj_cat_m, top_pids_m)

    # ---- Hours by project by employee (flat) — current month + YTD ----
    pe_hours_m = defaultdict(float)
    pe_hours_y = defaultdict(float)
    for r in ts_rows:
        if r["y"] != CURRENT_YEAR or not r["pid"] or not r["emp"]: continue
        pe_hours_y[(r["pid"], r["emp"])] += r["hrs"]
        if r["m"] == CURRENT_MONTH:
            pe_hours_m[(r["pid"], r["emp"])] += r["hrs"]
    def pe_list(d):
        return _cap([{"name": proj_name(pid), "pm": pid_pm.get(pid, ""), "pid": pid,
                       "employee": emp, "hours": round(h, 1)}
                      for (pid, emp), h in sorted(d.items(), key=lambda x:-x[1])
                      if h > 0], TOP_PE_FLAT)
    hours_pe_m = pe_list(pe_hours_m)
    hours_pe_y = pe_list(pe_hours_y)

    # ---- Expenses by project by employee (flat) ----
    pe_exp_m = defaultdict(float)
    pe_exp_y = defaultdict(float)
    for r in ex_rows:
        if r["y"] != CURRENT_YEAR or not r["emp"]: continue
        pid = r["pid"] or "(no-pid)"
        pe_exp_y[(pid, r["emp"])] += r["amt"]
        if r["m"] == CURRENT_MONTH:
            pe_exp_m[(pid, r["emp"])] += r["amt"]
    def pe_amount_list(d):
        return _cap([{"name": proj_name(pid), "pm": pid_pm.get(pid, ""), "pid": pid,
                       "employee": emp, "amount": round(a, 2)}
                      for (pid, emp), a in sorted(d.items(), key=lambda x:-x[1])
                      if a > 0], TOP_PE_FLAT)
    exp_pe_m = pe_amount_list(pe_exp_m)
    exp_pe_y = pe_amount_list(pe_exp_y)

    # ---- Expenses by Employee (totals only) ----
    emp_m = defaultdict(float)
    emp_y = defaultdict(float)
    emp_projects = defaultdict(set)
    for r in ex_rows:
        if r["y"] != CURRENT_YEAR or not r["emp"]: continue
        emp_y[r["emp"]] += r["amt"]
        emp_projects[r["emp"]].add(r["pid"] or "(no-pid)")
        if r["m"] == CURRENT_MONTH:
            emp_m[r["emp"]] += r["amt"]
    exp_emp_m = _cap([{"name": e, "amount": round(a, 2)}
                       for e, a in sorted(emp_m.items(), key=lambda x:-x[1]) if a > 0], TOP_EMP_FLAT)
    exp_emp_y = _cap([{"name": e, "amount": round(a, 2), "projects": len(emp_projects[e])}
                       for e, a in sorted(emp_y.items(), key=lambda x:-x[1]) if a > 0], TOP_EMP_FLAT)

    # ---- Per-PM monthly totals: billable + expense + total hours ----
    # NM = months carried per PM. Was hardcoded to 5 (Jan..May) and did NOT move
    # with CURRENT_MONTH, so these series silently disagreed with the chart labels.
    NM = CURRENT_MONTH
    pm_bill = defaultdict(lambda: [0.0]*NM)
    pm_exp  = defaultdict(lambda: [0.0]*NM)
    pm_hrs  = defaultdict(lambda: [0.0]*NM)
    for r in ts_rows:
        if r["y"] != CURRENT_YEAR or not (1 <= r["m"] <= NM): continue
        pm = pid_pm.get(r["pid"], "") or r.get("pm") or ""
        if pm:
            pm_bill[pm][r["m"]-1] += r["util"]
            pm_hrs[pm][r["m"]-1]  += r["hrs"]
    for r in ex_rows:
        if r["y"] != CURRENT_YEAR or not (1 <= r["m"] <= NM): continue
        pm = pid_pm.get(r["pid"], "") or r.get("pm") or ""
        if pm: pm_exp[pm][r["m"]-1] += r["amt"]

    def monthly_obj(b, e, h=None):
        entry = {}
        if any(b): entry["b"] = [round(v, 1) for v in b]
        if any(e): entry["e"] = [round(v, 2) for v in e]
        if h is not None and any(h): entry["h"] = [round(v, 1) for v in h]
        return entry

    monthly_by_pm    = {pm: monthly_obj(pm_bill.get(pm,[0]*NM), pm_exp.get(pm,[0]*NM))
                       for pm in (set(pm_bill) | set(pm_exp))}
    monthly_by_pm    = {k: v for k, v in monthly_by_pm.items() if v}
    monthly_h_by_pm  = {pm: {"h":[round(v,1) for v in pm_hrs[pm]]}
                       for pm in pm_hrs if any(pm_hrs[pm])}

    # ---- Per-PID monthly (top TOP_MONTHLY_PIDS by combined activity) ----
    pid_activity = {pid: sum(bill_by_pid_per_month.get(pid,[0]*CURRENT_MONTH)) +
                          sum(hours_by_pid_per_month.get(pid,[0]*CURRENT_MONTH))
                    for pid in set(bill_by_pid_per_month) | set(hours_by_pid_per_month)}
    top_pids_for_monthly = _cap(sorted(pid_activity, key=lambda x:-pid_activity[x]), TOP_MONTHLY_PIDS)

    pid_exp_per_month = defaultdict(lambda: [0.0]*CURRENT_MONTH)
    for r in ex_rows:
        if r["y"] != CURRENT_YEAR or not (1 <= r["m"] <= CURRENT_MONTH): continue
        pid = r["pid"] or "(no-pid)"
        pid_exp_per_month[pid][r["m"]-1] += r["amt"]

    monthly_by_pid = {}
    monthly_h_by_pid = {}
    for pid in top_pids_for_monthly:
        b = bill_by_pid_per_month.get(pid, [0]*CURRENT_MONTH)
        e = pid_exp_per_month.get(pid, [0]*CURRENT_MONTH)
        h = hours_by_pid_per_month.get(pid, [0]*CURRENT_MONTH)
        entry = monthly_obj(b, e)
        if entry: monthly_by_pid[pid] = entry
        if any(h): monthly_h_by_pid[pid] = {"h": [round(v, 1) for v in h]}

    # ---- Filter lookups ----
    pms_set  = {pm for pm in pid_pm.values() if pm}
    pms_set |= {r["pm"] for r in ts_rows + ex_rows if r["pm"]}
    pids_set = {r["pid"] for r in ts_rows + ex_rows if r["pid"]}

    # ---- Transactions (for drill-down) -------------------------------------
    # Compact format: lookup tables + integer-indexed rows. Keeps ~100K
    # timesheet rows + ~10K expense rows down to ~3 MB instead of ~12 MB.
    # Filter to current year only — older data isn't displayed anyway.
    emp_set, task_set, cat_set, pid_set_tx = set(), set(), set(), set()
    for r in ts_rows:
        if r["y"] != CURRENT_YEAR or not r["pid"]: continue
        if r["emp"]:  emp_set.add(r["emp"])
        if r["task"]: task_set.add(r["task"])
        pid_set_tx.add(r["pid"])
    for r in ex_rows:
        if r["y"] != CURRENT_YEAR or not r["pid"]: continue
        if r["emp"]: emp_set.add(r["emp"])
        if r["cat"]: cat_set.add(r["cat"])
        pid_set_tx.add(r["pid"])

    emp_list  = sorted(emp_set)
    task_list = sorted(task_set)
    cat_list  = sorted(cat_set)
    tx_pid_list = sorted(pid_set_tx)
    emp_to_i  = {e:i for i,e in enumerate(emp_list)}
    task_to_i = {t:i for i,t in enumerate(task_list)}
    cat_to_i  = {c:i for i,c in enumerate(cat_list)}
    pid_to_i  = {p:i for i,p in enumerate(tx_pid_list)}

    # Compact timesheet rows: [pid_i, emp_i, task_i, "MMDD", hrs, util]
    # Compact expense rows:   [pid_i, emp_i, cat_i, "MMDD", amt]
    ts_compact = []
    for r in ts_rows:
        if r["y"] != CURRENT_YEAR or not r["pid"]: continue
        ts_compact.append([
            pid_to_i[r["pid"]],
            emp_to_i.get(r["emp"], -1),
            task_to_i.get(r["task"], -1),
            r["dd"],
            round(r["hrs"], 2),
            round(r["util"], 2),
        ])
    ex_compact = []
    for r in ex_rows:
        if r["y"] != CURRENT_YEAR or not r["pid"]: continue
        ex_compact.append([
            pid_to_i[r["pid"]],
            emp_to_i.get(r["emp"], -1),
            cat_to_i.get(r["cat"], -1),
            r["dd"],
            round(r["amt"], 2),
        ])

    transactions = {
        "emp_list":  emp_list,
        "task_list": task_list,
        "cat_list":  cat_list,
        "pid_list":  tx_pid_list,
        "ts":        ts_compact,
        "ex":        ex_compact,
    }

    # ---- Assemble final dictionary ----
    return {
        "period_label_this": _PERIOD_THIS,
        "period_label_ytd":  _PERIOD_YTD,
        "kpi": {
            "pm_hours_this_month":           round(pm_hours_m, 1),
            "billable_hours_ytd":            round(bill_y, 1),
            "field_expenses_ytd":            round(exp_y, 2),
            "total_field_expenses_this_month": round(exp_m, 2),
        },
        "kpi_extra": {
            "pm_hours_ytd":                  round(pm_hours_y, 1),
            "billable_hours_this_month":     round(bill_m, 1),
            "field_expenses_this_month":     round(exp_m, 2),
            "field_expenses_ytd":            round(exp_y, 2),
        },
        "sparkline": {
            "labels":         MONTHS,
            "billable_hours": bill_series,
            "field_expenses": exp_series,
        },
        "phases_this_month":          phases_this_month,
        "phases_ytd":                 phases_ytd,
        "categories_ytd":             categories_ytd,
        "categories_this_month":      categories_this_month,
        "top_hours_by_project_this_month": top_hours,
        "top_billable_by_project_ytd":     top_bill,
        "top_expenses_by_project_this_month": top_exp_m,
        "largest_expenses_ytd_by_project":    top_exp_y,
        "expenses_by_project_category_ytd":        exp_proj_cat_y,
        "expenses_by_project_category_this_month": exp_proj_cat_m,
        "hours_by_project_employee_this_month":    hours_pe_m,
        "hours_by_project_employee_ytd":           hours_pe_y,
        "expenses_by_project_employee_ytd":        exp_pe_y,
        "expenses_by_project_employee_this_month": exp_pe_m,
        "expenses_by_employee_this_month":         exp_emp_m,
        "expenses_by_employee_ytd":                exp_emp_y,
        "monthly_by_pm":   monthly_by_pm,
        "monthly_h_by_pm": monthly_h_by_pm,
        "monthly_by_pid":  monthly_by_pid,
        "monthly_h_by_pid":monthly_h_by_pid,
        "filters": {
            "pms":  sorted(pms_set),
            "pids": sorted(pids_set),
        },
        "transactions": transactions,
    }


# ----------------------------------------------------------------------
# Write JS file
# ----------------------------------------------------------------------
def write_js(data: dict, out_path: Path):
    """Write data as window.DASH_DATA = {...}; with one item per line for
    array-of-object keys, so the file is moderately diffable."""
    BULK_ARRAY_KEYS = {
        "phases_this_month", "phases_ytd",
        "categories_ytd", "categories_this_month",
        "top_hours_by_project_this_month", "top_billable_by_project_ytd",
        "top_expenses_by_project_this_month", "largest_expenses_ytd_by_project",
        "expenses_by_project_category_ytd", "expenses_by_project_category_this_month",
        "hours_by_project_employee_this_month", "hours_by_project_employee_ytd",
        "expenses_by_project_employee_ytd",    "expenses_by_project_employee_this_month",
        "expenses_by_employee_this_month",     "expenses_by_employee_ytd",
    }
    lines = ["window.DASH_DATA = {"]
    items = list(data.items())
    for i, (k, v) in enumerate(items):
        last = i == len(items) - 1
        sep  = "" if last else ","
        if k in BULK_ARRAY_KEYS and isinstance(v, list):
            lines.append(f'"{k}":[')
            for j, item in enumerate(v):
                c = "" if j == len(v) - 1 else ","
                lines.append(json.dumps(item, separators=(",", ":"), ensure_ascii=False) + c)
            lines.append("]" + sep)
        elif k == "filters" and isinstance(v, dict):
            lines.append(f'"{k}":{json.dumps(v, separators=(",", ":"), ensure_ascii=False)}' + sep)
        else:
            lines.append(f'"{k}":' + json.dumps(v, separators=(",", ":"), ensure_ascii=False) + sep)
    lines.append("};")

    mode_label = "DEPLOY (generous caps, ~600-800 KB)" if DEPLOY_MODE else "LOCAL (uncapped, ~2.5 MB)"
    header = (
        "/* Auto-generated by regenerate_pm_hours_data.py\n"
        f"   Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
        f"   Mode: {mode_label}\n"
        "   Do not edit by hand — rerun the script when the source xlsx changes.\n"
        "*/\n"
    )
    out_path.write_text(header + "\n".join(lines) + "\n", encoding="utf-8")


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------
def find_source(folder: Path) -> Path:
    candidates = sorted(folder.glob("*CM PM Report*.xlsx"),
                        key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        candidates = sorted(folder.glob("*PM Hours*Expenses*.xlsx"),
                            key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        sys.exit("ERROR: no CM PM Report .xlsx found in folder; pass the path explicitly.")
    return candidates[0]


def main():
    here = Path(__file__).parent
    # Strip flags from positional args
    positional = [a for a in sys.argv[1:] if not a.startswith("--")]
    if positional:
        src = Path(positional[0])
    else:
        src = find_source(here)
    if not src.exists():
        sys.exit(f"ERROR: source file not found: {src}")

    print(f"Mode: {'DEPLOY (generous caps)' if DEPLOY_MODE else 'LOCAL (uncapped)'}")
    print(f"Reading: {src.name}")
    ts_rows, ex_rows = read_xlsx(src)
    print(f"  Timesheet rows: {len(ts_rows):,}")
    print(f"  Expense rows:   {len(ex_rows):,}")

    print("Computing aggregates...")
    data = build_data(ts_rows, ex_rows)

    out = here / "pm_hours_data.js"
    write_js(data, out)
    print(f"Wrote: {out} ({out.stat().st_size:,} bytes)")

    print("\nKPIs:")
    print(f"  PM Hours This Month:     {data['kpi']['pm_hours_this_month']:>12,.1f}")
    print(f"  PM Hours YTD:            {data['kpi_extra']['pm_hours_ytd']:>12,.1f}")
    print(f"  Billable YTD:            {data['kpi']['billable_hours_ytd']:>12,.1f}")
    print(f"  Billable This Month:     {data['kpi_extra']['billable_hours_this_month']:>12,.1f}")
    print(f"  Field Expenses YTD:      ${data['kpi']['field_expenses_ytd']:>12,.2f}")
    print(f"  Field Exp This Month:    ${data['kpi']['total_field_expenses_this_month']:>12,.2f}")
    print(f"\nFilters:")
    print(f"  PMs:  {len(data['filters']['pms'])}")
    print(f"  PIDs: {len(data['filters']['pids'])}")
    print("\nDone. Refresh the dashboard to see the updated numbers.")


if __name__ == "__main__":
    main()
