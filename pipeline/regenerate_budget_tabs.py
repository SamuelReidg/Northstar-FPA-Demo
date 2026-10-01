#!/usr/bin/env python3
"""
regenerate_budget_tabs.py
=========================
Rebuilds the downstream "budget utilization" tabs from the weekly FP&A & QC
Project Report export. Filter logic reverse-engineered from and validated
row-for-row against the hand-built June workbooks (>70%, 90%&>$50k, 100% Rev
Rec FF all reproduce exactly), and re-validated against the archived June
workbook after the 08/18/2026 export-format change. The regression fixture is
demo_data/fixtures/budget_regression.xlsx (old flat layout, "As of June 16,
2026"); `python tools/gen_budget.py --check` asserts its numbers:
    rows 717 / orig_n 717 / >70% 228 / 90%&>$50k 76 / 100% Rev Rec FF 58,
    weighted overrun 1.3558% (original) / 0.9813% (revised ETC).

FORMAT-AGNOSTIC LOADER (rewritten 2026-08-18)
---------------------------------------------
Nothing about the layout is hardcoded any more. The loader finds the header
block by locating "Dimension name" in column B (that row = LBL), then reads:
    LBL-3 = group header row      LBL-2 = sub header row
    LBL   = column label row      LBL+1 = first data row
Every value column is located by its (group, sub) header pair, so inserted
columns cannot shift the read:
    Contract Value        / Inception To Date
    Revenue               / Inception To Date
    Invoiced Amount       / Inception To Date          (NOT the Year To Date twin)
    Total Revised Budget  / Orig Budget + Adjustments
    Total Budget|Original Budget / Original
    Total Budget Adjustment / ETC
    Total Costs           / Inception To Date          (NOT the YTD / Month Ending twins)
Recognized % and Budget Spent % are SYSTEM values that cannot be reproduced
from base columns (their rounding drives the 100%-Rev-Rec and >70% boundaries),
so they are read from whichever columns carry the sub-row labels "REV REC %"
and "Budget Spent %". These are matched on the SUB row, never on the label row
-- the 08/2026 export repeats the label "Summary" across Budget Spent %,
Gross profit and Margin %, so matching there would silently grab Margin %.
If either column is missing, or if more than 5% of rows carry a non-numeric
value in them, the loader HARD-FAILS rather than falling back to computed
ratios (that fallback silently shifts the 100% Rev Rec FF count).

ROW SELECTION
-------------
The 08/2026 export is the full project universe (every status) presented as an
indented hierarchy -- "All Projects" > "Active Projects" / "All Other Projects"
> project rows -- with subtotal rows. Only leaf rows are read; group and total
rows are skipped. Rows are then filtered to ACTIVE_STATUSES, and office
overhead pseudo-projects (any ID containing "OVHD") are dropped: they carry $0
contract and $0 budget against millions in cost, which distorts the revised-ETC
weighted overrun (demo 09/29 package: 1.39% -> 8.53% with overhead and admin
buckets left in).

ACTIVE_STATUSES holds the statuses the FP&A lead counts as active. The baseline (2026-08-18)
was "only statuses that were in scope before the 08/2026 vocabulary change", so
"Final Billing/Invoice", "On Hold", "Canceled" and the rest of that export's new
vocabulary are deliberately NOT included. One exception has been granted since:
"Transferred Entity" (new in the 09/2026 export: the international block moved to
the Gulf entity) IS counted as active per the FP&A lead, 2026-09-14 -- it is
treated like "In Progress", so it is NOT added to
FINAL_STATUSES.

Base fields (internal row order, 0-based):
  0 id  1 name  2 status  3 billing  4 location  5 PM
  6 ContractValue  7 Revenue  8 Invoiced  9 RevisedBudget  10 OriginalBudget
  11 BudgetAdjustment/ETC  12 TotalCosts
Computed: 13 RemainingBudget = 9-12   14 OrigBudgetRemaining = 10-12
  15 teamwork Status (unused)  16 Recognized %  17 Budget Spent %

Output tabs: All Active Projects | Projects >70% Exec View | Projects >70% |
  Projects 90% & > $50k | 100% Rev Rec FF | Weighted Overrun Calculation
(Weighted Overrun uses live SUMIF formulas against All Active Projects.)

Usage: python regenerate_budget_tabs.py "<export>.xlsx" [out.xlsx]
"""
import sys
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC_TAB = "All Active Projects"
FIRST_DATA_ROW = 12
LEAF_INDENT = 6
NONNUM_TOLERANCE = 0.05

# Drop rows with no commercial basis: no contract value, or no budget of any kind.
# This is the cleanup that used to be done by hand before pasting the tab. Validated
# 2026-08-18 against the manual workbook -- with it the engine reproduces the manual
# >70%, 90%&>$50k and 100% Rev Rec FF lists project-for-project, and the
# weighted-overrun totals to the cent. Without it, administrative cost buckets
# ("Sales and Marketing", "Company Meeting 2026", ...) and transferred projects whose
# contract sits in another entity's report carry cost against zero contract/budget and
# inflate revised-ETC weighted overrun (demo 09/29 package: 1.39% -> 3.55%).
# No effect on the archived June fixture (0 rows dropped) -- that tab was already cleaned.
REQUIRE_COMMERCIAL_BASIS = True

ACTIVE_STATUSES = {
    "In Progress",
    "Transferred Entity",                    # added 2026-09-14 per the FP&A lead: counts as active/in progress
    "Final Reporting/Review",
    "Final Reporting/Review/Lab/Curation",   # retired in the 08/2026 export; kept for older workbooks
    "Lab/Curation",
    "Cost Tracking on Closed Project",
}
FINAL_STATUSES = {
    "Final Reporting/Review", "Final Reporting/Review/Lab/Curation",
    "Lab/Curation", "Cost Tracking on Closed Project",
}
GROUP_ROW_PREFIXES = ("total ", "all projects", "active projects", "all other projects")

COLSPEC = {
    ("Contract Value", "Inception To Date"): "contract",
    ("Revenue", "Inception To Date"): "revenue",
    ("Invoiced Amount", "Inception To Date"): "invoiced",
    ("Total Revised Budget", "Orig Budget + Adjustments"): "revbud",
    ("Total Budget", "Original"): "budget",
    ("Original Budget", "Original"): "budget",
    ("Total Budget Adjustment", "ETC"): "etc",
    ("Total Costs", "Inception To Date"): "cost",
}
# "invoiced" is carried through for reference only and is absent from the
# pre-filtered weekly export (2026-08-18) -> optional, not required.
REQUIRED = ("contract", "revenue", "revbud", "budget", "etc", "cost")

FONT = "Arial"
TITLE_F  = Font(name=FONT, bold=True, size=11)
HDR_F    = Font(name=FONT, bold=True, size=10, color="FFFFFF")
SUB_F    = Font(name=FONT, italic=True, size=9, color="404040")
NOTE_F   = Font(name=FONT, size=9, color="404040")
BODY_F   = Font(name=FONT, size=10)
BOLD_F   = Font(name=FONT, bold=True, size=10)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
THIN     = Side(style="thin", color="D9D9D9")
BORDER   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY    = '$#,##0;($#,##0);"-"'
PCT      = '0.0%'
PCT3     = '0.000%'
CENTER   = Alignment(horizontal="center", vertical="center", wrap_text=True)


def num(x):
    return x if isinstance(x, (int, float)) else 0.0


def _pick_sheet(wb):
    if SRC_TAB in wb.sheetnames:
        return wb[SRC_TAB]
    for ws in wb.worksheets:
        for r in range(1, min(ws.max_row, 60) + 1):
            if str(ws.cell(r, 2).value).strip() == "Dimension name":
                return ws
    raise SystemExit('ERROR: no sheet with a "Dimension name" header found. '
                     'Tabs: %s' % wb.sheetnames)


def _labels_row(ws):
    for r in range(1, min(ws.max_row, 60) + 1):
        if str(ws.cell(r, 2).value).strip() == "Dimension name":
            return r
    raise SystemExit('ERROR: could not find the "Dimension name" header row on '
                     'sheet "%s".' % ws.title)


def _map_cols(ws, lbl):
    """Locate every value column by its (group, sub) header pair."""
    grp, sub = lbl - 3, lbl - 2
    if grp < 1:
        raise SystemExit("ERROR: header block starts above row 4; layout not recognised.")
    cols = {}
    for c in range(1, ws.max_column + 1):
        g = str(ws.cell(grp, c).value or "").strip()
        s = str(ws.cell(sub, c).value or "").strip()
        key = COLSPEC.get((g, s))
        if key and key not in cols:
            cols[key] = c
        if s == "REV REC %" and "recpct" not in cols:
            cols["recpct"] = c
        if s == "Budget Spent %" and "bspct" not in cols:
            cols["bspct"] = c
    missing = [k for k in REQUIRED if k not in cols]
    if missing:
        raise SystemExit(
            "ERROR: could not locate these columns by header text: %s\n"
            "       Found: %s\n"
            "       The export layout changed in a way the loader does not recognise."
            % (missing, {k: get_column_letter(v) for k, v in sorted(cols.items())}))
    for k, label in (("recpct", "REV REC %"), ("bspct", "Budget Spent %")):
        if k not in cols:
            raise SystemExit(
                'ERROR: no column carries the sub-header "%s".\n'
                "       These are system values whose rounding sets the 70%%/100%% "
                "boundaries and cannot be recomputed. Refusing to guess." % label)
    return cols


# ---------- E3000 companion (Northstar Gulf entity), added 2026-09-29 ----------
# From the 09/29/2026 export, the "Transferred Entity" projects keep their budget and
# historic cost in the main (USD) report, but their contract value, invoicing and any
# post-transfer cost move to a separate report for the new entity:
# "FP&A & QC Project Report E3000.xlsx" (in SAR, project IDs suffixed "-Int").
# FP&A lead, 2026-09-29: use that file as well, going forward.
# Merge rule: a main-report project whose ID matches an E3000 project (suffix removed)
# gets E3000 contract / revenue / invoiced / cost ADDED (converted to USD); budgets stay
# from the main report (both files carry the same budget). Budget Spent % is recomputed
# as merged cost / revised budget; REV REC % is recomputed as merged revenue / contract.
# E3000 projects with no main-report row are appended as their own rows (location
# "NH Gulf", which build_overrun_dashboard.MAP rolls into International).
# The companion is picked up automatically when a file whose name contains "E3000"
# sits beside the main report AND carries the same "As of" date (so last week's
# workbook in the same folder is never merged with this week's companion).
SAR_PER_USD = 3.75
COMPANION_TAG = "E3000"
COMPANION_SUFFIX = "-Int"


def _asof(path):
    try:
        ws = _pick_sheet(load_workbook(path, data_only=True, read_only=True))
        return str(ws.cell(3, 1).value or "").strip()
    except Exception:
        return None


def find_companion(path):
    import os, glob
    d = os.path.dirname(os.path.abspath(path))
    me = os.path.abspath(path)
    want = _asof(path)
    for f in sorted(glob.glob(os.path.join(d, "*.xlsx"))):
        if os.path.abspath(f) == me or COMPANION_TAG.lower() not in os.path.basename(f).lower():
            continue
        if os.path.basename(f).startswith("~$"):
            continue
        if _asof(f) == want:
            return f
    return None


def _load_companion(path):
    """E3000 projects in scope, converted to USD, keyed by main-report project ID."""
    global REQUIRE_COMMERCIAL_BASIS
    keep = REQUIRE_COMMERCIAL_BASIS
    REQUIRE_COMMERCIAL_BASIS = False
    try:
        rows, _ = load_active_projects(path, companion=False)
    finally:
        REQUIRE_COMMERCIAL_BASIS = keep
    out = {}
    for m in rows:
        pid = str(m[0])
        key = pid[:-len(COMPANION_SUFFIX)] if pid.endswith(COMPANION_SUFFIX) else pid
        m = list(m)
        for i in (6, 7, 8, 9, 10, 11, 12):
            if isinstance(m[i], (int, float)):
                m[i] = m[i] / SAR_PER_USD
        m[0] = key
        if key in out:   # same project under both "ID" and "ID-Int": add them up
            prev = out[key]
            for i in (6, 7, 8, 9, 10, 11, 12):
                prev[i] = num(prev[i]) + num(m[i])
            rb = num(prev[9])
            prev[17] = num(prev[12]) / rb if rb else 0.0
            prev[16] = num(prev[7]) / num(prev[6]) if num(prev[6]) else 0.0
        else:
            out[key] = m
    return out


def load_active_projects(path, verbose=False, companion=None):
    comp = {}
    comp_path = None
    if companion is not False:
        comp_path = companion or find_companion(path)
        if comp_path:
            comp = _load_companion(comp_path)
    merged = set()
    wb = load_workbook(path, data_only=True)
    ws = _pick_sheet(wb)
    lbl = _labels_row(ws)
    c = _map_cols(ws, lbl)
    rows = []
    nonnum_rec = nonnum_bs = 0
    skip_group = skip_status = skip_ovhd = skip_nobasis = 0
    for r in range(lbl + 1, ws.max_row + 1):
        raw = ws.cell(r, 1).value
        if raw is None or not str(raw).strip():
            continue
        s = str(raw)
        indent = len(s) - len(s.lstrip(" "))
        pid = s.strip()
        if (indent and indent < LEAF_INDENT) or pid.lower().startswith(GROUP_ROW_PREFIXES):
            skip_group += 1
            continue
        status = str(ws.cell(r, 3).value or "").strip()
        if status not in ACTIVE_STATUSES:
            skip_status += 1
            continue
        if "OVHD" in pid.upper():
            skip_ovhd += 1
            continue
        name = ws.cell(r, 2).value
        m = [pid, name.strip() if isinstance(name, str) else name, status,
             ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 6).value,
             ws.cell(r, c["contract"]).value, ws.cell(r, c["revenue"]).value,
             (ws.cell(r, c["invoiced"]).value if "invoiced" in c else None),
             ws.cell(r, c["revbud"]).value,
             ws.cell(r, c["budget"]).value, ws.cell(r, c["etc"]).value,
             ws.cell(r, c["cost"]).value]
        cm = comp.get(pid)
        if cm is not None:
            for i in (6, 7, 8, 12):
                m[i] = num(m[i]) + num(cm[i])
            for i in (9, 10, 11):
                if not num(m[i]):
                    m[i] = cm[i]
            merged.add(pid)
        contract, revenue = num(m[6]), num(m[7])
        revbud, budget, cost = num(m[9]), num(m[10]), num(m[12])
        if REQUIRE_COMMERCIAL_BASIS and (contract <= 0 or (budget <= 0 and revbud <= 0)):
            skip_nobasis += 1
            continue
        recv = ws.cell(r, c["recpct"]).value
        if not isinstance(recv, (int, float)):
            nonnum_rec += 1
            recv = revenue / contract if contract else 0.0
        bsv = ws.cell(r, c["bspct"]).value
        if not isinstance(bsv, (int, float)):
            nonnum_bs += 1
            bsv = cost / revbud if revbud else 0.0
        if cm is not None:
            recv = revenue / contract if contract else 0.0
            bsv = cost / revbud if revbud else 0.0
        m += [revbud - cost, budget - cost, None, recv, bsv]
        rows.append(m)
    n_added = 0
    for key, cm in comp.items():
        if key in merged:
            continue
        contract = num(cm[6])
        if REQUIRE_COMMERCIAL_BASIS and (contract <= 0 or (num(cm[10]) <= 0 and num(cm[9]) <= 0)):
            continue
        cm = list(cm[:13]) + [num(cm[9]) - num(cm[12]), num(cm[10]) - num(cm[12]), None, cm[16], cm[17]]
        rows.append(cm)
        n_added += 1
    if not rows:
        raise SystemExit("ERROR: no rows matched ACTIVE_STATUSES on sheet '%s'. "
                         "Statuses in scope: %s" % (ws.title, sorted(ACTIVE_STATUSES)))
    n = len(rows)
    for cnt, label in ((nonnum_rec, "REV REC %"), (nonnum_bs, "Budget Spent %")):
        if cnt > NONNUM_TOLERANCE * n:
            raise SystemExit(
                'ERROR: %d of %d rows (%.0f%%) have a non-numeric "%s".\n'
                "       Usually this means the export carries uncalculated formulas.\n"
                "       Open the workbook in Excel, let it calculate, save, and re-run.\n"
                "       Refusing to fall back to computed ratios -- it silently shifts "
                "the 70%%/100%% boundary counts." % (cnt, n, cnt / n * 100, label))
    asof = ws.cell(3, 1).value or "As of (date)"
    if verbose:
        print("  sheet=%s labels_row=%d cols=%s"
              % (ws.title, lbl, {k: get_column_letter(v) for k, v in sorted(c.items())}))
        print("  kept=%d  skipped: group/total=%d status=%d overhead=%d no-basis=%d"
              % (n, skip_group, skip_status, skip_ovhd, skip_nobasis))
    if comp_path:
        print("  E3000 companion %s: %d merged into main-report rows, %d added as new rows"
              % (comp_path, len(merged), n_added))
    return rows, str(asof).replace("As of ", "").strip()


# ---------- filters (validated exact, both dates) ----------
def passes_cleanup(m, min_bs):
    contract, revbud = num(m[6]), num(m[9])
    etc, rembud = num(m[11]), num(m[13])
    status, bs = m[2], num(m[17])
    if bs < min_bs:        return False
    if contract == 0:      return False
    if status in FINAL_STATUSES:
        if etc > 0 and rembud > 1.0:               return False
        if revbud > 0 and rembud > 0.10 * revbud:  return False
        return True
    if etc != 0 and rembud > 0.01 * revbud:        return False
    return True


def filter_gt70(rows):   return [m for m in rows if passes_cleanup(m, 0.70)]
def filter_90_50k(rows): return [m for m in rows if passes_cleanup(m, 0.90) and num(m[6]) > 50000]
def filter_100_ff(rows): return [m for m in rows if m[3] == "Fixed fee" and num(m[16]) >= 1.0]


def _hdr_row(ws, ncol, row):
    for c in range(1, ncol + 1):
        cell = ws.cell(row, c)
        cell.font = HDR_F; cell.fill = HDR_FILL
        cell.alignment = CENTER; cell.border = BORDER


def _widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------- All Active Projects (so SUMIF formulas resolve) ----------
def write_all_active(wb, rows, asof):
    ws = wb.create_sheet(SRC_TAB)
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Northstar Heritage, LLC"; ws["A1"].font = TITLE_F
    ws["A2"] = "FP&A & QC Project Report"; ws["A2"].font = SUB_F
    ws["A3"] = f"As of {asof}"; ws["A3"].font = SUB_F
    grp = {7: "Contract Value", 8: "Revenue", 9: "Invoiced Amount",
           10: "Total Revised Budget", 11: "Total Budget",
           12: "Total Budget Adjustment", 13: "Total Costs", 17: "Recognized %"}
    sub = {7: "Inception To Date", 8: "Inception To Date", 9: "Inception To Date",
           10: "Orig Budget + Adjustments", 11: "Original", 12: "ETC",
           13: "Inception To Date", 14: "Remaining Budget",
           15: "Orig Budget Remaining", 17: "REV REC %", 18: "Budget Spent %"}
    for c, t in grp.items(): ws.cell(8, c, t).font = SUB_F
    for c, t in sub.items(): ws.cell(9, c, t).font = SUB_F
    labels = ["Project ID", "Dimension name", "Project status", "Project Billing type",
              "Location Name", "Project Manager Name", "EOP", "Actual", "Actual",
              "Actual", "Actual", "Actual", "Actual", "Remaining Budget",
              "Original Budget Remaining", "teamwork Status", "Actual", "Summary"]
    for c, t in enumerate(labels, 1):
        ws.cell(11, c, t)
    _hdr_row(ws, 18, 11)
    rr = 12
    for m in rows:
        for c in range(1, 19):
            cell = ws.cell(rr, c, m[c - 1]); cell.font = BODY_F
            if 7 <= c <= 15:    cell.number_format = MONEY
            elif c in (17, 18): cell.number_format = PCT
        rr += 1
    _widths(ws, [13, 38, 17, 16, 16, 22, 14, 14, 14, 16, 14, 14, 14, 14, 16, 13, 12, 12])
    ws.freeze_panes = "A12"
    return ws


# ---------- standard filter tab ----------
def write_standard_tab(wb, title, recs, asof, notes, total_label):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws["D1"] = "Summary:"; ws["D1"].font = TITLE_F
    r = 2
    for note in notes:
        ws.cell(r, 4, note).font = NOTE_F; r += 1
    ws.cell(r, 4, total_label).font = BOLD_F
    grp = {7: "Contract Value", 8: "Revenue", 9: "Invoiced Amount",
           10: "Total Revised Budget", 11: "Total Budget",
           12: "Total Budget Adjustment", 13: "Total Costs", 15: "Recognized %"}
    sub = {7: "Inception To Date", 8: "Inception To Date", 9: "Inception To Date",
           10: "Orig Budget + Adjustments", 11: "Original", 12: "ETC",
           13: "Inception To Date", 14: "Inception To Date", 15: "REV REC %",
           16: "Budget Spent %"}
    for c, t in grp.items(): ws.cell(8, c, t).font = SUB_F
    for c, t in sub.items(): ws.cell(9, c, t).font = SUB_F
    for c in range(7, 16): ws.cell(10, c, asof).font = SUB_F
    labels = ["", "Dimension name", "Project status", "Project Billing type",
              "Location Name", "Project Manager Name", "EOP", "Actual", "Actual",
              "Actual", "Actual", "Actual", "Actual", "Remaining Budget",
              "Actual", "Summary"]
    for c, t in enumerate(labels, 1):
        ws.cell(11, c, t)
    _hdr_row(ws, 16, 11)
    src_map = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 16, 17]
    rr = 12
    for m in recs:
        for c, si in enumerate(src_map, 1):
            cell = ws.cell(rr, c, m[si]); cell.font = BODY_F
            if 7 <= c <= 14:    cell.number_format = MONEY
            elif c in (15, 16): cell.number_format = PCT
        rr += 1
    _widths(ws, [13, 40, 17, 16, 16, 22, 14, 14, 14, 16, 14, 14, 14, 16, 12, 12])
    ws.freeze_panes = "A12"
    return ws


# ---------- executive reshape ----------
def write_exec_tab(wb, title, recs, asof, notes, total_label):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws["D1"] = "Summary:"; ws["D1"].font = TITLE_F
    r = 2
    for note in notes:
        ws.cell(r, 4, note).font = NOTE_F; r += 1
    ws.cell(r, 4, total_label).font = BOLD_F
    grp = {6: "Contract Value", 7: "Total Revised Budget", 8: "Original Budget",
           9: "ETC", 10: "Revenue Recognized", 11: "Rev Rec %", 12: "Cost",
           13: "Budget Remaining", 14: "Budget Spent %", 15: "Project Gross Profit"}
    for c, t in grp.items(): ws.cell(8, c, t).font = SUB_F
    for c in range(6, 16): ws.cell(9, c, asof).font = SUB_F
    labels = ["Project ID", "Project Name", "Project Billing type", "Project status",
              "Project Manager Name", "Actual", "Actual", "Actual", "Actual",
              "Actual", "Actual", "Actual", "Actual", "Actual", "Actual", "Location Name"]
    for c, t in enumerate(labels, 1):
        ws.cell(10, c, t)
    _hdr_row(ws, 16, 10)
    rr = 11
    for m in recs:
        vals = [m[0], m[1], m[3], m[2], m[5], num(m[6]), num(m[9]), num(m[10]),
                num(m[11]), num(m[7]), num(m[16]), num(m[12]), num(m[13])]
        for c, v in enumerate(vals, 1):
            ws.cell(rr, c, v).font = BODY_F
        ws.cell(rr, 14, f"=IF(G{rr}=0,0,L{rr}/G{rr})").font = BODY_F
        ws.cell(rr, 15, f"=J{rr}-L{rr}").font = BODY_F
        ws.cell(rr, 16, m[4]).font = BODY_F
        for c in (6, 7, 8, 9, 10, 12, 13, 15): ws.cell(rr, c).number_format = MONEY
        ws.cell(rr, 11).number_format = PCT
        ws.cell(rr, 14).number_format = PCT
        rr += 1
    _widths(ws, [13, 40, 16, 17, 22, 14, 16, 14, 12, 16, 11, 14, 15, 13, 16, 16])
    ws.freeze_panes = "A11"
    return ws


# ---------- weighted overrun (live SUMIF) ----------
def write_weighted_overrun_tab(wb, title, rows):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    for c, t in {1: "PMO", 2: "Project", 3: "Amount", 5: "PMO",
                 6: "Overrun Amount", 7: "Total Revised Budgets",
                 8: "Weighted Overrun"}.items():
        cell = ws.cell(1, c, t)
        cell.font = HDR_F; cell.fill = HDR_FILL
        cell.alignment = CENTER; cell.border = BORDER
    overrun = sorted((m for m in rows if num(m[13]) < 0), key=lambda m: num(m[13]))
    r = 2
    for m in overrun:
        ws.cell(r, 1, str(m[4]).strip()).font = BODY_F
        ws.cell(r, 2, m[1]).font = BODY_F
        c = ws.cell(r, 3, num(m[13])); c.font = BODY_F; c.number_format = MONEY
        r += 1
    F = {}
    for m in overrun:
        loc = str(m[4]).strip(); F[loc] = F.get(loc, 0.0) + num(m[13])
    r = 2
    for loc in sorted(F.keys(), key=lambda k: F[k]):
        ws.cell(r, 5, loc).font = BODY_F
        cf = ws.cell(r, 6, f"=SUMIF($A:$A,E{r},$C:$C)"); cf.font = BODY_F; cf.number_format = MONEY
        cg = ws.cell(r, 7, f"=SUMIF('{SRC_TAB}'!$E:$E,$E{r},'{SRC_TAB}'!$J:$J)"); cg.font = BODY_F; cg.number_format = MONEY
        ch = ws.cell(r, 8, f"=IF(G{r}=0,0,F{r}/G{r})"); ch.font = BODY_F; ch.number_format = PCT3
        r += 1
    _widths(ws, [16, 42, 14, 3, 16, 16, 20, 16])
    ws.freeze_panes = "A2"
    return ws


def build(src, out):
    rows, asof = load_active_projects(src)
    gt70, p9050, ff100 = filter_gt70(rows), filter_90_50k(rows), filter_100_ff(rows)
    notes = ["Removed projects with ETC>1% budget remaining",
             "Removed $0 Contract Value projects",
             "Removed previously identified ETC in final/reportingstatus",
             "Removed projects in final reporting with >10% budget remaining"]
    wb = Workbook(); wb.remove(wb.active)
    write_all_active(wb, rows, asof)
    write_exec_tab(wb, "Projects >70% Exec View", gt70, asof, notes, f"Total: {len(gt70)}")
    write_standard_tab(wb, "Projects >70% ", gt70, asof, notes, f"Total: {len(gt70)}")
    write_standard_tab(wb, "Projects 90% & > $50k", p9050, asof, notes, f"Total: {len(p9050)}")
    write_standard_tab(wb, "100% Rev Rec FF", ff100, asof, [], f"Total:  {len(ff100)}")
    write_weighted_overrun_tab(wb, "Weighted Overrun Calculation", rows)
    # Northstar branding (2026-09-30): ERP project names can still carry the pre-rebrand
    # company name (e.g. "... (Polaris Child)"); the downloadable workbook shows the Northstar name.
    for _ws in wb.worksheets:
        for _row in _ws.iter_rows():
            for _c in _row:
                if isinstance(_c.value, str) and 'polaris' in _c.value.lower():
                    _c.value = (_c.value.replace('Polaris Heritage', 'Northstar Heritage')
                                .replace('POLARIS', 'NORTHSTAR').replace('Polaris', 'Northstar'))
    wb.save(out)
    n_over = sum(1 for m in rows if num(m[13]) < 0)
    print(f"as-of {asof}: rows={len(rows)}  >70%={len(gt70)}  90%&>50k={len(p9050)}  "
          f"100%FF={len(ff100)}  overrun={n_over}  -> {out}")
    return rows, gt70, p9050, ff100


def main():
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    build(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "Rebuilt Budget Tabs.xlsx")


if __name__ == "__main__":
    main()
