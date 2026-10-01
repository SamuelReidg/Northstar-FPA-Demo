#!/usr/bin/env python3
"""
weekly_summary.py  --  one-command weekly budget-utilization summary.

Computes every figure in the weekly email from the "All Active Projects"
export, builds the two overrun-by-PMO tables, and (if last week's export is
supplied) auto-fills the "up from / down from last week" deltas. Then emits
both the rebuilt workbook and the summary Word document.

Everything is computed from the data -- nothing is hardcoded -- and reuses the
validated loader/filters in regenerate_budget_tabs.py.

Usage:
    python weekly_summary.py "<this_week.xlsx>" ["<last_week.xlsx>"]

Outputs (next to this script):
    Rebuilt Budget Tabs.xlsx
    Budget Utilization Summary <date>.docx
    summary_data.json   (intermediate, consumed by build_summary.js)

Open definitions still to confirm with the FP&A lead (clearly marked below):
  1. LATE_STAGE rule -- "have ETCs or are on-budget" among In Progress
     late-stage projects. Current rule lands ~1pp above the reported figures.
  2. ORIGINAL-budget project count -- counts projects with original budget > 0
     (matches the reported overrun %); the manual report has listed a couple more.
Both are isolated in functions so the rule can be swapped without touching the
rest of the pipeline.
"""
import sys, json, subprocess, os
from collections import defaultdict
import regenerate_budget_tabs as eng

n = eng.num
HERE = os.path.dirname(os.path.abspath(__file__))


def metrics(path):
    rows, asof = eng.load_active_projects(path)
    # indices: D billing 3, E loc 4, G contract 6, H rev 7, J revbud 9,
    #          K origbud 10, L etc 11, M cost 12, N rembud 13, recpct 16
    total = len(rows)
    withK = [m for m in rows if n(m[10]) > 0]            # original budget present
    ovK   = [m for m in withK if n(m[12]) > n(m[10])]    # cost > original budget
    sumK  = sum(n(m[10]) for m in withK)
    orig_w = sum(n(m[12]) - n(m[10]) for m in ovK) / sumK * 100 if sumK else 0
    ovJ   = [m for m in rows if n(m[12]) > n(m[9])]      # cost > revised budget
    sumJ  = sum(n(m[9]) for m in rows)
    rev_w = sum(n(m[12]) - n(m[9]) for m in ovJ) / sumJ * 100 if sumJ else 0

    gt70  = len(eng.filter_gt70(rows))
    gt90  = len(eng.filter_90_50k(rows))
    ff100 = len(eng.filter_100_ff(rows))

    def pmo_table(budcol):
        ov, tot = defaultdict(float), defaultdict(float)
        for m in rows:
            loc, bud, cost = str(m[4]).strip(), n(m[budcol]), n(m[12])
            if bud > 0:
                tot[loc] += bud
                if cost > bud:
                    ov[loc] += cost - bud
        res = [[loc, round(ov[loc]), round(tot[loc]), round(ov[loc] / tot[loc] * 100, 2)]
               for loc in tot if ov[loc] > 0]
        res.sort(key=lambda x: -x[3])
        T = sum(x[1] for x in res); Dn = sum(tot.values())
        return res, round(T), round(Dn), round(T / Dn * 100, 2) if Dn else 0

    orig_tbl, oT, oD, oWO = pmo_table(10)
    rev_tbl,  rT, rD, rWO = pmo_table(9)

    return {
        "asof": asof, "total": total,
        "orig_n": len(withK), "orig_overrun_n": len(ovK),
        "orig_overrun_pct": round(len(ovK) / len(withK) * 100, 1) if withK else 0,
        "orig_weighted": round(orig_w, 2),
        "rev_overrun_n": len(ovJ),
        "rev_overrun_pct": round(len(ovJ) / total * 100, 2) if total else 0,
        "rev_weighted": round(rev_w, 2),
        "gt70": gt70, "gt90_50k": gt90, "ff100": ff100,
        "late": late_stage(rows),
        "orig_tbl": orig_tbl, "orig_T": oT, "orig_D": oD, "orig_WO": oWO,
        "rev_tbl": rev_tbl, "rev_T": rT, "rev_D": rD, "rev_WO": rWO,
    }


# ---- OPEN DEFINITION #1 : late-stage "have ETCs or are on-budget" ----
def late_stage(rows):
    """Among In Progress projects at each revenue-recognition threshold, the
    share that either carry an ETC (budget adjustment > 0) or are within the
    revised budget. NOTE: lands ~1pp above the manually reported shares -- confirm
    exact rule (status set, 'has ETC' test, on-budget basis) to finalize."""
    def ok(m):
        return n(m[11]) > 0 or n(m[12]) <= n(m[9])
    def share(thr):
        g = [m for m in rows if m[2] == "In Progress" and n(m[16]) >= thr]
        return round(sum(1 for m in g if ok(m)) / len(g) * 100) if g else 0
    return {"p100": share(1.0), "p90": share(0.90), "p70": share(0.70)}


def delta_up(cur, prev):
    if prev is None: return ""
    if cur > prev:   return f", up from {prev} last week"
    if cur < prev:   return f", down from {prev} last week"
    return ", the same as last week"


def delta_pct(cur, prev):
    if prev is None: return "."
    if abs(cur - prev) < 0.005: return f", the same as last week."
    return (f", {'up' if cur > prev else 'down'} from {prev:.2f}% last week.")


def main():
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    cur = metrics(sys.argv[1])
    prev = metrics(sys.argv[2]) if len(sys.argv) > 2 else None

    # rebuilt workbook for this week
    eng.build(sys.argv[1], os.path.join(HERE, "Rebuilt Budget Tabs.xlsx"))

    p = (lambda k: prev[k] if prev else None)
    payload = {
        **cur,
        "orig_weighted_delta": delta_pct(cur["orig_weighted"], p("orig_weighted")),
        "rev_weighted_delta":  delta_pct(cur["rev_weighted"],  p("rev_weighted")),
        "gt70_delta":  delta_up(cur["gt70"],     p("gt70")),
        "gt90_delta":  delta_up(cur["gt90_50k"], p("gt90_50k")),
        "ff100_delta": delta_up(cur["ff100"],    p("ff100")),
    }
    json.dump(payload, open(os.path.join(HERE, "summary_data.json"), "w"))

    print(f"=== {cur['asof']} ===")
    print(f"Original: {cur['orig_n']} w/ orig budget, {cur['orig_overrun_pct']}% overrun, "
          f"weighted {cur['orig_weighted']}%{payload['orig_weighted_delta']}")
    print(f"Revised : {cur['total']} active, {cur['rev_overrun_pct']}% overrun, "
          f"weighted {cur['rev_weighted']}%{payload['rev_weighted_delta']}")
    print(f">70%={cur['gt70']}{payload['gt70_delta']} | 90%&>50k={cur['gt90_50k']}{payload['gt90_delta']} | "
          f"100%FF={cur['ff100']}{payload['ff100_delta']}")
    print(f"Late-stage (In Progress): 100%+={cur['late']['p100']}%  90%={cur['late']['p90']}%  70%={cur['late']['p70']}%")

    out_doc = os.path.join(HERE, f"Budget Utilization Summary {cur['asof'].replace(', ','-').replace(' ','-')}.docx")
    subprocess.run(["node", os.path.join(HERE, "build_summary.js"), out_doc], check=True)
    print(f"Saved -> {out_doc}")


if __name__ == "__main__":
    main()
