#!/usr/bin/env python3
"""
Regenerate ed3_extra.js — the PMO Dashboard "Key Observations" + "Backlog" data.

Backlog = SUM( Contract Value [col H, Inception-To-Date]  -  Revenue Recognized [col W, ITD] )
          over projects whose Project status [col D] is In Progress or Transferred Entity,
          grouped by Location Name [col AA] mapped to the dashboard PMOs.

Key Observations are auto-generated as short, professional-but-conversational notes:
a utilization read (what's going well, or direction to improve — redeploy / business development)
plus a pipeline note from the backlog (light book -> BD focus; deep book -> execution; one
dominant project -> wrap-up / client-concentration exposure). No dollar/percent figures and no
revenue-recognized-vs-contract-value language. Tone is measured and honest — not overly positive.
The text is editable by hand afterward.

Trend uses recent momentum (last 3 monthly actuals), so an office that peaked and rolled over
reads as "easing back / declining" rather than "climbing".

Usage:
    python regenerate_pmo_extra.py "Utilization Report.xlsx" [util_data.js] [out.js]

Source sheet: "OTC-Project Profitability CM to"  (headers in rows 2-5, data from row 7)
Output:       ed3_extra.js (one PMO per line) -> window.ED3_EXTRA = { "<PMO>": {...}, ... }
"""
import sys, json, hashlib
from collections import defaultdict
import openpyxl

SHEET = "OTC-Project Profitability CM to"
PMOS = ["Big Sky","Bluegrass","Built Environment","Cascade","Chesapeake","Coastal Bend","Corporate",
"Front Range","Geospatial","Great Lakes","Gulf Coast","High Desert","Hill Country","International","Island",
"Lakeshore","Maritime","Meridian Central","Meridian Southwest","Paleontology","Piedmont","Prairie",
"Red Rock","Sales & Marketing","Sierra"]
# OTC "Location Name" -> dashboard PMO. Same-name locations map to themselves; the rest are
# roll-ups consistent with the Early Warning System dashboard. Adjust here if offices change.
MAP = {**{p: p for p in PMOS},
    "NH Gulf (USD)":"International","NH Gulf":"International",
    "Global LLC International":"International","International (Global LLC)":"International"}
# Excluded (no PMO on the dashboard): Sales and Marketing (the ERP location, as opposed to the
# Sales & Marketing department), Archive Services, blank.
# Case-insensitive lookup: the OTC tab has logged some locations in a different case (e.g. a
# lowercased 'nh gulf (usd)' for part of the International book). Match on .lower() so they
# aren't dropped.
MAPL = {k.lower(): v for k, v in MAP.items()}

# Statuses counted as active for the backlog panels. "Transferred Entity" is new in the
# 09/2026 export: most of the International book moved out of "In Progress" into it, which
# silently removed that whole book from the backlog. FP&A's ruling 2026-09-14 is that it
# counts as active/in progress, consistent with ACTIVE_STATUSES in regenerate_budget_tabs.py.
# Compared lowercased.
ACTIVE_STATUSES = {"in progress", "transferred entity"}
# ---------------------------------------------------------------------------
# Backlog basis (changed 2026-08-19)
# ---------------------------------------------------------------------------
# The headline total counts POSITIVE backlog only (remaining work still to be
# earned). Projects where Revenue Recognized EXCEEDS Contract Value produce a
# negative figure (e.g. a change-order project recognized well past its contract
# value). Netting those into the headline made the panel impossible to reconcile:
# an office's headline read well below the sum of its visible rows, because the
# most negative project sorts last and fell outside the old 12-row cap, so it
# never rendered.
#
# Now: total = positive backlog only, and negative-backlog projects are ALWAYS
# listed (appended after the positives) so over-recognized work stays visible
# instead of silently depressing the number. Their magnitude also travels as
# `negbk` for the sub-caption. Zero-backlog projects are omitted from the list
# (no information, and they pushed real rows out of view) but still counted in
# `count`, which remains every In Progress row incl. the office OVHD line.
PROJECT_CAP = 200  # was 12; comfortably above the largest office's positive list, so the
                   # rows now always reconcile to the headline.
DEEP_BACKLOG = 2_000_000
THIN_BACKLOG = 250_000
CONCENTRATION = 0.45
MANY_PROJECTS = 60


def load_util(path):
    if not path: return {}, {}
    try:
        t = open(path, encoding="utf-8").read()
        d = json.loads(t[t.index("{"):t.rindex("}")+1])
        return d.get("util_targets", {}), d.get("util_actuals", {})
    except Exception as e:
        print(f"  (util_data.js unavailable: {e}; observations omit utilization)", file=sys.stderr)
        return {}, {}


def pick(pmo, options):
    """Stable per-PMO choice so wording varies between offices but doesn't change run to run."""
    return options[int(hashlib.md5(pmo.encode()).hexdigest(), 16) % len(options)]


def util_line(pmo, TGT, ACT):
    a = [x for x in ACT.get(pmo, []) if x is not None]
    tg = TGT.get(pmo, [])
    if not a or not tg: return None
    act = sum(a)/len(a); tgt = sum(tg[:len(a)])/len(a); d = (act - tgt) * 100
    # recent momentum over the last 3 actuals (falls back to full span if <3 points)
    recent = a[-1] - a[-3] if len(a) >= 3 else a[-1] - a[0]
    tr = "rising" if recent >= 0.02 else ("slipping" if recent <= -0.02 else "steady")
    band = "strong" if d >= 3 else ("on" if d > -3 else ("soft" if d > -8 else "weak"))
    L = {
    ("strong","rising"): [f"{pmo} is performing well, running above target and continuing to improve. The priority now is sustaining it — keep the pipeline full and monitor workloads so the team isn't overextended.",
        f"{pmo} is in a strong position and still trending up, comfortably above target. Focus on keeping the next wave of work ready so the momentum holds."],
    ("strong","steady"): [f"{pmo} is holding steadily above target — a solid, stable position. Maintaining a consistent flow of work is the main priority.",
        f"{pmo} continues to run above target with little fluctuation. Things are in good shape; the focus is sustaining the workload."],
    ("strong","slipping"): [f"{pmo} remains above target but has eased back over the past several weeks. Worth confirming upcoming work is lined up so the recent softening doesn't carry on.",
        f"{pmo} is still above target, though utilization has trended down recently. A look at the near-term schedule would help keep it from sliding further."],
    ("on","rising"): [f"{pmo} is tracking at target and trending upward — a healthy position. A deeper pipeline would help push it further ahead.",
        f"{pmo} has improved to right around target. Continuing to bring in work should move it higher."],
    ("on","steady"): [f"{pmo} is holding at target — a stable position. The emphasis is on keeping work flowing consistently so it doesn't drift.",
        f"{pmo} is sitting at target with little movement. Staying close to the pipeline will keep it there."],
    ("on","slipping"): [f"{pmo} is near target but has softened over recent weeks. A good time to review the upcoming workload and firm up the pipeline before it slips further.",
        f"{pmo} is around target with a recent downward drift. Worth getting ahead of it — line up the next round of work and keep staff billable."],
    ("soft","rising"): [f"{pmo} is below target but improving. Continue converting sold work into billable hours to keep closing the gap.",
        f"{pmo} remains under target, though the trend is moving in the right direction. Maintaining proposal activity and staffing onto billable work should help."],
    ("soft","steady"): [f"{pmo} has been running below target with little movement. Worth a closer look at staffing and pipeline — redeploying staff to busier offices or focusing on proposals would help lift billable time.",
        f"{pmo} is holding below target. Some direction would help here: redeploy where possible and prioritize landing the next round of work."],
    ("soft","slipping"): [f"{pmo} is below target and continuing to slip. This warrants attention — prioritize getting staff onto billable work, redeploy where possible, and step up business development.",
        f"{pmo} is trending down and sitting under target. Focus on moving people onto billable projects and accelerating proposals to reverse it."],
    ("weak","rising"): [f"{pmo} is well below target, though showing some recovery. Keep pushing — move staff onto billable work and pursue the pipeline actively to rebuild utilization.",
        f"{pmo} remains well under target but is recovering. Sustain that by redeploying available staff and maintaining strong proposal activity."],
    ("weak","steady"): [f"{pmo} is well below target and flat. This needs direction — redeploy staff toward busier teams and prioritize business development to lift billable time.",
        f"{pmo} continues to run well short of target with little change. Rebalancing staff toward busier offices and focusing on the pipeline should be the priority."],
    ("weak","slipping"): [f"{pmo} is well below target and still declining, which makes it a priority for attention. Redeploy staff to where the work is, manage the bench closely, and focus on winning new work.",
        f"{pmo} is under target and trending lower — a pressing situation. Move people onto billable work, keep the bench in check, and prioritize new business."],
    }
    return pick(pmo, L[(band, tr)])


def backlog_line(pmo, pl, tot=None):
    # `tot` is the positive-only backlog (the headline). Falls back to the net sum
    # so the function still works if called the old way.
    if tot is None:
        tot = sum(p["bk"] for p in pl)
    n = len(pl)
    if n == 0:
        return pick(pmo, ["There are no geographic projects mapped to this group, so utilization is the primary lever.",
            "With no project backlog mapped here, the focus is on how the team's time is deployed."])
    top = max(pl, key=lambda p: p["bk"]); share = (top["bk"]/tot if tot else 0)
    if share >= CONCENTRATION and top["bk"] > 0:
        return pick(pmo, ["The backlog is concentrated in one large project, so it's worth lining up follow-on work early to avoid a gap when it winds down — and to limit the office's exposure to a single client.",
            "A single large project accounts for most of the backlog; plan for its completion now and watch the client concentration that comes with it."])
    if tot < THIN_BACKLOG:
        return pick(pmo, ["The backlog is relatively light, so business development should remain a priority.",
            "There isn't much backlog behind the team, which keeps pipeline and proposals front of mind."])
    if tot > DEEP_BACKLOG or n > MANY_PROJECTS:
        return pick(pmo, ["There's a substantial backlog to deliver, so the emphasis shifts to staffing and execution, with attention to margins as projects progress.",
            "A strong backlog of active work is in hand; the focus is execution and ensuring nothing slips as jobs near completion."])
    return pick(pmo, ["There's a reasonable backlog to staff against, which should support utilization as current work winds down.",
        "A solid book of work is in place; worth tracking which projects are wrapping to keep the team transitioning smoothly."])


def main():
    xlsx = sys.argv[1] if len(sys.argv) > 1 else "Utilization Report.xlsx"
    util = sys.argv[2] if len(sys.argv) > 2 else "util_data.js"
    out  = sys.argv[3] if len(sys.argv) > 3 else "ed3_extra.js"
    TGT, ACT = load_util(util)
    ws = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)[SHEET]
    # 0-based col idx: A=0 B=1 C=2 D=3 H=7 W=22 AA=26 AC=28
    proj = defaultdict(list); unmapped = defaultdict(lambda: [0.0, 0])
    for r in ws.iter_rows(min_row=7, values_only=True):
        pid = r[0]
        if pid in (None, "", "All Projects"): continue
        if str(r[3]).strip().lower() not in ACTIVE_STATUSES: continue
        bk = (r[7] or 0) - (r[22] or 0); loc = r[26]
        pmo = MAPL.get(str(loc).strip().lower()) if loc is not None else None
        if pmo is None:
            unmapped[str(loc)][0] += bk; unmapped[str(loc)][1] += 1; continue
        proj[pmo].append({"id": str(pid).strip(), "name": (str(r[1]).strip() if r[1] else ""),
                          "pm": (str(r[28]).strip() if r[28] else ""), "bk": round(bk, 0)})
    extra = {}
    for pmo in PMOS:
        pl  = sorted(proj.get(pmo, []), key=lambda p: -p["bk"])
        pos = [p for p in pl if p["bk"] > 0]
        neg = [p for p in pl if p["bk"] < 0]
        pos_total = sum(p["bk"] for p in pos)
        neg_total = sum(p["bk"] for p in neg)          # negative or 0
        shown = pos[:PROJECT_CAP] + neg                # negatives ALWAYS shown
        parts = [x for x in (util_line(pmo, TGT, ACT), backlog_line(pmo, pl, pos_total)) if x]
        extra[pmo] = {"backlog": round(pos_total, 0), "negbk": round(-neg_total, 0),
                      "count": len(pl),
                      "projects": [{"id": p["id"], "name": p["name"], "pm": p["pm"], "bk": p["bk"]} for p in shown],
                      "obs": " ".join(parts)}
    # Emit one PMO per line (not one giant line) so Claude's file tools can write it
    # straight to the shared network drive — no drag-drop step needed.
    with open(out, "w", encoding="utf-8") as f:
        f.write("/* PMO backlog + observations — generated by regenerate_pmo_extra.py from the\n")
        f.write("   'OTC-Project Profitability CM to' tab. Observation text is editable. */\n")
        f.write("window.ED3_EXTRA = {\n")
        keys = list(extra.keys())
        for i, k in enumerate(keys):
            tail = "," if i < len(keys) - 1 else ""
            f.write(json.dumps(k, ensure_ascii=False) + ": " + json.dumps(extra[k], ensure_ascii=False) + tail + "\n")
        f.write("};\n")
    g = sum(extra[p]["backlog"] for p in PMOS); u = sum(v[0] for v in unmapped.values())
    n = sum(extra[p]["negbk"] for p in PMOS)
    print(f"Wrote {out}: {len(PMOS)} PMOs, positive backlog ${g:,.0f}, "
          f"over-recognized ${n:,.0f}, unmapped ${u:,.0f} "
          f"({', '.join(sorted(unmapped)) or 'none'}).")
    # ---- INVARIANT: every PMO's listed positive rows must sum to its headline. ----
    # This is the whole point of the 2026-08-19 change. If it ever fails, the panel is
    # lying again (a row is contributing to the total without being rendered), so fail
    # LOUDLY rather than print a warning somebody scrolls past. $1 tolerance for rounding.
    bad = []
    for p in PMOS:
        shown = sum(x["bk"] for x in extra[p]["projects"] if x["bk"] > 0)
        if abs(shown - extra[p]["backlog"]) > 1:
            bad.append(f"{p}: rows ${shown:,.0f} vs headline ${extra[p]['backlog']:,.0f}")
    if bad:
        raise SystemExit(
            "\n*** BACKLOG RECONCILIATION FAILED — DO NOT DEPLOY THIS FILE ***\n  "
            + "\n  ".join(bad)
            + f"\n\nLikely cause: a PMO now has more than PROJECT_CAP ({PROJECT_CAP}) "
              "positive-backlog projects, so real rows are being truncated out of the\n"
              "list while still counting toward the headline. Raise PROJECT_CAP and re-run."
        )
    print(f"  reconciliation: OK — all {len(PMOS)} PMOs' listed rows tie to their headline")


if __name__ == "__main__":
    main()
