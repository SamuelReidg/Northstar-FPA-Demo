#!/usr/bin/env python3
"""
Regenerate Early Warning System data blocks, FILTERED TO IN PROGRESS PROJECTS ONLY.

Source workbook: "SR of Projects @ Risk - Automated <date>.xlsx", sheet "Project Details ".
Usage:  python regenerate_ews_data.py "SR of Projects @ Risk - Automated 9.29.26.xlsx" "September 26, 2026"
        (the as-of label defaults to $FPA_AS_OF, formatted "Month D, YYYY", when it is set)

Outputs (written to the current directory):
  ews_data_prefix.txt  -> the DATA object prefix (everything up to ', "map_path"')
  ews_byoffice.txt     -> `const BYOFFICE = {...};`
  ews_offsite.txt      -> the offsite block, e.g. {"total": 3, "crit": 1}
  ews_projects.txt     -> `window.PROJECTS = [...];` for the Projects tab (project-recovery.html)
and prints the aggregates incl. crit_pmos / total_pmos for the "PMOs with Critical Risk" tile.

HOW THE DASHBOARD IS UPDATED: the regenerated blocks are applied to "early-warning-system.html"
and "project-recovery.html" (tools/apply_ews_blocks.py does this for the demo; the workflow
notes describe the equivalent hand edit).

Definitions (reverse-engineered from the live dashboard and validated against it):
  budget    = Revised Budget
  cost_pct  = "Task Overrun %"  * 100   (= Total Task Overrun $ / Revised Budget)
  hours_pct = "Hours Overrun %" * 100
  ovr$      = Total Task Overrun ($)
  type      = "Cost"  if cost_pct >= hours_pct else "Hours"
  tier(m = max(cost_pct, hours_pct)):  m>=25 Critical | m>=10 Warning | m>=5 Watch | else On Track
  cost_var_pct  = sum(ovr$  for at-risk) / sum(budget    for at-risk) * 100
  hours_var_pct = sum(ovrHrs for at-risk) / sum(actualHrs for at-risk) * 100
  dot color (dom) by an office's critical share s:  s>=.25 Critical | s>=.10 Warning | s>=.05 Watch | else On Track
  (the map's color legend therefore uses the SAME 25/10/5 numbers as the project tier bands.)
  "at-risk" = any tier other than On Track.
"""
import openpyxl, json, sys, os, datetime
from collections import defaultdict, Counter

SRC = sys.argv[1] if len(sys.argv) > 1 else 'SR of Projects @ Risk 5.31.26.xlsx'
if len(sys.argv) > 2:
    AS_OF = sys.argv[2]
elif os.environ.get('FPA_AS_OF'):      # reproducible builds: FPA_AS_OF=YYYY-MM-DD
    _d = datetime.date.fromisoformat(os.environ['FPA_AS_OF'])
    AS_OF = _d.strftime('%B ') + str(_d.day) + _d.strftime(', %Y')
else:
    AS_OF = 'May 31, 2026'
STATUS_FILTER = 'In Progress'

# Offices that have a dot on the map: name -> (x, y) in the 944x496 viewBox
# (x = 680 + 4*lon, y = 422.285 - 229.127*ln(tan(pi/4 + lat/2)) -- the base map's projection).
OFFICE_XY = {
 "Red Rock":(233.4,271.7),"Sierra":(194.8,258.0),"Cascade":(187.6,225.7),"International":(897.5,321.4),
 "High Desert":(256.2,269.3),"Maritime":(370.0,291.8),"Front Range":(259.7,244.4),"Great Lakes":(322.4,231.1),
 "Big Sky":(235.8,216.5),"Chesapeake":(374.0,252.8),"Coastal Bend":(354.8,296.9),"Meridian Central":(306.8,261.8),
 "Lakeshore":(369.6,230.6),"Meridian Southwest":(200.9,251.8),"Hill Country":(303.4,267.5),"Piedmont":(360.8,267.4),
 "Gulf Coast":(290.4,306.5),"Bluegrass":(342.0,257.6),"Prairie":(290.6,259.3),"Built Environment":(260.7,253.5),
 "Island":(413.6,349.0),"Paleontology":(236.1,286.1),
}
# Source "Location Name" -> mapped office (with roll-ups). Anything not listed -> offsite.
LOC2OFFICE = {
 "Red Rock":"Red Rock","Great Lakes":"Great Lakes","Coastal Bend":"Coastal Bend","Sierra":"Sierra",
 "Cascade":"Cascade","High Desert":"High Desert","Front Range":"Front Range","Big Sky":"Big Sky",
 "Chesapeake":"Chesapeake","Lakeshore":"Lakeshore","Hill Country":"Hill Country","Piedmont":"Piedmont",
 "Gulf Coast":"Gulf Coast","Bluegrass":"Bluegrass","Prairie":"Prairie","Island":"Island",
 "Built Environment":"Built Environment","Paleontology":"Paleontology","Maritime":"Maritime",
 "Meridian Central":"Meridian Central","Meridian Southwest":"Meridian Southwest",
 "International":"International","NH Gulf (USD)":"International",
}

def num(v):
    try: return float(v)
    except: return None

wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
ws = wb['Project Details ']
projs = []
for r in ws.iter_rows(min_row=5, values_only=True):
    if not r[0]:
        continue
    if (str(r[3]).strip() if r[3] else '') != STATUS_FILTER:
        continue
    cp = (num(r[10]) or 0) * 100
    hp = (num(r[16]) or 0) * 100
    m = max(cp, hp)
    tier = 'Critical' if m >= 25 else 'Warning' if m >= 10 else 'Watch' if m >= 5 else 'On Track'
    loc = (str(r[4]).strip() if r[4] else '')
    projs.append({'id': str(r[0]).strip(), 'name': (str(r[1]).strip() if r[1] else ''),
                  'office': LOC2OFFICE.get(loc, None), 'loc': loc, 'pm': (str(r[5]).strip() if r[5] else ''),
                  'budget': num(r[7]) or 0, 'totcost': num(r[8]) or 0, 'ovr': num(r[9]) or 0, 'cost_pct': cp,
                  'budhrs': num(r[13]) or 0, 'acthrs': num(r[14]) or 0, 'ovrhrs': num(r[15]) or 0, 'hours_pct': hp,
                  'type': 'Cost' if cp >= hp else 'Hours', 'tier': tier})

total = len(projs)
TIERS = ['Critical', 'Warning', 'Watch', 'On Track']
tc_all = Counter(p['tier'] for p in projs)
tier_counts = {t: tc_all.get(t, 0) for t in TIERS}
tier_pct = {t: round(tier_counts[t] / total * 100, 1) if total else 0 for t in TIERS}
atrisk = [p for p in projs if p['tier'] != 'On Track']
cost_ovr = round(sum(max(p['ovr'], 0) for p in atrisk))
bud_base = sum(p['budget'] for p in atrisk if p['budget'] > 0)
cost_var_pct = round(cost_ovr / bud_base * 100, 1) if bud_base else 0
hours_ovr = round(sum(max(p['ovrhrs'], 0) for p in atrisk))
acth_base = sum(p['acthrs'] for p in atrisk if p['acthrs'] > 0)
hours_var_pct = round(hours_ovr / acth_base * 100, 1) if acth_base else 0
by_ovr = sorted(projs, key=lambda p: p['ovr'], reverse=True)
crit = [p for p in by_ovr if p['tier'] == 'Critical']
alerts = [{'id': p['id'], 'name': p['name'], 'type': p['type'],
           'var': round(p['cost_pct'] if p['type'] == 'Cost' else p['hours_pct'], 1), 'tier': p['tier']}
          for p in crit[:6]]
top = [{'id': p['id'], 'name': p['name'], 'type': p['type'], 'budget': round(p['budget']),
        'cost_pct': round(p['cost_pct'], 1), 'hours_pct': round(p['hours_pct'], 1),
        'ovr': round(p['ovr']), 'tier': p['tier'], 'pm': p['pm'], 'loc': p['loc']} for p in by_ovr[:10]]
office_groups = defaultdict(list); offsite = []
for p in projs:
    if p['office'] and p['office'] in OFFICE_XY:
        office_groups[p['office']].append(p)
    else:
        offsite.append(p)

def dom_tier(cn, tot):
    s = cn / tot if tot else 0
    return 'Critical' if s >= 0.25 else 'Warning' if s >= 0.10 else 'Watch' if s >= 0.05 else 'On Track'

geo = []
for office, (x, y) in OFFICE_XY.items():
    grp = office_groups.get(office, [])
    if not grp:
        continue
    tc = Counter(p['tier'] for p in grp)
    geo.append({"loc": office, "x": x, "y": y, "total": len(grp),
                "ovr": round(sum(max(p['ovr'], 0) for p in grp)),
                "dom": dom_tier(tc.get('Critical', 0), len(grp)),
                "Critical": tc.get('Critical', 0), "Warning": tc.get('Warning', 0),
                "Watch": tc.get('Watch', 0), "On Track": tc.get('On Track', 0)})
offsite_block = {"total": len(offsite), "crit": sum(1 for p in offsite if p['tier'] == 'Critical')}
offices_with = [o for o in OFFICE_XY if office_groups.get(o)]
total_pmos = len(offices_with)
crit_pmos = sum(1 for o in offices_with if any(p['tier'] == 'Critical' for p in office_groups[o]))

data_prefix = ('const DATA = {"as_of": ' + json.dumps(AS_OF) + ', "total": ' + str(total) + ', '
               '"tiers": ' + json.dumps(tier_counts) + ', "tier_pct": ' + json.dumps(tier_pct) + ', '
               '"cost_var_pct": ' + str(cost_var_pct) + ', "cost_ovr": ' + str(cost_ovr) + ', '
               '"hours_var_pct": ' + str(hours_var_pct) + ', "hours_ovr": ' + str(hours_ovr) + ', '
               '"active_alerts": ' + str(tier_counts['Critical']) + ', '
               '"alerts": ' + json.dumps(alerts) + ', "top": ' + json.dumps(top) + ', '
               '"geo": ' + json.dumps(geo, separators=(",", ":")))
open('ews_data_prefix.txt', 'w').write(data_prefix)
TLET = {'Critical': 'C', 'Warning': 'W', 'Watch': 'T', 'On Track': 'O'}
byoffice = {}
for office, grp in office_groups.items():
    rows = sorted(grp, key=lambda p: p['ovr'], reverse=True)[:25]
    byoffice[office] = [[p['id'], p['name'], ('C' if p['type'] == 'Cost' else 'H'), round(p['budget']),
                         round(p['cost_pct']), round(p['hours_pct']), TLET[p['tier']]] for p in rows]
open('ews_byoffice.txt', 'w').write('const BYOFFICE = ' + json.dumps(byoffice, separators=(',', ':')) + ';')
open('ews_offsite.txt', 'w').write(json.dumps(offsite_block))

# ---- window.PROJECTS for the Projects tab (project-recovery.html): In Progress at-risk only ----
# Row schema: [id,name,type(C/H),costVar%,hoursVar%,costOvr$,hoursOvr#,budget,FAC$,hoursAtCompl,PM,location,tier(C/W/T)]
TLET2 = {'Critical': 'C', 'Warning': 'W', 'Watch': 'T'}
atrisk_sorted = sorted(atrisk, key=lambda p: (max(p['cost_pct'], p['hours_pct']), p['ovr']), reverse=True)
proj_rows = [[p['id'], p['name'], ('C' if p['type'] == 'Cost' else 'H'),
             round(p['cost_pct']), round(p['hours_pct']), round(p['ovr']), round(p['ovrhrs']),
             round(p['budget']), round(p['budget'] + p['ovr']), round(p['budhrs'] + p['ovrhrs']),
             p['pm'], p['loc'], TLET2[p['tier']]] for p in atrisk_sorted]
open('ews_projects.txt', 'w').write('window.PROJECTS = ' + json.dumps(proj_rows, separators=(',', ':'), ensure_ascii=False) + ';')

print('total:', total)
print('tiers:', tier_counts, 'sum=', sum(tier_counts.values()))
print('tier_pct:', tier_pct, 'sum=', round(sum(tier_pct.values()), 1))
print('cost_var_pct:', cost_var_pct, 'cost_ovr:', cost_ovr)
print('hours_var_pct:', hours_var_pct, 'hours_ovr:', hours_ovr)
print('active_alerts:', tier_counts['Critical'], 'offsite:', offsite_block)
print('crit_pmos / total_pmos:', crit_pmos, '/', total_pmos)
print('geo offices:', len(geo), 'sum geo totals:', sum(g['total'] for g in geo),
      '+ offsite', offsite_block['total'], '=', sum(g['total'] for g in geo) + offsite_block['total'])
print('window.PROJECTS (at-risk) rows:', len(proj_rows))
