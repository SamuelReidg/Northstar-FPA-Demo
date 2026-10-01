import json
from collections import defaultdict, OrderedDict
ed = json.load(open('site_data/Employee_Dashboard.json'))['rows']
emp = json.load(open('employees.json'))
def pn(v):
    try: return float(v)
    except (TypeError, ValueError): return 0.0
id2off = {}
for e in emp['employees']:
    eid = e.get('id')
    if eid is None: continue
    ids = str(int(eid)) if isinstance(eid, (int, float)) else str(eid)
    id2off[ids] = (e.get('office') or '').strip()
wk = json.load(open('site_data/PMO_Weekly.json'))['rows']
office_order = []
for ri in range(5, 31):
    if ri < len(wk) and len(wk[ri]) > 1 and wk[ri][1]:
        nm = str(wk[ri][1]).strip()
        if nm and nm.lower() != 'total': office_order.append(nm)
NWK = 52
def newrow():
    return {'util':[0.0]*NWK,'non':[0.0]*NWK,'tot':[0.0]*NWK,'U':0.0,'T':0.0,'admin':0.0,'pto':0.0,'sm':0.0,'ot':0.0,'tasks':defaultdict(float),'projects':OrderedDict(),'wbill':[0.0]*NWK,'wunbill':[0.0]*NWK,'wadmin':[0.0]*NWK,'wpto':[0.0]*NWK,'wsm':[0.0]*NWK}
acc = {o: newrow() for o in office_order}
unknown = defaultdict(int); maxnzweek = -1
for r in ed[2:]:
    if not r: continue
    rid = r[0]
    if rid in (None, ''): continue
    ids = str(int(rid)) if isinstance(rid, (int, float)) else str(rid)
    off = id2off.get(ids)
    if off not in acc:
        if off: unknown[off] += 1
        continue
    a = acc[off]
    for j in range(NWK):
        u = pn(r[1+j] if 1+j < len(r) else 0); n = pn(r[53+j] if 53+j < len(r) else 0); t = pn(r[105+j] if 105+j < len(r) else 0)
        a['util'][j]+=u; a['non'][j]+=n; a['tot'][j]+=t; a['U']+=u; a['T']+=t
        if u>0 and j>maxnzweek: maxnzweek=j
        a['wbill'][j]+=pn(r[196+j] if 196+j<len(r) else 0); a['wunbill'][j]+=pn(r[248+j] if 248+j<len(r) else 0)
        a['wadmin'][j]+=pn(r[300+j] if 300+j<len(r) else 0); a['wpto'][j]+=pn(r[352+j] if 352+j<len(r) else 0); a['wsm'][j]+=pn(r[404+j] if 404+j<len(r) else 0)
    a['admin']+=pn(r[158] if len(r)>158 else 0); a['pto']+=pn(r[159] if len(r)>159 else 0); a['sm']+=pn(r[160] if len(r)>160 else 0); a['ot']+=pn(r[161] if len(r)>161 else 0)
    for k in range(8):
        tnm = r[162+k*2] if 162+k*2<len(r) else None; thr = pn(r[163+k*2] if 163+k*2<len(r) else 0)
        if isinstance(tnm,str) and tnm.strip() and thr>0: a['tasks'][tnm.strip()]+=thr
    for k in range(6):
        pnm = r[184+k] if 184+k<len(r) else None; sec = r[190+k] if 190+k<len(r) else None
        if isinstance(pnm,str) and pnm.strip():
            key=pnm.strip()
            if key not in a['projects']: a['projects'][key]=[0, sec if isinstance(sec,str) else '']
            a['projects'][key][0]+=1
print('Unknown offices:', dict(unknown)); print('Max nonzero util week (0-based):', maxnzweek)
# Overtime KPI matches the Overtime tab (Summary - OT sheet, by PROJECT LOCATION, col 13 = "Total")
# instead of employee-home-office OT (Employee Dashboard col 161). Exact office row only — the
# sub-office "(PMO)" lines (e.g. "<Sub-office> (<PMO>)") are NOT rolled up, per FP&A
# decision 2026-07-28, so the card mirrors exactly what the Overtime tab displays.
import json as _json
OT_BY_OFFICE = {}
try:
    for _r in _json.load(open('site_data/Summary_OT.json'))['rows']:
        if _r and isinstance(_r[0], str) and _r[0].strip():
            _nm = _r[0].strip()
            if _nm.lower() in ('total','grand total','historical offices','office name'): continue
            _v = _r[13] if len(_r) > 13 and isinstance(_r[13], (int, float)) else None
            if _v is not None: OT_BY_OFFICE.setdefault(_nm.lower(), _v)
    print('OT override loaded from Summary_OT.json for', len(OT_BY_OFFICE), 'offices')
except Exception as _e:
    print('  [warn] OT override from Summary_OT.json failed:', _e)
def build_row(name,a):
    row=[None]*458; row[0]=name
    for j in range(NWK):
        row[1+j]=round(a['util'][j],4); row[53+j]=round(a['non'][j],4); row[105+j]=round(a['tot'][j],4)
    row[157]=(a['U']/a['T']) if a['T']>0 else None
    row[158]=round(a['admin'],4); row[159]=round(a['pto'],4); row[160]=round(a['sm'],4)
    row[161]=round(OT_BY_OFFICE.get(name.strip().lower(), 0.0), 4)  # OT from Overtime tab (Summary - OT col13), not home-office
    for k,(tnm,thr) in enumerate(sorted(a['tasks'].items(),key=lambda kv:-kv[1])[:8]): row[162+k*2]=tnm; row[163+k*2]=round(thr,2)
    for k,(pnm,(cnt,sec)) in enumerate(sorted(a['projects'].items(),key=lambda kv:(-kv[1][0],kv[0]))[:6]): row[184+k]=pnm; row[190+k]=sec
    for j in range(NWK):
        row[196+j]=round(a['wbill'][j],4); row[248+j]=round(a['wunbill'][j],4); row[300+j]=round(a['wadmin'][j],4); row[352+j]=round(a['wpto'][j],4); row[404+j]=round(a['wsm'][j],4)
    row[456]=round(sum(a['wbill']),2); row[457]=round(sum(a['wunbill']),2)
    return row
rows=[ed[0],ed[1]]
for o in office_order: rows.append(build_row(o,acc[o]))
json.dump({'tab':'PMO Dashboard','rows':rows}, open('site_data/PMO_Dashboard.json','w'))
print('Wrote site_data/PMO_Dashboard.json with', len(rows)-2, 'office rows')
