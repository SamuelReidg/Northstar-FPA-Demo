#!/usr/bin/env python3
"""
apply_categories.py — patch the PMO & Employee Dashboard payloads (ed3Data/ed2Data)
of the merged utilization site with data the Employee Dashboard sheet leaves broken:

  (1) Admin / PTO / S&M  — blank in the sheet; sourced from the 'Raw Data' tab.
  (2) Top 6 Projects     — '#N/A' in the sheet for most employees in recent exports;
                           rebuilt from 'Raw Data' hours by Project ID, named via the
                           'OTC-Project Profitability CM to' tab (≈100% coverage) and
                           sector via 'Proj. Sector' (partial — older projects show no
                           sector label, which is expected).

Billable / Non-Billable are already correct in the payload and left untouched.

Category classification (per Raw Data row; Utilized=col5, Hours=col7, Task=col10, WeekEnd=col18):
  utilized (col5>0): 'Overhead Non-Posting' -> Non-Billable ; else -> Billable (already in payload)
  unutilized (col5==0): PTO list -> PTO ; S&M list -> S&M ; any other 'OVHD' task -> Admin (catch-all)

Usage:  python apply_categories.py "<Utilization Report.xlsx>" <in.html> <out.html>
Needs employees.json (from extract_data.py) in cwd.
"""
import sys, json, re, datetime, collections
from python_calamine import CalamineWorkbook

norm = lambda s: re.sub(r'\s+',' ',str(s)).strip().lower()
PTO = {norm(x) for x in ["OVHD - Labor - Holiday","OVHD - Labor - Vacation (Salaried)","OVHD - Labor - PTO (Hourly Employee)","OVHD - Labor - Sick (Salaried/Temps)","OVHD - Labor - Parental Leave"]}
SM  = {norm(x) for x in ["OVHD - Labor - Business Development","OVHD - Labor - Marketing","OVHD - Labor - Proposal Writing","OVHD - Production - Proposal Editing"]}
sid=lambda x: str(int(x)) if isinstance(x,(int,float)) else str(x).strip()

def main():
    xlsx, in_html, out_html = sys.argv[1], sys.argv[2], sys.argv[3]
    html=open(in_html,encoding="utf-8").read()
    def getp(did):
        p=f'<script id="{did}" type="application/json">'; i=html.find(p); j=html.find('</script>',i)
        return json.loads(html[i+len(p):j])
    ed2,ed3=getp('ed2Data'),getp('ed3Data')
    NW=len(ed2['date_iso']); idx={d:k for k,d in enumerate(ed2['date_iso'])}
    def wi(we):
        if isinstance(we,(datetime.datetime,datetime.date)): we=we.strftime("%Y-%m-%d")
        else: we=str(we)[:10]
        return idx.get(we)
    id2off={sid(e['id']):e['office'] for e in json.load(open("employees.json"))['employees']}

    wb=CalamineWorkbook.from_path(xlsx)
    # project name/sector lookups: prefer Projects.csv (full master), fall back to workbook sheets
    pname={}; psec={}
    for r in wb.get_sheet_by_name("OTC-Project Profitability CM to").to_python(skip_empty_area=False)[6:]:
        if r and r[0] not in (None,'','All Projects'): pname[sid(r[0])]=str(r[1]).strip() if r[1] else ''
    for r in wb.get_sheet_by_name("Proj. Sector").to_python(skip_empty_area=False)[1:]:
        if r and r[0] not in (None,''): psec[sid(r[0])]=str(r[6]).strip() if len(r)>6 and r[6] else ''
    import os, csv as _csv
    if os.path.exists("Projects.csv"):
        for r in _csv.DictReader(open("Projects.csv",encoding="utf-8-sig")):
            pid=(r.get("Project ID") or "").strip()
            if not pid: continue
            nm=(r.get("Project name") or "").strip(); sc=(r.get("Project Sector") or "").strip()
            if nm: pname[pid]=nm
            if sc: psec[pid]=sc

    cat=collections.defaultdict(lambda:{c:[0.0]*NW for c in ('admin','pto','sm')})
    eproj=collections.defaultdict(lambda: collections.defaultdict(float))   # emp -> pid -> hrs
    for r in wb.get_sheet_by_name("Raw Data").to_python(skip_empty_area=False)[1:]:
        if not r or r[0] in (None,''): continue
        eid=sid(r[0]); u=r[5] if isinstance(r[5],(int,float)) else 0
        h=r[7] if len(r)>7 and isinstance(r[7],(int,float)) else 0
        pid=sid(r[4]) if r[4] not in (None,'') else ''
        if pid not in ('','OVHD') and h: eproj[eid][pid]+=h
        if u and u>0: continue
        if not h: continue
        task=norm(r[10]) if len(r)>10 and r[10] not in (None,'') else ''
        if   task in PTO: c='pto'
        elif task in SM:  c='sm'
        elif task.startswith('ovhd'): c='admin'
        else: continue
        k=wi(r[18] if len(r)>18 else None)
        if k is not None: cat[eid][c][k]+=h
    rnd=lambda a:[round(x,2) for x in a]; Z=[0.0]*NW
    def top6(pidh):
        return [{"name":(pname.get(p) or p),"sector":psec.get(p,'')} for p,_ in sorted(pidh.items(),key=lambda x:-x[1])[:6]]

    for e in ed2['employees']:
        eid=sid(e['id']); c=cat.get(eid)
        if c:
            e['w_admin'],e['w_pto'],e['w_sm']=rnd(c['admin']),rnd(c['pto']),rnd(c['sm'])
            e['admin'],e['pto'],e['sm']=round(sum(c['admin']),2),round(sum(c['pto']),2),round(sum(c['sm']),2)
        else:
            e['w_admin']=Z[:]; e['w_pto']=Z[:]; e['w_sm']=Z[:]; e['admin']=e['pto']=e['sm']=0
        e['projects']=top6(eproj[eid]) if eproj.get(eid) else []
    # office aggregation
    offc=collections.defaultdict(lambda:{c:[0.0]*NW for c in ('admin','pto','sm')})
    offp=collections.defaultdict(lambda: collections.defaultdict(float))
    for eid,c in cat.items():
        off=id2off.get(eid)
        if off:
            for k in ('admin','pto','sm'):
                for w in range(NW): offc[off][k][w]+=c[k][w]
    for eid,ph in eproj.items():
        off=id2off.get(eid)
        if off:
            for p,h in ph.items(): offp[off][p]+=h
    for o in ed3['employees']:
        a=offc.get(o['name'])
        if a:
            o['w_admin'],o['w_pto'],o['w_sm']=rnd(a['admin']),rnd(a['pto']),rnd(a['sm'])
            o['admin'],o['pto'],o['sm']=round(sum(a['admin']),2),round(sum(a['pto']),2),round(sum(a['sm']),2)
        else:
            o['w_admin']=Z[:]; o['w_pto']=Z[:]; o['w_sm']=Z[:]; o['admin']=o['pto']=o['sm']=0
        o['projects']=top6(offp[o['name']]) if offp.get(o['name']) else []
    def setp(S,did,obj):
        p=f'<script id="{did}" type="application/json">'; i=S.find(p); j=S.find('</script>',i)
        return S[:i+len(p)]+json.dumps(obj,ensure_ascii=False)+S[j:]
    html=setp(html,'ed2Data',ed2); html=setp(html,'ed3Data',ed3)
    open(out_html,"w",encoding="utf-8").write(html)
    na=sum(1 for d in (ed2,ed3) for x in d['employees'] for pr in x.get('projects',[]) if str(pr.get('name')).strip() in ('#N/A','','None'))
    print(f"apply_categories: cats+projects patched; residual blank/#N/A project tiles={na}")

if __name__=="__main__": main()
