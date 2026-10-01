#!/usr/bin/env python3
"""
regenerate_emp_extra.py — inject richer, role- & trend-aware Key Observations into
the Employee Dashboard's embedded ed2Data payload (no separate file).

For each employee, adds to its ed2Data object:
  obs : 2-3 sentence narrative that reads the data like an analyst —
        * role context (Job Title from 'Current Employee Roster'): leadership/BD roles
          are expected to split time into management + business development, so a lower
          billable rate is partly by design; production roles should be mostly billable.
        * the driver of any gap to target (PTO / admin / S&M), and the *trend* — e.g.
          S&M ramping up as the office's utilization (PMO Weekly) softens => leaning into
          contract generation.
        * office backlog context.  No hard numbers; editable.
  pmo : office->PMO key for the Office Backlog panel (reads ED3_EXTRA).

Inputs: ed2Data (post apply_categories) + ed3Data (office weekly util) from the HTML,
ed3_extra.js (backlog), and the workbook (roster Job Title). Rewrites HTML in place.

Usage:  python regenerate_emp_extra.py <site.html> <Utilization Report.xlsx> [ed3_extra.js]
"""
import sys, json, hashlib
from python_calamine import CalamineWorkbook

MAP_ROLLUP = {"NH Gulf (USD)":"International","NH Gulf":"International",
 "Global LLC International":"International",
 # display-only split of the International roster (build_site.EMPLOYEE_OFFICE_OVERRIDES):
 # the backlog panel and observations still read the International book.
 "International (Global LLC)":"International"}
THIN=250_000; DEEP=2_000_000; CONC=0.45
sid=lambda x: str(int(x)) if isinstance(x,(int,float)) else str(x).strip()
def pick(key,opts): return opts[int(hashlib.md5(key.encode()).hexdigest(),16)%len(opts)]

def role_bucket(title):
    t=(title or '').lower()
    if not t: return 'unknown'
    if any(k in t for k in ['principal','director','president','vp','vice president','chief','officer','program manager','service line']):
        return 'leader'
    if any(k in t for k in ['lead','manager','supervisor']):
        return 'lead'
    if any(k in t for k in ['archaeolog','paleontolog','historian','technician','gis','geolog','scientist','engineer','specialist','analyst','editor','designer','mountaineer','ecosystem','architect','biologist','curator']):
        return 'production'
    return 'other'

def trend(arr, totarr=None):
    """direction of a weekly series over active weeks; if totarr given, use share."""
    act=[k for k,t in enumerate(totarr or arr) if (t or 0)>0]
    if len(act)<4: return 'flat'
    vals=[]
    for k in act:
        v=arr[k] or 0
        if totarr is not None:
            tt=totarr[k] or 0; v=(v/tt) if tt>0 else 0
        vals.append(v)
    h=len(vals)//2
    early=sum(vals[:h])/max(h,1); late=sum(vals[h:])/max(len(vals)-h,1)
    if late-early>=0.04: return 'rising'
    if early-late>=0.04: return 'falling'
    return 'flat'

def util_trend(util_pct):
    v=[x for x in util_pct if x is not None]
    if len(v)<4: return 'flat'
    h=len(v)//2; early=sum(v[:h])/h; late=sum(v[h:])/(len(v)-h)
    if late-early>=0.04: return 'rising'
    if early-late>=0.04: return 'softening'
    return 'flat'

def make_obs(emp, title, office_util_pct, band_bk):
    eid=emp['id']; bucket=role_bucket(title)
    ytd=emp.get('ytd_util'); tgt=emp.get('target')
    adm=emp.get('admin') or 0; pto=emp.get('pto') or 0; sm=emp.get('sm') or 0
    bil=emp.get('billable') or 0; non=emp.get('unbillable') or 0
    tot=adm+pto+sm+bil+non
    share=lambda v: (v/tot) if tot>0 else 0
    d=(ytd-tgt) if (ytd is not None and tgt is not None) else None
    sm_tr=trend(emp.get('w_sm',[]), emp.get('total_wk'))
    pto_tr=trend(emp.get('w_pto',[]), emp.get('total_wk'))
    off_tr=util_trend(office_util_pct)
    title_nice=(title or '').strip()
    art='an' if title_nice[:1].lower() in 'aeiou' else 'a'
    P=[]

    # 1) role + standing
    if bucket=='leader':
        if d is None or d>=-0.05:
            P.append(pick(eid+'1',[
                f"As {art} {title_nice}, time is expected to span billable work, running the office, and winning new business, so a lower billable rate is partly by design — and they're tracking about where that role should land.",
                f"The {title_nice} role is a leadership seat, so the split across delivery, management, and business development matters more than a raw billable rate — and the current level reads as on-plan for the role."]))
        else:
            P.append(pick(eid+'1',[
                f"As {art} {title_nice}, some distance below the billable target is expected, but the current gap is wider than leadership duties alone would explain.",
                f"Even allowing for the management and business-development side of {art} {title_nice}'s role, billable time is running lighter than expected."]))
    elif bucket in ('production','lead'):
        if d is not None and d< -0.03:
            P.append(pick(eid+'1',[
                f"The {title_nice} role is meant to be largely billable, so the shortfall to target is worth a look.",
                f"For {art} {title_nice}, most hours should land on billable projects — so running under target stands out."]))
        else:
            P.append(pick(eid+'1',[
                f"Billable project work makes up the bulk of this {title_nice}'s time — right where it should be for the role.",
                f"As {art} {title_nice}, they're carrying a healthy billable load, on or above target."]))
    else:
        P.append(pick(eid+'1',[
            "Most of this person's time is going to billable project work." if (d is None or d>=-0.03) else "Billable time is running below target this year."]))

    # 2) the analytical driver + trend
    if bucket=='leader' and (share(sm)>=0.12 or sm_tr=='rising'):
        if sm_tr=='rising' and off_tr=='softening':
            P.append(pick(eid+'2',[
                "Tellingly, business-development time has stepped up through the year just as the office's utilization softened — a deliberate lean into contract generation to refill the pipeline, which is exactly where a principal's hours should go when work tightens.",
                "The standout pattern: as the office's utilization eased, their proposal and business-development time climbed — a clear pivot toward bringing in work, and the right instinct for the role."]))
        elif sm_tr=='rising':
            P.append(pick(eid+'2',[
                "Notably, business-development time has ramped up over the year — a deliberate push to generate new work and keep the office's pipeline full, the right focus for a leadership seat.",
                "Their proposal and business-development time has climbed as the year progressed, pointing to an increasing focus on winning the next round of work."]))
        else:
            P.append(pick(eid+'2',[
                "A meaningful share of the year has gone to business development and proposals — consistent with a leader keeping the office's pipeline full.",
                "Notable proposal and business-development time suggests an active focus on generating the next round of work."]))
    elif d is not None and d<-0.03:
        if pto>=adm and pto>=sm and pto>0:
            P.append(pick(eid+'2',[
                "The time below target traces mostly to paid time off rather than any shortage of work — likely a temporary, calendar-driven dip.",
                "Most of the gap is paid time off, so billable capacity itself looks intact."]))
        elif adm>=sm and adm>=pto and adm>0:
            P.append(pick(eid+'2',[
                "Administrative and overhead work is the largest non-billable draw — worth checking whether some of it can be redirected toward project work.",
                "A sizable slice of time is going to admin and overhead; some of that may be re-deployable onto billable projects."]))
        elif sm>0 and (share(sm)>=0.12 or sm_tr=='rising'):
            P.append(pick(eid+'2',[
                "Much of the non-billable time is going to proposals and business development, which turns a softer billable stretch into pipeline-building.",
                "The non-billable time skews toward business development — a constructive use of the slower period."]))
        else:
            P.append(pick(eid+'2',[
                "There's no single overhead or leave driver behind the gap, so it's worth confirming enough billable work is queued up.",
                "Nothing obvious (leave or overhead) explains the shortfall, which points to making sure the project pipeline is keeping them busy."]))

    # 3) backlog context
    if band_bk=='light':
        P.append(pick(eid+'3',["With the office's backlog on the lighter side, that focus on new business is well-timed.",
            "The office's book of work is light, so pipeline-building is the right priority."]))
    elif band_bk=='deep':
        P.append(pick(eid+'3',["There's a solid book of work in the office to deliver, so converting it into billable hours is the priority.",
            "The office is carrying a strong backlog — plenty to staff against."]))
    elif band_bk=='concentrated':
        P.append("The office's backlog leans heavily on one large project, so lining up follow-on work matters.")
    elif band_bk=='moderate':
        P.append("The office has a reasonable backlog to staff against.")
    return " ".join(P)

def main():
    htmlpath=sys.argv[1]; xlsx=sys.argv[2]; ed3p=sys.argv[3] if len(sys.argv)>3 else "ed3_extra.js"
    html=open(htmlpath,encoding="utf-8").read()
    def get(did):
        p=f'<script id="{did}" type="application/json">'; i=html.find(p); j=html.find('</script>',i); return json.loads(html[i+len(p):j]), i+len(p), j
    ed2,a2,b2=get('ed2Data'); ed3,_,_=get('ed3Data')
    off_util={o['name']:o.get('util_pct_wk',[]) for o in ed3['employees']}
    t=open(ed3p,encoding="utf-8").read(); ED3=json.loads(t[t.index("{"):t.rindex("}")+1]); keys=set(ED3)
    # roster titles
    title={}
    for r in CalamineWorkbook.from_path(xlsx).get_sheet_by_name("Current Employee Roster").to_python(skip_empty_area=False)[1:]:
        if r and r[0] not in (None,''): title[sid(r[0])]=str(r[7]).strip() if len(r)>7 and r[7] else ''
    def resolve(off): 
        if off in keys: return off
        m=MAP_ROLLUP.get(off); return m if m in keys else off
    def band(pmo):
        e=ED3.get(pmo)
        if not e or not e.get('count'): return 'none'
        tt=e.get('backlog',0); top=max((x.get('bk',0) for x in e.get('projects',[])),default=0)
        if tt>0 and top/tt>=CONC: return 'concentrated'
        if tt<THIN: return 'light'
        if tt>DEEP: return 'deep'
        return 'moderate'
    for emp in ed2['employees']:
        pmo=resolve(emp.get('office') or ''); emp['pmo']=pmo
        emp['obs']=make_obs(emp, title.get(emp['id'],''), off_util.get(pmo,[]), band(pmo))
    html=html[:a2]+json.dumps(ed2,ensure_ascii=False)+html[b2:]
    open(htmlpath,"w",encoding="utf-8").write(html)
    print(f"injected role/trend-aware obs for {len(ed2['employees'])} employees; roster matched {sum(1 for e in ed2['employees'] if title.get(e['id']))}")

if __name__=="__main__": main()
