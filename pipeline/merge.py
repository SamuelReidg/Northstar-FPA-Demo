import re, json, sys
# Weekly MERGE: fresh data sections from the build_site.py donor, spliced into the LIVE page
# (which carries the hand-applied design).  See the operating manual, "MERGE".
#
#   python3 merge.py                              normal weekly refresh
#   python3 merge.py --refresh-carried-forward    also take the donor's versions of the two
#                                                 carried-forward tabs (only when their retired
#                                                 source sheets were extracted into site_data/,
#                                                 e.g. rebuilding the site from a fresh package)
REFRESH_CARRIED = '--refresh-carried-forward' in sys.argv
LIVE = open('LIVE_index.html', encoding='utf-8').read()
DONOR = open('Northstar Heritage Utilization Site.html', encoding='utf-8').read()

def section_spans(html):
    starts = [(m.start(), m.group(1)) for m in re.finditer(r'<section id="([^"]+)"', html)]
    ends = [m.end() for m in re.finditer(r'</section>', html)]
    assert len(starts) == len(ends) == 16, (len(starts), len(ends))
    spans = []
    for i,(s,sid) in enumerate(starts):
        e = ends[i]
        assert e > s
        if i+1 < len(starts): assert e <= starts[i+1][0], f"overlap {sid}"
        spans.append((sid,s,e))
    return spans

live_spans = section_spans(LIVE)
donor_spans = section_spans(DONOR)
assert [x[0] for x in live_spans]==[x[0] for x in donor_spans], "order mismatch"
donor_txt = {sid: DONOR[s:e] for sid,s,e in donor_spans}

KEEP_ENTIRE={'Summary_OVHD_FT_Employee','Recovery_Planning'}
if REFRESH_CARRIED:
    for sid in sorted(KEEP_ENTIRE):
        assert '<tbody>' in donor_txt[sid] and '<tr' in donor_txt[sid], \
            f"--refresh-carried-forward: the donor's {sid} is an empty placeholder (was its retired sheet extracted?)"
    KEEP_ENTIRE=set()
KEEP_PAYLOAD={'PMO_Dashboard','Employee_Dashboard'}

merged=LIVE
reps=[]
for sid,s,e in live_spans:
    if sid in KEEP_ENTIRE or sid in KEEP_PAYLOAD: continue
    reps.append((s,e,donor_txt[sid]))
for s,e,new in sorted(reps,key=lambda x:x[0],reverse=True):
    merged=merged[:s]+new+merged[e:]

def extract(html,sid):
    m=re.search(r'(<script id="%s" type="application/json">)(.*?)(</script>)'%sid,html,re.S)
    assert m,f"{sid} not found"; return m.group(2)
def swap(html,sid,nj):
    pat=re.compile(r'(<script id="%s" type="application/json">)(.*?)(</script>)'%sid,re.S)
    h,n=pat.subn(lambda m:m.group(1)+nj+m.group(3),html,count=1)
    assert n==1,f"swap {sid}={n}"; return h

ed3=extract(DONOR,'ed3Data'); ed2=extract(DONOR,'ed2Data')
json.loads(ed3); json.loads(ed2)
merged=swap(merged,'ed3Data',ed3)
merged=swap(merged,'ed2Data',ed2)
open('MERGED.html','w',encoding='utf-8').write(merged)
print("MERGED bytes:",len(merged))
print("sections:",len(section_spans(merged)))
print("fresh-replaced:",len(reps))
print("kept-live entire:",sorted(KEEP_ENTIRE) or "none (--refresh-carried-forward)")
print("payload-swapped:",sorted(KEEP_PAYLOAD))
