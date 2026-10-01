#!/usr/bin/env python3
"""build_overrun_dashboard.py -- elegant, data-driven Project Dashboard (PMO Performance).
Every figure computed from the weekly "All Active Projects" export; auto-fits the viewport;
3-column PMO ranking with click-through project buildup. Location roll-ups via MAP.
Usage: python build_overrun_dashboard.py "<this_week.xlsx>" ["<last_week.xlsx>"]"""
import sys, os, html, json
from collections import defaultdict
import regenerate_budget_tabs as eng
n = eng.num
HERE = os.path.dirname(os.path.abspath(__file__))

# Location -> consolidated PMO roll-ups. The export's "Location Name" is the ERP location,
# which is finer than (or older than) the PMO the dashboard reports on: sub-locations,
# retired office codes, the acquired firm's pre-integration location and the international
# entities all roll up into one PMO row. Any location not listed passes through as its own PMO.
MAP = {
    "PIEDMONT - INACTIVE": "Piedmont", "Piedmont - Inactive": "Piedmont",   # retired office code
    "Sierra North": "Sierra", "Sierra South": "Sierra",
    "Gulf Coast East": "Gulf Coast", "Gulf Coast West": "Gulf Coast",
    "NH Gulf (USD)": "International",
    "NH Gulf": "International",                     # E3000 companion rows (2026-09-29)
    "Global LLC International": "International",
    "Meridian": "Meridian Central",                 # acquired firm's pre-integration location
}
def pmoname(loc):
    return MAP.get(loc, loc)


# Companion workbook produced alongside the page and offered by the Export button.
XLSX_FILE = "budget-watchtower-workbook.xlsx"

MONTHS = {"january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
          "july":7,"august":8,"september":9,"october":10,"november":11,"december":12}


def download_name(asof):
    """Friendly, dated filename for the browser download (falls back gracefully)."""
    try:
        month, day, year = asof.replace(",", "").split()
        return "Northstar Heritage - Budget Watchtower %02d-%02d-%s.xlsx" % (
            MONTHS[month.lower()], int(day), year)
    except Exception:
        return "Northstar Heritage - Budget Watchtower.xlsx"


XPORT_JS = """
(function(){
  var b=document.getElementById('xbtn'), t=document.getElementById('xtoast');
  if(!b||!t) return;
  function toast(m,ms){t.textContent=m;t.classList.add('show');
    clearTimeout(t._z);t._z=setTimeout(function(){t.classList.remove('show');},ms||2600);}
  b.addEventListener('click',function(){toast('Downloading the Budget Watchtower workbook\\u2026');});
})();
"""


def compute(path):
    rows, asof = eng.load_active_projects(path)
    total = len(rows)
    withK = [m for m in rows if n(m[10]) > 0]
    ovK   = [m for m in withK if n(m[12]) > n(m[10])]
    sumK  = sum(n(m[10]) for m in withK)
    ov_amt = sum(n(m[12]) - n(m[10]) for m in ovK)
    orig_w = ov_amt / sumK * 100 if sumK else 0
    ovJ   = [m for m in rows if n(m[12]) > n(m[9])]
    sumJ  = sum(n(m[9]) for m in rows)
    rev_w = sum(n(m[12]) - n(m[9]) for m in ovJ) / sumJ * 100 if sumJ else 0

    ov, tot = defaultdict(float), defaultdict(float)
    for m in rows:
        loc = pmoname(str(m[4]).strip()); bud = n(m[10]); cost = n(m[12])
        if bud > 0:
            tot[loc] += bud
            if cost > bud:
                ov[loc] += cost - bud
    pmo = [[loc, ov[loc], tot[loc], ov[loc] / tot[loc] * 100] for loc in tot if ov[loc] > 0]
    pmo.sort(key=lambda x: -x[1])

    detail = defaultdict(list)
    for m in rows:
        loc = pmoname(str(m[4]).strip()); bud = n(m[10]); cost = n(m[12])
        if bud > 0 and cost > bud:
            detail[loc].append([str(m[1]).strip(), round(bud), round(cost), round(cost - bud)])
    for k in detail:
        detail[k].sort(key=lambda x: -x[3])

    def ls(thr):
        g = [m for m in rows if m[2] == "In Progress" and n(m[16]) >= thr]
        ok = [m for m in g if n(m[11]) > 0 or n(m[12]) <= n(m[9])]
        return round(len(ok) / len(g) * 100) if g else 0

    return {
        "asof": asof, "total": total, "orig_n": len(withK),
        "orig_overrun_pct": round(len(ovK) / len(withK) * 100, 1) if withK else 0,
        "orig_weighted": orig_w, "sumK": sumK, "ov_amt": ov_amt,
        "rev_overrun_pct": round(len(ovJ) / total * 100, 2) if total else 0,
        "rev_weighted": rev_w,
        "gt70": len(eng.filter_gt70(rows)), "gt90": len(eng.filter_90_50k(rows)),
        "ff100": len(eng.filter_100_ff(rows)),
        "late": (ls(1.0), ls(0.90), ls(0.70)),
        "pmo": pmo, "detail": dict(detail),
    }


def millions(v):
    return "${:,.2f}M".format(v / 1e6)


def pmo_rows(rows):
    out = []
    for (loc, o, t, p) in rows:
        out.append('<tr><td class="pn">%s</td><td class="n">$%s</td><td class="n">($%s)</td><td class="n over">%.2f%%</td></tr>'
                   % (html.escape(str(loc)), format(t, ",.0f"), format(abs(o), ",.0f"), p))
    return "\n".join(out)


def gauge(pct):
    C = 326.7; frac = min(abs(pct) / 10.0, 1.0); off = C * (1 - frac)
    return ('<svg viewBox="0 0 130 130" class="gauge">\n'
            '          <circle cx="65" cy="65" r="52" class="g-track"/>\n'
            '          <circle cx="65" cy="65" r="52" class="g-val" stroke-dasharray="326.7" '
            'stroke-dashoffset="%.1f" transform="rotate(135 65 65)"/>\n'
            '          <text x="65" y="63" class="g-num">%.2f%%</text>\n'
            '          <text x="65" y="80" class="g-cap">VS &le;5%% TARGET</text>\n'
            '        </svg>' % (off, pct))


def delta_phrase(cur, prev, fmt):
    if prev is None: return "", ""
    if cur < prev: return "&darr; from %s last week" % fmt(prev), "dn"
    if cur > prev: return "&uarr; from %s last week" % fmt(prev), "up"
    return "no change from last week", "flat"


def stat_delta(cur, prev):
    if prev is None: return ""
    if cur > prev: return "&uarr; from %d last week" % prev
    if cur < prev: return "&darr; from %d last week" % prev
    return "no change from last week"


def up(cur, prev, fmt=lambda x: str(x)):
    if prev is None: return ""
    if cur > prev: return ", up from %s last week" % fmt(prev)
    if cur < prev: return ", down from %s last week" % fmt(prev)
    return ", the same as last week"


HEAD = '''<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Budget Watchtower — Northstar Heritage</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#F7F5F0; --paper:#FCFBF7; --forest:#264D3D; --forest-soft:#2F5C49;
  --sage:#9AAE9A; --sage-tint:#E7EBE3; --brass:#B08D57; --terra:#B65E4A;
  --ink:#2D2D2D; --muted:#7C7A72; --hair:#ECE8DE;
}
*{box-sizing:border-box}
html,body{margin:0;height:100%;overflow:hidden}
body{background:var(--bg);color:var(--ink);
  font-family:"Inter","Helvetica Neue",Helvetica,Arial,sans-serif;
  font-size:14px;line-height:1.5;-webkit-font-smoothing:antialiased;}
#fitviewport{height:100vh;width:100vw;display:flex;align-items:center;justify-content:center;overflow:hidden}
#fit{transform-origin:center center;will-change:transform}
.serif{font-family:"Cormorant Garamond",Georgia,"Times New Roman",serif}
.num{font-variant-numeric:tabular-nums}
a.back{display:inline-flex;gap:7px;align-items:center;color:var(--muted);
  text-decoration:none;font-size:12px;letter-spacing:.04em;margin:0 0 0 40px}
a.back:hover{color:var(--forest)}
.topbar{display:flex;align-items:center;justify-content:space-between;padding:0 40px 0 0}
a.xbtn{display:inline-flex;gap:7px;align-items:center;text-decoration:none;cursor:pointer;
  background:var(--paper);color:var(--forest);border:1px solid var(--hair);border-radius:999px;
  padding:6px 15px;font-size:11px;letter-spacing:.08em;text-transform:uppercase;font-weight:600;
  box-shadow:0 1px 2px rgba(38,77,61,.04);transition:background .15s,border-color .15s,color .15s}
a.xbtn:hover{background:var(--forest);border-color:var(--forest);color:#fff}
a.xbtn:hover svg{stroke:#fff}
a.xbtn svg{width:13px;height:13px;fill:none;stroke:var(--brass);stroke-width:1.8;
  stroke-linecap:round;stroke-linejoin:round;transition:stroke .15s}
.xtoast{position:fixed;left:50%;bottom:26px;transform:translateX(-50%) translateY(8px);
  background:var(--forest);color:#fff;padding:10px 18px;border-radius:999px;font-size:12px;
  letter-spacing:.03em;box-shadow:0 12px 30px rgba(38,77,61,.30);opacity:0;visibility:hidden;
  transition:opacity .18s,transform .18s;z-index:60}
.xtoast.show{opacity:1;visibility:visible;transform:translateX(-50%) translateY(0)}
.wrap{max-width:1560px;margin:0 auto;padding:4px 40px 10px}
.zone{margin:0 0 18px}
.ztag{display:flex;align-items:baseline;gap:12px;margin:0 0 12px}
.ztag .no{font-family:"Cormorant Garamond",serif;font-size:14px;color:var(--brass);font-weight:600;letter-spacing:.12em}
.ztag h2{margin:0;font-size:12px;letter-spacing:.24em;text-transform:uppercase;color:var(--forest);font-weight:600}
.ztag .rule{flex:1;height:1px;background:var(--hair)}
.rankhint{font-size:10px;color:var(--muted);letter-spacing:.04em;white-space:nowrap}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}
.topgrid{display:grid;grid-template-columns:1.15fr repeat(4,1fr);gap:16px;align-items:stretch}
.titleblock{display:flex;flex-direction:column;justify-content:center;gap:8px;padding:2px 4px}
.titleblock .eyebrow{color:var(--brass);font-size:10px;letter-spacing:.32em;font-weight:600;text-transform:uppercase}
.titleblock h1{font-weight:600;color:var(--forest);font-size:34px;line-height:.98;margin:0;letter-spacing:-.01em}
.kpi{background:var(--paper);border-radius:14px;padding:16px 18px;
  box-shadow:0 1px 2px rgba(38,77,61,.04),0 10px 26px rgba(38,77,61,.05);
  display:flex;flex-direction:column;gap:7px}
.kpi .ico{width:20px;height:20px;stroke:var(--sage);fill:none;stroke-width:1.6}
.kpi .v{font-family:"Cormorant Garamond",serif;font-weight:600;font-size:34px;line-height:.9;color:var(--forest)}
.kpi .v.neg{color:var(--terra)}
.kpi .l{font-size:10.5px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);font-weight:600}
.kpi .s{font-size:11.5px;color:var(--muted);line-height:1.35;margin-top:-2px}
.kpi .up{color:var(--terra)} .kpi .down{color:var(--forest-soft)}
.kpi.ring{align-items:center;text-align:center;justify-content:space-between;gap:4px}
.gauge{width:92px;height:92px;margin:0}
.g-track{fill:none;stroke:var(--sage-tint);stroke-width:9}
.g-val{fill:none;stroke:var(--brass);stroke-width:9;stroke-linecap:round}
.g-num{font-family:"Cormorant Garamond",serif;font-size:30px;fill:var(--forest);text-anchor:middle;font-weight:600}
.g-cap{font-size:7px;fill:var(--muted);text-anchor:middle;letter-spacing:.22em;font-weight:600}
.narr{display:grid;grid-template-columns:1fr 1fr 2fr;gap:16px}
.lscompliance{margin-top:14px;padding-top:14px;border-top:1px solid var(--hair)}
.lstitle{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:var(--forest);font-weight:700;margin-bottom:12px}
.lscompliance .bar{max-width:470px}
.ncard{background:var(--paper);border-radius:14px;padding:18px 22px;
  box-shadow:0 1px 2px rgba(38,77,61,.04),0 10px 26px rgba(38,77,61,.05)}
.ncard .hd{display:flex;align-items:center;gap:10px;margin-bottom:12px}
.ncard .hd .dotlbl{font-size:11px;letter-spacing:.16em;text-transform:uppercase;color:var(--forest);font-weight:700}
.tip{position:relative;cursor:help;border-bottom:1px dotted var(--sage)}
.tipbox{position:absolute;left:0;top:calc(100% + 9px);width:300px;background:var(--forest);color:#fff;
  padding:12px 14px;border-radius:10px;font-size:11.5px;line-height:1.55;letter-spacing:0;
  text-transform:none;font-weight:400;box-shadow:0 12px 30px rgba(38,77,61,.30);
  opacity:0;visibility:hidden;transform:translateY(-4px);transition:opacity .15s,transform .15s;z-index:30;pointer-events:none}
.tipbox::before{content:"";position:absolute;left:22px;top:-6px;width:12px;height:12px;background:var(--forest);transform:rotate(45deg)}
.tipbox b{display:block;margin-bottom:7px;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--sage-tint);font-weight:700}
.tipbox em{display:block;margin-top:7px;font-style:normal;color:var(--sage-tint);font-size:10.5px}
.tip:hover .tipbox{opacity:1;visibility:visible;transform:translateY(0)}
.ncard .hd .tgt{margin-left:auto;font-size:10.5px;color:var(--muted);letter-spacing:.04em}
.lead{display:flex;align-items:flex-end;gap:16px;margin-bottom:4px}
.lead .big{font-family:"Cormorant Garamond",serif;font-size:38px;font-weight:600;line-height:.85;color:var(--forest)}
.lead .delta{font-size:12px;margin-bottom:5px}
.lead .delta.dn{color:var(--forest-soft)} .lead .delta.up{color:var(--terra)} .lead .delta.flat{color:var(--muted)}
.lead .cap{font-size:10.5px;letter-spacing:.16em;text-transform:uppercase;color:var(--muted);margin-bottom:6px;font-weight:600}
.say{font-size:12.5px;line-height:1.5;color:#46453F;margin:10px 0 12px;max-width:46ch}
.rsum{margin:14px 0 0;padding-left:16px;font-size:11.5px;line-height:1.45;color:#46453F}
.rsum>li{margin:5px 0}
.rsum b{color:var(--forest)}
.rsum ul{margin:4px 0 0;padding-left:15px;list-style:circle}
.rsum ul li{margin:3px 0;color:var(--muted);font-size:11px}
.bar{margin:0 0 10px}
.bar:last-child{margin-bottom:0}
.bar .br{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px}
.bar .br .pct{font-family:"Cormorant Garamond",serif;font-size:17px;color:var(--forest);font-weight:600}
.bar .br .bl{font-size:9.5px;color:var(--muted);text-align:right;max-width:62%;line-height:1.25}
.track{height:6px;border-radius:6px;background:var(--sage-tint);overflow:hidden}
.fill{height:100%;border-radius:6px;background:var(--forest)}
.rankgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;align-items:stretch}
.rank{background:var(--paper);border-radius:14px;padding:6px 10px 4px;
  box-shadow:0 1px 2px rgba(38,77,61,.04),0 10px 26px rgba(38,77,61,.05)}
table{width:100%;border-collapse:collapse}
thead th{font-size:10px;letter-spacing:.13em;text-transform:uppercase;color:var(--muted);
  font-weight:600;text-align:right;padding:4px 10px;border-bottom:1px solid var(--hair)}
thead th:nth-child(1){text-align:left}
tbody td{padding:3px 10px;border-bottom:1px solid var(--hair);font-size:12px}
tbody tr:last-child td{border-bottom:none}
.rankgrid tbody tr.clk{cursor:pointer}
.rankgrid tbody tr.clk:hover{background:#F3F1E9}
.pn{color:var(--ink);font-weight:500}
td.n{text-align:right;font-variant-numeric:tabular-nums;color:#46453F;white-space:nowrap}
td.n.over{color:var(--terra);font-weight:600}
.ovl{position:fixed;inset:0;background:rgba(38,77,61,.45);display:none;align-items:center;justify-content:center;z-index:100;padding:24px}
.ovl.show{display:flex}
.modal{background:var(--paper);border-radius:16px;max-width:760px;width:100%;max-height:84vh;display:flex;flex-direction:column;box-shadow:0 30px 80px rgba(0,0,0,.32);overflow:hidden}
.modal .mhd{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:20px 26px 8px}
.modal .mt{font-family:"Cormorant Garamond",serif;font-size:26px;color:var(--forest);font-weight:600;line-height:1.05}
.modal .mcls{cursor:pointer;font-size:26px;color:var(--muted);line-height:1;flex:0 0 auto}
.modal .mcls:hover{color:var(--forest)}
.modal .msub{padding:0 26px 14px;color:var(--muted);font-size:12px;letter-spacing:.02em;border-bottom:1px solid var(--hair)}
.mwrap{overflow:auto;padding:2px 14px 12px}
.modal table{width:100%;border-collapse:collapse}
.modal thead th{position:sticky;top:0;background:var(--paper);font-size:10px;letter-spacing:.13em;text-transform:uppercase;color:var(--muted);font-weight:600;text-align:right;padding:10px 14px;border-bottom:1px solid var(--hair)}
.modal thead th:nth-child(1),.modal thead th:nth-child(2){text-align:left}
.modal tbody td{padding:7px 14px;border-bottom:1px solid var(--hair);font-size:12px}
.modal tbody tr:last-child td{border-bottom:none}
.mrk{color:var(--muted);width:34px;font-family:"Cormorant Garamond",serif;font-size:13px}
.mpn{color:var(--ink);max-width:340px}
.mn{text-align:right;font-variant-numeric:tabular-nums;color:#46453F;white-space:nowrap}
.mover{color:var(--terra);font-weight:600}
.methrow{display:grid;grid-template-columns:1fr 1.4fr 1fr;gap:16px;align-items:stretch}
.ncard.meth{display:flex;gap:13px;align-items:flex-start}
.mico{flex:0 0 26px;width:26px;height:26px;border-radius:50%;border:1px solid var(--brass);color:var(--brass);display:flex;align-items:center;justify-content:center;font-style:italic;font-weight:700;font-family:"Cormorant Garamond",serif;font-size:15px}
.meth .mh,.mdef .mh{margin:0 0 6px;font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--forest);font-weight:700}
.meth .mcode,.mdef .mcode{font-family:"Cormorant Garamond",Georgia,serif;font-size:13.5px;color:#46453F;line-height:1.55;display:block}
.ncard.mdef{display:flex;flex-direction:column;justify-content:center}
</style></head>
'''

TAIL_SCRIPT = '''
<script>
(function(){
  var fit = document.getElementById('fit');
  function apply(){
    fit.style.transform = 'none';
    var w = fit.offsetWidth, h = fit.offsetHeight;
    if(!w || !h) return;
    var pad = 16;
    var s = Math.min((window.innerWidth - pad) / w, (window.innerHeight - pad) / h);
    if(s > 1) s = 1;
    fit.style.transform = 'scale(' + s + ')';
  }
  window.addEventListener('resize', apply);
  window.addEventListener('load', apply);
  if(document.fonts && document.fonts.ready){ document.fonts.ready.then(apply); }
  apply();
})();
</script>
'''

HANDLERS = '''(function(){
  var ovl=document.getElementById('ovl'),mttl=document.getElementById('mttl'),
      msub=document.getElementById('msub'),mrows=document.getElementById('mrows');
  function fmt(v){return '$'+Math.round(v).toLocaleString();}
  function esc(s){return String(s).replace(/[&<>]/g,function(c){return c==='&'?'&amp;':c==='<'?'&lt;':'&gt;';});}
  function openPmo(name){
    var d=DETAIL[name];
    if(!d){return;}
    var tot=0,i;for(i=0;i<d.length;i++){tot+=d[i][3];}
    mttl.textContent=name;
    msub.textContent=d.length+' overrun project'+(d.length===1?'':'s')+' \\u00b7 '+fmt(tot)+' total overrun above original budget';
    var h='';
    for(i=0;i<d.length;i++){var r=d[i];
      h+='<tr><td class="mrk">'+(i+1)+'</td><td class="mpn">'+esc(r[0])+'</td><td class="mn">'+fmt(r[1])+'</td><td class="mn">'+fmt(r[2])+'</td><td class="mn mover">('+fmt(r[3])+')</td></tr>';}
    mrows.innerHTML=h;
    ovl.classList.add('show');
  }
  function closeM(){ovl.classList.remove('show');}
  var trs=document.querySelectorAll('.rankgrid tbody tr');
  for(var i=0;i<trs.length;i++){(function(tr){
    tr.classList.add('clk');
    tr.addEventListener('click',function(){
      var pn=tr.querySelector('.pn');
      if(pn){openPmo(pn.textContent.trim());}
    });
  })(trs[i]);}
  document.getElementById('mcls').addEventListener('click',closeM);
  ovl.addEventListener('click',function(e){if(e.target===ovl){closeM();}});
  document.addEventListener('keydown',function(e){if(e.key==='Escape'){closeM();}});
})();'''


def build(cur, prev, out):
    p = (lambda k: prev[k] if prev else None)
    N = len(cur["pmo"]); third = -(-N // 3)
    t1, t2, t3 = cur["pmo"][:third], cur["pmo"][third:2 * third], cur["pmo"][2 * third:]
    late = cur["late"]

    gtxt, gcls = delta_phrase(round(cur["orig_weighted"], 2), round(p("orig_weighted"), 2) if prev else None, lambda x: "%.2f%%" % x)
    if gtxt:
        klass = "down" if gcls in ("dn", "flat") else "up"
        gsub = 'Target &le; 5%% &middot; <span class="%s">%s</span> last week' % (klass, gtxt.replace(" last week", ""))
    else:
        gsub = "Target &le; 5%"
    o_txt, o_cls = delta_phrase(round(cur["orig_weighted"], 2), round(p("orig_weighted"), 2) if prev else None, lambda x: "%.2f%%" % x)
    r_txt, r_cls = delta_phrase(round(cur["rev_weighted"], 2), round(p("rev_weighted"), 2) if prev else None, lambda x: "%.2f%%" % x)
    o_delta = '<div class="delta %s">%s</div>' % (o_cls, o_txt) if o_txt else ""
    r_delta = '<div class="delta %s">%s</div>' % (r_cls, r_txt) if r_txt else ""
    ov_m = millions(cur["ov_amt"])

    body = '''<body>
<div id="fitviewport"><div id="fit">
<div class="topbar">
  <a class="back" href="index.html">&larr; Back to hub</a>
  <a class="xbtn" id="xbtn" href="{xlsx_href}" download="{xlsx_name}"
     title="Download the full Budget Watchtower workbook — All Active Projects, &gt;70% Exec View, &gt;70%, 90% &amp; &gt;$50k, 100% Rev Rec FF and Weighted Overrun Calculation">
    <svg viewBox="0 0 24 24"><path d="M12 3v12"/><path d="M7.5 10.5L12 15l4.5-4.5"/><path d="M4 20h16"/></svg>
    Export to Excel
  </a>
</div>
<div class="wrap">

  <section class="zone">
    <div class="topgrid">

      <div class="titleblock">
        <div class="eyebrow">Northstar Heritage &middot; PMO Performance</div>
        <h1 class="serif">Budget Watchtower</h1>
      </div>

      <div class="kpi">
        <svg class="ico" viewBox="0 0 24 24"><rect x="3.5" y="3.5" width="7" height="7" rx="1.4"/><rect x="13.5" y="3.5" width="7" height="7" rx="1.4"/><rect x="3.5" y="13.5" width="7" height="7" rx="1.4"/><rect x="13.5" y="13.5" width="7" height="7" rx="1.4"/></svg>
        <div class="v num">{orig_n}</div>
        <div class="l">Active Projects</div>
        <div class="s">Projects with an original budget</div>
      </div>

      <div class="kpi ring">
        <div class="l">Weighted Overrun</div>
        {gauge}
      </div>

      <div class="kpi">
        <svg class="ico" viewBox="0 0 24 24"><path d="M12 1v22M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>
        <div class="v num">{sumK}</div>
        <div class="l">Total Original Budget</div>
        <div class="s">Across all PMOs</div>
      </div>

      <div class="kpi">
        <svg class="ico" viewBox="0 0 24 24" style="stroke:var(--terra)"><path d="M23 18l-9.5-9.5-5 5L1 6"/><path d="M17 18h6v-6"/></svg>
        <div class="v num neg">({ov_m})</div>
        <div class="l">Total Overrun Amount</div>
        <div class="s">Costs above original budget</div>
      </div>

    </div>
  </section>

  <section class="zone">
    <div class="ztag"><span class="no serif">01</span><h2>Executive Summary</h2><span class="rule"></span></div>
    <div class="narr">

      <div class="ncard">
        <div class="hd"><span class="dotlbl tip">Original Cost Budget<span class="tipbox"><b>Original Budget &mdash; Weighted Overrun %</b>( Total Costs &minus; Original Budget ) &divide; Sum of all Original Budgets &times; 100</span></span></div>
        <div class="lead">
          <div><div class="cap">Weighted Overrun</div><div class="big serif">{orig_w}%</div></div>
          {o_delta}
        </div>
        <p class="say">Across <b>{orig_n}</b> active projects carrying an original budget, <b>{orig_op}%</b> are in overrun. Weighted overrun on original budgets sits at <b>{orig_w}%</b> against the 5% target.</p>
      </div>

      <div class="ncard">
        <div class="hd"><span class="dotlbl">Late-Stage Compliance</span></div>
        <div class="bar">
          <div class="br"><span class="pct num">{l0}%</span><span class="bl">100% rev. rec. projects with ETCs &middot; target 100%</span></div>
          <div class="track"><div class="fill" style="width:{l0}%"></div></div>
        </div>
        <div class="bar">
          <div class="br"><span class="pct num">{l1}%</span><span class="bl">90% rev. rec. projects with ETCs &middot; target 100%</span></div>
          <div class="track"><div class="fill" style="width:{l1}%"></div></div>
        </div>
        <div class="bar">
          <div class="br"><span class="pct num">{l2}%</span><span class="bl">70% rev. rec. projects with ETCs &middot; target 100%</span></div>
          <div class="track"><div class="fill" style="width:{l2}%"></div></div>
        </div>
      </div>

      <div class="ncard">
        <div class="hd"><span class="dotlbl tip">Revised (ETC) Budget<span class="tipbox"><b>Revised ETC &mdash; Weighted Overrun %</b>( Total Costs &minus; Total Revised Budget ) &divide; Sum of all Revised Budgets &times; 100<em>Total Revised Budget = Original Budget + Budget Adjustments / ETC.</em></span></span></div>
        <div class="lead">
          <div><div class="cap">Weighted Overrun</div><div class="big serif">{rev_w}%</div></div>
          {r_delta}
        </div>
        <ul class="rsum">
          <li><b>{gt70}</b> projects with over 70% of budget spent (additional filters summarized in the report){g70_up}.</li>
          <li><b>{gt90}</b> projects with over 90% of budget spent and contract values greater than $50K{g90_up}.</li>
          <li><b>{ff100}</b> active fixed-fee projects with 100% revenue recognition{ff_up}.</li>
          <li>Of the late-stage projects:
            <ul>
              <li><b>{l0}%</b> of projects with 100%+ revenue recognition have ETCs or are on-budget (target 100%).</li>
              <li><b>{l1}%</b> of projects with 90% revenue recognition have ETCs or are on-budget (target 100%).</li>
              <li><b>{l2}%</b> of projects with 70% revenue recognition have ETCs or are on-budget (target 100%).</li>
            </ul>
          </li>
        </ul>
      </div>

    </div>
  </section>

  <section class="zone">
    <div class="ztag"><span class="no serif">02</span><h2>PMO Performance Ranking</h2><span class="rule"></span><span class="rankhint">click a PMO for its project buildup</span></div>
    <div class="rankgrid">

      <div class="rank">
        <table>
          <thead><tr><th>PMO</th><th>Budget</th><th>Overrun</th><th>Overrun&nbsp;%</th></tr></thead>
          <tbody>
{t1}
          </tbody>
        </table>
      </div>

      <div class="rank">
        <table>
          <thead><tr><th>PMO</th><th>Budget</th><th>Overrun</th><th>Overrun&nbsp;%</th></tr></thead>
          <tbody>
{t2}
          </tbody>
        </table>
      </div>

      <div class="rank">
        <table>
          <thead><tr><th>PMO</th><th>Budget</th><th>Overrun</th><th>Overrun&nbsp;%</th></tr></thead>
          <tbody>
{t3}
          </tbody>
        </table>
      </div>

    </div>
  </section>

  <section class="zone" style="margin-bottom:0">
    <div class="ztag"><span class="no serif">03</span><h2>Methodology</h2><span class="rule"></span></div>
    <div class="methrow">

      <div class="ncard meth">
        <span class="mico">i</span>
        <div>
          <div class="mh">Original Budget &mdash; Weighted Overrun %</div>
          <code class="mcode">Weighted Overrun % = ( (Total Costs &minus; Original Budget) &divide; Sum of All Original Budgets ) &times; 100</code>
        </div>
      </div>

      <div class="ncard meth">
        <span class="mico">i</span>
        <div>
          <div class="mh">Revised ETC &mdash; Weighted Overrun %</div>
          <code class="mcode">Weighted Overrun % = ( (Total Costs &minus; (Original Budget + Budget Adjustments/ETC, i.e. Total Revised Budgets)) &divide; Sum of All Revised Budgets ) &times; 100</code>
        </div>
      </div>

      <div class="ncard mdef">
        <div class="mh">Total Revised Budget</div>
        <code class="mcode">Original Budget + Budget Adjustments / ETC</code>
      </div>

    </div>
  </section>

</div>
</div></div>'''.format(
        orig_n="{:,}".format(cur["orig_n"]), gauge=gauge(cur["orig_weighted"]), gsub=gsub,
        sumK=millions(cur["sumK"]), ov_m=ov_m, orig_w="%.2f" % cur["orig_weighted"],
        orig_op=cur["orig_overrun_pct"], o_delta=o_delta, rev_w="%.2f" % cur["rev_weighted"],
        r_delta=r_delta, gt70=cur["gt70"], gt90=cur["gt90"], ff100=cur["ff100"],
        l0=late[0], l1=late[1], l2=late[2],
        g70_up=up(cur["gt70"], p("gt70")), g90_up=up(cur["gt90"], p("gt90")), ff_up=up(cur["ff100"], p("ff100")),
        t1=pmo_rows(t1), t2=pmo_rows(t2), t3=pmo_rows(t3),
        xlsx_href=html.escape(XLSX_FILE), xlsx_name=html.escape(download_name(cur["asof"])))

    modal = ('\n<div id="ovl" class="ovl">\n  <div class="modal">\n'
             '    <div class="mhd"><span id="mttl" class="mt"></span><span id="mcls" class="mcls">&times;</span></div>\n'
             '    <div id="msub" class="msub"></div>\n'
             '    <div class="mwrap">\n      <table>\n'
             '        <thead><tr><th>#</th><th>Project</th><th>Original Budget</th><th>Total Cost</th><th>Overrun</th></tr></thead>\n'
             '        <tbody id="mrows"></tbody>\n      </table>\n    </div>\n  </div>\n</div>\n')
    modal += ('<div id="xtoast" class="xtoast"></div>\n')
    detail_js = json.dumps(cur["detail"], separators=(",", ":"))
    handlers = ('<script>\nvar DETAIL=' + detail_js + ';\n' + HANDLERS + '\n'
                + XPORT_JS + '\n</script>\n')
    open(out, "w", encoding="utf-8").write(HEAD + body + modal + TAIL_SCRIPT + handlers + "</body></html>\n")
    return out


def main():
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    cur = compute(sys.argv[1])
    prev = compute(sys.argv[2]) if len(sys.argv) > 2 else None
    out = build(cur, prev, os.path.join(HERE, "budget-overrun-summary.html"))
    xlsx = os.path.join(HERE, XLSX_FILE)
    eng.build(sys.argv[1], xlsx)
    print("Workbook -> %s  (offered as \"%s\")" % (xlsx, download_name(cur["asof"])))
    print("as-of %s: orig_n=%d weighted=%.2f%% budget=%s overrun=%s PMOs=%d"
          % (cur["asof"], cur["orig_n"], cur["orig_weighted"], millions(cur["sumK"]),
             millions(cur["ov_amt"]), len(cur["pmo"])))
    print("Saved -> %s" % out)


if __name__ == "__main__":
    main()
