#!/usr/bin/env python3
"""Fail loudly when the Excel export's side-car data files have fallen behind index.html.

FP&A (2026-09-29): the 09/22 refresh rebuilt the site but skipped `regenerate_raw_data.py`
and the copy of `proposal_pivot.js`, so the export's Raw Data and Proposal Hours sheets sat
on week-ending 9/12 while every DOM-built sheet moved to 9/19.  Nothing surfaced it — the
export has no reason to complain — and it went unnoticed for a week.

Run this at the END of every weekly refresh, after the files are in place:

    python3 check_export_freshness.py                 # checks the folder it lives in
    python3 check_export_freshness.py --dir "/path"

Exit status 0 = everything ties, 1 = something is stale (message says what to re-run),
2 = could not determine (missing file / unreadable), which is also worth a look.

The page carries the same check at export time (`checkExportFreshness` at the end of
index.html) so a stale workbook is visible to whoever runs the export, not just here.
"""
import argparse, json, os, re, sys

def _payload(path, marker='{"generated"'):
    with open(path, encoding='utf-8') as f:
        t = f.read()
    i = t.index(marker)
    return json.loads(t[i:].rstrip().rstrip(';'))

def site_latest_week(index_html):
    """The site's own latest completed week: pmoReducedData.as_of, else ed2Data."""
    with open(index_html, encoding='utf-8', errors='replace') as f:
        s = f.read()
    m = re.search(r'<script id="pmoReducedData" type="application/json">(.*?)</script>', s, re.S)
    if m:
        try:
            v = json.loads(m.group(1)).get('as_of')
            if v:
                return str(v)[:10], 'pmoReducedData.as_of'
        except Exception:
            pass
    m = re.search(r'<script id="ed2Data" type="application/json">(.*?)</script>', s, re.S)
    if m:
        try:
            d = json.loads(m.group(1))
            i = d.get('last_week_idx')
            if i is not None and d.get('date_iso') and d['date_iso'][i]:
                return str(d['date_iso'][i])[:10], 'ed2Data.date_iso[last_week_idx]'
        except Exception:
            pass
    return None, None

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--dir', default=os.path.dirname(os.path.abspath(__file__)))
    a = ap.parse_args()
    d = a.dir
    idx = os.path.join(d, 'index.html')
    if not os.path.exists(idx):
        print(f'CANNOT CHECK: {idx} not found'); return 2
    site, src = site_latest_week(idx)
    if not site:
        print('CANNOT CHECK: no latest week found in index.html'); return 2
    print(f'index.html latest completed week: {site}  (from {src})')

    stale, unknown = [], []

    p = os.path.join(d, 'raw_data.js')
    try:
        raw = _payload(p)
        wi = raw['columns'].index('Week End')
        mx = max((r[wi] for r in raw['rows'] if r[wi]), default='')
        print(f'raw_data.js          latest Week End: {mx}   ({len(raw["rows"]):,} rows, generated {raw.get("generated")})')
        if mx < site:
            stale.append(('raw_data.js', mx, 'regenerate_raw_data.py'))
    except Exception as e:
        unknown.append(f'raw_data.js: {e}')

    p = os.path.join(d, 'proposal_pivot.js')
    try:
        pp = _payload(p)
        mx = str(pp['weeks'][-1])[:10]
        print(f'proposal_pivot.js    latest Week End: {mx}   (generated {pp.get("generated")})')
        if mx < site:
            stale.append(('proposal_pivot.js', mx, 'regenerate_proposal_pivot.py'))
    except Exception as e:
        unknown.append(f'proposal_pivot.js: {e}')

    if unknown:
        for u in unknown:
            print('CANNOT CHECK:', u)
    if stale:
        print()
        for name, mx, script in stale:
            print(f'STALE: {name} stops at {mx}, site is through {site} -> re-run {script}')
        return 1
    if unknown:
        return 2
    print('\nOK: every export side-car file is current.')
    return 0

if __name__ == '__main__':
    sys.exit(main())
