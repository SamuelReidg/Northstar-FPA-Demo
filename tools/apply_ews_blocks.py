#!/usr/bin/env python3
"""apply_ews_blocks.py -- splice the blocks written by pipeline/regenerate_ews_data.py into the pages.

This is the scripted form of the monthly hand edit described in the EWS workflow notes:

  early-warning-system.html
    * `const DATA = {...` : everything before `, "map_path"` is replaced by ews_data_prefix.txt
      (map_path / viewBox are never touched) and the `"offsite": {...}` value by ews_offsite.txt;
    * the hard-coded "PMOs with Critical Risk" tile  NN / NN  (offices with a Critical project /
      offices on the map -- the crit_pmos / total_pmos the regen prints);
    * `const BYOFFICE = {...};` and the injected `Object.assign(BYOFFICE, {...});` drill-down block;
  project-recovery.html
    * the injected `window.PROJECTS = [...]` assignment after the schema comment (ews_projects.txt)
      and the "as of" in that comment;
  ews-header-sidebar.html (unlinked header mock-up)
    * its "Data as of:" pill.

usage: python3 tools/apply_ews_blocks.py <dir with ews_*.txt> <site dir>
"""
import json, os, re, sys


def main():
    src, site = sys.argv[1], sys.argv[2]
    rd = lambda n: open(os.path.join(src, n), encoding='utf-8').read()
    prefix, byoffice, offsite, projects = rd('ews_data_prefix.txt'), rd('ews_byoffice.txt'), \
        rd('ews_offsite.txt'), rd('ews_projects.txt')
    as_of = json.loads(re.search(r'"as_of": ("[^"]*")', prefix).group(1))
    geo = json.loads(prefix[prefix.index('"geo": ') + 7:])
    crit_pmos, total_pmos = sum(1 for g in geo if g['Critical'] > 0), len(geo)
    by_obj = byoffice[len('const BYOFFICE = '):].rstrip().rstrip(';')

    # ---- early-warning-system.html
    p = os.path.join(site, 'early-warning-system.html')
    s = open(p, encoding='utf-8').read()
    a = s.index('const DATA = {'); b = s.index(', "map_path"', a)
    s = s[:a] + prefix + s[b:]
    s, n1 = re.subn(r'"offsite": \{[^}]*\}', '"offsite": ' + offsite.strip(), s, count=1)
    s, n2 = re.subn(r'(<div class="ml">PMOs with Critical Risk</div><div class="mv">)\d+(<span[^>]*> / )\d+(</span>)',
                    lambda m: m.group(1) + str(crit_pmos) + m.group(2) + str(total_pmos) + m.group(3), s, count=1)
    s, n3 = re.subn(r'^const BYOFFICE = .*$', lambda m: byoffice.strip(), s, count=1, flags=re.M)
    s, n4 = re.subn(r'^Object\.assign\(BYOFFICE, .*\);$', lambda m: 'Object.assign(BYOFFICE, ' + by_obj + ');', s,
                    count=1, flags=re.M)
    s, n5 = re.subn(r'(// === Filtered to IN PROGRESS projects only \()[^)]*(\))', r'\g<1>as of ' + as_of + r'\2', s, count=1)
    assert (n1, n2, n3, n4, n5) == (1, 1, 1, 1, 1), (n1, n2, n3, n4, n5)
    open(p, 'w', encoding='utf-8').write(s)

    # ---- project-recovery.html
    p = os.path.join(site, 'project-recovery.html')
    s = open(p, encoding='utf-8').read()
    marker = 'this assignment overrides the array above ===\n'
    a = s.index(marker) + len(marker)
    b = s.index('\n', a)
    assert s[a:b].startswith('window.PROJECTS = [')
    s = s[:a] + projects.strip() + s[b:]
    s, n6 = re.subn(r'(Critical\+Warning\+Watch\), as of ).*?( — this assignment)', r'\g<1>' + as_of + r'\2', s, count=1)
    assert n6 == 1
    open(p, 'w', encoding='utf-8').write(s)

    # ---- ews-header-sidebar.html
    p = os.path.join(site, 'ews-header-sidebar.html')
    s = open(p, encoding='utf-8').read()
    s, n7 = re.subn(r'Data as of: [A-Z][a-z]+ \d{1,2}, \d{4}', 'Data as of: ' + as_of, s, count=1)
    assert n7 == 1
    open(p, 'w', encoding='utf-8').write(s)
    print('applied: as of %s | PMOs with Critical Risk %d / %d' % (as_of, crit_pmos, total_pmos))


if __name__ == '__main__':
    main()
