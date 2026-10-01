"""Northstar Heritage — the fictional company behind the demo.

Everything in the demo data package is generated here from a fixed seed: offices, people,
projects, clients, and a full year of daily timesheet entries. Every source export the
pipeline reads (the utilization workbook, the budget report, the approvals exports, the
project-risk workbook, PM hours, proposals, revenue forecast, location drivers) is written
from this one world by the `tools/gen_*.py` writers, so the numbers agree across dashboards
the way they would in a real company.

Nothing here is derived from any real organisation's data. Names come from a fixed fictional pool (tools/name_pool.py); project
names, clients and places are assembled from generic word lists below.

    from demo_world import world
    W = world()            # cached after the first build (tools/.cache/world.pkl)
"""
from __future__ import annotations

import datetime as dt
import math
import os
import pickle
import random
from collections import defaultdict

SEED = 20260926
AS_OF = dt.date(2026, 9, 26)            # last completed week-ending Saturday in the demo
FY = 2026
FIRST_WEEK_END = dt.date(2026, 1, 3)
HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, '.cache', 'world.pkl')
COMPANY = 'Northstar Heritage'
COMPANY_SHORT = 'Northstar'
EMAIL_DOMAIN = 'northstar-heritage.example'

# ── Offices / PMOs ───────────────────────────────────────────────────────────
# kind: regional | service | admin | acquired | international
# (name, kind, state, FT headcount, PT headcount, FT target, dept code, work location)
OFFICES = [
    ('Big Sky',            'regional',      'MT', 10, 3, 0.80, 210, 'Big Sky'),
    ('Bluegrass',          'regional',      'KY',  8, 2, 0.78, 220, 'Bluegrass'),
    ('Cascade',            'regional',      'OR', 20, 4, 0.84, 230, 'Cascade'),
    ('Chesapeake',         'regional',      'MD', 14, 4, 0.81, 240, 'Chesapeake'),
    ('Front Range',        'regional',      'CO', 16, 3, 0.72, 250, 'Front Range'),
    ('Great Lakes',        'regional',      'WI', 18, 6, 0.83, 260, 'Great Lakes'),
    ('Gulf Coast',         'regional',      'TX', 12, 3, 0.74, 270, 'Gulf Coast'),
    ('High Desert',        'regional',      'NM', 13, 5, 0.82, 280, 'High Desert'),
    ('Island',             'regional',      'PR',  7, 2, 0.75, 290, 'Island'),
    ('Lakeshore',          'regional',      'NY',  3, 2, 0.75, 300, 'Lakeshore'),
    ('Hill Country',       'regional',      'AR',  9, 4, 0.76, 310, 'Hill Country'),
    ('Piedmont',           'regional',      'NC',  8, 3, 0.74, 320, 'Piedmont'),
    ('Prairie',            'regional',      'KS',  7, 2, 0.77, 330, 'Prairie'),
    ('Red Rock',           'regional',      'AZ', 24, 9, 0.82, 340, 'Red Rock'),
    ('Sierra',             'regional',      'CA', 13, 3, 0.76, 350, 'Sierra'),
    ('Coastal Bend',       'regional',      'FL', 17, 8, 0.80, 360, 'Coastal Bend'),
    ('Built Environment',  'service',       'CO', 12, 1, 0.72, 400, 'Front Range'),
    ('Paleontology',       'service',       'AZ',  7, 3, 0.78, 410, 'Red Rock'),
    ('Maritime',           'service',       'FL',  5, 2, 0.70, 420, 'Coastal Bend'),
    ('Geospatial',         'service',       'OR',  4, 0, 0.60, 430, 'Cascade'),
    ('Meridian Central',   'acquired',      'MO',  8, 1, 0.86, 500, 'Meridian Central'),
    ('Meridian Southwest', 'acquired',      'NV',  5, 1, 0.72, 510, 'Meridian Southwest'),
    ('International',      'international', 'INTL', 12, 0, 0.84, 600, 'International'),
    ('Corporate',          'admin',         'CO', 18, 1, 0.00, 100, 'Northstar Heritage Corporate'),
    ('Sales & Marketing',  'admin',         'CO',  6, 0, 0.00, 110, 'Northstar Heritage Corporate'),
    ('Archive Services',   'admin',         'CO',  0, 0, 0.00, 120, 'Northstar Heritage Corporate'),
]
OFFICE_NAMES = [o[0] for o in OFFICES]
# Location names as timesheets carry them (the ERP's "Location name"), per PMO.
LOCATION_OF = {o[0]: o[0] for o in OFFICES}
LOCATION_OF.update({'Sales & Marketing': 'Sales and Marketing',
                    'International': 'NH Gulf (USD)'})
# Location IDs as the ERP's employee master stores them (Approvals dashboard).
LOCATION_ID = {o[0]: 'L%d' % (8000 + i * 10) for i, o in enumerate(OFFICES)}
# Two International staff are billed through a separate entity and shown on their own row.
INTL_ENTITY_LABEL = 'International (Global LLC)'
# A small office the business asked to keep off the monthly PMO table (demo of that rule).
HIDDEN_ON_MONTHLY = {'lakeshore'}

SECTORS = ['Energy', 'Transportation', 'State/Local Government', 'Federal Government', 'Water',
           'Telecommunications', 'Commercial Development', 'Mining', 'Residential Development',
           'Communication', 'Education', 'Tribal', 'Parks & Recreation', 'Utilities']
BILLING = ['Fixed fee', 'Time and material']
PROJECT_STATUSES = ['In Progress', 'Final Reporting/Review', 'Lab/Curation', 'Cost Tracking on Closed Project',
                    'Final Billing/Invoice', 'On Hold', 'Completed', 'Transferred Entity']

# Generic word lists for project and client names (no real places or organisations).
PLACES = ['Aspen Ridge', 'Bramble Creek', 'Cedar Hollow', 'Copper Flats', 'Dove Canyon', 'Eagle Bend',
          'Elk Meadow', 'Falcon Mesa', 'Fern Valley', 'Foxglove', 'Granite Pass', 'Hawk Point',
          'Heron Marsh', 'Juniper Wash', 'Kestrel Butte', 'Larkspur', 'Lone Pine', 'Maple Fork',
          'Creosote Draw', 'Millbrook', 'Oak Terrace', 'Plover Cove', 'Pine Hollow', 'Quail Run',
          'Raven Rock', 'Redbud', 'Sage Flat', 'Sandpiper', 'Silver Lake', 'Sparrow Hill',
          'Stone Mill', 'Sumac Ridge', 'Hackberry Bend', 'Thistle Down', 'Timber Gap', 'Tansy Flat',
          'Willow Springs', 'Wren Hollow', 'Yarrow Point', 'Blue Heron', 'Coyote Springs',
          'Sea Oat', 'Ember Valley', 'Goldfinch', 'Harbor View', 'Slate Ridge', 'Lantern Bay',
          'Meadowlark', 'North Fork', 'Pebble Beach Road', 'Riverbend', 'Saltgrass', 'Tamarack',
          'Twin Buttes', 'Wildrye', 'Windmill Flat', 'Rushgrass Draw', 'Badger Hollow']
PROJECT_KINDS = ['Solar', 'Wind Farm', 'Transmission Line', 'Pipeline', 'Highway Widening',
                 'Bridge Replacement', 'Water Main', 'Reservoir', 'Fiber Route', 'Cell Tower',
                 'Housing Development', 'Mixed-Use Development', 'Trail', 'Park Improvements',
                 'School Campus', 'Levee Repair', 'Rail Siding', 'Substation', 'Wastewater Plant',
                 'Mine Expansion', 'Airport Apron', 'Harbor Dredging', 'Canal Lining', 'Battery Storage']
PROJECT_SERVICES = ['Class III Survey', 'Cultural Resources Monitoring', 'Phase I Survey',
                    'Phase II Testing', 'Data Recovery', 'Historic Resources Evaluation',
                    'Paleontological Monitoring', 'Architectural History Survey', 'Section 106 Consultation',
                    'Cultural Resources Assessment', 'Archaeological Testing', 'Built Environment Survey',
                    'Maritime Remote Sensing Survey', 'GIS Mapping', 'Curation', 'Environmental Review Support']
CLIENT_A = ['Summit', 'Blue Mesa', 'Ridgeline', 'Prairie Wind', 'Keystone', 'Evergreen', 'Silverline',
            'Northfield', 'Clearwater', 'Ironbridge', 'Sunrise', 'Pioneer', 'Harborlight', 'Granite State',
            'Redtail', 'Coastal', 'Mesa Verde', 'High Plains', 'Lakecrest', 'Riverstone', 'Stonebridge',
            'Timberline', 'Westwind', 'Highland', 'Brightwater', 'Canyon', 'Cardinal', 'Heartland']
CLIENT_B = ['Energy Partners', 'Transit Authority', 'Power Cooperative', 'Water District', 'Development Group',
            'Engineering, Inc.', 'Renewables LLC', 'Pipeline Company', 'Telecom Holdings', 'Builders',
            'Land Company', 'Utilities', 'Infrastructure Group', 'Resources Corp.', 'Planning Associates',
            'Environmental Consultants', 'Construction Co.', 'Port Authority', 'County', 'Rail Partners']
TASKS_FIELD = ['Prefield', 'Fieldwork', 'Reporting', 'Project Management', 'Background Research',
               'Monitoring', 'Lab Analysis', 'Site Recordation', 'GIS & Mapping', 'Draft Report',
               'Final Report', 'Agency Consultation', 'Curation', 'Data Recovery Excavation',
               'Testing', 'Survey', 'Desktop Review']

# Overhead tasks (Project ID "OVHD"). (task, time type, weight in the non-utilized mix)
OVHD_TASKS = [
    ('OVHD - Labor - Administrative',               'IND', 34),
    ('OVHD - Labor - Meetings',                     'IND', 16),
    ('OVHD - Labor - Proposal Writing',             'IND', 12),
    ('OVHD - Production - Proposal Editing',        'IND', 2),
    ('OVHD - Labor - Business Development',         'IND', 5),
    ('OVHD - Labor - Marketing',                    'IND', 3),
    ('OVHD - Labor - Preparation',                  'IND', 4),
    ('OVHD - Labor - Company Meeting',              'IND', 1),
    ('OVHD - Labor - Employee Relations & Reviews', 'IND', 1),
    ('OVHD - Labor - Recruiting',                   'IND', 1),
    ('OVHD - Labor - Expense Reports',              'IND', 1),
    ('OVHD - Labor - Mentorship & Training',        'IND', 1),
    ('OVHD - Labor - Local Fleet Management',       'IND', 1),
    ('OVHD - Labor - National Fleet Management',    'IND', 1),
    ('OVHD - Labor - Service Line Program Mgmt',    'IND', 2),
    ('OVHD - Labor - Tech - Support And Systems Administration', 'IND', 1),
    ('OVHD - Labor - Onboarding Training',          'IND', 1),
    ('OVHD - Labor - Non Billable Drive Time',      'IND', 1),
    ('OVHD - Labor - Facilities Management',        'IND', 1),
]
TASK_VACATION = ('OVHD - Labor - Vacation (Salaried)', 'Vac')
TASK_PTO = ('OVHD - Labor - PTO (Hourly Employee)', 'PTO')
TASK_HOLIDAY = ('OVHD - Labor - Holiday', 'HOL')
TASK_SICK = ('OVHD - Labor - Sick (Salaried/Temps)', 'Sick')
TASK_TRAINING = ('OVHD - Labor - Required Training', 'IND')
TASK_PARENTAL = ('OVHD - Labor - Parental Leave', 'MAT')
NON_POSTING = 'Overhead Non-Posting'
NON_POSTING_REGIONAL = 'CA - Overhead Non-Posting'   # regional variant (not stripped by the exact rule)
HOLIDAYS = [dt.date(2026, 1, 1), dt.date(2026, 1, 19), dt.date(2026, 5, 25), dt.date(2026, 7, 3),
            dt.date(2026, 9, 7), dt.date(2026, 11, 26), dt.date(2026, 11, 27), dt.date(2026, 12, 25)]

TITLES = {
    'field': ['Field Technician', 'Staff Archaeologist', 'Associate Archaeologist I', 'Associate Archaeologist II'],
    'senior': ['Senior Archaeologist', 'Senior Architectural Historian', 'Senior Paleontologist',
               'Project Manager', 'Team Lead', 'Regional Principal Investigator', 'GIS Specialist'],
    'lead': ['Office Principal', 'Regional Principal', 'Service Line Lead', 'Office Operations Lead'],
    'admin': ['Billing Specialist', 'Payroll Specialist', 'Human Resources Generalist', 'Proposal Coordinator',
              'Accountant', 'Payroll Supervisor', 'Director of Marketing', 'Account Executive',
              'Database Specialist', 'Systems Engineer', 'Senior Copy Editor', 'Admin Assistant',
              'Staff Accountant', 'IT Support Specialist', 'Marketing Coordinator'],
}
SERVICE_TITLES = {'Built Environment': 'Architectural Historian', 'Paleontology': 'Paleontologist',
                  'Maritime': 'Maritime Archaeologist', 'Geospatial': 'GIS Specialist'}


def week_ends(until=dt.date(2026, 12, 26)):
    d, out = FIRST_WEEK_END, []
    while d <= until:
        out.append(d); d += dt.timedelta(days=7)
    return out


def eomonth(d):
    nxt = (d.replace(day=28) + dt.timedelta(days=4))
    return nxt - dt.timedelta(days=nxt.day)


def q(x, step=0.25):
    return round(round(x / step) * step, 2)


class World:
    pass


from name_pool import FIRST_NAMES, LAST_NAMES


def build(seed=SEED):
    rnd = random.Random(seed)
    W = World()
    W.seed, W.as_of, W.company = seed, AS_OF, COMPANY
    W.weeks = week_ends()
    W.offices = [dict(name=n, kind=k, state=s, ft=ft, pt=pt, target=t, dept=dc, work_location=wl,
                      location=LOCATION_OF[n], location_id=LOCATION_ID[n])
                 for (n, k, s, ft, pt, t, dc, wl) in OFFICES]
    W.office = {o['name']: o for o in W.offices}

    # ── people ──────────────────────────────────────────────────────────────
    used = set()

    def person():
        while True:
            first, last = rnd.choice(FIRST_NAMES), rnd.choice(LAST_NAMES)
            mi = rnd.choice('ABCDEFGHJKLMNPRSTW')
            if (first, last) not in used:
                used.add((first, last)); return first, last, mi

    emps = []
    next_id = [70000]

    def new_emp(office, ft, status='Active', role='field', intl=False):
        first, last, mi = person()
        if intl:
            eid = ('S%d' if rnd.random() < 0.6 else 'E%d') % (4100 + len(emps))
        else:
            next_id[0] += rnd.randint(1, 9); eid = next_id[0]
        o = W.office[office]
        if role == 'admin':
            title = rnd.choice(TITLES['admin']); target = 0.0
        elif role == 'lead':
            title = rnd.choice(TITLES['lead']); target = round(rnd.choice([0.30, 0.40, 0.50, 0.55]), 2)
        elif role == 'senior':
            title = SERVICE_TITLES.get(office) and ('Senior ' + SERVICE_TITLES[office]) or rnd.choice(TITLES['senior'])
            target = round(min(0.95, max(0.55, rnd.gauss(o['target'], 0.06))), 2)
        else:
            base = SERVICE_TITLES.get(office)
            title = ('Staff ' + base) if base else rnd.choice(TITLES['field'])
            target = 0.95 if ft == 'Part Time' else round(min(0.95, max(0.6, rnd.gauss(o['target'] + 0.05, 0.05))), 2)
        if intl:
            title = rnd.choice(['International Field Director', 'Senior Archaeologist', 'Staff Archaeologist',
                                'International Report Coordinator'])
        hourly = 'Salary Exempt' if role in ('lead', 'admin') and rnd.random() < 0.85 or role == 'senior' and rnd.random() < 0.6 \
            else 'Hourly Non Exempt'
        hire = dt.date(rnd.randint(2008, 2026) if status == 'Active' else rnd.randint(2012, 2025),
                       rnd.randint(1, 12), rnd.randint(1, 28))
        if status == 'Active' and hire > AS_OF - dt.timedelta(days=21):
            hire = AS_OF - dt.timedelta(days=rnd.randint(30, 900))
        e = dict(id=eid, first=first, last=last, mi=mi, name='%s, %s' % (last, first),
                 roster_name='%s, %s %s.' % (last, first, mi), office=office, ft=ft,
                 flsa=hourly, title=title, role=role, target=target if office_target_ok(office) else 0.0,
                 status=status, hire=hire, term=None, supervisor=None, intl=intl,
                 email='%s%s@%s' % (first[0].lower(), last.lower().replace("'", ''), EMAIL_DOMAIN),
                 sage_id='E%05d' % (30900 + len(emps)), paygroup=rnd.choice(['Bi-Weekly_West', 'Bi-Weekly_East']))
        e['billable'] = 'Billable' if e['target'] > 0 else 'Non-Billable'
        # hourly cost (loaded) and bill rates by seniority; these drive project labor cost
        base = {'field': 27, 'senior': 46, 'lead': 72, 'admin': 38}[role] * (1.12 if intl else 1.0)
        e['cost_rate'] = round(base * rnd.uniform(0.88, 1.18) * 1.55, 2)          # burdened
        e['bill_rate'] = round(e['cost_rate'] * rnd.uniform(1.65, 2.05), 0)
        emps.append(e)
        return e

    def office_target_ok(office):
        return W.office[office]['kind'] != 'admin'

    for o in W.offices:
        n, kind = o['name'], o['kind']
        intl = kind == 'international'
        if kind == 'admin':
            for i in range(o['ft']):
                new_emp(n, 'Full Time', role='admin')
            for i in range(o['pt']):
                new_emp(n, 'Part Time', role='admin')
            continue
        if o['ft']:
            new_emp(n, 'Full Time', role='lead', intl=intl)
        for i in range(max(0, o['ft'] - 1)):
            new_emp(n, 'Full Time', role='senior' if rnd.random() < 0.35 else 'field', intl=intl)
        for i in range(o['pt']):
            new_emp(n, 'Part Time', role='field')

    # separated staff: some left during FY (they have partial-year hours), most earlier (roster only)
    active_count = len(emps)
    for i in range(int(active_count * 1.1)):
        o = rnd.choice([x for x in W.offices if x['kind'] not in ('admin',) and (x['ft'] + x['pt'])])
        e = new_emp(o['name'], rnd.choice(['Part Time', 'Part Time', 'Full Time']),
                    status=rnd.choice(['Terminated', 'Terminated', 'Resigned', 'Laid off']),
                    role='field', intl=o['kind'] == 'international')
        if i < active_count * 0.12:                               # left during FY 2026
            e['term'] = dt.date(2026, rnd.randint(1, 8), rnd.randint(1, 28))
            e['hire'] = min(e['hire'], dt.date(2025, rnd.randint(1, 12), rnd.randint(1, 28)))
        else:
            e['hire'] = dt.date(rnd.randint(2012, 2024), rnd.randint(1, 12), rnd.randint(1, 28))
            e['term'] = e['hire'] + dt.timedelta(days=rnd.randint(40, 1200))
            if e['term'] > dt.date(2025, 12, 15):
                e['term'] = dt.date(2025, rnd.randint(1, 11), rnd.randint(1, 28))

    # growth: ~9% of today's field staff joined during FY2026, weighted toward the spring
    # field-season ramp (March–June), so headcount climbs through the year
    field = [x for x in emps if x['status'] == 'Active' and x['role'] == 'field']
    for e in rnd.sample(field, int(len(field) * 0.09) + 4):
        m = rnd.choices(range(1, 10), weights=[3, 6, 12, 14, 12, 8, 5, 4, 3])[0]
        e['hire'] = dt.date(2026, m, rnd.randint(1, 28))

    # the two International staff billed through the separate entity
    intl_staff = [e for e in emps if e['office'] == 'International' and e['status'] == 'Active']
    W.intl_entity_ids = [str(e['id']) for e in intl_staff[1:3]]

    # supervisors: office lead supervises the office; leads report to a corporate VP
    corp = [e for e in emps if e['office'] == 'Corporate']
    for e, title in zip(corp, ['Chief Executive Officer', 'Vice President of Operations', 'Chief Financial Officer',
                               'Controller', 'Director of Finance', 'Director of Human Resources']):
        e['title'] = title
    ceo, vp = corp[0], corp[1]
    for o in W.offices:
        members = [e for e in emps if e['office'] == o['name']]
        lead = next((e for e in members if e['role'] == 'lead' and e['status'] == 'Active'), None)
        for e in members:
            e['supervisor'] = (lead or vp)['name'] if e is not lead else vp['name']
        if o['kind'] == 'admin':
            for e in members:
                e['supervisor'] = ceo['name'] if e is not ceo else ''
    W.employees = emps
    W.emp = {str(e['id']): e for e in emps}

    # ── clients & projects ─────────────────────────────────────────────────
    clients = set()
    while len(clients) < 160:
        a, b = rnd.choice(CLIENT_A), rnd.choice(CLIENT_B)
        clients.add(('%s County' % a) if b == 'County' else '%s %s' % (a, b))
    W.clients = sorted(clients)
    projects = []
    pid = [30000]
    field_offices = [o for o in W.offices if o['kind'] in ('regional', 'service', 'acquired', 'international')]
    for o in field_offices:
        n_proj = max(14, int((o['ft'] + o['pt']) * rnd.uniform(3.2, 4.4)))
        pms = [e for e in emps if e['office'] == o['name'] and e['status'] == 'Active' and e['role'] in ('lead', 'senior')] \
            or [e for e in emps if e['office'] == o['name'] and e['status'] == 'Active']
        for i in range(n_proj):
            pid[0] += rnd.randint(1, 23)
            yr = rnd.choice([23, 24, 25, 25, 26, 26, 26])
            place, kind, svc = rnd.choice(PLACES), rnd.choice(PROJECT_KINDS), rnd.choice(PROJECT_SERVICES)
            name = rnd.choice(['{p} {k} {s}', '{p} {k} - {s}', '{c}: {p} {k} {s}', '{p} {s}'])\
                .format(p=place, k=kind, s=svc, c=rnd.choice(['Task Order %d' % rnd.randint(1, 40), 'On-Call', 'Phase %s' % rnd.choice('I II III'.split())]))
            billing = rnd.choice(BILLING) if rnd.random() < 0.8 else 'Fixed fee'
            contract = round(math.exp(rnd.gauss(math.log(105000), 1.15)) / 100) * 100
            contract = max(3500, min(contract, 5_200_000))
            status = rnd.choices(PROJECT_STATUSES, weights=[52, 9, 4, 3, 16, 5, 9, 2])[0]
            if o['kind'] == 'international' and rnd.random() < 0.6:
                status = 'Transferred Entity'
            start = dt.date(2000 + yr, rnd.randint(1, 12), rnd.randint(1, 28))
            if start > AS_OF: start = AS_OF - dt.timedelta(days=rnd.randint(20, 200))
            dur = rnd.randint(60, 900)
            p = dict(id='%d-PC-%05d' % (yr, pid[0]), name=name, office=o['name'], location=o['location'],
                     pm=rnd.choice(pms)['name'], client=rnd.choice(W.clients), sector=rnd.choice(SECTORS),
                     billing=billing, contract=float(contract), status=status, start=start,
                     end=start + dt.timedelta(days=dur), svc=svc)
            ntask = rnd.randint(1, 5)
            p['tasks'] = rnd.sample(TASKS_FIELD, ntask)
            # economics: contract -> original budget at a target margin that depends on billing
            # type and sector (T&M is priced at rate so it carries a steadier margin; fixed fee
            # is bid tighter and varies more). Costs are NOT drawn here: they are summed from the
            # timesheet hours actually charged (x each person's cost rate) plus expenses, in
            # _finalize_projects(), so overruns come from real labor, not from a random number.
            tm = billing == 'Time and material'
            sector_adj = {'Federal Government': -0.03, 'Energy': 0.03, 'Telecommunications': 0.02,
                          'Tribal': -0.02, 'Mining': 0.02}.get(p['sector'], 0.0)
            p['target_margin'] = round(min(0.48, max(0.12, rnd.gauss(0.31 if tm else 0.27, 0.035 if tm else 0.07) + sector_adj)), 4)
            p['orig_budget'] = round(contract * (1 - p['target_margin']), 2)
            p['change_order'] = rnd.random() < (0.28 if dur > 365 else 0.15)
            p['revised_budget'] = round(p['orig_budget'] * (rnd.uniform(1.06, 1.38) if p['change_order'] else 1.0), 2)
            if p['change_order']:
                p['contract'] = round(p['contract'] * (p['revised_budget'] / p['orig_budget']) * rnd.uniform(0.98, 1.03), -2)
            p['expense_share'] = round(rnd.uniform(0.06, 0.22), 3)       # per diem, lodging, vehicles, subs
            # labor intensity: how much of budget the team is likely to burn — PM effect + noise
            p['burn_bias'] = rnd.gauss(0.01, 0.12)
            if rnd.random() < 0.13:                  # scope creep / field surprises on a minority of jobs
                p['burn_bias'] += rnd.uniform(0.2, 0.65)
            p['size'] = contract
            projects.append(p)
    # project managers carry a persistent over/under-run tendency (a real driver of overruns)
    pm_bias = {}
    for p in projects:
        pm_bias.setdefault(p['pm'], rnd.gauss(0.0, 0.11))
        p['burn_bias'] += pm_bias[p['pm']]
    W.projects = projects
    W.project = {p['id']: p for p in projects}
    W.projects_by_office = defaultdict(list)
    for p in projects:
        if p['status'] in ('In Progress', 'Final Reporting/Review', 'Lab/Curation', 'Transferred Entity'):
            W.projects_by_office[p['office']].append(p)

    _timesheets(W, rnd)
    _finalize_projects(W, rnd)
    _opportunities(W, rnd)
    _expenses(W, rnd)
    return W


def season(office_kind, d):
    """Field-season multiplier on utilization: the regional field offices run hot May–September,
    soften over the winter and dip around the holidays; service lines and corporate are flatter."""
    doy = d.timetuple().tm_yday
    wave = math.sin((doy - 80) / 365.0 * 2 * math.pi)           # peaks late June
    amp = {'regional': 0.07, 'acquired': 0.06, 'international': 0.04, 'service': 0.03}.get(office_kind, 0.0)
    dip = -0.06 if (d.month == 12 and d.day > 18) or (d.month == 1 and d.day < 8) else 0.0
    return amp * wave + dip


def pto_probability(d, flsa):
    """Time off clusters in summer and the year-end holidays."""
    base = 0.055 if flsa == 'Salary Exempt' else 0.045
    if d.month in (7, 8): base *= 1.9
    if d.month == 12: base *= 2.2
    if d.month in (3,) : base *= 1.2
    return base


def _timesheets(W, rnd):
    """Daily timesheet entries for FY2026 through AS_OF, in the ERP's 'Raw Data' shape.

    Relationships built in (so the aggregates look like a real firm's):
      * utilization = person's target + persistent person effect + persistent office effect
        (an AR(1) drift) + the field-season curve + weekly noise;
      * PTO clusters in July/August and late December; holidays follow the calendar;
      * part-time field staff work more in the field season and many are idle in winter;
      * utilized hours are allocated to the office's live projects in proportion to each
        project's remaining labor budget and only inside its start/end window, so big jobs
        absorb more hours and a project's cost is the hours actually charged to it;
      * a small share of project time is logged as non-posting; hourly field staff pick up
        overtime in busy (high-demand) weeks.
    """
    rows = []
    doc = [0]
    office_effect = {o['name']: rnd.gauss(0, 0.035) for o in W.offices}
    drift = {}
    for o in W.offices:
        acc, series = 0.0, []
        for _ in W.weeks:
            acc = 0.82 * acc + rnd.gauss(0, 0.018); series.append(acc)
        drift[o['name']] = series
    W.office_effect = office_effect
    last_week = max(w for w in W.weeks if w <= W.as_of)
    # remaining labor budget per project (hours), consumed as time is charged
    remaining = {}
    for p in W.projects:
        prior = 0.0 if p['start'].year >= 2026 else min(0.85, max(0.0, (dt.date(2026, 1, 1) - p['start']).days / max(1, (p['end'] - p['start']).days)))
        p['prior_share'] = round(prior * 0.8, 3)                 # share of labor spent before FY2026
        labor_budget = p['revised_budget'] * (1 - p['expense_share'])
        remaining[p['id']] = labor_budget * (1 - prior) * (1 + p['burn_bias']) / 62.0   # ~blended $/hr
    for e in W.employees:
        if e['office'] == 'Archive Services':
            continue
        if e['term'] and e['term'] < dt.date(2026, 1, 1):
            continue
        o = W.office[e['office']]
        e_eff = rnd.gauss(0, 0.05)
        e['util_effect'] = e_eff
        pool = W.projects_by_office.get(e['office']) or W.projects_by_office['Red Rock']
        # a person works on a stable handful of projects, refreshed as jobs start and finish
        roster = rnd.sample(pool, min(len(pool), rnd.randint(4, 9)))
        seasonal_pt = e['ft'] == 'Part Time' and rnd.random() < 0.6
        for wi, we in enumerate(W.weeks):
            if we > last_week:
                break
            ws = we - dt.timedelta(days=6)
            days = [ws + dt.timedelta(days=k) for k in range(1, 6)]     # Mon–Fri
            days = [d for d in days if d >= e['hire'] and (not e['term'] or d <= e['term'])]
            if not days:
                continue
            mid = days[len(days) // 2]
            sea = season(o['kind'], mid)
            if e['ft'] == 'Part Time':
                idle = 0.55 - 3.5 * sea if seasonal_pt else 0.15
                if rnd.random() < max(0.05, min(0.85, idle)):
                    continue
                total = q(min(38, max(4, rnd.gauss(22 + 60 * sea, 7))))
            else:
                total = 40.0 if e['flsa'] != 'Salary Exempt' else q(rnd.choice([40, 40, 40, 41, 42, 44]))
            total = q(total * len(days) / 5)
            parts = []   # (project id|'OVHD', task, time type, hours, utilized?, fixed day)
            hol = [d for d in days if d in HOLIDAYS]
            if hol and e['ft'] == 'Full Time':
                for d in hol:
                    parts.append(('OVHD', TASK_HOLIDAY[0], TASK_HOLIDAY[1], 8.0, False, d))
            if e['ft'] == 'Full Time' and rnd.random() < pto_probability(mid, e['flsa']):
                pto = q(rnd.choice([4, 8, 8, 16, 24, 40]))
                task = TASK_VACATION if e['flsa'] == 'Salary Exempt' else TASK_PTO
                pto = min(pto, total - 8 * len(hol))
                if pto > 0:
                    parts.append(('OVHD', task[0], task[1], pto, False, None))
            if rnd.random() < (0.03 if mid.month in (1, 2, 12) else 0.015):
                parts.append(('OVHD', TASK_SICK[0], TASK_SICK[1], q(rnd.choice([4, 8])), False, None))
            if rnd.random() < (0.09 if mid.month in (1, 2) else 0.04):     # training is front-loaded
                parts.append(('OVHD', TASK_TRAINING[0], TASK_TRAINING[1], q(rnd.uniform(1, 4)), False, None))
            used = sum(p[3] for p in parts)
            work = max(0.0, total - used)
            live = [p for p in roster if p['start'] <= mid <= p['end'] + dt.timedelta(days=45)]
            if len(live) < 2:
                fresh = [p for p in pool if p['start'] <= mid <= p['end'] + dt.timedelta(days=45) and p not in roster]
                if fresh:
                    roster.append(rnd.choice(fresh)); live = [p for p in roster if p['start'] <= mid <= p['end'] + dt.timedelta(days=45)]
            demand = sum(max(0.0, remaining[p['id']]) for p in live)
            demand_effect = 0.04 * math.tanh((demand - 400) / 600.0) if live else -0.25
            if e['target'] <= 0:
                rate = max(0.0, rnd.gauss(0.03, 0.03))
            else:
                mu = e['target'] - 0.045 + e_eff + office_effect[e['office']] + drift[e['office']][wi] + sea + demand_effect
                rate = min(1.0, max(0.0, rnd.gauss(mu, 0.075)))
            util = q(work * rate) if live or e['target'] <= 0 else 0.0
            if not live and e['target'] <= 0:
                util = 0.0
            nonutil = q(work - util)
            if util > 0 and live:
                w_ = [max(1.0, remaining[p['id']]) ** 0.7 for p in live]
                k = min(len(live), rnd.choice([1, 1, 2, 2, 3]))
                chosen = []
                cand, cw = list(live), list(w_)
                for _ in range(k):
                    i = rnd.choices(range(len(cand)), weights=cw)[0]
                    chosen.append((cand.pop(i), cw.pop(i)))
                s = sum(wt for _, wt in chosen)
                for p, wt in chosen:
                    h = q(util * wt / s)
                    if h <= 0: continue
                    remaining[p['id']] -= h
                    if rnd.random() < 0.08:
                        np_h = q(h * rnd.uniform(0.1, 0.45))
                        if np_h > 0:
                            parts.append((p['id'], NON_POSTING, 'DIR', np_h, True, None)); h = q(h - np_h)
                    if h > 0:
                        parts.append((p['id'], rnd.choice(p['tasks']), 'DIR', h, True, None))
            if nonutil > 0:
                weights = [w * (3 if e['role'] == 'lead' and 'Proposal' in t else 1)
                             * (5 if e['office'] == 'Sales & Marketing' and ('Business' in t or 'Marketing' in t) else 1)
                             * (1.6 if 'Proposal' in t and mid.month in (1, 2, 3, 10, 11) else 1)
                           for t, _, w in OVHD_TASKS]
                picks = rnd.choices(OVHD_TASKS, weights=weights, k=rnd.randint(1, 3))
                split = [rnd.random() + 0.3 for _ in picks]; s = sum(split)
                for (t_, tt, _), sh in zip(picks, split):
                    h = q(nonutil * sh / s)
                    if h > 0:
                        parts.append(('OVHD', t_, tt, h, False, None))
            if e['flsa'] == 'Hourly Non Exempt' and e['ft'] == 'Full Time' and e['target'] > 0 and live \
                    and rnd.random() < max(0.0, 0.05 + 1.1 * sea + demand_effect):
                p = max(live, key=lambda x: remaining[x['id']])
                h = q(rnd.uniform(1, 9)); remaining[p['id']] -= h
                parts.append((p['id'], rnd.choice(p['tasks']), 'OT', h, True, None))
            if rnd.random() < 0.004 and live:
                parts.append((rnd.choice(live)['id'], NON_POSTING, 'DIR', -q(rnd.uniform(1, 6)), True, None))
            if o['state'] == 'CA' and rnd.random() < 0.02:
                parts.append(('OVHD', NON_POSTING_REGIONAL, 'DIR', q(rnd.uniform(1, 4)), True, None))
            for (pid, task, tt, hrs, is_util, day) in parts:
                if hrs == 0:
                    continue
                if day is not None:
                    chunks = [(day, hrs)]
                else:
                    n = 1 if abs(hrs) <= 4 else min(len(days), max(1, int(abs(hrs) // 6)))
                    ds = rnd.sample(days, min(n, len(days)))
                    per = q(hrs / len(ds)); chunks = [(d, per) for d in ds[:-1]] + [(ds[-1], q(hrs - per * (len(ds) - 1)))]
                for d, h in chunks:
                    if h == 0: continue
                    doc[0] += 1
                    rows.append(dict(emp=str(e['id']), date=d, week=we, project=pid, task=task, tt=tt,
                                     hours=h, util=h if is_util else 0.0, doc=doc[0]))
    W.entries = rows


def _finalize_projects(W, rnd):
    """Roll the charged timesheet hours up into each project's cost, revenue and billing.

    cost_to_date   = prior-year cost (share of labor before FY2026) + FY2026 labor (hours x
                     each person's burdened cost rate) + expenses (a fixed share of labor);
    budget spent % = cost_to_date / revised budget;
    revenue        = fixed fee: % complete (cost / revised budget, capped at 100%) x contract;
                     T&M: charged hours x bill rates (never above contract);
    invoiced       = revenue less an unbilled lag that grows for recent work (30–60 days).
    Hierarchical totals (office, PMO, company) are always sums of these project figures.
    """
    labor = defaultdict(float); hours = defaultdict(float); billed = defaultdict(float)
    recent = defaultdict(float)
    for r in W.entries:
        if r['project'] == 'OVHD':
            continue
        e = W.emp[r['emp']]
        mult = 1.5 if r['tt'] == 'OT' else 1.0
        labor[r['project']] += r['hours'] * e['cost_rate'] * mult
        hours[r['project']] += r['hours']
        if r['task'] != NON_POSTING:
            billed[r['project']] += r['hours'] * e['bill_rate']
        if (W.as_of - r['date']).days <= 45:
            recent[r['project']] += r['hours'] * e['bill_rate']
    for p in W.projects:
        rb = p['revised_budget']
        fy_labor = labor[p['id']]
        prior = p['prior_share'] * rb * (1 - p['expense_share']) * (1 + p['burn_bias'])
        if p['status'] in ('Completed', 'Final Billing/Invoice', 'Cost Tracking on Closed Project') and fy_labor == 0:
            prior = rb * (1 - p['expense_share']) * min(1.3, max(0.7, 1 + p['burn_bias']))
        labor_total = prior + fy_labor
        cost = labor_total * (1 + p['expense_share'] / (1 - p['expense_share']))
        p['labor_cost'] = round(labor_total, 2)
        p['expense_cost'] = round(cost - labor_total, 2)
        p['cost_to_date'] = round(cost, 2)
        p['fy_hours'] = round(hours[p['id']], 2)
        p['budget_spent'] = round(cost / rb, 4) if rb else 0.0
        if p['billing'] == 'Fixed fee':
            pct = min(1.0, p['budget_spent'])
            if p['status'] in ('Completed', 'Final Billing/Invoice'):
                pct = 1.0
            p['revenue'] = round(p['contract'] * pct, 2)
        else:
            prior_rev = p['prior_share'] * p['contract']
            p['revenue'] = round(min(p['contract'], prior_rev + billed[p['id']]), 2)
        p['rev_rec'] = round(p['revenue'] / p['contract'], 4) if p['contract'] else 0.0
        lag = recent[p['id']] * rnd.uniform(0.45, 0.8)
        p['invoiced'] = round(max(0.0, p['revenue'] - lag), 2)
        p['backlog'] = round(p['contract'] - p['revenue'], 2)
        p['margin_to_date'] = round((p['revenue'] - cost) / p['revenue'], 4) if p['revenue'] else None


OPP_STAGES = ['Prospecting', 'Qualification', 'Proposal Development', 'Proposal Submitted', 'Negotiation',
              'Follow-Up', 'Closed Won', 'Closed Lost', 'No Bid']


def _opportunities(W, rnd):
    """Sales pipeline: opportunities per office, with proposal-writing hours from the timesheets
    assigned to them. Win rates differ by office and sector; won opportunities are the source of
    this year's new projects (their start dates line up with the close dates)."""
    opps = []
    oid = [5000]
    win_rate = {o['name']: min(0.55, max(0.2, rnd.gauss(0.36, 0.08))) for o in W.offices}
    field = [o for o in W.offices if o['kind'] in ('regional', 'service', 'acquired', 'international')]
    for o in field:
        members = [e for e in W.employees if e['office'] == o['name'] and e['status'] == 'Active' and e['role'] in ('lead', 'senior')]
        if not members:
            continue
        n = int((o['ft'] + o['pt']) * rnd.uniform(5.2, 7.6))
        for i in range(n):
            oid[0] += rnd.randint(1, 7)
            created = dt.date(2025, 10, 1) + dt.timedelta(days=rnd.randint(0, 360))
            if created > W.as_of: created = W.as_of - dt.timedelta(days=rnd.randint(1, 60))
            close = created + dt.timedelta(days=rnd.randint(20, 120))
            amount = round(math.exp(rnd.gauss(math.log(85000), 1.0)) / 500) * 500
            if close <= W.as_of:
                stage = 'Closed Won' if rnd.random() < win_rate[o['name']] else rnd.choice(['Closed Lost', 'Closed Lost', 'No Bid'])
            else:
                stage = rnd.choice(OPP_STAGES[:6])
            place, kind, svc = rnd.choice(PLACES), rnd.choice(PROJECT_KINDS), rnd.choice(PROJECT_SERVICES)
            client = rnd.choice(W.clients)
            opps.append(dict(id='OPP-%05d' % oid[0], name='%s | %s %s %s' % (client, place, kind, svc) if rnd.random() < 0.4
                             else '%s %s %s' % (place, kind, svc), client=client, office=o['name'],
                             owner=rnd.choice(members)['name'], stage=stage, amount=float(amount),
                             created=created, close=close, sector=rnd.choice(SECTORS),
                             probability={'Prospecting': 10, 'Qualification': 20, 'Proposal Development': 40,
                                          'Proposal Submitted': 50, 'Negotiation': 75, 'Follow-Up': 50,
                                          'Closed Won': 100, 'Closed Lost': 0, 'No Bid': 0}[stage],
                             proposal_hours=0.0))
    # assign proposal-writing / editing hours to the office's open-at-the-time opportunities
    by_office = defaultdict(list)
    for op in opps:
        by_office[op['office']].append(op)
    W.proposal_hours = []
    for r in W.entries:
        if r['project'] != 'OVHD' or 'Proposal' not in r['task']:
            continue
        e = W.emp[r['emp']]
        cands = [op for op in by_office.get(e['office'], []) if op['created'] <= r['date'] <= op['close']]
        if not cands:
            cands = by_office.get(e['office'], []) or opps
        op = rnd.choice(cands)
        op['proposal_hours'] += r['hours']
        r['opportunity'] = op['id']
        W.proposal_hours.append(dict(emp=r['emp'], date=r['date'], week=r['week'], opp=op['id'], hours=r['hours'], task=r['task']))
    W.opportunities = opps


EXPENSE_TYPES = [('Per Diem', 55, 0.30), ('Lodging', 118, 0.25), ('Mileage', 46, 0.2), ('Field Supplies', 64, 0.1),
                 ('Vehicle Rental', 210, 0.06), ('Airfare', 380, 0.04), ('Equipment Rental', 150, 0.05)]
ER_STATES = ['Draft', 'Submitted', 'Approved', 'Paid', 'Declined', 'Reversal', 'Partially Approved']


def _expenses(W, rnd):
    """Staff expense reports. Field staff on project weeks generate per diem / lodging / mileage
    lines charged to the projects they worked; reports are filed a few days after the trip and
    move Draft -> Submitted -> Approved -> Paid, so only recent reports are still open."""
    by_emp_week = defaultdict(lambda: defaultdict(float))
    for r in W.entries:
        if r['project'] != 'OVHD' and r['hours'] > 0:
            by_emp_week[(r['emp'], r['week'])][r['project']] += r['hours']
    reports, er = [], [20000]
    lines = []
    for (emp, week), projs in sorted(by_emp_week.items()):
        e = W.emp[emp]
        field_hours = sum(projs.values())
        if field_hours < 12 or rnd.random() > (0.28 if e['role'] in ('field', 'senior') else 0.06):
            continue
        er[0] += 1
        filed = week + dt.timedelta(days=rnd.randint(1, 12))
        if filed > W.as_of + dt.timedelta(days=2):
            filed = W.as_of + dt.timedelta(days=rnd.randint(0, 2))
        age = (W.as_of - filed).days
        if age <= 7:
            state = rnd.choices(['Draft', 'Submitted', 'Approved'], weights=[3, 4, 2])[0]
        elif age <= 30:
            state = rnd.choices(['Draft', 'Submitted', 'Approved', 'Paid'], weights=[1, 2, 4, 5])[0]
        else:
            state = rnd.choices(['Draft', 'Submitted', 'Approved', 'Paid', 'Declined', 'Reversal'], weights=[0.25, 0.35, 2, 22, 0.4, 0.2])[0]
        total = 0.0
        pids = list(projs)
        for d in range(rnd.randint(1, 4)):
            et, base, _ = rnd.choices(EXPENSE_TYPES, weights=[x[2] for x in EXPENSE_TYPES])[0]
            amt = round(base * rnd.uniform(0.6, 2.4) * (rnd.randint(1, 4) if et in ('Per Diem', 'Lodging') else 1), 2)
            pid = rnd.choice(pids)
            lines.append(dict(er='EXP%06d' % er[0], emp=emp, project=pid, type=et, amount=amt,
                              date=week - dt.timedelta(days=rnd.randint(0, 4))))
            total += amt
        gl = filed + dt.timedelta(days=rnd.randint(0, 3)) if state in ('Submitted', 'Approved', 'Paid', 'Draft') else filed
        reports.append(dict(er='EXP%06d' % er[0], emp=emp, week=week, filed=filed, gl=gl, state=state,
                            total=round(total, 2), last_modified=min(W.as_of + dt.timedelta(days=2), filed + dt.timedelta(days=rnd.randint(0, 20))),
                            reason=rnd.choice(['Field travel', 'Fieldwork per diem', 'Site visit', 'Monitoring travel',
                                               'Survey mobilization', 'Client meeting', 'Training travel'])))
    W.expense_reports, W.expense_lines = reports, lines


def world(rebuild=False):
    if not rebuild and os.path.exists(CACHE):
        with open(CACHE, 'rb') as f:
            W = World(); W.__dict__.update(pickle.load(f)); return W
    W = build()
    os.makedirs(os.path.dirname(CACHE), exist_ok=True)
    with open(CACHE, 'wb') as f:
        pickle.dump(W.__dict__, f)
    return W


if __name__ == '__main__':
    W = world(rebuild=True)
    act = [e for e in W.employees if e['status'] == 'Active']
    print('offices', len(W.offices), 'employees', len(W.employees), 'active', len(act),
          'projects', len(W.projects), 'timesheet rows', len(W.entries))
