"""Synthetic round-5 files (round 4 with rep phones in the data file, no People Report, presence areas per store, tracker rep names).
Synthetic round-4 files (round 3 plus Store Zip, a revisit, Vendor Work Validation Fail, the BDS No Show and Data Review tracker and Parts Orders Detail).
Synthetic round-2 files. Fake stores, towns and people. Maximo export (CM + PM rows), Power BI "Export" sheet, People Report.
Dates are relative to today so 'this week' and 'last week' are always real. Writes ./samples/*.xlsx and ./samples/expected.json."""
import openpyxl, datetime as dt, json, os, random, collections

os.makedirs('samples', exist_ok=True)
random.seed(11)
NOW = dt.datetime.now(dt.timezone.utc).replace(second=0, microsecond=0)
TODAY = NOW.date()
MON = TODAY - dt.timedelta(days=TODAY.weekday())  # this week's Monday
OFF = {'EDT': -4, 'CDT': -5, 'MDT': -6, 'MST': -7, 'PDT': -7, 'AKDT': -8, 'HST': -10}

MH = ('work_order_number', 'work_order_description', 'work_order_status', 'work_type', 'region', 'group', 'district', 'site', 'city', 'state', 'not_to_exceed', 'earliest_start', 'repair_expectation', 'response_expectation', 'controllable_execution', 'arrive_by', 'work_order_eta', 'first_check_in', 'last_check_in', 'last_check_out', 'repair_by', 'invoice_by', 'craft', 'vendor', 'asset_name', 'asset_number', 'asset_description', 'person_group', 'asset_tag', 'area_of_store', 'problem', 'lvm_status', 'pml_reviewed', 'redirect_reason', 'last_visit_status', 'last_updated', 'reported_date', 'job_plan', 'converted_response', 'converted_repair', 'status_date', 'work_logs', 'last_vendor_work_log', 'work_order_id', 'pm_contract', 'lvm_l1_esc', 'lvm_l2_esc', 'lvm_l3_esc', 'lvm_l4_esc', 'lvm_l5_esc', 'vendor_wo_val', 'external', 'id')
STORES = [  # site, town, state, zone, region, group, district, DM, rep
    ('T0102', 'Maple Falls', 'PA', 'EDT', 400, 493, 421, 'Dana Test-DM1', 'Rory Test-Rep1'),
    ('T0347', 'Cedar Point', 'OH', 'EDT', 400, 493, 425, 'Dana Test-DM1', 'Sky Test-Rep2'),
    ('T0581', 'Riverbend', 'IL', 'CDT', 200, 210, 214, 'Eli Test-DM2', 'Quinn Test-Rep3'),
    ('T0733', 'Pine Hollow', 'TX', 'CDT', 200, 230, 233, 'Eli Test-DM2', None),
    ('T0912', 'Mesa Verde', 'AZ', 'MST', 300, 310, 312, 'Fay Test-DM3', 'Jules Test-Rep4'),
    ('T1048', 'Silver Lake', 'CO', 'MDT', 300, 320, 322, 'Fay Test-DM3', 'Robin Test-Rep5'),
    ('T1290', 'Harbor View', 'CA', 'PDT', 300, 340, 341, 'Gus Test-DM4', 'Sam Test-Rep6'),
    ('T1455', 'Glacier Bay', 'AK', 'AKDT', 300, 340, 349, 'Gus Test-DM4', 'Alex Test-Rep7'),
    ('T1622', 'Palm Grove', 'HI', 'HST', 300, 340, 349, 'Gus Test-DM4', 'Kit Test-Rep8'),
    ('T1877', 'Oak Ridge', 'GA', 'EDT', 100, 120, 127, 'Hal Test-DM5', 'Lee Test-Rep9')]  # T1877 is not in the Power BI file
NO_VISITS = 'T1877'
DESC = ['Apple discovery table displays missing', 'Audio wall demo units not powered', 'Gaming endcap needs reset', 'Phone fixture security tethers broken']
PROBLEMS = ['BDS_NO_SHOW', 'BDS_DATA_REVIEW']

def maximo(t, zone):
    return (t + dt.timedelta(hours=OFF[zone])).strftime('%m/%d/%Y %I:%M %p ') + zone

LATE = []
NAMED = []

def mrow(n, status, wtype, store, due_h, problem):
    site, town, st, zone, reg, grp, dist = store[:7]
    los = NOW + dt.timedelta(hours=due_h)
    if wtype == 'CM': LATE.append(status in ('APPR', 'VDECWO') and due_h < 0)
    es = los - dt.timedelta(days=10)
    ab = los - dt.timedelta(days=7)
    d = dict.fromkeys(MH)
    d.update(work_order_number=9160000000 + n, work_order_description=(f'{site} Oura Ring 5 Sizing Ring Endcap Fixture DPCI 057-03-0470 marked as not present' if site == 'T0581' and not NAMED else f'{site} {random.choice(DESC)}'), work_order_status=status,
             work_type=wtype, region=reg, group=grp, district=dist, site=site, city=town, state=st, not_to_exceed=0,
             earliest_start=maximo(es, zone), arrive_by=maximo(ab, zone), repair_by=maximo(los, zone), vendor=308017,
             asset_name='EXPERIENTIAL_DISPLAYS', area_of_store='ELECTRONICS', problem=problem, lvm_status='Not Available',
             last_updated=(NOW - dt.timedelta(hours=1)).strftime('%m/%d/%Y %I:%M %p UTC'), reported_date=maximo(es, zone),
             status_date=maximo(es, zone), work_order_id=135000000 + n, external=1, id=9160000000 + n,
             vendor_wo_val='Pass' if n % 4 == 0 else 'Fail' if n % 7 == 0 else None)
    if site == 'T0581' and wtype == 'CM' and not NAMED: NAMED.append(n)
    return [d[h] for h in MH]

def maximo_rows():
    R, n = [], 0
    plan = [('APPR', -300), ('APPR', -30), ('APPR', 5), ('APPR', 200), ('VDECWO', -90), ('VDECWO', 20), ('VDECWO', 150), ('VINPRG', -50), ('VINPRG', 100), ('VDECWC', -10)]
    for k, store in enumerate(STORES):
        for j in range(2 + k % 3):
            status, h = plan[(k * 3 + j) % len(plan)]
            n += 1
            R.append(mrow(n, status, 'CM', store, h + k, PROBLEMS[n % 2]))
    for k in range(6):  # PM rows: skipped
        n += 1
        R.append(mrow(n, 'APPR', 'PM', STORES[k], -100, 'HQ_USE_ONLY'))
    return R

def save_maximo(name, rows, headers=MH):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'in'
    ws.append(list(headers))
    for r in rows: ws.append(r)
    wb.save(f'samples/{name}')

rows = maximo_rows()
save_maximo('export_ok.xlsx', rows)
save_maximo('export_pm_only.xlsx', [r for r in rows if r[3] == 'PM'])
i = MH.index('arrive_by')
save_maximo('export_missing_col.xlsx', [r[:i] + r[i + 1:] for r in rows], MH[:i] + MH[i + 1:])
bad = [r[:] for r in rows]; bad[2][2] = 'WAPPR'
save_maximo('export_bad_status.xlsx', bad)

# ---- Power BI "Export" sheet
PH = ['FOM Sup Name', 'FOM Territory', 'Reps Manager', 'Rep Name', 'Assignment ID', 'Rep Person ID', 'Rep Position Type', 'Rep Position Title', 'Rep Position ID', 'Rep Home Phone', 'Rep Mobile Phone', 'Rep Email', 'Wave Active', 'Wave Start Date', 'Wave End Date', 'BDS Store ID', 'Chain Store #', 'Current Priority', 'Call Form Name', 'Wave Name', 'Type of Program', 'Visit Role', 'Visit Type', 'Assignment Creation Date', 'VisitGoalReasonName', 'QA Reason', 'Materials to Rep Description', 'Parts to Rep Description', 'Materials to Store Description', 'Parts to Store Description', 'Client Service Lead', 'Store Name', 'Chain Name', 'Store Address', 'Store City', 'Store State', 'Store Zip', 'Store Phone #', 'Market Name (CBSA)', 'Rep to Store Distance (mi)', 'Store Ranking', 'Reason Not Brought', 'Visit Date Scheduled', 'Date Report Entered', 'Scheduled Time In', 'Scheduled Time Out', 'Total Goal Hours', 'In-Store Hours Actual', 'In-Store Differential', 'Hours Reason Adjusted', 'Shipped Materials Required', 'Materials Destination', 'Have Materials Shipped', 'Date Available', 'Admin Time Hours', 'Original Drive Time', 'Reported Drive Time', 'Reason Adjusted', 'Priority 7 days out']
BF = 'Target Break-Fix Continuity 2026'
EXP = collections.OrderedDict()
prow, aid = [], 0

def prow_add(store, wk_off, form=BF, chain='Target', state=None, day=None, reported=True, rep=True, goal=None):
    """One assignment. wk_off: weeks from this Monday (-3..+1). state: 'visited' | 'scheduled' | 'none'."""
    global aid
    aid += 1
    site, town, st, zone, reg, grp, dist, dm, repname = store
    k = STORES.index(store)
    start = MON + dt.timedelta(weeks=wk_off)
    d = dict.fromkeys(PH)
    d.update({'FOM Sup Name': 'Pat Test-Sup', 'FOM Territory': dm, 'Reps Manager': 'Pat Test-Sup' if k < 5 else 'Max Test-Sup', 'Rep Name': repname if rep and repname else '(Open) - OPEN - Test',
              'Assignment ID': 5000 + aid, 'Rep Person ID': (900000 + k) if rep and repname else None, 'Rep Mobile Phone': f'555010{k}' if rep and repname else None, 'Rep Home Phone': f'555020{k}' if rep and repname and k % 2 == 0 else None, 'Rep Email': 'not-shown@example.com', 'Wave Start Date': dt.datetime.combine(start, dt.time()),
              'Wave End Date': dt.datetime.combine(start + dt.timedelta(days=4), dt.time(23, 59)), 'BDS Store ID': 70000 + int(site[1:]), 'Chain Store #': int(site[1:]),
              'Current Priority': 'Priority 1 - First Tier', 'Call Form Name': form, 'Wave Name': f'Wave {703 + wk_off} Tier 1', 'Visit Type': 'Original-1 Rep',
              'Store Name': f'Target {town}', 'Chain Name': chain, 'Store City': town, 'Store State': st, 'Store Address': f'{int(site[1:])} Test Way', 'Store Zip': f'{10000 + int(site[1:])}', 'VisitGoalReasonName': goal, 'Total Goal Hours': 1})
    if state in ('visited', 'scheduled'):
        d['Visit Date Scheduled'] = dt.datetime.combine(start + dt.timedelta(days=day), dt.time())
    if state == 'visited':
        d['Date Report Entered'] = dt.datetime.combine(start + dt.timedelta(days=day), dt.time())
        d['In-Store Hours Actual'] = 1.1
    prow.append([d[h] for h in PH])
    return d

def label(d0, d1):
    f = lambda x: x.strftime('%b ') + str(x.day)
    return f'{f(d0)} – {f(d1)}'

for k, store in enumerate(STORES):
    if store[0] == NO_VISITS: continue
    n = int(store[0][1:])
    weeks = []
    for wk in (-3, -2, -1, 0):
        # visited, except: store 1 missed last week (not scheduled), store 2 scheduled last week with no report, store 3 has no rep this week
        state = 'visited'
        if k == 1 and wk == -1: state = 'none'
        if k == 2 and wk == -1: state = 'scheduled'
        if wk == 0 and k >= 5: state = 'scheduled'
        if wk == 0 and k == 4: state = 'none'
        rep = not (k == 3 and wk == 0)
        if k == 3 and wk == 0: state = 'none'
        day = 1 + (k % 3)
        prow_add(store, wk, state=state, day=day, rep=rep, goal='QA Revisit' if k == 0 and wk == -3 else None)
        s0 = MON + dt.timedelta(weeks=wk)
        weeks.append({'wave': label(s0, s0 + dt.timedelta(days=4)), 'rep': (store[8] if rep and store[8] else 'No rep assigned'),
                      'status': {'visited': 'Visited ' + label(s0 + dt.timedelta(days=day), s0 + dt.timedelta(days=day)).split(' – ')[0],
                                 'scheduled': 'Scheduled ' + label(s0 + dt.timedelta(days=day), s0).split(' – ')[0] + ', no report yet',
                                 'none': 'Not scheduled'}[state]})
    EXP[str(n)] = {'dm': store[7], 'weeks': list(reversed(weeks))}
    prow_add(store, 1, state='scheduled', day=2)  # next week: not loaded (wave not started)
    prow_add(store, 0, form='Target Nintendo Holiday Q4 2026', state='scheduled', day=2)  # other form: dropped
    prow_add(store, 0, form=BF, chain='Best Buy', state='visited', day=2)  # other chain: dropped

started = [r for r in prow if r[PH.index('Chain Name')] == 'Target' and r[PH.index('Call Form Name')] == BF and r[PH.index('Wave Start Date')].date() <= MON]
nxt = MON + dt.timedelta(days=7)

def save_pbi(name, rows, headers=PH, sheet='Export', footer=True):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = sheet
    ws.append(list(headers))
    for r in rows: ws.append(r)
    if footer:
        ws.append([]); ws.append(['No filters applied']); ws.append(['Exported data exceeded the allowed volume. Some data may have been omitted.'])
    wb.save(f'samples/{name}')

save_pbi('data_ok.xlsx', prow)
j = PH.index('Date Report Entered')
save_pbi('data_missing_col.xlsx', [r[:j] + r[j + 1:] for r in prow], PH[:j] + PH[j + 1:])
save_pbi('data_no_breakfix.xlsx', [r for r in prow if r[PH.index('Call Form Name')] != BF])
save_pbi('data_not_started.xlsx', [r for r in prow if r[PH.index('Wave Start Date')].date() > MON])
save_pbi('data_wrong_sheet.xlsx', prow, sheet='Data')

mon_rows = [r for r in rows if r[3] == 'CM']
cnt = collections.Counter(r[2] for r in mon_rows)
exp = {'monday': MON.isoformat(), 'cm': len(mon_rows), 'late': sum(LATE), 'in_week': sum(1 for r in started if r[PH.index('Visit Date Scheduled')] and MON <= r[PH.index('Visit Date Scheduled')].date() < nxt), 'pm': sum(r[3] == 'PM' for r in rows), 'status': dict(cnt), 'problems': dict(collections.Counter(r[MH.index('problem')] for r in mon_rows)),
       'started_rows': len(started), 'stores': len({r[PH.index('Chain Store #')] for r in started}),
       'vis_status': dict(collections.Counter(('Completed' if r[j] else ('Sched this week' if r[PH.index('Visit Date Scheduled')] and MON <= r[PH.index('Visit Date Scheduled')].date() < nxt else 'Not sched this week')) for r in started)),
       'cards': EXP, 'no_visits_site': NO_VISITS}
json.dump(exp, open('samples/expected.json', 'w'), indent=1, default=str)
print(json.dumps({k: v for k, v in exp.items() if k != 'cards'}))


# ---- Presence and Functionality (fake): rows 1-40 are the report's dashboard area, the table header is on row 41
DISPLAYS = ['Oura Ring 5 Sizing Ring Endcap Fixture DPCI 057-03-0470', 'Bose 2ft Flex + SoundLink Plus Display DPCI 008-94-0219', 'Fitbit/Google Base Display DPCI 057-94-3484',
            'Logitech D56 Non-Interactive Mouse Display DPCI 056-94-0430', 'Hiro 65” (Red Channel) DPCI 008-94-0244']
PH2 = ['Visit Date', 'Region', 'Group', 'District', 'Store #', 'Display Name', 'Present?', 'Functional? (Are ALL interactive elements of the display functional?)', 'Display Source\n(Where the display is shipped from)', 'Stock Status\n(Display inventory availability per Target HQ/Brand)', 'Reason display is not Present', 'Reason display is not Functional', 'Brand']
def presence_file(name, mutate=None, sheet='Presence and Functionality', header_row=41):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = sheet
    ws['A1'] = 'Interactives Program Presence and Functionality'; ws['A22'] = 'Presence and Functionality Calculator'; ws['A28'] = '=COUNTIF(Table4[Present?],"Yes*")/COUNTA(Table4[Present?])'
    ws['G31'] = 'Region'; ws['H31'] = 'No'
    for c, t in enumerate(PH2, 1): ws.cell(header_row, c, t)
    rows, exp = [], {}
    for k, st in enumerate(STORES):
        site = st[0]
        if site in ('T0347', NO_VISITS): continue   # not in the report
        vd = dt.datetime.combine(MON - dt.timedelta(days=3 - k % 3), dt.time())
        recs = []
        for j, dsp in enumerate(DISPLAYS):
            present, func, why_np, why_nf = 'Yes', 'Yes', None, None
            if site == 'T0581' and j == 0: present, func, why_np = 'No', None, 'Does not show on stores planogram'
            elif site == 'T0581' and j == 3: present, func, why_np = 'No', None, 'Display not in store, ordered through BDS Target Support Center'
            elif site == 'T0581' and j == 2: func, why_nf = 'No', 'Damage'
            elif site == 'T0102' and j == 1: func, why_nf = 'No', 'Power, Demo Content'
            recs.append((vd, st[4], st[5], st[6], int(site[1:]), dsp, present, func, 'Brand/Vendor', 'In Stock', why_np, why_nf, dsp.split()[0]))
        rows += recs
        asked = [r for r in recs if r[7] is not None]
        bad = [r for r in recs if r[6] == 'No' or r[7] == 'No']
        exp[str(int(site[1:]))] = {'date': vd.strftime('%b ') + str(vd.day) + ', ' + str(vd.year), 'present': f"{sum(r[6]=='Yes' for r in recs)} of {len(recs)}",
                                   'working': f"{sum(r[7]=='Yes' for r in asked)} of {len(asked)} present displays",
                                   'bad': [('Not present' if r[6] == 'No' else 'Not working', r[5], r[10] or r[11], '057-03-0470' in r[5] and site == 'T0581') for r in bad]}
    if mutate: mutate(rows)
    for i, r in enumerate(rows): [ws.cell(header_row + 1 + i, c + 1, v) for c, v in enumerate(r)]
    wb.save(f'samples/{name}')
    return exp, len(rows)
exp3, nrows = presence_file('presence_ok.xlsx')
presence_file('presence_wrong_sheet.xlsx', sheet='Sheet1')
presence_file('presence_no_header.xlsx', header_row=41)  # replaced below
wb = openpyxl.load_workbook('samples/presence_no_header.xlsx'); wb.active['A41'] = 'Date'; wb.save('samples/presence_no_header.xlsx')
wb = openpyxl.load_workbook('samples/presence_ok.xlsx'); wb.active['K41'] = 'Reason not there'; wb.save('samples/presence_missing_col.xlsx')
wb = openpyxl.load_workbook('samples/presence_ok.xlsx'); wb.active['G45'] = 'Maybe'; wb.save('samples/presence_bad_answer.xlsx')
e = json.load(open('samples/expected.json')); e['presence'] = exp3; e['presence_rows'] = nrows; e['named_wo'] = 9160000000 + NAMED[0]
json.dump(e, open('samples/expected.json', 'w'), indent=1, default=str)
print('presence rows', nrows, 'named wo', 9160000000 + NAMED[0])


# ---- BDS No Show and Data Review tracker (fake notes). Work orders by number; one stored as "<n>.0" text like the real file.
cm = [r for r in rows if r[3] == 'CM']
TH = ['Work Order', 'Site', 'Description', 'Problem Code', 'Vendor Work Validation', 'Status', 'Status Date', 'Report Date', 'LOS Finish', 'Store Phone', 'Comments ', 'Action on Maximo', 'Maximo', 'Closed In Maximo', 'Rep', 'FOM (DM), DC, RDS', 'Date of Update to FOM & PML', 'Status\r\nEmailed-FOM\r\n', 'Response from FOM or PML', None]
wb = openpyxl.Workbook(); ns = wb.active; ns.title = 'No show 9.7.20 - Present'; dr = wb.create_sheet('Data Review 9.7.20- Present'); rs = wb.create_sheet('Responses')
ns.append(TH); dr.append([h for h in TH if h != None][:15] + ['FOM', 'Date of Update to FOM & PML', 'Status\r\nEmailed-FOM\r\n', 'Response from FOM or PML'])
rs.append(['Category', 'WO Scenario', 'BDS Status Update, Pending Resolution', 'Log Note Verbiage back to Target'])
rs.append(['PML Hours\nBDS NO SHOW ', 'Test scenario: visit outside PML hours', 'Test status: pending PML', 'Test log note: the visit was outside PML hours.'])
rs.append(['Submission Issue', 'Test scenario: duplicate WO', None, 'Test log note: this WO duplicates an earlier one.'])
# The four scenarios Reply can tell apart from the files, worded like the real Responses sheet.
rs.append(['Missed Visit\nBDS NO SHOW', 'Missed visit reported, BDS did miss visit', 'BDS will move WO to VINPRG', 'BDS confirmed visit was missed; next rep visit scheduled for XX/XX.'])
rs.append(['Missed Visit\nBDS NO SHOW', 'Missed visit reported, BDS did not miss visit', 'BDS will move WO to VDECWO', 'BDS rep visited on XX/XX/XX at X:XX; next rep visit scheduled for XX/XX.'])
rs.append(['Accepted\nDATA REVIEW', 'If the rep reporting does not align with what the PML is stating', 'BDS will move WO to VINPRG', 'Test log note: BDS will notify the rep\'s Field Manager.'])
rs.append(['MISC\nDATA REVIEW', 'If the rep reporting already aligns with what the PML is stating is present or non functional', 'BDS will move WO to VDECWO', 'Test log note: rep reporting aligns, declining this WO.'])
def mtext(d): return d.strftime('%m/%d/%Y %I:%M %p') + ' EDT'
TRK = {}
REPS_IN_TRACKER = {'T0102': 'Rory Test-Rep1\nNot In Data', 'T0347': 'sky test-rep2, Quinn Test-Rep3'}
for i, r in enumerate(cm[:8]):
    wo = r[0]; site = r[7]
    sheet = ns if r[30] == 'BDS_NO_SHOW' else dr
    upd = dt.datetime.combine(TODAY - dt.timedelta(days=2 + i), dt.time())
    sd = dt.datetime.combine(TODAY - dt.timedelta(days=3 + i), dt.time(9, 15))
    when = mtext(sd) if i % 2 == 0 else sd
    vals = [f'{wo}.0' if i == 0 else wo, site, r[1], r[30], 'Pass', 'APPR', when, when, when, None,
            f'Test note {i}: rep visited, display reset.\nFollow-up sent.' if i % 2 == 0 else f'Test note {i}: no visit found.', 'VINPRG\r\nVCOMP' if i % 2 == 0 else 'VDECWO', 'Y', 'N' if i % 3 else 'Y', REPS_IN_TRACKER.get(site, 'Test Rep'), 'Test DM',
            upd, 'Y', 'Test response from PML' if i == 0 else None]
    sheet.append(vals)
    TRK[str(wo)] = {'sheet': 'No show' if sheet is ns else 'Data Review', 'action': vals[11].replace('\r\n', '\n'), 'update': upd.strftime('%b ') + str(upd.day) + ', ' + str(upd.year)}
# History: older rows not in the export, spread over the last 14 months, so No Show has months, repeat stores and reps.
HIST = []
for j in range(18):
    sheet = ns if j % 3 else dr
    store = ['T0581', 'T0581', 'T0347', 'T0912', 'T0581', 'T1048'][j % 6]
    sd = dt.datetime.combine(TODAY - dt.timedelta(days=40 + j * 23), dt.time(10, 0))
    wo = 9150000000 + j
    vals = [wo, store, f'{store} Test history row {j}', 'BDS_NO_SHOW' if sheet is ns else 'BDS_DATA_REVIEW', 'Pass', 'APPR', sd, sd, sd, None,
            f'Test history note {j}', 'VINPRG\r\nCOMP', 'Y', 'Y', ['Quinn Test-Rep3', 'Rory Test-Rep1', 'Unknown Test-Person'][j % 3], 'Test DM', sd, 'Y' if j % 2 else 'N', 'Test reply' if j % 4 == 1 else None]
    sheet.append(vals)
    HIST.append({'wo': str(wo), 'store': str(int(store[1:])), 'sheet': 'No show' if sheet is ns else 'Data Review', 'days': 40 + j * 23})
wb.save('samples/tracker_ok.xlsx')
wb = openpyxl.Workbook(); wb.active.title = 'Responses'; wb.active.append(['Scenario', 'Response']); wb.save('samples/tracker_no_sheet.xlsx')

# ---- Parts Orders Detail (fake). Matched by store; Ticket ID is not a work order number.
PCOLS = ['Parts Order ID', 'Program Name', 'Parts Display', 'Part Name', 'Part Number', 'Ticket ID', 'Transmission Status', 'Reason', 'Stock Status Name', 'Tracking Number', 'Ship Date', 'Shipping Provider Name', 'Placement Status Name', 'Status Notes', 'Order Created Date', 'Chain Store Number', 'Issue Type']
def porder(i, store, part, reason, created, track=None, sent=None, placed=None, notes=None, typ='Part Ordered'):
    d = dict.fromkeys(PCOLS)
    d.update({'Parts Order ID': 800 + i, 'Program Name': 'Target Continuity', 'Parts Display': 'Test Display Kit', 'Part Name': part, 'Part Number': f'P-{100 + i}', 'Ticket ID': 3300000 + i,
              'Transmission Status': sent, 'Reason': reason, 'Stock Status Name': 'In Stock' if reason else None, 'Tracking Number': track, 'Ship Date': dt.datetime.combine(TODAY - dt.timedelta(days=1), dt.time()) if track else None,
              'Shipping Provider Name': 'FedEx' if track else None, 'Placement Status Name': placed, 'Status Notes': notes, 'Order Created Date': dt.datetime.combine(TODAY - dt.timedelta(days=created), dt.time()),
              'Chain Store Number': store, 'Issue Type': typ})
    return [d[h] for h in PCOLS]
PO = [porder(1, '581', 'Test Riser Shelf', 'Broken', 5, track='123456789012', sent='shipped by C-Exchange'),
      porder(2, '581', 'Test Power Cable', 'Missing', 2),
      porder(3, '102', 'Test Header Sign', 'Missing', 6, sent='canceled by account team', placed='Cancelled', notes='Test: wrong part ordered'),
      porder(4, '102', None, None, 3, typ='Troubleshooting Only')]
wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Parts Orders Detail'; ws.append(PCOLS); [ws.append(r) for r in PO]; wb.save('samples/parts_ok.xlsx')
wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Parts Orders Detail'; ws.append([c for c in PCOLS if c != 'Tracking Number']); wb.save('samples/parts_missing_col.xlsx')
e = json.load(open('samples/expected.json')); e['tracker'] = TRK; e['tracker_history'] = HIST; e['parts_orders'] = len(PO); e['revisit_store'] = '102'
json.dump(e, open('samples/expected.json', 'w'), indent=1, default=str)
print('tracker rows', len(TRK), 'parts', len(PO))
