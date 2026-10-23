// Clinic Eligibility file (central Sohar -> this add-in): pack -> workbook -> load() -> prepare -> derive, Clinic Work writes,
// and the nav. Uses the byte-identical copies of Sohar's core.js / rules.js. Synthetic data only (made-up names, MRNs 9000xx,
// made-up payer rules); dates are relative to today because the add-in derives at the current time.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'); const path = require('path'); const vm = require('vm');
const C = require('../core.js');
const R = require('../rules.js');
const PR = require('../payer-rules.js');

const day = n => { const d = new Date(); d.setHours(12, 0, 0, 0); d.setDate(d.getDate() + n); return d; };
const mdy = n => { const d = day(n); return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`; };
const iso = n => { const d = day(n); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
const NOW = new Date();

// ---------- the admin side (as Sohar builds it) ----------
const HDR = ['Pat Name', 'Appt Dt', 'Beg Tm', 'Event', 'Loc Name', 'Per Nbr', 'Md Rc', 'Birth Dt', 'Sts', 'Guar Name', 'Guar Birth Dt', 'Guar Zip',
  'Payer Name (Cht Guar)', 'Pol Nbr (Cht Guar)', 'Sec Payer (Cht Guar)', 'Rendering'];
const row = o => HDR.map(k => (o[k] == null ? '' : o[k]));
const MD = 'Reyes, Sam MD', NP = 'Lark, Jordan PMHNP', LOC = 'San Rafael';
const LOCSTATE = { 'San Rafael': 'CA', 'Walnut Creek': 'CA' };
const MAP = [
  { nextgenPayerName: 'Blue Shield CA', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Exact' },
  { nextgenPayerName: 'Cigna', payerId: '62308', payerName: 'Cigna', confidence: 'Exact' },
];
const CFG = { asOf: new Date(Date.UTC(day(0).getFullYear(), day(0).getMonth(), day(0).getDate())), scope: 'upcoming', locationState: LOCSTATE, teleKeywords: ['Tele'], specialty: 'PSY', payerMap: MAP };
const ap = (per, mrn, name, dob, n, time, event, loc, payer, pol, extra) => row({ 'Per Nbr': per, 'Md Rc': mrn, 'Pat Name': name, 'Birth Dt': dob, 'Appt Dt': mdy(n),
  'Beg Tm': time, Event: event, 'Loc Name': loc, Sts: 'Expected', 'Payer Name (Cht Guar)': payer, 'Pol Nbr (Cht Guar)': pol, 'Guar Zip': '94901', Rendering: MD, ...(extra || {}) });
const IMPORT = () => [HDR,
  ap('501', '900001', 'Quillfeather, Testa', '4/5/1981', 5, '9:00 AM', 'Telepsych', LOC, 'Cigna', 'C501'),
  ap('501', '900001', 'Quillfeather, Testa', '4/5/1981', 7, '10:00 AM', 'Esketamine Treatment', LOC, 'Cigna', 'C501', { Rendering: NP }),
  ap('502', '900002', 'Brambleton, Bravoline', '6/7/1975', 4, '10:30 AM', 'Telepsych', LOC, 'Blue Shield CA', 'BS502', { 'Sec Payer (Cht Guar)': 'Cigna' }),
  ap('503', '900003', 'Kiddington, Juniper', '8/9/2012', 6, '1:00 PM', 'Telepsych', LOC, '', '', { 'Guar Name': 'Parentisson, Guardia', 'Guar Birth Dt': '3/3/1970' }),
  ap('504', '900004', 'Distantly, Otherine', '2/2/1960', 5, '9:00 AM', 'Telepsych', 'Walnut Creek', 'Cigna', 'C504'),
  ap('505', '900005', 'Onenamesson', '1/1/1966', 8, '2:00 PM', 'Telepsych', LOC, 'Cigna', 'C505'),
];
const NEVER = ['Quillfeather', 'Testa', 'Brambleton', 'Bravoline', 'Kiddington', 'Juniper', 'Parentisson', 'Guardia', 'Onenamesson', 'Distantly', 'Otherine', '4/5/1981', '04/05/1981', '94901', '900004'];
const VER_HDR = ['requestId', 'id', 'url', 'created', 'resolved', 'memberId', 'payerName', 'status', 'statusReason', 'firstName', 'lastName',
  'planName', 'innCopay', 'innTotalDeductible', 'innRemainingDeductible', 'innOutOfPocket', 'innOutOfPocketRemaining'];
const DISC_HDR = ['requestId', 'id', 'created', 'resolved', 'firstName', 'lastName', 'dateOfBirth', 'memberId', 'responsePayerId', 'responsePayerName'];
const pick = (hdr, o) => hdr.map(k => (o[k] == null ? '' : String(o[k])));
const FILES = () => [
  { name: 'verification output.csv', rows: [VER_HDR, ...[
    { requestId: 'v-501', id: '900001', url: 'https://app.soharhealth.com/verifications/v-501', created: iso(-1), resolved: iso(-1) + 'T17:00:00Z', memberId: 'C501', payerName: 'Cigna', status: 'complete.eligible',
      firstName: 'Testa', lastName: 'Quillfeather', planName: 'SYNTH OPEN ACCESS', innCopay: '30', innTotalDeductible: '1000', innRemainingDeductible: '100', innOutOfPocket: '4000', innOutOfPocketRemaining: '2500' },
    { requestId: 'v-502', id: '900002', created: iso(-1), resolved: iso(-1) + 'T18:00:00Z', memberId: 'BS502', payerName: 'Blue Shield of California', status: 'error.member.id' },
    { requestId: 'v-504', id: '900004', created: iso(-1), resolved: iso(-1) + 'T18:00:00Z', memberId: 'C504', payerName: 'Cigna', status: 'complete.eligible', innCopay: '20' },
  ].map(o => pick(VER_HDR, o))] },
  { name: 'discovery output.csv', rows: [DISC_HDR, ...[
    { requestId: 'd-503', id: '900003', created: iso(-1), resolved: iso(-1) + 'T19:00:00Z' },
    { requestId: 'g-503', id: '900003-G', created: iso(-1), resolved: iso(-1) + 'T20:00:00Z', memberId: 'BSG503', responsePayerId: '94036', responsePayerName: 'Blue Shield of California' },
  ].map(o => pick(DISC_HDR, o))] },
];
const RULES = PR.read([PR.HEADERS, ['Cigna', '', 'NP', 'N', 'No', 'No', 'No', 'No', 'cigna', ''], ['Cigna', '', 'MD', 'N', 'N/A', 'Yes', 'Yes', 'Yes', 'cigna', '']]).rules;
const blankRow = (cols, o) => Object.assign(Object.fromEntries(cols.map(c => [c, ''])), o);
function adminSrc(){
  const built = C.buildPatients(IMPORT(), CFG);
  const merged = C.mergeSoharResults(IMPORT(), FILES(), CFG);
  const patients = C.applyReverifySkips(built.patients, merged.results, { ...C.REVERIFY_DEFAULTS, today: NOW }).patients.filter(p => p['In Scope'] === 'Yes');
  const fresh = C.buildTasks(merged.results);
  const fix = fresh.find(t => t.MRN === '900002' && t.Task === 'Fix member ID in NextGen');
  return { fixKey: fix['Task Key'], src: {
    patients, results: merged.results, review: built.review, tasks: C.keepDone(fresh, [{ 'Task Key': fix['Task Key'], Done: 'x' }]),
    learningRows: [blankRow(C.PAYER_LEARNING_COLS, { 'NextGen Payer': 'Cigna', Confidence: 'Exact', 'Sohar payerName': 'Cigna', 'Patients agreed': '4' }), blankRow(C.PAYER_LEARNING_COLS, { 'NextGen Payer': 'Blue Shield CA', Confidence: 'Auto-Medium', 'Payer errors': '2' })],
    metricRows: [blankRow(C.PAYER_METRIC_COLS, { 'NextGen Payer': 'Blue Shield CA', 'Failed verification patients': '1', 'Most common failure': 'error.member.id (1)', Recurring: 'Yes', 'Days with issues': '4' })],
    cfg: { locationState: LOCSTATE }, payerRuleRows: R.payerRuleRows(RULES), runAt: NOW, version: '0.14.0',
  } };
}
// the sheets as Excel holds them: header in row 1, values only (what XL.buildWorkbook writes)
const gridsOf = pack => Object.fromEntries(pack.sheets.map(s => [s.name, [s.columns, ...s.rows.map(r => s.columns.map(c => (r[c] == null ? '' : r[c])))]]));

// ---------- a mock Excel workbook that load() can read ----------
function mockBook(sheets){
  const reads = [];
  const ws = s => ({ name: s.name, visibility: s.hidden ? 'Hidden' : 'Visible',
    getUsedRangeOrNullObject(){ const h = s.rows.length, w = Math.max(0, ...s.rows.map(r => r.length)); return h && w ? { isNullObject: false, load(){}, rowIndex: 0, columnIndex: 0, rowCount: h, columnCount: w } : { isNullObject: true, load(){} }; },
    getRangeByIndexes: (r, c, h, w) => ({ load(){ reads.push([s.name, r, h, w]); }, get values(){ return Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => { const v = (s.rows[r + i] || [])[c + j]; return v == null ? '' : v; })); } }) });
  const ctx = { sync: async () => {}, workbook: { worksheets: { load(){}, items: sheets.map(ws) } } };
  return { Excel: { run: fn => fn(ctx) }, reads };
}
const bookOf = pack => pack.sheets.map(s => ({ name: s.name, hidden: !!s.hidden, rows: gridsOf(pack)[s.name] }));
function loadApp(Excel, extra){
  const els = {}, on = {};
  const el = () => ({ addEventListener(){}, appendChild(){}, remove(){}, focus(){}, querySelector: () => null, querySelectorAll: () => [], hidden: true, innerHTML: '', value: '', classList: { toggle(){} } });
  // timers never hold the test process open (toasts, the full screen relay's 2-minute timeout)
  const later = (f, ms) => { const h = setTimeout(f, ms); if (h.unref) h.unref(); return h; };
  const sandbox = { console, setTimeout: later, clearTimeout, Excel, Office: undefined, window: {}, location: { href: 'https://x/taskpane.html', search: '' },
    document: { querySelector: s => els[s] || (els[s] = el()), querySelectorAll: () => [], addEventListener(k, f){ (on[k] = on[k] || []).push(f); }, createElement: el, head: el() } };
  sandbox.TE = require('../engine.js'); sandbox.TC = require('../cadence.js'); sandbox.PayerRules = PR; sandbox.SoharCore = C; sandbox.SoharRules = R; sandbox.self = sandbox;
  Object.assign(sandbox, extra || {});
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8') + '\n;this.__api = { OfficeStore, RemoteStore, RELAY_OPS, S, prepare, reload, render, seWrite, shownViews, navGroups, currentView };', sandbox);
  // fire the page's own click / change handlers with a fake event (each returns its promise)
  const fire = (k, e) => Promise.all((on[k] || []).map(f => f(e)));
  return { api: sandbox.__api, els, doc: sandbox.document, fire };
}
// an app whose store is the mock workbook (load) plus a writeSheet log. A Clinic Work change goes onto the rows on the sheet
// (the last write), as OfficeStore does it; the log holds what was written and the rows come back.
async function appWith(sheets){
  const { Excel, reads } = mockBook(sheets); const { api, els, doc, fire } = loadApp(Excel);
  const writes = [], docs = {};
  const office = api.OfficeStore();
  let work = (sheets.find(s => s.name === 'Clinic Work') || {}).rows || [];
  const writeSheet = async (n, aoa, fmt) => {
    if (!(fmt && fmt.change)){ writes.push([n, aoa, fmt]); return; }
    const rows = R.workRows(work, fmt.change); work = [R.CLINIC_WORK_COLS].concat(rows.map(x => R.CLINIC_WORK_COLS.map(c => x[c])));
    writes.push([n, work, fmt]); return { table: 'ok', rows };
  };
  api.S.store = { live: true, load: office.load, writeSheet, getDoc: k => docs[k] ?? null, setDoc: async (k, v) => { docs[k] = v; } };
  await api.reload();
  return { api, els, doc, writes, docs, reads, fire };
}

test('pack -> workbook -> load() -> prepare -> derive gives the clinic list; no names or DOBs anywhere', async () => {
  const { src } = adminSrc();
  const pack = R.clinicPack(LOC, src);
  const { api, els, docs } = await appWith(bookOf(pack));
  const S = api.S;
  assert.ok(S.SE && S.E, 'the file unpacked and derived');
  assert.deepEqual(S.SE, R.clinicUnpack(gridsOf(pack)), 'load() + prepare() read the same tables as the file holds');
  assert.equal(S.SE.payerRules.length, RULES.length, 'payer rules from the hidden sheet');
  // the admin's own derivation for this clinic, now
  const cfg = { today: NOW, sfhpProviders: [], reverifyDays: 7, reverifyNewMonth: false, payerRules: R.payerRulesFrom(src.payerRuleRows), locationState: LOCSTATE };
  const flags = R.allFlags(src.patients, src.results, cfg);
  const admin = R.clinicRows(LOC, src.patients, src.results, flags, NOW);
  assert.deepEqual(S.E.rows.map(r => r.key), admin.map(r => r.key));
  assert.deepEqual(S.E.rows.map(r => r.key), ['900002', '900001', '900003', '900005']);
  assert.deepEqual(S.E.rows.map(r => r.LASFIR), admin.map(r => C.lasfir(r['Last Name'], r['First Name'])));
  const strip = r => Object.fromEntries(Object.entries(r).filter(([k]) => !['Last Name', 'First Name', 'DOB', 'LASFIR', 'Notes'].includes(k)));
  assert.deepEqual(S.E.rows.map(strip), admin.map(strip), 'same rows as the admin, minus names and DOB');
  assert.deepEqual(S.E.flags, new Map([...flags].filter(([k]) => S.E.rows.some(r => r.key === k))));
  const t = (mrn, task) => S.E.tasks.find(x => x.MRN === mrn && x.Task === task);
  assert.equal(t('900002', 'Fix member ID in NextGen').doneBy, 'admin');
  assert.equal(t('900001', 'Update benefits in NextGen').clinic, true);
  assert.deepEqual(S.E.benefits.rows.map(b => [b.MRN, b.LASFIR, b['OOP Status']]).filter(b => b[0] === '900001'), [['900001', 'QUITES', 'Not met']]);
  assert.equal(S.view, 'eligibility', 'a clinic file opens on Eligibility');
  assert.equal(docs['Office.AutoShowTaskpaneWithDocument'], true, 'first open saves the auto-show setting');
  // every view draws, LASFIR + MRN only
  for (const v of ['eligibility', 'benefits', 'payers', 'setup']){
    S.view = v; api.render(); const html = els['#app'].innerHTML;
    assert.ok(html.includes('QUITES') || v === 'payers' || v === 'setup', v);
    for (const s of NEVER) assert.ok(!html.includes(s), `"${s}" is not on the ${v} view`);
    if (v !== 'setup') assert.ok(html.includes('Clinic Eligibility</b> · San Rafael · data from'), 'header line');
    assert.ok(!html.includes('>Schedule<') && !html.includes('>Patients<'), 'template tools are hidden');
  }
  S.view = 'eligibility'; api.render(); const html = els['#app'].innerHTML;
  assert.ok(html.includes('MRN 900001') && html.includes('data-act="seDone"') && html.includes('data-act="seRv"'));
  assert.ok(/Insurance found|Choose/.test(html));
});

test('load() reads a clinic file\'s own sheets whole (hidden, every column, 500 rows a sync) and never as templates or the appointment list', async () => {
  const wide = Array.from({ length: 25 }, (_, i) => (i === 24 ? 'stray' : ''));
  const results = [['Key', 'Record'], ...Array.from({ length: 1200 }, (_, i) => [String(900100 + i), '{}'])]; results[3] = results[3].concat(wide.slice(2));
  const tpl = [[], ['Reyes MD, Samuel'], ['Effective From:', 46300], ['Location:', 'Eastside'], ['Provider:', 'Reyes', 'Reyes', 'Reyes', 'Reyes', 'Reyes'], ['Time', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'], [480 / 1440, 'New Patient', '', '', '', '']];
  const appts = [['Appt Date']].concat(Array.from({ length: 60 }, () => ['1/1/2026']));
  const sheets = [
    { name: 'Read Me', rows: tpl }, { name: 'Clinic Work', rows: tpl.concat([R.CLINIC_WORK_COLS]) },
    { name: 'SE Info', hidden: true, rows: [['Setting', 'Value'], ['Kind', 'Clinic Eligibility'], ['Location', LOC], ['Format', '1']] },
    { name: 'SE Results', hidden: true, rows: results }, { name: 'SE Done', hidden: true, rows: appts },
    { name: 'Sheet1', rows: tpl },
  ];
  const { Excel, reads } = mockBook(sheets); const { api } = loadApp(Excel);
  const raw = await api.OfficeStore().load();
  assert.deepEqual(Object.keys(raw.grids), ['Sheet1'], 'Read Me and Clinic Work are never template grids');
  assert.equal(raw.appt, null, 'SE sheets never match the appointment-list rule');
  assert.deepEqual(Object.keys(raw.se).sort(), ['Clinic Work', 'SE Done', 'SE Info', 'SE Results']);
  assert.deepEqual(reads.filter(x => x[0] === 'SE Results').map(x => [x[1], x[2], x[3]]), [[0, 500, 25], [500, 500, 25], [1000, 201, 25]], 'all 25 columns, 500 rows a read');
  assert.equal(raw.se['SE Results'].length, 1201);
  assert.equal(raw.se['SE Results'][3][24], 'stray');
  // the same sheets without SE Info: exactly as before (Read Me is a template tab, the big sheet the appointment list)
  const plain = sheets.filter(s => s.name !== 'SE Info').map(s => s.name === 'SE Done' ? Object.assign({}, s, { name: 'Appt export', hidden: false }) : s);
  const { Excel: E2 } = mockBook(plain); const raw2 = await loadApp(E2).api.OfficeStore().load();
  assert.equal(raw2.se, undefined);
  assert.deepEqual(Object.keys(raw2.grids).sort(), ['Clinic Work', 'Read Me', 'Sheet1']);
  assert.equal(raw2.appt.sheet, 'Appt export');
  // lookup tabs are read under their own keys (they were filed under 'ref,Ref' and never seen)
  const lookups = plain.concat([{ name: 'Ref', rows: [['Event', 'Class'], ['Consult', 'Visit']] }, { name: 'providers', rows: [['NextGen Provider', 'Label'], ['Reyes MD, Samuel', 'Reyes']] }]);
  const raw3 = await loadApp(mockBook(lookups).Excel).api.OfficeStore().load();
  assert.deepEqual(Object.keys(raw3.cfg).sort(), ['prov', 'ref']);
  assert.equal(raw3.cfg.prov[1][1], 'Reyes');
  assert.ok(!('Ref' in raw3.grids) && !('providers' in raw3.grids), 'lookup tabs are never template grids');
  // a wrong file shows its error on Setup and leaves the template tools as they are
  const bad = sheets.map(s => s.name === 'SE Info' ? Object.assign({}, s, { rows: [['Setting', 'Value'], ['Kind', 'Sohar Clinic'], ['Format', '1']] }) : s);
  const { api: a3, els } = await appWith(bad);
  assert.equal(a3.S.SE, null);
  assert.match(a3.S.data.seErr, /not a Clinic Eligibility file/);
  assert.ok(els['#app'].innerHTML.includes('not a Clinic Eligibility file'));
});

test('Clinic Work: each change upserts by Kind + Key and rewrites the whole sheet, one write at a time', async () => {
  const { src, fixKey } = adminSrc();
  const { api, writes } = await appWith(bookOf(R.clinicPack(LOC, src)));
  const S = api.S;
  const r1 = () => S.E.rows.find(r => r.key === '900001');
  assert.ok(r1()['Verification Skipped'], 'verified yesterday: skipped this week');
  await api.seWrite({ Kind: 'Re-verify anyway', Key: '900001', MRN: '900001', LASFIR: 'QUITES', Value: 'On' });
  assert.equal(r1()['Re-verify Anyway'], 'Yes', 'the list re-derives at once');
  const ben = S.E.tasks.find(t => t.MRN === '900001' && t.Task === 'Update benefits in NextGen')['Task Key'];
  // two quick changes: queued, and the last write holds both
  const p1 = api.seWrite({ Kind: 'Re-verify anyway', Key: '900001', MRN: '900001', LASFIR: 'QUITES', Value: 'Off' });
  const p2 = api.seWrite({ Kind: 'Task done', Key: ben, MRN: '900001', LASFIR: 'QUITES', Value: 'Yes' });
  await Promise.all([p1, p2]);
  await api.seWrite({ Kind: 'Note', Key: '900002', MRN: '900002', LASFIR: 'BRABRA', Note: 'New card at check-in' });
  await api.seWrite({ Kind: 'Payer note', Key: 'Cigna', Value: 'Cigna Open Access Plus', Note: 'Group 0042' });
  assert.equal(writes.length, 5);
  assert.ok(writes.every(w => w[0] === 'Clinic Work' && w[2].keepFormat && w[2].inPlace && w[2].text && w[2].table === 'tblClinicWork' && w[1][0].join() === R.CLINIC_WORK_COLS.join()), 'as text, over the old rows, inside the table');
  const last = JSON.parse(JSON.stringify(writes[4][1])).slice(1).map(v => Object.fromEntries(R.CLINIC_WORK_COLS.map((c, i) => [c, v[i]])));
  assert.deepEqual(last.map(r => [r.Kind, r.Key, r.Value]), [['Re-verify anyway', '900001', 'Off'], ['Task done', ben, 'Yes'], ['Note', '900002', ''], ['Payer note', 'Cigna', 'Cigna Open Access Plus']]);
  assert.ok(last.every(r => r.By === 'Clinic San Rafael' && /^\d{4}-\d\d-\d\dT/.test(r.When)));
  assert.equal(r1()['Re-verify Anyway'], '');
  assert.deepEqual([S.E.tasks.find(t => t['Task Key'] === ben).Done, S.E.tasks.find(t => t['Task Key'] === ben).doneBy], ['Yes', 'clinic']);
  assert.equal(S.E.tasks.find(t => t['Task Key'] === fixKey).doneBy, 'admin');
  assert.equal(S.E.notes.get('900002').Note, 'New card at check-in');
  assert.equal(S.E.payers.find(p => p.key === 'cigna').note.cardSays, 'Cigna Open Access Plus');
  // what the admin gets back from that sheet
  const m = R.mergeClinicWork(writes[4][1], { location: LOC, requests: [], tasks: src.tasks, sentDone: S.SE.done, notes: [], payerNotes: [] });
  assert.deepEqual([m.done.get(ben), m.notes.length, m.payerNotes.length], [`Yes (clinic ${LOC})`, 1, 1]);
  // reopening the saved file picks the edits up from Clinic Work
  const book = bookOf(R.clinicPack(LOC, src)).map(s => s.name === 'Clinic Work' ? Object.assign({}, s, { rows: writes[4][1] }) : s);
  const again = await appWith(book);
  assert.equal(again.api.S.E.notes.get('900002').Note, 'New card at check-in');
  assert.equal(again.api.S.data.work.length, 4);
});

test('nav: groups with sub-tabs; a clinic file shows only the Sohar views unless template tools are on', () => {
  const { api } = loadApp(mockBook([]).Excel);
  const st = o => Object.assign({ SE: null, R: null, opts: {}, view: 'overview' }, o);
  const plain = x => JSON.parse(JSON.stringify(x));   // arrays from the app's vm realm
  const groups = s => plain(api.navGroups(s).map(g => [g.k, g.views]));
  // no clinic file: the same views as before
  assert.deepEqual(plain(api.shownViews(st({ R: {} }))), ['overview', 'templates', 'access', 'cadence', 'funnel', 'recall', 'visits', 'impact', 'trends', 'setup']);
  assert.deepEqual(groups(st({ R: {} })), [['schedule', ['overview', 'templates', 'impact', 'trends', 'visits']], ['patients', ['cadence', 'funnel', 'recall']], ['coverage', ['access']], ['setup', ['setup']]]);
  assert.equal(api.currentView(st({})), 'setup', 'no data yet: Setup');
  assert.equal(api.currentView(st({ R: {}, view: 'recall' })), 'recall');
  // clinic file
  assert.deepEqual(groups(st({ SE: {} })), [['coverage', ['eligibility', 'benefits', 'payers']], ['setup', ['setup']]]);
  assert.equal(api.currentView(st({ SE: {}, view: 'overview' })), 'eligibility');
  assert.deepEqual(groups(st({ SE: {}, opts: { templates: 'On' } })).map(g => g[1].length), [5, 3, 4, 1]);
  assert.equal(api.currentView(st({ SE: {}, opts: { templates: 'On' }, view: 'overview' })), 'eligibility', 'template views need the appointment list');
  assert.equal(api.currentView(st({ SE: {}, R: {}, opts: { templates: 'On' }, view: 'overview' })), 'overview');
});

// ---------- a stateful mock sheet for the Clinic Work write: ops apply on sync (as Office.js does, no rollback), and a
// General cell parses what it is given the way Excel does ('0090001' -> 90001, '=...' is a formula and fails here) ----------
function liveSheet(rows, failSync){
  const g = rows.map(r => r.map(v => ({ v, f: 'General' })));
  const cell = (r, c) => { while (g.length <= r) g.push([]); while (g[r].length <= c) g[r].push({ v: '', f: 'General' }); return g[r][c]; };
  const parse = (v, f) => { if (f === '@' || typeof v !== 'string') return v; if (v.startsWith('=')) throw new Error('InvalidArgument: formula'); return /^-?\d+(\.\d+)?$/.test(v) ? +v : v; };
  let q = [], n = 0;
  const range = (r0, c0, h, w) => {
    const each = (v, fn) => { assert.ok(v.length === h && v.every(x => x.length === w), `array must be ${h}×${w}`); v.forEach((x, i) => x.forEach((y, j) => fn(cell(r0 + i, c0 + j), y))); };
    return { set numberFormat(v){ q.push(() => each(v, (c, y) => { c.f = y; })); }, set values(v){ q.push(() => each(v, (c, y) => { c.v = parse(y, c.f); })); },
      clear(a){ q.push(() => { for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) { const c = cell(r0 + i, c0 + j); c.v = ''; if (a !== 'Contents') c.f = 'General'; } }); } };
  };
  const ws = { name: 'Clinic Work', isNullObject: false, getRange: () => range(0, 0, g.length || 1, Math.max(1, ...g.map(r => r.length))), getRangeByIndexes: range,
    getUsedRangeOrNullObject(){ const u = { load(){} }; q.push(() => { const used = g.map((r, i) => [i, r.reduce((m, c, j) => (c.v !== '' ? j + 1 : m), 0)]).filter(x => x[1]);
      if (!used.length) u.isNullObject = true; else Object.assign(u, { isNullObject: false, rowIndex: 0, columnIndex: 0, rowCount: used[used.length - 1][0] + 1, columnCount: Math.max(...used.map(x => x[1])) }); }); return u; } };
  const ctx = { sync: async () => { const ops = q; q = []; n++; if (failSync && failSync(n)) throw new Error('RequestThrottled'); ops.forEach(op => op()); },
    workbook: { worksheets: { getItemOrNullObject: () => ws, add(){ throw new Error('no add'); } } } };
  const values = () => g.map(r => { const x = r.map(c => c.v); while (x.length && x[x.length - 1] === '') x.pop(); return x; }).filter(r => r.length);
  return { Excel: { run: fn => fn(ctx) }, values, formats: () => g.map(r => r.map(c => c.f)) };
}
const workAoa = rows => [R.CLINIC_WORK_COLS].concat(rows.map(x => R.CLINIC_WORK_COLS.map(c => (x[c] == null ? '' : x[c]))));

test('Clinic Work is written as text: leading-zero MRN keys, group numbers and notes starting with = come back as typed', async () => {
  const work = R.workRows([], { Kind: 'Re-verify anyway', Key: '0090001', MRN: '0090001', LASFIR: 'QUITES', Value: 'On', By: 'Clinic San Rafael' }, NOW);
  const all = R.workRows(R.workRows(work, { Kind: 'Note', Key: '0090001', MRN: '0090001', LASFIR: 'QUITES', Note: '=> new card at check-in', By: 'Clinic San Rafael' }, NOW),
    { Kind: 'Payer note', Key: 'Cigna', Value: '00123', Note: '1/2 of the card is cut off', By: 'Clinic San Rafael' }, NOW);
  // the old General write: the key turns into a number (and an = note can't be written at all)
  const old = liveSheet([R.CLINIC_WORK_COLS]);
  await loadApp(old.Excel).api.OfficeStore().writeSheet('Clinic Work', workAoa(work), { keepFormat: true });
  assert.equal(old.values()[1][1], 90001, 'the mock coerces like Excel');
  const sh = liveSheet([R.CLINIC_WORK_COLS]);
  await loadApp(sh.Excel).api.OfficeStore().writeSheet('Clinic Work', workAoa(all), { keepFormat: true, inPlace: true, text: true });
  const back = sh.values();
  assert.deepEqual(back, workAoa(all), 'every cell reads back as the same string');
  assert.ok(sh.formats().slice(0, back.length).every(r => r.slice(0, R.CLINIC_WORK_COLS.length).every(f => f === '@')));
  // so the clinic's tick and note still key on the patient after a reload
  const objs = back.slice(1).map(v => Object.fromEntries(R.CLINIC_WORK_COLS.map((c, i) => [c, v[i]])));
  assert.deepEqual(objs.filter(r => r.Key === '0090001').map(r => r.Kind), ['Re-verify anyway', 'Note']);
});

test('a failed Clinic Work write keeps the rows already on the sheet; a shorter write clears what is left below', async () => {
  const before = workAoa([{ Kind: 'Task done', Key: '900001|Update benefits in NextGen', MRN: '900001', Value: 'Yes', By: 'Clinic San Rafael', When: '2026-10-20T10:00:00.000Z' },
    { Kind: 'Note', Key: '900002', MRN: '900002', Note: 'New card at check-in', By: 'Clinic San Rafael', When: '2026-10-20T10:05:00.000Z' }]);
  const next = before.concat([['Re-verify anyway', '900001', '900001', 'QUITES', 'On', '', 'Clinic San Rafael', '2026-10-21T09:00:00.000Z']]);
  // the values sync fails (throttled, timed out, relay dropped): nothing was cleared first, so both earlier rows stay
  const sh = liveSheet(before.map(r => r.slice()), n => n === 3);   // sync 1 finds the sheet, 2 reads its size, 3 writes
  await assert.rejects(loadApp(sh.Excel).api.OfficeStore().writeSheet('Clinic Work', next, { keepFormat: true, inPlace: true, text: true }), /RequestThrottled/);
  assert.deepEqual(sh.values(), before);
  // the clear-then-write path loses them (why Clinic Work no longer uses it)
  const old = liveSheet(before.map(r => r.slice()), n => n === 3);   // 1 finds the sheet, 2 clears it, 3 writes
  await assert.rejects(loadApp(old.Excel).api.OfficeStore().writeSheet('Clinic Work', next, { keepFormat: true }), /RequestThrottled/);
  assert.deepEqual(old.values(), []);
  // a write shorter or narrower than what is there clears the rest, after the new rows are in
  const sh2 = liveSheet(next.map(r => r.concat(['stray'])));
  await loadApp(sh2.Excel).api.OfficeStore().writeSheet('Clinic Work', before, { keepFormat: true, inPlace: true, text: true });
  assert.deepEqual(sh2.values(), before);
});

test('ticking a box re-draws the list and puts keyboard focus back on the same control', async () => {
  const { src } = adminSrc();
  const { api, els, doc } = await appWith(bookOf(R.clinicPack(LOC, src)));
  const seen = []; let focused = 0;
  els['#app'].querySelector = sel => { seen.push(sel); return { focus(){ focused++; } }; };
  doc.activeElement = { dataset: { act: 'seRv', m: '900001' } };
  await api.seWrite({ Kind: 'Re-verify anyway', Key: '900001', MRN: '900001', LASFIR: 'QUITES', Value: 'On' });
  doc.activeElement = { dataset: { act: 'seDone', k: '900001|Say "hi" \\ there', m: '900001' } };
  await api.seWrite({ Kind: 'Task done', Key: '900001|Say "hi" \\ there', MRN: '900001', LASFIR: 'QUITES', Value: 'Yes' });
  assert.deepEqual(seen, ['[data-act="seRv"][data-m="900001"]', '[data-act="seDone"][data-k="900001|Say \\"hi\\" \\\\ there"][data-m="900001"]']);
  assert.equal(focused, 2);
  // nothing focused (a mouse click on the page): no selector, no focus call
  doc.activeElement = undefined;
  await api.seWrite({ Kind: 'Re-verify anyway', Key: '900001', MRN: '900001', LASFIR: 'QUITES', Value: 'Off' });
  assert.equal(seen.length, 2);
});

// ---------- Clinic Work inside tblClinicWork: a stateful sheet with a table that keeps Excel's bounds rules. Ops apply per sync,
// in order, with no rollback. resize (ExcelApi 1.13) must keep the header row and overlap the table; rows.add inserts at the
// table's end and shifts the cells below it down (Microsoft Learn: "Any rows below the inserted row are shifted downwards").
// The admin's flow ("List rows present in a table") sees only inTable(). o: {noResize, failResize, failAdd, failSync(n)} ----------
function tableSheet(rows, tbl, o = {}){
  const g = rows.map(r => r.map(v => ({ v, f: '@' })));   // what an earlier text write left
  const cell = (r, c) => { while (g.length <= r) g.push([]); while (g[r].length <= c) g[r].push({ v: '', f: 'General' }); return g[r][c]; };
  const parse = (v, f) => { if (f === '@' || typeof v !== 'string') return v; if (v.startsWith('=')) throw new Error('InvalidArgument: formula'); return /^-?\d+(\.\d+)?$/.test(v) ? +v : v; };
  const t = tbl ? Object.assign({ name: 'tblClinicWork', r0: 0, c0: 0, cols: R.CLINIC_WORK_COLS.length }, tbl) : null;
  let q = [], n = 0;
  const range = (r0, c0, h, w) => {
    const each = (v, fn) => { assert.ok(v.length === h && v.every(x => x.length === w), `array must be ${h}×${w}`); v.forEach((x, i) => x.forEach((y, j) => fn(cell(r0 + i, c0 + j), y))); };
    // a read: load, then the values as they are once that sync has run
    let got = null;
    return { _b: { r0, c0, h, w }, load(){ q.push(() => { got = Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => (g[r0 + i] && g[r0 + i][c0 + j] ? g[r0 + i][c0 + j].v : ''))); }); },
      get values(){ assert.ok(got, 'values read before a sync'); return got; },
      set numberFormat(v){ q.push(() => each(v, (c, y) => { c.f = y; })); }, set values(v){ q.push(() => each(v, (c, y) => { c.v = parse(y, c.f); })); },
      clear(a){ q.push(() => { for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) { const c = cell(r0 + i, c0 + j); c.v = ''; if (a !== 'Contents') c.f = 'General'; } }); } };
  };
  const used = () => { const u = g.map((r, i) => [i, r.reduce((m, c, j) => (c.v !== '' ? j + 1 : m), 0)]).filter(x => x[1]);
    return u.length ? { isNullObject: false, rowIndex: 0, columnIndex: 0, rowCount: u[u.length - 1][0] + 1, columnCount: Math.max(...u.map(x => x[1])) } : { isNullObject: true }; };
  const table = t && { isNullObject: false,
    getRange(){ const r = { load(){} }; q.push(() => Object.assign(r, { rowIndex: t.r0, columnIndex: t.c0, rowCount: t.rows, columnCount: t.cols })); return r; },
    resize: o.noResize ? undefined : rg => { q.push(() => { if (o.failResize) throw new Error('ApiNotFound'); const b = rg._b;
      assert.equal(b.r0, t.r0, 'the header stays in its row'); assert.ok(b.c0 < t.c0 + t.cols && b.c0 + b.w > t.c0, 'the new range overlaps the table'); Object.assign(t, { rows: b.h, cols: b.w, c0: b.c0 }); }); },
    rows: { add: (i, values) => { q.push(() => {
      if (o.failAdd) throw new Error('InvalidOperation');
      assert.ok(i == null && values.every(r => r.length === t.cols), 'rows at the end, as wide as the table');
      const k = values.length, at = t.r0 + t.rows, last = g.length - 1;
      for (let j = t.c0; j < t.c0 + t.cols; j++){
        for (let r = last; r >= at; r--){ const s = cell(r, j), d = cell(r + k, j); d.v = s.v; d.f = s.f; }
        for (let x = 0; x < k; x++){ const c = cell(at + x, j); c.f = 'General'; c.v = parse(values[x][j - t.c0], 'General'); }
      }
      t.rows += k; }); } } };
  const ws = { name: 'Clinic Work', isNullObject: false, getRangeByIndexes: range,
    getUsedRangeOrNullObject(){ const u = { load(){} }; q.push(() => Object.assign(u, used())); return u; },
    tables: { getItemOrNullObject: nm => (t && nm === t.name ? table : { isNullObject: true }) } };
  const ctx = { sync: async () => { const ops = q; q = []; n++; if (o.failSync && o.failSync(n)) throw new Error('RequestThrottled'); ops.forEach(op => op()); },
    workbook: { worksheets: { getItemOrNullObject: () => ws, add(){ throw new Error('no add'); } } } };
  const trim = r => { const x = r.slice(); while (x.length && x[x.length - 1] === '') x.pop(); return x; };
  const values = () => g.map(r => trim(r.map(c => c.v))).filter(r => r.length);
  const inTable = () => Array.from({ length: t.rows }, (_, i) => Array.from({ length: t.cols }, (_, j) => (g[t.r0 + i] && g[t.r0 + i][t.c0 + j] ? g[t.r0 + i][t.c0 + j].v : '')));
  return { Excel: { run: fn => fn(ctx) }, values, inTable, table: () => Object.assign({}, t), formats: () => g.map(r => r.map(c => c.f)), syncs: () => n };
}
const FMT = { keepFormat: true, inPlace: true, text: true, table: 'tblClinicWork' };
const wr = (kind, key, extra) => Object.assign({ Kind: kind, Key: key, MRN: String(key).split('|')[0], LASFIR: 'QUITES', By: 'Clinic San Rafael' }, extra);
const CHANGES = [wr('Task done', '0090001|Update benefits in NextGen', { Value: 'Yes', Note: '=> called the patient, new card scanned' }), wr('Note', '0090001', { Note: 'New card at check-in' }),
  wr('Re-verify anyway', '900002', { Value: 'On' }), wr('Payer note', 'Cigna', { Value: '00123', Note: 'Group 0042' })];
const plainRows = rows => JSON.parse(JSON.stringify(rows));

test('every Clinic Work write leaves all its rows inside tblClinicWork, as text, and the table never shrinks', async () => {
  // as Sohar exports it: header + one blank row (A1:H2)
  const sh = tableSheet([R.CLINIC_WORK_COLS], { rows: 2 });
  const store = loadApp(sh.Excel).api.OfficeStore();
  let work = [];
  for (const ch of CHANGES){
    work = R.workRows(work, ch, NOW);
    const res = await store.writeSheet('Clinic Work', workAoa(work), FMT);
    assert.equal(res.table, 'ok');
    assert.deepEqual(sh.inTable(), workAoa(work).length < 2 ? workAoa(work).concat([R.CLINIC_WORK_COLS.map(() => '')]) : workAoa(work), `after ${work.length} rows the flow sees them all`);
  }
  assert.deepEqual(sh.values(), workAoa(work), 'nothing outside the table');
  assert.equal(sh.inTable()[1][2], '0090001', 'leading zeros kept');
  assert.equal(sh.inTable()[1][5], '=> called the patient, new card scanned', 'a note starting with = is text');
  // a table already bigger than the rows stays that size
  const big = tableSheet([R.CLINIC_WORK_COLS], { rows: 9 });
  assert.equal((await loadApp(big.Excel).api.OfficeStore().writeSheet('Clinic Work', workAoa(work.slice(0, 2)), FMT)).table, 'ok');
  assert.equal(big.table().rows, 9);
  // other sheets never get a table result
  const plain = tableSheet([['A']], null);
  assert.equal(await loadApp(plain.Excel).api.OfficeStore().writeSheet('Clinic Work', [['A'], ['1']], { keepFormat: true, inPlace: true, text: true }), undefined);
});

test('rows an older write left below the table are pulled in by the next write (resize, or rows.add where resize is missing)', async () => {
  const old = R.workRows([], CHANGES.slice(0, 3), NOW);   // 3 rows written by the old code: the table still ends at row 2
  for (const o of [{}, { noResize: true }, { failResize: true }]){
    const sh = tableSheet(workAoa(old), { rows: 2 }, o);
    assert.equal(sh.inTable().length - 1, 1, 'the flow saw one row of three');
    // what the pane does: read the sheet (all rows), apply the change, write the whole sheet
    const next = R.workRows(R.tableObjects(sh.values()), CHANGES[3], NOW);
    const res = await loadApp(sh.Excel).api.OfficeStore().writeSheet('Clinic Work', workAoa(next), FMT);
    const how = JSON.stringify(o);
    assert.equal(res.table, 'ok', how);
    assert.deepEqual(sh.inTable(), workAoa(next), `${how}: all four rows inside the table`);
    assert.deepEqual(sh.values(), workAoa(next), `${how}: no copies left below it (and '00123' is still text)`);
    assert.ok(sh.formats().slice(0, 5).every(r => r.slice(0, 8).every(f => f === '@')), `${how}: added rows are text`);
  }
});

test('a Clinic Work table write is fail-safe: the rows written stay whatever fails, and the result says the table is short', async () => {
  const before = workAoa(R.workRows([], CHANGES.slice(0, 2), NOW));
  const next = workAoa(R.workRows(R.tableObjects(before), CHANGES.slice(2), NOW));
  const write = async (sh, aoa) => loadApp(sh.Excel).api.OfficeStore().writeSheet('Clinic Work', aoa, FMT);
  // resize and rows.add both fail: the rows are on the sheet (the clinic loses nothing), the table is as it was
  const a = tableSheet(before, { rows: 3 }, { failResize: true, failAdd: true });
  assert.equal((await write(a, next)).table, 'short');
  assert.deepEqual(a.values(), next);
  assert.equal(a.table().rows, 3);
  // the values sync fails: nothing changes, and the error reaches seWrite (sync 1 finds the sheet, 2 reads its size, 3 writes)
  const b = tableSheet(before, { rows: 3 }, { failSync: n => n === 3 });
  await assert.rejects(write(b, next), /RequestThrottled/);
  assert.deepEqual([b.values(), b.table().rows], [before, 3]);
  // the table lookup itself fails (sync 5): still a saved write, reported short
  const c = tableSheet(before, { rows: 3 }, { failSync: n => n === 5 });
  assert.equal((await write(c, next)).table, 'short');
  assert.deepEqual(c.values(), next);
  // the clean-up after rows.add fails (no resize: 7 adds, 8 reads the bounds, 9 clears): the table has every row, copies sit below
  const d = tableSheet(before, { rows: 3 }, { noResize: true, failSync: n => n === 9 });
  assert.equal((await write(d, next)).table, 'ok');
  assert.deepEqual(d.inTable(), next);
  assert.ok(d.values().length > next.length, 'the shifted copies are still below the table');
  const objs = R.tableObjects(d.values()).filter(r => r.Kind);
  assert.equal(new Set(objs.map(r => r.Kind + '|' + r.Key)).size, 4, 'the copies below repeat rows already in the table');
  // no table on the sheet: written, reported missing
  const e = tableSheet(before, null);
  assert.equal((await write(e, next)).table, 'missing');
  assert.deepEqual(e.values(), next);
});

test('seWrite asks for the table, and says once when the rows could not go inside it', async () => {
  const { src } = adminSrc();
  const { api, els } = await appWith(bookOf(R.clinicPack(LOC, src)));
  const toasts = []; els['#toast'] = { appendChild: d => toasts.push([d.className, d.innerHTML]) };
  const fmts = []; let res = { table: 'ok' };
  api.S.store.writeSheet = async (n, aoa, fmt) => { fmts.push(fmt); return res; };
  await api.seWrite(wr('Re-verify anyway', '900001', { Value: 'On' }));
  assert.equal(toasts.length, 0);
  res = { table: 'short' };
  await api.seWrite(wr('Re-verify anyway', '900001', { Value: 'Off' }));
  await api.seWrite(wr('Re-verify anyway', '900001', { Value: 'On' }));
  assert.equal(toasts.length, 1, 'once a session');
  assert.equal(toasts[0][0], 'msg warn');
  assert.match(toasts[0][1], /Saved on the Clinic Work tab\. The admin’s inbox reads only the rows inside the Clinic Work table.*send the saved file to the admin/);
  assert.ok(fmts.every(f => f.table === 'tblClinicWork'));
  assert.equal(api.S.E.rows.find(r => r.key === '900001')['Re-verify Anyway'], 'Yes', 'the change stands');
});

test('full screen: the table option goes through the relay and the result comes back', async () => {
  const sent = []; let rx = null;
  const Office = { onReady: cb => cb({ host: 'None' }), HostType: { Excel: 'Excel' }, EventType: { DialogParentMessageReceived: 'parent' },
    context: { ui: { addHandlerAsync: (k, h) => { rx = h; }, messageParent: s => sent.push(s) } } };
  const { api } = loadApp(mockBook([]).Excel, { Office });
  assert.ok(api.RELAY_OPS.includes('writeSheet'));
  const rs = api.RemoteStore();
  const aoa = workAoa(R.workRows([], CHANGES[0], NOW));
  const p = rs.writeSheet('Clinic Work', aoa, FMT);
  const parts = sent.map(s => JSON.parse(s)).sort((x, y) => x.i - y.i);
  const msg = JSON.parse(parts.map(x => x.d).join(''));
  assert.equal(msg.op, 'writeSheet');
  assert.deepEqual(msg.args, [ 'Clinic Work', plainRows(aoa), FMT ]);
  rx({ message: JSON.stringify({ c: 'r1', i: 0, n: 1, d: JSON.stringify({ rid: msg.rid, ok: true, res: { table: 'short' } }) }) });
  assert.equal((await p).table, 'short');
  // the Clinic Work change is plain JSON: it crosses the relay unchanged, and the rows on the sheet come back
  const fmt = Object.assign({ change: Object.assign({ When: NOW.toISOString() }, CHANGES[0]) }, FMT);
  assert.deepEqual(plainRows(fmt), fmt, 'nothing lost to JSON');
  sent.length = 0;
  const p2 = rs.writeSheet('Clinic Work', null, fmt);
  const msg2 = JSON.parse(sent.map(s => JSON.parse(s)).sort((x, y) => x.i - y.i).map(x => x.d).join(''));
  assert.deepEqual(msg2.args, ['Clinic Work', null, fmt]);
  const rows = R.workRows([], CHANGES[0], NOW);
  rx({ message: JSON.stringify({ c: 'r2', i: 0, n: 1, d: JSON.stringify({ rid: msg2.rid, ok: true, res: { table: 'ok', rows } }) }) });
  assert.deepEqual(plainRows((await p2).rows), plainRows(rows));
});

// two panes (two staff) on one clinic file: each pane reads the file once, then keeps working from its own copy
async function paneOn(sh, pack){
  const a = await appWith(bookOf(pack));
  a.api.S.store.writeSheet = loadApp(sh.Excel).api.OfficeStore().writeSheet;
  return a.api;
}
const cw = rows => R.tableObjects(rows).map(r => [r.Kind, r.Key, r.Value, r.Note, r.By]);

test('two people in one clinic file: each change goes onto the rows on the sheet now, so neither wipes out the other', async () => {
  const { src } = adminSrc();
  const pack = R.clinicPack(LOC, src);
  const sh = tableSheet([R.CLINIC_WORK_COLS], { rows: 2 });   // as Sohar exports it
  const A = await paneOn(sh, pack), B = await paneOn(sh, pack);
  const ben = A.S.E.tasks.find(t => t.MRN === '900001' && t.Task === 'Update benefits in NextGen')['Task Key'];
  const rv = (by, v) => wr('Re-verify anyway', '900001', { Value: v, By: by });
  // Ana first: nothing new on the sheet, so her list is left as it was drawn
  const pa = A.seWrite(rv('Ana', 'On')); const mineA = A.S.data.work; await pa;
  assert.equal(A.S.data.work, mineA, 'no unseen rows: no second re-draw');
  // Ben's pane never saw Ana's tick; his note goes onto the sheet beside it
  const pb = B.seWrite(wr('Note', '900002', { Note: 'New card at check-in', By: 'Ben' })); const mineB = B.S.data.work;
  assert.equal(mineB.length, 1, 'Ben’s own copy has only his note');
  await pb;
  assert.deepEqual(cw(sh.values()), [['Re-verify anyway', '900001', 'On', '', 'Ana'], ['Note', '900002', '', 'New card at check-in', 'Ben']], 'Ana’s tick survived Ben’s save');
  assert.notEqual(B.S.data.work, mineB);
  assert.deepEqual(cw(B.S.data.work), cw(sh.values()), 'Ben now sees Ana’s row');
  assert.equal(B.S.E.rows.find(r => r.key === '900001')['Re-verify Anyway'], 'Yes', 'and his list shows it');
  // interleaved: Ana ticks a task; Ben (whose copy has no tick) adds a done note to it; the tick stays, from the sheet
  await A.seWrite(wr('Task done', ben, { Value: 'Yes', By: 'Ana' }));
  await B.seWrite(wr('Task done', ben, { Note: 'Called the patient', By: 'Ben' }));
  const aTask = () => A.S.E.tasks.find(t => t['Task Key'] === ben);
  assert.deepEqual([aTask().Done, aTask().clinicNote], ['Yes', ''], 'Ana’s pane hasn’t seen Ben’s note yet');
  // same Kind + Key: the last writer wins, and there is still one row for it
  await A.seWrite(rv('Ana', 'Off'));
  assert.deepEqual([aTask().Done, aTask().clinicNote], ['Yes', 'Called the patient'], 'her next save brings in Ben’s rows');
  await B.seWrite(rv('Ben', 'On'));
  const want = [['Note', '900002', '', 'New card at check-in', 'Ben'], ['Task done', ben, 'Yes', 'Called the patient', 'Ben'], ['Re-verify anyway', '900001', 'On', '', 'Ben']];
  assert.deepEqual(cw(sh.values()), want);
  assert.deepEqual(sh.inTable(), sh.values(), 'every row inside tblClinicWork for the admin’s flow');
  assert.ok(sh.formats().slice(0, 4).every(r => r.slice(0, 8).every(f => f === '@')), 'as text');
  assert.deepEqual(cw(B.S.data.work), want);
  assert.equal(A.S.E.rows.find(r => r.key === '900001')['Re-verify Anyway'], '', 'Ana still shows her own Off');
  await A.seWrite(wr('Payer note', 'Cigna', { Value: 'Cigna Open Access Plus', By: 'Ana' }));
  assert.equal(A.S.E.rows.find(r => r.key === '900001')['Re-verify Anyway'], 'Yes', 'until her next save brings in Ben’s On');
  // what the admin gets back holds both people's changes
  const m = R.mergeClinicWork(sh.inTable(), { location: LOC, requests: [], tasks: src.tasks, sentDone: B.S.SE.done, notes: [], payerNotes: [] });
  assert.deepEqual([m.done.get(ben), m.doneNotes.get(ben).split(' (')[0], m.notes.length, m.payerNotes.length], [`Yes (clinic ${LOC})`, 'Called the patient', 1, 1]);
});

test('two people, same Kind + Key, the later save carrying the earlier When: the sheet and the admin pick the same winner', async () => {
  const { src } = adminSrc();
  const pack = R.clinicPack(LOC, src);
  const sh = tableSheet([R.CLINIC_WORK_COLS], { rows: 2 });
  const A = await paneOn(sh, pack), B = await paneOn(sh, pack);
  const rv = (by, v, when) => wr('Re-verify anyway', '900001', Object.assign({ Value: v, By: by }, when ? { When: when } : {}));
  const row = () => R.tableObjects(sh.values()).find(r => r.Kind === 'Re-verify anyway');
  // Ben's PC clock runs 5 s ahead; his On lands first, and the admin's flow reads the file before Ana saves
  const benWhen = new Date(Date.now() + 5000).toISOString();
  await B.seWrite(rv('Ben', 'On', benWhen));
  const inbox = R.tableObjects(sh.inTable()).map(r => Object.assign({ Location: LOC, File: `Clinic Eligibility - ${LOC}.xlsx`, Received: new Date().toISOString() }, r));
  const merge = (rows, prev) => R.mergeClinicWork(R.inboxWork(rows, new Date()).get(LOC), Object.assign({ location: LOC, requests: [], tasks: src.tasks, sentDone: B.S.SE.done, notes: [], payerNotes: [], now: new Date() }, prev || {}));
  const asked = m => m.requests.some(r => r.MRN === '900001');   // a Re-verify Requests row
  const m1 = merge(inbox);
  assert.equal(asked(m1), true);
  // Ana's Off is stamped now, before Ben's When, but it is the last write: the sheet keeps it, restamped just after Ben's
  const pa = A.seWrite(rv('Ana', 'Off')); const mine = A.S.data.work;
  const anaWhen = mine.find(r => r.Kind === 'Re-verify anyway').When;
  assert.ok(anaWhen < benWhen, 'her pane stamped the earlier When');
  await pa;
  assert.deepEqual([row().Value, row().By], ['Off', 'Ana'], 'last writer wins on the sheet');
  assert.ok(Date.parse(row().When) > Date.parse(benWhen), 'and its When is now after Ben’s');
  assert.notEqual(A.S.data.work, mine, 'the restamped row came back to her pane');
  assert.deepEqual(A.S.data.work.find(r => r.Kind === 'Re-verify anyway'), row());
  // the flow appends Ana's row too (a new Kind|Key|When); the admin applies it after Ben's, so it agrees with the sheet
  inbox.push(Object.assign({ Location: LOC, File: `Clinic Eligibility - ${LOC}.xlsx`, Received: new Date().toISOString() }, row()));
  const m2 = merge(inbox, { requests: m1.requests, applied: m1.applied });
  assert.deepEqual([asked(m2), m2.counts.reverifyRemoved], [false, 1], 'Ana’s Off is the newest When, so it applies');
  assert.equal(asked(merge(inbox)), false, 'a fresh admin reading both rows gets Off too');
  // Ben's pane: his next save brings in Ana's Off
  await B.seWrite(wr('Note', '900002', { Note: 'New card at check-in', By: 'Ben' }));
  assert.equal(B.S.E.rows.find(r => r.key === '900001')['Re-verify Anyway'], '');
});

test('a Clinic Work change reads the sheet by header name, and a failed write still leaves the earlier rows', async () => {
  // columns in another order: each row is read by its header and written back in CLINIC_WORK_COLS order
  const odd = R.CLINIC_WORK_COLS.slice().reverse();
  const note = { When: '2026-10-20T10:05:00.000Z', By: 'Ben', Note: 'New card at check-in', Value: '', LASFIR: 'BRABRA', MRN: '900002', Key: '900002', Kind: 'Note' };
  const sh = tableSheet([odd, odd.map(c => note[c])], { rows: 2 });
  const fmt = Object.assign({ change: Object.assign({ When: NOW.toISOString() }, CHANGES[2]) }, FMT);
  const res = await loadApp(sh.Excel).api.OfficeStore().writeSheet('Clinic Work', null, fmt);
  assert.equal(res.table, 'ok');
  assert.deepEqual(cw(res.rows), [['Note', '900002', '', 'New card at check-in', 'Ben'], ['Re-verify anyway', '900002', 'On', '', 'Clinic San Rafael']]);
  assert.deepEqual(sh.inTable(), workAoa(res.rows));
  assert.equal(res.rows[1].When, NOW.toISOString(), 'the change’s own When, so the pane’s copy and the sheet match');
  // sync 1 finds the sheet, 2 its size, 3 reads the rows, 4 writes: a failed write changes nothing
  const before = workAoa(R.workRows([], CHANGES.slice(0, 2), NOW));
  const b = tableSheet(before, { rows: 3 }, { failSync: n => n === 4 });
  await assert.rejects(loadApp(b.Excel).api.OfficeStore().writeSheet('Clinic Work', null, fmt), /RequestThrottled/);
  assert.deepEqual([b.values(), b.table().rows], [before, 3]);
});

test('done notes back from the admin: each task shows the note the admin has, in plain words, and the clinic can still add its own', async () => {
  const { src, fixKey } = adminSrc();
  const ben = src.tasks.find(t => t.MRN === '900001' && t.Task === 'Update benefits in NextGen')['Task Key'];
  // a bring-back wrote one Clinic Note (it quotes the patient's name); the admin typed the other
  const notes = { [ben]: 'called Testa Quillfeather, new card scanned (Front desk, San Rafael, 10/21)', [fixKey]: 'billing has the new card' };
  const tasks = src.tasks.map(t => (notes[t['Task Key']] ? Object.assign({}, t, { 'Clinic Note': notes[t['Task Key']] }) : t));
  const pack = R.clinicPack(LOC, Object.assign({}, src, { tasks }));
  const { api, els, writes, fire } = await appWith(bookOf(pack));
  const S = api.S, task = k => S.E.tasks.find(t => t['Task Key'] === k);
  const sent = 'called QUITES, new card scanned (Front desk, San Rafael, 10/21)';
  assert.deepEqual([task(ben).sentNote, task(ben).clinicNote, task(ben).Done], [sent, '', ''], 'a note alone doesn’t tick the task');
  S.view = 'eligibility'; api.render(); let html = els['#app'].innerHTML;
  assert.ok(html.includes(`The admin has this note:</span> ${sent}`));
  assert.ok(html.includes('The admin has this note:</span> billing has the new card'), 'on a task the admin did too');
  for (const s of NEVER) assert.ok(!html.includes(s), `"${s}" is not on the list`);
  assert.ok(html.includes(`data-k="${ben}" data-m="900001" aria-label="Add a done note for Update benefits in NextGen">Note</button>`), 'the clinic can still add a note');
  // the dialog shows the admin's note and starts empty
  await fire('click', { target: { closest: s => (s === '[data-act]' ? { tagName: 'BUTTON', dataset: { act: 'seTaskNote', k: ben } } : null) } });
  const dlg = els['#sheet'].innerHTML;
  assert.ok(dlg.includes(`The admin has this note:</span> ${sent}`) && dlg.includes('A new note takes the place of the admin’s note above'));
  assert.match(dlg, /<textarea[^>]*><\/textarea>/);
  for (const s of NEVER) assert.ok(!dlg.includes(s), `"${s}" is not in the dialog`);
  els['#seNoteText'] = { value: 'card photo emailed too' };
  await fire('click', { target: { closest: s => (s === '[data-act]' ? { tagName: 'BUTTON', dataset: { act: 'seTaskNoteSave', k: ben } } : null) } });
  assert.deepEqual([task(ben).sentNote, task(ben).clinicNote], [sent, 'card photo emailed too'], 'side by side; neither replaces the other');
  const row = R.tableObjects(writes[writes.length - 1][1]).find(r => r.Kind === 'Task done' && r.Key === ben);
  assert.deepEqual([row.Value, row.Note], ['', 'card photo emailed too']);
  html = els['#app'].innerHTML;
  assert.ok(html.indexOf('The admin has this note:') < html.indexOf('Your note:</span> card photo emailed too'));
  // a file made before SE Done had the Clinic Note column: no admin note line, Done as before
  const old = bookOf(pack).map(s => (s.name === 'SE Done' ? Object.assign({}, s, { rows: s.rows.map(r => r.slice(0, 2)) }) : s));
  assert.deepEqual(old.find(s => s.name === 'SE Done').rows[0], ['Task Key', 'Done']);
  const o = await appWith(old);
  o.api.S.view = 'eligibility'; o.api.render();
  assert.ok(!o.els['#app'].innerHTML.includes('The admin has this note'));
  assert.equal(o.api.S.E.tasks.find(t => t['Task Key'] === fixKey).doneBy, 'admin');
});

test('done notes: a Note button by each clinic task; a note keeps the tick, a tick keeps the note, and the admin gets it from the file or the inbox', async () => {
  const { src } = adminSrc();
  const { api, els, writes, fire } = await appWith(bookOf(R.clinicPack(LOC, src)));
  const S = api.S;
  const task = mrn => S.E.tasks.find(t => t.MRN === mrn && t.Task === 'Update benefits in NextGen');
  const key = task('900001')['Task Key'];
  S.view = 'eligibility'; api.render();
  let html = els['#app'].innerHTML;
  const count = s => html.split(s).length - 1;
  assert.ok(count('data-act="seDone"') > 0);
  assert.equal(count('data-act="seTaskNote"'), count('data-act="seDone"'), 'one note button per task tick');
  // open the dialog: LASFIR + MRN only
  await fire('click', { target: { closest: s => (s === '[data-act]' ? { tagName: 'BUTTON', dataset: { act: 'seTaskNote', k: key } } : null) } });
  const dlg = els['#sheet'].innerHTML;
  assert.ok(dlg.includes('Done note · QUITES') && dlg.includes('MRN 900001') && dlg.includes('data-act="seTaskNoteSave"'));
  for (const s of NEVER) assert.ok(!dlg.includes(s), `"${s}" is not in the dialog`);
  const save = async note => { els['#seNoteText'] = { value: note }; await fire('click', { target: { closest: s => (s === '[data-act]' ? { tagName: 'BUTTON', dataset: { act: 'seTaskNoteSave', k: key } } : null) } }); };
  const tick = async on => fire('change', { target: { dataset: { act: 'seDone', k: key, m: '900001' }, checked: on } });
  const lastRow = () => R.tableObjects(writes[writes.length - 1][1]).find(r => r.Kind === 'Task done' && r.Key === key);
  // a note with the box unticked: a note-only row, Done stays as the file has it
  await save('  Called the patient, new card scanned ');
  assert.deepEqual([lastRow().Value, lastRow().Note], ['', 'Called the patient, new card scanned']);
  assert.deepEqual([task('900001').Done, task('900001').clinicNote], ['', 'Called the patient, new card scanned']);
  html = els['#app'].innerHTML;
  assert.ok(html.includes('Your note:</span> Called the patient, new card scanned') && html.includes('>Edit note<'));
  const m0 = R.mergeClinicWork(writes[writes.length - 1][1], { location: LOC, requests: [], tasks: src.tasks, sentDone: S.SE.done, notes: [], payerNotes: [] });
  assert.equal(m0.done.has(key), false, 'note only: the admin’s Done is left alone');
  assert.match(m0.doneNotes.get(key), /^Called the patient, new card scanned \(Clinic San Rafael, San Rafael, \d{1,2}\/\d{1,2}\)$/);
  // tick it: the note stays; edit the note: the tick stays; clear it: the tick stays
  await tick(true);
  assert.deepEqual([lastRow().Value, lastRow().Note], ['Yes', 'Called the patient, new card scanned']);
  await save('Called the patient twice');
  assert.deepEqual([lastRow().Value, lastRow().Note, task('900001').Done], ['Yes', 'Called the patient twice', 'Yes']);
  const m1 = R.mergeClinicWork(writes[writes.length - 1][1], { location: LOC, requests: [], tasks: src.tasks, sentDone: S.SE.done, notes: [], payerNotes: [] });
  assert.deepEqual([m1.done.get(key), m1.doneNotes.get(key).split(' (')[0]], [`Yes (clinic ${LOC})`, 'Called the patient twice']);
  // the same rows through the central inbox (what the flow appends) apply the same way
  const inbox = R.tableObjects(writes[writes.length - 1][1]).map(r => Object.assign({ Location: LOC, File: `Clinic Eligibility - ${LOC} 2026-10-23.xlsx`, Received: NOW.toISOString() }, r));
  const byLoc = R.inboxWork(inbox, NOW);
  const m2 = R.mergeClinicWork(byLoc.get(LOC), { location: LOC, requests: [], tasks: src.tasks, sentDone: S.SE.done, notes: [], payerNotes: [], now: NOW });
  assert.deepEqual([m2.done.get(key), m2.doneNotes.get(key)], [m1.done.get(key), m1.doneNotes.get(key)]);
  await save('');
  assert.deepEqual([lastRow().Value, lastRow().Note, task('900001').clinicNote], ['Yes', '', '']);
  assert.ok(!els['#app'].innerHTML.includes('Your note:</span>'));
  const m3 = R.mergeClinicWork(writes[writes.length - 1][1], { location: LOC, requests: [], tasks: src.tasks, sentDone: S.SE.done, notes: [], payerNotes: [] });
  assert.equal(m3.doneNotes.has(key), false, 'a cleared note never erases the admin’s');
  // reopening the saved file shows the note again
  await save('Called the patient twice');
  const again = await appWith(bookOf(R.clinicPack(LOC, src)).map(s => (s.name === 'Clinic Work' ? Object.assign({}, s, { rows: writes[writes.length - 1][1] }) : s)));
  const t2 = again.api.S.E.tasks.find(t => t['Task Key'] === key);
  assert.deepEqual([t2.Done, t2.clinicNote], ['Yes', 'Called the patient twice']);
});

test('a dialog takes keyboard focus, keeps Tab inside, and gives focus back to the button that opened it', async () => {
  const { src } = adminSrc();
  const { api, els, doc, fire } = await appWith(bookOf(R.clinicPack(LOC, src)));
  const S = api.S;
  const key = S.E.tasks.find(t => t.MRN === '900001' && t.Task === 'Update benefits in NextGen')['Task Key'];
  S.view = 'eligibility'; api.render();
  // a fake focusable: focus() makes it the active element and logs it
  const log = [];
  const node = (name, dataset) => { const n = { name, dataset: dataset || {}, focus(){ doc.activeElement = n; log.push(name); } }; return n; };
  const opener = node('note button', { act: 'seTaskNote', k: key, m: '900001' });
  const app = []; els['#app'].querySelector = sel => { app.push(sel); return node('note button again', opener.dataset); };
  const text = node('textarea'), save = node('save', { act: 'seTaskNoteSave', k: key }), cancel = node('cancel', { act: 'close' });
  const sheet = doc.querySelector('#sheet');
  sheet.querySelector = sel => (sel === 'textarea, input, select, button' ? text : null);
  sheet.querySelectorAll = () => [text, save, cancel];
  const sel = `[data-act="seTaskNote"][data-k="${key}"][data-m="900001"]`;
  const open = async () => { doc.activeElement = opener; await fire('click', { target: { closest: s => (s === '[data-act]' ? { tagName: 'BUTTON', dataset: opener.dataset } : null) } }); };
  // the keydown handler is synchronous: returns whether it took the key
  const press = (k, shiftKey) => { let took = false; fire('keydown', { key: k, shiftKey: !!shiftKey, preventDefault(){ took = true; } }); return took; };
  // opening moves focus to the note box
  await open();
  assert.equal(els['#modal'].hidden, false);
  assert.equal(doc.activeElement, text);
  // Tab walks the dialog's own controls and wraps; Shift+Tab wraps back
  assert.equal(press('Tab'), true); assert.equal(doc.activeElement, save);
  press('Tab'); assert.equal(doc.activeElement, cancel);
  press('Tab'); assert.equal(doc.activeElement, text);
  press('Tab', true); assert.equal(doc.activeElement, cancel);
  // Escape closes it and focus goes back to the Note button
  press('Escape');
  assert.equal(els['#modal'].hidden, true);
  assert.deepEqual([app[app.length - 1], doc.activeElement.name], [sel, 'note button again']);
  // a closed dialog leaves Tab alone
  assert.equal(press('Tab'), false);
  // Save re-draws the list; focus lands on the same Note button in the new list
  await open(); app.length = 0;
  els['#seNoteText'] = { value: 'Called the patient' };
  await fire('click', { target: { closest: s => (s === '[data-act]' ? { tagName: 'BUTTON', dataset: save.dataset } : null) } });
  assert.deepEqual(app, [sel, sel], 'once on close, once after the re-draw');
  assert.equal(doc.activeElement.name, 'note button again');
  assert.equal(S.E.tasks.find(t => t['Task Key'] === key).clinicNote, 'Called the patient');
});
