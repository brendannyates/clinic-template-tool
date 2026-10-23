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
function loadApp(Excel){
  const els = {};
  const el = () => ({ addEventListener(){}, appendChild(){}, querySelectorAll: () => [], hidden: true, innerHTML: '', value: '', classList: { toggle(){} } });
  const sandbox = { console, setTimeout, clearTimeout, Excel, Office: undefined, window: {}, location: { href: 'https://x/taskpane.html', search: '' },
    document: { querySelector: s => els[s] || (els[s] = el()), querySelectorAll: () => [], addEventListener(){}, createElement: el, head: el() } };
  sandbox.TE = require('../engine.js'); sandbox.TC = require('../cadence.js'); sandbox.PayerRules = PR; sandbox.SoharCore = C; sandbox.SoharRules = R; sandbox.self = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8') + '\n;this.__api = { OfficeStore, S, prepare, reload, render, seWrite, shownViews, navGroups, currentView };', sandbox);
  return { api: sandbox.__api, els, doc: sandbox.document };
}
// an app whose store is the mock workbook (load) plus a writeSheet log
async function appWith(sheets){
  const { Excel, reads } = mockBook(sheets); const { api, els, doc } = loadApp(Excel);
  const writes = [], docs = {};
  const office = api.OfficeStore();
  api.S.store = { live: true, load: office.load, writeSheet: async (n, aoa, fmt) => { writes.push([n, aoa, fmt]); }, getDoc: k => docs[k] ?? null, setDoc: async (k, v) => { docs[k] = v; } };
  await api.reload();
  return { api, els, doc, writes, docs, reads };
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
  assert.ok(writes.every(w => w[0] === 'Clinic Work' && w[2].keepFormat && w[2].inPlace && w[2].text && w[1][0].join() === R.CLINIC_WORK_COLS.join()), 'as text, over the old rows');
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
