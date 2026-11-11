// CTT 1.2.0: the Clinic Eligibility file v2 in the pane. Clinic Day (written by this add-in from the hidden tables and
// Clinic Work), OK to collect, desk collections, pricing from the file's fee tables, and the patient lookup client.
// Synthetic data only: made-up names, MRNs 9001xx, made-up rates. Dates are relative to today (the pane derives at now).
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'); const path = require('path'); const vm = require('vm');
const C = require('../core.js');
const R = require('../rules.js');
const F = require('../fees.js');
const PR = require('../payer-rules.js');
const LK = require('../lookup.js');
const PRICE = require('../pricing.js');

const day = n => { const d = new Date(); d.setHours(12, 0, 0, 0); d.setDate(d.getDate() + n); return d; };
const mdy = n => { const d = day(n); return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`; };
const mmdd = n => { const d = day(n); return `${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}/${d.getFullYear()}`; };
const NOW = new Date();
const HDR = ['Pat Name', 'Appt Dt', 'Beg Tm', 'Event', 'Loc Name', 'Per Nbr', 'Md Rc', 'Birth Dt', 'Sts', 'Guar Name', 'Guar Birth Dt', 'Guar Zip',
  'Payer Name (Cht Guar)', 'Pol Nbr (Cht Guar)', 'Sec Payer (Cht Guar)', 'Rendering'];
const row = o => HDR.map(k => (o[k] == null ? '' : o[k]));
const LOCSTATE = { 'San Rafael': 'CA', 'Castro': 'CA' };
const MAP = [{ nextgenPayerName: 'Cigna', payerId: '62308', payerName: 'Cigna', confidence: 'Exact' }];
const CFG = { asOf: new Date(Date.UTC(day(0).getFullYear(), day(0).getMonth(), day(0).getDate())), scope: 'upcoming', locationState: LOCSTATE, teleKeywords: ['Tele'], specialty: 'PSY', payerMap: MAP };
const ap = (per, mrn, name, n, time, event, loc, extra) => row({ 'Per Nbr': per, 'Md Rc': mrn, 'Pat Name': name, 'Birth Dt': '3/4/1980', 'Appt Dt': mdy(n), 'Beg Tm': time, Event: event,
  'Loc Name': loc, Sts: 'Expected', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': `C${per}`, 'Guar Zip': '94114', Rendering: 'Reyes, Sam MD', ...(extra || {}) });
const IMPORT = () => [HDR,
  ap('601', '900101', 'Larkspur, Wendeline', 2, '9:00 AM', 'Med Management', 'San Rafael'),
  ap('601', '900101', 'Larkspur, Wendeline', 4, '10:00 AM', 'Esketamine Treatment', 'San Rafael'),
  ap('602', '900102', 'Quincefield, Ambrose', 1, '1:00 PM', 'Med Management', 'Castro'),
  ap('602', '900102', 'Quincefield, Ambrose', 6, '2:30 PM', 'Med Management', 'San Rafael'),
];
const NEVER = ['Larkspur', 'Wendeline', 'Quincefield', 'Ambrose', '3/4/1980', '03/04/1980', '94114', '128.21'];
const VER_HDR = ['requestId', 'id', 'url', 'created', 'resolved', 'memberId', 'payerName', 'status', 'firstName', 'lastName', 'planName', 'innCopay',
  'innTotalDeductible', 'innRemainingDeductible', 'innOutOfPocket', 'innOutOfPocketRemaining'];
const ver = o => VER_HDR.map(k => (o[k] == null ? '' : String(o[k])));
const iso0 = new Date(NOW.getTime() - 864e5).toISOString();
const FILES = () => [{ name: 'verification output.csv', rows: [VER_HDR,
  ver({ requestId: 'v-601', id: '900101', url: 'https://app.soharhealth.com/verifications/v-601', created: iso0.slice(0, 10), resolved: iso0, memberId: 'C601', payerName: 'Cigna', status: 'complete.eligible',
    firstName: 'Wendeline', lastName: 'Larkspur', planName: 'SYNTH OPEN ACCESS', innCopay: '30', innTotalDeductible: '1000', innRemainingDeductible: '100', innOutOfPocket: '4000', innOutOfPocketRemaining: '2500' }),
  ver({ requestId: 'v-602', id: '900102', url: 'https://app.soharhealth.com/verifications/v-602', created: iso0.slice(0, 10), resolved: iso0, memberId: 'C602', payerName: 'Cigna', status: 'complete.eligible',
    planName: 'SYNTH OPEN ACCESS', innCopay: '20', innTotalDeductible: '500', innRemainingDeductible: '0', innOutOfPocket: '3000', innOutOfPocketRemaining: '1000' }),
]}];
const L = o => Object.assign(Object.fromEntries(R.LEDGER_COLS.map(c => [c, ''])), o);
const FEE_ROWS = ['99214', '99213'].map(code => ({ Tab: 'Marin', Locality: 'Locality 52', Locations: 'San Rafael', Code: code, Description: 'Synthetic visit', Payer: 'CIGNA', Level: 'MD', Rate: code === '99214' ? 150 : 100, Note: '' }))
  .concat(['99214', '99213'].map(code => ({ Tab: 'Marin', Locality: 'Locality 52', Locations: 'San Rafael', Code: code, Description: 'Synthetic visit', Payer: 'CIGNA', Level: 'NP', Rate: code === '99214' ? 120 : 80, Note: '' })));
function src(extra){
  const built = C.buildPatients(IMPORT(), CFG);
  const merged = C.mergeSoharResults(IMPORT(), FILES(), CFG);
  const patients = C.applyReverifySkips(built.patients, merged.results, { ...C.REVERIFY_DEFAULTS, today: NOW }).patients.filter(p => p['In Scope'] === 'Yes');
  return Object.assign({ patients, results: merged.results, review: built.review, tasks: [], learningRows: [], metricRows: [], cfg: { locationState: LOCSTATE }, payerRuleRows: [], runAt: NOW, version: '0.29.0',
    ledger: [L({ MRN: '900101', 'Appt Date': mmdd(2), Time: '9:00 AM', Location: 'San Rafael', 'Projected Patient Responsibility': '45', Allowed: '128.21' })],
    cards: [{ MRN: '900101', 'Card On File': 'Yes', Source: 'Luma checklist', 'As Of': mmdd(-1) }],
    consents: [], desk: [], history: [], copayActivity: [],
    fees: { columns: F.FEE_COLS, rows: FEE_ROWS, rules: [
      { Kind: 'locations', Record: [{ 'NextGen Location': 'San Rafael', 'Fee Tab': 'Marin', Source: 'Manual', 'Auto Match': '', Note: '' }] },
      { Kind: 'payerMap', Record: [{ Payer: 'Cigna', 'Sohar payerId': '62308', 'Fee Tab': 'Marin', Region: 'CA', 'Fee Payer': 'CIGNA', Confidence: 'Exact', Source: 'Manual', 'Auto Match': '', Note: '' }] },
      { Kind: 'codes', Record: F.ESTIMATE_CODE_SEEDS }, { Kind: 'defaults', Record: [] },
      { Kind: 'config', Record: { teleKeywords: ['Tele'], locationState: LOCSTATE } },
    ] },
  }, extra || {});
}
const gridsOf = pack => Object.fromEntries(pack.sheets.map(s => [s.name, [s.columns, ...s.rows.map(r => s.columns.map(c => (r[c] == null ? '' : r[c])))]]));
const bookOf = pack => pack.sheets.map(s => ({ name: s.name, hidden: !!s.hidden, rows: gridsOf(pack)[s.name] }));
function mockBook(sheets){
  const ws = s => ({ name: s.name, visibility: s.hidden ? 'Hidden' : 'Visible',
    getUsedRangeOrNullObject(){ const h = s.rows.length, w = Math.max(0, ...s.rows.map(r => r.length)); return h && w ? { isNullObject: false, load(){}, rowIndex: 0, columnIndex: 0, rowCount: h, columnCount: w } : { isNullObject: true, load(){} }; },
    getRangeByIndexes: (r, c, h, w) => ({ load(){}, get values(){ return Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => { const v = (s.rows[r + i] || [])[c + j]; return v == null ? '' : v; })); } }) });
  const ctx = { sync: async () => {}, workbook: { worksheets: { load(){}, items: sheets.map(ws) } } };
  return { run: fn => fn(ctx) };
}
function loadApp(Excel, extra){
  const els = {}, on = {}, store = {};
  const el = () => ({ addEventListener(){}, appendChild(){}, remove(){}, focus(){}, querySelector: () => null, querySelectorAll: () => [], hidden: true, innerHTML: '', value: '', classList: { toggle(){} } });
  const later = (f, ms) => { const h = setTimeout(f, ms); if (h.unref) h.unref(); return h; };
  const sandbox = { console, setTimeout: later, clearTimeout, Excel, Office: undefined, window: {}, location: { href: 'https://x/taskpane.html', search: '' },
    localStorage: { getItem: k => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); } },
    document: { querySelector: s => els[s] || (els[s] = el()), querySelectorAll: () => [], addEventListener(k, f){ (on[k] = on[k] || []).push(f); }, createElement: el, head: el() } };
  sandbox.TE = require('../engine.js'); sandbox.TC = require('../cadence.js'); sandbox.PayerRules = PR; sandbox.SoharCore = C; sandbox.SoharRules = R; sandbox.SoharFees = F; sandbox.CTTPricing = PRICE; sandbox.self = sandbox;
  Object.assign(sandbox, extra || {});
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8') + '\n;this.__api = { OfficeStore, S, reload, render, seWrite };', sandbox);
  const fire = (k, e) => Promise.all((on[k] || []).map(f => f(e)));
  return { api: sandbox.__api, els, fire, store };
}
const settle = () => new Promise(r => setTimeout(r, 30));
async function appWith(sheets, extra){
  const { api, els, fire, store } = loadApp(mockBook(sheets), extra);
  const writes = [];
  let work = (sheets.find(s => s.name === 'Clinic Work') || {}).rows || [];
  const writeSheet = async (n, aoa, fmt) => {
    if (!(fmt && fmt.change)){ writes.push([n, aoa, fmt]); return { table: 'ok' }; }
    const rows = R.workRows(work, fmt.change); work = [R.CLINIC_WORK_COLS].concat(rows.map(x => R.CLINIC_WORK_COLS.map(c => x[c])));
    writes.push([n, work, fmt]); return { table: 'ok', rows };
  };
  api.S.store = { live: true, load: api.OfficeStore().load, writeSheet, getDoc: () => null, setDoc: async () => {} };
  api.S.dayDelay = 0;
  await api.reload(); api.S.name = 'Desk A'; await settle();
  return { api, els, fire, writes, store };
}
const click = (fire, ds) => fire('click', { target: { closest: sel => (sel === '[data-act]' ? { tagName: 'BUTTON', dataset: ds } : null) } });
const dayWrites = writes => writes.filter(w => w[0] === 'Clinic Day');

test('Clinic Day: written from the file on open (no MRN, no names), with the ledger estimate, a priced visit and the guest\'s visit here', async () => {
  const pack = R.clinicPack('San Rafael', src());
  const { api, writes } = await appWith(bookOf(pack));
  assert.ok(api.S.E && api.S.E.appts.length === 3, 'three visits here');
  const dw = dayWrites(writes);
  assert.equal(dw.length, 1, 'one write after open');
  const [name, aoa, fmt] = dw[0];
  assert.equal(name, 'Clinic Day');
  assert.deepEqual(aoa[0], R.CLINIC_DAY_COLS);
  assert.deepEqual(JSON.parse(JSON.stringify(fmt)), { keepFormat: true, inPlace: true, text: true, table: 'tblClinicDay' });
  const rows = aoa.slice(1).map(r => Object.fromEntries(R.CLINIC_DAY_COLS.map((c, i) => [c, r[i]])));
  assert.deepEqual(JSON.parse(JSON.stringify(rows.map(r => r.LASFIR))), ['LARWEN', 'LARWEN', 'QUIAMB']);
  assert.equal(rows[0]['Est. owes'], '$45', 'the rotation\'s estimate');
  assert.equal(rows[1]['Est. owes'], 'No estimate: check DED/OOP', 'esketamine is never priced');
  assert.match(rows[2]['Est. owes'], /^\$[\d.,]+ \(est\.\)$/, 'the guest\'s visit priced from the file\'s fee tables');
  assert.match(rows[0]['Card on file'], /^Yes/);
  const text = JSON.stringify(aoa);
  assert.ok(!/9001\d\d/.test(text), 'no MRN on Clinic Day');
  for (const s of NEVER) assert.ok(!text.includes(s), `"${s}" is not on Clinic Day`);
  // nothing changed: no second write
  api.render(); await settle();
  assert.equal(dayWrites(writes).length, 1);
});

test('OK to collect and Record a collection go on Clinic Work with who and when, and Clinic Day follows', async () => {
  const { api, els, fire, writes } = await appWith(bookOf(R.clinicPack('San Rafael', src())));
  await click(fire, { act: 'seOk', m: '900102', v: 'Yes' }); await settle();
  let w = writes.filter(x => x[0] === 'Clinic Work').pop()[1];
  const ok = R.tableObjects(w).find(r => r.Kind === 'OK to collect');
  assert.deepEqual([ok.Key, ok.Value, ok.By], ['900102', 'Yes', 'Desk A']);
  assert.ok(ok.When);
  let last = dayWrites(writes).pop()[1].slice(1).find(r => r[2] === 'QUIAMB');
  assert.match(last[R.CLINIC_DAY_COLS.indexOf('OK to collect')], /^Yes · Desk A/);
  // a collection: amount, method, note; a card number typed in the note is taken out
  els['#seAmt'] = { value: '30' }; els['#seMethod'] = { value: 'Cash' }; els['#seCollectNote'] = { value: 'copay 4111 1111 1111 1111' };
  await click(fire, { act: 'seCollectSave', m: '900101' }); await settle();
  w = writes.filter(x => x[0] === 'Clinic Work').pop()[1];
  const col = R.tableObjects(w).find(r => r.Kind === 'Collected');
  assert.equal(col.Value, '30.00');
  assert.equal(col.Note, 'Cash · copay [card number removed]');
  assert.match(col.Key, /^900101\|\d{4}-\d{2}-\d{2}T/);
  assert.equal(api.S.E.collected.length, 1, 'pending until the admin applies it');
  els['#seAmt'] = { value: 'thirty' };
  await click(fire, { act: 'seCollectSave', m: '900101' });
  assert.equal(R.tableObjects(writes.filter(x => x[0] === 'Clinic Work').pop()[1]).filter(r => r.Kind === 'Collected').length, 1, 'a word is not an amount');
});

test('an older clinic file (no Clinic Day tab) gets no Clinic Day write: the add-in never adds a sheet to a clinic file', async () => {
  const pack = R.clinicPack('San Rafael', src());
  const sheets = bookOf(pack).filter(s => s.name !== 'Clinic Day');
  const { api, writes } = await appWith(sheets);
  assert.ok(api.S.E);
  assert.equal(dayWrites(writes).length, 0);
});

test('lookup.js: one host, https only, a reason unless scheduled, only known fields kept, the key decides the clinic', async () => {
  const H = 'flow.example.test';
  assert.equal(LK.checkSetup({ url: 'https://flow.example.test/x', key: 'k'.repeat(20) }, H).ok, true);
  assert.match(LK.checkSetup({ url: 'https://evil.example.test/x', key: 'k'.repeat(20) }, H).error, /not the clinic lookup address/);
  assert.match(LK.checkSetup({ url: 'http://flow.example.test/x', key: 'k'.repeat(20) }, H).error, /not the clinic lookup address/);
  assert.match(LK.checkSetup({ url: 'https://flow.example.test/x', key: 'short' }, H).error, /whole clinic key/);
  assert.match(LK.checkSetup({ url: 'https://flow.example.test/x', key: 'k'.repeat(20) }, '').error, /not switched on/);
  assert.throws(() => LK.buildRequest({ mrn: '900101', purpose: 'unscheduled', staff: 'Desk A' }), /Pick why/);
  assert.throws(() => LK.buildRequest({ mrn: '900101', purpose: 'unscheduled', reason: 'Other (add a note)', staff: 'Desk A' }), /short note/);
  assert.throws(() => LK.buildRequest({ mrn: 'Larkspur', purpose: 'scheduled', staff: 'Desk A' }), /digits only/);
  const b = LK.buildRequest({ mrn: ' 900101 ', purpose: 'scheduled', staff: 'Desk A', requestId: 'r1' });
  assert.deepEqual(Object.keys(b).sort(), ['key', 'mrn', 'note', 'purpose', 'reason', 'requestId', 'staff', 'v']);
  assert.ok(!('clinic' in b) && !('location' in b), 'the request never names the clinic: the key does');
  const sent = [];
  const fetchOk = async (u, o) => { sent.push([u, o]); return { ok: true, status: 200, json: async () => ({ ok: true, tier: 'seen-12mo', patient: { LASFIR: 'LARWEN', MRN: '900101', Copay: '30', Name: 'Larkspur, Wendeline', DOB: '3/4/1980' }, history: [{ When: '11/1', Kind: 'Copay', Amount: '30', Card: '4111' }], extra: 'x' }) }; };
  const res = await LK.call(fetchOk, { url: 'https://flow.example.test/x', key: 'k'.repeat(20) }, b, { host: H });
  assert.equal(res.tier, 'seen-12mo');
  assert.deepEqual(Object.keys(res.patient), LK.PATIENT_FIELDS);
  assert.ok(!JSON.stringify(res).includes('Larkspur') && !JSON.stringify(res).includes('1980') && !JSON.stringify(res).includes('4111'), 'fields the add-in does not show are dropped');
  const [u, o] = sent[0];
  assert.equal(u, 'https://flow.example.test/x');
  assert.equal(o.method, 'POST'); assert.equal(o.headers['Content-Type'], 'text/plain;charset=UTF-8', 'no preflight'); assert.equal(o.credentials, 'omit'); assert.equal(o.referrerPolicy, 'no-referrer');
  assert.equal(JSON.parse(o.body).key, 'k'.repeat(20));
  await assert.rejects(LK.call(async () => ({ ok: false, status: 401, json: async () => ({}) }), { url: 'https://flow.example.test/x', key: 'k'.repeat(20) }, b, { host: H }), /key was not accepted/);
  await assert.rejects(LK.call(async () => { throw new TypeError('Failed to fetch'); }, { url: 'https://flow.example.test/x', key: 'k'.repeat(20) }, b, { host: H }), /could not be reached/);
});

test('CSP: connect-src is \'self\' plus exactly the lookup host (none while LOOKUP_HOST is blank); no other network script', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'taskpane.html'), 'utf8');
  const csp = (html.match(/Content-Security-Policy" content="([^"]+)"/) || [])[1];
  assert.equal(csp, `connect-src 'self'${LK.LOOKUP_HOST ? ` https://${LK.LOOKUP_HOST}` : ''}`);
  const scripts = [...html.matchAll(/<script src="([^"]+)"/g)].map(m => m[1]);
  assert.deepEqual(scripts.filter(s => /^https?:/.test(s)), ['https://appsforoffice.microsoft.com/lib/1/hosted/office.js']);
  for (const f of ['fees.js', 'pricing.js', 'lookup.js']) assert.ok(scripts.some(s => s.startsWith(f + '?v=')), f);
  const app = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  // two fetches: the install page's manifest download (same site) and the lookup, which goes through lookup.js's host check
  assert.deepEqual(app.match(/\bfetch\([^)]*\)/g), ['fetch(u, o)', "fetch('manifest.xml', { cache: 'no-store' })"]);
  assert.ok(!/localStorage\.setItem\([^)]*lk\.res|LS\.set\([^)]*res/.test(app), 'a lookup result is never stored');
});

test('lookup in the pane: shown in memory only; nothing goes on a sheet or in browser storage but the setup', async () => {
  const H = 'flow.example.test';
  const LKH = Object.assign({}, LK, { LOOKUP_HOST: H, checkSetup: (c, h) => LK.checkSetup(c, h === undefined ? H : h), call: (f, c, b, o) => LK.call(f, c, b, Object.assign({ host: H }, o)) });
  let posted = null;
  const fetch = async (u, o) => { posted = JSON.parse(o.body); return { ok: true, status: 200, json: async () => ({ ok: true, tier: 'scheduled', patient: { LASFIR: 'LARWEN', Copay: '30' }, history: [] }) }; };
  const { api, els, fire, writes, store } = await appWith(bookOf(R.clinicPack('San Rafael', src())), { CTTLookup: LKH, fetch });
  store.cttLookup = JSON.stringify({ url: `https://${H}/x`, key: 'k'.repeat(20) });
  const before = writes.length;
  await click(fire, { act: 'seMore', m: '900101' }); await settle();
  assert.deepEqual([posted.mrn, posted.purpose, posted.staff], ['900101', 'scheduled', 'Desk A']);
  assert.equal(api.S.lk.res.patient.Copay, '30');
  assert.ok(els['#sheet'].innerHTML.includes('nothing is saved in this file'));
  assert.equal(writes.length, before, 'no sheet write');
  assert.deepEqual(Object.keys(store), ['cttLookup'], 'only the setup is in browser storage');
  await click(fire, { act: 'lkClose' });
  assert.equal(api.S.lk, null, 'dropped when the box closes');
});

test('pricing.js: null without fee tables; an intervention is never priced; MD vs NP level from the provider', () => {
  assert.equal(PRICE.pricer({ fees: [] }, { SR: R, SF: F, SC: C }), null);
  assert.equal(PRICE.levelOf('Lark, Jordan PMHNP'), 'np');
  assert.equal(PRICE.levelOf('Reyes, Sam MD'), 'md');
  const u = R.clinicUnpack(gridsOf(R.clinicPack('San Rafael', src())));
  const p = PRICE.pricer(u, { SR: R, SF: F, SC: C });
  assert.equal(typeof p, 'function');
  assert.deepEqual(p({ Key: '900101', 'Per Nbr': '601', Event: 'Esketamine Treatment', Provider: 'Reyes, Sam MD' }), { note: 'No estimate: check DED/OOP' });
  const md = p({ Key: '900102', 'Per Nbr': '602', Event: 'Med Management', Provider: 'Reyes, Sam MD' });
  const np = p({ Key: '900102', 'Per Nbr': '602', Event: 'Med Management', Provider: 'Lark, Jordan PMHNP' });
  assert.ok(md.amount > 0 && np.amount > 0, JSON.stringify([md, np]));
});
