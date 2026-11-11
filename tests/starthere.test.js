// Sohar v0.30 clinic workbook (Start Here + the clinic's Appointment List, LASFIR only): CTT opens it as a plain workbook
// (no SE Info), reads the Appointment List for templates, access and cadence, and keys patients on MRN with the LASFIR
// code as the patient label. Synthetic data only: made-up names and MRNs 9002xx.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'); const path = require('path'); const vm = require('vm');
const C = require('../core.js');
const R = require('../rules.js');
const PR = require('../payer-rules.js');
const TE = require('../engine.js');
const TC = require('../cadence.js');

const HEAD = ['Res Name', 'Pat Name', 'Appt Date', 'Beg Tm', 'Event', 'Details', 'Loc Name', 'MRN', 'Status', 'Dur', 'Birth Dt', 'Rendering', 'Per Nbr', 'Guar Name',
  'Payer Name (Cht Guar)', 'Pol Nbr (Cht Guar)', 'Crt Dt'];
const row = (o) => HEAD.map((h) => (o[h] == null ? '' : o[h]));
const appt = (name, mrn, per, date, time, event, status, loc) => row({ 'Res Name': 'Reyes, Sam MD', 'Pat Name': name, 'Appt Date': date, 'Beg Tm': time, Event: event,
  Details: `${name} asked about a copay`, 'Loc Name': loc, MRN: mrn, Status: status, Dur: '30', 'Birth Dt': '2/3/1971', Rendering: 'Reyes, Sam MD', 'Per Nbr': per,
  'Guar Name': 'Guardia, Pat', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'ZXQ123', 'Crt Dt': '10/1/2026' });
const GRID = [['Appointments Listing'], [''], HEAD,
  appt('Featherstone, Ophelia', '900201', '701', '10/14/2026', '9:00 AM', 'Med Management', 'Kept', 'San Rafael'),
  appt('Featherstone, Ophelia', '900201', '701', '11/11/2026', '9:00 AM', 'Med Management', 'Kept', 'San Rafael'),
  appt('Featherstone, Ophelia', '900201', '701', '12/9/2026', '9:00 AM', 'Med Management', 'Expected', 'San Rafael'),
  appt('Quintero, Bram', '900202', '702', '11/12/2026', '10:00 AM', 'Telepsych', 'Expected', 'Castro'),
  ['Total Appointments: 4']];

function packGrids() {
  const pack = R.clinicPack('San Rafael', { eligibility: false, importGrid: GRID, runAt: new Date(2026, 10, 10, 9), version: '0.30.0' });
  return { pack, grids: Object.fromEntries(pack.sheets.map((s) => [s.name, [s.columns, ...s.rows.map((r) => s.columns.map((c) => (r[c] == null ? '' : r[c])))]])) };
}

test('engine reads the LASFIR Appointment List: patient column found, codes kept, MRN key; names never there', () => {
  const { grids } = packGrids();
  const ap = TE.parseAppts(grids['Appointment List']);
  assert.ok(!ap.error, ap.error);
  assert.equal(ap.has.pt, true);
  assert.equal(ap.has.mrn, true);
  assert.deepEqual([...new Set(ap.rows.map((r) => r.pt))], ['FEAOPH']);
  assert.equal(TE.lasfir('FEAOPH'), 'FEAOPH', 'a code stays a code');
  assert.equal(TE.lasfir('Featherstone, Ophelia'), 'FEAOPH', 'a name still becomes a code');
  assert.equal(ap.rows.length, 3, 'only this clinic');
  const v = TC.visits(ap.rows, [], { asOf: TE.parseDate('11/10/2026') });
  const list = Array.isArray(v) ? v : v.list;
  assert.ok(list.length > 0 && list.every((x) => x.code === 'FEAOPH' && /^m:/.test(x.key)), `keyed on MRN, labelled with the code (${list.length} visits)`);
  const text = JSON.stringify(grids);
  for (const s of ['Featherstone', 'Ophelia', 'Quintero', 'Guardia', '2/3/1971', 'ZXQ123', 'copay', 'Castro']) assert.ok(!text.includes(s), `"${s}" is not in the file`);
});

test('a Start Here workbook loads like any workbook: no SE Info, the Appointment List is the appointment tab, Start Here is not a template', async () => {
  const { pack, grids } = packGrids();
  const sheets = pack.sheets.map((s) => ({ name: s.name, rows: grids[s.name] }));
  const ws = (s) => ({ name: s.name, visibility: 'Visible',
    getUsedRangeOrNullObject(){ const h = s.rows.length, w = Math.max(0, ...s.rows.map((r) => r.length)); return h && w ? { isNullObject: false, load(){}, rowIndex: 0, columnIndex: 0, rowCount: h, columnCount: w } : { isNullObject: true, load(){} }; },
    getRangeByIndexes: (r, c, h, w) => ({ load(){}, get values(){ return Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => { const v = (s.rows[r + i] || [])[c + j]; return v == null ? '' : v; })); } }) });
  const ctx = { sync: async () => {}, workbook: { worksheets: { load(){}, items: sheets.map(ws) } } };
  const el = () => ({ addEventListener(){}, appendChild(){}, remove(){}, focus(){}, querySelector: () => null, querySelectorAll: () => [], hidden: true, innerHTML: '', value: '', classList: { toggle(){} } });
  const els = {};
  const sandbox = { console, setTimeout: (f, ms) => { const h = setTimeout(f, ms); if (h.unref) h.unref(); return h; }, clearTimeout, Excel: { run: (fn) => fn(ctx) }, Office: undefined, window: {},
    location: { href: 'https://x/taskpane.html', search: '' }, document: { querySelector: (s) => els[s] || (els[s] = el()), querySelectorAll: () => [], addEventListener(){}, createElement: el, head: el() },
    TE, TC, PayerRules: PR, SoharCore: C, SoharRules: R };
  sandbox.self = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8') + '\n;this.__api = { OfficeStore, prepare };', sandbox);
  const raw = await sandbox.__api.OfficeStore().load();
  assert.equal(raw.se, undefined, 'not a Clinic Eligibility file');
  assert.equal(raw.appt.sheet, 'Appointment List');
  const d = sandbox.__api.prepare(raw);
  assert.equal(d.se, null);
  assert.equal(d.templates.length, 0, 'Start Here holds no template grid');
  assert.equal(d.ap.rows.length, 3);
});
