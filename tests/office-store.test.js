// Mock-Excel test of the workbook writes: every values/formulas/numberFormat array must match its range size,
// which is what Excel on the web rejects with InvalidArgument when it doesn't.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'); const path = require('path'); const vm = require('vm');

function colNum(s){ let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }
function dims(addr){
  const rr = addr.match(/^(\d+):(\d+)$/); if (rr) return { rows: rr[2] - rr[1] + 1, cols: 16384 };   // whole rows
  const m = addr.replace(/\$/g, '').match(/^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/); if (!m) throw new Error('bad address ' + addr);
  if (!m[2]) return { rows: 1048576, cols: (m[3] ? colNum(m[3]) : colNum(m[1])) - colNum(m[1]) + 1 };
  const r1 = +m[2], c1 = colNum(m[1]), r2 = m[4] ? +m[4] : r1, c2 = m[3] ? colNum(m[3]) : c1;
  return { rows: r2 - r1 + 1, cols: c2 - c1 + 1 };
}
function mockExcel(existing = []){
  const log = [];
  const range = (label, d) => {
    const check = (k, v) => { assert.ok(Array.isArray(v) && v.length === d.rows && v.every(r => Array.isArray(r) && r.length === d.cols), `${k} on ${label} must be ${d.rows}×${d.cols}, got ${v.length}×${v[0] && v[0].length}`); log.push([label, k, v]); };
    const r = { format: { font: {}, fill: {}, borders: { getItem: () => ({}) }, autofitColumns(){} }, dataValidation: { load(){}, rule: null, clear(){ log.push([label, 'dvClear']); } }, clear(a){ log.push([label, 'clear', a]); }, delete(d){ log.push([label, 'delete', d]); }, select(){} };
    for (const k of ['values', 'formulas', 'numberFormat']) Object.defineProperty(r, k, { set: v => check(k, v) });
    return r;
  };
  const sheet = name => ({ name, getRange: a => range(name + '!' + (a || 'all'), a ? dims(a) : { rows: 1048576, cols: 16384 }), getRangeByIndexes: (r, c, h, w) => range(`${name}!R${r}C${c}`, { rows: h, cols: w }), freezePanes: { freezeRows(){} }, activate(){} });
  const ctx = { sync: async () => {}, workbook: { worksheets: { getItemOrNullObject: n => existing.includes(n) ? sheet(n) : { isNullObject: true }, getItem: n => ({ copy: () => { log.push(['copy', n]); return sheet(n + ' (2)'); } }), add: n => { log.push(['add', n]); return sheet(n); } } } };
  return { Excel: { run: fn => fn(ctx) }, log };
}
function loadApp(Excel){
  const el = () => ({ addEventListener(){}, appendChild(){}, querySelectorAll: () => [], hidden: true, innerHTML: '', classList: { toggle(){} } });
  const sandbox = { console, setTimeout, clearTimeout, Excel, Office: undefined, window: {}, location: { href: 'https://x/taskpane.html', search: '' },
    document: { querySelector: el, querySelectorAll: () => [], addEventListener(){}, createElement: el, head: el() } };
  sandbox.TE = require('../engine.js'); sandbox.TC = require('../cadence.js'); sandbox.PayerRules = require('../payer-rules.js'); sandbox.self = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8') + '\n;this.__api = { OfficeStore, S };', sandbox);
  return sandbox.__api;
}

test('createTemplate writes a grid whose formulas and formats fit their ranges', async () => {
  const { Excel, log } = mockExcel(); const { OfficeStore } = loadApp(Excel);
  const times = []; for (let m = 480; m < 1080; m += 30) times.push(m);
  const cells = [[], ['Reyes MD, Samuel'], ['Effective From:', 46300, '', 'Effective To:', ''], ['Location:', 'Eastside', 'Weekly Clinic Template', '', 'Last Modified:', 46300, '', '', 'Event Types', 'Slots', 'Hours', 'M', 'T', 'W', 'TH', 'F'], ['Provider:', 'Reyes', 'Reyes', 'Reyes', 'Reyes', 'Reyes'], ['Time', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']];
  for (const m of times) cells.push([m / 1440, 'Admin', '', '', '', '']);
  await OfficeStore().createTemplate({ sheet: 'TPL Reyes 2026-10-15', cells, times, step: 30 });
  assert.ok(log.some(x => x[0] === 'add' && x[1] === 'TPL Reyes 2026-10-15'));
  const f = log.filter(x => x[1] === 'formulas').map(x => x[2][0][0]);
  assert.ok(f.some(x => x.includes('UNIQUE(TOCOL(B7:F26,1))')), 'event list formula spans the grid');
  assert.ok(f.some(x => x.includes("'Slot Types'!$B$2:$B$300=\"Clinical\"")));
});

test('writeSheet and replaceAppts write ragged rows as full rectangles', async () => {
  const { Excel, log } = mockExcel(); const { OfficeStore } = loadApp(Excel);
  const st = OfficeStore();
  await st.writeSheet('TA Summary', [['A', 'B', 'C'], [1, 2], [3]], { numFmt: { 1: '0.0%' } });
  const rows = Array.from({ length: 2600 }, (_, i) => i ? ['', i, 46000 + i, '9:00 A'] : ['', 'MRN', 'Appt Date', 'Time']);
  await st.replaceAppts('Appointment List', rows, 2);
  const vals = log.filter(x => x[1] === 'values');
  assert.equal(vals.reduce((n, x) => n + x[2].length, 0), 3 + 2600);
});

test('the template the add-in writes parses back as a template', () => {
  const TE = require('../engine.js');
  const times = [480, 510, 540];
  const cells = [[], ['Reyes MD, Samuel'], ['Effective From:', 46300, '', 'Effective To:', ''], ['Location:', 'Eastside', 'Weekly Clinic Template', '', 'Last Modified:', 46300], ['Provider:', 'Reyes', 'Reyes', 'Reyes', 'Reyes', 'Reyes'], ['Time', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']];
  for (const m of times) cells.push([m / 1440, 'New Patient', 'Admin', '', '', 'Telepsych']);
  const [t] = TE.parseTemplates('TPL', cells);
  assert.equal(t.label, 'Reyes'); assert.equal(t.effFrom, 46300); assert.equal(t.location, 'Eastside'); assert.equal(t.slots.length, 9); assert.equal(t.step, 30);
});

test('copyTemplate duplicates the tab, writes headers, deletes other grids bottom up and repairs the dropdown', async () => {
  const { Excel, log } = mockExcel(); const { OfficeStore } = loadApp(Excel);
  await OfficeStore().copyTemplate({ src: 'Sheet1', sheet: 'TPL Kim 2026-09-01', writes: [[3, 1, 'Effective From:'], [3, 2, 46300, 'm/d/yyyy']], clearSlots: true,
    slots: [33, 2, 18, 5], slotsAfter: [6, 2, 18, 5], deleteRows: [[57, 104], [0, 29]], dvSource: "='Slot Types'!$A$2:$A$300" });
  assert.ok(log.some(x => x[0] === 'copy' && x[1] === 'Sheet1'));
  assert.deepEqual(log.filter(x => x[1] === 'delete').map(x => x[0]), ['Sheet1 (2)!58:105', 'Sheet1 (2)!1:30']);
  assert.ok(log.some(x => x[1] === 'clear' && x[2] === 'Contents'), 'emptying the grid keeps its formatting');
  assert.ok(log.some(x => x[1] === 'dvClear'));
});

test('lookup tabs keep their formatting when rewritten', async () => {
  const { Excel, log } = mockExcel(['Slot Types']); const { OfficeStore } = loadApp(Excel);
  await OfficeStore().writeSheet('Slot Types', [['Slot type', 'Kind'], ['Admin', 'Blocked']], { keepFormat: true });
  assert.deepEqual(log.filter(x => x[1] === 'clear').map(x => x[2]), ['Contents']);
  const { Excel: E2, log: l2 } = mockExcel(); await loadApp(E2).OfficeStore().writeSheet('TA Summary', [['A']]);
  assert.deepEqual(l2.filter(x => x[1] === 'clear').map(x => x[2]), [undefined]);
});
