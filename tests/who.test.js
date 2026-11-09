// "Enter your name", the loading screen and the full-screen relay for both. Made-up staff names only.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'); const path = require('path'); const vm = require('vm');

function memStorage(seed){ const m = new Map(Object.entries(seed || {})); return { getItem: k => (m.has(k) ? m.get(k) : null), setItem: (k, v) => m.set(k, String(v)), removeItem: k => m.delete(k), _m: m }; }
function loadApp({ storage, docs } = {}){
  const els = {}, on = {};
  const el = () => ({ addEventListener(){}, appendChild(){}, remove(){}, focus(){}, querySelector: () => null, querySelectorAll: () => [], hidden: true, innerHTML: '', value: '', classList: { toggle(){} } });
  const later = (f, ms) => { const h = setTimeout(f, ms); if (h.unref) h.unref(); return h; };
  const sandbox = { console, setTimeout: later, clearTimeout, Office: undefined, window: { scrollTo(){} }, location: { href: 'https://x/taskpane.html', search: '' }, localStorage: storage || memStorage(),
    document: { querySelector: s => els[s] || (els[s] = el()), querySelectorAll: () => [], addEventListener(k, f){ (on[k] = on[k] || []).push(f); }, createElement: el, head: el() } };
  sandbox.TE = require('../engine.js'); sandbox.TC = require('../cadence.js'); sandbox.PayerRules = require('../payer-rules.js'); sandbox.SoharCore = require('../core.js'); sandbox.SoharRules = require('../rules.js'); sandbox.self = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8') + '\n;this.__api = { S, OfficeStore, RemoteStore, RELAY_OPS, reload, render, seBy, userList, withUser, pickOrder, cleanName, wbKeyOf, whoRec, whoToday, whoPick, dayIso, BOOT };', sandbox);
  const api = sandbox.__api, d = docs || {};
  const steps = [];
  api.S.store = { live: true, docs: d, load: async step => { for (let i = 0; i < 4; i++) if (step){ step(i, [{ tabs: 3 }, { scanned: 2 }, { lookups: [] }, { appt: null, se: false }][i]); steps.push(i); } return { names: [], grids: {}, cfg: {}, appt: null, hidden: [] }; },
    writeSheet: async () => {}, getDoc: k => d[k] ?? null, setDoc: async (k, v) => { d[k] = v; } };
  api.S.askWho = true;
  const fire = (k, e) => Promise.all((on[k] || []).map(f => f(e)));
  const click = (act, data = {}) => { const t = { dataset: Object.assign({ act }, data), tagName: 'BUTTON', closest: s => (s === '[data-act]' ? t : null) }; return fire('click', { target: t }); };
  return { api, els, docs: d, steps, sandbox, click };
}
const html = o => o.els['#app'].innerHTML;
// the app runs in its own vm realm: compare its arrays and objects as plain JSON
const eq = (a, b, m) => assert.deepEqual(JSON.parse(JSON.stringify(a)), b, m);

test('name helpers: clean, dedupe, add once, order by this computer, today only', () => {
  const { api } = loadApp();
  assert.equal(api.cleanName('  Ana   Ruiz  '), 'Ana Ruiz');
  assert.equal(api.cleanName('x'.repeat(40)).length, 30);
  eq(api.userList(['Ana', { name: 'ana ' }, '', null, { name: 'Ben', added: '2026-11-01' }]).map(u => u.name), ['Ana', 'Ben']);
  eq(api.userList('not a list'), []);
  const L = api.withUser([{ name: 'Ana', added: '2026-11-01' }], 'Cam', '2026-11-07');
  eq(JSON.parse(JSON.stringify(L)), [{ name: 'Ana', added: '2026-11-01' }, { name: 'Cam', added: '2026-11-07' }]);
  assert.equal(api.withUser(L, 'cam').length, 2, 'the same name in other case is not added twice');
  const o = api.pickOrder([{ name: 'Ana' }, { name: 'Ben' }, { name: 'Cam' }], ['Cam']);
  eq(o.map(u => [u.name, u.here]), [['Cam', true], ['Ana', false], ['Ben', false]]);
  const rec = api.whoPick(api.whoRec(''), 'Ben', '2026-11-07');
  assert.equal(api.whoToday(rec, '2026-11-07'), 'Ben');
  assert.equal(api.whoToday(rec, '2026-11-08'), '', 'a pick lasts the day');
  eq(api.whoPick(rec, 'Ana', '2026-11-08').used, ['Ana', 'Ben']);
  eq(api.whoRec('{bad'), { day: '', name: '', used: [] });
  assert.equal(api.wbKeyOf('https://x/Book.xlsx?web=1'), api.wbKeyOf('https://X/book.xlsx'));
  assert.equal(api.wbKeyOf(''), '');
});

test('first open: loading screen with steps, then "Enter your name"; adding a name saves it on the workbook and today on this computer', async () => {
  const o = loadApp();
  const seen = [];
  const orig = o.api.S.store.load; o.api.S.store.load = async step => { const r = await orig(step && ((i, x) => { step(i, x); seen.push(html(o)); })); return r; };
  await o.api.reload();
  eq(o.steps, [0, 1, 2, 3], 'the pane passes the step callback on first open');
  assert.ok(seen[0].includes('class="boot"') && seen[0].includes('class="spinner"') && seen[0].includes(o.api.BOOT[1]), 'loading screen with a spinner and steps');
  assert.ok(seen[0].includes('3 tabs'), 'a finished step says what it found');
  assert.equal(o.api.S.boot.i, 5, 'every step done');
  let h = html(o);
  assert.ok(h.includes('Enter your name') && h.includes('Nobody has added their name to this workbook yet'));
  assert.ok(!h.includes('data-act="pickUser"'));
  o.els['#whoNew'].value = '  Ana  ';
  await o.click('addUser');
  eq(o.docs.ctUsers.map(u => u.name), ['Ana'], 'stored in document settings');
  assert.ok(o.docs.ctWbId, 'a workbook without a URL gets an id for this computer’s memory');
  assert.equal(o.api.S.name, 'Ana');
  h = html(o);
  assert.ok(h.includes('data-act="switchUser"') && h.includes('<b>Ana</b>') && h.includes('Switch user'), 'name shows with Switch user');
  const key = [...o.sandbox.localStorage._m.keys()].find(k => k.startsWith('ctWho:'));
  assert.ok(key, 'today’s pick kept per workbook');
  assert.equal(JSON.parse(o.sandbox.localStorage.getItem(key)).day, o.api.dayIso());
  // a later reload is a refresh: no loading screen, no step callback
  o.steps.length = 0; await o.api.reload(); eq(o.steps, []);
  assert.ok(!html(o).includes('Enter your name'));
});

test('the same day reopens straight in; the next day asks again with this computer’s names first', async () => {
  const docs = { ctUsers: [{ name: 'Ben', added: '2026-11-01' }, { name: 'Ana', added: '2026-11-01' }], ctWbId: 'id-demo' };
  const storage = memStorage();
  const a = loadApp({ storage, docs }); await a.api.reload();
  assert.ok(html(a).includes('Pick yours below, or add it if you’re new.'));
  await a.click('pickUser', { v: 'Ben' });
  assert.equal(a.api.S.name, 'Ben');
  const b = loadApp({ storage, docs }); await b.api.reload();
  assert.equal(b.api.S.name, 'Ben'); assert.ok(!html(b).includes('Enter your name'));
  // yesterday's pick: ask again, nobody picked, Ben first and marked
  const key = [...storage._m.keys()].find(k => k.startsWith('ctWho:'));
  storage.setItem(key, JSON.stringify({ day: '2000-01-01', name: 'Ben', used: ['Ben'] }));
  const c = loadApp({ storage, docs }); await c.api.reload();
  assert.equal(c.api.S.name, '');
  const h = html(c);
  assert.ok(h.indexOf('data-v="Ben"') < h.indexOf('data-v="Ana"') && h.includes('On this computer'));
  // a pick whose name was taken off the list doesn't count
  storage.setItem(key, JSON.stringify({ day: c.api.dayIso(), name: 'Dee', used: ['Dee'] }));
  const d = loadApp({ storage, docs }); await d.api.reload(); assert.equal(d.api.S.name, '');
  // Switch user, then stay
  await c.click('pickUser', { v: 'Ana' }); await c.click('switchUser');
  assert.ok(html(c).includes('Stay as Ana'));
  await c.click('pickCancel'); assert.ok(!html(c).includes('Enter your name'));
});

test('Clinic Work By is the picked name; Node tests without Excel never ask', async () => {
  const o = loadApp(); o.api.S.askWho = false; await o.api.reload();
  assert.ok(!html(o).includes('Enter your name'));
  o.api.S.SE = { meta: { Location: 'San Rafael' } };
  assert.equal(o.api.seBy(), 'Clinic San Rafael');
  o.api.S.name = 'Ana'; assert.equal(o.api.seBy(), 'Ana');
});

test('full screen: addUser is relayed, the pane’s pick and Dev Update ids come down, picks go up', async () => {
  const o = loadApp();
  assert.ok(o.api.RELAY_OPS.includes('addUser'));
  const sent = []; let handler = null;
  o.sandbox.Office = { EventType: { DialogParentMessageReceived: 'p' }, context: { ui: { addHandlerAsync: (k, f) => { handler = f; }, messageParent: s => sent.push(JSON.parse(s)) } } };
  const rs = o.api.RemoteStore();
  assert.equal(typeof rs.addUser, 'function'); assert.equal(typeof rs.send, 'function');
  o.api.S.mode = 'dialog'; o.api.S.store = rs;
  const chunk = obj => JSON.stringify({ c: 'z' + Math.random(), i: 0, n: 1, d: JSON.stringify(obj) });
  const p = rs.load();
  const req = JSON.parse(sent[0].d); assert.equal(req.op, 'load');
  handler({ message: chunk({ rid: req.rid, ok: true, res: { who: 'Ana', docs: { ctUsers: [{ name: 'Ana' }] } } }) });
  const r = await p; assert.equal(r.who, 'Ana'); eq(rs.getDoc('ctUsers'), [{ name: 'Ana' }]);
  o.api.S.data = { names: [], grids: {}, cfg: {}, ap: null, templates: [], hidden: [], work: [] };
  handler({ message: chunk({ ev: 'who', name: 'Ben' }) }); assert.equal(o.api.S.name, 'Ben');
  let got = null; o.sandbox.window.DevUpdates = { receive: m => { got = m; return m.k === 'devupd'; } }; o.sandbox.DevUpdates = o.sandbox.window.DevUpdates;
  handler({ message: chunk({ k: 'devupd', seen: ['a'] }) }); eq(got.seen, ['a']);
  sent.length = 0; await o.click('pickUser', { v: 'Cam' });
  eq(JSON.parse(sent[0].d), { ev: 'who', name: 'Cam' }, 'a pick in full screen goes to the pane');
});

test('OfficeStore.addUser re-reads saved settings first, so a name added elsewhere stays', async () => {
  const o = loadApp();
  let mem = { ctUsers: [{ name: 'Ana', added: '2026-11-01' }] }; const saved = { ctUsers: [{ name: 'Ana', added: '2026-11-01' }, { name: 'Ben', added: '2026-11-02' }] };
  o.sandbox.Office = { context: { document: { settings: { get: k => mem[k], set: (k, v) => { mem[k] = v; }, saveAsync: cb => { Object.assign(saved, mem); cb(); }, refreshAsync: cb => { mem = JSON.parse(JSON.stringify(saved)); cb(); } } } } };
  const list = await o.api.OfficeStore().addUser('Cam');
  eq(list.map(u => u.name), ['Ana', 'Ben', 'Cam']);
  eq(saved.ctUsers.map(u => u.name), ['Ana', 'Ben', 'Cam']);
});

test('toolbar: refresh, full screen, light / dark; the theme is kept and passed both ways; the bell home is full screen only', async () => {
  const o = loadApp({ docs: { ctUsers: [{ name: 'Ana', added: '2026-11-01' }], ctWbId: 'id-t' } });
  o.api.S.canFull = true; await o.api.reload(); await o.click('pickUser', { v: 'Ana' });
  let h = html(o);
  const tools = h.slice(h.indexOf('<div class="tools">'));
  const acts = [...tools.matchAll(/<button class="iconbtn[^"]*" data-act="(\w+)" data-tip="([^"]+)"/g)].map(m => [m[1], m[2]]);
  eq(acts, [['refresh', 'Refresh'], ['full', 'Full screen'], ['theme', 'Dark mode']]);
  assert.ok(tools.includes('<span class="tcur"><svg') && tools.includes('<span class="tsl">/</span><span class="tgo"><svg'), 'whole sun and moon with the slash');
  assert.ok(!h.includes('du-home'), 'no bell home in the pane');
  const sent = []; o.api.S.mode = 'pane';
  await o.click('theme');
  assert.equal(o.api.S.theme, 'dark'); assert.equal(o.sandbox.localStorage.getItem('ctTheme'), 'dark');
  h = html(o); assert.ok(h.includes('data-tip="Light mode"') && h.includes('aria-pressed="true"') && / swap"/.test(h), 'swap animation on the change');
  // full screen: the home spot right of the name, and the theme goes up to the pane
  o.api.S.mode = 'dialog'; o.api.S.store.send = m => sent.push(m); o.api.S.store.close = () => {};
  o.api.render(); h = html(o);
  assert.ok(/<span class="bname">Clinic Template Tool<\/span><span class="du-home" aria-hidden="true"><\/span>/.test(h), 'du-home right of the name');
  assert.ok(h.includes('data-act="closeFull" data-tip="Exit full screen"'));
  await o.click('theme'); eq(sent, [{ ev: 'theme', theme: 'light' }]);
  // the loading and name screens have no home spot
  o.api.S.picking = true; o.api.render(); assert.ok(!html(o).includes('du-home'));
  const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  assert.ok(src.includes("displayDialogAsync(url, { height: 95, width: 95, displayInIframe: true }") && !/displayInIframe: false/.test(src), 'in-page dialog only');
  assert.ok(src.includes("'?mode=dialog&theme=' + S.theme"), 'full screen opens in the pane’s theme');
});

test('release: APP_VERSION, package.json and every ?v= agree', () => {
  const v = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8').match(/const APP_VERSION = '([\d.]+)'/)[1];
  assert.equal(v, '1.1.5');
  assert.equal(JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'package.json'), 'utf8')).version, v);
  const found = [...fs.readFileSync(path.join(__dirname, '..', 'taskpane.html'), 'utf8').matchAll(/\?v=([\d.]+)/g)].map(m => m[1]);
  assert.ok(found.length); eq([...new Set(found)], [v]);
});
