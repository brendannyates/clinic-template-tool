// Clinic Template Tool: Excel add-in front end. Analysis lives in engine.js (TE).
// Clinic Eligibility files use Sohar's own core.js / rules.js (SoharCore / SoharRules, byte-identical copies).
'use strict';
const APP_VERSION = '1.1.0';   // bump with every ?v= in taskpane.html
const SC = SoharCore, SR = SoharRules;
const $ = s => document.querySelector(s);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pct = (x, d = 0) => x == null || isNaN(x) ? '—' : (x * 100).toFixed(d) + '%';
const num = x => x == null || isNaN(x) ? '—' : Math.round(x).toLocaleString();
const days = x => x == null ? '—' : x + (x === 1 ? ' day' : ' days');

const SHEETS = { appt: 'Appointment List', ref: 'Ref', st: 'Slot Types', prov: 'Providers', hist: 'TA 3NA History', eskst: 'PT ESK Status', pcr: 'Payer Credential Rules' };
const OUT_PREFIX = 'TA ', PT_PREFIX = 'PT ';   // TA tabs are counts only; PT tabs list patients by LASFIR code
const VIEWS = [['overview', 'Overview'], ['templates', 'Templates'], ['access', 'Access'], ['cadence', 'Cadence'], ['funnel', 'Funnel'], ['recall', 'Recall'], ['visits', 'Visits'], ['impact', 'Provider impact'], ['trends', 'Trends'], ['eligibility', 'Eligibility'], ['benefits', 'Benefits'], ['payers', 'Payers'], ['setup', 'Setup']];
const VIEW_LABEL = Object.fromEntries(VIEWS);
// top tabs; a group with more than one visible view gets a sub-tab row
const NAV = [['schedule', 'Schedule', ['overview', 'templates', 'impact', 'trends', 'visits']], ['patients', 'Patients', ['cadence', 'funnel', 'recall']], ['coverage', 'Coverage', ['eligibility', 'benefits', 'payers', 'access']], ['setup', 'Setup', ['setup']]];
const SE_VIEWS = ['eligibility', 'benefits', 'payers'];   // need a Clinic Eligibility file; every other view but Setup needs the appointment list
const SE_WORK_TABLE = 'tblClinicWork';   // the Clinic Work table Sohar's clinicPack makes; the admin's Power Automate flow reads it
const AUTO_SHOW = 'Office.AutoShowTaskpaneWithDocument';
const S = { view: 'overview', data: null, R: null, f: { from: '', to: '', loc: '', prov: '', sl: '' }, opts: { lateHours: 24, target: 14, asOf: '', releaseHours: 24, convDays: 90, lapsedDays: 60, overdueMult: 1.5, tmsCount: 30, maintDays: 180, matchDays: 10, dayStart: '08:00', dayEnd: '18:00' }, heatMetric: 'fill', trendType: '', impactWhy: '', recallState: 'call', store: null, mode: 'pane', full: false,
  SE: null, E: null, sub: {}, sePill: '', seQ: '', seBen: '' };

// ---------- small UI helpers ----------
function toast(kind, html, ms = 5000){ const d = document.createElement('div'); d.className = 'msg ' + kind; d.innerHTML = html; $('#toast').appendChild(d); setTimeout(() => d.remove(), ms); }
function busy(on, msg){ $('#busy').hidden = !on; if (msg) $('#busyMsg').textContent = msg; }
// a control's data-act/k/m/v/p as a selector, so it can be found again after #app is re-drawn
function actSel(el){ const d = el && el.dataset; return d && d.act ? ['act', 'k', 'm', 'v', 'p'].filter(k => d[k] != null).map(k => `[data-${k}="${String(d[k]).replace(/["\\]/g, '\\$&')}"]`).join('') : ''; }
// the dialog takes keyboard focus (its first field, else first button, or `first`) and closing it puts focus back on the
// control that opened it, found again by actSel (a Save re-draws #app); otherwise focus falls to the top of the page
let modalFrom = '';
const FOCUSABLE = 'textarea, input, select, button';
function openModal(html, first){
  if ($('#modal').hidden) modalFrom = actSel(document.activeElement);
  $('#sheet').innerHTML = html; $('#modal').hidden = false;
  const el = (first && $('#sheet').querySelector(first)) || $('#sheet').querySelector(FOCUSABLE); if (el) el.focus();
}
function closeModal(){
  $('#modal').hidden = true; $('#sheet').innerHTML = ''; const r = askRes; askRes = null;
  const sel = modalFrom; modalFrom = ''; const el = sel && $('#app').querySelector(sel); if (el) el.focus();
  if (r) r(false);
}
// Office add-ins block window.confirm, so yes/no questions use the in-pane modal. Closing it any other way = Cancel.
let askRes = null;
function ask(html, ok = 'Continue'){
  return new Promise(res => {
    // focus starts on Cancel: these confirm a replace, so Enter alone should not do it
    openModal(`<h2>Please confirm</h2><div>${html}</div><div class="row" style="margin-top:12px"><button class="btn primary" data-act="askYes">${esc(ok)}</button><button class="btn" data-act="askNo">Cancel</button></div>`, '[data-act="askNo"]');
    askRes = res;
  });
}
$('#modal').addEventListener('click', e => { if (e.target.id === 'modal') closeModal(); });
// Tab and Shift+Tab stay inside an open dialog (aria-modal alone does not keep them there)
document.addEventListener('keydown', e => {
  if ($('#modal').hidden) return;
  if (e.key === 'Escape') return closeModal();
  if (e.key !== 'Tab') return;
  const f = Array.from($('#sheet').querySelectorAll(FOCUSABLE)).filter(x => !x.disabled); if (!f.length) return;
  const i = f.indexOf(document.activeElement), j = e.shiftKey ? (i <= 0 ? f.length - 1 : i - 1) : (i < 0 || i === f.length - 1 ? 0 : i + 1);
  e.preventDefault(); f[j].focus();
});

// ---------- workbook access ----------
function OfficeStore(){
  const CHUNK = 2500;
  async function readBig(ctx, ws, rows, cols, chunk = CHUNK){
    const out = [];
    for (let r = 0; r < rows; r += chunk){ const g = ws.getRangeByIndexes(r, 0, Math.min(chunk, rows - r), cols); g.load('values'); await ctx.sync(); out.push(...g.values); }
    return out;
  }
  async function sheetFor(ctx, name){ const ws = ctx.workbook.worksheets.getItemOrNullObject(name); await ctx.sync(); if (!ws.isNullObject) return { ws, fresh: false }; return { ws: ctx.workbook.worksheets.add(name), fresh: true }; }
  // Clinic Work: the sheet keeps the last write per Kind + Key, the admin applies the newest When. When the sheet already has
  // that Kind + Key at or after this change's When (another PC's clock runs ahead, or this write waited behind a slow one),
  // restamp the change just after it (max(now, that When + 1 ms)) so both sides pick the same winner.
  function lastWins(cur, ch){
    const id = r => `${String(r.Kind ?? '').trim().toLowerCase()}|${String(r.Key ?? '').trim().toLowerCase()}`, mine = id(ch);
    const prev = Math.max(-Infinity, ...SR.tableObjects(cur).filter(r => id(r) === mine).map(r => Date.parse(String(r.When ?? '').trim())).filter(n => !isNaN(n)));
    const at = Date.parse(String(ch.When ?? '').trim());
    return prev < (isNaN(at) ? Date.now() : at) ? ch : Object.assign({}, ch, { When: new Date(Math.max(Date.now(), prev + 1)).toISOString() });
  }
  // text: '@' on the range in the same batch, before the values, so Excel keeps '0090001' and '=> new card' as typed
  async function writeRows(ctx, ws, aoa, r0 = 0, text = false){
    const w = Math.max(1, ...aoa.map(r => r.length)); const rows = aoa.map(r => { const x = r.slice(); while (x.length < w) x.push(''); return x.map(v => v == null ? '' : text ? String(v) : v); });
    for (let i = 0; i < rows.length; i += 1000){ const part = rows.slice(i, i + 1000); const g = ws.getRangeByIndexes(r0 + i, 0, part.length, w); if (text) g.numberFormat = part.map(r => r.map(() => '@')); g.values = part; await ctx.sync(); }
    return w;
  }
  // fmt.table (Clinic Work): Power Automate's "List rows present in a table" sees only rows inside the table, so after the
  // in-place write the table grows to header + every row (never shrinks; header in row 1, as Sohar writes it). That also pulls in
  // rows an older write left below it. resize needs ExcelApi 1.13; else rows.add, which inserts and pushes what is below down, so
  // those copies are cleared after. Whatever fails, the rows already written stay on the sheet; the result says if the table has them.
  async function fitTable(ctx, ws, name, aoa, w){
    const t = ws.tables.getItemOrNullObject(name); await ctx.sync();
    if (t.isNullObject) return { table: 'missing' };
    const bounds = async () => { const r = t.getRange(); r.load('rowIndex,columnIndex,rowCount,columnCount'); await ctx.sync(); return r; };
    const b = await bounds();
    if (b.rowIndex !== 0 || b.columnIndex !== 0) return { table: 'short' };
    if (b.rowCount >= aoa.length && b.columnCount >= w) return { table: 'ok' };
    try { t.resize(ws.getRangeByIndexes(0, 0, Math.max(b.rowCount, aoa.length), Math.max(b.columnCount, w))); await ctx.sync(); return { table: 'ok' }; }
    catch (e){ console.warn('Table resize failed; adding rows instead', e); }
    if (b.rowCount >= aoa.length) return { table: 'short' };   // only columns were missing: rows.add can't add those
    try {
      const add = aoa.slice(b.rowCount).map(r => Array.from({ length: b.columnCount }, (_, j) => (r[j] == null ? '' : String(r[j]))));
      t.rows.add(null, add.map(r => r.map(() => '')));   // blank first: in a General cell '0090001' turns into a number and '=…' into a formula
      const g = ws.getRangeByIndexes(b.rowCount, 0, add.length, b.columnCount); g.numberFormat = add.map(r => r.map(() => '@')); g.values = add;
      await ctx.sync();
    } catch (e){ console.warn('Adding table rows failed', e); return { table: 'short' }; }
    try {
      const a = t.getRange(); a.load('rowIndex,columnIndex,rowCount,columnCount'); const u = ws.getUsedRangeOrNullObject(true); u.load('rowIndex,rowCount,columnIndex,columnCount'); await ctx.sync();
      const end = u.isNullObject ? 0 : u.rowIndex + u.rowCount, top = a.rowIndex + a.rowCount;
      if (end > top){ ws.getRangeByIndexes(top, 0, end - top, Math.max(a.columnCount, u.columnIndex + u.columnCount)).clear('Contents'); await ctx.sync(); }
    } catch (e){ console.warn('Clearing below the table failed', e); }   // copies left below the table are harmless (last row per Kind + Key wins)
    return { table: 'ok' };
  }
  const settings = () => Office.context.document.settings;
  return {
    live: true,
    load: () => Excel.run(async ctx => {
      const wss = ctx.workbook.worksheets; wss.load('items/name,items/visibility'); await ctx.sync();
      const info = wss.items.map(ws => { const u = ws.getUsedRangeOrNullObject(true); u.load('rowCount,columnCount,rowIndex,columnIndex'); return { ws, name: ws.name, hidden: ws.visibility !== 'Visible', u }; });
      await ctx.sync();
      const dims = x => x.u.isNullObject ? null : { rows: x.u.rowIndex + x.u.rowCount, cols: x.u.columnIndex + x.u.columnCount };
      const out = { names: info.map(x => x.name), grids: {}, cfg: {}, appt: null, hidden: info.filter(x => x.hidden).map(x => x.name) };
      // a Clinic Eligibility file (it has SE Info): its own sheets are data, never templates or the appointment list
      const seFile = info.some(x => x.name.toLowerCase() === SR.SE_SHEETS.info.toLowerCase());
      const seNames = Object.values(SR.SE_SHEETS).map(n => n.toLowerCase());
      const isSE = x => seFile && seNames.includes(x.name.toLowerCase());
      const apptX = info.find(x => !isSE(x) && !x.hidden && dims(x) && /^(appointment list|appt list)/i.test(x.name)) || info.find(x => !isSE(x) && dims(x) && /appointment|appt/i.test(x.name) && dims(x).rows > 50);
      const cfgKey = n => (Object.entries(SHEETS).find(([k, v]) => k !== 'appt' && v.toLowerCase() === n.toLowerCase()) || [])[0];
      // templates live on small visible sheets; hidden sheets and report tabs are never read
      const small = info.filter(x => x !== apptX && !isSE(x) && dims(x) && !x.hidden && !x.name.startsWith(OUT_PREFIX) && !x.name.startsWith(PT_PREFIX) && !cfgKey(x.name) && dims(x).rows <= 800 && dims(x).cols <= 80);
      for (let i = 0; i < small.length; i += 6){
        const batch = small.slice(i, i + 6).map(x => { const d = dims(x); const g = x.ws.getRangeByIndexes(0, 0, d.rows, d.cols); g.load('values'); return { x, g }; });
        await ctx.sync(); batch.forEach(b => { out.grids[b.x.name] = b.g.values; });
      }
      for (const x of info){ const k = cfgKey(x.name); const d = dims(x); if (k && d) out.cfg[k] = await readBig(ctx, x.ws, d.rows, Math.min(d.cols, 20)); }
      if (apptX){ const d = dims(apptX); out.appt = { sheet: apptX.name, values: await readBig(ctx, apptX.ws, d.rows, Math.min(d.cols, 60)) }; }
      // SE sheets are hidden and hold 1-3 KB JSON cells: every column, 500 rows a sync (Excel on the web caps a response at 5 MB)
      if (seFile){ out.se = {}; for (const x of info){ const d = dims(x); if (isSE(x) && d && x.name.toLowerCase() !== SR.SE_SHEETS.readMe.toLowerCase()) out.se[x.name] = await readBig(ctx, x.ws, d.rows, d.cols, 500); } }
      return out;
    }),
    // replace a sheet's contents (report and config tabs)
    // fmt.change (Clinic Work; a plain object with its own When, so full screen relays it as is): two people can work one clinic
    // file, so the change is applied (SoharRules.workRows, by header name) to the rows on the sheet now, read in this same run,
    // and aoa is not used. Writing the pane's own copy would drop rows the other person saved since this pane read the sheet.
    // The result then carries rows: the Clinic Work rows written.
    writeSheet: (name, aoa, fmt = {}) => Excel.run(async ctx => {
      const { ws, fresh } = await sheetFor(ctx, name);
      const keep = fmt.keepFormat && !fresh;     // lookup tabs a PM may have colored: replace values only
      let merged = null;
      const apply = cur => { merged = SR.workRows(cur, lastWins(cur, fmt.change)); return [SR.CLINIC_WORK_COLS].concat(merged.map(x => SR.CLINIC_WORK_COLS.map(c => x[c]))); };
      const done = res => (merged ? Object.assign({}, res, { rows: merged }) : res);
      // inPlace (Clinic Work): write over the old rows first and clear what is left below only after that, so a failed write keeps them
      if (fmt.inPlace && !fresh){
        const u = ws.getUsedRangeOrNullObject(true); u.load('rowIndex,rowCount,columnIndex,columnCount'); await ctx.sync();
        const rows = u.isNullObject ? 0 : u.rowIndex + u.rowCount, cols = u.isNullObject ? 0 : u.columnIndex + u.columnCount;
        if (fmt.change) aoa = apply(rows ? await readBig(ctx, ws, rows, cols, 500) : []);
        const w = await writeRows(ctx, ws, aoa, 0, fmt.text);
        if (rows > aoa.length) ws.getRangeByIndexes(aoa.length, 0, rows - aoa.length, Math.max(cols, w)).clear('Contents');
        if (cols > w && aoa.length) ws.getRangeByIndexes(0, w, Math.min(rows, aoa.length), cols - w).clear('Contents');
        await ctx.sync();
        return done(fmt.table ? await fitTable(ctx, ws, fmt.table, aoa, w).catch(e => { console.warn(e); return { table: 'short' }; }) : undefined);
      }
      if (fmt.change) aoa = apply([]);   // a new sheet has no rows yet
      ws.getRange().clear(keep ? 'Contents' : undefined); await ctx.sync();
      const w = await writeRows(ctx, ws, aoa, 0, fmt.text);
      if (keep){ await ctx.sync(); return done(undefined); }
      const hdr = ws.getRangeByIndexes(0, 0, 1, w); hdr.format.font.bold = true; hdr.format.fill.color = '#DCEFED';
      for (const [c, f] of Object.entries(fmt.numFmt || {})) if (aoa.length > 1) ws.getRangeByIndexes(1, +c, aoa.length - 1, 1).numberFormat = Array.from({ length: aoa.length - 1 }, () => [f]);
      ws.freezePanes.freezeRows(1);
      try { ws.getRangeByIndexes(0, 0, Math.min(aoa.length, 400), w).format.autofitColumns(); } catch {}
      await ctx.sync();
      return done(fmt.table ? { table: 'missing' } : undefined);   // a sheet made here has no table
    }),
    replaceAppts: (sheetName, aoa, dateCol) => Excel.run(async ctx => {
      const { ws } = await sheetFor(ctx, sheetName || SHEETS.appt); ws.getRange().clear(); await ctx.sync();
      await writeRows(ctx, ws, aoa);
      if (dateCol != null && aoa.length > 1) ws.getRangeByIndexes(0, dateCol, aoa.length, 1).numberFormat = aoa.map(() => ['m/d/yyyy']);
      await ctx.sync();
    }),
    createTemplate: spec => Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.add(spec.sheet);
      await writeRows(ctx, ws, spec.cells);
      const n = spec.times.length, last = 6 + n;     // grid rows 7..last (1-based)
      ws.getRange(`A7:A${last}`).numberFormat = spec.times.map(() => ['h:mm AM/PM']);
      ws.getRange('B3').numberFormat = [['m/d/yyyy']]; ws.getRange('E3').numberFormat = [['m/d/yyyy']]; ws.getRange('F4').numberFormat = [['m/d/yyyy']];
      const g = ws.getRange(`B7:F${last}`);
      ws.getRange('I5').formulas = [[`=IFERROR(SORT(UNIQUE(TOCOL(B7:F${last},1))),"")`]];
      ws.getRange('J5').formulas = [[`=IF(I5#="","",COUNTIF($B$7:$F$${last},I5#))`]];
      ws.getRange('K5').formulas = [[`=IF(I5#="","",COUNTIF($B$7:$F$${last},I5#)*${spec.step}/60)`]];
      ['B', 'C', 'D', 'E', 'F'].forEach((c, i) => { ws.getRange(String.fromCharCode(76 + i) + '5').formulas = [[`=IF($I5#="","",COUNTIF(${c}$7:${c}$${last},$I5#))`]]; });
      ws.getRange('I2').values = [['Clinical slots / week']];
      ws.getRange('J2').formulas = [[`=SUMPRODUCT(COUNTIF(B7:F${last},FILTER('${SHEETS.st}'!$A$2:$A$300,'${SHEETS.st}'!$B$2:$B$300="Clinical","~")))`]];
      ws.getRange('A2').format.font.size = 14; ws.getRange('A2').format.font.bold = true;
      ['A3', 'D3', 'A4', 'E4', 'A5', 'A6:F6', 'I2', 'I4:P4'].forEach(a => { ws.getRange(a).format.font.bold = true; });
      ws.getRange('A6:F6').format.fill.color = '#DCEFED'; ws.getRange('I4:P4').format.fill.color = '#DCEFED';
      const b = g.format.borders; ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideVertical', 'InsideHorizontal'].forEach(k => { b.getItem(k).style = 'Continuous'; b.getItem(k).color = '#C9CFC7'; });
      ws.getRange('A:A').format.columnWidth = 70; ws.getRange('B:F').format.columnWidth = 150; ws.getRange('I:I').format.columnWidth = 150;
      try { g.dataValidation.rule = { list: { inCellDropDown: true, source: `='${SHEETS.st}'!$A$2:$A$300` } }; g.dataValidation.errorAlert = { showAlert: false }; } catch {}
      ws.activate(); await ctx.sync();
    }),
    // duplicate a PM's own template tab so its colors, conditional formats, dropdowns, merges and formulas carry over,
    // then drop the other grids on it. Writes use the source tab's coordinates; rows are deleted last, bottom up.
    copyTemplate: spec => Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(spec.src).copy('End'); ws.name = spec.sheet; await ctx.sync();
      for (const [r, c, v, f] of spec.writes){ const g = ws.getRangeByIndexes(r, c, 1, 1); g.values = [[v]]; if (f) g.numberFormat = [[f]]; }
      if (spec.clearSlots) ws.getRangeByIndexes(...spec.slots).clear('Contents');
      await ctx.sync();
      for (const [r0, r1] of spec.deleteRows) ws.getRange(`${r0 + 1}:${r1 + 1}`).delete('Up');
      await ctx.sync();
      // a dropdown that pointed at another grid's list now reads #REF; point it at the Slot Types list instead
      const sl = ws.getRangeByIndexes(...spec.slotsAfter); sl.dataValidation.load('rule'); await ctx.sync();
      const list = sl.dataValidation.rule && sl.dataValidation.rule.list;
      if (!list || /#REF/i.test(String(list.source))){ try { sl.dataValidation.clear(); sl.dataValidation.rule = { list: { inCellDropDown: true, source: spec.dvSource } }; sl.dataValidation.errorAlert = { showAlert: false }; } catch {} }
      ws.activate(); ws.getRangeByIndexes(spec.slotsAfter[0], spec.slotsAfter[1], 1, 1).select(); await ctx.sync();
    }),
    getDoc: k => { try { return settings().get(k); } catch { return null; } },
    setDoc: (k, v) => new Promise(res => { try { settings().set(k, v); settings().saveAsync(() => res()); } catch { res(); } }),
  };
}

// Full screen: an in-page Office dialog that relays its workbook calls through the pane.
const CH = 12000;
function sendChunks(fn, obj){ const s = JSON.stringify(obj); const id = Math.random().toString(36).slice(2); const n = Math.max(1, Math.ceil(s.length / CH)); for (let i = 0; i < n; i++) fn(JSON.stringify({ c: id, i, n, d: s.slice(i * CH, (i + 1) * CH) })); }
function makeRx(handler){ const buf = {}; return str => { let m; try { m = JSON.parse(str); } catch { return; } if (!m || !m.c) return; const b = buf[m.c] || (buf[m.c] = []); b[m.i] = m.d; if (b.filter(x => x != null).length === m.n){ delete buf[m.c]; try { handler(JSON.parse(b.join(''))); } catch (e){ console.error(e); } } }; }
const RELAY_OPS = ['load', 'writeSheet', 'replaceAppts', 'createTemplate', 'copyTemplate', 'setDoc'];
function RemoteStore(){
  const pend = new Map(); let docs = {};
  const rx = makeRx(m => { if (m.rid && pend.has(m.rid)){ const p = pend.get(m.rid); pend.delete(m.rid); m.ok ? p.res(m.res) : p.rej(new Error(m.err)); } });
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, a => rx(a.message));
  const up = obj => sendChunks(s => Office.context.ui.messageParent(s), obj);
  const call = (op, args) => new Promise((res, rej) => { const rid = Math.random().toString(36).slice(2); pend.set(rid, { res, rej }); up({ rid, op, args }); setTimeout(() => { if (pend.has(rid)){ pend.delete(rid); rej(new Error('Excel didn’t answer. Keep the add-in panel open behind full screen, then try again.')); } }, 120000); });
  return {
    live: true, remote: true,
    async load(){ const d = await call('load'); docs = d.docs || {}; return d; },
    writeSheet: (...a) => call('writeSheet', a), replaceAppts: (...a) => call('replaceAppts', a), createTemplate: (...a) => call('createTemplate', a), copyTemplate: (...a) => call('copyTemplate', a),
    getDoc: k => docs[k] ?? null, async setDoc(k, v){ docs[k] = v; await call('setDoc', [k, v]); },
    close(){ up({ ev: 'close' }); },
  };
}
let DLG = null;
function openFull(){
  if (DLG || S.mode !== 'pane') return;
  const url = location.href.replace(/[?#].*$/, '') + '?mode=dialog';
  Office.context.ui.displayDialogAsync(url, { height: 96, width: 96, displayInIframe: true }, res => {
    if (res.status !== Office.AsyncResultStatus.Succeeded){ toast('bad', 'Couldn’t open full screen: ' + esc(res.error && res.error.message)); return; }
    DLG = res.value; S.full = true; render();
    const reply = obj => { try { sendChunks(s => DLG && DLG.messageChild(s), obj); } catch (e){ console.error(e); } };
    const rx = makeRx(async m => {
      if (m.ev === 'close'){ try { DLG.close(); } catch {} DLG = null; S.full = false; reload(); return; }
      const { rid, op, args } = m;
      try {
        if (!RELAY_OPS.includes(op)) throw new Error('Unknown request');
        let r = await S.store[op](...(args || []));
        if (op === 'load') r = Object.assign({}, r, { docs: { ctOpts: S.store.getDoc('ctOpts'), ctEsk: S.store.getDoc('ctEsk'), ctEskAt: S.store.getDoc('ctEskAt'), ctFilters: S.store.getDoc('ctFilters'), [AUTO_SHOW]: S.store.getDoc(AUTO_SHOW) } });
        reply({ rid, ok: true, res: r ?? null });
      } catch (e){ reply({ rid, ok: false, err: e.message || String(e) }); }
    });
    DLG.addEventHandler(Office.EventType.DialogMessageReceived, a => rx(a.message));
    DLG.addEventHandler(Office.EventType.DialogEventReceived, () => { DLG = null; S.full = false; reload(); });
  });
}

// ---------- reading the workbook into the model ----------
function cfgRows(values, spec){
  if (!values || !values.length) return [];
  let hr = values.findIndex(r => r.some(v => Object.values(spec).some(re => re.test(String(v).trim())))); if (hr < 0) return [];
  const h = values[hr].map(v => String(v).trim()); const idx = {};
  for (const [k, re] of Object.entries(spec)){ const i = h.findIndex(x => re.test(x)); if (i >= 0) idx[k] = i; }
  return values.slice(hr + 1).filter(r => r.some(v => String(v).trim())).map(r => Object.fromEntries(Object.entries(idx).map(([k, i]) => [k, r[i]])));
}
const SPEC = {
  ref: { event: /^event$/i, resType: /^resource/i, sched: /schedule\s*class/i, cls: /^event\s*class/i, sl: /^service\s*line/i, elig: /eligible/i, role: /^role$/i },
  st: { type: /^slot\s*type$|^type$/i, kind: /^kind$/i, accepts: /^accepts?/i, notes: /^notes?$/i },
  prov: { name: /^nextgen|^provider$/i, label: /template\s*label|^label$/i, cred: /^cred/i, include: /^include/i },
  hist: { date: /^snapshot/i, prov: /^provider$/i, type: /^slot\s*type$/i, next: /^next/i, third: /^3na/i, open: /^open/i },
};
// Clinic Eligibility file: central Sohar's export for one clinic, unpacked with Sohar's own rules.js. Clinic Work = the clinic's edits.
function seUnpack(raw){
  if (!raw.se) return { se: null, seErr: null, work: [] };
  const pick = n => { const k = Object.keys(raw.se).find(x => x.toLowerCase() === n.toLowerCase()); return k ? raw.se[k] : null; };
  const work = SR.tableObjects(pick(SR.SE_SHEETS.work) || []);
  try { return { se: SR.clinicUnpack(raw.se), seErr: null, work }; }
  catch (e){ console.error(e); return { se: null, seErr: e.message || String(e), work }; }
}
function prepare(raw){
  const templates = [];
  for (const [name, values] of Object.entries(raw.grids || {})) templates.push(...TE.parseTemplates(name, values));
  const ap = raw.appt ? TE.parseAppts(raw.appt.values) : null;
  if (ap && raw.appt) ap.sheet = raw.appt.sheet;
  const se = seUnpack(raw);
  return {
    se: se.se, seErr: se.seErr, work: se.work,
    names: raw.names || [], hidden: raw.hidden || [], grids: raw.grids || {}, templates, ap,
    ref: cfgRows(raw.cfg.ref, SPEC.ref), refExists: !!raw.cfg.ref,
    st: cfgRows(raw.cfg.st, SPEC.st), stExists: !!raw.cfg.st,
    prov: cfgRows(raw.cfg.prov, SPEC.prov).map(p => ({ name: String(p.name || '').trim(), label: String(p.label || '').trim(), cred: String(p.cred || '').trim(), include: p.include })), provExists: !!raw.cfg.prov,
    eskStatus: raw.cfg.eskst ? TC.readEskStatus(raw.cfg.eskst) : null,
    payerRules: raw.cfg.pcr ? PayerRules.read(raw.cfg.pcr) : null,
    hist: cfgRows(raw.cfg.hist, SPEC.hist).map(h => ({ date: TE.parseDate(h.date), prov: String(h.prov || ''), type: String(h.type || ''), next: h.next === '' ? null : +h.next, third: h.third === '' ? null : +h.third, open: +h.open || 0 })),
  };
}
const stRows = () => S.data.st.length ? S.data.st : TE.DEFAULT_SLOT_TYPES.map(r => ({ type: r[0], kind: r[1], accepts: r[2] }));
function asOfSerial(){ const o = TE.parseDate(S.opts.asOf); if (o != null) return o; const r = S.data && S.data.ap && S.data.ap.runAt; return r != null ? Math.floor(r) : TE.todaySerial(); }
// cadence and recall settings, with the defaults from the spec
function copts(){
  const o = S.opts, tm = (v, d) => { const x = TE.parseTime(v); return x == null ? d : x; };
  return { releaseHours: +o.releaseHours || 24, convDays: +o.convDays || 90, lapsedDays: +o.lapsedDays || 60, overdueMult: +o.overdueMult || 1.5, tmsCount: +o.tmsCount || 30,
    maintDays: +o.maintDays || 180, matchDays: +o.matchDays || 10, dayStart: tm(o.dayStart, 480), dayEnd: tm(o.dayEnd, 1080) };
}
function run(){
  const d = S.data; if (!d || !d.ap || d.ap.error){ S.R = null; S.C = null; return; }
  const f = S.f; const P = s => TE.parseDate(s); const co = copts();
  S.R = TE.analyze({
    templates: d.templates.map(t => Object.assign({}, t)), appts: d.ap.rows, ref: d.ref, slotTypes: stRows(), providers: d.prov,
    opts: { asOf: asOfSerial(), from: P(f.from), to: P(f.to), lateHours: +S.opts.lateHours || 24, provs: f.prov ? new Set([f.prov]) : null, locs: f.loc ? new Set([f.loc]) : null, sls: f.sl ? new Set([f.sl]) : null,
      releaseHours: co.releaseHours, dayStart: co.dayStart, dayEnd: co.dayEnd },
  });
  runCadence(co);
  runPayerRules();
}
// ESK visits against the Payer Credential Rules sheet (needs the sheet and a payer column on the report)
function runPayerRules(){
  const d = S.data, R = S.R, pr = d.payerRules; S.P = null;
  if (!pr || pr.error || !pr.rules.length || !d.ap.has.payer) return;
  const one = x => x ? new Set([x]) : null;
  S.P = TC.eskCredentials(d.ap.rows, d.ref, pr.rules, PayerRules, n => PayerRules.credOf(credOf(n)), new Set(R.people), { asOf: R.asOf, from: R.from, to: R.to, provs: one(S.f.prov), locs: one(S.f.loc) });
}
// patient-level views need a patient column; everything else works without one
function runCadence(co){
  const d = S.data, R = S.R; S.C = null; S.Cerr = null;
  if (!d.ap.has.pt && !d.ap.has.mrn) return;
  try {
    const one = x => x ? new Set([x]) : null;
    const o = Object.assign({}, co, { asOf: R.asOf, from: R.from, to: R.to, provs: one(S.f.prov), sls: one(S.f.sl), locs: one(S.f.loc) });
    const v = TC.visits(d.ap.rows, d.ref, o, new Set(R.people));
    const cad = TC.cadence(v.list, o);
    const esk = d.eskStatus && !d.eskStatus.error ? d.eskStatus : null;
    const rc = TC.recall(v.list, o, cad, esk);
    const now = d.ap.runAt != null && !S.opts.asOf ? d.ap.runAt : R.asOf + co.dayStart / 1440;   // report run time, else the start of the as-of day
    S.C = { o, v, cad, rc, now, esk, fun: TC.funnel(v.list, R, o), flex: TC.flexUsage(R, v.list, o), rel: TC.releaseList(R, o, now), matches: TC.slotMatches(rc, R, v.list, o, now) };
  } catch (e){ console.error(e); S.Cerr = e.message || String(e); }
}
// the clinic's view at this moment (fresh re-verify days, passed visits drop out); Clinic Work overrides ticks and notes
function runSE(){
  S.E = null; S.Eerr = null; if (!S.SE) return;
  try { S.E = SR.clinicDerive(S.SE, S.data.work, new Date()); } catch (e){ console.error(e); S.Eerr = e.message || String(e); }
}
async function reload(){
  busy(true, 'Reading the workbook…');
  try {
    const raw = await S.store.load();
    S.data = prepare(raw);
    S.opts = Object.assign(S.opts, S.store.getDoc('ctOpts') || {});
    S.esk = S.store.getDoc('ctEsk') || null;
    S.eskAt = S.store.getDoc('ctEskAt') || null;
    const first = !S.SE && S.data.se; S.SE = S.data.se;
    run(); runSE(); S.err = null;
    if (first) S.view = 'eligibility';
    // first open of a clinic file: save the setting that opens the pane with the document every time
    if (S.SE && S.mode === 'pane' && !S.store.getDoc(AUTO_SHOW)) Promise.resolve(S.store.setDoc(AUTO_SHOW, true)).catch(() => {});
  } catch (e){ console.error(e); S.err = e.message || String(e); }
  finally { busy(false); render(); }
}

// ---------- charts ----------
const CHARTS = {};
function lineChart(id, title, pts, fmt, yMaxFixed){
  const W = 360, H = 150, L = 38, R = 10, T = 8, B = 22;
  const ys = pts.map(p => p.y).filter(y => y != null);
  if (pts.length < 2 || !ys.length) return `<div class="chart"><div class="ttl">${esc(title)}</div><div class="small muted">Not enough weeks yet.</div></div>`;
  const yMax = yMaxFixed != null ? yMaxFixed : Math.max(...ys) * 1.1 || 1;
  const x = i => L + (W - L - R) * i / (pts.length - 1), y = v => T + (H - T - B) * (1 - v / yMax);
  let d = '', pen = false; pts.forEach((p, i) => { if (p.y == null){ pen = false; return; } d += (pen ? 'L' : 'M') + x(i).toFixed(1) + ' ' + y(p.y).toFixed(1); pen = true; });
  const grid = [0, .5, 1].map(f => `<line x1="${L}" x2="${W - R}" y1="${y(yMax * f)}" y2="${y(yMax * f)}" stroke="var(--line)" stroke-width="1"/><text x="${L - 6}" y="${y(yMax * f) + 4}" text-anchor="end" font-size="10" fill="var(--muted)">${esc(fmt(yMax * f))}</text>`).join('');
  CHARTS[id] = { pts, fmt, x: pts.map((p, i) => x(i)), y: pts.map(p => p.y == null ? null : y(p.y)), W, H, T, B };
  return `<div class="chart"><div class="ttl">${esc(title)}</div><svg viewBox="0 0 ${W} ${H}" data-chart="${id}" role="img" aria-label="${esc(title)} by week">${grid}
    <path d="${d}" fill="none" stroke="var(--s1)" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>
    <line class="xh" x1="0" x2="0" y1="${T}" y2="${H - B}" stroke="var(--muted)" stroke-width="1" opacity="0"/><circle class="dotp" r="4.5" fill="var(--s1)" stroke="var(--surface)" stroke-width="2" opacity="0"/>
    <text x="${L}" y="${H - 6}" font-size="10" fill="var(--muted)">${esc(pts[0].x)}</text><text x="${W - R}" y="${H - 6}" font-size="10" text-anchor="end" fill="var(--muted)">${esc(pts[pts.length - 1].x)}</text>
    <rect x="${L}" y="${T}" width="${W - L - R}" height="${H - T - B}" fill="transparent"/></svg></div>`;
}
document.addEventListener('mousemove', e => {
  const svg = e.target.closest && e.target.closest('svg[data-chart]'); const tip = $('#tip');
  document.querySelectorAll('svg[data-chart] .xh, svg[data-chart] .dotp').forEach(n => { if (!svg || !svg.contains(n)) n.setAttribute('opacity', 0); });
  if (!svg){ if (!e.target.closest || !e.target.closest('[data-tip]')) tip.hidden = true; else { tip.textContent = e.target.closest('[data-tip]').dataset.tip; tip.hidden = false; tip.style.left = (e.clientX + 12) + 'px'; tip.style.top = (e.clientY + 12) + 'px'; } return; }
  const c = CHARTS[svg.dataset.chart]; if (!c) return;
  const r = svg.getBoundingClientRect(); const vx = (e.clientX - r.left) * c.W / r.width;
  let best = 0; c.x.forEach((xx, i) => { if (Math.abs(xx - vx) < Math.abs(c.x[best] - vx)) best = i; });
  const p = c.pts[best]; svg.querySelector('.xh').setAttribute('x1', c.x[best]); svg.querySelector('.xh').setAttribute('x2', c.x[best]); svg.querySelector('.xh').setAttribute('opacity', .5);
  const dot = svg.querySelector('.dotp'); if (c.y[best] != null){ dot.setAttribute('cx', c.x[best]); dot.setAttribute('cy', c.y[best]); dot.setAttribute('opacity', 1); } else dot.setAttribute('opacity', 0);
  tip.innerHTML = `<b>Week of ${esc(p.x)}</b><br>${esc(p.y == null ? 'No data' : c.fmt(p.y))}${p.sub ? '<br><span class="muted">' + esc(p.sub) + '</span>' : ''}`; tip.hidden = false; tip.style.left = (e.clientX + 12) + 'px'; tip.style.top = (e.clientY + 12) + 'px';
});
const SEQ = ['var(--q1)', 'var(--q2)', 'var(--q3)', 'var(--q4)', 'var(--q5)', 'var(--q6)', 'var(--q7)'];
const seq = f => f == null ? 'transparent' : SEQ[Math.max(0, Math.min(6, Math.floor(f * 7 - 1e-9)))];
const seqInk = f => f != null && f > 0.57 ? '#fff' : 'var(--ink)';
const barCell = (v, max, label) => `<td class="n bar-cell"><i style="width:${max ? Math.max(0, Math.min(100, 100 * v / max)) : 0}%"></i><span>${label}</span></td>`;

// ---------- views ----------
function credOf(n){ const p = (S.data.prov || []).find(x => x.name === n); return (p && p.cred) || TE.credentialOf(n); }
const target = () => +S.opts.target || 14;
function threeChip(v){ if (v == null) return '<span class="chip mute">None open</span>'; const t = target(); return `<span class="chip ${v <= t / 2 ? 'ok' : v <= t ? 'warn' : 'bad'}">${days(v)}</span>`; }

function filtersHTML(){
  const R = S.R, d = S.data; if (!R) return '';
  const locs = [...new Set(d.ap.rows.map(a => a.loc).filter(Boolean))].sort();
  const provs = R.people.slice().sort((a, b) => (R.tByProv.has(b) - R.tByProv.has(a)) || a.localeCompare(b));
  const ref = TE.buildRef(d.ref);
  const sls = [...new Set([...new Set(d.ap.rows.map(a => a.event))].map(ev => TE.classify(ref, ev).sl).flatMap(s => String(s).split(',').map(x => x.trim())).filter(x => x && x !== 'N/A'))].sort();
  const opt = (v, cur, label) => `<option value="${esc(v)}"${v === cur ? ' selected' : ''}>${esc(label || v)}</option>`;
  return `<div class="filters">
    <label>From<input type="date" data-f="from" value="${esc(S.f.from || TE.iso(R.from))}"></label>
    <label>To<input type="date" data-f="to" value="${esc(S.f.to || TE.iso(R.to))}"></label>
    <label>Location<select data-f="loc">${opt('', S.f.loc, 'All locations')}${locs.map(l => opt(l, S.f.loc)).join('')}</select></label>
    <label>Provider<select data-f="prov">${opt('', S.f.prov, 'All providers')}${provs.map(p => opt(p, S.f.prov, p + (R.tByProv.has(p) ? '' : ' (no template)'))).join('')}</select></label>
    <label>Service line<select data-f="sl">${opt('', S.f.sl, 'All service lines')}${sls.map(s => opt(s, S.f.sl)).join('')}</select></label>
    <span class="small muted" style="padding-bottom:6px">As of ${esc(TE.mdy(R.asOf))}${S.data.ap.runAt != null && !S.opts.asOf ? ' (report run time)' : ''}. Weeks run Sun–Sat.</span>
  </div>`;
}

function insights(R){
  const out = [];
  const typeRows = R.byType.filter(t => t.supply >= 40);
  for (const t of typeRows.filter(t => t.fit != null && t.fit < 0.5 && t.kept >= 20).sort((a, b) => a.fit - b.fit).slice(0, 2)){
    const mix = Object.entries(R.offMix[t.key] || {}).sort((a, b) => b[1] - a[1])[0];
    out.push(`<b>${esc(t.key)}</b> slots: only ${pct(t.fit)} of kept visits were the intended type${mix ? `; most often used for <b>${esc(mix[0])}</b> (${num(mix[1])})` : ''}. Either the slot type or the booking rules need a look.`);
  }
  for (const t of typeRows.filter(t => t.fill != null && t.fill < 0.25).sort((a, b) => a.fill - b.fill).slice(0, 2)) out.push(`<b>${esc(t.key)}</b> slots are rarely used: ${pct(t.fill)} filled across ${num(t.supply)} slots. Candidates to convert.`);
  const hours = R.byHour.filter(h => typeof h.key === 'number' && h.aKept + h.aNS >= 40 && h.nsRate != null).sort((a, b) => b.nsRate - a.nsRate);
  if (hours[0] && R.total.nsRate != null && hours[0].nsRate > R.total.nsRate * 1.4) out.push(`No-shows peak at <b>${TE.hourLabel(hours[0].key)}</b> (${pct(hours[0].nsRate, 1)} vs ${pct(R.total.nsRate, 1)} overall).`);
  const dw = R.byDow.filter(x => x.supply >= 40 && x.fill != null).sort((a, b) => a.fill - b.fill);
  if (dw.length > 1 && dw[dw.length - 1].fill - dw[0].fill > 0.1) out.push(`<b>${esc(dw[0].key)}</b> fills least (${pct(dw[0].fill)}); <b>${esc(dw[dw.length - 1].key)}</b> fills most (${pct(dw[dw.length - 1].fill)}).`);
  const imp = R.byProv.filter(p => p.inBlocked + p.offTemplate + p.dbl > 0).sort((a, b) => (b.inBlocked + b.offTemplate + b.dbl) - (a.inBlocked + a.offTemplate + a.dbl))[0];
  if (imp) out.push(`<b>${esc(imp.key)}</b> has the most visits outside clinical template time: ${num(imp.inBlocked)} in admin/lunch/meeting, ${num(imp.offTemplate)} outside template hours or days,${num(imp.dbl)} double-booked (${num(imp.overMin / 60)} h).`);
  const slow = R.threeNa.filter(x => x.prov !== 'All providers' && x.third != null && x.third > target()).sort((a, b) => b.third - a.third).slice(0, 3);
  if (slow.length) out.push(`3NA over the ${target()}-day target: ${slow.map(x => `${esc(x.type)} with ${esc(x.prov.split(',')[0])} (${days(x.third)})`).join('; ')}.`);
  if (!R.hasLate) out.push(`Late cancels can’t be measured yet: the report has no cancel date. Add <b>Cancel Date/Time</b> (and <b>Created Date</b> for lead time and 3NA history) to the NextGen report.`);
  return out;
}
function provTable(R){
  const rows = R.byProv.filter(p => p.supply || p.aAll).sort((a, b) => (R.tByProv.has(b.key) - R.tByProv.has(a.key)) || b.supply - a.supply || b.aAll - a.aAll);
  return `<div class="tw"><table><thead><tr><th>Provider</th><th>Cred</th><th class="n">Clinical slots</th><th class="n">Filled</th><th class="n">Kept use</th><th class="n">Template fit</th><th class="n">Open</th><th class="n">To admin</th><th class="n">PTO</th><th class="n">Visits</th><th class="n">No-show</th><th class="n">Cancel</th><th class="n">Late cx</th><th class="n">Outside template</th></tr></thead><tbody>
  ${rows.map(p => `<tr><td>${esc(p.key)}${R.tByProv.has(p.key) ? '' : ' <span class="chip mute">No template</span>'}</td><td>${esc(credOf(p.key))}</td><td class="n">${p.supply ? num(p.supply) : '—'}</td>${p.supply ? barCell(p.fill, 1, pct(p.fill)) : '<td class="n">—</td>'}<td class="n">${pct(p.util)}</td><td class="n">${pct(p.fit)}</td><td class="n">${p.supply ? num(p.open) : '—'}</td><td class="n">${p.supply ? num(p.admin) : '—'}</td><td class="n">${p.supply ? num(p.pto) : '—'}</td><td class="n">${num(p.aAll)}</td><td class="n">${pct(p.nsRate, 1)}</td><td class="n">${pct(p.cxRate, 1)}</td><td class="n">${R.hasLate ? pct(p.lcRate, 1) : '—'}</td><td class="n">${num(p.inBlocked + p.offTemplate)}</td></tr>`).join('')}
  </tbody></table></div>`;
}
function vOverview(){
  const R = S.R, t = R.total;
  const k = (l, v, s) => `<div class="kpi"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`;
  const any = R.threeNa.filter(x => x.prov === 'All providers' && x.third != null).sort((a, b) => b.third - a.third)[0];
  return `<div class="kpis">
    ${k('Clinical template slots', num(t.supply), `${TE.mdy(R.from)} – ${TE.mdy(R.to)}`)}
    ${k('Filled', pct(t.fill), 'slots with a kept or no-show visit')}
    ${k('Template fit', pct(t.fit), 'kept visits that matched the slot type')}
    ${k('No-show rate', pct(t.nsRate, 1), `${num(t.aNS)} of ${num(t.aKept + t.aNS)} visits`)}
    ${k('Cancel rate', pct(t.cxRate, 1), 'all cancels, incl. reschedules')}
    ${k('Late cancel rate', R.hasLate ? pct(t.lcRate, 1) : '—', R.hasLate ? `within ${S.opts.lateHours} h of the visit` : 'needs cancel date on the report')}
    ${k('Booked ahead', pct(t.futFill), 'of template slots after the as-of date')}
    ${k('Longest 3NA', any ? days(any.third) : '—', any ? esc(any.type) + ', all providers' : 'no open slots found')}
    ${k('Visits outside template', num(t.inBlocked + t.offTemplate), `${num(t.dbl)} double-booked`)}
  </div>
  <div class="card"><h2>What stands out</h2><ul class="ins">${insights(R).map(x => `<li>${x}</li>`).join('') || '<li class="muted">Nothing unusual in this range.</li>'}</ul></div>
  <div class="card"><h2>By provider</h2>${provTable(R)}<div class="small muted" style="margin-top:6px">Filled and template fit use past template slots only. Visit rates count every visit booked on the provider, with or without a template.</div></div>`;
}
function vTemplates(){
  const R = S.R; const st = TE.buildSlotTypes(stRows());
  const rows = R.byType.filter(t => t.supply || t.futSupply).sort((a, b) => b.supply - a.supply);
  const maxS = Math.max(1, ...rows.map(r => r.supply));
  const mixOf = k => { const e = Object.entries(R.offMix[k] || {}).sort((a, b) => b[1] - a[1]).slice(0, 2); return e.map(x => `${esc(x[0])} (${x[1]})`).join(', '); };
  const heatM = S.heatMetric; const metricOf = c => c ? ({ fill: c.fill, util: c.util, ns: c.nsRate, cx: c.cxRate, kept: c.aKept }[heatM]) : null;
  const H = new Map(R.heat.map(h => [h.key, h])); const hours = [...new Set(R.heat.filter(h => h.supply || h.futSupply).map(h => +String(h.key).split('|')[1]))].sort((a, b) => a - b);
  const vals = R.heat.map(metricOf).filter(v => v != null); const maxV = heatM === 'kept' ? Math.max(1, ...vals) : (heatM === 'ns' || heatM === 'cx' ? Math.max(0.01, ...vals) : 1);
  const fmtV = v => v == null ? '' : heatM === 'kept' ? num(v) : pct(v, heatM === 'ns' ? 1 : 0);
  const heat = `<div class="tw"><table class="heat"><thead><tr><th></th>${hours.map(h => `<th class="n">${TE.hourLabel(h)}</th>`).join('')}</tr></thead><tbody>${[1, 2, 3, 4, 5].map(d => `<tr><td>${TE.DAYS[d].slice(0, 3)}</td>${hours.map(h => { const c = H.get(d + '|' + h); const v = metricOf(c); const f = v == null ? null : v / maxV; return `<td class="h" style="background:${seq(f)};color:${seqInk(f)}" data-tip="${esc(`${TE.DAYS[d]} ${TE.hourLabel(h)}: ${c ? `${num(c.supply)} slots, ${pct(c.fill)} filled, ${num(c.aKept)} kept, no-show ${pct(c.nsRate, 1)}` : 'no slots'}`)}">${fmtV(v)}</td>`; }).join('')}</tr>`).join('')}</tbody></table></div>
    <div class="legend">Low <i style="background:linear-gradient(90deg,var(--q1),var(--q4),var(--q7))"></i> High${heatM === 'ns' || heatM === 'cx' ? ` (scale tops out at ${pct(maxV, 1)})` : ''}</div>`;
  const tplRows = S.data.templates.map(t => ({ t, prov: (R.templates.find(x => x.sheet === t.sheet && x.anchor[0] === t.anchor[0] && x.anchor[1] === t.anchor[1]) || {}).provider }));
  const wk = t => { let c = 0, b = 0; t.slots.forEach(s => { TE.slotTypeOf(st, s.type).kind === 'clinical' ? c++ : b++; }); return { c, b }; };
  return `<div class="card"><h2>Slot types</h2><div class="tw"><table><thead><tr><th>Slot type</th><th class="n">Slots</th><th class="n">Filled</th><th class="n">Template fit</th><th class="n">Kept as intended</th><th class="n">Kept, other visit</th><th class="n">Release bookings</th><th class="n">No-show</th><th class="n">Open</th><th class="n">To admin</th><th class="n">Booked ahead</th><th>Most common other visits</th></tr></thead><tbody>
    ${rows.map(r => `<tr><td>${esc(r.key)}${st.get(String(r.key).toLowerCase()) ? '' : ' <span class="chip warn" data-tip="Not on the Slot Types tab; using a guess">?</span>'}</td>${barCell(r.supply, maxS, num(r.supply))}<td class="n">${pct(r.fill)}</td><td class="n">${pct(r.fit)}</td><td class="n">${num(r.keptFit)}</td><td class="n">${num(r.keptOff)}</td><td class="n">${num(r.keptRel)}</td><td class="n">${num(r.ns)}</td><td class="n">${num(r.open)}</td><td class="n">${num(r.admin)}</td><td class="n">${pct(r.futFill)}</td><td class="small">${mixOf(r.key)}</td></tr>`).join('')}
    </tbody></table></div><div class="small muted" style="margin-top:6px">Template fit uses the Accepts column on the Slot Types tab. A release booking (made under ${copts().releaseHours} business hours before the visit) may take any slot, so it doesn’t count against fit. "To admin" = a clinical slot taken by an admin or meeting event in NextGen. PTO slots are left out of the slot counts.</div></div>
  ${flexCard()}
  <div class="card"><h2>Day of week × time of day <span class="sp"></span><select data-act="heatMetric">${[['fill', 'Filled %'], ['util', 'Kept use %'], ['ns', 'No-show %'], ['cx', 'Cancel %'], ['kept', 'Kept visits']].map(([v, l]) => `<option value="${v}"${v === heatM ? ' selected' : ''}>${l}</option>`).join('')}</select></h2>${heat}</div>
  <div class="cols"><div class="card"><h2>Morning vs afternoon</h2><table><thead><tr><th></th><th class="n">Slots</th><th class="n">Filled</th><th class="n">Fit</th><th class="n">No-show</th><th class="n">Cancel</th></tr></thead><tbody>${R.byPart.filter(p => p.key !== 'No time').map(p => `<tr><td>${p.key}</td><td class="n">${num(p.supply)}</td><td class="n">${pct(p.fill)}</td><td class="n">${pct(p.fit)}</td><td class="n">${pct(p.nsRate, 1)}</td><td class="n">${pct(p.cxRate, 1)}</td></tr>`).join('')}</tbody></table></div>
  <div class="card"><h2>Templates in this workbook</h2><div class="tw"><table><thead><tr><th>Tab</th><th>Label</th><th>Linked provider</th><th>Location</th><th>Applies from</th><th class="n">Clinical / wk</th><th class="n">Blocked / wk</th></tr></thead><tbody>${tplRows.map(({ t, prov }) => { const w = wk(t); return `<tr><td>${esc(t.sheet)}</td><td>${esc(t.label)}</td><td>${prov ? esc(prov) : '<span class="chip bad">Not linked</span>'}</td><td>${esc(t.location)}</td><td>${t.effFrom != null ? TE.mdy(t.effFrom) : '<span class="muted">All dates</span>'}</td><td class="n">${w.c}</td><td class="n">${w.b}</td></tr>`; }).join('') || '<tr><td colspan="7" class="muted">No template grids found. Use Setup → New template tab.</td></tr>'}</tbody></table></div>
  <div class="small muted" style="margin-top:6px">A template with no Effective From date applies to every date. Add one when a provider’s template changes, and keep the old tab so earlier weeks are measured against the template that was in place.</div></div></div>`;
}
function vAccess(){
  const R = S.R;
  const all = R.threeNa.filter(x => x.prov === 'All providers'); const per = R.threeNa.filter(x => x.prov !== 'All providers');
  const tbl = rows => `<table><thead><tr><th>Provider</th><th>Slot type</th><th class="n">Next available</th><th class="n">3NA</th><th class="n">Open slots ahead</th><th>Third open slot</th></tr></thead><tbody>${rows.map(x => `<tr><td>${esc(x.prov)}</td><td>${esc(x.type)}</td><td class="n">${days(x.next)}</td><td class="n">${threeChip(x.third)}</td><td class="n">${num(x.open)}</td><td>${x.thirdDate != null ? TE.mdy(x.thirdDate) : x.full ? '<span class="muted">Fully booked to end of report</span>' : ''}</td></tr>`).join('')}</tbody></table>`;
  const fut = R.byWeek.filter(w => w.futSupply).slice(0, 12);
  // 3NA trend: saved snapshots, or rebuilt weekly when the report has created dates
  const types = [...new Set([...(S.data.hist || []).map(h => h.type), ...((R.threeNaHist || []).map(h => h.type))])].sort();
  const tt = S.trendType && types.includes(S.trendType) ? S.trendType : (types.includes('Any clinical slot') ? 'Any clinical slot' : types[0]);
  let trend = '';
  if (R.threeNaHist && tt){ const pts = R.threeNaHist.filter(h => h.type === tt).map(h => ({ x: TE.mdy(h.week), y: h.third })); trend = lineChart('t3na', `3NA, ${tt} (rebuilt from created dates)`, pts, v => Math.round(v) + ' d'); }
  else if (tt){ const byD = new Map(); (S.data.hist || []).filter(h => h.type === tt && h.prov === 'All providers').forEach(h => byD.set(h.date, h.third)); const pts = [...byD.entries()].sort((a, b) => a[0] - b[0]).map(([d, v]) => ({ x: TE.mdy(d), y: v })); trend = lineChart('t3na', `3NA, ${tt} (saved snapshots)`, pts, v => Math.round(v) + ' d'); }
  const esk = (S.esk ? eskCard() : '') + payerCard();
  return `<div class="card"><h2>Third next available (3NA)<span class="sp"></span><button class="btn sm" data-act="snapshot">Save 3NA snapshot</button></h2>
    <div class="small muted">Calendar days from ${TE.mdy(R.asOf)} to the third open template slot of each type. Open = no Kept, Expected or No Show visit on the report, and no PTO or admin booked. Target: ${target()} days (change in Setup).</div>
    <h3>All providers</h3><div class="tw">${tbl(all)}</div><h3>By provider</h3><div class="tw">${tbl(per)}</div>
    <div class="small muted" style="margin-top:6px">Each slot counts once. A visit that needs two slots in a row (for example a 60-minute consult) may wait longer than this shows.</div></div>
  <div class="cols"><div class="card"><h2>Booked ahead by week</h2><table><thead><tr><th>Week of</th><th class="n">Template slots</th><th class="n">Booked</th><th class="n">Booked %</th></tr></thead><tbody>${fut.map(w => `<tr><td>${TE.mdy(w.key)}</td><td class="n">${num(w.futSupply)}</td><td class="n">${num(w.futBooked)}</td>${barCell(w.futFill, 1, pct(w.futFill))}</tr>`).join('') || '<tr><td colspan="4" class="muted">No future weeks on the report.</td></tr>'}</tbody></table></div>
  <div class="card"><h2>3NA over time<span class="sp"></span>${types.length ? `<select data-act="trendType">${types.map(t => `<option${t === tt ? ' selected' : ''}>${esc(t)}</option>`).join('')}</select>` : ''}</h2>${trend || `<div class="small muted">${R.hasBooked ? '' : 'Save a snapshot each week (it’s stored on the TA 3NA History tab), or add Created Date to the NextGen report so the tool can rebuild past weeks.'}</div>`}</div></div>${esk}`;
}
function eskCard(){
  const m = S.esk, R = S.R; const st = TE.buildSlotTypes(stRows());
  const eskSlots = prov => { const list = R.tByProv.get(prov) || []; const t = list[list.length - 1]; if (!t) return null; return t.slots.filter(s => { const x = TE.slotTypeOf(st, s.type); return !x.any && TE.accepts(x, { cls: 'EFU', event: 'Esketamine Follow Up', sl: 'ESK' }); }).length; };   // dedicated ESK slots; Open Schedule would overstate it
  const rows = (m.providers || []).map(p => { const nm = TE.matchProvider(String(p.Provider || ''), R.people) || String(p.Provider || ''); return { p, nm, ours: eskSlots(nm) }; });
  return `<div class="card"><h2>${esc(m.modality || 'ESK')} access model <span class="chip mute">${esc(m.clinic || '')} · ${esc(String(m.generated || '').slice(0, 10))}</span></h2>
    <div class="tw"><table><thead><tr><th>Provider (tracker)</th><th>Matched to</th><th>Cred</th><th class="n">Patients</th><th class="n">Need / wk</th><th class="n">Tracker template slots / wk</th><th class="n">ESK slots / wk in this template</th><th class="n">Booked / wk</th></tr></thead><tbody>
    ${rows.map(({ p, nm, ours }) => `<tr><td>${esc(p.Provider)}</td><td>${esc(nm)}</td><td>${esc(p.Credential)}</td><td class="n">${num(+p.Patients)}</td><td class="n">${esc(p['Need per week'])}</td><td class="n">${esc(p['Template slots per week'])}</td><td class="n">${ours == null ? '—' : ours}</td><td class="n">${esc(p['Booked per week'])}</td></tr>`).join('')}
    </tbody></table></div><div class="small muted" style="margin-top:6px">From the ESK Tracker’s Access export. Need per week is the panel’s expected visits; compare it with the slots this template workbook gives ESK.</div></div>`;
}
// ---------- cadence, funnel, recall ----------
const ROLE_LABEL = { consult: 'Consult', start: 'Start', treatment: 'Treatment', followup: 'Follow-up', admin: 'Admin' };
const STATE = { overdue: ['warn', 'Overdue'], lapsed: ['bad', 'Lapsed'], unaccounted: ['bad', 'Unaccounted'], maintenance: ['warn', 'Maintenance due'], active: ['ok', 'Booked'], current: ['mute', 'Not due'], completed: ['mute', 'Completed'], paused: ['mute', 'Paused'], discontinued: ['mute', 'Discontinued'] };
const stateChip = s => `<span class="chip ${(STATE[s] || ['mute'])[0]}">${esc((STATE[s] || [0, s])[1])}</span>`;
const code = c => c ? esc(c) : '<span class="muted" data-tip="Pat Name isn’t in “Last, First” form, so no LASFIR code">(check name)</span>';
const gapD = x => x == null ? '—' : Math.round(x) + ' d';
function needPatients(){
  if (S.Cerr) return `<div class="card"><div class="msg bad">Couldn’t work out cadence: ${esc(S.Cerr)}</div></div>`;
  return `<div class="card"><h2>Needs a patient column</h2><div class="msg info">Cadence, funnel and recall follow each patient from visit to visit. Add <b>Pat Name</b> (and <b>Md Rc</b>, so two patients with the same name stay apart) to the NextGen report, then import it again. The tool shows and writes LASFIR codes only.</div></div>`;
}
function vCadence(){
  const C = S.C; if (!C) return needPatients();
  const rows = C.cad.rows.filter(r => r.patients); const H = C.cad.hist;
  const hc = (n, tot) => { const f = tot ? n / tot : null; return `<td class="h" style="background:${seq(n ? f : null)};color:${seqInk(f)}" data-tip="${esc(`${n} of ${tot} returns`)}">${tot && n ? pct(f) : ''}</td>`; };
  const pts = C.cad.patients;
  return `<div class="card"><h2>How often patients come back</h2>
    <div class="tw"><table class="heat"><thead><tr><th>Provider</th><th>Service line</th><th class="n">Patients</th><th class="n">Kept visits</th><th class="n">Visits / patient</th><th class="n">Returning</th><th class="n">Mean gap</th><th class="n">Median gap</th><th class="n">Usual return</th><th class="n">No-show</th><th class="n">Tele</th>${H.map(h => `<th class="n">${esc(h)}</th>`).join('')}</tr></thead><tbody>
    ${rows.map(r => `<tr${r.prov === 'All providers' ? ' class="tot"' : ''}><td>${esc(r.prov)}</td><td>${esc(r.bucket)}</td><td class="n">${num(r.patients)}</td><td class="n">${num(r.kept)}</td><td class="n">${r.perPt == null ? '—' : r.perPt.toFixed(1)}</td><td class="n">${num(r.returning)}</td><td class="n">${gapD(r.mean)}</td><td class="n">${gapD(r.median)}</td><td class="n">${r.modeWk == null ? '—' : `${r.modeWk === 1 ? '≤1' : r.modeWk} wk <span class="muted">(${pct(r.modeShare)})</span>`}</td><td class="n">${pct(r.nsRate, 1)}</td><td class="n">${pct(r.telePct)}</td>${r.hist.map(n => hc(n, r.n)).join('')}</tr>`).join('') || '<tr><td colspan="19" class="muted">No kept visits in this range.</td></tr>'}
    </tbody></table></div>
    <div class="small muted" style="margin-top:6px">Kept visits only; same-day bookings count once. Gaps are days between a patient’s kept visits in the same service line, credited to the provider they last saw in it (TMS and ESK treatments on chairs count toward that provider). Usual return is the most common gap in whole weeks, with the share of gaps that fall in it. Consults start a chain but aren’t part of it (see Funnel). No-show rate = no-shows ÷ (kept + no-shows).</div></div>
  <div class="card"><h2>By patient <span class="chip mute">${num(pts.length)}</span></h2>
    <div class="tw" style="max-height:480px"><table><thead><tr><th>Patient</th><th>Provider</th><th>Service line</th><th class="n">Kept</th><th>First seen</th><th>Last seen</th><th class="n">Mean gap</th><th class="n">Last gap</th><th>Modality</th></tr></thead><tbody>
    ${pts.slice(0, 400).map(p => `<tr><td>${code(p.code)}</td><td>${esc(p.prov)}</td><td>${esc(p.bucket)}</td><td class="n">${num(p.kept)}</td><td>${TE.mdy(p.first)}</td><td>${TE.mdy(p.last)}</td><td class="n">${gapD(p.meanGap)}</td><td class="n">${gapD(p.lastGap)}</td><td class="small">${p.telePct == null ? '' : `Tele ${pct(p.telePct)} / In person ${pct(1 - p.telePct)}`}</td></tr>`).join('')}
    </tbody></table></div>${pts.length > 400 ? `<div class="small muted">Showing 400 of ${num(pts.length)}. Write patient tabs for the full list.</div>` : ''}</div>`;
}
function vFunnel(){
  const C = S.C; if (!C) return needPatients();
  const rows = C.fun; const co = C.o;
  const flag = r => r.flag === 'over' ? '<span class="chip warn">More than needed</span>' : r.flag === 'under' ? '<span class="chip bad">Short</span>' : r.flag === 'ok' ? '<span class="chip ok">About right</span>' : '';
  const h = x => x == null ? '—' : x.toFixed(1);
  return `<div class="card"><h2>Consult to start</h2>
    <div class="tw"><table><thead><tr><th>Provider</th><th>Service line</th><th class="n">Consult slots</th><th class="n">Consults kept</th><th class="n">Consult fill</th><th class="n">Started</th><th class="n">Pending</th><th class="n">Conversion</th><th class="n">Median days to start</th><th class="n">Start hours offered</th><th class="n">Start hours used</th><th class="n">Start hours needed</th><th></th></tr></thead><tbody>
    ${rows.map(r => `<tr><td>${esc(r.prov)}</td><td>${esc(r.sl)}</td><td class="n">${num(r.consultSlots)}</td><td class="n">${num(r.consults)}</td><td class="n">${pct(r.consultFill)}</td><td class="n">${num(r.converted)}</td><td class="n">${num(r.pending)}</td><td class="n">${pct(r.rate)}</td><td class="n">${r.medDays == null ? '—' : days(Math.round(r.medDays))}</td><td class="n">${h(r.startHoursOffered)}</td><td class="n">${h(r.startHoursUsed)}</td><td class="n">${h(r.startHoursNeeded)}</td><td>${flag(r)}</td></tr>`).join('') || '<tr><td colspan="13" class="muted">No kept consults in this range.</td></tr>'}
    </tbody></table></div>
    <div class="small muted" style="margin-top:6px">A consult converts when the same patient has a kept start (Provider MT, ESK start and similar) in that service line within ${co.convDays} days, with any provider. Consults from the last ${co.convDays} days that haven’t started yet are pending and left out of the rate. Start hours needed = kept consults × conversion × start visit length; size start time from this, not from the consult hours offered. Slots come from this provider’s template; Open Schedule and other flex slots count toward both consults and starts.</div></div>`;
}
function flexCard(){
  const C = S.C; if (!C || !C.flex.length) return '';
  const roles = ['consult', 'start', 'treatment', 'followup'];
  return `<div class="card"><h2>Flex slots: what actually fills them</h2><div class="tw"><table><thead><tr><th>Provider</th><th>Slot type</th><th>Accepts</th><th class="n">Offered</th><th class="n">Used</th><th class="n">Use</th>${roles.map(r => `<th class="n">${ROLE_LABEL[r]}</th>`).join('')}<th class="n">Release bookings</th></tr></thead><tbody>
    ${C.flex.map(r => `<tr><td>${esc(r.prov)}</td><td>${esc(r.type)}</td><td class="small">${r.accepts.map(x => ROLE_LABEL[x]).join(', ') || 'Anything'}</td><td class="n">${num(r.offered)}</td><td class="n">${num(r.used)}</td><td class="n">${pct(r.useRate)}</td>${roles.map(x => `<td class="n">${num(r.roles[x] || 0)}</td>`).join('')}<td class="n">${num(r.release)} <span class="muted">${r.used ? '(' + pct(r.releaseShare) + ')' : ''}</span></td></tr>`).join('')}
    </tbody></table></div><div class="small muted" style="margin-top:6px">Flex = a slot type whose Accepts list takes two or more kinds of visit. If a consult/start flex pool is mostly starts, it has quietly become start overflow.</div></div>`;
}
function vRecall(){
  const C = S.C; if (!C) return needPatients();
  const counts = {}; C.rc.forEach(r => { counts[r.state] = (counts[r.state] || 0) + 1; });
  const call = C.rc.filter(r => TC.CALL.has(r.state));
  const pick = S.recallState === 'call' ? call : S.recallState ? C.rc.filter(r => r.state === S.recallState) : C.rc;
  const opt = (v, l) => `<option value="${v}"${v === S.recallState ? ' selected' : ''}>${l}</option>`;
  const eskAge = S.eskAt != null ? TE.todaySerial() - S.eskAt : null;
  const esk = S.data.eskStatus;
  const eskLine = esk && esk.error ? `<span class="chip bad">${esc(esk.error)}</span>` : esk ? `<span class="chip ${eskAge != null && eskAge > 7 ? 'warn' : 'mute'}">ESK Tracker: ${num(esk.rows.length)} patients${S.eskAt != null ? ', imported ' + TE.mdy(S.eskAt) : ''}</span>` : '<span class="chip mute">No ESK Tracker status loaded</span>';
  const t = (d, m) => `${TE.DAYS[TE.dow(d)].slice(0, 3)} ${TE.mdy(d)} ${TE.hhmm(m)}`;
  const m = C.matches;
  return `<div class="card"><h2>Recall <span class="sp"></span><select data-act="recallState">${opt('call', `Call list (${num(call.length)})`)}${opt('', `Everyone (${num(C.rc.length)})`)}${Object.keys(STATE).filter(k => counts[k]).map(k => opt(k, `${STATE[k][1]} (${num(counts[k])})`)).join('')}</select>
      <button class="btn sm primary" data-act="copyOutreach">Copy outreach list</button><label class="btn sm">Import ESK status…<input type="file" accept=".csv,.xlsx,.xls" data-act="importEskStatus" hidden></label><button class="btn sm" data-act="writePatients">Write patient tabs</button></h2>
    <div class="row" style="margin-bottom:6px">${Object.keys(STATE).filter(k => counts[k]).map(k => `<span class="small">${stateChip(k)} ${num(counts[k])}</span>`).join('')}${eskLine}</div>
    <div class="tw" style="max-height:520px"><table><thead><tr><th>Patient</th><th>Provider</th><th>Service line</th><th>Clinic</th><th>Last seen</th><th class="n">Days since</th><th class="n">Usual interval</th><th>State</th><th>Since then</th><th>Next booked</th><th>ESK Tracker</th><th>Return date</th><th>Pause reason</th></tr></thead><tbody>
    ${pick.slice(0, 500).map(r => `<tr><td>${code(r.code)}</td><td>${esc(r.prov)}</td><td>${esc(r.bucket)}</td><td>${esc(r.loc)}</td><td>${TE.mdy(r.last)}</td><td class="n">${num(r.daysSince)}</td><td class="n">${r.exp == null ? '—' : r.exp + ' d'}</td><td>${stateChip(r.state)}</td><td>${r.post ? `<span class="chip warn">${esc(r.post)}</span>` : ''}</td><td>${r.next != null ? TE.mdy(r.next) : ''}</td><td class="small">${esc(r.tracker)}</td><td>${r.ret != null ? TE.mdy(r.ret) : ''}</td><td class="small">${esc(r.reason)}</td></tr>`).join('') || '<tr><td colspan="13" class="muted">Nobody in this group.</td></tr>'}
    </tbody></table></div>${pick.length > 500 ? `<div class="small muted">Showing 500 of ${num(pick.length)}. Write patient tabs for the full list.</div>` : ''}
    <div class="small muted" style="margin-top:6px">As of ${TE.mdy(S.R.asOf)}. Overdue = more than ${C.o.overdueMult}× the service line’s usual return interval since the last kept visit, with nothing booked. Lapsed = over ${C.o.lapsedDays} days. TMS counts as completed after ${C.o.tmsCount} treatments and a follow-up, and comes back as maintenance due after ${C.o.maintDays} days. Paused and discontinued come from the ESK Tracker import; an overdue ESK patient the tracker doesn’t list is unaccounted. “Since then” flags a no-show or cancel after the last kept visit.</div></div>
  <div class="cols"><div class="card"><h2>Open slots for the call list <span class="chip mute">next ${C.o.matchDays} business days</span></h2>
    <div class="tw" style="max-height:420px"><table><thead><tr><th>Patient</th><th>Service line</th><th>Slot with</th><th>When</th><th>Slot type</th><th>Why it fits</th></tr></thead><tbody>
    ${m.slice(0, 300).map(x => `<tr><td>${code(x.code)}</td><td>${esc(x.bucket)}</td><td>${esc(x.prov)}${x.prov === x.patientProv ? '' : ' <span class="chip mute">other provider</span>'}</td><td>${t(x.date, x.min)}</td><td>${esc(x.type)}</td><td>${x.why === 'released' ? '<span class="chip ok">Released</span>' : x.why === 'flex' ? '<span class="chip mute">Flex</span>' : 'Takes this visit'}</td></tr>`).join('') || '<tr><td colspan="6" class="muted">No open template slots match.</td></tr>'}
    </tbody></table></div><div class="small muted" style="margin-top:6px">Up to three per patient: their own provider first, then the soonest. Released = inside the ${C.o.releaseHours}-business-hour window, open to any visit.</div></div>
  <div class="card"><h2>Opening for any visit <span class="chip mute">release window</span></h2>
    <div class="tw" style="max-height:420px"><table><thead><tr><th>Provider</th><th>Slot</th><th>Slot type</th><th>Status</th></tr></thead><tbody>
    ${C.rel.map(x => `<tr><td>${esc(x.prov)}</td><td>${t(x.date, x.min)}</td><td>${esc(x.type)}</td><td>${x.released ? '<span class="chip ok">Open to any visit now</span>' : `Opens ${t(Math.floor(x.enters), Math.round((x.enters % 1) * 1440))}`}</td></tr>`).join('') || '<tr><td colspan="4" class="muted">No open slots in or near the window.</td></tr>'}
    </tbody></table></div><div class="small muted" style="margin-top:6px">Open template slots less than ${C.o.releaseHours} business hours away (Mon–Fri, ${TE.hhmm(C.o.dayStart)}–${TE.hhmm(C.o.dayEnd)}), plus the ones that cross into the window by the end of the next business day. Counted from ${t(Math.floor(C.now), Math.round((C.now % 1) * 1440))}.</div></div></div>`;
}

function payerCard(){
  const d = S.data, pr = d.payerRules;
  if (!pr) return '';
  if (pr.error) return `<div class="card"><h2>Esketamine payer rules</h2><div class="msg bad">${esc(pr.error)}</div></div>`;
  if (!d.ap.has.payer) return `<div class="card"><h2>Esketamine payer rules</h2><div class="msg info">Add <b>Payer Name</b> to the NextGen report so each ESK visit can be checked against the Payer Credential Rules tab.</div></div>`;
  const P = S.P; if (!P) return '';
  const t = r => `${TE.mdy(r.date)} ${TE.hhmm(r.start)}`;
  const B = { approved: ['ok', 'Approved'], 'not approved': ['warn', 'Not approved'], pending: ['mute', 'Pending'], unknown: ['mute', '—'] };
  const vis = rows => `<div class="tw" style="max-height:360px"><table><thead><tr><th>Visit</th><th>Patient</th><th>Provider</th><th>Cred</th><th>Event</th><th>Payer</th><th>Problem</th></tr></thead><tbody>${rows.slice(0, 300).map(r => `<tr><td>${t(r)}</td><td>${r.code ? esc(r.code) : ''}</td><td>${esc(r.prov)}</td><td>${esc(r.cred)}</td><td>${esc(r.event)}</td><td>${esc(r.payer)}</td><td class="small">${esc(r.why)}</td></tr>`).join('')}</tbody></table></div>`;
  return `<div class="card"><h2>Esketamine payer rules <span class="chip mute">${num(P.rows.length)} ESK visits checked</span></h2>
    ${P.upcoming.length ? `<h3>Upcoming visits the payer won’t pay this provider for (${num(P.upcoming.length)})</h3>${vis(P.upcoming)}<div class="small muted">Move these to a provider the payer allows, or confirm with billing.</div>` : '<div class="msg ok">Every upcoming ESK visit is booked with a credential the payer allows (or the payer isn’t on the list).</div>'}
    ${P.unclear.length ? `<div class="small muted" style="margin-top:6px">${num(P.unclear.length)} upcoming visits match a payer with different rules by state (for example Anthem CA and WA) and the clinic name doesn’t say which. Add the state to the clinic name, or check those by hand.</div>` : ''}
    ${P.past.length ? `<h3>Past visits in the range outside the payer’s rules (${num(P.past.length)})</h3>${vis(P.past)}` : ''}
    <h3>By payer</h3><div class="tw" style="max-height:360px"><table><thead><tr><th>Payer on the report</th><th>Rule</th><th class="n">ESK visits</th><th class="n">With an MD</th><th class="n">With an NP</th><th>NP may bill</th><th>Besse</th><th class="n">Outside the rules</th></tr></thead><tbody>
    ${P.byPayer.map(g => `<tr><td>${esc(g.payer)}</td><td>${g.rule ? esc(g.ambiguous ? g.rule + ' (state?)' : g.rule) : '<span class="muted">Not on the list</span>'}</td><td class="n">${num(g.visits)}</td><td class="n">${num(g.md)}</td><td class="n">${num(g.np)}</td><td class="small">${esc(g.npSteps || '')}</td><td>${g.rule ? `<span class="chip ${B[g.besse][0]}">${B[g.besse][1]}</span>` : ''}</td><td class="n">${g.blocked ? `<span class="chip bad">${num(g.blocked)}</span>` : '0'}</td></tr>`).join('')}
    </tbody></table></div>
    <div class="small muted" style="margin-top:6px">From the Payer Credential Rules tab (Setup). ESK visits on providers only; chairs carry no credential. The NextGen list can’t tell a first treatment from later dosing, so dosing visits are checked as observation unless the event name says “first”. A blank Besse cell shows as pending.</div></div>`;
}

function rateTable(rows, label, showLead){
  const R = S.R; const max = Math.max(1, ...rows.map(r => r.aAll));
  return `<div class="tw"><table><thead><tr><th>${label}</th><th class="n">Visits</th><th class="n">Kept</th><th class="n">No-show</th><th class="n">Cancel</th><th class="n">Late cx</th>${showLead ? '<th class="n">Median lead</th>' : ''}</tr></thead><tbody>
  ${rows.map(r => `<tr><td>${esc(r.key)}</td>${barCell(r.aAll, max, num(r.aAll))}<td class="n">${num(r.aKept)}</td><td class="n">${pct(r.nsRate, 1)}</td><td class="n">${pct(r.cxRate, 1)}</td><td class="n">${R.hasLate ? pct(r.lcRate, 1) : '—'}</td>${showLead ? `<td class="n">${r.leadMed == null ? '—' : days(Math.round(r.leadMed))}</td>` : ''}</tr>`).join('')}</tbody></table></div>`;
}
function vVisits(){
  const R = S.R, lead = R.hasBooked;
  const srt = a => a.filter(r => r.aAll).sort((x, y) => y.aAll - x.aAll);
  const hrs = R.byHour.filter(h => typeof h.key === 'number' && h.aAll).sort((a, b) => a.key - b.key).map(h => Object.assign({}, h, { key: TE.hourLabel(h.key) }));
  const dws = [1, 2, 3, 4, 5, 6, 0].map(d => R.byDow.find(x => x.key === TE.DAYS[d])).filter(x => x && x.aAll);
  return `<div class="cols"><div class="card"><h2>By service line</h2>${rateTable(srt(R.bySL), 'Service line', lead)}</div><div class="card"><h2>By visit class</h2>${rateTable(srt(R.byClass), 'Event class', lead)}</div></div>
  <div class="card"><h2>By appointment type</h2>${rateTable(srt(R.byEvent), 'Event', lead)}</div>
  <div class="cols"><div class="card"><h2>By day of week</h2>${rateTable(dws, 'Day', lead)}</div><div class="card"><h2>By start time</h2>${rateTable(hrs, 'Hour', lead)}</div></div>
  <div class="small muted">Providers only (chairs and rooms are left out). Rates use past dates in the range. ${lead ? 'Lead time = days from booking to the visit.' : 'Add Created Date to the report to see lead time.'}</div>`;
}
function vImpact(){
  const R = S.R;
  const byP = new Map(); for (const e of R.exceptions){ const k = e.prov; const x = byP.get(k) || { blocked: 0, lunch: 0, off: 0, noday: 0, dbl: 0, min: 0, days: new Set() }; if (/Lunch/.test(e.why)) x.lunch++; if (/^In |Runs into/.test(e.why)) x.blocked++; if (/Outside template hours/.test(e.why)) x.off++; if (/No template that day/.test(e.why)) x.noday++; if (/Double-booked/.test(e.why)) x.dbl++; x.min += e.min || 0; x.days.add(e.date); byP.set(k, x); }
  const provs = [...R.tByProv.keys()].filter(p => !S.f.prov || p === S.f.prov);
  const kept = p => (R.byProv.find(x => x.key === p) || {}).aKept || 0;
  const whys = ['', 'In Lunch', 'In Admin', 'In Meeting', 'Runs into', 'Outside template hours', 'No template that day', 'Double-booked'];
  const ex = R.exceptions.filter(e => !S.impactWhy || e.why.includes(S.impactWhy));
  return `<div class="card"><h2>How much the schedule leans on the provider</h2><div class="tw"><table><thead><tr><th>Provider</th><th class="n">Kept visits</th><th class="n">In admin / lunch / meeting</th><th class="n">…of those in lunch</th><th class="n">Outside template hours</th><th class="n">On a non-template day</th><th class="n">Double-booked</th><th class="n">Extra hours</th><th class="n">Days affected</th></tr></thead><tbody>
    ${provs.map(p => { const x = byP.get(p) || { blocked: 0, lunch: 0, off: 0, noday: 0, dbl: 0, min: 0, days: new Set() }; return `<tr><td>${esc(p)}</td><td class="n">${num(kept(p))}</td><td class="n">${num(x.blocked)}</td><td class="n">${num(x.lunch)}</td><td class="n">${num(x.off)}</td><td class="n">${num(x.noday)}</td><td class="n">${num(x.dbl)}</td><td class="n">${(x.min / 60).toFixed(1)}</td><td class="n">${num(x.days.size)}</td></tr>`; }).join('')}
    </tbody></table></div><div class="small muted" style="margin-top:6px">Kept, expected and no-show visits on past dates. Extra hours = visit minutes that fall in blocked time or outside the template.</div></div>
  <div class="card"><h2>Visits outside the template <span class="sp"></span><select data-act="impactWhy">${whys.map(w => `<option value="${esc(w)}"${w === S.impactWhy ? ' selected' : ''}>${w || 'All reasons'}</option>`).join('')}</select></h2>
    <div class="tw" style="max-height:480px"><table><thead><tr><th>Date</th><th>Time</th><th>Provider</th><th>Event</th><th>Status</th><th>Why</th></tr></thead><tbody>${ex.slice(0, 400).map(e => `<tr><td>${TE.mdy(e.date)}</td><td>${TE.hhmm(e.time)}</td><td>${esc(e.prov)}</td><td>${esc(e.event)}</td><td>${esc(e.status)}</td><td>${esc(e.why)}</td></tr>`).join('') || '<tr><td colspan="6" class="muted">None in this range.</td></tr>'}</tbody></table></div>
    ${ex.length > 400 ? `<div class="small muted">Showing 400 of ${num(ex.length)}. Write report tabs for the full list.</div>` : ''}</div>`;
}
function vTrends(){
  const R = S.R; const wk = R.byWeek.filter(w => w.key >= R.from && w.key + 6 <= R.to && (w.supply || w.aAll));   // whole weeks only; a partial first or last week skews every rate
  const P = f => wk.map(w => ({ x: TE.mdy(w.key), y: f(w), sub: `${num(w.supply)} slots · ${num(w.aAll)} visits` }));
  return `<div class="card"><h2>Weekly trends</h2><div class="charts">
    ${lineChart('c1', 'Filled %', P(w => w.fill), v => pct(v), 1)}
    ${lineChart('c2', 'Template fit %', P(w => w.fit), v => pct(v), 1)}
    ${lineChart('c3', 'No-show rate', P(w => w.nsRate), v => pct(v, 1))}
    ${lineChart('c4', 'Cancel rate', P(w => w.cxRate), v => pct(v, 1))}
    ${R.hasLate ? lineChart('c5', 'Late cancel rate', P(w => w.lcRate), v => pct(v, 1)) : ''}
    ${lineChart('c6', 'Kept visits', P(w => w.aKept), v => num(v))}
    ${lineChart('c7', 'Visits outside template', P(w => w.inBlocked + w.offTemplate), v => num(v))}
  </div><div class="small muted" style="margin-top:6px">Whole weeks only (Sunday to Saturday) inside the From–To range. Hover a chart for the week’s value. Filters above apply.</div></div>`;
}
// ---------- Clinic Eligibility (central Sohar's file for one clinic; patients as LASFIR + MRN only) ----------
const PILL_CHIP = { red: 'bad', amber: 'warn', blue: 'info', grey: 'mute' };
const PILL_BY_LABEL = new Map(SR.PILLS.map(p => [p.label, p]));
const pillChip = l => `<span class="chip ${PILL_CHIP[(PILL_BY_LABEL.get(l) || {}).color] || 'mute'}">${esc(l)}</span>`;
const splitList = (s, sep) => String(s ?? '').split(sep).map(x => x.trim()).filter(Boolean);
const money = v => { const n = parseFloat(String(v ?? '').replace(/[$,]/g, '')); return isNaN(n) ? null : n; };
const usd = v => { const n = money(v); return n == null ? '' : '$' + n.toLocaleString('en-US', { maximumFractionDigits: 0 }); };
const seTag = r => r.LASFIR ? esc(r.LASFIR) : '<span class="muted">(no name)</span>';
const seId = r => esc(SC.idLabel(r) || 'No MRN');
const seWhen = s => { const d = new Date(s); return isNaN(d) ? '' : d.toLocaleString('en-US', { month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit' }); };
const SEV_WORD = { high: 'Action', medium: 'Check' };
// Sohar status codes and the Payer Map's "Choose in NextGen" in words front-desk staff can act on
function plainStatus(s){
  const t = String(s || '').trim();
  if (!t || t === 'Not in Sohar file') return 'Not checked yet';
  if (/^complete\.eligible/i.test(t)) return 'Eligible';
  if (/^complete\.ineligible/i.test(t)) return 'Not eligible';
  if (/^pending/i.test(t)) return 'Pending';
  const e = t.match(/^error\.?([^(\s]*)/i); if (e) return `Error${e[1] ? `: ${e[1].replace(/[._]+/g, ' ')}` : ''}`;
  if (/^coverage found/i.test(t)) return 'Insurance found';
  if (/^no coverage/i.test(t)) return 'No insurance found';
  return t;
}
function plainChoose(c){
  const s = String(c || ''); const m = s.match(/^Not in Payer Map: (.*?)(?: \([^)]*\))?$/);
  if (m) return `Insurance found: ${esc(m[1])}. Ask billing which NextGen plan to pick.`;
  return /^One of: /.test(s) ? `Choose one of: ${esc(s.slice(8))}` : `Choose “${esc(s)}”`;
}
// deductible / OOP meter: share met, then what is left
function meterBar(total, remaining, status){
  if (status === 'None') return '<span class="muted">None</span>';
  const t = money(total), r = money(remaining);
  if (t == null || r == null || t <= 0) return status ? esc(status) : '';
  const f = Math.max(0, Math.min(1, (t - r) / t));
  return `<span class="meter" role="img" aria-label="${pct(f)} met"><i style="width:${Math.round(f * 100)}%"></i></span> ${status === 'Met' ? '<b style="color:var(--ok)">Met</b>' : `${esc(usd(r))} left`} <span class="muted">of ${esc(usd(t))}</span>`;
}
const mcls = s => s === 'Met' ? 'met' : s === 'Approaching' ? 'appr' : '';
function seTaskMap(){ const T = new Map(); for (const t of S.E.tasks){ const k = String(t['Task Key'] || '').split('|')[0]; if (!T.has(k)) T.set(k, []); T.get(k).push(t); } return T; }
const seOpen = (r, T) => (T.get(r.key) || []).filter(t => t.clinic && !t.Done);
// Action / Check lines for a card: flag items (per-visit repeats collapsed, as in Sohar), then the admin's build review
function seItems(r){
  const rv = S.E.review.filter(x => x.MRN === r.key && x.Severity !== 'Low');
  const flags = SR.collapseFlags(rv.filter(x => x.Source === 'Flag').map(x => ({ code: x.Issue, label: x.Pill, severity: String(x.Severity).toLowerCase(), detail: x.Detail, appt: { date: x['Appt Date'], event: x.Event, rendering: x.Rendering } })));
  return flags.concat(rv.filter(x => x.Source === 'Build').map(x => ({ severity: String(x.Severity).toLowerCase(), detail: x.Detail ? `${x.Issue}: ${x.Detail}` : x.Issue })));
}
function seHeadHTML(){
  const m = S.SE.meta, run = SC.parseDate(m['Run Date']); const age = run ? Math.round((SC.parseDate(new Date()) - run) / 864e5) : 0;
  return `<div class="filters"><span class="small"><b>Clinic Eligibility</b> · ${esc(m.Location)} · data from ${esc(m['Run Date'])}</span>${age > 7 ? `<span class="chip warn">${num(age)} days old: ask the admin for a new file</span>` : ''}</div>`;
}
const seErrCard = () => `<div class="card"><h2>Clinic Eligibility</h2><div class="msg bad">Couldn’t work out the list: ${esc(S.Eerr)}</div></div>`;
function vEligibility(){
  if (S.Eerr || !S.E) return seErrCard();
  const rows = S.E.rows, T = seTaskMap();
  const count = new Map(); for (const r of rows) for (const l of splitList(r.Pills, ';')) count.set(l, (count.get(l) || 0) + 1);
  const withTasks = rows.filter(r => seOpen(r, T).length).length;
  if (S.sePill && !(S.sePill === 'tasks' ? withTasks : count.has(S.sePill))) S.sePill = '';
  const chip = (v, label, n, cls) => `<button class="chip ${cls}" data-act="sePill" data-v="${esc(v)}" aria-pressed="${S.sePill === v}">${esc(label)} ${num(n)}</button>`;
  return `<div class="card"><h2>Eligibility <span class="chip mute">${num(rows.length)} patients</span></h2>
    <div class="small muted">Everyone whose next visit is here, soonest first. Do the tasks before the visit. Your ticks, re-verify requests and notes are saved on the Clinic Work tab, and the admin picks them up from your clinic’s folder.</div>
    <div class="chips">${chip('', 'All', rows.length, 'mute')}${withTasks ? chip('tasks', 'Has tasks', withTasks, 'warn') : ''}${SR.PILLS.filter(p => count.has(p.label)).map(p => chip(p.label, p.label, count.get(p.label), PILL_CHIP[p.color] || 'mute')).join('')}</div>
    <input type="search" class="sesearch" data-act="seQ" value="${esc(S.seQ)}" placeholder="Search LASFIR or MRN" aria-label="Search LASFIR or MRN"></div>
  <div id="seList">${seListHTML(T)}</div>`;
}
function seListHTML(T = seTaskMap()){
  const q = S.seQ.trim().toLowerCase();
  const rows = S.E.rows.filter(r => (!S.sePill || (S.sePill === 'tasks' ? seOpen(r, T).length : splitList(r.Pills, ';').includes(S.sePill))) && (!q || [r.LASFIR, r.MRN].some(v => String(v || '').toLowerCase().includes(q))));
  if (!rows.length) return `<div class="card small muted">${S.E.rows.length ? 'No one matches this filter.' : 'No upcoming patients in this file. Ask the admin for a new one.'}</div>`;
  const days = new Map(); for (const r of rows){ const d = r['Next Appt'] || 'No date'; if (!days.has(d)) days.set(d, []); days.get(d).push(r); }
  const wd = s => { const d = SC.parseDate(s); return d ? d.toLocaleDateString('en-US', { weekday: 'short', timeZone: 'UTC' }) + ' ' : ''; };
  return `<div class="ptgrid">${[...days].map(([d, rs]) => `<h3 class="day">${esc(wd(d) + d)} <span class="chip mute">${rs.length}</span></h3>${rs.map(r => ptCard(r, T)).join('')}`).join('')}</div>`;
}
function ptCard(r, T){
  const pills = splitList(r.Pills, ';'); const shown = pills.includes('Kaiser/VA MM no auth') ? pills.filter(l => l !== 'Kaiser/VA MM') : pills;   // the filter category only repeats it
  const high = pills.some(l => (PILL_BY_LABEL.get(l) || {}).severity === 'high');
  const items = seItems(r), notes = splitList(r.Notes, '|');
  const tasks = T.get(r.key) || [], mine = tasks.filter(t => t.clinic), adminOpen = tasks.filter(t => !t.clinic && !t.Done).length;
  const note = S.E.notes.get(r.key);
  const ver = r['Verified Date'] ? `Verified ${esc(r['Verified Date'])}${r['Days Since Verified'] !== '' ? ` (${esc(days(+r['Days Since Verified']))} ago)` : ''}` : 'Not verified in Sohar yet';
  const ben = [['Deductible', r.Deductible, r['Deductible Remaining'], r['Deductible Status']], ['OOP max', r['OOP Max'], r['OOP Remaining'], r['OOP Status']]]
    .map(([l, t, rem, s]) => { const b = meterBar(t, rem, s); return b ? `<div class="mrow ${mcls(s)}"><span class="ml">${l}</span>${b}</div>` : ''; }).join('');
  const k = esc(r.key);
  return `<div class="card pt${high ? ' hi' : ''}">
    <div class="pt-h"><span class="muted">${esc(r['Next Time'])}</span><b>${seTag(r)}</b><span class="small muted">${seId(r)}</span></div>
    <div class="small">${esc([r['Next Event'], r['Next Provider']].filter(Boolean).join(' · '))}</div>
    ${shown.length ? `<div class="chips">${shown.map(pillChip).join('')}</div>` : ''}
    <dl><dt>Payer</dt><dd>${esc(r['Primary Payer'] || 'None on file')}${r['Secondary Payer'] ? ` <span class="muted">+ ${esc(r['Secondary Payer'])}</span>` : ''}</dd>
      ${r['Choose in NextGen'] ? `<dt>NextGen</dt><dd>${plainChoose(r['Choose in NextGen'])}</dd>` : ''}
      <dt>Sohar</dt><dd>${esc(plainStatus(r['Sohar Status']))}${r.Plan ? ` · ${esc(r.Plan)}` : ''}<br><span class="small muted">${ver}</span></dd>
      ${r.Copay !== '' || r.Coinsurance !== '' ? `<dt>Copay</dt><dd>${r.Copay !== '' ? esc(usd(r.Copay) || r.Copay) : '—'}${r.Coinsurance !== '' ? ` · coinsurance ${esc(r.Coinsurance)}` : ''}</dd>` : ''}</dl>
    ${ben}
    ${items.length ? `<ul class="items">${items.map(t => `<li class="sev-${esc(t.severity || 'high')}">${SEV_WORD[t.severity] ? `<b>${SEV_WORD[t.severity]}:</b> ` : ''}${esc(t.detail)}</li>`).join('')}</ul>` : ''}
    ${mine.map(seTaskHTML).join('')}
    ${adminOpen ? `<div class="small muted">${num(adminOpen)} more task${adminOpen === 1 ? '' : 's'} for the admin.</div>` : ''}
    ${r['Verification Skipped'] || r['Re-verify Anyway'] ? `<div class="msg info small">${r['Verification Skipped'] ? `${esc(r['Verification Skipped'])}, so this patient isn’t on this week’s verification file (next check ${esc(r['Next Check'] || '?')}).` : 'Verified recently.'}
      <label class="ck"><input type="checkbox" data-act="seRv" data-m="${k}"${r['Re-verify Anyway'] ? ' checked' : ''}> Re-verify anyway (new card, or another reason to check)</label></div>` : ''}
    ${note ? `<div class="msg info small"><b>Your note:</b> ${esc(note.Note)} <span class="muted">· ${esc(note.By)}${seWhen(note.When) ? ', ' + esc(seWhen(note.When)) : ''}</span></div>` : ''}
    ${notes.length ? `<details class="more"><summary>Sohar notes (${notes.length})</summary><ul class="items">${notes.map(t => `<li>${esc(t)}</li>`).join('')}</ul></details>` : ''}
    <div class="pt-act"><button class="btn sm" data-act="seNote" data-m="${k}">${note ? 'Edit note' : 'Add note'}</button></div>
  </div>`;
}
// a clinic task: the Done tick, and next to it the done note (the Note on the task's Clinic Work "Task done" row). Under it, the
// task's Clinic Note on the admin's NextGen Tasks when the file was made (sentNote, from SE Done): what the admin already has.
const seSentHTML = t => (t.sentNote ? `<div class="small tnote"><span class="muted">The admin has this note:</span> ${esc(t.sentNote)}</div>` : '');
function seTaskHTML(t){
  const tk = esc(t['Task Key']), m = esc(String(t['Task Key'] || '').split('|')[0]);
  return `<div class="task"><label class="ck"><input type="checkbox" data-act="seDone" data-k="${tk}" data-m="${m}"${t.Done ? ' checked' : ''}${t.doneBy === 'admin' ? ' disabled' : ''}><span><b>${esc(t.Task)}</b>${t.Details ? ` <span class="small muted">${esc(t.Details)}</span>` : ''}${t.doneBy === 'admin' ? ' <span class="chip ok">Done by the admin</span>' : ''}</span></label>
    <button class="btn sm" data-act="seTaskNote" data-k="${tk}" data-m="${m}" aria-label="${esc((t.clinicNote ? 'Edit the done note for ' : 'Add a done note for ') + t.Task)}">${t.clinicNote ? 'Edit note' : 'Note'}</button></div>
    ${seSentHTML(t)}${t.clinicNote ? `<div class="small tnote"><span class="muted">Your note:</span> ${esc(t.clinicNote)}</div>` : ''}`;
}
function seTaskNoteDialog(key){
  const t = S.E.tasks.find(x => x['Task Key'] === key), r = t && seRow(String(key).split('|')[0]); if (!t || !r.key) return;
  openModal(`<h2>Done note · ${seTag(r)} · ${seId(r)}</h2>
    <div class="small"><b>${esc(t.Task)}</b>${t.Details ? ` <span class="muted">${esc(t.Details)}</span>` : ''}</div>
    ${seSentHTML(t)}
    <div class="fld">What you did<textarea id="seNoteText" maxlength="500" placeholder="${esc('For example: called the patient, new card scanned')}">${esc(t.clinicNote || '')}</textarea></div>
    <div class="small muted">Saved with this task on the Clinic Work tab, and the admin sees it next to the task. You can save a note before you tick the box. LASFIR and MRN are enough: don’t type the patient’s name or date of birth. ${t.sentNote ? 'A new note takes the place of the admin’s note above once it reaches them. ' : ''}Clearing the box takes the note off this file; a note the admin already has stays.</div>
    <div class="row" style="margin-top:12px"><button class="btn primary" data-act="seTaskNoteSave" data-k="${esc(key)}">Save note</button><button class="btn" data-act="close">Cancel</button></div>`);
}
function vBenefits(){
  if (S.Eerr || !S.E) return seErrCard();
  const B = S.E.benefits, sum = SC.benefitsSummary(B), ap = +((S.SE.settings || {}).approachingPct) || 0.2;
  if (S.seBen && !B.rows.some(b => b.Flag === S.seBen)) S.seBen = '';
  const rows = B.rows.filter(b => !S.seBen || b.Flag === S.seBen);
  const k = (l, v, s) => `<div class="kpi"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`;
  const chip = (v, label, n, cls) => `<button class="chip ${cls}" data-act="seBen" data-v="${esc(v)}" aria-pressed="${S.seBen === v}">${esc(label)} ${num(n)}</button>`;
  const flagChip = f => f ? `<span class="chip ${/met$/.test(f) ? 'ok' : 'warn'}">${esc(f)}</span>` : '';
  return `<div class="kpis">${k('OOP max met', num(sum.oopMet), 'the plan pays in full')}${k('OOP max approaching', num(sum.oopApproaching), `within ${pct(ap)} of it`)}${k('Deductible met', num(sum.dedMet), 'copay or coinsurance only')}${k('Deductible approaching', num(sum.dedApproaching), `within ${pct(ap)} of it`)}</div>
  <div class="card"><h2>Benefits <span class="chip mute">${num(sum.verified)} verified</span></h2>
    <div class="small muted">Latest Sohar verification for each patient, closest to the out-of-pocket max first. Met means the patient has paid the whole amount for the plan year, so the plan pays more of each visit. Use these to set expectations about what the patient owes.</div>
    <div class="chips">${chip('', 'All', B.rows.length, 'mute')}${[['OOP max met', sum.oopMet, 'ok'], ['OOP max approaching', sum.oopApproaching, 'warn'], ['Deductible met', sum.dedMet, 'ok'], ['Deductible approaching', sum.dedApproaching, 'warn']].filter(x => x[1]).map(x => chip(x[0], x[0], x[1], x[2])).join('')}</div>
    <div class="tw"><table class="ben"><thead><tr><th>Patient</th><th>Payer</th><th>Benefits</th></tr></thead><tbody>
    ${rows.map(b => { const m = (l, t, rem, s) => { const x = meterBar(t, rem, s); return x ? `<div class="mrow ${mcls(s)}"><span class="ml">${l}</span>${x}</div>` : ''; };
      return `<tr><td><b>${seTag(b)}</b> <span class="small muted">${seId(b)}</span><br><span class="small muted">${esc([b['Next Appt'], b['Next Time']].filter(Boolean).join(' '))}</span>${b.Flag ? '<br>' + flagChip(b.Flag) : ''}</td><td>${esc(b['Sohar Payer'])}${b.Plan ? `<br><span class="small muted">${esc(b.Plan)}</span>` : ''}</td>
        <td><div class="small">Copay ${b.Copay !== '' ? esc(usd(b.Copay) || b.Copay) : '—'}${b.Coinsurance !== '' ? ` · coinsurance ${esc(b.Coinsurance)}` : ''}</div>${m('OOP max', b['OOP Max'], b['OOP Remaining'], b['OOP Status'])}${m('Deductible', b.Deductible, b['Deductible Remaining'], b['Deductible Status'])}</td></tr>`; }).join('') || '<tr><td colspan="3" class="muted">No verified benefits in this file.</td></tr>'}
    </tbody></table></div>
    ${sum.pending ? `<div class="small muted" style="margin-top:6px">${num(sum.pending)} more wait on a verification of insurance Sohar found; their benefits come with a later file.</div>` : ''}</div>`;
}
// how far the Payer Map's match can be trusted, in words: [chip class, short label, sentence]
function payerTrust(l){
  if (!l) return ['mute', 'No history yet', 'Sohar hasn’t seen enough of this payer to say. Check the card at check-in.'];
  const c = String(l.Confidence || '').toLowerCase(), to = esc(l['Sohar payerName'] || l['Sohar payerId'] || 'a Sohar payer');
  if (c === 'not in payer map') return ['bad', 'Not mapped', 'Sohar can’t check this payer yet, so these patients get a search for any insurance instead. Check the card and ask billing which NextGen plan to use.'];
  if (c === 'not a payer') return ['mute', 'Not insurance', 'This NextGen payer isn’t an insurance plan (for example self-pay).'];
  if (['exact', 'confirmed', 'probable'].includes(c)) return ['ok', 'Reliable', `A person matched it to ${to}.`];
  if (c === 'learned') return ['ok', 'Reliable', `Sohar agreed with the match to ${to} on ${esc(l['Patients agreed'])} patients.`];
  if (c === 'auto-high') return ['ok', 'Likely right', `Matched by name to ${to}.`];
  if (c === 'auto-medium') return ['warn', 'A guess', `Matched by name to ${to}, but nobody has confirmed it.`];
  return ['warn', 'Not sure', 'The match needs a person to check it.'];
}
function vPayers(){
  if (S.Eerr || !S.E) return seErrCard();
  const P = S.E.payers;
  const found = p => { const l = p.learning || {}, m = p.metrics || {}, out = [];
    if (+l['Patients agreed']) out.push(`Sohar agreed on ${num(+l['Patients agreed'])} patient${+l['Patients agreed'] === 1 ? '' : 's'}`);
    if (+l['Payer errors']) out.push(`${num(+l['Payer errors'])} came back with a payer error`);
    if (+l['Discovery other payer']) out.push(`a search found a different payer for ${num(+l['Discovery other payer'])} of ${num(+l['Discovery checked'])}${l['Most common other payer'] ? ` (most often ${esc(l['Most common other payer'])})` : ''}`);
    if (+m['Failed verification patients']) out.push(`${num(+m['Failed verification patients'])} failed verification${m['Most common failure'] ? ` (most often ${esc(m['Most common failure'])})` : ''}`);
    return out; };
  return `<div class="card"><h2>Payers <span class="chip mute">${num(P.length)}</span></h2><div class="small muted">Each insurance your patients carry in NextGen, and how far the match to Sohar can be trusted. When it isn’t reliable, check the card and ask billing. A payer note goes to the admin; it never changes the mapping by itself.</div></div>
  <div class="ptgrid">${P.map(p => { const [cls, label, why] = payerTrust(p.learning); const f = found(p), m = p.metrics || {}, n = p.note && (p.note.cardSays || p.note.note) ? p.note : null;
    return `<div class="card pt"><div class="pt-h"><b>${esc(p.payer)}</b><span class="small muted">${num(p.patients)} patient${p.patients === 1 ? '' : 's'}</span></div>
      <div class="chips"><span class="chip ${cls}">${esc(label)}</span>${m.Recurring === 'Yes' ? `<span class="chip warn">Keeps coming up (${esc(m['Days with issues'])} days)</span>` : ''}</div>
      <div class="small" style="margin-top:4px">${why}</div>
      ${f.length ? `<div class="small" style="margin-top:4px"><b>What Sohar found:</b> ${f.join('; ')}.</div>` : ''}
      ${p.learning && p.learning.Suggestion ? `<div class="small" style="margin-top:4px"><b>Suggestion:</b> ${esc(p.learning.Suggestion)}</div>` : ''}
      ${cls !== 'ok' ? '<div class="msg warn small">Check the card at check-in, and ask billing if it doesn’t match.</div>' : ''}
      ${n ? `<div class="msg info small"><b>Your note:</b> ${n.cardSays ? `card says “${esc(n.cardSays)}”` : ''}${n.cardSays && n.note ? '. ' : ''}${esc(n.note)} <span class="muted">· ${esc(n.by)}${seWhen(n.when) ? ', ' + esc(seWhen(n.when)) : ''}</span></div>` : ''}
      <div class="pt-act"><button class="btn sm" data-act="sePayerNote" data-p="${esc(p.payer)}">${n ? 'Edit payer note' : 'Add payer note'}</button></div></div>`; }).join('') || '<div class="card small muted">No payers in this file.</div>'}</div>`;
}
function seNoteDialog(key){
  const r = S.E.rows.find(x => x.key === key); if (!r) return; const n = S.E.notes.get(key);
  openModal(`<h2>Note · ${seTag(r)} · ${seId(r)}</h2>
    <div class="fld">Note for the admin<textarea id="seNoteText" maxlength="1000">${esc(n ? n.Note : '')}</textarea></div>
    <div class="small muted">Saved on the Clinic Work tab; the admin picks it up from your clinic’s folder. LASFIR and MRN are enough: don’t type the patient’s name or date of birth. Clear the box to remove the note.</div>
    <div class="row" style="margin-top:12px"><button class="btn primary" data-act="seNoteSave" data-m="${esc(key)}">Save note</button><button class="btn" data-act="close">Cancel</button></div>`);
}
function sePayerNoteDialog(name){
  const p = S.E.payers.find(x => x.payer === name); const n = p && p.note;
  openModal(`<h2>Payer note · ${esc(name)}</h2>
    <div class="fld">What the card says (plan name, group, payer phone)<input id="seCard" maxlength="200" value="${esc(n ? n.cardSays : '')}"></div>
    <div class="fld">Note for the admin<textarea id="seNoteText" maxlength="1000">${esc(n ? n.note : '')}</textarea></div>
    <div class="small muted">The admin reads this and decides what to change. No member IDs or patient names here. Clear both boxes to remove the note.</div>
    <div class="row" style="margin-top:12px"><button class="btn primary" data-act="sePayerNoteSave" data-p="${esc(name)}">Save note</button><button class="btn" data-act="close">Cancel</button></div>`);
}
// a tick or chip re-draws all of #app: put keyboard focus back on the same control (matched by its data-act/k/m/v/p) so a
// keyboard or screen-reader user keeps their place in the list
function renderKeepFocus(){
  const sel = actSel(document.activeElement);
  render();
  const el = sel && $('#app').querySelector(sel); if (el) el.focus();
}
// Clinic Work: the whole small sheet is rewritten on each change (last row per Kind + Key wins); writes run one at a time.
// The list updates at once from this pane's copy; writeSheet applies the same change to the rows on the sheet at write time
// (someone else may have saved since) and hands them back.
let workQ = Promise.resolve();
function seBy(){ let n = ''; try { n = localStorage.getItem('ctWho') || ''; } catch {} return n.trim() || `Clinic ${S.SE.meta.Location}`; }
function seWrite(change){
  // When stamped once here, so this copy and the rows written at the sheet match (writeSheet restamps it only when the sheet
  // already has a later When for the same Kind + Key; the rows then differ and come back below)
  const ch = Object.assign({ By: seBy(), When: new Date().toISOString() }, change);
  const rows = SR.workRows(S.data.work, ch);
  S.data.work = rows; runSE(); renderKeepFocus();
  const p = workQ.then(() => S.store.writeSheet(SR.SE_SHEETS.work, null, { keepFormat: true, inPlace: true, text: true, table: SE_WORK_TABLE, change: ch })); workQ = p.catch(() => {});
  return p.then(r => {
    // the sheet had rows this pane hadn't seen: show them. Skipped while a later change is on its way (its write hands back newer rows).
    if (r && Array.isArray(r.rows) && S.data.work === rows && JSON.stringify(r.rows) !== JSON.stringify(rows)){ S.data.work = r.rows; runSE(); renderKeepFocus(); }
    // saved, but outside the table the admin's flow reads: say so once a session
    if (r && r.table && r.table !== 'ok' && !S.seTableWarned){ S.seTableWarned = true; toast('warn', `Saved on the Clinic Work tab. The admin’s inbox reads only the rows inside the Clinic Work table, and ${r.table === 'missing' ? 'this file doesn’t have one' : 'it couldn’t be made bigger'}, so send the saved file to the admin in OneDrive or Teams.`, 12000); }
  }, e => { console.error(e); toast('bad', `Couldn’t save that on the Clinic Work tab: ${esc(e.message || String(e))}. Try again after the refresh.`, 9000); return reload(); });
}
const seRow = key => S.E.rows.find(x => x.key === key) || {};
function seSetupCard(){
  const d = S.data, se = S.SE;
  if (!se) return d.seErr ? `<div class="card"><h2>Clinic Eligibility file</h2><div class="msg bad">${esc(d.seErr)}</div></div>` : '';
  let who = ''; try { who = localStorage.getItem('ctWho') || ''; } catch {}
  return `<div class="card"><h2>Clinic Eligibility file</h2>
    <div class="small">${esc(se.meta.Location)} · data from ${esc(se.meta['Run Date'])}${S.E ? ` · ${num(S.E.counts.patients)} patients · ${num(S.E.counts.clinicTasks)} open tasks for you` : ''} · ${num(d.work.length)} change${d.work.length === 1 ? '' : 's'} on Clinic Work</div>
    <div class="fld">Your name or initials (goes in By on Clinic Work)<input data-act="seWho" maxlength="60" value="${esc(who)}" placeholder="${esc('Clinic ' + se.meta.Location)}"></div>
    <div class="small muted">Central Sohar made this file for your clinic. It holds PHI (MRNs, insurance, visits): keep it in your clinic’s folder in the Clinic Eligibility library (MHS OneDrive or Teams only). Excel on the web saves your changes on its own and the admin picks them up from there, so you don’t need to send the file back. If the file isn’t in the library, or the admin asks for it, save it and send it to the admin through OneDrive or Teams. Don’t edit the hidden tabs.</div>
    <label class="ck" style="margin-top:8px"><input type="checkbox" data-act="seTemplates"${S.opts.templates === 'On' ? ' checked' : ''}> Show template tools <span class="small muted">(templates, cadence, recall and access; they need a NextGen appointment list)</span></label>
    <div class="small muted" style="margin-top:6px">Clinic Template Tool ${esc(APP_VERSION)} · Sohar file ${esc(se.meta['Sohar Version'])}</div></div>`;
}
function vSetup(){
  if (S.SE && S.opts.templates !== 'On') return seSetupCard();
  const d = S.data, ap = d.ap, R = S.R;
  const has = ap && !ap.error ? ap.has : {};
  const flag = (ok, t) => `<span class="chip ${ok ? 'ok' : 'warn'}">${ok ? '✓' : '—'} ${t}</span>`;
  const linkRows = d.templates.map(t => { const tt = R && R.templates.find(x => x.sheet === t.sheet && x.anchor[0] === t.anchor[0] && x.anchor[1] === t.anchor[1]); return { t, prov: tt && tt.provider }; });
  const hidden = S.SE ? d.hidden.filter(n => !/^se /i.test(n) && n !== SHEETS.pcr) : d.hidden;   // a clinic file’s own data tabs are read
  return seSetupCard() + `<div class="card"><h2>Appointment list</h2>
    ${!ap ? `<div class="msg warn">No Appointment List tab yet. Import the NextGen appointment report below.</div>` : ap.error ? `<div class="msg bad">${esc(ap.error)} (tab “${esc(ap.sheet)}”)</div>` :
    `<div class="small">Tab <b>${esc(ap.sheet)}</b>: ${num(ap.rows.length)} visits, ${TE.mdy(ap.min)} – ${TE.mdy(ap.max)}${ap.runAt != null ? `, report run ${TE.mdy(ap.runAt)} ${TE.hhmm(Math.round((ap.runAt % 1) * 1440))}` : ''}.</div>
     <div class="row" style="margin:6px 0">${flag(true, 'Date, time, provider, event, status')}${flag(has.dur, 'Duration / end time')}${flag(has.loc, 'Location')}${flag(has.cancel, 'Cancel date/time')}${flag(has.booked, 'Created date')}${flag(has.pt || has.mrn, 'Patient (Pat Name, Md Rc)')}${flag(has.payer, 'Payer Name')}</div>
     ${has.cancel && has.booked ? '' : `<div class="msg info">Add <b>${[!has.cancel && 'Cancel Date and Cancel Time', !has.booked && 'Created Date (when the visit was booked)'].filter(Boolean).join('</b> and <b>')}</b> to the NextGen report. ${!has.cancel ? 'Cancel date/time is what separates late cancels from early ones. ' : ''}${!has.booked ? 'Created date gives lead time and lets the tool rebuild 3NA for past weeks.' : ''}</div>`}`}
    <div class="row"><label class="btn primary">Import NextGen report…<input type="file" accept=".xls,.xlsx,.csv" data-act="importAppts" hidden></label><span class="small muted">Replaces the Appointment List tab. Use the full date range you want to study plus the weeks ahead.</span></div></div>
  <div class="card"><h2>Provider templates<span class="sp"></span><button class="btn sm primary" data-act="newTemplate">New template tab</button></h2>
    <div class="small muted">Every visible tab is scanned for a weekly grid: a Monday–Friday header row, times down the left, slot types in the cells, and Provider:/Location:/Effective From: labels above. The New template tab button makes one in the standard layout; you can also paste a copy of an existing template.</div>
    <table style="margin-top:6px"><thead><tr><th>Tab</th><th>Label</th><th>Linked to</th><th>Slots</th><th>Hours</th></tr></thead><tbody>${linkRows.map(({ t, prov }) => `<tr><td>${esc(t.sheet)}</td><td>${esc(t.label || '(no label)')}</td><td>${prov ? esc(prov) : '<span class="chip bad">Not linked</span> <span class="small muted">set it on the Providers tab</span>'}</td><td>${t.slots.length} × ${t.step} min</td><td>${TE.hhmm(t.start)}–${TE.hhmm(t.end)}</td></tr>`).join('') || '<tr><td colspan="5" class="muted">None found.</td></tr>'}</tbody></table>
    ${hidden.length ? `<div class="small muted" style="margin-top:6px">Hidden tabs aren’t read: ${hidden.map(esc).join(', ')}.</div>` : ''}</div>
  <div class="card"><h2>Lookup tabs</h2>
    <table><tbody>
      <tr><td><b>${SHEETS.st}</b></td><td class="small">Slot type → Clinical or Blocked, and which visits it accepts (event class, service line or event; comma = any of, + = all of, * = anything).</td><td>${d.stExists ? flag(true, `${d.st.length} rows`) : flag(false, 'Using built-in defaults')}</td><td><button class="btn sm" data-act="cfgSt">${d.stExists ? 'Add missing types' : 'Create'}</button></td></tr>
      <tr><td><b>${SHEETS.ref}</b></td><td class="small">NextGen event → event class, service line, whether it counts (holds and reference rows don’t), and Role (Consult, Start, Treatment, Follow-up) for cadence and the funnel.</td><td>${d.refExists ? flag(!(R && R.unmapped.length), R && R.unmapped.length ? `${R.unmapped.length} events not listed` : `${d.ref.length} rows`) : flag(false, 'Missing')}</td><td><button class="btn sm" data-act="cfgRef">${d.refExists ? 'Add missing events' : 'Create'}</button></td></tr>
      <tr><td><b>${SHEETS.pcr}</b></td><td class="small">Which payers let an MD or an NP bill each esketamine visit, and Besse approval. Import the shared Payer Credential Rules file from OneDrive (Add-In Manifests); edit the tab when a payer changes.</td><td>${d.payerRules ? (d.payerRules.error ? flag(false, 'Can’t read it') : flag(true, `${d.payerRules.rules.length} rows`)) : flag(false, 'Not imported')}</td><td><label class="btn sm">Import…<input type="file" accept=".csv,.xlsx,.xls" data-act="importPayerRules" hidden></label></td></tr>
      <tr><td><b>${SHEETS.prov}</b></td><td class="small">NextGen provider → template label, credential, include. Leave Include blank to let the tool decide (people yes, chairs and rooms no).</td><td>${d.provExists ? flag(true, `${d.prov.length} rows`) : flag(false, 'Missing')}</td><td><button class="btn sm" data-act="cfgProv">${d.provExists ? 'Add missing providers' : 'Create'}</button></td></tr>
    </tbody></table>
    ${R && R.unknownTypes.length ? `<div class="msg warn">Slot types in templates but not on ${SHEETS.st}: ${R.unknownTypes.map(x => esc(x.type)).join(', ')}.</div>` : ''}</div>
  <div class="cols"><div class="card"><h2>Settings</h2>
    <div class="fld">Late cancel window (hours before the visit)<input type="number" min="1" max="168" data-o="lateHours" value="${esc(S.opts.lateHours)}"></div>
    <div class="fld">3NA target (days)<input type="number" min="1" max="120" data-o="target" value="${esc(S.opts.target)}"></div>
    <div class="fld">As-of date (blank = report run time)<input type="date" data-o="asOf" value="${esc(S.opts.asOf)}"></div>
    <h3>Cadence and recall</h3>
    <div class="row"><div class="fld" style="flex:1">Release window (business hours)<input type="number" min="1" max="120" data-o="releaseHours" value="${esc(S.opts.releaseHours)}"></div><div class="fld" style="flex:1">Business day starts<input type="time" data-o="dayStart" value="${esc(S.opts.dayStart)}"></div><div class="fld" style="flex:1">Business day ends<input type="time" data-o="dayEnd" value="${esc(S.opts.dayEnd)}"></div></div>
    <div class="row"><div class="fld" style="flex:1">Overdue after (× usual interval)<input type="number" min="1" max="5" step="0.1" data-o="overdueMult" value="${esc(S.opts.overdueMult)}"></div><div class="fld" style="flex:1">Lapsed after (days)<input type="number" min="7" max="730" data-o="lapsedDays" value="${esc(S.opts.lapsedDays)}"></div><div class="fld" style="flex:1">Match horizon (business days)<input type="number" min="1" max="60" data-o="matchDays" value="${esc(S.opts.matchDays)}"></div></div>
    <div class="row"><div class="fld" style="flex:1">Consult converts within (days)<input type="number" min="7" max="365" data-o="convDays" value="${esc(S.opts.convDays)}"></div><div class="fld" style="flex:1">TMS course (treatments)<input type="number" min="1" max="60" data-o="tmsCount" value="${esc(S.opts.tmsCount)}"></div><div class="fld" style="flex:1">Maintenance after (days)<input type="number" min="30" max="730" data-o="maintDays" value="${esc(S.opts.maintDays)}"></div></div></div>
  <div class="card"><h2>Share results</h2><div class="small muted">Writes summary tabs (prefixed “TA ”) with the current filters. Counts only, no patient fields. Safe to share with anyone who can open this workbook. Recall → Write patient tabs writes the “PT ” tabs, which list patients by LASFIR code.</div>
    <div class="row" style="margin-top:8px"><button class="btn primary" data-act="writeReports">Write report tabs</button><button class="btn" data-act="snapshot">Save 3NA snapshot</button></div>
    <h3>ESK access model</h3><div class="small muted">Import the export from the ESK Tracker’s Access tab (Export for template tool). Only the summary tabs are read; patient tabs in that file are ignored.</div>
    <div class="row" style="margin-top:6px"><label class="btn">Import access model…<input type="file" accept=".xlsx" data-act="importEsk" hidden></label>${S.esk ? `<span class="small muted">Loaded: ${esc(S.esk.modality)} · ${esc(S.esk.clinic)} · ${esc(String(S.esk.generated).slice(0, 10))}</span>` : ''}</div></div></div>`;
}

// ---------- navigation ----------
// A Clinic Eligibility file shows only the Sohar views unless Setup → Show template tools is on. Other workbooks: as before.
function shownViews(st = S){
  const tools = !st.SE || (st.opts && st.opts.templates) === 'On';
  return VIEWS.map(v => v[0]).filter(k => SE_VIEWS.includes(k) ? !!st.SE : k === 'setup' || tools);
}
const viewOff = (k, st = S) => !SE_VIEWS.includes(k) && k !== 'setup' && !st.R;   // template views need the appointment list
function navGroups(st = S){
  const vis = shownViews(st);
  return NAV.map(([k, l, vs]) => ({ k, l, views: vs.filter(v => vis.includes(v)) })).filter(g => g.views.length);
}
// the view to draw: the chosen one if it is shown and has its data, else Eligibility (clinic file) or Setup
function currentView(st = S){
  const ready = !!st.R || !!st.SE; const v = ready ? st.view : 'setup';
  return shownViews(st).includes(v) && !viewOff(v, st) ? v : st.SE ? 'eligibility' : 'setup';
}
function navHTML(view){
  const groups = navGroups(); const cur = groups.find(g => g.views.includes(view));
  const go = g => { const on = g.views.filter(v => !viewOff(v)); return g === cur ? view : on.includes(S.sub[g.k]) ? S.sub[g.k] : on[0] || g.views[0]; };
  const top = `<nav class="tabs">${groups.map(g => `<button data-view="${go(g)}"${g === cur ? ' aria-current="page"' : ''}${g.views.every(v => viewOff(v)) ? ' disabled' : ''}>${esc(g.views.length === 1 ? VIEW_LABEL[g.views[0]] : g.l)}</button>`).join('')}</nav>`;
  const sub = cur && cur.views.length > 1 ? `<nav class="tabs sub" aria-label="${esc(cur.l)}">${cur.views.map(k => `<button data-view="${k}"${k === view ? ' aria-current="page"' : ''}${viewOff(k) ? ' disabled' : ''}>${esc(VIEW_LABEL[k])}</button>`).join('')}</nav>` : '';
  if (cur) S.sub[cur.k] = view;
  return top + sub;
}

function render(){
  for (const k in CHARTS) delete CHARTS[k];
  if (S.err) { $('#app').innerHTML = `<main><div class="card"><h2>Can’t read the workbook</h2><div class="msg bad">${esc(S.err)}</div><button class="btn primary" data-act="refresh">Try again</button></div></main>`; return; }
  if (!S.data) return;
  const view = currentView(); const se = SE_VIEWS.includes(view);
  const body = { overview: vOverview, templates: vTemplates, access: vAccess, cadence: vCadence, funnel: vFunnel, recall: vRecall, visits: vVisits, impact: vImpact, trends: vTrends, eligibility: vEligibility, benefits: vBenefits, payers: vPayers, setup: vSetup }[view]();
  $('#app').innerHTML = `<header class="top"><div class="bar"><div class="brand"><span class="dot" title="Clinic Template Tool ${esc(APP_VERSION)}">CT</span><span>Clinic Template Tool</span><small>${S.R && !se ? `${S.R.tByProv.size} templates · ${num(S.data.ap.rows.length)} visits` : ''}</small></div>
    <button class="iconbtn" data-act="refresh" title="Re-read the workbook" aria-label="Refresh">⟳</button>
    ${S.mode === 'pane' && S.canFull ? `<button class="iconbtn" data-act="full" title="Full screen" aria-label="Full screen">⛶</button>` : ''}${S.mode === 'dialog' ? `<button class="iconbtn" data-act="closeFull" title="Close full screen" aria-label="Close full screen">✕</button>` : ''}</div>
    ${navHTML(view)}
    ${se ? seHeadHTML() : S.R && view !== 'setup' ? filtersHTML() : ''}</header><main>${body}</main>`;
}

// ---------- actions ----------
document.addEventListener('click', async e => {
  const v = e.target.closest('[data-view]'); if (v && !v.disabled){ S.view = v.dataset.view; render(); window.scrollTo(0, 0); return; }
  const a = e.target.closest('[data-act]'); if (!a || a.tagName === 'SELECT' || a.tagName === 'INPUT') return;
  const act = a.dataset.act;
  try {
    if (act === 'refresh') await reload();
    else if (act === 'full') openFull();
    else if (act === 'closeFull') S.store.close();
    else if (act === 'newTemplate') newTemplateDialog();
    else if (act === 'makeTemplate') await makeTemplate();
    else if (act === 'close' || act === 'askNo') closeModal();
    else if (act === 'askYes'){ const r = askRes; askRes = null; closeModal(); if (r) r(true); }
    else if (act === 'cfgSt') await writeSlotTypes();
    else if (act === 'cfgRef') await writeRef();
    else if (act === 'cfgProv') await writeProviders();
    else if (act === 'writeReports') await writeReports();
    else if (act === 'snapshot') await saveSnapshot();
    else if (act === 'dlManifest') downloadManifest();
    else if (act === 'copyOutreach') await copyOutreach();
    else if (act === 'writePatients') await writePatientTabs();
    else if (act === 'sePill'){ S.sePill = S.sePill === a.dataset.v ? '' : a.dataset.v; renderKeepFocus(); }
    else if (act === 'seBen'){ S.seBen = S.seBen === a.dataset.v ? '' : a.dataset.v; renderKeepFocus(); }
    else if (act === 'seNote') seNoteDialog(a.dataset.m);
    else if (act === 'sePayerNote') sePayerNoteDialog(a.dataset.p);
    else if (act === 'seNoteSave'){ const r = seRow(a.dataset.m), note = $('#seNoteText').value.trim(); closeModal(); await seWrite({ Kind: 'Note', Key: a.dataset.m, MRN: r.MRN, LASFIR: r.LASFIR, Note: note }); }
    else if (act === 'seTaskNote') seTaskNoteDialog(a.dataset.k);
    // no Value: workRows keeps the tick as it is (a note-only row leaves Done alone)
    else if (act === 'seTaskNoteSave'){ const k = a.dataset.k, r = seRow(k.split('|')[0]), note = $('#seNoteText').value.trim(); closeModal(); await seWrite({ Kind: 'Task done', Key: k, MRN: r.MRN, LASFIR: r.LASFIR, Note: note }); }
    else if (act === 'sePayerNoteSave'){ const card = $('#seCard').value.trim(), note = $('#seNoteText').value.trim(); closeModal(); await seWrite({ Kind: 'Payer note', Key: a.dataset.p, Value: card, Note: note }); }
  } catch (err){ console.error(err); busy(false); toast('bad', esc(err.message || String(err)), 9000); }
});
// search re-draws the list only, so the box keeps focus while typing
document.addEventListener('input', e => { const t = e.target; if (t.dataset && t.dataset.act === 'seQ' && S.E){ S.seQ = t.value; const box = $('#seList'); if (box) box.innerHTML = seListHTML(); } });
document.addEventListener('change', async e => {
  const t = e.target;
  if (t.dataset.f){ S.f[t.dataset.f] = t.value; run(); render(); return; }
  if (t.dataset.o){ S.opts[t.dataset.o] = t.type === 'number' ? +t.value : t.value; await S.store.setDoc('ctOpts', S.opts); run(); render(); return; }
  const act = t.dataset.act;
  try {
    if (act === 'heatMetric'){ S.heatMetric = t.value; render(); }
    else if (act === 'trendType'){ S.trendType = t.value; render(); }
    else if (act === 'impactWhy'){ S.impactWhy = t.value; render(); }
    else if (act === 'recallState'){ S.recallState = t.value; render(); }
    else if (act === 'importEskStatus' && t.files[0]) await importEskStatus(t.files[0]);
    else if (act === 'importPayerRules' && t.files[0]) await importPayerRules(t.files[0]);
    else if (act === 'importAppts' && t.files[0]) await importAppts(t.files[0]);
    else if (act === 'importEsk' && t.files[0]) await importEsk(t.files[0]);
    else if (act === 'tplCopy') fillTemplateDefaults();
    else if (act === 'seDone'){ const r = seRow(t.dataset.m); await seWrite({ Kind: 'Task done', Key: t.dataset.k, MRN: r.MRN, LASFIR: r.LASFIR, Value: t.checked ? 'Yes' : 'No' }); }
    else if (act === 'seRv'){ const r = seRow(t.dataset.m); await seWrite({ Kind: 'Re-verify anyway', Key: t.dataset.m, MRN: r.MRN, LASFIR: r.LASFIR, Value: t.checked ? 'On' : 'Off' }); }
    else if (act === 'seTemplates'){ S.opts.templates = t.checked ? 'On' : ''; await S.store.setDoc('ctOpts', S.opts); render(); }
    else if (act === 'seWho'){ try { localStorage.setItem('ctWho', t.value.trim()); } catch {} }
  } catch (err){ console.error(err); busy(false); toast('bad', esc(err.message || String(err)), 9000); }
});

let xlsxP = null;
function loadXLSX(){ return xlsxP || (xlsxP = new Promise(res => { if (window.XLSX) return res(true); const s = document.createElement('script'); s.src = 'assets/xlsx.full.min.js?v=' + APP_VERSION; s.onload = () => res(!!window.XLSX); s.onerror = () => { xlsxP = null; res(false); }; document.head.appendChild(s); })); }
async function readFileAoa(file){
  if (!(await loadXLSX())) throw new Error('Couldn’t load the spreadsheet reader. Check the connection and try again.');
  const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false, raw: false });
  return wb;
}
async function importAppts(file){
  busy(true, 'Reading the report…');
  const wb = await readFileAoa(file);
  let best = null;
  for (const n of wb.SheetNames){ const aoa = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }); const h = TE.findHeader(aoa); if (h){ best = { aoa, h }; break; } }
  busy(false);
  if (!best) throw new Error('That file has no Appt Date / Provider / Event / Status header row. Use the NextGen Appointments Listing report.');
  // keep only columns that have something in them
  const W = Math.max(...best.aoa.map(r => r.length)); const keepC = []; for (let c = 0; c < W; c++) if (best.aoa.some(r => String(r[c] ?? '').trim())) keepC.push(c);
  const aoa = best.aoa.map(r => keepC.map(c => r[c] ?? ''));
  const parsed = TE.parseAppts(aoa);
  if (!(await ask(`Replace the <b>${esc(S.data.ap ? S.data.ap.sheet : SHEETS.appt)}</b> tab with ${parsed.rows.length.toLocaleString()} visits (${TE.mdy(parsed.min)} – ${TE.mdy(parsed.max)})?`, 'Replace'))) return;
  busy(true, `Writing ${parsed.rows.length.toLocaleString()} visits…`);
  await S.store.replaceAppts(S.data.ap ? S.data.ap.sheet : SHEETS.appt, aoa, keepC.indexOf(best.h.cols.date) >= 0 ? keepC.indexOf(best.h.cols.date) : null);
  busy(false); toast('ok', 'Appointment list updated.'); await reload();
}
async function importEsk(file){
  const wb = await readFileAoa(file);
  const sheets = {}; for (const n of ['About', 'Providers', 'DayPart', 'Frequencies', 'Weekly']) if (wb.Sheets[n]) sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' });
  const m = TE.readAccessModel(sheets); if (m.error) throw new Error(m.error);
  delete m.weekly;   // keep the stored copy small
  S.esk = m; await S.store.setDoc('ctEsk', m); toast('ok', `Loaded the ${esc(m.modality)} access model (${m.providers.length} providers).`); S.view = 'access'; render();
}

// lookup tabs: create or top up, never overwrite what a PM typed
async function writeSlotTypes(){
  const have = new Set(S.data.st.map(r => String(r.type).toLowerCase()));
  const base = S.data.stExists ? S.data.st.map(r => [r.type, r.kind, r.accepts, r.notes || '']) : TE.DEFAULT_SLOT_TYPES.map(r => r.slice());
  const inBase = new Set(base.map(r => String(r[0]).toLowerCase()));
  for (const t of S.data.templates) for (const s of t.slots){ const k = s.type.toLowerCase(); if (!inBase.has(k)){ inBase.add(k); const g = TE.slotTypeOf(new Map(), s.type); base.push([s.type, g.kind === 'blocked' ? 'Blocked' : 'Clinical', g.any ? '*' : '', 'Added from a template; fill in Accepts']); } }
  busy(true, 'Writing Slot Types…'); await S.store.writeSheet(SHEETS.st, [['Slot type', 'Kind', 'Accepts', 'Notes']].concat(base), { keepFormat: true }); busy(false);
  toast('ok', `${SHEETS.st} tab has ${base.length} slot types.`); await reload();
}
async function writeRef(){
  const R = S.R; const roleGuess = r => { const c = TE.classify(TE.buildRef([r]), r.event); return c.kind === 'pt' ? ROLE_LABEL[c.role] : ''; };
  const rows = S.data.ref.map(r => [r.event, r.resType, r.sched, r.cls, r.sl, r.elig, r.role || roleGuess(r)]);
  const seen = new Set(rows.map(r => String(r[0]).toLowerCase()));
  const events = S.data.ap && !S.data.ap.error ? [...new Set(S.data.ap.rows.map(a => a.event).filter(Boolean))] : [];
  let added = 0;
  for (const ev of events.sort()) if (!seen.has(ev.toLowerCase())){ const g = TE.classify(new Map(), ev); rows.push([ev, 'Provider', g.kind === 'ignore' ? 'Reference/Hold Only' : g.kind === 'nonpt' ? 'Admin' : 'Clinical', g.cls, g.sl, g.kind === 'ignore' ? 'No' : 'Always', g.kind === 'pt' ? ROLE_LABEL[g.role] : '']); added++; }
  busy(true, 'Writing Ref…'); await S.store.writeSheet(SHEETS.ref, [['Event', 'Resource Type', 'Schedule Classification', 'Event Class', 'Service Line', 'Template Eligible', 'Role']].concat(rows), { keepFormat: true }); busy(false);
  toast('ok', added ? `Added ${added} events to ${SHEETS.ref}. Check their Event Class, Service Line and Role.` : `${SHEETS.ref} already lists every event.`); await reload();
}
async function writeProviders(){
  const R = S.R; const rows = S.data.prov.map(p => [p.name, p.label, p.cred, p.include ?? '']);
  const seen = new Set(rows.map(r => r[0]));
  const names = S.data.ap && !S.data.ap.error ? [...new Set(S.data.ap.rows.map(a => a.prov).filter(Boolean))].sort() : [];
  for (const n of names) if (!seen.has(n)){ const t = R && R.templates.find(x => x.provider === n); rows.push([n, t ? t.label : '', TE.credentialOf(n), TE.isPersonResource(n) ? '' : 'No']); }
  busy(true, 'Writing Providers…'); await S.store.writeSheet(SHEETS.prov, [['NextGen provider', 'Template label', 'Credential', 'Include']].concat(rows), { keepFormat: true }); busy(false);
  toast('ok', `${SHEETS.prov} tab has ${rows.length} rows.`); await reload();
}

// new template tab in the standard (provider configurable) layout
function newTemplateDialog(){
  const R = S.R; const people = R ? R.people.slice().sort() : [];
  openModal(`<h2>New template tab</h2>
    <div class="fld">Provider (as in NextGen)<select id="tpProv">${people.map(p => `<option>${esc(p)}</option>`).join('')}<option value="">Someone not on the report…</option></select></div>
    <div class="fld">Label on the template (usually last name)<input id="tpLabel"></div>
    <div class="fld">Location<input id="tpLoc"></div>
    <div class="fld">Effective from<input id="tpFrom" type="date" value="${TE.iso(TE.todaySerial())}"></div>
    <div class="row"><div class="fld" style="flex:1">Day starts<input id="tpStart" type="time" value="08:00"></div><div class="fld" style="flex:1">Day ends<input id="tpEnd" type="time" value="18:00"></div><div class="fld" style="flex:1">Slot length (min)<input id="tpStep" type="number" value="30" min="5" max="120"></div></div>
    <div class="fld">Start from<select id="tpCopy" data-act="tplCopy"><option value="">A blank grid</option>${S.data.templates.map((t, i) => `<option value="${i}">Copy ${esc(t.sheet)} · ${esc(t.label)}</option><option value="${i}:e">${esc(t.sheet)} · ${esc(t.label)} layout, empty grid</option>`).join('')}</select></div>
    <div class="small muted">Copying an existing template is how you record a revision: give the copy a new Effective From date and edit it. The old tab keeps measuring the weeks before. A copy keeps that tab’s colors, dropdowns and formulas, and its times; the start, end and slot length above only apply to a blank grid.</div>
    <div class="row" style="margin-top:12px"><button class="btn primary" data-act="makeTemplate">Create tab</button><button class="btn" data-act="close">Cancel</button></div>`);
  const sync = () => { const p = $('#tpProv').value; $('#tpLabel').value = p ? (TE.lastName(p).replace(/^./, c => c.toUpperCase())) : ''; const locs = {}; (S.data.ap.rows || []).filter(a => a.prov === p && a.loc).forEach(a => { locs[a.loc] = (locs[a.loc] || 0) + 1; }); $('#tpLoc').value = Object.entries(locs).sort((a, b) => b[1] - a[1]).map(x => x[0])[0] || ''; };
  $('#tpProv').addEventListener('change', sync); sync();
}
function fillTemplateDefaults(){ const i = $('#tpCopy').value; if (i === '') return; const t = S.data.templates[parseInt(i, 10)]; $('#tpStart').value = hm(t.start); $('#tpEnd').value = hm(t.end); $('#tpStep').value = t.step; if (t.location) $('#tpLoc').value = t.location; }
const hm = m => String(Math.floor(m / 60)).padStart(2, '0') + ':' + String(m % 60).padStart(2, '0');
async function makeTemplate(){
  const prov = $('#tpProv').value, label = $('#tpLabel').value.trim(), loc = $('#tpLoc').value.trim(), from = TE.parseDate($('#tpFrom').value);
  const start = TE.parseTime($('#tpStart').value), end = TE.parseTime($('#tpEnd').value), step = +$('#tpStep').value || 30;
  if (!label) throw new Error('Give the template a label.'); if (start == null || end == null || end <= start) throw new Error('Check the start and end times.');
  const pick = $('#tpCopy').value; const src = pick === '' ? null : S.data.templates[parseInt(pick, 10)];
  const sheetName = base => { let name = base, n = 2; while (S.data.names.some(x => x.toLowerCase() === name.toLowerCase())) name = (base.slice(0, 28) + ' ' + n++); return name; };
  if (src && S.data.grids[src.sheet]){
    // copy the PM's own tab so its colors, dropdowns and formulas come along
    const sibs = S.data.templates.filter(t => t.sheet === src.sheet);
    const plan = TE.copyPlan(src, sibs, S.data.grids[src.sheet], { label, title: src.title && src.at.title ? (prov || label) : '', location: loc, from, today: TE.todaySerial(), empty: pick.endsWith(':e') });
    if (plan.error) throw new Error(plan.error);
    const name = sheetName(`TPL ${label} ${from != null ? TE.iso(from) : ''}`.replace(/[\\/?*\[\]:]/g, '').slice(0, 31).trim());
    closeModal(); busy(true, 'Copying the template tab…');
    await S.store.copyTemplate(Object.assign(plan, { src: src.sheet, sheet: name, dvSource: `='${SHEETS.st}'!$A$2:$A$300` }));
    busy(false); toast('ok', `Created “${esc(name)}” from “${esc(src.sheet)}” with its formatting. ${plan.clearSlots ? 'Fill the grid' : 'Edit the slots'}, then Refresh.`); await reload();
    return;
  }
  const times = []; for (let m = start; m < end; m += step) times.push(m);
  const cells = [[], [prov || label], ['Effective From:', from ?? '', '', 'Effective To:', ''], ['Location:', loc, 'Weekly Clinic Template', '', 'Last Modified:', TE.todaySerial()], ['Provider:', label, label, label, label, label], ['Time', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']];
  for (const m of times){ const row = [m / 1440]; for (let d = 1; d <= 5; d++){ const s = src && src.slots.find(x => x.dow === d && x.min === m); row.push(s ? s.type : ''); } cells.push(row); }
  cells[3][8] = 'Event Types'; cells[3][9] = 'Slots'; cells[3][10] = 'Hours'; ['M', 'T', 'W', 'TH', 'F'].forEach((x, i) => { cells[3][11 + i] = x; });
  const base = `TPL ${label} ${from != null ? TE.iso(from) : ''}`.replace(/[\\/?*\[\]:]/g, '').slice(0, 31).trim();
  const name = sheetName(base);
  closeModal(); busy(true, 'Creating the template tab…');
  await S.store.createTemplate({ sheet: name, cells, times, step });
  busy(false); toast('ok', `Created “${esc(name)}”. Fill the grid from the dropdowns, then Refresh.`); await reload();
}

// report tabs: counts only
async function writeReports(){
  const R = S.R; if (!R) return;
  const P = x => x == null ? '' : Math.round(x * 1000) / 1000;
  const scope = `${TE.mdy(R.from)} – ${TE.mdy(R.to)}; as of ${TE.mdy(R.asOf)}${S.f.loc ? '; ' + S.f.loc : ''}${S.f.prov ? '; ' + S.f.prov : ''}${S.f.sl ? '; ' + S.f.sl : ''}`;
  busy(true, 'Writing report tabs…');
  const pf = { numFmt: {} };
  const provH = ['Provider', 'Credential', 'Has template', 'Clinical slots', 'Filled %', 'Kept use %', 'Template fit %', 'Open slots', 'Slots to admin', 'PTO slots', 'Visits', 'Kept', 'No-show %', 'Cancel %', 'Late cancel %', 'Visits in blocked time', 'Visits outside template hours', 'Double-booked', 'Booked ahead %', 'Scope'];
  const pr = [...R.byProv].sort((a, b) => b.supply - a.supply).map(p => [p.key, credOf(p.key), R.tByProv.has(p.key) ? 'Yes' : 'No', p.supply, P(p.fill), P(p.util), P(p.fit), p.open, p.admin, p.pto, p.aAll, p.aKept, P(p.nsRate), P(p.cxRate), R.hasLate ? P(p.lcRate) : '', p.inBlocked, p.offTemplate, p.dbl, P(p.futFill), scope]);
  const t = R.total; pr.push(['All', '', '', t.supply, P(t.fill), P(t.util), P(t.fit), t.open, t.admin, t.pto, t.aAll, t.aKept, P(t.nsRate), P(t.cxRate), R.hasLate ? P(t.lcRate) : '', t.inBlocked, t.offTemplate, t.dbl, P(t.futFill), scope]);
  const pc = c => Object.fromEntries(c.map(i => [i, '0.0%']));
  await S.store.writeSheet('TA Summary', [provH].concat(pr), { numFmt: pc([4, 5, 6, 12, 13, 14, 18]) });
  await S.store.writeSheet('TA Slot Types', [['Slot type', 'Slots', 'Filled %', 'Template fit %', 'Kept as intended', 'Kept, other visit', 'No-show', 'Open', 'To admin', 'PTO', 'Booked ahead %', 'Most common other visits', 'Kept, release booking']].concat(R.byType.filter(r => r.supply || r.futSupply).sort((a, b) => b.supply - a.supply).map(r => [r.key, r.supply, P(r.fill), P(r.fit), r.keptFit, r.keptOff, r.ns, r.open, r.admin, r.pto, P(r.futFill), Object.entries(R.offMix[r.key] || {}).sort((a, b) => b[1] - a[1]).slice(0, 3).map(x => `${x[0]} (${x[1]})`).join(', '), r.keptRel])), { numFmt: pc([2, 3, 10]) });
  const vr = (dim, rows) => rows.filter(r => r.aAll).sort((a, b) => b.aAll - a.aAll).map(r => [dim, r.key, r.aAll, r.aKept, r.aNS, r.aCx, P(r.nsRate), P(r.cxRate), R.hasLate ? P(r.lcRate) : '', r.leadMed == null ? '' : r.leadMed]);
  await S.store.writeSheet('TA Visit Types', [['Group', 'Value', 'Visits', 'Kept', 'No-show', 'Cancelled', 'No-show %', 'Cancel %', 'Late cancel %', 'Median lead days']].concat(vr('Service line', R.bySL), vr('Event class', R.byClass), vr('Event', R.byEvent), vr('Day', R.byDow), vr('Hour', R.byHour.filter(h => typeof h.key === 'number').map(h => Object.assign({}, h, { key: TE.hourLabel(h.key) })))), { numFmt: pc([6, 7, 8]) });
  await S.store.writeSheet('TA Weekly', [['Week of (Sun)', 'Clinical slots', 'Filled %', 'Template fit %', 'Visits', 'Kept', 'No-show %', 'Cancel %', 'Late cancel %', 'Visits outside template', 'Future template slots', 'Booked ahead %']].concat(R.byWeek.map(w => [w.key, w.supply, P(w.fill), P(w.fit), w.aAll, w.aKept, P(w.nsRate), P(w.cxRate), R.hasLate ? P(w.lcRate) : '', w.inBlocked + w.offTemplate, w.futSupply, P(w.futFill)])), { numFmt: Object.assign({ 0: 'm/d/yyyy' }, pc([2, 3, 6, 7, 8, 11])) });
  const hours = [...new Set(R.heat.map(h => +String(h.key).split('|')[1]))].sort((a, b) => a - b); const H = new Map(R.heat.map(h => [h.key, h]));
  await S.store.writeSheet('TA Heatmap', [['Filled % by day and hour'].concat(hours.map(TE.hourLabel))].concat([1, 2, 3, 4, 5].map(d => [TE.DAYS[d]].concat(hours.map(h => { const c = H.get(d + '|' + h); return c && c.supply ? P(c.fill) : ''; })))), { numFmt: Object.fromEntries(hours.map((h, i) => [i + 1, '0%'])) });
  await S.store.writeSheet('TA 3NA', [['As of', 'Provider', 'Slot type', 'Next available (days)', '3NA (days)', 'Open slots ahead', 'Third open slot']].concat(R.threeNa.map(x => [R.asOf, x.prov, x.type, x.next ?? '', x.third ?? '', x.open, x.thirdDate ?? ''])), { numFmt: { 0: 'm/d/yyyy', 6: 'm/d/yyyy' } });
  await S.store.writeSheet('TA Exceptions', [['Date', 'Time', 'Provider', 'Event', 'Status', 'Why', 'Minutes outside template']].concat(R.exceptions.map(e => [e.date, TE.hhmm(e.time), e.prov, e.event, e.status, e.why, e.min])), { numFmt: { 0: 'm/d/yyyy' } });
  const C = S.C; let extra = '';
  if (C){
    await S.store.writeSheet('TA Cadence', [['Provider', 'Service line', 'Patients', 'Kept visits', 'Visits per patient', 'Returning patients', 'Mean gap (days)', 'Median gap (days)', 'Mode (wks)', 'Mode share', 'No-shows', 'Cancels', 'No-show %'].concat(C.cad.hist, ['Scope'])].concat(C.cad.rows.map(r => [r.prov, r.bucket, r.patients, r.kept, P(r.perPt), r.returning, P(r.mean), r.median ?? '', r.modeWk ?? '', P(r.modeShare), r.ns, r.cx, P(r.nsRate)].concat(r.hist, [scope]))), { numFmt: pc([9, 12]) });
    await S.store.writeSheet('TA Funnel', [['Provider', 'Service line', 'Consult slots', 'Consult hours', 'Consults kept', 'Consult fill %', 'Started', 'Pending', 'Conversion %', 'Median days to start', 'Start visit (min)', 'Start hours offered', 'Start hours used', 'Start hours needed', 'Offered vs needed', 'Scope']].concat(C.fun.map(r => [r.prov, r.sl, r.consultSlots, P(r.consultHours), r.consults, P(r.consultFill), r.converted, r.pending, P(r.rate), r.medDays ?? '', r.startLen, P(r.startHoursOffered), P(r.startHoursUsed), P(r.startHoursNeeded), r.flag, scope])), { numFmt: pc([5, 8]) });
    await S.store.writeSheet('TA Flex Usage', [['Provider', 'Slot type', 'Accepts', 'Offered', 'Used', 'Use %', 'Consult', 'Start', 'Treatment', 'Follow-up', 'Release bookings', 'Release share', 'Scope']].concat(C.flex.map(r => [r.prov, r.type, r.accepts.map(x => ROLE_LABEL[x]).join(', ') || 'Anything', r.offered, r.used, P(r.useRate), r.roles.consult || 0, r.roles.start || 0, r.roles.treatment || 0, r.roles.followup || 0, r.release, P(r.releaseShare), scope])), { numFmt: pc([5, 11]) });
    extra = ', TA Cadence, TA Funnel, TA Flex Usage';
  }
  if (S.P){
    await S.store.writeSheet('TA ESK Payer Rules', [['Payer on the report', 'Rule', 'ESK visits', 'With an MD', 'With an NP', 'NP may bill', 'Besse', 'Outside the rules', 'Upcoming outside the rules', 'Scope']].concat(S.P.byPayer.map(g => [g.payer, g.rule, g.visits, g.md, g.np, g.npSteps || '', g.rule ? g.besse : '', g.blocked, g.upcomingBlocked, scope])));
    extra += ', TA ESK Payer Rules';
  }
  busy(false); toast('ok', `Wrote TA Summary, TA Slot Types, TA Visit Types, TA Weekly, TA Heatmap, TA 3NA, TA Exceptions${extra}.`);
}
// patient-level tabs: LASFIR codes only, never names or record numbers
async function writePatientTabs(){
  const C = S.C; if (!C) return;
  if (!(await ask('Write PT Recall, PT Slot Matches and PT Cadence Patients? They list patients by LASFIR code (no names or record numbers), so share this workbook only with people who may see patient-level schedules.', 'Write tabs'))) return;
  busy(true, 'Writing patient tabs…');
  await S.store.writeSheet('PT Recall', [['As of', 'LASFIR', 'Provider', 'Service line', 'Clinic', 'Last seen', 'Days since', 'Usual interval (days)', 'State', 'Since last visit', 'Next booked', 'ESK Tracker', 'Return date', 'Pause reason']].concat(C.rc.map(r => [S.R.asOf, r.code || '(check name)', r.prov, r.bucket, r.loc, r.last, r.daysSince, r.exp ?? '', (STATE[r.state] || [0, r.state])[1], r.post, r.next ?? '', r.tracker, r.ret ?? '', r.reason])), { numFmt: { 0: 'm/d/yyyy', 5: 'm/d/yyyy', 10: 'm/d/yyyy', 12: 'm/d/yyyy' } });
  await S.store.writeSheet('PT Slot Matches', [['LASFIR', 'Service line', 'State', 'Usual provider', 'Slot provider', 'Slot date', 'Slot time', 'Slot type', 'Why it fits']].concat(C.matches.map(x => [x.code || '(check name)', x.bucket, (STATE[x.state] || [0, x.state])[1], x.patientProv, x.prov, x.date, TE.hhmm(x.min), x.type, x.why])), { numFmt: { 5: 'm/d/yyyy' } });
  await S.store.writeSheet('PT Cadence Patients', [['LASFIR', 'Provider', 'Service line', 'Kept visits', 'First seen', 'Last seen', 'Mean gap (days)', 'Last gap (days)', 'Tele %']].concat(C.cad.patients.map(p => [p.code || '(check name)', p.prov, p.bucket, p.kept, p.first, p.last, p.meanGap == null ? '' : Math.round(p.meanGap * 10) / 10, p.lastGap ?? '', p.telePct == null ? '' : Math.round(p.telePct * 1000) / 1000])), { numFmt: { 4: 'm/d/yyyy', 5: 'm/d/yyyy', 8: '0%' } });
  busy(false); toast('ok', 'Wrote PT Recall, PT Slot Matches and PT Cadence Patients.');
}
// the call list grouped by clinic, as a table to paste into an email or Teams (LASFIR only, like the Confirmations tool)
async function copyOutreach(){
  const C = S.C; if (!C) return;
  const list = (S.recallState === 'call' ? C.rc.filter(r => TC.CALL.has(r.state)) : S.recallState ? C.rc.filter(r => r.state === S.recallState) : C.rc);
  if (!list.length){ toast('warn', 'Nobody on this list.'); return; }
  const first = new Map(); for (const x of C.matches){ const k = x.code + '|' + x.bucket; if (!first.has(k)) first.set(k, x); }
  const cols = [['Pat Name', r => r.code || '(check name)'], ['Service line', r => r.bucket], ['Provider', r => r.prov], ['Last seen', r => TE.mdy(r.last)], ['Days since', r => r.daysSince], ['Usual interval', r => r.exp == null ? '' : r.exp + ' d'], ['State', r => (STATE[r.state] || [0, r.state])[1]], ['Since last visit', r => r.post], ['ESK Tracker', r => r.tracker], ['First open slot', r => { const m = first.get(r.code + '|' + r.bucket); return m ? `${TE.mdy(m.date)} ${TE.hhmm(m.min)} ${m.type} (${m.prov.split(',')[0]})` : ''; }]];
  const groups = new Map(); for (const r of list){ const k = r.loc || 'No location'; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(r); }
  const th = 'style="background:#1f4e79;color:#fff;text-align:left;padding:6px 8px;font-family:Aptos Display,Aptos,Calibri,Arial,sans-serif;border:1px solid #ccc;font-size:12px"';
  const td = 'style="padding:6px 8px;font-family:Aptos Display,Aptos,Calibri,Arial,sans-serif;border:1px solid #ccc;font-size:12px;vertical-align:top"';
  let html = '', text = '';
  for (const [loc, rows] of [...groups.entries()].sort((a, b) => a[0].localeCompare(b[0]))){
    html += `<p style="margin:14px 0 6px;font-family:Aptos,Calibri,Arial,sans-serif"><b>${esc(loc)}</b> (${rows.length})</p><table style="border-collapse:collapse"><tr>${cols.map(c => `<th ${th}>${esc(c[0])}</th>`).join('')}</tr>${rows.map(r => `<tr>${cols.map(c => `<td ${td}>${esc(c[1](r))}</td>`).join('')}</tr>`).join('')}</table>`;
    text += `${loc} (${rows.length})\n${cols.map(c => c[0]).join('\t')}\n${rows.map(r => cols.map(c => c[1](r)).join('\t')).join('\n')}\n\n`;
  }
  const done = () => toast('ok', `Copied ${list.length} patients in ${groups.size} clinic${groups.size === 1 ? '' : 's'}.`);
  try { if (navigator.clipboard && window.ClipboardItem){ await navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([html], { type: 'text/html' }), 'text/plain': new Blob([text], { type: 'text/plain' }) })]); done(); return; } } catch {}
  // the clipboard API is often blocked inside Office: select a hidden copy of the table and use the copy command
  const box = document.createElement('div'); box.contentEditable = 'true'; box.style.cssText = 'position:fixed;left:-9999px;top:0'; box.innerHTML = html; document.body.appendChild(box);
  const rg = document.createRange(); rg.selectNodeContents(box); const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(rg);
  let ok = false; try { ok = document.execCommand('copy'); } catch {}
  sel.removeAllRanges(); box.remove();
  if (ok) done(); else toast('bad', 'Couldn’t copy here. Use Write patient tabs and copy from PT Recall instead.', 9000);
}
async function importEskStatus(file){
  const wb = await readFileAoa(file);
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' });
  const r = TC.readEskStatus(aoa); if (r.error) throw new Error(r.error);
  busy(true, 'Saving ESK Tracker status…');
  await S.store.writeSheet(SHEETS.eskst, [['LASFIR', 'Clinic', 'Status', 'Last Dose Date', 'Pause Reason', 'Expected Return Date']].concat(r.rows.map(x => [x.code, x.clinic, x.status, x.lastDose ?? '', x.reason, x.ret ?? ''])), { numFmt: { 3: 'm/d/yyyy', 5: 'm/d/yyyy' } });
  await S.store.setDoc('ctEskAt', TE.todaySerial());
  busy(false); toast('ok', `Loaded ESK Tracker status for ${r.rows.length} patients${r.skipped ? ` (${r.skipped} rows without a LASFIR code skipped)` : ''}.`); await reload();
}
async function importPayerRules(file){
  const wb = await readFileAoa(file);
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: '' });
  const r = PayerRules.read(aoa); if (r.error) throw new Error(r.error);
  if (S.data.payerRules && !S.data.payerRules.error && !(await ask(`Replace the <b>${esc(SHEETS.pcr)}</b> tab (${S.data.payerRules.rules.length} rows) with this file (${r.rules.length} rows)? Edits made on the tab will be lost.`, 'Replace'))) return;
  busy(true, 'Saving payer rules…');
  await S.store.writeSheet(SHEETS.pcr, PayerRules.toRows(r.rules));
  busy(false); toast('ok', `Loaded ${r.rules.length} payer rules${r.skipped ? ` (${r.skipped} rows without MD or NP skipped)` : ''}.`); await reload();
}
async function saveSnapshot(){
  const R = S.R; if (!R) return;
  const keep = (S.data.hist || []).filter(h => h.date !== R.asOf);
  const rows = keep.map(h => [h.date, h.prov, h.type, h.next ?? '', h.third ?? '', h.open]).concat(R.threeNa.map(x => [R.asOf, x.prov, x.type, x.next ?? '', x.third ?? '', x.open]));
  rows.sort((a, b) => a[0] - b[0] || String(a[1]).localeCompare(String(b[1])) || String(a[2]).localeCompare(String(b[2])));
  busy(true, 'Saving the 3NA snapshot…');
  await S.store.writeSheet(SHEETS.hist, [['Snapshot date', 'Provider', 'Slot type', 'Next available (days)', '3NA (days)', 'Open slots ahead']].concat(rows), { numFmt: { 0: 'm/d/yyyy' } });
  busy(false); toast('ok', `Saved 3NA as of ${TE.mdy(R.asOf)}${S.f.prov || S.f.loc ? ' (with the current filters)' : ''}.`); await reload();
}

// ---------- install page outside Excel ----------
function notInExcelHTML(){
  return `<main style="max-width:640px"><div class="card"><h2><span class="brand"><span class="dot">CT</span></span> Clinic Template Tool</h2>
    <p>This page is an Excel add-in. It measures provider templates against the NextGen appointment list inside a template workbook.</p>
    <ol><li>Click <b>Download manifest</b>.</li><li>Open the workbook in Excel on the web → <b>Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in</b>, and choose the file.</li><li>A <b>Template Tool</b> button appears on the Home tab.</li></ol>
    <button class="btn primary" data-act="dlManifest">Download manifest</button>
    <div class="small muted" style="margin-top:10px">Data stays in the workbook. This site only serves the app’s code.</div></div></main>`;
}
async function downloadManifest(){
  const base = location.href.replace(/[?#].*$/, '').replace(/[^/]*$/, '').replace(/\/$/, ''); const host = new URL(base).origin;
  const x = (await (await fetch('manifest.xml', { cache: 'no-store' })).text()).replaceAll('__BASE__', base).replaceAll('__HOST__', host);
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([x], { type: 'application/xml' })); a.download = 'clinic-template-tool-manifest.xml'; a.click();
}

(function boot(){
  S.mode = /[?&]mode=dialog\b/.test(location.search) ? 'dialog' : 'pane';
  if (window.__TEST_STORE__){ S.store = window.__TEST_STORE__; S.canFull = true; reload(); return; }   // headless tests
  if (typeof Office === 'undefined'){ $('#app').innerHTML = notInExcelHTML(); return; }
  let settled = false;
  Office.onReady(info => {
    settled = true;
    if (S.mode === 'dialog'){ S.store = RemoteStore(); reload(); return; }   // the dialog has no Excel host of its own
    if (info.host !== Office.HostType.Excel){ $('#app').innerHTML = notInExcelHTML(); return; }
    try { S.canFull = Office.context.requirements.isSetSupported('DialogApi', '1.2'); } catch { S.canFull = false; }
    S.store = OfficeStore(); reload();
  });
  setTimeout(() => { if (!settled) $('#app').innerHTML = notInExcelHTML(); }, 4000);
})();
