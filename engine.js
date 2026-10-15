// Clinic Template Tool engine: parses provider template grids and the NextGen appointment list,
// lays the template over the calendar, and measures how the schedule actually used it.
// Pure functions only (no Excel calls) so tests can run it under Node.
// Dates are Excel serial day numbers (whole days); times are minutes after midnight.
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TE = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const DAY_RE = [/^sun(day)?$/i, /^mon(day)?$/i, /^tue(s|sday)?$/i, /^wed(nesday)?$/i, /^thu(r|rs|rsday)?$/i, /^fri(day)?$/i, /^sat(urday)?$/i];
  const lc = s => String(s == null ? '' : s).trim().toLowerCase();
  const txt = s => String(s == null ? '' : s).trim();

  // ---------- dates and times ----------
  const dow = serial => ((serial - 1) % 7 + 7) % 7;              // 0 = Sunday (valid for serials after Feb 1900)
  const weekStart = serial => serial - dow(serial);               // weeks run Sunday to Saturday
  function serialFromYMD(y, m, d){ return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5); }
  function ymd(serial){ const t = new Date(Date.UTC(1899, 11, 30) + serial * 864e5); return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() }; }
  function iso(serial){ if (serial == null || isNaN(serial)) return ''; const x = ymd(Math.floor(serial)); return `${x.y}-${String(x.m).padStart(2, '0')}-${String(x.d).padStart(2, '0')}`; }
  function mdy(serial){ if (serial == null || isNaN(serial)) return ''; const x = ymd(Math.floor(serial)); return `${x.m}/${x.d}/${x.y}`; }
  function todaySerial(){ const n = new Date(); return serialFromYMD(n.getFullYear(), n.getMonth() + 1, n.getDate()); }

  // time of day -> minutes. Accepts Excel fractions, "1:30 P", "1:30 PM", "13:30", "9a".
  function parseTime(v){
    if (v == null || v === '') return null;
    if (typeof v === 'number'){ const f = v - Math.floor(v); return Math.round(f * 1440) % 1440; }
    const m = String(v).trim().match(/^(\d{1,2})(?::(\d{2}))?(?::\d{2})?\s*([ap])?\.?\s*m?\.?$/i);
    if (!m) return null;
    let h = +m[1]; const mi = +(m[2] || 0); const ap = (m[3] || '').toLowerCase();
    if (ap === 'p' && h < 12) h += 12; if (ap === 'a' && h === 12) h = 0;
    if (h > 23 || mi > 59) return null;
    return h * 60 + mi;
  }
  // date (or date-time) -> serial with fraction. Accepts Excel serials, m/d/yyyy [h:mm AM], m/d/yy, yyyy-mm-dd.
  function parseDateTime(v){
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v > 20000 && v < 80000 ? v : null;
    if (v instanceof Date) return serialFromYMD(v.getFullYear(), v.getMonth() + 1, v.getDate()) + (v.getHours() * 60 + v.getMinutes()) / 1440;
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:[ T]+(.+))?$/), y, mo, d, rest;
    if (m){ mo = +m[1]; d = +m[2]; y = +m[3]; if (y < 100) y += 2000; rest = m[4]; }
    else if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T]+(.+))?$/))){ y = +m[1]; mo = +m[2]; d = +m[3]; rest = m[4]; }
    else return null;
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    let out = serialFromYMD(y, mo, d);
    if (rest){ const t = parseTime(rest.replace(/(\d)\s*(am|pm)$/i, '$1 $2').replace(/\s+(a|p)\.?m\.?$/i, ' $1')); if (t != null) out += t / 1440; }
    return out;
  }
  const parseDate = v => { const x = parseDateTime(v); return x == null ? null : Math.floor(x); };
  const hhmm = min => { if (min == null) return ''; const h = Math.floor(min / 60), m = min % 60; return `${((h + 11) % 12) + 1}:${String(m).padStart(2, '0')} ${h < 12 ? 'AM' : 'PM'}`; };
  const hourLabel = h => `${((h + 11) % 12) + 1} ${h < 12 ? 'AM' : 'PM'}`;

  // ---------- provider names ----------
  const CRED_WORDS = new Set(['md', 'do', 'np', 'pmhnp', 'bc', 'fnp', 'aprn', 'pa', 'pac', 'dr', 'phd', 'psyd', 'rn', 'mba', 'mph', 'jr', 'sr', 'ii', 'iii']);
  const nameTokens = s => lc(s).replace(/[^a-z\s,'-]/g, ' ').split(/[\s,]+/).map(t => t.replace(/^['-]+|['-]+$/g, '')).filter(t => t.length >= 2 && !CRED_WORDS.has(t));
  // NextGen lists people as "Last CRED, First" and resources (chairs, rooms) without a comma.
  const isPersonResource = s => /,/.test(s) && !/\b(magv|magventure|brainsway|chair|room|ivk\s*\d|esk\s*\d|tms\s*\d|device)\b/i.test(s);
  function credentialOf(s){ const t = ' ' + lc(s).replace(/[^a-z]/g, ' ') + ' '; if (/ (md|do) /.test(t)) return 'MD'; if (/ (pmhnp|np|fnp|aprn|pmhnp bc) /.test(t)) return 'NP'; if (/ (pa|pac) /.test(t)) return 'PA'; return ''; }
  const lastName = s => (nameTokens(String(s).split(',')[0].replace(/\b(md|do|np|pmhnp|pmhnp-bc|bc|dr)\b\.?/ig, ''))[0]) || nameTokens(s)[0] || '';
  // Template label ("Patel", "Jordan Okafor, NP", "Samuel Reyes, MD") -> NextGen provider name.
  function matchProvider(label, names){
    const toks = nameTokens(label); if (!toks.length) return null;
    const exact = names.find(n => lc(n) === lc(label)); if (exact) return exact;
    const hits = names.filter(n => toks.includes(lastName(n)));
    if (hits.length === 1) return hits[0];
    if (hits.length > 1){ const best = hits.map(n => ({ n, s: nameTokens(n).filter(t => toks.includes(t)).length })).sort((a, b) => b.s - a.s); if (best[0].s > best[1].s) return best[0].n; }
    return null;
  }

  // ---------- template grids ----------
  // Finds every weekly grid on a sheet: a row with Monday..Friday headers, a time column to its left,
  // and event types in the cells. Header labels (Location:, Provider:, Last Modified:, Effective From:,
  // Effective To:) are read from the few rows above the grid.
  const LABELS = { location: /^location\s*:?$/i, provider: /^provider\s*:?$/i, modified: /^last\s*modified\s*:?$/i, from: /^effective\s*(from|start|date)?\s*:?$/i, to: /^effective\s*(to|through|thru|end)\s*:?$/i };
  function parseTemplates(sheet, values, formats){
    const out = []; if (!values || !values.length) return out;
    const R = values.length, C = Math.max(...values.map(r => r.length));
    const cell = (r, c) => (r >= 0 && r < R && c >= 0 && values[r] && c < values[r].length) ? values[r][c] : '';
    const used = new Set();
    for (let r = 0; r < R; r++){
      for (let c = 1; c < C; c++){
        if (used.has(r + ':' + c) || !DAY_RE[1].test(txt(cell(r, c)))) continue;
        // weekday header row: Monday at c, then consecutive weekdays
        const cols = []; for (let k = 0; k < 7 && c + k < C; k++){ const v = txt(cell(r, c + k)); const d = DAY_RE.findIndex(re => re.test(v)); if (d < 0) break; cols.push({ c: c + k, d }); }
        if (cols.length < 3 || cols[1].d !== 2) continue;
        cols.forEach(x => used.add(r + ':' + x.c));
        const tc = c - 1; const slots = []; const times = [];
        let rr = r + 1, blankRun = 0;
        for (; rr < R && rr < r + 80; rr++){
          const t = parseTime(cell(rr, tc));
          if (t == null || (typeof cell(rr, tc) === 'number' && cell(rr, tc) >= 1)){ if (times.length){ if (++blankRun > 0) break; } continue; }
          blankRun = 0; times.push({ r: rr, t });
        }
        if (times.length < 2) continue;
        const steps = times.slice(1).map((x, i) => x.t - times[i].t).filter(x => x > 0);
        const step = steps.length ? steps.sort((a, b) => a - b)[Math.floor(steps.length / 2)] : 30;
        times.forEach((x, i) => { cols.forEach(col => { const v = txt(cell(x.r, col.c)); if (v) slots.push({ dow: col.d, min: x.t, len: step, type: v, addr: [x.r, col.c] }); }); });
        // header labels
        const hdr = {}; let title = '';
        for (let hr = Math.max(0, r - 5); hr <= r; hr++){
          for (let hc = Math.max(0, tc - 1); hc <= tc + cols.length + 2; hc++){
            const v = txt(cell(hr, hc)); if (!v) continue;
            for (const [k, re] of Object.entries(LABELS)) if (re.test(v) && hdr[k] == null){
              for (let j = hc + 1; j <= hc + 8; j++){ const w = cell(hr, j); if (w !== '' && w != null && !Object.values(LABELS).some(q => q.test(txt(w)))){ hdr[k] = w; break; } }
              if (hdr[k] == null) hdr[k] = '';
            }
          }
        }
        for (let hr = Math.max(0, r - 6); hr < r && !title; hr++){ const v = txt(cell(hr, tc)) || txt(cell(hr, tc + 1)); if (v && /[a-z]{2,}/i.test(v) && !Object.values(LABELS).some(q => q.test(v)) && nameTokens(v).length && !/template|schedule/i.test(v)) title = v; }
        const label = txt(hdr.provider) || title;
        out.push({
          sheet, anchor: [r, c], label, title, location: txt(hdr.location),
          modified: parseDate(hdr.modified), effFrom: parseDate(hdr.from), effTo: parseDate(hdr.to),
          step, start: times[0].t, end: times[times.length - 1].t + step, days: [...new Set(slots.map(s => s.dow))].sort(), slots,
        });
      }
    }
    return out;
  }

  // ---------- reference tables ----------
  const NONPT_CLASSES = new Set(['admin', 'meeting', 'pto', 'clinical training', 'training', 'lunch']);
  function buildRef(rows){
    const m = new Map();
    for (const r of rows || []){ const ev = lc(r.event); if (!ev) continue; m.set(ev, { event: txt(r.event), resType: txt(r.resType), sched: txt(r.sched), cls: txt(r.cls) || txt(r.event), sl: txt(r.sl), elig: txt(r.elig) }); }
    return m;
  }
  // kind: 'pt' (a patient visit), 'nonpt' (admin/meeting/PTO booked on the provider), 'ignore' (holds, reference rows)
  function classify(ref, event){
    const r = ref.get(lc(event));
    if (r){
      const s = lc(r.sched), el = lc(r.elig), c = lc(r.cls);
      let kind = 'pt';
      if (/^reference/.test(s) || el === 'no' || el === 'never') kind = 'ignore';
      else if (NONPT_CLASSES.has(c) || s === 'admin') kind = 'nonpt';
      return { kind, cls: r.cls || r.event, sl: r.sl || 'Unassigned', mapped: true, pto: c === 'pto' };
    }
    const e = lc(event);
    if (/\bhold\b|coverage|on call|in clinic/.test(e)) return { kind: 'ignore', cls: txt(event), sl: 'Unassigned', mapped: false };
    if (/\b(admin|meeting|lunch|pto|training|peer to peer)\b/.test(e)) return { kind: 'nonpt', cls: txt(event), sl: 'N/A', mapped: false, pto: /\bpto\b/.test(e) };
    return { kind: 'pt', cls: txt(event), sl: guessServiceLine(e), mapped: false };
  }
  function guessServiceLine(e){
    e = lc(e);
    if (/\b(esk|esketamine|spravato)\b/.test(e)) return 'ESK';
    if (/\b(ivk|iv ketamine|ketamine)\b/.test(e)) return 'IVK';
    if (/\b(tms|mt\b|magv|brainsway|theta|mindfulone|mindfulconnect)/.test(e)) return 'TMS';
    if (/new patient|transition|med manage|telepsych|adhd/.test(e)) return 'Medication Management';
    return 'Unassigned';
  }
  // Slot types: "Accepts" is a list of terms separated by commas (any term may match); a term joins
  // tokens with + (all must match). A token matches the visit's event class, a service line, or the
  // exact event name. * accepts any patient visit.
  function buildSlotTypes(rows){
    const m = new Map();
    for (const r of rows || []){
      const t = lc(r.type); if (!t) continue;
      const kind = /^block|non|no$/i.test(txt(r.kind)) ? 'blocked' : 'clinical';
      const terms = txt(r.accepts).split(/[,;\n]/).map(x => x.split('+').map(lc).filter(Boolean)).filter(x => x.length);
      m.set(t, { type: txt(r.type), kind, terms, any: terms.some(x => x.includes('*')) });
    }
    return m;
  }
  function slotTypeOf(stMap, type){
    const st = stMap.get(lc(type)); if (st) return st;
    const e = lc(type);
    const blocked = /\b(admin|lunch|meeting|break|blocked|pto|out|off|huddle|training)\b/.test(e);
    return { type: txt(type), kind: blocked ? 'blocked' : 'clinical', terms: [[e]], any: /open/.test(e), unknown: true };
  }
  function accepts(st, ap){
    if (st.kind !== 'clinical') return false;
    if (st.any) return true;
    const toks = new Set([lc(ap.cls), lc(ap.event), ...String(ap.sl || '').split(',').map(lc)].filter(Boolean));
    return st.terms.some(term => term.every(t => t === '*' || toks.has(t)));
  }

  // ---------- appointment list ----------
  const COLS = {
    date: /^(appt\.?\s*)?date$|^appointment date$/i, time: /^(start\s*)?time$|^appt\.?\s*time$|^start$/i, end: /^end\s*(tm|time)?$/i,
    dur: /^duration|^length|^mins?$/i, prov: /^(provider|resource|rendering|rendering provider|resource name)$/i, event: /^(event|appt\.?\s*type|appointment type|event name|type)$/i,
    status: /^(status|appt\.?\s*status|sts)$/i, loc: /^(location|site|clinic|loc)$/i,
    booked: /^(created|create(d)?\s*(date|on|dt)|date\s*created|scheduled\s*(on|date)|date\s*scheduled|booked(\s*(on|date))?|appt\.?\s*made|made\s*on|entered(\s*(on|date))?|entry\s*date|create\s*timestamp)$/i,
    bookedTime: /^(create(d)?\s*time|scheduled\s*time|booked\s*time|entry\s*time)$/i,
    cancel: /^(cancel(l)?(ed)?\s*(date|on|dt|date\/time|timestamp)?|date\s*cancel(l)?(ed)?|cxl\s*(date|dt))$/i,
    cancelTime: /^(cancel(l)?(ed)?\s*time|cxl\s*time)$/i,
    reason: /^(cancel(l)?(ation)?\s*reason|reason|cxl\s*reason|cancel\s*code)$/i,
  };
  function findHeader(values){
    for (let r = 0; r < Math.min(values.length, 40); r++){
      const row = values[r] || []; const map = {};
      row.forEach((v, c) => { const h = txt(v); if (!h) return; for (const [k, re] of Object.entries(COLS)) if (map[k] == null && re.test(h)){ map[k] = c; break; } });
      if (map.date != null && map.event != null && map.status != null && map.prov != null) return { row: r, cols: map, headers: row.map(txt) };
    }
    return null;
  }
  // The NextGen report prints its run time in the first rows ("9/14/26 12:59 PM"); that's when Expected
  // visits were last true, so it's the default as-of date for 3NA.
  function findRunDate(values, hdrRow){
    for (let r = 0; r < Math.min(hdrRow, 12); r++) for (const v of values[r] || []){
      const s = txt(v); if (/^from\s/i.test(s)) continue;
      if (/\d{1,2}\/\d{1,2}\/\d{2,4}\s+\d{1,2}:\d{2}/.test(s)){ const d = parseDateTime(s.replace(/\s+([AP])M?$/i, ' $1')); if (d) return d; }
    }
    return null;
  }
  function normStatus(s){ const v = lc(s); if (/^kept|^checked|^complete|^arrived|^seen/.test(v)) return 'Kept'; if (/no.?show|^ns$/.test(v)) return 'No Show'; if (/cancel|^cx|resched/.test(v)) return 'Cancelled'; if (/expect|sched|confirm|booked|pending/.test(v)) return 'Expected'; return txt(s) || 'Expected'; }
  function parseAppts(values){
    const h = findHeader(values); if (!h) return { error: 'Couldn’t find the header row. The appointment list needs Appt Date, Provider, Event and Status columns.' };
    const c = h.cols, rows = []; let skipped = 0;
    for (let r = h.row + 1; r < values.length; r++){
      const v = values[r] || []; const date = parseDate(v[c.date]);
      if (date == null){ if (v.some(x => txt(x))) skipped++; continue; }
      const start = parseTime(v[c.time]); const endT = c.end != null ? parseTime(v[c.end]) : null;
      let dur = c.dur != null && v[c.dur] !== '' && !isNaN(+v[c.dur]) ? +v[c.dur] : null;
      if (dur == null && start != null && endT != null && endT > start) dur = endT - start;
      let booked = c.booked != null ? parseDateTime(v[c.booked]) : null;
      if (booked != null && c.bookedTime != null){ const t = parseTime(v[c.bookedTime]); if (t != null) booked = Math.floor(booked) + t / 1440; }
      let cancelAt = c.cancel != null ? parseDateTime(v[c.cancel]) : null;
      if (cancelAt != null && c.cancelTime != null){ const t = parseTime(v[c.cancelTime]); if (t != null) cancelAt = Math.floor(cancelAt) + t / 1440; }
      rows.push({ date, start, dur, prov: txt(v[c.prov]), event: txt(v[c.event]), status: normStatus(v[c.status]), rawStatus: txt(v[c.status]), loc: c.loc != null ? txt(v[c.loc]) : '', booked, cancelAt, reason: c.reason != null ? txt(v[c.reason]) : '' });
    }
    const dates = rows.map(x => x.date);
    return { rows, header: h, runAt: findRunDate(values, h.row), skipped, min: dates.length ? Math.min(...dates) : null, max: dates.length ? Math.max(...dates) : null,
      has: { booked: c.booked != null, cancel: c.cancel != null, reason: c.reason != null, dur: c.dur != null || c.end != null, loc: c.loc != null } };
  }

  // ---------- analysis ----------
  const ACTIVE = new Set(['Kept', 'No Show', 'Expected']);
  function blankAcc(){ return { supply: 0, kept: 0, keptFit: 0, keptOff: 0, ns: 0, open: 0, admin: 0, pto: 0, futSupply: 0, futBooked: 0, aKept: 0, aNS: 0, aCx: 0, aLC: 0, aCxKnown: 0, aExp: 0, aAll: 0, inBlocked: 0, offTemplate: 0, dbl: 0, overMin: 0, lead: [] }; }
  function rates(a){
    const d = (n, m) => m ? n / m : null;
    return { ...a, fill: d(a.kept + a.ns, a.supply), util: d(a.kept, a.supply), fit: d(a.keptFit, a.kept), nsRate: d(a.aNS, a.aKept + a.aNS), cxRate: d(a.aCx, a.aKept + a.aNS + a.aCx), lcRate: a.aCxKnown ? d(a.aLC, a.aKept + a.aNS + a.aCx) : null,
      missRate: d(a.aNS + (a.aCxKnown ? a.aLC : 0), a.aKept + a.aNS + a.aCx), futFill: d(a.futBooked, a.futSupply), leadMed: median(a.lead), lead: undefined };
  }
  function median(xs){ if (!xs || !xs.length) return null; const s = [...xs].sort((a, b) => a - b); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; }

  // inputs: { templates:[parsed grids], appts:[parsed rows], ref:[rows], slotTypes:[rows], providers:[{name,label,cred,include}], opts }
  // opts: { asOf, from, to, lateHours (24), provs:Set|null, locs:Set|null, sls:Set|null, threeNaDays }
  function analyze(input){
    const o = Object.assign({ lateHours: 24 }, input.opts || {});
    const ref = buildRef(input.ref), stMap = buildSlotTypes(input.slotTypes);
    const allAppts = input.appts || [];
    const asOf = o.asOf != null ? Math.floor(o.asOf) : todaySerial();
    const minD = allAppts.length ? Math.min(...allAppts.map(a => a.date)) : asOf, maxD = allAppts.length ? Math.max(...allAppts.map(a => a.date)) : asOf;
    const from = o.from != null ? o.from : minD, to = o.to != null ? o.to : Math.min(asOf, maxD);

    // providers: who is a person, how templates map to them
    const pmeta = new Map();
    for (const p of input.providers || []) if (p.name) pmeta.set(p.name, p);
    const names = [...new Set(allAppts.map(a => a.prov).filter(Boolean))];
    const people = names.filter(n => { const p = pmeta.get(n); if (p && p.include != null && p.include !== '') return !/^n|^false$/i.test(String(p.include)); return isPersonResource(n); });
    const tByProv = new Map(); const unlinked = [];
    for (const t of input.templates || []){
      const fixed = [...pmeta.values()].find(p => p.label && lc(p.label) === lc(t.label) && p.name);
      const name = fixed ? fixed.name : matchProvider(t.label, people.length ? people : names) || matchProvider(t.title, people.length ? people : names);
      t.provider = name || null;
      if (!name){ unlinked.push(t); continue; }
      if (!tByProv.has(name)) tByProv.set(name, []); tByProv.get(name).push(t);
    }
    // which template version applies on a date: the latest Effective From on or before it; blank = from the start
    for (const list of tByProv.values()) list.sort((a, b) => (a.effFrom ?? -1) - (b.effFrom ?? -1));
    const versionOn = (prov, d) => { const list = tByProv.get(prov); if (!list) return null; let pick = null; for (const t of list){ if ((t.effFrom == null || t.effFrom <= d) && (t.effTo == null || t.effTo >= d)) pick = t; } return pick; };

    const keep = a => (!o.provs || o.provs.has(a.prov)) && (!o.locs || o.locs.has(a.loc));
    const cls = new Map(); const clsOf = a => { let c = cls.get(a.event); if (!c){ c = classify(ref, a.event); cls.set(a.event, c); } return c; };
    const unmapped = new Map();
    const appts = [];
    for (const a of allAppts){
      if (!a.prov || !people.includes(a.prov)) continue;                 // providers only, no chairs or rooms
      const c = clsOf(a); if (!c.mapped) unmapped.set(a.event, (unmapped.get(a.event) || 0) + 1);
      if (c.kind === 'ignore') continue;
      const x = Object.assign({}, a, { kind: c.kind, cls: c.cls, sl: c.sl, pto: !!c.pto });
      if (o.sls && x.kind === 'pt' && !String(x.sl).split(',').map(s => s.trim()).some(s => o.sls.has(s))) continue;
      appts.push(x);
    }

    // lay the templates over the calendar
    const slots = []; const slotIdx = new Map(); const unknownTypes = new Map();
    const horizon = maxD;
    for (const [prov, list] of tByProv){
      if (o.provs && !o.provs.has(prov)) continue;
      for (let d = Math.min(from, asOf); d <= horizon; d++){
        const t = versionOn(prov, d); if (!t) continue;
        const wd = dow(d); const day = t.slots.filter(s => s.dow === wd); if (!day.length) continue;
        if (o.locs && t.location && !o.locs.has(t.location) && ![...o.locs].some(l => lc(l).includes(lc(t.location)) || lc(t.location).includes(lc(l)))) continue;
        const key = prov + '|' + d; const arr = [];
        for (const s of day){
          const st = slotTypeOf(stMap, s.type); if (st.unknown) unknownTypes.set(s.type, (unknownTypes.get(s.type) || 0) + 1);
          const sl = { prov, date: d, min: s.min, len: s.len, type: st.type, st, tpl: t, occ: [], nonpt: null, pto: false };
          arr.push(sl); slots.push(sl);
        }
        arr.sort((a, b) => a.min - b.min); slotIdx.set(key, arr);
      }
    }
    // place each visit into the slots it overlaps
    const exceptions = [];
    for (const a of appts){
      a.slots = [];
      const arr = slotIdx.get(a.prov + '|' + a.date);
      const len = a.dur || (arr && arr[0] ? arr[0].len : 30);
      if (arr && a.start != null) for (const s of arr) if (s.min < a.start + len && a.start < s.min + s.len) a.slots.push(s);
      if (a.kind === 'nonpt'){ for (const s of a.slots){ if (a.status === 'Cancelled') continue; if (a.pto) s.pto = true; else s.nonpt = s.nonpt || a; } continue; }
      for (const s of a.slots){ if (ACTIVE.has(a.status)) s.occ.push(a); (s.hist || (s.hist = [])).push(a); }
    }

    // per-slot outcome
    const inWin = d => d >= from && d <= to;
    for (const s of slots){
      const past = s.date <= asOf;
      if (s.st.kind !== 'clinical'){ s.out = 'blocked'; continue; }
      if (s.pto){ s.out = 'pto'; continue; }
      const act = s.occ;
      if (!act.length){ s.out = s.nonpt ? 'admin' : past ? 'open' : 'avail'; continue; }
      if (past){
        const kept = act.filter(a => a.status === 'Kept');
        if (kept.length) s.out = kept.some(a => accepts(s.st, a)) ? 'keptFit' : 'keptOff';
        else if (act.some(a => a.status === 'No Show')) s.out = 'ns';
        else s.out = 'unresolved';                          // still "Expected" on a past date
      } else s.out = 'booked';
    }

    // what fills each slot type when it isn't the intended visit
    const offMix = {};
    for (const s of slots) if (s.out === 'keptOff' && s.date >= from && s.date <= to){ const m = offMix[s.type] || (offMix[s.type] = {}); for (const a of s.occ) if (a.status === 'Kept') m[a.event] = (m[a.event] || 0) + 1; }

    // aggregation
    const groups = {};
    const bump = (dim, key, fn) => { const g = groups[dim] || (groups[dim] = new Map()); let a = g.get(key); if (!a){ a = blankAcc(); g.set(key, a); } fn(a); };
    const slotDims = s => ({ all: 'All', prov: s.prov, type: s.type, dow: DAYS[dow(s.date)], hour: Math.floor(s.min / 60), part: s.min < 720 ? 'AM' : 'PM', week: weekStart(s.date), heat: dow(s.date) + '|' + Math.floor(s.min / 60), loc: s.tpl.location || '' });
    for (const s of slots){
      const past = s.date <= asOf;
      if (s.st.kind !== 'clinical') continue;
      const dims = slotDims(s);
      for (const [dim, key] of Object.entries(dims)){
        if (past && inWin(s.date)) bump(dim, key, a => {
          if (s.out === 'pto'){ a.pto++; return; }
          a.supply++;
          if (s.out === 'keptFit'){ a.kept++; a.keptFit++; } else if (s.out === 'keptOff'){ a.kept++; a.keptOff++; } else if (s.out === 'ns') a.ns++; else if (s.out === 'admin') a.admin++; else if (s.out === 'open') a.open++;
        });
        if (!past && s.out !== 'pto') bump(dim, key, a => { a.futSupply++; if (s.out === 'booked') a.futBooked++; });
      }
    }
    const lateCut = (o.lateHours || 24) / 24;
    for (const a of appts){
      if (a.kind !== 'pt') continue;
      const firstSlot = a.slots[0];
      if (inWin(a.date) && a.date <= asOf){
        const dims = { all: 'All', prov: a.prov, event: a.event, cls: a.cls, sl: a.sl, dow: DAYS[dow(a.date)], hour: a.start != null ? Math.floor(a.start / 60) : 'No time', part: a.start == null ? 'No time' : a.start < 720 ? 'AM' : 'PM', week: weekStart(a.date), heat: a.start != null ? dow(a.date) + '|' + Math.floor(a.start / 60) : null, loc: a.loc, type: firstSlot ? firstSlot.type : 'Outside template' };
        let lc_ = false;
        if (a.status === 'Cancelled' && a.cancelAt != null){ const startAt = a.date + (a.start || 0) / 1440; lc_ = startAt - a.cancelAt < lateCut; }
        const lead = a.booked != null && a.status !== 'Cancelled' ? a.date - Math.floor(a.booked) : null;
        for (const [dim, key] of Object.entries(dims)){
          if (key == null) continue;
          bump(dim, key, x => {
            x.aAll++;
            if (a.status === 'Kept') x.aKept++; else if (a.status === 'No Show') x.aNS++; else if (a.status === 'Cancelled'){ x.aCx++; if (a.cancelAt != null){ x.aCxKnown++; if (lc_) x.aLC++; } } else x.aExp++;
            if (lead != null && lead >= 0) x.lead.push(lead);
          });
        }
      }
    }
    // provider impact: visits in blocked time, outside the template, double-booked
    const tplProvs = new Set(tByProv.keys());
    for (const a of appts){
      if (a.kind !== 'pt' || !ACTIVE.has(a.status) || !inWin(a.date) || a.date > asOf || !tplProvs.has(a.prov)) continue;
      const len = a.dur || 30; let why = null;
      if (!slotIdx.has(a.prov + '|' + a.date)) why = 'No template that day';
      else if (!a.slots.length) why = 'Outside template hours';
      else if (a.slots.every(s => s.st.kind !== 'clinical')) why = 'In ' + a.slots.map(s => s.type).filter((v, i, z) => z.indexOf(v) === i).join(' / ');
      else if (a.slots.some(s => s.st.kind !== 'clinical')) why = 'Runs into ' + a.slots.filter(s => s.st.kind !== 'clinical').map(s => s.type).filter((v, i, z) => z.indexOf(v) === i).join(' / ');
      const dbl = a.slots.some(s => s.st.kind === 'clinical' && s.occ.length > 1 && s.occ[0] !== a);
      const k = { all: 'All', prov: a.prov, dow: DAYS[dow(a.date)], week: weekStart(a.date) };
      if (why || dbl){
        const over = why ? (why.startsWith('Runs into') ? a.slots.filter(s => s.st.kind !== 'clinical').reduce((n, s) => n + Math.min(s.min + s.len, a.start + len) - Math.max(s.min, a.start), 0) : len) : 0;
        for (const [dim, key] of Object.entries(k)) bump(dim, key, x => { if (why && (why.startsWith('In ') || why.startsWith('Runs into'))) x.inBlocked++; else if (why) x.offTemplate++; if (dbl) x.dbl++; x.overMin += over; });
        exceptions.push({ date: a.date, time: a.start, prov: a.prov, event: a.event, status: a.status, why: [why, dbl ? 'Double-booked' : null].filter(Boolean).join('; '), min: over });
      }
    }

    // 3NA: calendar days from the as-of date to the third open slot of each type
    const nth = (list, n) => list.length >= n ? list[n - 1].date - asOf : null;
    const avail = slots.filter(s => s.date > asOf && s.out === 'avail').sort((a, b) => a.date - b.date || a.min - b.min);
    const threeNa = [];
    const byType = new Map(); for (const s of avail){ const k = s.type; if (!byType.has(k)) byType.set(k, []); byType.get(k).push(s); }
    const typeSupply = new Map(); for (const s of slots) if (s.date > asOf && s.st.kind === 'clinical' && s.out !== 'pto') typeSupply.set(s.prov + '|' + s.type, (typeSupply.get(s.prov + '|' + s.type) || 0) + 1);
    for (const [type, list] of byType){
      threeNa.push({ prov: 'All providers', type, next: nth(list, 1), third: nth(list, 3), open: list.length, nextDate: list[0] ? list[0].date : null, thirdDate: list[2] ? list[2].date : null });
      const ps = new Map(); for (const s of list){ if (!ps.has(s.prov)) ps.set(s.prov, []); ps.get(s.prov).push(s); }
      for (const [prov, l] of ps) threeNa.push({ prov, type, next: nth(l, 1), third: nth(l, 3), open: l.length, nextDate: l[0].date, thirdDate: l[2] ? l[2].date : null });
    }
    // types that exist in the template but have no open slot left in the horizon
    for (const [k] of typeSupply){ const [prov, type] = k.split('|'); if (!threeNa.some(x => x.prov === prov && x.type === type)) threeNa.push({ prov, type, next: null, third: null, open: 0, full: true }); }
    // and by visit class: the next slots whose type accepts it
    const clsSeen = [...new Set(appts.filter(a => a.kind === 'pt').map(a => a.cls))];
    const threeNaClass = clsSeen.map(c => { const sample = appts.find(a => a.cls === c); const l = avail.filter(s => s.st.terms.length && !s.st.any ? accepts(s.st, sample) : false); const anyL = avail.filter(s => accepts(s.st, sample)); return { cls: c, sl: sample.sl, next: nth(anyL, 1), third: nth(anyL, 3), thirdDedicated: nth(l, 3), open: anyL.length }; }).filter(x => x.open || appts.some(a => a.cls === x.cls && a.date > asOf));

    // historical 3NA, rebuilt from booked/cancel dates when the report has them
    let threeNaHist = null;
    if (allAppts.some(a => a.booked != null)){
      threeNaHist = [];
      const clinical = slots.filter(s => s.st.kind === 'clinical' && s.out !== 'pto' && s.out !== 'admin').sort((a, b) => a.date - b.date || a.min - b.min);
      const firstSun = weekStart(from) + 7;
      for (let D = firstSun; D <= Math.min(asOf, to); D += 7){
        // open on day D = not yet booked by then, or booked and later cancelled after D
        const heldOn = a => a.booked != null && Math.floor(a.booked) <= D && (a.status !== 'Cancelled' || (a.cancelAt != null && a.cancelAt > D));
        const openAt = clinical.filter(s => s.date > D && !(s.hist || []).some(heldOn));   // slots are already in date order
        const per = new Map(); for (const s of openAt){ if (!per.has(s.type)) per.set(s.type, []); per.get(s.type).push(s); }
        for (const [type, l] of per) threeNaHist.push({ week: D, type, third: l.length >= 3 ? l[2].date - D : null });
        threeNaHist.push({ week: D, type: 'Any clinical slot', third: openAt.length >= 3 ? openAt[2].date - D : null });
      }
    }

    const fin = dim => { const g = groups[dim]; if (!g) return []; return [...g.entries()].map(([k, a]) => Object.assign({ key: k }, rates(a))); };
    const total = fin('all')[0] || rates(blankAcc());
    return {
      asOf, from, to, minD, maxD, people, templates: input.templates || [], tByProv, unlinked,
      unmapped: [...unmapped.entries()].map(([event, n]) => ({ event, n, guess: classify(new Map(), event) })).sort((a, b) => b.n - a.n),
      unknownTypes: [...unknownTypes.entries()].map(([type, n]) => ({ type, n })),
      total, byProv: fin('prov'), byType: fin('type'), byEvent: fin('event'), byClass: fin('cls'), bySL: fin('sl'), byDow: fin('dow'), byHour: fin('hour'), byPart: fin('part'), byWeek: fin('week').sort((a, b) => a.key - b.key), heat: fin('heat'), byLoc: fin('loc'),
      threeNa: threeNa.sort((a, b) => (a.prov === 'All providers' ? 0 : 1) - (b.prov === 'All providers' ? 0 : 1) || a.prov.localeCompare(b.prov) || a.type.localeCompare(b.type)), threeNaClass, threeNaHist,
      offMix,
      exceptions: exceptions.sort((a, b) => a.date - b.date || (a.time || 0) - (b.time || 0)),
      hasLate: allAppts.some(a => a.cancelAt != null), hasBooked: allAppts.some(a => a.booked != null),
      slotCount: slots.length, apptCount: appts.length,
    };
  }

  // ---------- ESK / IVK access model import (mhs-access-model v1) ----------
  function readAccessModel(sheets){
    const about = sheets.About || [];
    const kv = Object.fromEntries(about.slice(1).map(r => [txt(r[0]), r[1]]));
    if (txt(kv.Schema) !== 'mhs-access-model') return { error: 'This file isn’t an access-model export (About tab is missing or has another schema).' };
    const tbl = name => { const rows = sheets[name] || []; if (!rows.length) return []; const h = rows[0].map(txt); return rows.slice(1).filter(r => r.some(v => txt(v))).map(r => Object.fromEntries(h.map((k, i) => [k, r[i]]))); };
    // patient-level tabs (Patients, Moves) are never read
    return { version: +kv['Schema version'] || 1, modality: txt(kv.Modality), clinic: txt(kv.Clinic), generated: txt(kv.Generated), providers: tbl('Providers'), dayPart: tbl('DayPart'), frequencies: tbl('Frequencies'), weekly: tbl('Weekly') };
  }

  // ---------- defaults ----------
  const DEFAULT_SLOT_TYPES = [
    ['Open Schedule', 'Clinical', '*', 'Any visit'],
    ['New Patient', 'Clinical', 'Intake', ''],
    ['Telepsych', 'Clinical', 'Medication Management, TMS Follow Up, ADHD', 'Tele follow-ups'],
    ['Interventional Consult', 'Clinical', 'Consult, TMS Consult', ''],
    ['Esketamine Follow Up', 'Clinical', 'EFU, ESK, ESK Start', ''],
    ['Provider MT', 'Clinical', 'TMS Start, MT Redo, TMS', 'Motor threshold'],
    ['TMS Follow Up', 'Clinical', 'TMS Follow Up', ''],
    ['IV Ketamine Start/Index', 'Clinical', 'IVK Start', ''],
    ['IV Ketamine Follow Up', 'Clinical', 'IVK FU', ''],
    ['IV Ketamine Consult', 'Clinical', 'Consult+IVK', ''],
    ['IVK Post Check', 'Clinical', 'IVK FU, IVK', ''],
    ['Flex Post Check', 'Clinical', 'IVK FU, TMS Follow Up, Medication Management', 'TMS F/U or tele when no index patient'],
    ['Med Management', 'Clinical', 'Medication Management', ''],
    ['Admin', 'Blocked', '', ''], ['Lunch', 'Blocked', '', ''], ['Meeting', 'Blocked', '', ''],
  ];

  return { DAYS, dow, weekStart, iso, mdy, hhmm, hourLabel, parseTime, parseDate, parseDateTime, serialFromYMD, todaySerial,
    nameTokens, isPersonResource, credentialOf, matchProvider, lastName, parseTemplates, buildRef, classify, buildSlotTypes, slotTypeOf, accepts,
    findHeader, parseAppts, analyze, readAccessModel, guessServiceLine, DEFAULT_SLOT_TYPES };
});
