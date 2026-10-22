// Clinic Template Tool cadence, recall and funnel: how often each provider's patients come back, who has stopped coming,
// which open slots could take them, and how many consults turn into starts. Pure functions on the parsed appointment
// list and the engine's analysis (no Excel calls), so tests can run it under Node.
// Patients are keyed by Md Rc when the report has it, else by name. Everything shown or written uses LASFIR only.
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./engine.js'));
  else root.TC = factory(root.TE);
})(typeof self !== 'undefined' ? self : this, function (TE) {
  'use strict';

  const lc = s => String(s == null ? '' : s).trim().toLowerCase();
  const txt = s => String(s == null ? '' : s).trim();
  const DEFAULTS = { releaseHours: 24, convDays: 90, lapsedDays: 60, overdueMult: 1.5, tmsCount: 30, maintDays: 180, matchDays: 10, dayStart: 480, dayEnd: 1080 };
  const HIST = [['≤1 wk', 1], ['2 wk', 2], ['3 wk', 3], ['4 wk', 4], ['6 wk', 6], ['8 wk', 8], ['12 wk', 12], ['>12 wk', Infinity]];
  const TELE = /tele|virtual|video|\btel\b/i;
  const CALL = new Set(['overdue', 'lapsed', 'unaccounted', 'maintenance']);
  const median = xs => { if (!xs.length) return null; const s = [...xs].sort((a, b) => a - b); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
  const mean = xs => xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : null;
  const bucketOf = sl => txt(String(sl || '').split(',')[0]) || 'Unassigned';
  const slTokens = sl => String(sl || '').split(',').map(txt).filter(Boolean);
  const isKept = v => v.status === 'Kept';
  function addBusinessDays(d, n){ let x = d; while (n > 0){ x++; const w = TE.dow(x); if (w !== 0 && w !== 6) n--; } return x; }

  // ---------- visits ----------
  // One record per patient visit (chairs and rooms included: TMS and ESK treatments are booked on them).
  // o: { asOf, releaseHours, dayStart, dayEnd }, people: Set of NextGen names that are providers (not chairs or rooms).
  function visits(appts, ref, o, people){
    o = Object.assign({}, DEFAULTS, o);
    const refMap = TE.buildRef(ref); const memo = new Map(); const out = []; let noKey = 0;
    for (const a of appts || []){
      if (!a.prov) continue;
      let c = memo.get(a.event); if (!c){ c = TE.classify(refMap, a.event); memo.set(a.event, c); }
      if (c.kind !== 'pt') continue;
      const key = a.mrn ? 'm:' + lc(a.mrn) : a.pt ? 'n:' + lc(a.pt) : '';
      if (!key){ noKey++; continue; }
      const status = a.status === 'Expected' && !txt(a.rawStatus) && a.date < o.asOf ? 'Kept' : a.status;   // blank status on a past date counts as kept
      const at = a.date + (a.start || 0) / 1440;
      const leadH = a.booked != null && a.start != null ? TE.businessHours(a.booked, at, o.dayStart, o.dayEnd) : null;
      out.push({ key, code: TE.lasfir(a.pt), prov: a.prov, person: people ? people.has(a.prov) : TE.isPersonResource(a.prov),
        date: a.date, at, dur: a.dur, event: a.event, cls: c.cls, sl: c.sl, bucket: bucketOf(c.sl), role: c.role, status, loc: a.loc,
        tele: TELE.test(a.event), leadH, release: leadH != null && leadH < o.releaseHours });
    }
    out.sort((a, b) => a.at - b.at);
    return { list: out, noKey, named: out.filter(v => v.code).length };
  }

  // A patient's visits in one service line. The owner is the provider (a person, not a chair) they last saw in it.
  function chains(vs, asOf){
    const m = new Map();
    for (const v of vs){ const k = v.key + '|' + v.bucket; let c = m.get(k); if (!c){ c = { key: v.key, code: '', bucket: v.bucket, visits: [] }; m.set(k, c); } c.visits.push(v); if (v.code) c.code = v.code; }
    for (const c of m.values()){
      const past = c.visits.filter(v => v.date <= asOf);
      const pick = [...past].reverse().find(v => v.person && isKept(v)) || [...c.visits].reverse().find(v => v.person) || c.visits[c.visits.length - 1];
      c.owner = pick.prov;
      const lastV = [...past].reverse().find(isKept) || c.visits[c.visits.length - 1]; c.loc = lastV.loc;
    }
    return m;
  }
  // kept visit days in a chain: same-day duplicates collapse so they never make zero-day gaps
  function keptDays(list){ const days = []; for (const v of list) if (isKept(v) && v.role !== 'consult' && v.role !== 'admin' && days[days.length - 1] !== v.date) days.push(v.date); return days; }

  function gapStats(gaps){
    const hist = HIST.map(() => 0);
    if (!gaps.length) return { n: 0, mean: null, median: null, modeWk: null, modeShare: null, expDays: null, hist };
    const counts = new Map();
    for (const g of gaps){ const wk = Math.round(g / 7); hist[HIST.findIndex(b => wk <= b[1])]++; const k = Math.max(1, wk); counts.set(k, (counts.get(k) || 0) + 1); }
    const [modeWk, n] = [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0] - b[0])[0];
    return { n: gaps.length, mean: mean(gaps), median: median(gaps), modeWk, modeShare: n / gaps.length, expDays: modeWk * 7, hist };
  }

  // ---------- cadence ----------
  // Per provider (chain owner) × service line over from..to: patients, kept visits, return intervals, no-shows.
  // Consults anchor a chain but aren't part of it; consult-to-start lives in the funnel.
  // o: { from, to, asOf, provs, sls, locs }
  function cadence(vs, o){
    const ch = chains(vs, o.asOf);
    const acc = new Map(); const patients = [];
    const get = (prov, bucket) => { const k = prov + '|' + bucket; let a = acc.get(k); if (!a){ a = { prov, bucket, patients: 0, kept: 0, returning: 0, gaps: [], ns: 0, cx: 0, tele: 0 }; acc.set(k, a); } return a; };
    for (const c of ch.values()){
      if (o.sls && !o.sls.has(c.bucket)) continue;
      if (o.provs && !o.provs.has(c.owner)) continue;
      const win = c.visits.filter(v => v.date >= o.from && v.date <= o.to && v.date <= o.asOf && (!o.locs || o.locs.has(v.loc)) && v.role !== 'consult' && v.role !== 'admin');
      const days = keptDays(win);
      const ns = win.filter(v => v.status === 'No Show').length, cx = win.filter(v => v.status === 'Cancelled').length;
      if (!days.length && !ns && !cx) continue;
      const gaps = days.slice(1).map((d, i) => d - days[i]);
      const keptV = win.filter(isKept), tele = keptV.filter(v => v.tele).length;
      for (const a of [get(c.owner, c.bucket), get('All providers', c.bucket)]){
        a.patients++; a.kept += days.length; if (days.length >= 2) a.returning++; a.gaps.push(...gaps); a.ns += ns; a.cx += cx; a.tele += tele;
      }
      if (days.length) patients.push({ code: c.code, key: c.key, prov: c.owner, bucket: c.bucket, kept: days.length, first: days[0], last: days[days.length - 1],
        meanGap: mean(gaps), lastGap: gaps.length ? gaps[gaps.length - 1] : null, telePct: keptV.length ? tele / keptV.length : null, ns, cx });
    }
    const rows = [...acc.values()].map(a => Object.assign({ prov: a.prov, bucket: a.bucket, patients: a.patients, kept: a.kept, perPt: a.patients ? a.kept / a.patients : null, returning: a.returning,
      ns: a.ns, cx: a.cx, nsRate: a.kept + a.ns ? a.ns / (a.kept + a.ns) : null, telePct: a.kept ? a.tele / a.kept : null }, gapStats(a.gaps)));
    rows.sort((a, b) => (a.prov === 'All providers' ? 0 : 1) - (b.prov === 'All providers' ? 0 : 1) || a.bucket.localeCompare(b.bucket) || b.patients - a.patients);
    const byBucket = new Map(rows.filter(r => r.prov === 'All providers').map(r => [r.bucket, r]));
    return { rows, byBucket, patients: patients.sort((a, b) => a.prov.localeCompare(b.prov) || a.bucket.localeCompare(b.bucket) || b.kept - a.kept), chains: ch, hist: HIST.map(h => h[0]) };
  }

  // ---------- recall ----------
  // Per patient per service line, exactly one state:
  //   active (something booked ahead), paused / discontinued / completed (ESK Tracker), completed (TMS course inferred),
  //   maintenance (completed TMS long enough ago to offer maintenance), lapsed (> lapsedDays), overdue (> overdueMult × the
  //   service line's usual return interval), unaccounted (ESK, overdue or lapsed, and not in the tracker), current (not due).
  // esk: { rows:[{ code, clinic, status, lastDose, reason, ret }] } or null.
  function recall(vs, o, cad, esk){
    o = Object.assign({}, DEFAULTS, o);
    const ch = cad.chains; const out = [];
    const eskBy = new Map(); for (const r of (esk && esk.rows) || []){ if (!eskBy.has(r.code)) eskBy.set(r.code, []); eskBy.get(r.code).push(r); }
    const trackerFor = c => { const l = eskBy.get(c.code); if (!l) return null; return l.find(r => r.clinic && c.loc && (lc(c.loc).includes(lc(r.clinic)) || lc(r.clinic).includes(lc(c.loc)))) || (l.length === 1 ? l[0] : null); };
    for (const c of ch.values()){
      if (o.sls && !o.sls.has(c.bucket)) continue;
      if (o.provs && !o.provs.has(c.owner)) continue;
      if (o.locs && !o.locs.has(c.loc)) continue;
      const kept = c.visits.filter(v => isKept(v) && v.date <= o.asOf);
      if (!kept.some(v => v.role !== 'consult')) continue;            // consult only: that's the funnel, not recall
      const last = kept[kept.length - 1].date, daysSince = o.asOf - last;
      const future = c.visits.find(v => v.status === 'Expected' && v.date >= o.asOf);
      const after = c.visits.filter(v => v.date > last && v.date <= o.asOf && (v.status === 'No Show' || v.status === 'Cancelled'));
      const post = after.length ? after[after.length - 1].status : '';
      const b = cad.byBucket.get(c.bucket); const exp = b ? b.expDays : null;
      const isEsk = /\besk|esketamine|spravato/i.test(c.bucket);
      const tr = isEsk ? trackerFor(c) : null;
      let state;
      if (future) state = 'active';
      else if (tr && tr.status === 'discontinued') state = 'discontinued';
      else if (tr && tr.status === 'paused') state = 'paused';
      else if (tr && tr.status === 'completed') state = 'completed';
      else if (/\btms\b/i.test(c.bucket) && tmsDone(kept, o.tmsCount)) state = daysSince > o.maintDays ? 'maintenance' : 'completed';
      else if (daysSince > o.lapsedDays) state = 'lapsed';
      else if (exp && daysSince > o.overdueMult * exp) state = 'overdue';
      else state = 'current';
      if ((state === 'overdue' || state === 'lapsed') && isEsk && esk && !tr) state = 'unaccounted';
      const roles = {}; kept.forEach(v => { if (v.role !== 'consult') roles[v.role] = (roles[v.role] || 0) + 1; });
      const main = Object.entries(roles).sort((x, y) => y[1] - x[1])[0][0];
      const sample = [...kept].reverse().find(v => v.role === main);
      out.push({ code: c.code, key: c.key, prov: c.owner, bucket: c.bucket, loc: c.loc, last, daysSince, exp, overdueBy: daysSince - (exp ? Math.round(o.overdueMult * exp) : o.lapsedDays),
        state, post, next: future ? future.date : null, tracker: tr ? tr.status : (isEsk && esk ? 'Not in tracker' : ''), ret: tr ? tr.ret : null, reason: tr ? tr.reason : '', sample: { cls: sample.cls, event: sample.event, sl: sample.sl }, role: main });
    }
    return out.sort((a, b) => (CALL.has(b.state) - CALL.has(a.state)) || b.overdueBy - a.overdueBy);
  }
  // TMS course done: at least `count` treatments, then a follow-up after the last one
  function tmsDone(kept, count){
    const tx = kept.filter(v => v.role === 'treatment'); if (tx.length < count) return false;
    const lastTx = tx[tx.length - 1].date; return kept.some(v => v.role === 'followup' && v.date > lastTx);
  }

  // ---------- template slots: flex, release window, matches ----------
  // Representative visits per role (and per service line) so a slot type's Accepts list can be tested by role.
  function samples(vs){
    const m = new Map();
    for (const v of vs){ const k = v.role + '|' + v.bucket; if (!m.has(k)) m.set(k, new Map()); const s = m.get(k); const id = lc(v.cls) + '|' + lc(v.event); if (!s.has(id)) s.set(id, { cls: v.cls, event: v.event, sl: v.sl, role: v.role, bucket: v.bucket }); }
    const all = []; for (const s of m.values()) all.push(...s.values());
    return all;
  }
  // which roles a slot type takes; flex = two or more patient roles
  function slotRoles(st, smp){ const r = new Set(); for (const s of smp) if (s.role !== 'admin' && TE.accepts(st, s)) r.add(s.role); return [...r]; }
  function flexTypes(R, smp){
    const m = new Map();
    for (const s of R.slots){ if (m.has(s.type) || s.st.kind !== 'clinical') continue; const roles = slotRoles(s.st, smp); m.set(s.type, { roles, flex: s.st.any || roles.length >= 2 }); }
    return m;
  }
  // Flex slots: offered, used, which roles actually filled them, and how many of those were release bookings.
  function flexUsage(R, vs, o){
    const smp = samples(vs), types = flexTypes(R, smp); const acc = new Map();
    for (const s of R.slots){
      const t = types.get(s.type); if (!t || !t.flex || s.date > R.asOf || s.date < R.from || s.date > R.to || s.out === 'pto' || s.out === 'unresolved') continue;
      const k = s.prov + '|' + s.type; let a = acc.get(k); if (!a){ a = { prov: s.prov, type: s.type, accepts: t.roles, offered: 0, used: 0, release: 0, roles: {} }; acc.set(k, a); }
      a.offered++;
      const kept = s.occ.find(x => x.status === 'Kept');
      if (kept){ a.used++; a.roles[kept.role] = (a.roles[kept.role] || 0) + 1; if (kept.release) a.release++; }
    }
    return [...acc.values()].map(a => Object.assign(a, { useRate: a.offered ? a.used / a.offered : null, releaseShare: a.used ? a.release / a.used : null })).sort((a, b) => a.prov.localeCompare(b.prov) || b.offered - a.offered);
  }
  // Open template slots that are, or within a business day will be, inside the release window (any visit may take them).
  // now = report run time (serial date-time).
  function releaseList(R, o, now){
    o = Object.assign({}, DEFAULTS, o);
    const until = addBusinessDays(Math.floor(now), 1) + o.dayEnd / 1440;
    const out = [];
    for (const s of R.slots){
      if (s.out !== 'avail') continue;
      const at = s.date + s.min / 1440; if (at <= now) continue;
      const enters = TE.minusBusinessHours(at, o.releaseHours, o.dayStart, o.dayEnd);
      if (enters <= until) out.push({ prov: s.prov, date: s.date, min: s.min, type: s.type, released: enters <= now, enters });
    }
    return out.sort((a, b) => a.date - b.date || a.min - b.min);
  }
  // For each patient on the call list: open slots in the next matchDays business days that take their usual visit
  // (dedicated or flex), or that sit with their own provider inside the release window and serve their service line.
  function slotMatches(rc, R, vs, o, now, perPatient = 3){
    o = Object.assign({}, DEFAULTS, o);
    const smp = samples(vs), types = flexTypes(R, smp);
    const inLine = new Map();   // slot type|service line -> takes some visit in that line
    const servesLine = (s, bucket) => { const k = s.type + '|' + bucket; if (!inLine.has(k)) inLine.set(k, smp.some(x => x.bucket === bucket && x.role !== 'admin' && TE.accepts(s.st, x))); return inLine.get(k); };
    const horizon = addBusinessDays(Math.floor(now), o.matchDays);
    const open = R.slots.filter(s => s.out === 'avail' && s.date + s.min / 1440 > now && s.date <= horizon).sort((a, b) => a.date - b.date || a.min - b.min);
    const enters = new Map(open.map(s => [s, TE.minusBusinessHours(s.date + s.min / 1440, o.releaseHours, o.dayStart, o.dayEnd)]));
    const out = [];
    for (const p of rc){
      if (!CALL.has(p.state)) continue;
      const hits = [];
      for (const s of open){
        const ok = TE.accepts(s.st, p.sample), released = enters.get(s) <= now;
        const why = ok ? ((types.get(s.type) || {}).flex ? 'flex' : 'dedicated') : released && s.prov === p.prov && servesLine(s, p.bucket) ? 'released' : null;
        if (why) hits.push({ s, why });
      }
      const rank = { dedicated: 0, flex: 1, released: 2 };   // own provider first, then soonest day, then the best fit that day
      hits.sort((a, b) => (b.s.prov === p.prov) - (a.s.prov === p.prov) || a.s.date - b.s.date || rank[a.why] - rank[b.why] || a.s.min - b.s.min);
      for (const h of hits.slice(0, perPatient)) out.push({ code: p.code, bucket: p.bucket, state: p.state, patientProv: p.prov, prov: h.s.prov, date: h.s.date, min: h.s.min, type: h.s.type, why: h.why });
    }
    return out;
  }

  // ---------- funnel: consult to start ----------
  // Per consulting provider × the consult's service line: consults offered and kept, how many led to a start in the same
  // service line within convDays (any provider), and start hours offered and used versus start hours needed.
  // Consults too recent to have had convDays are "pending" and left out of the conversion rate.
  function funnel(vs, R, o){
    o = Object.assign({}, DEFAULTS, o);
    const inWin = v => v.date >= o.from && v.date <= o.to && v.date <= o.asOf && (!o.locs || o.locs.has(v.loc));
    const smp = samples(vs);
    const starts = new Map(); for (const v of vs) if (v.role === 'start' && isKept(v)){ if (!starts.has(v.key)) starts.set(v.key, []); starts.get(v.key).push(v); }
    const acc = new Map(); const seen = new Set();
    for (const v of vs){
      if (v.role !== 'consult' || !isKept(v) || !v.person || !inWin(v)) continue;
      if (o.provs && !o.provs.has(v.prov)) continue;
      const sl = txt(v.sl) || 'Unassigned', toks = slTokens(sl);
      if (o.sls && !toks.some(t => o.sls.has(t))) continue;
      const once = v.key + '|' + sl; if (seen.has(once)) continue; seen.add(once);
      const k = v.prov + '|' + sl; let a = acc.get(k); if (!a){ a = { prov: v.prov, sl, toks, consults: 0, converted: 0, pending: 0, days: [], sample: v }; acc.set(k, a); }
      a.consults++;
      const st = (starts.get(v.key) || []).find(s => s.date >= v.date && s.date - v.date <= o.convDays && (sl === 'Unassigned' || toks.includes(s.bucket)));
      if (st){ a.converted++; a.days.push(st.date - v.date); } else if (v.date + o.convDays > o.asOf) a.pending++;
    }
    const past = (R && R.slots || []).filter(s => s.st.kind === 'clinical' && s.date >= R.from && s.date <= Math.min(R.to, R.asOf) && s.out !== 'pto' && s.out !== 'unresolved');
    const rows = [];
    for (const a of acc.values()){
      const startSmp = smp.filter(s => s.role === 'start' && (a.sl === 'Unassigned' || a.toks.includes(s.bucket)));
      const mine = past.filter(s => s.prov === a.prov);
      const consultSlots = mine.filter(s => TE.accepts(s.st, a.sample));
      const startSlots = mine.filter(s => startSmp.some(x => TE.accepts(s.st, x)));
      const startVisits = vs.filter(v => v.role === 'start' && isKept(v) && v.prov === a.prov && inWin(v) && (a.sl === 'Unassigned' || a.toks.includes(v.bucket)));
      const durs = vs.filter(v => v.role === 'start' && isKept(v) && v.dur && (a.sl === 'Unassigned' || a.toks.includes(v.bucket))).map(v => v.dur);
      const startLen = median(durs) || median(startSlots.map(s => s.len)) || 60;
      const decided = a.consults - a.pending, rate = decided ? a.converted / decided : null;
      const offeredH = startSlots.reduce((n, s) => n + s.len, 0) / 60, usedH = startVisits.reduce((n, v) => n + (v.dur || startLen), 0) / 60;
      const neededH = rate == null ? null : a.consults * rate * startLen / 60;
      const flag = neededH == null || !startSlots.length ? '' : offeredH > neededH * 1.2 ? 'over' : offeredH < neededH * 0.8 ? 'under' : 'ok';
      rows.push({ prov: a.prov, sl: a.sl, consultSlots: consultSlots.length, consultHours: consultSlots.reduce((n, s) => n + s.len, 0) / 60, consults: a.consults,
        consultFill: consultSlots.length ? Math.min(1, a.consults / consultSlots.length) : null, converted: a.converted, pending: a.pending, rate, medDays: median(a.days),
        startLen, startHoursOffered: offeredH, startHoursUsed: usedH, startHoursNeeded: neededH, flag });
    }
    return rows.sort((a, b) => a.prov.localeCompare(b.prov) || b.consults - a.consults);
  }

  // ---------- esketamine payer credential rules ----------
  // Every ESK visit booked on a provider (not a chair), checked against the Payer Credential Rules sheet: may this
  // provider's credential bill this visit for the patient's payer? Upcoming ones can still be moved; past ones are billing
  // risk. The NextGen list can't tell a first treatment from later dosing, so only an event named "first" counts as one.
  // PR = the PayerRules module, rules = PR.read(sheet).rules, cred(prov) -> 'MD' | 'NP' | ''. o: { asOf, from, to, provs, locs }
  function eskCredentials(appts, ref, rules, PR, cred, people, o){
    const refMap = TE.buildRef(ref); const memo = new Map(); const rows = [];
    for (const a of appts || []){
      if (!a.prov || !people.has(a.prov) || a.status === 'Cancelled') continue;
      if (o.provs && !o.provs.has(a.prov)) continue;
      if (o.locs && !o.locs.has(a.loc)) continue;
      const future = a.date >= o.asOf;
      if (!future && (a.date < o.from || a.date > o.to)) continue;
      let c = memo.get(a.event); if (!c){ c = TE.classify(refMap, a.event); memo.set(a.event, c); }
      if (c.kind !== 'pt' || !slTokens(c.sl).some(t => /^(esk|esketamine|spravato)$/i.test(t))) continue;
      const step = PR.step(a.event, /\bfirst\b/i.test(a.event));
      const cr = cred(a.prov), m = PR.match(rules, a.payer, a.loc), r = PR.check(m, cr, step);
      rows.push({ date: a.date, start: a.start, prov: a.prov, cred: cr, event: a.event, status: a.status, loc: a.loc, payer: a.payer, rule: m ? m.payer : '',
        ambiguous: !!(m && m.ambiguous.length), step, ok: r.ok, why: r.why, besse: PR.besse(m), code: TE.lasfir(a.pt), future, m });
    }
    rows.sort((x, y) => x.date - y.date || (x.start || 0) - (y.start || 0));
    const byPayer = new Map();
    for (const r of rows){
      const k = (r.payer || '(no payer)') + '|' + r.rule; let g = byPayer.get(k);   // one payer name can fall under a CA and a WA rule
      if (!g){ g = { payer: r.payer || '(no payer)', rule: r.rule, ambiguous: r.ambiguous, besse: r.besse, visits: 0, md: 0, np: 0, blocked: 0, upcomingBlocked: 0, m: r.m }; byPayer.set(k, g); }
      g.visits++; if (r.cred === 'MD') g.md++; if (r.cred === 'NP') g.np++;
      if (r.ok === false){ g.blocked++; if (r.future) g.upcomingBlocked++; }
    }
    // what an NP may do for each payer: every ESK step, observation only, or none
    for (const g of byPayer.values()){
      const m = g.m;   // the first visit's match, which used its clinic for the state
      if (!m || m.ambiguous.length){ g.npSteps = m ? 'Depends on state' : ''; continue; }
      const ok = ['first', 'start', 'obs'].filter(s => PR.check(m, 'NP', s).ok === true);
      g.npSteps = !m.NP ? 'No NP row' : ok.length === 3 ? 'All visits' : ok.length === 0 ? 'None' : ok.map(s => PR.STEPS[s]).join(', ') + ' only';
    }
    return { rows, byPayer: [...byPayer.values()].sort((a, b) => b.visits - a.visits), upcoming: rows.filter(r => r.future && r.ok === false), past: rows.filter(r => !r.future && r.ok === false),
      unclear: rows.filter(r => r.future && r.ok === null && r.ambiguous) };
  }

  // ---------- ESK Tracker status import ----------
  // CSV or sheet with LASFIR, clinic, status, last dose date, pause reason, expected return date. Names are refused.
  const ESK_COLS = { code: /^(lasfir|patient\s*code|code|patient)$/i, clinic: /^(clinic|location|site)$/i, status: /^status$/i, lastDose: /^last\s*dose/i, reason: /reason/i, ret: /return/i };
  function readEskStatus(aoa){
    for (let r = 0; r < Math.min((aoa || []).length, 10); r++){
      const h = (aoa[r] || []).map(txt); const idx = {};
      for (const [k, re] of Object.entries(ESK_COLS)){ const i = h.findIndex(x => re.test(x)); if (i >= 0) idx[k] = i; }
      if (idx.code == null || idx.status == null) continue;
      const rows = []; let bad = 0;
      for (const v of aoa.slice(r + 1)){
        const code = txt(v[idx.code]).toUpperCase(); if (!code) continue;
        if (!/^[A-Z]{2,6}$/.test(code)){ bad++; continue; }
        const s = lc(v[idx.status]);
        const status = /discont|stopp|ended|inactive/.test(s) ? 'discontinued' : /pause|hold/.test(s) ? 'paused' : /complet|graduat|finish/.test(s) ? 'completed' : 'active';
        rows.push({ code, clinic: idx.clinic != null ? txt(v[idx.clinic]) : '', status, lastDose: idx.lastDose != null ? TE.parseDate(v[idx.lastDose]) : null, reason: idx.reason != null ? txt(v[idx.reason]) : '', ret: idx.ret != null ? TE.parseDate(v[idx.ret]) : null });
      }
      if (bad && bad >= rows.length) return { error: 'The patient column has names or IDs, not LASFIR codes. Export LASFIR only.' };
      return { rows, skipped: bad };
    }
    return { error: 'Couldn’t find LASFIR and Status columns. Use the ESK Tracker’s Export for Appointments file.' };
  }

  return { DEFAULTS, HIST, CALL, visits, chains, gapStats, cadence, recall, samples, slotRoles, flexTypes, flexUsage, releaseList, slotMatches, funnel, readEskStatus, eskCredentials, addBusinessDays };
});
