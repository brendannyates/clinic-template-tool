// node --test tests/cadence.test.js   (synthetic data only: made-up patients and providers)
const test = require('node:test');
const assert = require('node:assert/strict');
const TE = require('../engine.js');
const TC = require('../cadence.js');

const D = (y, m, d) => TE.serialFromYMD(y, m, d);
const MON = D(2026, 3, 2);   // a Monday
const t = (h, m = 0) => (h * 60 + m) / 1440;
const near = (a, b, eps = 1e-6) => assert.ok(Math.abs(a - b) < eps, `${a} ≉ ${b}`);

test('LASFIR follows the Confirmations rule', () => {
  assert.equal(TE.lasfir('Smith, John (Tom)'), 'SMIJOH');
  assert.equal(TE.lasfir('Ng, Bo'), 'NGBO');
  assert.equal(TE.lasfir('O’Brien-Lee, Ána Maria'), 'OBRANA');
  assert.equal(TE.lasfir('John Smith'), '');            // not "Last, First": no guess
  assert.equal(TE.lasfir(''), '');
});

test('business hours run Monday to Friday', () => {
  const fri3 = MON + 4 + t(15), mon11 = MON + 7 + t(11);
  assert.equal(TE.businessHours(fri3, mon11), 6);       // Friday 3-6 pm + Monday 8-11 am
  assert.equal(TE.businessHours(mon11, fri3), 0);
  near(TE.minusBusinessHours(mon11, 24), MON + 2 + t(17));   // 24 business hours before Monday 11 am = Wednesday 5 pm
  near(TE.minusBusinessHours(MON + t(10), 2), MON + t(8));
});

test('visit roles from the event, with the Ref Role column winning', () => {
  const r = (e, sl, given) => TE.roleOf(e, e, sl, given);
  assert.equal(r('Provider MT', 'TMS'), 'start');
  assert.equal(r('Esketamine Start', 'ESK'), 'start');
  assert.equal(r('Esketamine Follow Up', 'ESK'), 'treatment');
  assert.equal(r('IV Ketamine Start/Index', 'IVK'), 'treatment');
  assert.equal(r('TMS Consult In Person', 'TMS'), 'consult');
  assert.equal(r('Tel New Patient - Mature', 'Medication Management'), 'consult');
  assert.equal(r('TMS Follow Up', 'TMS'), 'followup');
  assert.equal(r('TMS Treatment', 'TMS'), 'treatment');
  assert.equal(r('Med Management Telepsych', 'Medication Management'), 'followup');
  assert.equal(r('IV Ketamine Follow Up', 'IVK', 'Follow up'), 'followup');
  const ref = TE.buildRef([{ event: 'Weird Visit', cls: 'Thing', sl: 'TMS', role: 'Start' }]);
  assert.equal(TE.classify(ref, 'Weird Visit').role, 'start');
});

test('the appointment list reads Pat Name, Md Rc and Loc Name', () => {
  const ap = TE.parseAppts([['Appt Date', 'Beg Tm', 'Rendering', 'Event', 'Status', 'Pat Name', 'Md Rc', 'Loc Name'], ['3/2/2026', '9:00 AM', 'Smith MD, Jane', 'TMS Consult In Person', 'Kept', 'Doe, Alex', '1001', 'Eastside Clinic']]);
  assert.equal(ap.has.pt, true); assert.equal(ap.has.mrn, true); assert.equal(ap.rows[0].loc, 'Eastside Clinic'); assert.equal(ap.rows[0].start, 540);
  assert.equal(ap.rows[0].pt, 'Doe, Alex'); assert.equal(ap.rows[0].mrn, '1001');
});

// ---------- one synthetic clinic ----------
const P = 'Smith MD, Jane', CHAIR = 'Eastside MagV1';
const REF = [
  { event: 'TMS Consult In Person', sched: 'Clinical', cls: 'Consult', sl: 'TMS', elig: 'Always' },
  { event: 'Provider MT', sched: 'Clinical', cls: 'TMS Start', sl: 'TMS', elig: 'Always' },
  { event: 'TMS Treatment', sched: 'Clinical', cls: 'TMS Treatment', sl: 'TMS', elig: 'Always' },
  { event: 'TMS Follow Up', sched: 'Clinical', cls: 'TMS Follow Up', sl: 'TMS', elig: 'Always' },
  { event: 'Med Management Telepsych', sched: 'Clinical', cls: 'Medication Management', sl: 'Medication Management', elig: 'Always' },
  { event: 'Esketamine Follow Up', sched: 'Clinical', cls: 'EFU', sl: 'ESK', elig: 'Always' },
];
const ST = [
  { type: 'Interventional Consult', kind: 'Clinical', accepts: 'Consult' }, { type: 'Provider MT', kind: 'Clinical', accepts: 'TMS Start' },
  { type: 'Telepsych', kind: 'Clinical', accepts: 'Medication Management' }, { type: 'Open Schedule', kind: 'Clinical', accepts: '*' },
  { type: 'New Patient', kind: 'Clinical', accepts: 'Intake' },
];
function grid(){
  return [[], ['Jane Smith, MD'], ['Effective From:', ''], ['Location:', 'Eastside', 'Weekly Clinic Template'], ['Provider:', 'Smith', 'Smith', 'Smith', 'Smith', 'Smith'],
    ['Time', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'],
    [t(9), 'Interventional Consult', 'Telepsych', '', '', ''], [t(9, 30), 'Provider MT', 'New Patient', '', '', ''], [t(10), 'Telepsych', '', '', '', ''], [t(10, 30), 'Open Schedule', '', '', '', ''], [t(11)]];
}
const v = (pt, mrn, date, h, m, event, status = 'Kept', extra = {}) => Object.assign({ date, start: h * 60 + m, dur: 30, prov: P, event, status, rawStatus: status, loc: 'Eastside', booked: date - 30, cancelAt: null, pt, mrn }, extra);
const ASOF = MON + 38;    // a Thursday
const NOW = ASOF + t(12);
function clinic(){
  const A = ['Doe, Alex', '1001'], B = ['Roe, Bea', '1002'], C = ['Poe, Cy', '1003'], Dd = ['Lo, Di', '1004'], E = ['Wu, Ed', '1005'];
  const rows = [
    // Medication Management: A every 3 weeks (plus a same-day duplicate and a no-show), B 4 then 3 weeks and booked ahead
    ...[-84, -63, -42, -21, 0].map(d => v(...A, MON + d, 10, 0, 'Med Management Telepsych')), v(...A, MON - 42, 10, 0, 'Med Management Telepsych'), v(...A, MON - 35, 10, 0, 'Med Management Telepsych', 'No Show'),
    ...[-49, -21, 0].map(d => v(...B, MON + d, 10, 0, 'Med Management Telepsych')), v(...B, MON + 45, 10, 0, 'Med Management Telepsych', 'Expected'),
    v(...C, MON - 70, 10, 0, 'Med Management Telepsych'), v(...C, MON - 50, 10, 0, 'Med Management Telepsych', 'No Show'),
    v(...Dd, MON + 20, 10, 0, 'Med Management Telepsych'),
    // release bookings: Open Schedule (flex) and a New Patient slot taken by a follow-up booked the same morning
    v('Lee, Lu', '1012', MON - 7, 10, 30, 'Med Management Telepsych', 'Kept', { booked: MON - 7 + t(8) }),
    v('Lee, Lu', '1012', MON - 6, 9, 30, 'Med Management Telepsych', 'Kept', { booked: MON - 6 + t(8) }),
    v('May, Mo', '1013', MON - 13, 9, 30, 'Med Management Telepsych'),
    // TMS: E consults, maps, 30 treatments on a chair, then a follow-up
    v(...E, MON - 120, 9, 0, 'TMS Consult In Person'), v(...E, MON - 100, 9, 30, 'Provider MT', 'Kept', { dur: 60 }), v(...E, MON - 50, 10, 0, 'TMS Follow Up'),
    v('Ivy, Ian', '1009', MON - 60, 9, 0, 'TMS Consult In Person'), v('Ivy, Ian', '1009', MON - 45, 9, 30, 'Provider MT', 'Kept', { dur: 60 }),
    v('Jay, Jo', '1010', MON - 140, 9, 0, 'TMS Consult In Person'),
    v('Kay, Kim', '1011', MON + 10, 9, 0, 'TMS Consult In Person'),
    // ESK: F weekly then stopped; G paused and H discontinued in the tracker
    ...[-70, -63, -56, -49, -42, -35, -28].map(d => v('Fox, Flo', '1006', MON + d, 11, 0, 'Esketamine Follow Up')),
    v('Gil, Gus', '1007', MON - 80, 11, 0, 'Esketamine Follow Up'), v('Hu, Hal', '1008', MON - 80, 11, 0, 'Esketamine Follow Up'),
  ];
  for (let d = MON - 98, n = 0; n < 30; d++){ const w = TE.dow(d); if (w === 0 || w === 6) continue; rows.push(v(...E, d, 8, 0, 'TMS Treatment', 'Kept', { prov: CHAIR })); n++; }
  const opts = { asOf: ASOF, from: MON - 150, to: ASOF, releaseHours: 24 };
  const R = TE.analyze({ templates: TE.parseTemplates('Tab', grid()), appts: rows, ref: REF, slotTypes: ST, providers: [], opts });
  const vs = TC.visits(rows, REF, opts, new Set(R.people)).list;
  return { R, vs, opts };
}

test('cadence: return intervals per provider and service line', () => {
  const { vs, opts } = clinic();
  const cad = TC.cadence(vs, opts);
  const mm = cad.rows.find(r => r.prov === P && r.bucket === 'Medication Management');
  assert.equal(mm.n, 7);                                   // A: 21,21,21,21  B: 28,21  L: 1  (A's same-day duplicate collapsed)
  assert.equal(mm.modeWk, 3); near(mm.modeShare, 5 / 7); assert.equal(mm.median, 21); assert.equal(mm.expDays, 21);
  assert.equal(mm.returning, 3); assert.equal(mm.ns, 2);
  assert.deepEqual(mm.hist, [1, 0, 5, 1, 0, 0, 0, 0]);
  // TMS treatments on the chair belong to the provider the patient last saw; the consult isn't in the chain
  const tms = cad.rows.find(r => r.prov === P && r.bucket === 'TMS');
  assert.ok(tms.kept >= 32); assert.equal(tms.modeWk, 1);
  assert.ok(!cad.rows.some(r => r.prov === CHAIR));
  const a = cad.patients.find(p => p.code === 'DOEALE'); assert.equal(a.kept, 5); assert.equal(a.lastGap, 21);
});

test('recall states: active, current, overdue, lapsed, completed, maintenance, tracker states, unaccounted', () => {
  const { vs, opts } = clinic();
  const cad = TC.cadence(vs, opts);
  const esk = TC.readEskStatus([['LASFIR', 'Clinic', 'Status', 'Last Dose Date', 'Pause Reason', 'Expected Return Date'], ['GILGUS', 'Eastside', 'Paused', '1/1/2026', 'Travel', '4/20/2026'], ['HUHAL', '', 'Discontinued', '', '', '']]);
  const rc = TC.recall(vs, opts, cad, esk);
  const st = (code, b) => rc.find(r => r.code === code && (!b || r.bucket === b));
  assert.equal(st('DOEALE').state, 'overdue');             // 38 days since, usual interval 21 → over 1.5×
  assert.equal(st('ROEBEA').state, 'active');
  assert.equal(st('LODI').state, 'current');
  assert.equal(st('POECY').state, 'lapsed'); assert.equal(st('POECY').post, 'No Show');
  assert.equal(st('WUED', 'TMS').state, 'completed');
  assert.equal(TC.recall(vs, Object.assign({}, opts, { maintDays: 60 }), cad, esk).find(r => r.code === 'WUED').state, 'maintenance');
  assert.equal(st('GILGUS').state, 'paused'); assert.equal(st('GILGUS').ret, D(2026, 4, 20));
  assert.equal(st('HUHAL').state, 'discontinued');
  assert.equal(st('FOXFLO').state, 'unaccounted');
  assert.equal(TC.recall(vs, opts, cad, null).find(r => r.code === 'FOXFLO').state, 'lapsed');   // no tracker loaded
  assert.ok(!rc.some(r => r.code === 'JAYJO' || r.code === 'KAYKIM'));                     // consult only: funnel, not recall
  assert.ok(TC.CALL.has(rc[0].state));                                                      // call list sorts first
});

test('funnel: conversion leaves out pending consults and sizes start hours from kept consults', () => {
  const { R, vs, opts } = clinic();
  const f = TC.funnel(vs, R, opts).find(r => r.prov === P && r.sl === 'TMS');
  assert.equal(f.consults, 4); assert.equal(f.converted, 2); assert.equal(f.pending, 1);
  near(f.rate, 2 / 3); assert.equal(f.medDays, 17.5); assert.equal(f.startLen, 60);
  near(f.startHoursNeeded, 4 * 2 / 3); assert.equal(f.startHoursUsed, 2);
  assert.equal(f.consultSlots, 54);                        // Interventional Consult + Open Schedule, 27 Mondays each
  assert.equal(f.startHoursOffered, 27); assert.equal(f.flag, 'over');
});

test('release bookings never count against template fit; flex usage shows who filled the slot', () => {
  const { R, vs } = clinic();
  const np = R.byType.find(x => x.key === 'New Patient');
  assert.equal(np.keptRel, 1); assert.equal(np.keptOff, 1); assert.equal(np.fit, 0);
  const fu = TC.flexUsage(R, vs, {});
  const open = fu.find(x => x.type === 'Open Schedule');
  assert.ok(open && open.used >= 1); assert.equal(open.roles.followup, 1); assert.equal(open.release, 1);
  assert.ok(!fu.some(x => x.type === 'Telepsych'));        // a single-role slot isn't flex
});

test('release window list and open-slot matches for the call list', () => {
  const { R, vs, opts } = clinic();
  const rel = TC.releaseList(R, opts, NOW);
  assert.equal(rel.length, 6); assert.equal(rel.filter(x => x.released).length, 4);
  near(rel.find(x => !x.released).enters, MON + 38 + t(15));   // Tuesday 9 am enters the window Thursday 3 pm
  const rc = TC.recall(vs, opts, TC.cadence(vs, opts), null);
  const m = TC.slotMatches(rc, R, vs, opts, NOW).filter(x => x.code === 'DOEALE');
  assert.equal(m.length, 3);
  assert.deepEqual(m.map(x => [x.date - MON, x.type, x.why]), [[42, 'Telepsych', 'dedicated'], [42, 'Open Schedule', 'flex'], [43, 'Telepsych', 'dedicated']]);
  // a released slot with the patient's own provider counts only if it serves their service line
  const tms = TC.slotMatches(rc.filter(r => r.code === 'IVYIAN'), R, vs, opts, NOW);
  assert.ok(tms.length && tms.every(x => x.type !== 'Telepsych'));
  assert.ok(tms.some(x => x.why === 'released' && x.type === 'Interventional Consult'));
});

test('ESK status import refuses names', () => {
  assert.ok(TC.readEskStatus([['LASFIR', 'Status'], ['Doe, Alex', 'Active']]).error);
  assert.ok(TC.readEskStatus([['Name', 'Thing']]).error);
  const ok = TC.readEskStatus([['LASFIR', 'Status'], ['doeale', 'On hold'], ['ROEBEA', 'Graduated']]);
  assert.deepEqual(ok.rows.map(r => [r.code, r.status]), [['DOEALE', 'paused'], ['ROEBEA', 'completed']]);
});
