// node --test tests/   (synthetic data only)
const test = require('node:test');
const assert = require('node:assert/strict');
const TE = require('../engine.js');

const D = (y, m, d) => TE.serialFromYMD(y, m, d);
const MON = D(2026, 3, 2);   // a Monday
const t = (h, m = 0) => (h * 60 + m) / 1440;

test('dates, times, weeks', () => {
  assert.equal(TE.dow(MON), 1);
  assert.equal(TE.weekStart(MON), MON - 1);                         // Sunday
  assert.equal(TE.parseTime('1:30 P'), 13 * 60 + 30);
  assert.equal(TE.parseTime('12:00 A'), 0);
  assert.equal(TE.parseTime('12:30 PM'), 12 * 60 + 30);
  assert.equal(TE.parseTime(0.375), 9 * 60);
  assert.equal(TE.parseDate('3/2/2026'), MON);
  assert.equal(TE.parseDate('3/2/26'), MON);
  assert.equal(TE.parseDate(MON + 0.5), MON);
  assert.ok(Math.abs(TE.parseDateTime('3/2/2026 2:15 PM') - (MON + t(14, 15))) < 1e-9);
  assert.equal(TE.iso(MON), '2026-03-02');
});

test('provider names', () => {
  const names = ['Patel MD, Priya Anne', 'Okafor PMHNP, Jordan', 'Reyes MD, Samuel Lee', 'Kim, MD Dana', 'Northside MagV1', 'Eastside IVK 1'];
  assert.equal(TE.matchProvider('Patel', names), 'Patel MD, Priya Anne');
  assert.equal(TE.matchProvider('Jordan Okafor, NP', names), 'Okafor PMHNP, Jordan');
  assert.equal(TE.matchProvider('Samuel Reyes, MD', names), 'Reyes MD, Samuel Lee');
  assert.equal(TE.matchProvider('Kim', names), 'Kim, MD Dana');
  assert.equal(TE.isPersonResource('Northside MagV1'), false);
  assert.equal(TE.isPersonResource('Eastside IVK 1'), false);
  assert.equal(TE.isPersonResource('Kim, MD Dana'), true);
  assert.equal(TE.credentialOf('Okafor PMHNP, Jordan'), 'NP');
  assert.equal(TE.credentialOf('Moreau DO, Alex'), 'MD');
});

// a grid in the provider-configurable layout (labels above, Time column, Monday..Friday)
function grid({ label = 'Smith', loc = 'Eastside', from = '', withTimeLabel = true } = {}){
  const g = [[], ['Jane Smith, MD'], ['Effective From:', from], ['Location:', loc, 'Weekly Clinic Template', '', 'Last Modified:', D(2026, 2, 1)], ['Provider:', label, label, label, label, label],
    [withTimeLabel ? 'Time' : '', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']];
  const rows = [[t(9), 'Interventional Consult', 'Telepsych', '', '', ''], [t(9, 30), 'Admin', 'Telepsych', '', '', ''], [t(10), 'New Patient', '', '', '', ''], [t(10, 30), 'Lunch', '', '', '', '']];
  return g.concat(rows, [[t(11)]]);
}

test('template grids parse in both layouts', () => {
  const [a] = TE.parseTemplates('Tab', grid());
  assert.equal(a.label, 'Smith'); assert.equal(a.location, 'Eastside'); assert.equal(a.step, 30);
  assert.equal(a.slots.length, 6); assert.deepEqual(a.days, [1, 2]);
  assert.equal(a.modified, D(2026, 2, 1)); assert.equal(a.effFrom, null);
  const stacked = grid({ label: 'Jones', withTimeLabel: false, from: D(2026, 3, 1) }).concat([[], []], grid({ label: 'Lee' }));
  const both = TE.parseTemplates('Stack', stacked);
  assert.deepEqual(both.map(x => x.label), ['Jones', 'Lee']);
  assert.equal(both[0].effFrom, D(2026, 3, 1));
});

const REF = [
  { event: 'TMS Consult In Person', sched: 'Clinical', cls: 'Consult', sl: 'TMS', elig: 'Always' },
  { event: 'Med Management Telepsych', sched: 'Clinical', cls: 'Medication Management', sl: 'Medication Management', elig: 'Always' },
  { event: 'Tel New Patient - Mature', sched: 'Clinical', cls: 'Intake', sl: 'Medication Management', elig: 'Always' },
  { event: 'Provider PTO', sched: 'Clinical', cls: 'PTO', sl: 'N/A', elig: 'Always' },
  { event: 'Meeting', sched: 'Clinical', cls: 'Meeting', sl: 'N/A', elig: 'Always' },
  { event: 'MT Hold - MagVenture', sched: 'Reference/Hold Only', cls: 'MT Hold', sl: 'TMS', elig: 'Always' },
];
const ST = TE.DEFAULT_SLOT_TYPES.map(r => ({ type: r[0], kind: r[1], accepts: r[2] }));
const P = 'Smith MD, Jane';
const ap = (date, h, m, event, status, extra = {}) => Object.assign({ date, start: h * 60 + m, dur: 30, prov: P, event, status, loc: 'Eastside', booked: null, cancelAt: null }, extra);

function scenario(appts, opts = {}){
  const tpl = TE.parseTemplates('Tab', grid());
  return TE.analyze({ templates: tpl, appts, ref: REF, slotTypes: ST, providers: [], opts: Object.assign({ asOf: MON + 7 }, opts) });
}

test('slot outcomes, fit, no-show, admin, double-book and blocked time', () => {
  const R = scenario([
    ap(MON, 9, 0, 'TMS Consult In Person', 'Kept'),             // consult slot, intended visit
    ap(MON, 9, 30, 'Med Management Telepsych', 'Kept'),         // admin slot: visit in blocked time
    ap(MON, 10, 0, 'Med Management Telepsych', 'Kept'),         // new patient slot, other visit
    ap(MON + 1, 9, 0, 'Med Management Telepsych', 'No Show'),   // telepsych slot, no-show
    ap(MON + 1, 9, 30, 'Med Management Telepsych', 'Kept'),
    ap(MON + 1, 9, 30, 'Med Management Telepsych', 'Kept'),     // double-booked
    ap(MON + 7, 9, 0, 'Meeting', 'Kept'),                       // a meeting takes the consult slot (as-of day counts as past)
    ap(MON + 7, 10, 0, 'Med Management Telepsych', 'Cancelled'),
    ap(MON + 7, 9, 0, 'MT Hold - MagVenture', 'Expected'),      // holds are ignored
  ]);
  const T = R.total;
  assert.equal(T.supply, 6);           // Mon + Tue of week 1, Mon of week 2 (the range ends on the as-of date)
  assert.equal(T.keptFit, 2);          // consult + one telepsych
  assert.equal(T.keptOff, 1);          // med mgmt in a new patient slot
  assert.equal(T.ns, 1);
  assert.equal(T.admin, 1);
  assert.equal(T.aKept, 5); assert.equal(T.aNS, 1); assert.equal(T.aCx, 1);
  assert.equal(T.inBlocked, 1); assert.equal(T.dbl, 1);
  assert.equal(R.exceptions.length, 2);
  assert.match(R.exceptions[0].why, /In Admin/);
  assert.deepEqual(R.offMix['New Patient'], { 'Med Management Telepsych': 1 });
  assert.equal(R.hasLate, false); assert.equal(T.lcRate, null);
});

test('late cancels need a cancel time', () => {
  const R = scenario([
    ap(MON, 9, 0, 'TMS Consult In Person', 'Cancelled', { cancelAt: MON + t(8) }),       // 1 h before: late
    ap(MON, 10, 0, 'Tel New Patient - Mature', 'Cancelled', { cancelAt: MON - 3 }),     // 3 days before: not late
    ap(MON + 1, 9, 0, 'Med Management Telepsych', 'Kept'),
  ]);
  assert.equal(R.hasLate, true);
  assert.equal(R.total.aLC, 1); assert.equal(R.total.aCx, 2);
  assert.ok(Math.abs(R.total.lcRate - 1 / 3) < 1e-9);
});

test('PTO removes supply; 3NA counts calendar days to the third open slot', () => {
  const asOf = MON + 2;   // Wednesday
  const R = scenario([
    ap(MON, 9, 0, 'Provider PTO', 'Kept', { dur: 90 }),                    // PTO covers Monday 9:00-10:30
    ap(MON + 7, 9, 0, 'TMS Consult In Person', 'Expected'),               // next Monday's consult is taken
    ap(MON + 60, 9, 0, 'MT Hold - MagVenture', 'Expected'),              // report runs this far ahead
  ], { asOf });
  assert.equal(R.total.pto, 2);   // consult + new patient on Monday
  const c = R.threeNa.find(x => x.prov === P && x.type === 'Interventional Consult');
  // open consult slots after Wed 3/4: 3/16, 3/23, 3/30... (3/9 is booked)
  assert.equal(c.next, MON + 14 - asOf);
  assert.equal(c.third, MON + 28 - asOf);
  const tele = R.threeNa.find(x => x.prov === P && x.type === 'Telepsych');
  assert.equal(tele.third, MON + 15 - asOf);  // Tue 3/10 9:00 and 9:30, then Tue 3/17
});

test('templates with Effective From dates apply only from that date', () => {
  const old = TE.parseTemplates('Old', grid())[0];
  const neu = TE.parseTemplates('New', grid({ from: MON + 7 }))[0];
  neu.slots = neu.slots.filter(s => s.dow === 1);   // new version drops Tuesday
  const R = TE.analyze({ templates: [old, neu], appts: [ap(MON, 9, 0, 'TMS Consult In Person', 'Kept'), ap(MON + 14, 9, 0, 'TMS Consult In Person', 'Kept')], ref: REF, slotTypes: ST, providers: [], opts: { asOf: MON + 14 } });
  const tue = R.byDow.find(x => x.key === 'Tuesday');
  assert.equal(tue.supply, 2);   // only the first week's Tuesday
});

test('3NA history is rebuilt from created and cancel dates', () => {
  const asOf = MON + 21;
  const R = scenario([
    ap(MON + 7, 9, 0, 'TMS Consult In Person', 'Kept', { booked: MON - 10 }),
    ap(MON + 14, 9, 0, 'TMS Consult In Person', 'Cancelled', { booked: MON - 10, cancelAt: MON + 10 }),
    ap(MON + 60, 9, 0, 'MT Hold - MagVenture', 'Expected'),
  ], { asOf, from: MON - 7 });
  assert.ok(R.threeNaHist && R.threeNaHist.length);
  const wk = R.threeNaHist.find(h => h.week === MON - 1 && h.type === 'Interventional Consult');
  // on Sunday 3/1 the 3/2 slot is open, 3/9 and 3/16 are booked, then 3/23, 3/30
  assert.equal(wk.third, MON + 28 - (MON - 1));
});

test('access model import ignores patient tabs and checks the schema', () => {
  const ok = TE.readAccessModel({ About: [['Field', 'Value'], ['Schema', 'mhs-access-model'], ['Schema version', 1], ['Modality', 'ESK'], ['Clinic', 'SR']], Providers: [['Modality', 'Provider', 'Need per week'], ['ESK', 'Patel', 12]], Patients: [['MRN'], ['123']] });
  assert.equal(ok.modality, 'ESK'); assert.equal(ok.providers[0]['Need per week'], 12); assert.equal(ok.patients, undefined);
  assert.ok(TE.readAccessModel({ About: [['Field', 'Value'], ['Schema', 'other']] }).error);
});

test('appointment list header, run date and status words', () => {
  const v = [['', 'Mindful Health Solutions'], ['', 'Appointments Listing'], ['', 'From 1/1/2026 to 12/31/2026'], ['9/14/26 12:59 PM'], [],
    ['', 'MRN', 'Location', 'Appt Date', 'Time', 'End Tm', 'Duration', 'Provider', 'Event', 'Status', 'Cancel Date', 'Created Date'],
    ['', '1', 'Eastside', MON, '9:00 A', '9:30 A', 30, P, 'TMS Consult In Person', 'Kept', '', '2/20/2026'],
    ['', '2', 'Eastside', '3/3/2026', '1:30 P', '2:00 P', 30, P, 'Meeting', 'Cancelled', '3/2/2026 4:00 PM', ''],
    ['Total Appointments (2)']];
  const r = TE.parseAppts(v);
  assert.equal(r.rows.length, 2); assert.equal(r.rows[1].date, MON + 1); assert.equal(r.rows[1].start, 13 * 60 + 30);
  assert.equal(Math.floor(r.runAt), D(2026, 9, 14));
  assert.ok(r.has.cancel && r.has.booked);
  assert.ok(Math.abs(r.rows[1].cancelAt - (MON + t(16))) < 1e-9);
  assert.equal(r.rows[0].booked, D(2026, 2, 20));
});

test('provider and location filters apply to visit rates, not just slots', () => {
  const Q = 'Okafor PMHNP, Jordan';
  const appts = [
    ap(MON, 9, 0, 'TMS Consult In Person', 'Kept'),
    ap(MON + 1, 9, 0, 'Med Management Telepsych', 'No Show'),
    ap(MON, 9, 0, 'Med Management Telepsych', 'No Show', { prov: Q, loc: 'Westside' }),
    ap(MON, 10, 0, 'Med Management Telepsych', 'No Show', { prov: Q, loc: 'Westside' }),
    ap(MON + 1, 10, 0, 'Med Management Telepsych', 'Cancelled', { prov: Q, loc: 'Westside' }),
  ];
  const all = scenario(appts), p = scenario(appts, { provs: new Set([P]) }), q = scenario(appts, { provs: new Set([Q]) });
  assert.equal(all.total.aAll, 5);
  assert.equal(p.total.aAll, 2); assert.equal(p.total.nsRate, 0.5);
  assert.equal(q.total.aAll, 3); assert.equal(q.total.nsRate, 1); assert.equal(q.total.supply, 0);   // no template for Q
  assert.equal(scenario(appts, { locs: new Set(['Westside']) }).total.aAll, 3);
});

test('slots still Expected on the as-of day stay out of fill', () => {
  const R = scenario([
    ap(MON + 7, 9, 0, 'TMS Consult In Person', 'Kept'),
    ap(MON + 7, 10, 0, 'Tel New Patient - Mature', 'Expected'),   // later today
  ]);
  const T = R.total;
  assert.equal(T.unres, 1);
  assert.equal(T.supply, 1);          // the window is just the as-of Monday: consult + new patient, minus the unresolved one
  assert.equal(T.fill, 1);
});

test('NextGen "Crt Dt" is read as the created date', () => {
  const v = [['Appt Date', 'Time', 'Provider', 'Event', 'Status', 'Duration', 'Location', 'Crt Dt', 'Mod Dt'],
    [MON, t(9), P, 'TMS Consult In Person', 'Kept', 30, 'Eastside', MON - 10, MON - 2]];
  const r = TE.parseAppts(v);
  assert.ok(r.has.booked);
  assert.equal(r.rows[0].booked, MON - 10);
});

test('copyPlan keeps one grid from a stacked tab and rewrites its header', () => {
  const v = grid({ label: 'Jones', withTimeLabel: false }).concat([[], []], grid({ label: 'Lee', from: D(2026, 3, 1) }));
  const ts = TE.parseTemplates('Stack', v);
  for (const [i, empty] of [[0, false], [1, false], [1, true]]){
    const plan = TE.copyPlan(ts[i], ts, v, { label: 'Kim', title: 'Dana Kim, MD', location: 'Northside', from: D(2026, 9, 1), today: D(2026, 8, 1), empty });
    assert.ok(!plan.error, plan.error);
    const w = v.map(r => r.slice());
    for (const [r, c, x] of plan.writes){ while (w[r].length <= c) w[r].push(''); w[r][c] = x; }
    if (plan.clearSlots){ const [r0, c0, h, n] = plan.slots; for (let r = r0; r < r0 + h; r++) for (let c = c0; c < c0 + n; c++) w[r][c] = ''; }
    for (const [a, b] of plan.deleteRows) w.splice(a, b - a + 1);
    const back = TE.parseTemplates('Copy', w);
    assert.equal(back.length, 1);
    assert.equal(back[0].label, 'Kim'); assert.equal(back[0].location, 'Northside'); assert.equal(back[0].effFrom, D(2026, 9, 1)); assert.equal(back[0].modified, D(2026, 8, 1));
    assert.equal(back[0].slots.length, empty ? 0 : 6);
    assert.equal(back[0].anchor[0] + 1, plan.slotsAfter[0]);
    assert.ok(w.every(r => !r.includes('Jones') && !r.includes('Lee')), 'the old label is gone from every cell');
  }
});

test('copyPlan adds an Effective From label when the grid has none', () => {
  const v = [[], [], ['', 'Location:', 'Eastside', 'Provider:', 'Smith', 'Last Modified:', D(2026, 2, 1)], ['', 'Time', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'],
    ['', t(9), 'Telepsych', 'Admin', '', '', ''], ['', t(9, 30), 'Telepsych', '', '', '', 'Telepsych'], ['', t(10)]];
  const [g] = TE.parseTemplates('Tab', v);
  const plan = TE.copyPlan(g, [g], v, { label: 'Smith', location: '', from: D(2026, 9, 1), today: D(2026, 8, 1) });
  const w = v.map(r => r.slice()); for (const [r, c, x] of plan.writes){ while (w[r].length <= c) w[r].push(''); w[r][c] = x; }
  assert.equal(TE.parseTemplates('Tab', w)[0].effFrom, D(2026, 9, 1));
  assert.equal(TE.parseTemplates('Tab', w)[0].location, 'Eastside');   // blank location in the dialog leaves it alone
});
