// pricing.js (CTT 1.2.0): price a visit the rotation has not, from the clinic file's own fee tables (SE Fees and SE Fee
// Rules), with Sohar's fees.js (a byte-identical copy). Pure: no Office.js, no DOM, no network. The amounts are estimates
// for the front desk and say so on Clinic Day ("(est.)"); the rotation's ledger estimate always wins when there is one.
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CTTPricing = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  const str = (v) => (v == null ? '' : String(v).trim());
  // NP, PA and PMHNP render at the NP level of the fee schedule; everyone else at MD.
  const levelOf = (provider) => (/\b(pmhnp|np|fnp|pa|pa-c|aprn)\b/i.test(str(provider)) ? 'np' : 'md');

  /**
   * pack: SR.clinicUnpack output. Libraries: {SR (rules.js), SF (fees.js), SC (core.js)}. Returns priceOf(appt) ->
   * {amount} or {note}, or null when the file carries no fee tables (older files, fee estimates Off at the admin).
   */
  function pricer(pack, lib) {
    const { SR, SF, SC } = lib || {};
    if (!pack || !SR || !SF || !SC || !Array.isArray(pack.fees) || !pack.fees.length) return null;
    const ft = SR.feeTables(pack.feeRules || []);
    const cfg = ft.config || {};
    // A visit here is priced at this clinic's fee tab, whatever the patient's next visit anywhere is (a guest's is elsewhere).
    const here = new Map();
    for (const a of pack.appts || []) if (!here.has(str(a['Per Nbr']))) here.set(str(a['Per Nbr']), a);
    const results = (pack.results || []).map((r) => {
      const a = here.get(str(r['Per Nbr']));
      return a ? Object.assign({}, r, { 'Next Appt': a.Date, 'Next Time': a.Time, 'Next Event': a.Event, 'Next Location': a.Location, 'Next Provider': a.Provider }) : r;
    });
    let prep;
    try {
      prep = SF.prepareEstimates({
        soharPayers: [], locationState: cfg.locationState || {}, feeRules: cfg.feeRules,
        results, scheduleRows: pack.fees, tabSummary: ft.tabSummary || [],
        locations: ft.locations || [], homes: [], payerMap: ft.payerMap || [], payerLog: [],
        codes: ft.codes || [], defaults: ft.defaults || [], feeFlags: ft.feeFlags || [],
        appts: (pack.appts || []).map((a) => ({ rendering: a.Provider, loc: a.Location, event: a.Event, per: a['Per Nbr'], date: a.Date, status: a.Status })),
        teleKeywords: cfg.teleKeywords || ['Tele'], cfgDefaults: cfg.cfgDefaults, when: '',
        pers: SC.benefitsView(results).rows.map((r) => r['Per Nbr']),
      });
    } catch (e) { return null; }
    const ctxByPer = new Map(Object.values(prep.contexts || {}).map((c) => [str(c.per), c]));
    return function priceOf(appt) {
      const a = appt || {};
      if (SR.isInterventionEvent && SR.isInterventionEvent(a.Event)) return { note: 'No estimate: check DED/OOP' };
      const ctx = ctxByPer.get(str(a['Per Nbr']));
      if (!ctx) return {};
      const ev = SF.codesForEvent(a.Event, ft.codes || []);
      if (ev.none) return { note: ev.note || 'No estimate' };
      let e;
      try { e = SF.estimatePatient(ctx, prep.tables, ev.codes.length ? { codes: ev.codes } : undefined); } catch (err) { return {}; }
      if (e.noEstimate) return { note: `No estimate: ${e.noEstimate}` };
      const lv = e[levelOf(a.Provider)];
      return lv && lv.patient != null ? { amount: lv.patient } : {};
    };
  }

  return { pricer, levelOf };
});
