// lookup.js (CTT 1.2.0): the one network call this add-in makes, to the "Clinic lookup" Power Automate flow, for detail the
// clinic file does not carry. Pure (fetch is passed in): no Office.js, no DOM. Rules (Brendan, 11/10; flagged for Compliance):
// - One address only: the flow's host is LOOKUP_HOST, and taskpane.html's connect-src allows that host and nothing else.
// - The flow address and the clinic key are typed once per computer by the admin and kept in this browser's storage, never
//   in the workbook, the repo or a log. The key decides the clinic; the request never names one.
// - A patient on this clinic's list needs no reason. Anyone else needs one; the flow asks for a note when the patient has
//   no history here, and queues "Request eligibility check" when the patient is not in the master at all.
// - What comes back is shown and dropped: never written to the workbook, Clinic Work or browser storage.
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CTTLookup = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  // The lookup flow's host, set once Brendan saves the flow (the part of its URL before /powerautomate or /workflows).
  // Blank: lookups are off. Must match the connect-src in taskpane.html (tests/lookup.test.js pins the two together).
  const LOOKUP_HOST = '';
  const STORE_KEY = 'cttLookup';
  const REASONS = ['Walk-in or same-day add', 'Patient called this clinic', 'Billing or copay question', 'Moving care to this clinic', 'Other (add a note)'];
  const PURPOSES = ['scheduled', 'unscheduled', 'request-check'];
  // Only these fields are shown, whatever the flow sends (minimum necessary on the way in, too).
  const PATIENT_FIELDS = ['LASFIR', 'MRN', 'Coverage', 'Payer', 'Plan', 'Copay', 'Coinsurance', 'Deductible Remaining', 'Deductible Status',
    'OOP Remaining', 'OOP Status', 'Card On File', 'OK to Collect', 'Next Visit Here', 'As Of'];
  const HISTORY_FIELDS = ['When', 'Kind', 'Detail', 'Amount'];
  const TIERS = ['scheduled', 'seen-12mo', 'no-history', 'not-found'];
  const str = (v) => (v == null ? '' : String(v).trim());

  /** {url, key} as typed -> {ok, error}. The URL must be https on LOOKUP_HOST (host passed in for tests). */
  function checkSetup(cfg, host) {
    const h = host === undefined ? LOOKUP_HOST : host;
    const c = cfg || {};
    if (!h) return { ok: false, error: 'Lookups are not switched on in this version yet.' };
    let u;
    try { u = new URL(str(c.url)); } catch (e) { return { ok: false, error: 'Paste the lookup address the admin gave you.' }; }
    if (u.protocol !== 'https:' || u.hostname.toLowerCase() !== h.toLowerCase()) return { ok: false, error: 'That is not the clinic lookup address. Ask the admin for it again.' };
    if (str(c.key).length < 16) return { ok: false, error: 'Paste the whole clinic key the admin gave you.' };
    return { ok: true };
  }

  /** The request body. q: {mrn, purpose, reason, note, staff, key, requestId}. Throws a readable error when something is missing. */
  function buildRequest(q) {
    const o = q || {};
    const mrn = str(o.mrn).replace(/\s+/g, '');
    if (!/^\d{3,12}$/.test(mrn)) throw new Error('Type the MRN (digits only).');
    const purpose = PURPOSES.includes(o.purpose) ? o.purpose : 'unscheduled';
    const reason = str(o.reason), note = str(o.note).slice(0, 300);
    if (purpose !== 'scheduled' && !reason) throw new Error('Pick why you need this patient\'s details.');
    if (/other/i.test(reason) && !note) throw new Error('Add a short note for "Other".');
    if (!str(o.staff)) throw new Error('Enter your name first (top of the pane).');
    return { v: 1, key: str(o.key), mrn, purpose, reason, note, staff: str(o.staff), requestId: str(o.requestId) || rid() };
  }
  function rid() {
    const c = typeof globalThis !== 'undefined' && globalThis.crypto;
    if (c && c.randomUUID) return c.randomUUID();
    return `r-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
  }

  /** Keep only the fields this add-in shows. Anything else the flow sends is dropped before it reaches the screen. */
  function clean(resp) {
    const r = resp && typeof resp === 'object' ? resp : {};
    const out = { ok: r.ok === true, tier: TIERS.includes(r.tier) ? r.tier : '', error: str(r.error), message: str(r.message).slice(0, 300),
      needReason: r.needReason === true, needNote: r.needNote === true, queued: r.queued === true, patient: null, history: [] };
    if (r.patient && typeof r.patient === 'object') out.patient = Object.fromEntries(PATIENT_FIELDS.map((k) => [k, str(r.patient[k])]));
    if (Array.isArray(r.history)) out.history = r.history.slice(0, 5).map((h) => Object.fromEntries(HISTORY_FIELDS.map((k) => [k, str(h && h[k])])));
    return out;
  }

  /** POST the request to the flow and return clean(response). fetchFn: window.fetch (or a test double). */
  async function call(fetchFn, cfg, body, opts) {
    const ok = checkSetup(cfg, opts && opts.host);
    if (!ok.ok) throw new Error(ok.error);
    const ms = (opts && opts.timeoutMs) || 20000;
    const ctl = typeof AbortController !== 'undefined' ? new AbortController() : null;
    const t = ctl ? setTimeout(() => ctl.abort(), ms) : null;
    let res;
    try {
      // text/plain: a simple request, so the browser sends no CORS preflight the flow cannot answer; the flow's Response
      // carries Access-Control-Allow-Origin for this add-in's site, so the answer is readable.
      res = await fetchFn(str(cfg.url), { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify(Object.assign({}, body, { key: str(cfg.key) })),
        signal: ctl ? ctl.signal : undefined, credentials: 'omit', cache: 'no-store', referrerPolicy: 'no-referrer' });
    } catch (e) {
      throw new Error(e && e.name === 'AbortError' ? 'The lookup took too long. Try again in a minute.' : 'The lookup could not be reached. Check the connection and try again.');
    } finally { if (t) clearTimeout(t); }
    let json = null;
    try { json = await res.json(); } catch (e) { json = null; }
    if (res.status === 401 || res.status === 403) throw new Error('The clinic key was not accepted. Ask the admin to set up the lookup again.');
    if (!res.ok && !json) throw new Error(`The lookup answered ${res.status}. Try again later.`);
    return clean(json);
  }

  return { LOOKUP_HOST, STORE_KEY, REASONS, PURPOSES, PATIENT_FIELDS, HISTORY_FIELDS, TIERS, checkSetup, buildRequest, clean, call };
});
