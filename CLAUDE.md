# Working on clinic-template-tool

Hard rules:
- `core.js`, `rules.js` and `payer-rules.js` are shared byte-identical with the Sohar Eligibility add-in (`payer-rules.js` also with the ESK Tracker). Never edit them here alone: change them in Sohar (with a Sohar test), then copy them over and check with `git hash-object core.js rules.js payer-rules.js` on both sides.
- Every release: bump `APP_VERSION` in `app.js`, `version` in `package.json`, and the `?v=` on every local script in `taskpane.html` (all the same value). Excel caches add-in files; without a new `?v=`, users keep running the old code.
- Patients show as LASFIR (+ MRN in the Clinic Eligibility views) only. Never show or write patient names or dates of birth.
- Synthetic data only in tests, fixtures and commits (made-up names, MRNs 9000xx). The real Payer Credential Rules table and any export never go in this repo.
- No network calls. The CSP in `taskpane.html` allows `connect-src 'self'` only, for the install page's manifest download.
- Workbooks without an `SE Info` tab must behave exactly as before. Clinic Eligibility sheets (`SE *`, `Clinic Work`, `Read Me`) are never template grids or the appointment list.
- Full screen is an Office dialog that relays workbook calls through the pane: a new store op goes in `RELAY_OPS` and `RemoteStore`, and a new document setting goes in the `docs` list in `openFull`.
- Clinic Work is written whole through `S.store.writeSheet` (rows from `SoharRules.workRows`) with `{ inPlace: true, text: true }`: as text (`'@'` before the values, so leading-zero MRNs and notes starting with `=` survive) and over the old rows, never clear-first (a failed write must leave the clinic's earlier ticks and notes). Its columns are a contract with Sohar's `bringBackClinicFiles`: `CLINIC_WORK_COLS`.

Workflow: edit → `npm test` → check the pane at about 350 px and full screen → commit with a plain-English message → push `main` → confirm GitHub Pages is live.
