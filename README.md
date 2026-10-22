# Clinic Template Tool

Excel add-in for Practice Managers. Put copies of provider templates in a workbook, add the NextGen appointment list, and the tool measures how the templates actually get used.

- App: https://brendannyates.github.io/clinic-template-tool/taskpane.html (open it outside Excel to download the manifest)
- Runs in Excel on the web. Data stays in the workbook; this site only serves the app's code.

## What goes in the workbook

| Tab | What it holds |
|---|---|
| Template tabs (any name) | One weekly grid per provider in the provider-configurable layout: `Provider:`, `Location:`, `Effective From:` labels, a `Time` column, Monday–Friday headers, a slot type in each cell. Several grids can share a tab. **Setup → New template tab** makes one; copy an old one to record a revision. A copy duplicates that tab, so its colors, dropdowns and formulas come along. |
| Appointment List | The NextGen *Appointments Listing* report. **Setup → Import NextGen report** writes it. |
| Slot Types | Slot type → Clinical or Blocked, and what it accepts (event class, service line or event; `,` = any of, `+` = all of, `*` = anything). |
| Ref | NextGen event → event class, service line, whether it counts (holds and reference rows don't), and Role (Consult, Start, Treatment, Follow-up). A blank Role is guessed from the event name. |
| Payer Credential Rules | Which payers let an MD or an NP bill each esketamine visit (first treatment, start, observation), and Besse approval. **Setup → Import** the shared file from OneDrive (`Add-In Manifests\Payer Credential Rules.csv`); edit the tab when a payer changes. The table is never in this repo. `payer-rules.js` is shared byte-identical with the Sohar Eligibility and ESK Tracker add-ins. |
| PT ESK Status | ESK Tracker status by LASFIR (active, paused, discontinued, completed). **Recall → Import ESK status** writes it. |
| Providers | NextGen provider → template label, credential, include. Chairs and rooms are left out automatically. |

Hidden tabs are never read.

## What it measures

- **Template fill and fit** by provider, slot type, day of week, hour and AM/PM: slots filled, kept use, kept visits that matched the slot type (and what filled them when they didn't), open slots, slots taken by admin, PTO.
- **Visit outcomes** by service line, event class, appointment type, day and hour: no-show, cancel and late-cancel rates, median lead time.
- **Access**: third next available (3NA, calendar days) by slot type, for each provider and all providers, plus booked-ahead % by week. **Save 3NA snapshot** keeps a weekly history on the `TA 3NA History` tab.
- **Provider impact**: visits in admin/lunch/meeting time, outside template hours or days, double-booked, extra hours.
- **Trends**: weekly (Sunday–Saturday) fill, fit, no-show, cancel and kept volume.
- **ESK access model**: imports the ESK Tracker's *Export for template tool* file (`mhs-access-model` v1) and compares panel need with the ESK slots in the template. IV ketamine will use the same format.

- **Esketamine payer rules** (Access tab): ESK visits booked with a provider whose credential the patient's payer won't pay for, upcoming (to move) and past (billing risk), plus a by-payer summary of what NPs may bill and Besse approval. Needs **Payer Name** on the report and the Payer Credential Rules tab.
- **Release bookings**: a visit booked less than 24 business hours ahead (Mon–Fri, 8–6, configurable) may take any slot, so it never counts against template fit.
- **Cadence**: per provider and service line, how often patients come back: kept visits, return gaps (mean, median, usual return in weeks and its share), no-shows, a gap histogram, and a per-patient list. Treatments on chairs count toward the provider the patient last saw. Consults start a chain but aren't part of it.
- **Funnel**: consults offered and kept, how many led to a start in the same service line within 90 days (recent consults are pending, not failures), and start hours offered and used against start hours needed (kept consults × conversion × start length).
- **Flex slots**: for slot types that accept more than one kind of visit, which kinds actually filled them.
- **Recall**: each patient per service line is booked, not due, overdue (1.5× the usual interval), lapsed (60 days), completed or maintenance due (TMS, inferred from 30 treatments and a follow-up), or paused / discontinued / unaccounted (ESK, from the tracker import). Lists open slots for the call list over the next 10 business days, slots entering the release window, and copies an outreach list grouped by clinic.

Late cancels, lead time, release bookings and rebuilt 3NA history need **Cancel Date/Time** and **Created Date** (Crt Dt) on the report. Cadence, funnel and recall need **Pat Name** and ideally **Md Rc**. The tool detects them by header. Thresholds are in **Setup → Settings**.

**Write report tabs** puts the results on `TA ` tabs (counts only, no patient fields). **Recall → Write patient tabs** writes `PT ` tabs that list patients by LASFIR code only (no names or record numbers).

## Development

No build step. `engine.js` holds the template analysis and `cadence.js` the cadence, recall and funnel logic (pure functions), `app.js` the UI and Excel calls, `taskpane.html` the page.

```
node --test tests/engine.test.js tests/office-store.test.js tests/cadence.test.js
```

Tests use synthetic data only.
