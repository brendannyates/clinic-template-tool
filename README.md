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

Hidden tabs are never read, except a Clinic Eligibility file's own data tabs (below).

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

## Clinic Eligibility file (from central Sohar)

Central Sohar Eligibility does every Sohar import, export, discovery and verification. For each clinic the admin saves a **Clinic Eligibility - <Location> YYYY-MM-DD.xlsx** file (Admin tab, **Clinic Template Tool file**) and puts it in that clinic's folder in the Clinic Eligibility library (SharePoint, one folder per clinic). Opening it opens this add-in.

- **What's in it**: only the patients whose next visit is at that clinic, their visits, Sohar results and benefits, the admin's build review for them, and what Sohar has learned about the payers they carry. Patient and guarantor names and dates of birth are taken out before the file is made; each patient has a **LASFIR** code (first three letters of the last name, then of the first name, e.g. Jane Smith = SMIJAN) and the MRN. The pane shows patients as LASFIR + MRN only. The file still holds PHI (MRNs, member IDs, insurance, visits): keep it in MHS OneDrive or Teams.
- **Tabs**: `Read Me` and `Clinic Work` are visible; `SE Info`, `SE Patients`, `SE Results`, `SE Review`, `SE Payers`, `SE Done` and `Payer Credential Rules` are hidden data tabs. Don't edit those. The add-in reads them with Sohar's own `core.js` and `rules.js`, so the list, flags, benefits and tasks match what the admin sees.
- **What the clinic sees**: **Coverage → Eligibility** (one card per patient by visit day: pills, payer, what to choose in NextGen, Sohar status, copay, deductible and OOP meters, Action/Check items, tasks to tick, Re-verify anyway, a note), **Benefits** (deductible and OOP max met or approaching) and **Payers** (how far each payer's mapping can be trusted, what Sohar found, a payer note). The template tools stay hidden until **Setup → Show template tools** is ticked.
- **Done notes**: each task has a **Note** button next to its tick ("called the patient, new card scanned"). The note is saved on the task's `Task done` row in `Clinic Work`, and the admin sees it in the `Clinic Note` column of NextGen Tasks. You can save a note before ticking the box, and ticking or unticking keeps the note. Clearing a note takes it off the file but doesn't erase one the admin already has. When the admin's NextGen Tasks already had a `Clinic Note` for the task as the file was made (yours brought back, or one the admin typed), it comes in the hidden `SE Done` tab (`Task Key`, `Done`, `Clinic Note`; older files without the column still open) and shows under the task as "The admin has this note: …". You can still add a new note; it shows beside the admin's and replaces it once it reaches the admin. Sohar takes patient names out of that note before the file is made.
- **Getting changes to the admin**: ticks, re-verify requests and notes are saved on the `Clinic Work` tab (Kind, Key, MRN, LASFIR, Value, Note, By, When; the last row per Kind + Key wins). Two people can work the same file: each change is added to the rows on the tab at the moment it saves, not to the copy your pane read when it opened, so nobody's changes are wiped out, and the pane then shows the other person's rows too. When both change the same thing (same Kind + Key), the last save wins, and the admin gets the same winner (if a PC's clock is behind, the add-in moves that save's When just past the earlier one). Save the file in your clinic's folder in the Clinic Eligibility library. Excel on the web saves on its own, and the admin's flow copies the `Clinic Work` rows into Sohar's clinic inbox from there, so you don't need to send the file back. If the file isn't in the library, or the admin asks for it, save it and send it through OneDrive or Teams; the admin brings it into Sohar (**Bring back clinic files**). Payer notes are for a person to act on; they never change the Payer Map by themselves.
- The flow reads only the rows inside the `tblClinicWork` table, so every save makes the table cover all its rows. In a file saved by an earlier version, rows left below the table are pulled in by the next save. If the table is missing or can't grow, the pane says so once and asks you to send the file instead.
- The file's data is from its run date. Visits that have passed drop out, and re-verify windows are counted from today. Ask the admin for a new file each week.

Workbooks without an `SE Info` tab work exactly as before.

## Development

No build step. `engine.js` holds the template analysis and `cadence.js` the cadence, recall and funnel logic (pure functions), `app.js` the UI and Excel calls, `taskpane.html` the page. `core.js`, `rules.js` and `payer-rules.js` are byte-identical copies of the Sohar Eligibility add-in's files (see CLAUDE.md).

```
npm test
```

(or `node --test tests/*.js`)

Tests use synthetic data only.
