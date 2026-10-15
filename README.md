# Clinic Template Tool

Excel add-in for Practice Managers. Put copies of provider templates in a workbook, add the NextGen appointment list, and the tool measures how the templates actually get used.

- App: https://brendannyates.github.io/clinic-template-tool/taskpane.html (open it outside Excel to download the manifest)
- Runs in Excel on the web. Data stays in the workbook; this site only serves the app's code.

## What goes in the workbook

| Tab | What it holds |
|---|---|
| Template tabs (any name) | One weekly grid per provider in the provider-configurable layout: `Provider:`, `Location:`, `Effective From:` labels, a `Time` column, Monday–Friday headers, a slot type in each cell. Several grids can share a tab. **Setup → New template tab** makes one; copy an old one to record a revision. |
| Appointment List | The NextGen *Appointments Listing* report. **Setup → Import NextGen report** writes it. |
| Slot Types | Slot type → Clinical or Blocked, and what it accepts (event class, service line or event; `,` = any of, `+` = all of, `*` = anything). |
| Ref | NextGen event → event class, service line, whether it counts (holds and reference rows don't). |
| Providers | NextGen provider → template label, credential, include. Chairs and rooms are left out automatically. |

Hidden tabs are never read.

## What it measures

- **Template fill and fit** by provider, slot type, day of week, hour and AM/PM: slots filled, kept use, kept visits that matched the slot type (and what filled them when they didn't), open slots, slots taken by admin, PTO.
- **Visit outcomes** by service line, event class, appointment type, day and hour: no-show, cancel and late-cancel rates, median lead time.
- **Access**: third next available (3NA, calendar days) by slot type, for each provider and all providers, plus booked-ahead % by week. **Save 3NA snapshot** keeps a weekly history on the `TA 3NA History` tab.
- **Provider impact**: visits in admin/lunch/meeting time, outside template hours or days, double-booked, extra hours.
- **Trends**: weekly (Sunday–Saturday) fill, fit, no-show, cancel and kept volume.
- **ESK access model**: imports the ESK Tracker's *Export for template tool* file (`mhs-access-model` v1) and compares panel need with the ESK slots in the template. IV ketamine will use the same format.

Late cancels, lead time and rebuilt 3NA history need two extra report columns: **Cancel Date/Time** and **Created Date**. The tool detects them by header.

**Write report tabs** puts the results on `TA ` tabs (counts only, no patient fields).

## Development

No build step. `engine.js` holds the analysis (pure functions), `app.js` the UI and Excel calls, `taskpane.html` the page.

```
node --test tests/engine.test.js tests/office-store.test.js
```

Tests use synthetic data only.
