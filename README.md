# Costco Internal Applications
A collection of internal applications built to automate day to day tasks

# Development Disclaimer
This is a personal hobby project. It is **not** developed on company time, has not been approved or sponsored by Costco, and does not use any proprietary or confidential company data. No actual internal data has been extracted or used in its development. The data models, schemas and workflows were designed based on general familiarity with how data is handled in a warehouse environment and observations made while on the clock in the normal course of work.

Company management has confirmed that there is no allocated time for this work. This project is maintained entirely outside of work hours, on personal hardware.

# Included Projects
### Infraction Notifier `v1.0.0`
Scans the employee attendance controller tabs for policy violations (tardiness, no-shows, etc.) and automatically issues Counseling Notice (CN) emails to payroll recipients. Uses a deduplication engine to avoid re-notifying on previously logged infractions. Supports dry-run mode for safe testing without sending emails.

### Absence Notifier `v0.2.1`
A time-windowed notification system that monitors an absence log spreadsheet and emails department managers about recent employee call-outs. Runs on a configurable rolling time window so managers are only alerted about new entries since the last check.

### Review Notifier `TBD`
Planned refactor of a tool from my last warehouse to notify managers and payroll of upcoming performance reviews.

### Auto Schedule Generator `v1.0.0`
A four-phase weekly schedule generation engine for warehouse departments. Ingests an employee roster with shift preferences and availability, then produces a complete schedule grid that respects seniority, preferred days off, minimum/maximum hour rules and coverage requirements. Outputs formatted schedule sheets directly to Google Sheets.

### COMET: Costco Operations, Management & Employee Tracking `v0.6.x`
A unified warehouse management web application built on Google Apps Script. Consolidates scheduling, attendance tracking, and infraction management into a single web based interface backed by Google Sheets. Includes an upgraded version of the Schedule Generator with multi-department generation and door count heatmap based staffing, a full CN (Counseling Notice) lifecycle pipeline, employee call out logging with FMLA tracking, and a UKG payroll data import workflow.

---

Future projects will be added here

---

## Deploying to Google App Script

All projects in this repo are Google App Script (GAS) applications backed by Google Sheets.

### Prerequisites
- A google workspace account with access to Google Sheets and Google Drive
- For pushing changes to COMET only: [clasp](https://github.com/google/clasp) installed globally (`npm install -g @google/clasp`) and authenticated (`clasp login`)

> **Note on file extensions:** All scripts in this repo use the `.js` extension rather than `.gs`. This is intentional — the projects were developed in VSCodium, which does not recognize `.gs` as a known format and provides no autocomplete or IntelliCode support for it. JavaScript and Google Apps Script are functionally the same language; GAS adds a set of Google-specific global APIs (like `SpreadsheetApp`, `GmailApp`, etc.) on top of standard JavaScript. Renaming any `.js` file to `.gs` before pasting into the GAS editor is not required — the editor accepts either extension.

### Infraction Notifier, Absence Notifier, Auto Schedule Generator

These projects are deployed manually through the GAS online editor:

1. Open (or create) the Google Sheet you want to attach the script to.
2. Go to **Extensions → Apps Script** to open the bound script editor.
3. Delete any default code in the editor.
4. For each `.js` file in the project folder, create a matching script file in the editor and paste the contents in.
5. Open `config.js` (in the editor) and fill in all configuration values for your spreadsheet (column mappings, recipient email addresses, sheet names, etc.).
6. Save all files, then run the entry-point function manually once to grant the required OAuth permissions.
7. Optionally set up a time-based trigger under **Triggers** (clock icon in the sidebar) to run the script on a schedule.

### COMET

COMET is a web app deployed via clasp:

1. Create a new GAS project at [script.google.com](https://script.google.com) (or use an existing one).
2. Copy the script ID from the project URL (`/projects/<scriptId>/edit`).
3. Update `comet/.clasp.json` with your script ID:
   ```json
   { "scriptId": "YOUR_SCRIPT_ID", "rootDir": "." }
   ```
4. From the `comet/` directory, push the files:
   ```bash
   cd comet/
   clasp push
   ```
5. In the GAS editor, go to **Deploy → New Deployment**, select type **Web app**, and configure access as needed.
6. Open `config.js` and set all configuration values (spreadsheet IDs, recipient emails, column positions, etc.) before going live.
7. Ensure `DRY_RUN` is set to `false` in `config.js` when deploying to production.