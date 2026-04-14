/**
 * ui.js — Custom menu and menu handler wrappers for the Infraction Notifier.
 * VERSION: 0.1.1
 *
 * This file registers the "Infraction Notifier" menu when the attendance
 * controller workbook is opened, and provides thin handler functions for each
 * menu item. Each handler wraps its worker function in a try/catch and shows
 * a toast notification so the manager always gets visible feedback.
 *
 * This file is the single onOpen registration point for ALL menus in the
 * attendance controller workbook. Having two separate onOpen functions would
 * cause only the last-defined one to run; all menus must be registered here.
 *
 * INFRACTION NOTIFIER MENU:
 *   Dry Run (Log Only)     → dryRunCNs()             — scans, logs, no emails sent
 *   Send CNs (Live)        → sendCNsDaily()           — scans and sends real emails
 *   Run Expiry Check       → expireCNsDaily()         — marks expired CNs, notifies payroll
 *   ──────────────
 *   Debug: Active Sheet    → menuDebugActiveSheet()   — logs all parsed events for the
 *                            currently selected employee tab; verify layout before live run
 *   Setup Config Sheet     → menuSetupConfigSheet()   — creates "Infraction Config" sheet
 *
 * CREATE TABS IN SHEET MENU (tabManager.js):
 *   Create Tabs W Color    → buildtabs()              — creates tabs from color template
 *   Create Tabs W/O Color  → buildtabs2()             — creates tabs from plain template
 *
 * CREATE INDIVIDUAL SHEETS MENU (tabManager.js):
 *   Create Sheets W Color  → buildSheetColor()        — creates separate files, color template
 *   Create Sheets W/O Color → buildSheetNoColor()     — creates separate files, plain template
 *
 * SORT TABS MENU (tabManager.js):
 *   Sort Tabs              → sortTabs()               — sorts all tabs alphabetically
 *
 * TRIGGER SETUP REMINDER:
 *   The daily automatic CN scan requires a time-driven trigger pointing to
 *   sendCNsDaily. Menu-based sends are available for immediate manual runs.
 */


// ---------------------------------------------------------------------------
// onOpen — Menu Registration
// ---------------------------------------------------------------------------

/**
 * Registers the "Infraction Notifier" custom menu when the workbook is opened.
 *
 * This is a simple trigger (not installable) so it fires automatically whenever
 * any user opens the spreadsheet. Menu items call the handler functions below.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // --- Infraction Notifier ---
  ui.createMenu('Infraction Notifier')
    .addItem('Dry Run (Log Only)',  'menuDryRun')
    .addItem('Send CNs (Live)',     'menuSendLive')
    .addSeparator()
    .addItem('Run Expiry Check',   'menuExpireCheck')
    .addSeparator()
    .addItem('Debug: Active Sheet', 'menuDebugActiveSheet')
    .addItem('Setup Config Sheet',  'menuSetupConfigSheet')
    .addToUi();

  // --- Create Tabs In Sheet (tabManager.js) ---
  ui.createMenu('Create Tabs In Sheet')
    .addItem('Create Tabs W Color',   'buildtabs')
    .addItem('Create Tabs W/O Color', 'buildtabs2')
    .addToUi();

  // --- Create Individual Sheets (tabManager.js) ---
  ui.createMenu('Create Individual Sheets')
    .addItem('Create Sheets W Color',   'buildSheetColor')
    .addItem('Create Sheets W/O Color', 'buildSheetNoColor')
    .addToUi();

  // --- Sort Tabs (tabManager.js) ---
  ui.createMenu('Sort Tabs')
    .addItem('Sort Tabs', 'sortTabs')
    .addToUi();
}


// ---------------------------------------------------------------------------
// Menu Handlers
// ---------------------------------------------------------------------------

/**
 * Runs the infraction scanner in dry-run mode.
 * No emails are sent and nothing is written to the CN_Log.
 * Results are visible in Apps Script → View → Logs.
 */
function menuDryRun() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  try {
    workbook.toast('Scanning for infractions (dry run — no emails will be sent)…', 'Infraction Notifier', 5);
    dryRunCNs(); // infractionEngine.js
    workbook.toast('Dry run complete. Open Apps Script logs to see results.', 'Done', 6);
  } catch (error) {
    console.error('ui: menuDryRun failed —', error);
    workbook.toast('Dry run failed. Check Apps Script logs for details.', 'Error', 8);
  }
}

/**
 * Runs the infraction scanner in live mode.
 * Emails are sent and new CNs are written to the CN_Log.
 * Only available when DRY_RUN is false in config.js — shows a warning if not.
 */
function menuSendLive() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  try {
    if (DRY_RUN) { // config.js
      SpreadsheetApp.getUi().alert(
        'Live send is disabled.\n\n' +
        'Set DRY_RUN = false in config.js to enable live email sending.'
      );
      return;
    }
    workbook.toast('Scanning and sending CN notifications…', 'Infraction Notifier', 5);
    sendCNsDaily(); // infractionEngine.js
    workbook.toast('Scan complete. Check CN_Log for issued CNs.', 'Done', 6);
  } catch (error) {
    console.error('ui: menuSendLive failed —', error);
    workbook.toast('Send failed. Check Apps Script logs for details.', 'Error', 8);
  }
}

/**
 * Runs the CN expiry check.
 * Marks Active CNs older than EXPIRY_DAYS as Expired in the CN_Log and
 * sends expiry notifications to payroll.
 */
function menuExpireCheck() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  try {
    workbook.toast('Checking for expired CNs…', 'Infraction Notifier', 5);
    expireCNsDaily(!!DRY_RUN); // cnLog.js — passes current dry-run mode
    workbook.toast('Expiry check complete. Check Apps Script logs for details.', 'Done', 6);
  } catch (error) {
    console.error('ui: menuExpireCheck failed —', error);
    workbook.toast('Expiry check failed. Check Apps Script logs for details.', 'Error', 8);
  }
}

/**
 * Parses the currently active sheet tab and logs all detected calendar events.
 *
 * Useful for verifying that the grid layout constants in config.js match the
 * actual attendance controller layout before running the full scan. Run this
 * on a known employee tab and check View → Logs to confirm events are parsed
 * correctly with the right dates and codes.
 */
function menuDebugActiveSheet() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const sheet    = workbook.getActiveSheet();
    const timeZone = Session.getScriptTimeZone();
    const ctx      = readEmployeeContext_(sheet); // calendarParser.js
    const year     = parseYearFromTitle_(ctx.yearTitle) || new Date().getFullYear();
    const events   = parseCalendarEvents_(sheet, year, timeZone, ctx); // calendarParser.js

    events.sort((a, b) => a.date.getTime() - b.date.getTime());

    console.log(`DEBUG: Sheet "${sheet.getName()}" — Employee: ${ctx.employeeName} | ID: ${ctx.employeeId} | Dept: ${ctx.department} | Year: ${year}`);
    console.log(`DEBUG: ${events.length} total event(s) parsed.`);

    events.forEach(e => {
      const dateStr = Utilities.formatDate(e.date, timeZone, 'yyyy-MM-dd');
      const status  = e.isIgnored ? 'IGNORED' : (e.isInfraction ? 'INFRACTION' : 'other');
      console.log(`  ${dateStr}  ${e.code.padEnd(4)}  ${status.padEnd(10)}  cell ${e.a1}`);
    });

    workbook.toast(
      `Parsed ${events.length} event(s) for ${ctx.employeeName || sheet.getName()}. Open Apps Script logs to review.`,
      'Debug Complete',
      8
    );
  } catch (error) {
    console.error('ui: menuDebugActiveSheet failed —', error);
    workbook.toast('Debug parse failed. Check Apps Script logs for details.', 'Error', 8);
  }
}

/**
 * Creates (or resets) the "Infraction Config" sheet in the active workbook.
 *
 * The config sheet has a single input cell (B2) where a manager pastes the
 * Google Spreadsheet ID of the dedicated CN Log workbook. Once set, all CN
 * records are written there instead of the attendance controller.
 *
 * GUARD: If the sheet already exists and B2 contains a non-empty value,
 * setup is skipped and the manager is shown the current value.
 */
function menuSetupConfigSheet() {
  const workbook    = SpreadsheetApp.getActiveSpreadsheet();
  const ui          = SpreadsheetApp.getUi();

  try {
    let configSheet = workbook.getSheetByName(INFRACTION_CONFIG_SHEET_NAME); // config.js

    if (configSheet) {
      const existingId = configSheet.getRange(LOG_SPREADSHEET_ID_CELL).getValue().toString().trim(); // config.js
      if (existingId) {
        ui.alert(
          `"${INFRACTION_CONFIG_SHEET_NAME}" is already set up.\n\n` +
          `Current CN Log Spreadsheet ID:\n${existingId}\n\n` +
          `To change it, edit cell ${LOG_SPREADSHEET_ID_CELL} directly.`
        );
        return;
      }
      // Sheet exists but B2 is blank — reset the layout
      configSheet.clearContents();
      configSheet.clearFormats();
    } else {
      configSheet = workbook.insertSheet(INFRACTION_CONFIG_SHEET_NAME);
    }

    writeConfigSheetLayout_(configSheet);
    workbook.setActiveSheet(configSheet);
    workbook.toast('Config sheet ready. Paste your CN Log spreadsheet ID into cell B2.', 'Setup Complete', 8);
  } catch (error) {
    console.error('ui: menuSetupConfigSheet failed —', error);
    workbook.toast('Setup failed. Check Apps Script logs for details.', 'Error', 8);
  }
}


// ---------------------------------------------------------------------------
// Config Sheet Layout
// ---------------------------------------------------------------------------

/**
 * Writes the labels and input cell to the Infraction Config sheet.
 *
 * Layout:
 *   A1 — Bold title: "Infraction Notifier — Configuration"
 *   A2 — Label: "CN Log Spreadsheet ID:"
 *   B2 — Input cell (blank; manager pastes the ID here)
 *   A4 — Instructional note
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} configSheet
 */
function writeConfigSheetLayout_(configSheet) {
  configSheet.getRange('A1')
    .setValue('Infraction Notifier — Configuration')
    .setFontWeight('bold')
    .setFontSize(13);

  configSheet.getRange('A2')
    .setValue('CN Log Spreadsheet ID:')
    .setFontWeight('bold');

  configSheet.getRange(LOG_SPREADSHEET_ID_CELL) // "B2" — config.js
    .setValue('')
    .setNote(
      'Paste the Google Spreadsheet ID of your dedicated CN Log workbook.\n\n' +
      'Find it in the spreadsheet URL:\n' +
      'docs.google.com/spreadsheets/d/ *** PASTE THIS PART *** /edit\n\n' +
      'Once set, all Counseling Notice records will be written to that workbook\n' +
      'instead of the attendance controller.\n\n' +
      'Leave blank to write the CN_Log into this workbook instead.'
    );

  configSheet.getRange('A4')
    .setValue(
      'Paste the CN Log Spreadsheet ID in cell B2. ' +
      'If left blank, CN records will be written into a CN_Log tab in this workbook. ' +
      'The external workbook option is recommended since the attendance controller ' +
      'is a corporate source of truth that should not be modified.'
    )
    .setFontStyle('italic')
    .setFontColor('#666666')
    .setWrap(true);

  configSheet.setColumnWidth(1, 200); // A — label
  configSheet.setColumnWidth(2, 250); // B — input
  configSheet.setColumnWidth(3, 420); // C — overflow/notes
}
