/**
 * ui.js — Apps Script triggers, custom menu, and menu handler functions.
 * VERSION: 0.2.0
 *
 * This file is the entry point for all user-initiated actions in the absence
 * notifier workbook. It contains:
 *   - onOpen():         Creates the "Call Log Admin" menu when the workbook opens.
 *   - Menu handlers:    Thin functions triggered by menu items that call workers
 *                       in other files and show toast confirmations.
 *
 * DESIGN RULE:
 *   Every function here is either a pure trigger (onOpen) or a thin wrapper
 *   that calls a worker function from another file and reports the result via
 *   toast or alert. No sheet logic, no email logic, and no time math lives here.
 *   If something goes wrong during a menu action, the trace leads to the worker
 *   file, not this one.
 *
 * GAS TRIGGER NOTES:
 *   - onOpen() is a simple trigger — it runs automatically on open with limited
 *     permissions. It cannot access external services or send email.
 *   - Menu handler functions run with the user's full permissions and can call
 *     GmailApp, SpreadsheetApp, etc.
 *   - The onEdit() installable trigger lives in autofill.js, not here, because
 *     it requires elevated permissions that simple triggers do not provide.
 */


// ---------------------------------------------------------------------------
// GAS Simple Trigger: onOpen
// ---------------------------------------------------------------------------

/**
 * Creates the "Call Log Admin" custom menu in the Google Sheets menu bar.
 *
 * Runs automatically every time the workbook is opened via the simple trigger
 * naming convention — no manual trigger installation required.
 *
 * Menu layout:
 *   Generate New Week Sheet   — Creates a formatted call log sheet for the current period/week
 *   Send Digest Now           — Manually sends the digest for the full current window
 *   ── separator ──
 *   Setup Config Sheet        — First-run setup; creates the Absence Config input sheet
 */
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet()
    .addMenu('Call Log Admin', [
      { name: 'Generate New Week Sheet', functionName: 'menuGenerateNewSheet'  },
      { name: 'Send Digest Now',         functionName: 'menuSendDigestNow'     },
      null, // Separator
      { name: 'Setup Config Sheet (First Run)', functionName: 'menuSetupConfigSheet' },
    ]);
}


// ---------------------------------------------------------------------------
// Menu Handlers
// ---------------------------------------------------------------------------

/**
 * Triggered by: Call Log Admin → Generate New Week Sheet
 *
 * Creates a new, formatted call log sheet titled with the current fiscal
 * period and week (e.g. "P3 W1"), or the "Week Ending MM/DD/YY" fallback
 * if no FY start date has been configured.
 *
 * Shows a success toast with the new sheet's title, or an alert if the sheet
 * already exists so the manager knows nothing was overwritten.
 *
 * The actual creation logic lives in sheetGenerator.js:generateNewCallLogSheet().
 * This function only calls it and handles the user-facing feedback.
 */
function menuGenerateNewSheet() {
  try {
    const sheetTitle = getActiveCallLogSheetName_(); // defined in sheetUtils.js

    // generateNewCallLogSheet() shows its own alert if the sheet already exists
    // and returns early — we check existence here to know whether to show a
    // success toast or let the alert from the generator serve as the response.
    const workbook = SpreadsheetApp.getActiveSpreadsheet();
    if (workbook.getSheetByName(sheetTitle)) {
      // generateNewCallLogSheet() will show the "already exists" alert — just call it
      generateNewCallLogSheet();
      return;
    }

    generateNewCallLogSheet(); // defined in sheetGenerator.js

    workbook.toast(
      `"${sheetTitle}" created and ready for entries.`,
      'Sheet Generated',
      6
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Error generating sheet:\n\n' + error.message
    );
    console.error('menuGenerateNewSheet error:', error);
  }
}

/**
 * Triggered by: Call Log Admin → Send Digest Now
 *
 * Manually runs the absence digest for the current time window without waiting
 * for the next scheduled trigger fire. Useful when a manager wants to push a
 * notification immediately after logging a late entry, or to test that the
 * notifier is configured correctly.
 *
 * This calls the same sendAbsenceDigest() function that the time-driven trigger
 * calls — there is no special "manual mode." The window used is the most
 * recently completed N-minute window, identical to what the automatic trigger
 * would use if it fired right now.
 *
 * Shows a toast to confirm the action was initiated. Email delivery (or the
 * absence of emails if no records matched) is reported to the console log.
 *
 * The digest logic lives in absenceNotifier.js:sendAbsenceDigest().
 */
function menuSendDigestNow() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Scanning the current window for absences...',
      'Sending Digest',
      4
    );

    sendAbsenceDigest(); // defined in absenceNotifier.js

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Digest run complete. Check the Apps Script logs for details.',
      'Done',
      6
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Error sending digest:\n\n' + error.message
    );
    console.error('menuSendDigestNow error:', error);
  }
}

/**
 * Triggered by: Call Log Admin → Setup Config Sheet (First Run)
 *
 * Creates (or re-initializes) the "Absence Config" sheet with an input cell
 * for the fiscal year start date. This is a one-time step when first deploying
 * the notifier. After setup, the manager fills in the FY P1W1 start date and
 * the sheet name calculation will produce "P# W#" labels automatically.
 *
 * Shows a toast on success. If the sheet is already configured with a real
 * date, setupAbsenceConfigSheet() shows an alert explaining the current value
 * rather than overwriting it.
 *
 * The setup logic lives in sheetGenerator.js:setupAbsenceConfigSheet().
 */
function menuSetupConfigSheet() {
  try {
    setupAbsenceConfigSheet(); // defined in sheetGenerator.js

    // setupAbsenceConfigSheet() shows its own alert if already configured,
    // so only show the success toast when we reach this point without an alert.
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `"${CONFIG_SHEET_NAME}" is ready. Enter the FY P1W1 start date in cell B2.`,
      'Setup Complete',
      8
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Error during setup:\n\n' + error.message
    );
    console.error('menuSetupConfigSheet error:', error);
  }
}
