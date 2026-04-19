/**
 * setup.js — First-run sheet bootstrap and onOpen menu for COMET.
 * VERSION: 0.1.0
 *
 * This file has two responsibilities:
 *
 *   1. FIRST-RUN SETUP — runCometSetup_() creates every sheet COMET needs
 *      if it doesn't already exist. It is safe to run multiple times; sheets
 *      that already exist are left untouched. Called from:
 *        - The Admin panel "Run Setup" button via api.js runSetup()
 *        - The "COMET" spreadsheet menu (backup for the GM)
 *
 *   2. SPREADSHEET MENU — onOpen() registers a minimal "COMET" menu so the
 *      GM has a fallback if the web app URL is unavailable. Menu items are
 *      intentionally limited — all real work happens through the web app.
 *
 * SHEETS CREATED:
 *   Employees       — master employee roster (UKG import target)
 *   COMET Config    — runtime settings (FY start, window minutes, etc.)
 *   CN_Log          — infraction records written by the infraction scanner
 *   Active CNs      — manager-facing view of currently active CNs
 *   (Expired CNs)   — archived CNs moved here on expiry; hidden from tab bar
 *
 * Call Log sheets are NOT created here — they are created on demand by
 * callLog.js whenever the first absence entry for a new week is logged.
 *
 * TAB COLOR CONVENTION:
 *   Blue   (#005DAA) — core data sheets (Employees, CN_Log)
 *   Red    (#E31837) — active manager views (Active CNs, COMET Config)
 *   Gray   (#B7B7B7) — archive/hidden sheets ((Expired CNs))
 *   Green  (#57BB8A) — Call Log sheets (created by callLog.js)
 *   White  (#FFFFFF) — attendance controller tabs (one per employee)
 */


// ---------------------------------------------------------------------------
// Spreadsheet Menu
// ---------------------------------------------------------------------------

/**
 * Registers the COMET spreadsheet menu when the workbook is opened.
 *
 * This is a simple trigger that fires whenever any user opens the spreadsheet.
 * It provides a fallback for the GM if the web app URL is unavailable, and
 * a one-click way to run setup on a fresh workbook.
 */
function onOpen() { // eslint-disable-line no-unused-vars
  SpreadsheetApp.getUi()
    .createMenu('COMET')
    .addItem('Run Setup', 'menuRunSetup')
    .addSeparator()
    .addItem('Open Web App', 'menuOpenWebApp')
    .addToUi();
}

/** Menu handler for Run Setup. */
function menuRunSetup() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const result = runCometSetup_();
    workbook.toast(
      `Setup complete — ${result.created.length > 0 ? result.created.join(', ') + ' created' : 'all sheets already exist'}.`,
      'COMET Setup',
      8
    );
  } catch (error) {
    console.error('setup: menuRunSetup failed —', error);
    workbook.toast('Setup failed. Check Apps Script logs for details.', 'Error', 8);
  }
}

/**
 * Shows the web app URL in an alert so the GM can copy it.
 * The URL is only available after at least one deployment exists.
 */
function menuOpenWebApp() {
  const url = ScriptApp.getService().getUrl();
  if (url) {
    SpreadsheetApp.getUi().alert('COMET Web App URL:\n\n' + url);
  } else {
    SpreadsheetApp.getUi().alert(
      'No web app deployment found.\n\n' +
      'Deploy COMET first:\n' +
      'Apps Script editor → Deploy → New Deployment → Web App'
    );
  }
}


// ---------------------------------------------------------------------------
// First-Run Setup
// ---------------------------------------------------------------------------

/**
 * Creates all sheets COMET needs if they do not already exist.
 *
 * Safe to run multiple times — existing sheets are left completely untouched.
 * Returns a summary of what was created so the caller can report to the user.
 *
 * @returns {{ created: string[], skipped: string[] }}
 */
function runCometSetup_() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];
  const skipped = [];

  const sheets = [
    {
      name:    EMPLOYEES_SHEET_NAME,       // config.js
      create:  () => createEmployeesSheet_(workbook),
    },
    {
      name:    COMET_CONFIG_SHEET_NAME,    // config.js
      create:  () => createCometConfigSheet_(workbook),
    },
    {
      name:    CN_LOG_SHEET_NAME,          // config.js
      create:  () => createCnLogSheet_(workbook),
    },
    {
      name:    ACTIVE_CNS_SHEET_NAME,      // config.js
      create:  () => createActiveCnsSheet_(workbook),
    },
    {
      name:    EXPIRED_CNS_SHEET_NAME,     // config.js
      create:  () => createExpiredCnsSheet_(workbook),
    },
  ];

  sheets.forEach(({ name, create }) => {
    if (workbook.getSheetByName(name)) {
      skipped.push(name);
    } else {
      create();
      created.push(name);
    }
  });

  SpreadsheetApp.flush();
  console.log(`setup: runCometSetup_ complete — created: [${created.join(', ')}], skipped: [${skipped.join(', ')}]`);
  return { created, skipped };
}


// ---------------------------------------------------------------------------
// Individual Sheet Creators
// ---------------------------------------------------------------------------

/**
 * Creates and formats the Employees sheet.
 * Called by runCometSetup_() and also by ukgImport.js getOrCreateEmployeesSheet_()
 * as a fallback — both guard against duplicate creation, so there is no conflict.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} workbook
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function createEmployeesSheet_(workbook) {
  const sheet = workbook.insertSheet(EMPLOYEES_SHEET_NAME);
  const headers = ['Name (Last, First)', 'Employee ID', 'Hire Date', 'Department', 'Status'];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#005DAA')
    .setFontColor('#FFFFFF');

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 90);
  sheet.setFrozenRows(1);
  sheet.setTabColor('#005DAA');

  return sheet;
}

/**
 * Creates and formats the COMET Config sheet.
 * Stores warehouse-level runtime settings as key/value pairs.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} workbook
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function createCometConfigSheet_(workbook) {
  const sheet = workbook.insertSheet(COMET_CONFIG_SHEET_NAME);

  sheet.getRange('A1').setValue('COMET Configuration').setFontWeight('bold').setFontSize(13);
  sheet.getRange('A2:B2').setValues([['Setting', 'Value']]).setFontWeight('bold')
    .setBackground('#005DAA').setFontColor('#FFFFFF');

  // Default config values
  const defaults = [
    ['windowMinutes',  '15'],
    ['fyStartMonth',   '9'],   // September (Costco fiscal year)
    ['dryRun',         'true'],
  ];
  sheet.getRange(3, 1, defaults.length, 2).setValues(defaults);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 200);
  sheet.setFrozenRows(2);
  sheet.setTabColor('#E31837');

  return sheet;
}

/**
 * Creates and formats the CN_Log sheet with the correct headers.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} workbook
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function createCnLogSheet_(workbook) {
  const sheet = workbook.insertSheet(CN_LOG_SHEET_NAME);

  sheet.getRange(1, 1, 1, CN_LOG_HEADERS.length)  // config.js
    .setValues([CN_LOG_HEADERS])
    .setFontWeight('bold')
    .setBackground('#005DAA')
    .setFontColor('#FFFFFF');

  sheet.setFrozenRows(1);
  sheet.setTabColor('#005DAA');

  return sheet;
}

/**
 * Creates and formats the Active CNs sheet with the correct headers.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} workbook
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function createActiveCnsSheet_(workbook) {
  const sheet = workbook.insertSheet(ACTIVE_CNS_SHEET_NAME);

  sheet.getRange(1, 1, 1, ACTIVE_CNS_HEADERS.length)  // config.js
    .setValues([ACTIVE_CNS_HEADERS])
    .setFontWeight('bold')
    .setBackground('#E31837')
    .setFontColor('#FFFFFF');

  sheet.setFrozenRows(1);
  sheet.setTabColor('#E31837');

  return sheet;
}

/**
 * Creates, formats, and hides the (Expired CNs) archive sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} workbook
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function createExpiredCnsSheet_(workbook) {
  const sheet = workbook.insertSheet(EXPIRED_CNS_SHEET_NAME);

  sheet.getRange(1, 1, 1, EXPIRED_CNS_HEADERS.length)  // config.js
    .setValues([EXPIRED_CNS_HEADERS])
    .setFontWeight('bold')
    .setBackground('#B7B7B7')
    .setFontColor('#FFFFFF');

  sheet.setFrozenRows(1);
  sheet.setTabColor('#B7B7B7');
  sheet.hideSheet();

  return sheet;
}
