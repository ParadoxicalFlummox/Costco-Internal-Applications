/**
 * setup.js — First-run sheet bootstrap and onOpen menu for COMET.
 * VERSION: 0.3.5
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
 *   COMET Config    — runtime settings stored as a single JSON blob (cell A3)
 *   Active CNs      — per-employee JSON row store for proposed/active CNs
 *   (Expired CNs)   — per-employee JSON row store for expired/rejected CNs; hidden
 *
 * CN sheet initialization is delegated to getOrCreateActiveCNsSheet_() and
 * getOrCreateExpiredCNsSheet_() in cnLog.js, which apply the correct two-row
 * header layout (title row + column headers row).
 *
 * Call Log sheets are NOT created here — they are created on demand by
 * callLog.js whenever the first absence entry for a new week is logged.
 *
 * TAB COLOR CONVENTION:
 *   Blue   (#005DAA) — core data sheets (Employees)
 *   Red    (#E31837) — active manager views (Active CNs, COMET Config)
 *   Gray   (#B7B7B7) — archive/hidden sheets ((Expired CNs))
 *   Green  (#57BB8A) — Call Log sheets (created by callLog.js)
 *   White  (#FFFFFF) — attendance controller tabs (one per employee)
 */


// ---------------------------------------------------------------------------
// Configuration Defaults
// ---------------------------------------------------------------------------

/**
 * Default values for the COMET Config JSON.
 * These are merged with any stored config to ensure all keys are present.
 */
const COMET_CONFIG_DEFAULTS = {
  sendEmails: false,
  dryRun: true,
  ukgImportLastRan: null,
  trafficHeatmaps: {},
};


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
    .addSeparator()
    // Data Validation
    .addItem('Validate Employee Profiles', 'menuValidateEmployees')
    .addItem('Verify Shift Definitions', 'menuVerifyShifts')
    .addItem('Check for Duplicate IDs', 'menuCheckDuplicates')
    .addSeparator()
    // Schedule Management
    .addItem('Validate Current Week', 'menuValidateSchedule')
    .addItem('Delete & Reset Week Sheet', 'menuResetWeekSheet')
    .addSeparator()
    // Diagnostics
    .addItem('Show Employee Summary', 'menuEmployeeSummary')
    .addItem('Show Last Import', 'menuShowLastImport')
    .addItem('List All Shifts', 'menuListShifts')
    .addSeparator()
    // Cross-Department
    .addItem('Find Over-Hours Employees', 'menuFindOverHours')
    .addItem('Check Scheduling Conflicts', 'menuCheckConflicts')
    .addSeparator()
    // Maintenance
    .addItem('Recalculate Seniority Ranks', 'menuCalculateSeniority')
    .addItem('Scan for Infractions', 'menuScanInfractions')
    .addSeparator()
    // DANGER ZONE
    .addItem('Wipe Employee Sheet', 'menuWipeEmployees')
    .addItem('Export Employees', 'menuExportEmployees')
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

/** Menu handler for recalculating seniority ranks. */
function menuCalculateSeniority() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  try {
    calculateAndWriteSeniorityRanks_(); // scheduleEngine.js
    workbook.toast('Seniority ranks calculated and written to column M.', 'Success', 5);
  } catch (error) {
    console.error('setup: menuCalculateSeniority failed —', error);
    workbook.toast('Seniority calculation failed. Check Apps Script logs.', 'Error', 8);
  }
}

/** Menu handler for scanning infractions. */
function menuScanInfractions() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  try {
    // Call the infraction scanner in dry-run mode (scan without sending emails)
    const result = runCNScan(false, false); // api.js
    if (result.ok) {
      workbook.toast(
        'Infraction scan complete.\n\n' +
        'Proposals: ' + (result.data.proposals || 0) + '\n' +
        'CNs issued: ' + (result.data.issued || 0),
        'Scan Results',
        10
      );
    } else {
      workbook.toast('Infraction scan failed: ' + (result.error || 'Unknown error'), 'Error', 8);
    }
  } catch (error) {
    console.error('setup: menuScanInfractions failed —', error);
    workbook.toast('Infraction scan failed. Check Apps Script logs.', 'Error', 8);
  }
}

// --- Data Validation Handlers ---

/** Validates all employee profiles for required fields. */
function menuValidateEmployees() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const sheet = workbook.getSheetByName(EMPLOYEES_SHEET_NAME);
    if (!sheet) throw new Error('Employees sheet not found.');

    const data = sheet.getDataRange().getValues();
    const issues = [];

    for (let row = EMPLOYEES_DATA_START_ROW; row <= data.length; row++) {
      const rowData = data[row - 1];
      const name = rowData[EMPLOYEE_COLUMN.NAME - 1];
      const ftpt = rowData[EMPLOYEE_COLUMN.FT_PT - 1];
      const hireDate = rowData[EMPLOYEE_COLUMN.HIRE_DATE - 1];
      const qualifiedShifts = rowData[EMPLOYEE_COLUMN.QUALIFIED_SHIFTS - 1];

      if (!name) continue; // Skip empty rows

      if (!ftpt) issues.push(name + ' — Missing FT/PT status');
      if (!hireDate) issues.push(name + ' — Missing hire date');
      if (!qualifiedShifts) issues.push(name + ' — Missing qualified shifts');
    }

    if (issues.length === 0) {
      workbook.toast('All employee profiles valid ✓', 'Validation Complete', 5);
    } else {
      const msg = 'Found ' + issues.length + ' issues:\n\n' + issues.slice(0, 10).join('\n') + (issues.length > 10 ? '\n... and ' + (issues.length - 10) + ' more' : '');
      SpreadsheetApp.getUi().alert(msg);
    }
  } catch (error) {
    console.error('setup: menuValidateEmployees failed —', error);
    workbook.toast('Validation failed. Check logs.', 'Error', 8);
  }
}

/** Verifies all referenced shifts are defined in settings. */
function menuVerifyShifts() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const sheet = workbook.getSheetByName(EMPLOYEES_SHEET_NAME);
    if (!sheet) throw new Error('Employees sheet not found.');

    const data = sheet.getDataRange().getValues();
    const referencedShifts = new Set();
    const issues = [];

    // Collect all shift references from employee profiles
    for (let row = EMPLOYEES_DATA_START_ROW; row <= data.length; row++) {
      const rowData = data[row - 1];
      const name = rowData[EMPLOYEE_COLUMN.NAME - 1];
      const preferred = rowData[EMPLOYEE_COLUMN.PREFERRED_SHIFT - 1];
      const qualified = rowData[EMPLOYEE_COLUMN.QUALIFIED_SHIFTS - 1];

      if (!name) continue;

      if (preferred && String(preferred).trim()) {
        referencedShifts.add(String(preferred).toLowerCase().trim());
      }

      if (qualified && String(qualified).trim()) {
        String(qualified).split(',').forEach(function (s) {
          referencedShifts.add(s.toLowerCase().trim());
        });
      }
    }

    // Check each shift exists in at least one department's settings (consolidated Settings sheet)
    referencedShifts.forEach(function (shift) {
      let found = false;
      const settingsSheet = workbook.getSheetByName(SETTINGS_SHEET_NAME); // config.js
      if (settingsSheet) {
        const settingsData = settingsSheet.getDataRange().getValues();
        // Each row has [Department, JSON settings]
        for (let row = 1; row < settingsData.length; row++) { // Start at 1 to skip header
          const jsonStr = String(settingsData[row][1] || '');
          if (jsonStr && jsonStr.startsWith('{')) {
            try {
              const settings = JSON.parse(jsonStr);
              const shifts = settings.shifts || [];
              for (let s = 0; s < shifts.length; s++) {
                if (String(shifts[s].name || '').toLowerCase().trim() === shift) {
                  found = true;
                  break;
                }
              }
            } catch (e) {
              // Skip row if JSON is invalid
            }
          }
          if (found) break;
        }
      }

      if (!found && shift.length > 0) issues.push(shift);
    });

    if (issues.length === 0) {
      workbook.toast('All shifts are defined ✓', 'Verification Complete', 5);
    } else {
      SpreadsheetApp.getUi().alert('Undefined shifts found:\n\n' + issues.join('\n'));
    }
  } catch (error) {
    console.error('setup: menuVerifyShifts failed —', error);
    workbook.toast('Verification failed. Check logs.', 'Error', 8);
  }
}

/** Checks for duplicate employee IDs. */
function menuCheckDuplicates() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const sheet = workbook.getSheetByName(EMPLOYEES_SHEET_NAME);
    if (!sheet) throw new Error('Employees sheet not found.');

    const data = sheet.getDataRange().getValues();
    const ids = {};
    const duplicates = [];

    for (let row = EMPLOYEES_DATA_START_ROW; row <= data.length; row++) {
      const rowData = data[row - 1];
      const id = String(rowData[EMPLOYEE_COLUMN.ID - 1] || '').trim();
      const name = rowData[EMPLOYEE_COLUMN.NAME - 1];

      if (!id || !name) continue;

      if (ids[id]) {
        duplicates.push(id + ' (' + name + ' & ' + ids[id] + ')');
      } else {
        ids[id] = name;
      }
    }

    if (duplicates.length === 0) {
      workbook.toast('No duplicate IDs found ✓', 'Check Complete', 5);
    } else {
      SpreadsheetApp.getUi().alert('Duplicate IDs found:\n\n' + duplicates.join('\n'));
    }
  } catch (error) {
    console.error('setup: menuCheckDuplicates failed —', error);
    workbook.toast('Check failed. Check logs.', 'Error', 8);
  }
}

// --- Schedule Management Handlers ---

/** Placeholder: validates current week schedule. */
function menuValidateSchedule() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Schedule validation not yet implemented.\nCheck the web app for detailed schedule review.',
    'Info',
    5
  );
}

/** Deletes a selected week sheet so it can be regenerated fresh. */
function menuResetWeekSheet() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const sheets = workbook.getSheets();
    const weekSheets = sheets.filter(function (s) {
      return s.getName().startsWith('Week_');
    });

    if (weekSheets.length === 0) {
      workbook.toast('No week sheets found.', 'Info', 5);
      return;
    }

    const sheetNames = weekSheets.map(function (s) { return s.getName(); }).sort();
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Select a week sheet to delete:\n\n' + sheetNames.slice(0, 5).join('\n') +
      (sheetNames.length > 5 ? '\n... and ' + (sheetNames.length - 5) + ' more' : '') +
      '\n\nEnter the full sheet name in the next prompt.',
      ui.ButtonSet.OK_CANCEL
    );

    if (response !== ui.Button.OK) return;

    const sheetName = ui.prompt('Enter the exact week sheet name to delete:', '', ui.ButtonSet.OK_CANCEL);
    if (sheetName.getSelectedButton() !== ui.Button.OK) return;

    const nameToDelete = sheetName.getResponseText().trim();
    const sheetToDelete = workbook.getSheetByName(nameToDelete);

    if (!sheetToDelete) {
      workbook.toast('Sheet "' + nameToDelete + '" not found.', 'Error', 5);
      return;
    }

    const confirm = ui.alert(
      'Delete "' + nameToDelete + '"?\n\nThis cannot be undone. You will need to regenerate the schedule.',
      ui.ButtonSet.YES_NO
    );

    if (confirm === ui.Button.YES) {
      workbook.deleteSheet(sheetToDelete);
      workbook.toast('Week sheet deleted. Regenerate the schedule via the web app.', 'Success', 8);
    }
  } catch (error) {
    console.error('setup: menuResetWeekSheet failed —', error);
    workbook.toast('Reset failed. Check logs.', 'Error', 8);
  }
}

// --- Diagnostics Handlers ---

/** Shows employee count breakdown. */
function menuEmployeeSummary() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const employees = getActiveEmployees_(); // ukgImport.js
    let ft = 0, pt = 0, lpt = 0, supervisors = 0, hybrid = 0;

    employees.forEach(function (emp) {
      if (emp.role && emp.role.toLowerCase().includes('supervisor')) supervisors++;
      if (emp.ftpt === 'FT') ft++;
      else if (emp.ftpt === 'LPT') lpt++;
      else pt++;
      // Simple hybrid check: if secondary dept exists
      if (emp.secondaryDepartments && emp.secondaryDepartments.length > 0) hybrid++;
    });

    const summary = 'Active Employees: ' + employees.length + '\n\n' +
      'FT: ' + ft + '\n' +
      'PT: ' + pt + '\n' +
      'LPT: ' + lpt + '\n' +
      'Supervisors: ' + supervisors + '\n' +
      'Hybrid (multi-dept): ' + hybrid;

    SpreadsheetApp.getUi().alert(summary);
  } catch (error) {
    console.error('setup: menuEmployeeSummary failed —', error);
    workbook.toast('Summary failed. Check logs.', 'Error', 8);
  }
}

/** Shows last UKG import timestamp. */
function menuShowLastImport() {
  try {
    const config = readCometConfig_();
    const lastImport = config.ukgImportLastRan || 'Never';
    SpreadsheetApp.getUi().alert('Last UKG Import:\n' + lastImport);
  } catch (error) {
    console.error('setup: menuShowLastImport failed —', error);
    SpreadsheetApp.getActiveSpreadsheet().toast('Lookup failed. Check logs.', 'Error', 8);
  }
}

/** Placeholder: lists all defined shifts. */
function menuListShifts() {
  SpreadsheetApp.getUi().alert(
    'Shift definitions by department:\n\n' +
    'Open Settings sheet to view all department shifts.\n' +
    'Each shift shows: Name, Start Time, End Time, Paid Hours'
  );
}

// --- Cross-Department Handlers ---

/** Placeholder: finds employees over max hours. */
function menuFindOverHours() {
  SpreadsheetApp.getUi().alert(
    'Over-hours check requires active week schedule.\n\n' +
    'Use the web app: Schedule tab → view Hours column.\n' +
    'Red indicates over max hours.'
  );
}

/** Placeholder: checks for double-bookings. */
function menuCheckConflicts() {
  SpreadsheetApp.getUi().alert(
    'Schedule conflict detection:\n\n' +
    'Use the web app Schedule tab → conflicts show in red.\n' +
    'Check hybrid employees are not scheduled in overlapping shifts.'
  );
}

// --- DANGER ZONE ---

/** DESTRUCTIVE: Wipes the Employees sheet and recreates it empty. */
function menuWipeEmployees() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const confirm1 = ui.alert(
      '⚠️ WARNING: This will DELETE all employee data.\n\n' +
      'You will lose all employee profiles, preferences, and history.\n\n' +
      'This cannot be undone. Are you sure?',
      ui.ButtonSet.YES_NO
    );

    if (confirm1 !== ui.Button.YES) return;

    const confirm2 = ui.alert(
      '🔴 LAST CHANCE: Type "DELETE ALL EMPLOYEES" in the next prompt to confirm.',
      ui.ButtonSet.OK_CANCEL
    );

    if (confirm2 !== ui.Button.OK) return;

    const response = ui.prompt(
      'Type exactly: DELETE ALL EMPLOYEES',
      '',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) return;

    if (response.getResponseText().trim() !== 'DELETE ALL EMPLOYEES') {
      workbook.toast('Cancelled — text did not match.', 'Cancelled', 5);
      return;
    }

    // Delete and recreate
    const sheet = workbook.getSheetByName(EMPLOYEES_SHEET_NAME);
    if (sheet) {
      workbook.deleteSheet(sheet);
    }

    // Recreate with headers
    const newSheet = workbook.insertSheet(EMPLOYEES_SHEET_NAME, 0);
    writeEmployeesSheetHeader_(newSheet); // ukgImport.js — single source of truth for headers
    newSheet.setTabColor(SHEET_TAB_COLORS.EMPLOYEES); // config.js

    workbook.toast('Employee sheet wiped and recreated. Ready for fresh import.', 'Success', 8);
  } catch (error) {
    console.error('setup: menuWipeEmployees failed —', error);
    workbook.toast('Wipe failed. Check logs.', 'Error', 8);
  }
}

/** Exports all employee data to CSV format for backup or migration. */
function menuExportEmployees() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const employees = getActiveEmployees_(); // ukgImport.js

    if (employees.length === 0) {
      workbook.toast('No employees to export.', 'Info', 5);
      return;
    }

    // Build CSV
    const headers = ['Name', 'ID', 'Hire Date', 'Department', 'Status', 'FT/PT', 'Role', 'Secondary Depts'];
    const rows = employees.map(function (emp) {
      return [
        emp.name || '',
        emp.id || '',
        emp.hireDate || '',
        emp.department || '',
        emp.status || '',
        emp.ftpt || '',
        emp.role || '',
        (emp.secondaryDepartments || []).join('; ')
      ];
    });

    const csv = headers.join(',') + '\n' + rows.map(function (row) {
      return row.map(function (cell) {
        return '"' + String(cell).replace(/"/g, '""') + '"';
      }).join(',');
    }).join('\n');

    // Show in alert for copy-paste
    const blob = Utilities.newBlob(csv, 'text/csv', 'employees_export_' + new Date().getTime() + '.csv');
    SpreadsheetApp.getUi().alert(
      'Export ready (' + employees.length + ' employees).\n\n' +
      'File: ' + blob.getName() + '\n\n' +
      'Save this file for backup or import into another system.'
    );

    // Create a temporary sheet with the data
    const exportSheet = workbook.insertSheet('_EXPORT_TMP_');
    exportSheet.appendRow(headers);
    rows.forEach(function (row) {
      exportSheet.appendRow(row);
    });

    workbook.toast(
      'Data exported to "_EXPORT_TMP_" sheet. Copy/download as needed, then delete the sheet.',
      'Export Complete',
      10
    );
  } catch (error) {
    console.error('setup: menuExportEmployees failed —', error);
    workbook.toast('Export failed. Check logs.', 'Error', 8);
  }
}


// ---------------------------------------------------------------------------
// Config Sheet Utilities
// ---------------------------------------------------------------------------

/**
 * Reads the full COMET configuration from the config sheet (JSON cell).
 * Merges stored config with defaults to ensure all keys are present.
 *
 * @returns {Object} The full config object with all keys
 */
function readCometConfig_() {
  try {
    const workbook = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = workbook.getSheetByName(COMET_CONFIG_SHEET_NAME);
    if (!configSheet) {
      console.warn('setup: COMET Config sheet not found; returning defaults');
      return Object.assign({}, COMET_CONFIG_DEFAULTS);
    }

    // Config is stored in cell A3 as a JSON string
    const configJson = configSheet.getRange('A3').getValue();
    if (!configJson) {
      console.warn('setup: Config cell A3 is empty; returning defaults');
      return Object.assign({}, COMET_CONFIG_DEFAULTS);
    }

    const stored = JSON.parse(String(configJson));
    // Merge with defaults to ensure all keys are present
    return Object.assign({}, COMET_CONFIG_DEFAULTS, stored);
  } catch (error) {
    console.error('setup: readCometConfig_() failed — ' + error.message);
    return Object.assign({}, COMET_CONFIG_DEFAULTS);
  }
}

/**
 * Writes the full COMET configuration to the config sheet (JSON cell).
 *
 * @param {Object} config — The full config object to write
 */
function writeCometConfig_(config) {
  try {
    const workbook = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = workbook.getSheetByName(COMET_CONFIG_SHEET_NAME);
    if (!configSheet) {
      throw new Error('COMET Config sheet not found');
    }

    const configJson = JSON.stringify(config, null, 2);
    configSheet.getRange('A3').setValue(configJson);
  } catch (error) {
    console.error('setup: writeCometConfig_() failed — ' + error.message);
  }
}

/**
 * Updates part of the COMET configuration and writes it back.
 * Reads current config, merges partial update, and writes.
 *
 * @param {Object} partial — Partial config object with keys to update
 */
function updateCometConfig_(partial) {
  try {
    const config = readCometConfig_();
    Object.assign(config, partial);
    writeCometConfig_(config);
  } catch (error) {
    console.error('setup: updateCometConfig_() failed — ' + error.message);
  }
}

/**
 * Reads a single configuration value from the COMET Config sheet.
 * DEPRECATED: Use readCometConfig_() instead. Kept for backwards compatibility.
 *
 * @param {string} key — Configuration key name
 * @returns {string|null} The configuration value, or null if key not found
 */
function getConfigSheetValue_(key) {
  const config = readCometConfig_();
  const value = config[key];
  return value != null ? String(value) : null;
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
      name: EMPLOYEES_SHEET_NAME,       // config.js
      create: () => createEmployeesSheet_(workbook),
    },
    {
      name: COMET_CONFIG_SHEET_NAME,    // config.js
      create: () => createCometConfigSheet_(workbook),
    },
    {
      name: ACTIVE_CNS_SHEET_NAME,      // config.js
      create: () => getOrCreateActiveCNsSheet_(),   // cnLog.js — applies two-row header layout
    },
    {
      name: EXPIRED_CNS_SHEET_NAME,     // config.js
      create: () => getOrCreateExpiredCNsSheet_(),  // cnLog.js — applies two-row header layout + hides sheet
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

  // Install time-based triggers
  installAutoSendTrigger_();
  installDailyInfractionScanTrigger_();

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
  writeEmployeesSheetHeader_(sheet); // ukgImport.js — single source of truth for headers
  sheet.setTabColor(SHEET_TAB_COLORS.EMPLOYEES); // config.js
  return sheet;
}

/**
 * Creates and formats the COMET Config sheet.
 * Stores warehouse-level runtime settings as a JSON blob in cell A3.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} workbook
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function createCometConfigSheet_(workbook) {
  const sheet = workbook.insertSheet(COMET_CONFIG_SHEET_NAME);

  sheet.getRange('A1').setValue('COMET Configuration').setFontWeight('bold').setFontSize(13);
  sheet.getRange('A2:B2').setValues([['Config (JSON)', '(do not edit manually)']]).setFontWeight('bold')
    .setBackground(SHEET_TAB_COLORS.EMPLOYEES).setFontColor(COLORS.HEADER_TEXT); // config.js

  // Write default config as JSON to cell A3
  const configJson = JSON.stringify(COMET_CONFIG_DEFAULTS, null, 2);
  sheet.getRange('A3').setValue(configJson).setWrap(true);

  sheet.setColumnWidth(1, 600);
  sheet.setColumnWidth(2, 200);
  sheet.setFrozenRows(2);
  sheet.setTabColor(SHEET_TAB_COLORS.ACTIVE); // config.js

  return sheet;
}



// ---------------------------------------------------------------------------
// Time-Based Triggers
// ---------------------------------------------------------------------------

/**
 * Installs a 30-minute time-based trigger for auto-sending absence notifications.
 * Called from: onOpen() menu or manually via sheet menu.
 * Effect: autoSendAbsenceNotifications_() runs every 30 minutes automatically.
 */
function installAutoSendTrigger_() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const triggerId = scriptProperties.getProperty('AUTOSEND_TRIGGER_ID');

  // Check if trigger already exists
  if (triggerId) {
    try {
      ScriptApp.getProjectTrigger(triggerId);
      console.log('setup: Auto-send trigger already installed (ID: ' + triggerId + ')');
      return;
    } catch (e) {
      // Trigger doesn't exist anymore, fall through and create a new one
      scriptProperties.deleteProperty('AUTOSEND_TRIGGER_ID');
    }
  }

  // Create a new 30-minute timer trigger
  const trigger = ScriptApp.newTrigger('autoSendAbsenceNotifications_')
    .timeBased()
    .everyMinutes(30)
    .create();

  scriptProperties.setProperty('AUTOSEND_TRIGGER_ID', trigger.getUniqueId());
  console.log('setup: Auto-send trigger installed successfully (ID: ' + trigger.getUniqueId() + ')');
}

/**
 * Installs a daily trigger to scan for infractions at 2-3 AM.
 * Called from: onOpen() menu or manually via sheet menu.
 * Effect: Infraction scanner runs daily between 2-3 AM to generate fresh CNs
 * before payroll clerks arrive in the morning.
 */
function installDailyInfractionScanTrigger_() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const triggerId = scriptProperties.getProperty('INFRACTION_SCAN_TRIGGER_ID');

  // Check if trigger already exists
  if (triggerId) {
    try {
      ScriptApp.getProjectTrigger(triggerId);
      console.log('setup: Daily infraction scan trigger already installed (ID: ' + triggerId + ')');
      return;
    } catch (e) {
      // Trigger doesn't exist anymore, fall through and create a new one
      scriptProperties.deleteProperty('INFRACTION_SCAN_TRIGGER_ID');
    }
  }

  // Create a daily trigger at 2 AM (between 2-3 AM window)
  const trigger = ScriptApp.newTrigger('runCNScan')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();

  scriptProperties.setProperty('INFRACTION_SCAN_TRIGGER_ID', trigger.getUniqueId());
  console.log('setup: Daily infraction scan trigger installed successfully (ID: ' + trigger.getUniqueId() + ')');
}

/**
 * Removes the auto-send trigger if it exists.
 * Called from: Spreadsheet menu for maintenance/debugging.
 */
function uninstallAutoSendTrigger_() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const triggerId = scriptProperties.getProperty('AUTOSEND_TRIGGER_ID');

  if (!triggerId) {
    console.log('setup: No auto-send trigger found.');
    return;
  }

  try {
    const trigger = ScriptApp.getProjectTrigger(triggerId);
    ScriptApp.deleteTrigger(trigger);
    scriptProperties.deleteProperty('AUTOSEND_TRIGGER_ID');
    console.log('setup: Auto-send trigger removed successfully.');
  } catch (e) {
    console.error('setup: Failed to remove auto-send trigger — ' + e.message);
  }
}

/**
 * Removes the daily infraction scan trigger if it exists.
 * Called from: Spreadsheet menu for maintenance/debugging.
 */
function uninstallDailyInfractionScanTrigger_() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const triggerId = scriptProperties.getProperty('INFRACTION_SCAN_TRIGGER_ID');

  if (!triggerId) {
    console.log('setup: No daily infraction scan trigger found.');
    return;
  }

  try {
    const trigger = ScriptApp.getProjectTrigger(triggerId);
    ScriptApp.deleteTrigger(trigger);
    scriptProperties.deleteProperty('INFRACTION_SCAN_TRIGGER_ID');
    console.log('setup: Daily infraction scan trigger removed successfully.');
  } catch (e) {
    console.error('setup: Failed to remove daily infraction scan trigger — ' + e.message);
  }
}

