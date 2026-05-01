/**
 * tabManager.js — Attendance controller JSON tab management.
 * VERSION: 0.4.2
 *
 * Manages per-employee attendance tabs in JSON format (Phase 2).
 * Each employee tab contains one row per year with attendance codes stored as JSON.
 *
 * Layout (per employee tab):
 *   Row 1:      [Year | Codes JSON | StoredAt]  (header)
 *   Row 2:      [2024 | {...codes...} | timestamp]
 *   Row 3:      [2025 | {...codes...} | timestamp]
 *   ...
 *
 * Replaces the old template-copy approach (generateAttendanceControllerTabs_) and
 * grid formatting entirely. This module now focuses on JSON row lifecycle: create,
 * read, update, and import.
 */


// ---------------------------------------------------------------------------
// Public Entry Points
// ---------------------------------------------------------------------------

/**
 * Initializes attendance JSON tabs for all active employees for the given year.
 *
 * For each employee: finds or creates their tab, checks if a year row exists,
 * and appends a new empty year row if not. Idempotent — existing year rows
 * are skipped so the function is safe to re-run.
 *
 * @param {number} year — Calendar year (e.g. 2026).
 * @returns {{ created: number, updated: number, skipped: number }}
 */
function initAttendanceJsonTabs_(year) {
  const workbook       = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId  = workbook.getId();
  const employees      = getActiveEmployees_(); // ukgImport.js
  const now            = new Date().toISOString();

  // Build a name→sheet map in one getSheets() call instead of one getSheetByName() per employee.
  const sheetsMap = {};
  workbook.getSheets().forEach(function(sheet) {
    sheetsMap[sheet.getName()] = sheet;
  });

  let created = 0;
  let updated = 0;
  let skipped = 0;

  const newTabNames    = [];
  const existingSheets = [];

  employees.forEach(function(employee) {
    const tabName = employee.name + ' - ' + employee.id;
    if (sheetsMap[tabName]) {
      existingSheets.push(sheetsMap[tabName]);
    } else {
      newTabNames.push(tabName);
    }
  });

  // --- Batch create all missing tabs in two Sheets API requests ---
  // This replaces N individual insertSheet() calls (each ~1 second) with one
  // batchUpdate request and one Values.batchUpdate request (~2–5 seconds total).
  if (newTabNames.length > 0) {
    // Request 1: create all missing sheets at once.
    const addRequests = newTabNames.map(function(title) {
      return { addSheet: { properties: { title: title } } };
    });
    Sheets.Spreadsheets.batchUpdate({ requests: addRequests }, spreadsheetId);

    // Request 2: write header + year row to every new tab in one call.
    const headerAndYear = [['Year', 'Codes', 'StoredAt'], [year, '{}', now]];
    const valueData = newTabNames.map(function(tabName) {
      return { range: sheetsRangeA1_(tabName, 'A1:C2'), values: headerAndYear };
    });
    Sheets.Spreadsheets.Values.batchUpdate(
      { data: valueData, valueInputOption: 'RAW' },
      spreadsheetId
    );

    created = newTabNames.length;
  }

  // --- Check existing tabs for missing year rows ---
  // These sheets already exist so no creation cost; just read + conditionally append.
  existingSheets.forEach(function(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= ATTENDANCE_JSON.DATA_START_ROW) {
      const numDataRows = lastRow - ATTENDANCE_JSON.DATA_START_ROW + 1;
      const yearValues = sheet.getRange(
        ATTENDANCE_JSON.DATA_START_ROW, ATTENDANCE_JSON.COL_YEAR, numDataRows, 1
      ).getValues();
      for (let i = 0; i < yearValues.length; i++) {
        if (String(yearValues[i][0]) === String(year)) {
          skipped++;
          return;
        }
      }
    }
    sheet.appendRow([year, '{}', now]);
    updated++;
  });

  SpreadsheetApp.flush();
  // sortWorkbookSheets_() deliberately omitted — moving 300+ tabs one-by-one
  // takes longer than the creation itself. Tabs sort on the next operation that
  // already calls sortWorkbookSheets_().
  return { created, updated, skipped };
}


/**
 * Converts a tab name + cell range into Sheets API A1 notation.
 * Apostrophes in the tab name are doubled per the Sheets API spec.
 *
 * @param {string} tabName — Sheet tab name (may contain spaces/apostrophes).
 * @param {string} range   — Cell range (e.g. "A1:C2").
 * @returns {string}
 */
function sheetsRangeA1_(tabName, range) {
  return "'" + tabName.replace(/'/g, "''") + "'!" + range;
}


/**
 * Updates or inserts an attendance code for a specific employee and date.
 *
 * Finds the employee's tab, locates the year row for the given date, parses
 * the JSON, merges/replaces codes for that date, and writes back the updated
 * JSON string.
 *
 * @param {string} employeeId — Employee ID.
 * @param {string} isoDate    — ISO date string (e.g. "2026-01-15").
 * @param {string[]} codes    — Array of attendance codes (e.g. ["TD", "SE"]).
 * @returns {{ success: boolean, message: string, updatedCodes: object|null }}
 */
function updateAttendanceCode_(employeeId, isoDate, codes) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const date = new Date(isoDate);
  const year = date.getFullYear();

  // Find the employee's tab
  const allSheets = workbook.getSheets();
  let employeeSheet = null;
  for (let i = 0; i < allSheets.length; i++) {
    const sheetName = allSheets[i].getName();
    if (sheetName.endsWith(' - ' + employeeId)) {
      employeeSheet = allSheets[i];
      break;
    }
  }

  if (!employeeSheet) {
    return { success: false, message: 'Employee tab not found: ' + employeeId, updatedCodes: null };
  }

  // Find the row for this year
  const lastRow = employeeSheet.getLastRow();
  if (lastRow < ATTENDANCE_JSON.DATA_START_ROW) {
    return { success: false, message: 'No data rows in tab for ' + employeeId, updatedCodes: null };
  }

  const numDataRows = lastRow - ATTENDANCE_JSON.DATA_START_ROW + 1;
  const dataRange = employeeSheet.getRange(ATTENDANCE_JSON.DATA_START_ROW,
                                          ATTENDANCE_JSON.COL_YEAR,
                                          numDataRows, 3);
  const allRows = dataRange.getValues();

  let yearRowIndex = -1;
  for (let i = 0; i < allRows.length; i++) {
    if (allRows[i][0] === year || String(allRows[i][0]) === String(year)) {
      yearRowIndex = i;
      break;
    }
  }

  if (yearRowIndex === -1) {
    return { success: false, message: 'Year row not found for ' + year, updatedCodes: null };
  }

  // Parse the JSON, update the date, write back
  const codesJson = allRows[yearRowIndex][1];
  let schedule = {};
  try {
    schedule = codesJson ? JSON.parse(codesJson) : {};
  } catch (e) {
    return { success: false, message: 'Failed to parse JSON: ' + e.toString(), updatedCodes: null };
  }

  schedule[isoDate] = codes;

  // Write back to the sheet
  const actualRow = ATTENDANCE_JSON.DATA_START_ROW + yearRowIndex;
  const updatedJson = JSON.stringify(schedule);
  employeeSheet.getRange(actualRow, ATTENDANCE_JSON.COL_CODES_JSON).setValue(updatedJson);

  SpreadsheetApp.flush();
  return { success: true, message: 'Updated', updatedCodes: schedule };
}


/**
 * Imports attendance data from the migration project output.
 *
 * Receives a structured payload of { employeeId, employeeName, years: [{ year, codes }] }
 * and upserts each year row into the employee's tab. If a year row exists, it is
 * overwritten; otherwise, a new row is appended.
 *
 * @param {{ employeeId, employeeName, years: Array<{ year: number, codes: object }> }} payload
 * @returns {{ success: boolean, message: string, employee: string, yearsImported: number }}
 */
function importAttendanceJson_(payload) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const tabName = payload.employeeName + ' - ' + payload.employeeId;

  // Find or create the tab
  let sheet = workbook.getSheetByName(tabName);
  if (!sheet) {
    sheet = workbook.insertSheet(tabName);
    sheet.getRange(ATTENDANCE_JSON.HEADER_ROW, 1, 1, 3)
      .setValues([['Year', 'Codes', 'StoredAt']]);
  }

  const now = new Date().toISOString();
  let yearsImported = 0;

  payload.years.forEach(function(yearRow) {
    const year = yearRow.year;
    const codes = yearRow.codes || {};

    // Find existing year row
    const lastRow = sheet.getLastRow();
    let found = false;

    if (lastRow >= ATTENDANCE_JSON.DATA_START_ROW) {
      const numDataRows = lastRow - ATTENDANCE_JSON.DATA_START_ROW + 1;
      const dataRange = sheet.getRange(ATTENDANCE_JSON.DATA_START_ROW, ATTENDANCE_JSON.COL_YEAR,
                                       numDataRows, 3);
      const allRows = dataRange.getValues();

      for (let i = 0; i < allRows.length; i++) {
        if (allRows[i][0] === year || String(allRows[i][0]) === String(year)) {
          // Overwrite existing year row
          const actualRow = ATTENDANCE_JSON.DATA_START_ROW + i;
          sheet.getRange(actualRow, ATTENDANCE_JSON.COL_CODES_JSON).setValue(JSON.stringify(codes));
          sheet.getRange(actualRow, ATTENDANCE_JSON.COL_STORED_AT).setValue(now);
          found = true;
          yearsImported++;
          break;
        }
      }
    }

    if (!found) {
      // Append new year row
      sheet.appendRow([year, JSON.stringify(codes), now]);
      yearsImported++;
    }
  });

  SpreadsheetApp.flush();
  return {
    success: true,
    message: 'Imported ' + yearsImported + ' year(s)',
    employee: payload.employeeName,
    yearsImported: yearsImported,
  };
}
