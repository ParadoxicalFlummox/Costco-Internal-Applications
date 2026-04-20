/**
 * calendarParser.js — Reads individual employee tabs from the attendance controller.
 * VERSION: 0.2.3
 *
 * This file owns all logic for turning a raw attendance controller sheet into
 * structured event objects that the infraction detector can work with.
 *
 * Two responsibilities:
 *
 *   1. EMPLOYEE CONTEXT: Reading the employee's name, ID, department, hire date,
 *      and the fiscal year from the fixed metadata cells defined in EMPLOYEE_FIELDS
 *      (config.js). These fields are standardized across Costco warehouses.
 *
 *   2. CALENDAR GRID PARSING: The attendance controller lays out each month as a
 *      grid of columns (one per day of the week) across three horizontal bands.
 *      Within each band there are four month blocks side by side. Within each
 *      month block, cells contain either a day number (1–31) or an attendance
 *      code (TD, NS, etc.). Day numbers act as anchors — a code cell "belongs to"
 *      the most recently seen day number in its column.
 *
 *      The parser iterates every band × every month block × every cell, builds
 *      a date for each code from the month name + anchor day + year extracted
 *      from D1, and emits one CalendarEvent object per code.
 *
 * OUTPUT SHAPE — CalendarEvent:
 *   {
 *     employeeName:  string  — From EMPLOYEE_FIELDS.employeeName (cell X1)
 *     employeeId:    string  — From EMPLOYEE_FIELDS.employeeId (cell X3)
 *     department:    string  — From EMPLOYEE_FIELDS.department (cell R3)
 *     hireDate:      Date    — From EMPLOYEE_FIELDS.hireDate (cell AD3)
 *     month:         string  — Month name from the grid header row
 *     date:          Date    — Resolved date for this event
 *     code:          string  — Normalized attendance code (uppercase, letters only)
 *     isInfraction:  boolean — True if the code is in INFRACTION_CODES or CODE_RULES
 *     isIgnored:     boolean — True if the code is in IGNORE_CODES
 *     a1:            string  — A1 address of the cell (e.g. "B12") for debugging
 *   }
 */


// ---------------------------------------------------------------------------
// Public Entry Points
// ---------------------------------------------------------------------------

/**
 * Reads the employee metadata cells from a single attendance controller tab.
 *
 * The returned context object is passed through to every downstream function
 * so that employee identity fields do not need to be re-read from the sheet
 * multiple times per scan.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet — One employee tab.
 * @returns {{ sheetName, employeeName, employeeId, department, hireDate, yearTitle }}
 */
function readEmployeeContext_(sheet) {
  const fields = EMPLOYEE_FIELDS; // defined in config.js
  return {
    sheetName: sheet.getName(),
    employeeName: String(sheet.getRange(fields.employeeName).getDisplayValue() || '').trim(),
    employeeId: String(sheet.getRange(fields.employeeId).getDisplayValue() || '').trim(),
    department: String(sheet.getRange(fields.department).getDisplayValue() || '').trim(),
    hireDate: sheet.getRange(fields.hireDate).getValue(),
    yearTitle: String(sheet.getRange(fields.yearTitle).getDisplayValue() || '').trim(),
  };
}

/**
 * Parses all calendar events from a single attendance controller tab.
 *
 * Iterates all three data bands and all four month blocks within each band.
 * For each non-empty cell that is not a day number, one CalendarEvent is
 * emitted per attendance code found in that cell (cells may contain multiple
 * codes separated by spaces, commas, or slashes).
 *
 * Events whose codes appear in IGNORE_CODES are included in the returned
 * array with isIgnored=true so callers can log them if needed, but the
 * infraction detector filters them out before counting.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet — The employee tab to parse.
 * @param {number} year   — The fiscal year (from parseYearFromTitle_).
 * @param {string} timeZone — The script time zone string.
 * @param {{ employeeName, employeeId, department, hireDate }} ctx — Employee context.
 * @returns {CalendarEvent[]}
 */
function parseCalendarEvents_(sheet, year, timeZone, ctx) {
  const lastRow = sheet.getLastRow();
  const events = [];

  DATA_BANDS.forEach((band, bandIndex) => { // DATA_BANDS defined in config.js
    // The grid for this band runs from firstGridRow up to lastGridRow (explicit
    // boundary defined in config.js). If lastGridRow is not set, fall back to
    // the next band's monthRow - 1, or the sheet's last row for the final band.
    const nextBandStart = band.lastGridRow
      ? band.lastGridRow + 1
      : (bandIndex + 1 < DATA_BANDS.length)
        ? DATA_BANDS[bandIndex + 1].monthRow
        : lastRow + 1;

    const gridStartRow = band.firstGridRow;
    const gridEndRow = nextBandStart - 1;
    const numRows = Math.max(0, gridEndRow - gridStartRow + 1);
    if (numRows <= 0) return;

    START_COLUMNS.forEach(startColA1 => { // START_COLUMNS defined in config.js
      const startColIndex = colLetterToIndex_(startColA1); // 1-based

      // Read the month name from the header row of this block
      const monthName = String(
        sheet.getRange(band.monthRow, startColIndex).getDisplayValue() || ''
      ).trim();
      const monthIndex = monthNameToIndex_(monthName); // 0–11, or null
      if (monthIndex === null) return; // blank or unrecognized month — skip block

      // Read the entire grid block in one call (numRows × DAY_COLS_PER_BLOCK)
      const values = sheet
        .getRange(gridStartRow, startColIndex, numRows, DAY_COLS_PER_BLOCK)
        .getDisplayValues();

      // lastDayByCol tracks the most recently seen day number for each column
      // so that code cells can be anchored to the correct calendar date.
      const lastDayByCol = new Array(DAY_COLS_PER_BLOCK).fill(null);

      for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
        for (let colOffset = 0; colOffset < DAY_COLS_PER_BLOCK; colOffset++) {
          const cellValue = String(values[rowOffset][colOffset] == null ? '' : values[rowOffset][colOffset]).trim();
          if (!cellValue) continue;

          const absRow = gridStartRow + rowOffset;
          const absCol = startColIndex + colOffset;
          const a1Addr = colIndexToLetter_(absCol) + absRow;

          // If this cell is a day number (1–31), update the anchor for this column
          const parsed = parseInt(cellValue, 10);
          if (!isNaN(parsed) && parsed >= 1 && parsed <= 31) {
            lastDayByCol[colOffset] = parsed;
            continue;
          }

          // Not a day number — treat as one or more attendance codes
          const dayNum = lastDayByCol[colOffset];
          if (dayNum === null) continue; // no anchor day yet; skip

          const codes = normalizeAndSplitCodes_(cellValue);
          codes.forEach(code => {
            if (!code) return;

            const isIgnored = IGNORE_CODES.indexOf(code) >= 0;          // config.js
            const inInfractionList = INFRACTION_CODES.indexOf(code) >= 0;    // config.js
            const hasCodeRule = !!(CODE_RULES && CODE_RULES[code]);        // config.js
            const isInfraction = (inInfractionList || hasCodeRule) && !isIgnored;

            events.push({
              employeeName: ctx.employeeName,
              employeeId: ctx.employeeId,
              department: ctx.department,
              hireDate: ctx.hireDate,
              month: monthName,
              date: new Date(year, monthIndex, dayNum),
              code: code,
              isInfraction: isInfraction,
              isIgnored: isIgnored,
              a1: a1Addr,
            });
          });
        }
      }
    });
  });

  return events;
}

/**
 * Returns true if the given sheet tab looks like an individual employee
 * attendance controller tab (not an instruction or summary tab).
 *
 * Employee tabs follow the pattern "Last, First - EmployeeNumber"
 * e.g. "Le, Tony - 1234578". The EMPLOYEE_TAB_PATTERN regex in config.js
 * is the authoritative definition of this format.
 *
 * @param {string} sheetName — The tab name to test.
 * @returns {boolean}
 */
function isEmployeeTab_(sheetName) {
  return EMPLOYEE_TAB_PATTERN.test(sheetName); // EMPLOYEE_TAB_PATTERN defined in config.js
}

/**
 * Extracts a four-digit calendar year from the year title string in cell D1.
 *
 * e.g. "2026 Attendance Controller" → 2026
 * Returns null if no four-digit year is found; callers fall back to the
 * current calendar year.
 *
 * @param {string} titleString — The raw string from cell D1.
 * @returns {number|null}
 */
function parseYearFromTitle_(titleString) {
  if (!titleString) return null;
  const match = String(titleString).match(/\b(19|20)\d{2}\b/);
  return match ? parseInt(match[0], 10) : null;
}


// ---------------------------------------------------------------------------
// Code Normalization
// ---------------------------------------------------------------------------

/**
 * Splits a cell value into individual attendance codes and normalizes each one.
 *
 * Cells may contain multiple codes separated by spaces, commas, slashes, or
 * semicolons (e.g. "TD/SE" or "NS, TD"). Each token is uppercased and stripped
 * of any non-letter characters so that the output is always clean uppercase
 * alpha strings suitable for comparison against INFRACTION_CODES and CODE_RULES.
 *
 * @param {string} cellValue — The raw display value from a grid cell.
 * @returns {string[]} Array of normalized uppercase code strings (may be empty).
 */
function normalizeAndSplitCodes_(cellValue) {
  return String(cellValue)
    .split(/[\s,\/;|]+/)
    .map(token => token.trim().toUpperCase().replace(/[^A-Z]/g, ''))
    .filter(Boolean);
}


// ---------------------------------------------------------------------------
// Month Name Lookup
// ---------------------------------------------------------------------------

/**
 * Converts a month name string to its 0-based JavaScript month index.
 *
 * Matching is case-insensitive. Returns null for blank or unrecognized strings
 * so the caller can skip the block.
 *
 * @param {string} monthName — e.g. "January", "JANUARY", "january"
 * @returns {number|null} 0 for January … 11 for December, or null.
 */
function monthNameToIndex_(monthName) {
  if (!monthName) return null;
  const lookup = {
    'JANUARY': 0, 'FEBRUARY': 1, 'MARCH': 2, 'APRIL': 3,
    'MAY': 4, 'JUNE': 5, 'JULY': 6, 'AUGUST': 7,
    'SEPTEMBER': 8, 'OCTOBER': 9, 'NOVEMBER': 10, 'DECEMBER': 11,
  };
  return lookup.hasOwnProperty(String(monthName).trim().toUpperCase())
    ? lookup[String(monthName).trim().toUpperCase()]
    : null;
}


// ---------------------------------------------------------------------------
// Column Index / A1 Utilities
// ---------------------------------------------------------------------------

/**
 * Converts a column letter string to a 1-based column index.
 * "A" → 1, "B" → 2, "Z" → 26, "AA" → 27, etc.
 *
 * @param {string} letter — e.g. "A", "I", "Q", "Y"
 * @returns {number} 1-based column index.
 */
function colLetterToIndex_(letter) {
  const s = String(letter).toUpperCase();
  let index = 0;
  for (let i = 0; i < s.length; i++) {
    const charCode = s.charCodeAt(i);
    if (charCode < 65 || charCode > 90) continue;
    index = index * 26 + (charCode - 64);
  }
  return index;
}

/**
 * Converts a 1-based column index to its A1-notation letter string.
 * 1 → "A", 26 → "Z", 27 → "AA", etc.
 *
 * @param {number} index — 1-based column index.
 * @returns {string} A1-notation column letter(s).
 */
function colIndexToLetter_(index) {
  let result = '';
  let remaining = index;
  while (remaining > 0) {
    const remainder = (remaining - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return result;
}
