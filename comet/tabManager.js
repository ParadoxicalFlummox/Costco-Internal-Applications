/**
 * tabManager.js — Attendance controller tab generation for COMET.
 * VERSION: 0.2.4
 *
 * Generates one attendance controller sheet tab per active employee, following
 * the standardized Costco layout that calendarParser.js already knows how to read:
 *
 *   - Tab name:  "Last, First - EmployeeNumber"  (matches EMPLOYEE_TAB_PATTERN)
 *   - D1:        "[Year] Attendance Controller"
 *   - X1:        Employee name
 *   - R3:        Department
 *   - X3:        Employee ID
 *   - AD3:       Hire date
 *   - 3 bands × 4 month blocks × 7 day columns — day numbers placed under the
 *     correct day-of-week column, 4 rows per week slot (1 day-number row +
 *     3 empty rows for attendance code entry).
 *
 * IDEMPOTENT: existing tabs are skipped, so the function is safe to run again
 * if it times out partway through or new employees are added later.
 *
 * EXECUTION TIME: GAS has a 6-minute limit. At ~200ms per sheet, a roster of
 * 100 employees takes roughly 20 seconds. If you have 200+ employees on a slow
 * script host, run Setup a second time — the skipped count confirms progress.
 */


// ---------------------------------------------------------------------------
// Public Entry Point
// ---------------------------------------------------------------------------

/**
 * Creates attendance controller tabs for all active employees for the given year.
 *
 * @param {number} year — Calendar year (e.g. 2026).
 * @returns {{ created: number, skipped: number }}
 */
function generateAttendanceControllerTabs_(year) {
  const workbook  = SpreadsheetApp.getActiveSpreadsheet();
  const employees = getActiveEmployees_(); // ukgImport.js

  const MONTHS = [
    'JANUARY', 'FEBRUARY', 'MARCH',    'APRIL',
    'MAY',     'JUNE',     'JULY',     'AUGUST',
    'SEPTEMBER','OCTOBER', 'NOVEMBER', 'DECEMBER',
  ];
  const DAY_HEADERS = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'];

  let created = 0;
  let skipped = 0;

  employees.forEach(function(emp) {
    const tabName = emp.name + ' - ' + emp.id;

    if (workbook.getSheetByName(tabName)) {
      skipped++;
      return;
    }

    const sheet = workbook.insertSheet(tabName);

    // ---- Metadata cells ----
    sheet.getRange(EMPLOYEE_FIELDS.yearTitle)   .setValue(year + ' Attendance Controller');
    sheet.getRange(EMPLOYEE_FIELDS.employeeName).setValue(emp.name);
    sheet.getRange(EMPLOYEE_FIELDS.department)  .setValue(emp.department);
    sheet.getRange(EMPLOYEE_FIELDS.employeeId)  .setValue(emp.id);
    if (emp.hireDate) sheet.getRange(EMPLOYEE_FIELDS.hireDate).setValue(emp.hireDate);

    // ---- Calendar grid ----
    DATA_BANDS.forEach(function(band, bandIndex) {
      START_COLUMNS.forEach(function(colLetter, blockIndex) {
        const monthIndex = bandIndex * 4 + blockIndex;
        if (monthIndex >= 12) return;

        const startColIdx = colLetterToIndex_(colLetter); // calendarParser.js

        // Month name header
        const monthNameCell = sheet.getRange(band.monthRow, startColIdx);
        monthNameCell.setValue(MONTHS[monthIndex]);
        monthNameCell.setFontWeight('bold');

        // Day-of-week header row
        sheet.getRange(band.dayOfWeekRow, startColIdx, 1, DAY_COLS_PER_BLOCK)
          .setValues([DAY_HEADERS])
          .setFontWeight('bold')
          .setBackground('#E8EAF6')
          .setHorizontalAlignment('center');

        // Day numbers — written as one batch per month block
        writeDayNumberGrid_(
          sheet, year, monthIndex,
          band.firstGridRow, band.lastGridRow, startColIdx
        );
      });
    });

    formatAttendanceControllerSheet_(sheet, year, emp);
    created++;
  });

  sortWorkbookSheets_(); // api.js
  return { created, skipped };
}


// ---------------------------------------------------------------------------
// Calendar Grid Writer
// ---------------------------------------------------------------------------

/**
 * Builds the day-number grid for one month block and writes it in a single
 * setValues() call to minimize GAS API round-trips.
 *
 * Layout: 4 rows per week slot (row 0 = day numbers, rows 1–3 = code entry).
 * This gives 6 week slots across 24 rows, which covers any month regardless
 * of start day (worst case: 31-day month starting on Saturday = 6 partial weeks).
 *
 * @param {Sheet}  sheet
 * @param {number} year         — 4-digit calendar year.
 * @param {number} monthIndex   — 0-based month index (0 = January).
 * @param {number} startRow     — First row of the grid (1-indexed).
 * @param {number} endRow       — Last row of the grid (inclusive).
 * @param {number} startColIdx  — First column of the month block (1-indexed).
 */
function writeDayNumberGrid_(sheet, year, monthIndex, startRow, endRow, startColIdx) {
  const numRows     = endRow - startRow + 1; // 24
  const rowsPerSlot = 4;

  // Build a blank numRows × 7 grid
  const grid = [];
  for (let r = 0; r < numRows; r++) {
    grid.push(['', '', '', '', '', '', '']);
  }

  const firstDay   = new Date(year, monthIndex, 1);
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();

  // Convert JS getDay() (0=Sun) to Mon-based offset (0=Mon … 6=Sun)
  let dow = firstDay.getDay();
  dow = (dow === 0) ? 6 : dow - 1;

  let day  = 1;
  let slot = 0;
  let col  = dow;

  while (day <= daysInMonth) {
    const rowIdx = slot * rowsPerSlot;
    if (rowIdx >= numRows) break; // safety — should not happen with 24-row grids
    grid[rowIdx][col] = day;
    day++;
    col++;
    if (col >= 7) {
      col = 0;
      slot++;
    }
  }

  sheet.getRange(startRow, startColIdx, numRows, DAY_COLS_PER_BLOCK).setValues(grid);
}


// ---------------------------------------------------------------------------
// Sheet Formatting
// ---------------------------------------------------------------------------

/**
 * Applies visual formatting to a newly created attendance controller tab so
 * it is readable on screen and clean when printed.
 *
 * Layout overview (columns A–AH, 34 columns):
 *   Each band occupies 4 blocks × 8 columns (7 day cols + 1 spacer) = 32 cols.
 *   Cols A–G, I–O, Q–W, Y–AE are the 7 day columns per month block.
 *   Cols H, P, X are visual spacers between blocks.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} year
 * @param {{ name: string, id: string, department: string }} emp
 */
function formatAttendanceControllerSheet_(sheet, year, emp) {
  // ---- Header rows (rows 1–4): employee identity ----
  sheet.getRange('A1:AH4').setBackground('#263238').setFontColor('#FFFFFF');
  sheet.getRange(EMPLOYEE_FIELDS.yearTitle)
    .setFontSize(14).setFontWeight('bold');
  sheet.getRange(EMPLOYEE_FIELDS.employeeName)
    .setFontSize(11).setFontWeight('bold');
  sheet.getRange(EMPLOYEE_FIELDS.department)
    .setFontSize(10);
  sheet.getRange(EMPLOYEE_FIELDS.employeeId)
    .setFontSize(10);
  if (emp.hireDate) {
    sheet.getRange(EMPLOYEE_FIELDS.hireDate).setFontSize(10);
  }
  // Freeze the identity rows so they stay visible while scrolling
  sheet.setFrozenRows(4);

  // ---- Column widths ----
  // Day columns 30px each; spacer columns narrower
  const totalCols = 34; // A(1)–AH(34)
  for (let col = 1; col <= totalCols; col++) {
    // Spacer columns: H(8), P(16), X(24)
    if (col === 8 || col === 16 || col === 24) {
      sheet.setColumnWidth(col, 10);
    } else {
      sheet.setColumnWidth(col, 32);
    }
  }

  // ---- Band formatting ----
  const BAND_HEADER_BG    = '#E8EAF6'; // lavender — month name + day-of-week rows
  const BAND_DAY_NUM_BG   = '#F5F5F5'; // light gray — day-number rows
  const BAND_CODE_BG      = '#FFFFFF'; // white — attendance code entry rows
  const BAND_ALT_BG       = '#FAFAFA'; // very light — alternating code rows 2 & 3
  const BORDER_COLOR      = '#B0BEC5'; // blue-gray border

  DATA_BANDS.forEach(function(band) {
    START_COLUMNS.forEach(function(colLetter) {
      const startCol = colLetterToIndex_(colLetter); // calendarParser.js
      const endCol   = startCol + DAY_COLS_PER_BLOCK - 1;

      // Month name row
      sheet.getRange(band.monthRow, startCol, 1, DAY_COLS_PER_BLOCK)
        .setBackground(BAND_HEADER_BG)
        .setFontWeight('bold')
        .setHorizontalAlignment('left')
        .setBorder(true, true, true, true, false, false, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);

      // Day-of-week header row
      sheet.getRange(band.dayOfWeekRow, startCol, 1, DAY_COLS_PER_BLOCK)
        .setBackground(BAND_HEADER_BG)
        .setBorder(true, true, true, true, true, false, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);

      // Grid rows: 4 rows per week slot × 6 slots = 24 rows
      const totalGridRows = band.lastGridRow - band.firstGridRow + 1;
      const rowsPerSlot   = 4;
      for (let slot = 0; slot < totalGridRows / rowsPerSlot; slot++) {
        const baseRow = band.firstGridRow + slot * rowsPerSlot;

        // Row 0 of slot: day numbers
        sheet.getRange(baseRow, startCol, 1, DAY_COLS_PER_BLOCK)
          .setBackground(BAND_DAY_NUM_BG)
          .setHorizontalAlignment('center')
          .setFontSize(8)
          .setBorder(true, true, false, true, true, false, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
        sheet.setRowHeight(baseRow, 18);

        // Rows 1–3 of slot: code entry rows
        for (let r = 1; r <= 3; r++) {
          const bg = r === 1 ? BAND_CODE_BG : BAND_ALT_BG;
          sheet.getRange(baseRow + r, startCol, 1, DAY_COLS_PER_BLOCK)
            .setBackground(bg)
            .setHorizontalAlignment('center')
            .setFontSize(8)
            .setBorder(
              false, true, r === 3, true, true, false,
              BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID
            );
          sheet.setRowHeight(baseRow + r, 20);
        }
      }
    });

    // Row heights for month and day-of-week header rows
    sheet.setRowHeight(band.monthRow,    20);
    sheet.setRowHeight(band.dayOfWeekRow, 18);

    // Spacer row between bands (the row just above each band's monthRow, except band 0)
    if (band.monthRow > 5) {
      sheet.setRowHeight(band.monthRow - 1, 10);
    }
  });

  // ---- Legend for infraction codes ----
  // Positioned below the main grid (rows 85+) for easy reference
  const legendStartRow = 85;
  sheet.getRange(legendStartRow, 1).setValue('INFRACTION CODE LEGEND').setFontWeight('bold').setFontSize(11);

  const infraCodes = [
    { code: 'TD', name: 'Tardy', color: '#FFCCCC' },
    { code: 'NS', name: 'No Show', color: '#FF9999' },
    { code: 'SE', name: 'Swiping Error', color: '#FFFF99' },
    { code: 'MP', name: 'Meal Period', color: '#CCFFCC' },
    { code: 'SZ', name: 'Suspension', color: '#CC99FF' },
  ];

  infraCodes.forEach(function(item, index) {
    const row = legendStartRow + 1 + index;
    const codeCell = sheet.getRange(row, 1);
    codeCell.setValue(item.code + ' — ' + item.name)
      .setBackground(item.color)
      .setFontSize(10)
      .setHorizontalAlignment('left')
      .setBorder(true, true, true, true, false, false, '#B0BEC5', SpreadsheetApp.BorderStyle.SOLID);
  });

  // ---- Conditional formatting for infraction codes ----
  // Apply conditional formatting to all code entry rows in all bands.
  // Each rule targets a specific code and applies a background color.
  const ranges = [];

  DATA_BANDS.forEach(function(band) {
    START_COLUMNS.forEach(function(colLetter) {
      const startCol = colLetterToIndex_(colLetter);
      const totalGridRows = band.lastGridRow - band.firstGridRow + 1;
      const rowsPerSlot = 4;

      for (let slot = 0; slot < totalGridRows / rowsPerSlot; slot++) {
        const baseRow = band.firstGridRow + slot * rowsPerSlot;

        // Rows 1–3 of each slot are code entry rows
        for (let r = 1; r <= 3; r++) {
          const codeEntryRow = baseRow + r;
          ranges.push(sheet.getRange(codeEntryRow, startCol, 1, DAY_COLS_PER_BLOCK));
        }
      }
    });
  });

  if (ranges.length > 0) {
    sheet.clearConditionalFormatRules();

    // TD — Tardy (light red)
    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('TD')
      .setBackground('#FFCCCC')
      .setRanges(ranges)
      .build();
    sheet.setConditionalFormatRules([rule]);

    // NS — No Show (darker red)
    rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('NS')
      .setBackground('#FF9999')
      .setRanges(ranges)
      .build();
    sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat([rule]));

    // SE — Swiping Error (yellow)
    rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('SE')
      .setBackground('#FFFF99')
      .setRanges(ranges)
      .build();
    sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat([rule]));

    // MP — Meal Period (light green)
    rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('MP')
      .setBackground('#CCFFCC')
      .setRanges(ranges)
      .build();
    sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat([rule]));

    // SZ — Suspension (light purple)
    rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('SZ')
      .setBackground('#CC99FF')
      .setRanges(ranges)
      .build();
    sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat([rule]));
  }

  // ---- Print settings ----
  // Set landscape orientation and fit all columns to page width
  sheet.setTabColor('#4A90D9'); // blue tab so attendance tabs are visually distinct
}

