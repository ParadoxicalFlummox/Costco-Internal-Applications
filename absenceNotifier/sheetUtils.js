/**
 * sheetUtils.js — Fiscal calendar math and call log sheet name calculation.
 * VERSION: 0.2.1
 *
 * This file owns two things:
 *
 *   1. SHEET NAME RESOLUTION: Determining the title of the current call log
 *      sheet. The title is either a fiscal period/week label ("P3 W1") when
 *      a FY start date is configured, or a "Week Ending MM/DD/YY" fallback
 *      when it is not. Both formats are supported so the notifier works
 *      out of the box even before a manager sets up the config sheet.
 *
 *   2. FISCAL CALENDAR MATH: Converting today's date into a period and week
 *      number using the Costco fiscal year structure:
 *        - 13 periods per fiscal year
 *        - 4 weeks per period
 *        - 7 days per week  (52 weeks total)
 *      Given a FY start date (the Monday of P1 W1) and any reference date,
 *      this file can calculate which period and week that date falls into.
 *
 * WHY SEPARATE FROM config.js:
 *   The FY start date is a configurable input (read from a sheet cell at
 *   runtime). The logic that transforms that input into a period/week label
 *   is non-trivial and testable in isolation. Keeping it here means config.js
 *   stays as pure constants, and this file handles the one computed value
 *   that drives sheet name lookups across the whole system.
 */


// ---------------------------------------------------------------------------
// Public Entry Point
// ---------------------------------------------------------------------------

/**
 * Returns the sheet tab title for the call log that covers the current date.
 *
 * Resolution order:
 *   1. Read the FY start date from the "Absence Config" sheet (cell B2).
 *   2. If a valid date is found, compute "P# W#" from the fiscal calendar.
 *   3. If the config sheet is missing, the cell is empty, or the value is not
 *      a valid date, fall back to "Week Ending MM/DD/YY".
 *
 * This function is called by sendAbsenceDigest() in absenceNotifier.js once
 * per trigger run to identify which sheet to scan for absence entries.
 *
 * @returns {string} The sheet title, e.g. "P3 W1" or "Week Ending 10/19/25".
 */
function getActiveCallLogSheetName_() {
  const fyStartDate = readFiscalYearStartDate_();

  if (fyStartDate) {
    return calculateFiscalPeriodWeekLabel_(fyStartDate, new Date());
  }

  return calculateWeekEndingLabel_(new Date());
}

/**
 * Returns the fiscal year start date as used in the sheet header row.
 * e.g. given a start date of Sep 1, 2025 and today = Apr 8, 2026 → "FY'26"
 *
 * This is used by sheetGenerator.js when writing the merged header in row 1
 * so the fiscal year label matches what managers expect to see.
 *
 * @returns {string} e.g. "FY'26"
 */
function getCurrentFiscalYearLabel_() {
  const fyStartDate = readFiscalYearStartDate_();
  if (!fyStartDate) return "FY'??";

  // The fiscal year label is derived from the calendar year the FY started in,
  // incremented by 1 (FY26 starts in the fall of calendar year 2025).
  const fiscalYear = fyStartDate.getFullYear() + 1;
  return `FY'${String(fiscalYear).slice(2)}`;
}


// ---------------------------------------------------------------------------
// Fiscal Calendar Calculation
// ---------------------------------------------------------------------------

/**
 * Converts a reference date into a "P# W#" fiscal period/week label.
 *
 * Costco's fiscal year structure:
 *   13 periods × 4 weeks × 7 days = 364 days (52 weeks)
 *
 * The calculation:
 *   1. Find how many whole days have elapsed from the FY start to the reference date.
 *   2. Divide by 7 to get the 0-based week index within the fiscal year.
 *   3. Divide that by 4 to get the 0-based period index; the remainder is the
 *      0-based week-within-period index.
 *   4. Add 1 to both to get human-readable 1-based numbers.
 *
 * Edge cases:
 *   - If the reference date is before the FY start date (daysElapsed < 0),
 *     the function returns "P? W?" rather than producing a nonsense label.
 *   - After week 52 (day 364+), the function wraps into the next fiscal year
 *     naturally because the math continues past 13 periods. For a more
 *     sophisticated implementation, the caller would pass the correct FY start
 *     date for that year. In practice, the manager updates the config sheet
 *     each year, so this is a one-time manual step.
 *
 * @param {Date} fyStartDate     — The Monday of P1 W1 for this fiscal year.
 * @param {Date} referenceDate   — The date to compute the label for (typically today).
 * @returns {string} e.g. "P3 W1"
 */
function calculateFiscalPeriodWeekLabel_(fyStartDate, referenceDate) {
  // Truncate both dates to midnight to avoid hour/DST drift affecting the day count.
  const startMidnight = new Date(fyStartDate);
  startMidnight.setHours(0, 0, 0, 0);

  const referenceMidnight = new Date(referenceDate);
  referenceMidnight.setHours(0, 0, 0, 0);

  const millisPerDay = 24 * 60 * 60 * 1000;
  const daysElapsed  = Math.floor((referenceMidnight - startMidnight) / millisPerDay);

  if (daysElapsed < 0) {
    // Reference date is before the fiscal year start — config is probably wrong.
    console.warn(`sheetUtils: Reference date ${referenceDate} is before FY start ${fyStartDate}. Returning fallback label.`);
    return calculateWeekEndingLabel_(referenceDate);
  }

  const weekIndex   = Math.floor(daysElapsed / 7); // 0-based week in the fiscal year
  const periodIndex = Math.floor(weekIndex / 4);   // 0-based period (0 = Period 1)
  const weekInPeriod = (weekIndex % 4);             // 0-based week within the period

  const period = periodIndex + 1;
  const week   = weekInPeriod + 1;

  return `P${period} W${week}`;
}

/**
 * Returns the fiscal period and week numbers for a given date as separate values.
 *
 * This is the same math as calculateFiscalPeriodWeekLabel_() but returns an object
 * instead of a formatted string. Used by sheetGenerator.js when it needs the period
 * and week numbers separately to write the merged row 1 header.
 *
 * @param {Date} fyStartDate   — The Monday of P1 W1.
 * @param {Date} referenceDate — The date to evaluate (typically today).
 * @returns {{ period: number, week: number } | null} null if reference is before FY start.
 */
function calculateFiscalPeriodAndWeek_(fyStartDate, referenceDate) {
  const startMidnight = new Date(fyStartDate);
  startMidnight.setHours(0, 0, 0, 0);

  const referenceMidnight = new Date(referenceDate);
  referenceMidnight.setHours(0, 0, 0, 0);

  const millisPerDay = 24 * 60 * 60 * 1000;
  const daysElapsed  = Math.floor((referenceMidnight - startMidnight) / millisPerDay);

  if (daysElapsed < 0) return null;

  const weekIndex    = Math.floor(daysElapsed / 7);
  const periodIndex  = Math.floor(weekIndex / 4);
  const weekInPeriod = weekIndex % 4;

  return {
    period: periodIndex + 1,
    week:   weekInPeriod + 1,
  };
}


// ---------------------------------------------------------------------------
// Week Date Range Label
// ---------------------------------------------------------------------------

/**
 * Builds a human-readable date range label for the current call log week.
 *
 * Produces the same style used by the AutoScheduler's week header:
 *   "April 6 – 12, 2026"
 *
 * The call log week runs Monday through Sunday. Given today's date, this
 * function finds the Monday that started the current week and the Sunday
 * that ends it, then formats them as a single range string.
 *
 * GAS's V8 Intl implementation does not produce clean output for partial date
 * option sets like { day, year } without month — it renders "(day: 12) 2026".
 * The label is therefore built manually:
 *   "{month} {startDay} – {endDay}, {year}"
 * where the month and year are taken from the Sunday (end of week) so that
 * weeks spanning a month boundary (e.g. March 31 – April 6) display the
 * ending month name, matching the AutoScheduler convention.
 *
 * @param {Date} referenceDate — The date to calculate the current week from (typically today).
 * @returns {string} e.g. "April 6 – 12, 2026"
 */
function calculateWeekDateRangeLabel_(referenceDate) {
  const dayOfWeek = referenceDate.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

  // Find the Monday that started this week.
  // If today is Sunday (0), Monday was 6 days ago.
  const daysFromMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  const monday = new Date(referenceDate);
  monday.setDate(referenceDate.getDate() - daysFromMonday);

  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);

  // Format: month and start day from Monday, end day and year from Sunday.
  // Using Sunday's month handles weeks that cross a month boundary correctly
  // (e.g. "March 31 – April 6" would become "April 31 – 6" which is wrong;
  // instead we use Monday's month for the start and Sunday for the rest).
  const startMonth = monday.toLocaleDateString('en-US', { month: 'long' });
  const startDay   = monday.getDate();
  const endDay     = sunday.getDate();
  const year       = sunday.getFullYear();

  // If the week crosses a month boundary, show both month names.
  const endMonth = sunday.toLocaleDateString('en-US', { month: 'long' });
  if (startMonth !== endMonth) {
    return `${startMonth} ${startDay} \u2013 ${endMonth} ${endDay}, ${year}`;
  }

  return `${startMonth} ${startDay} \u2013 ${endDay}, ${year}`;
}


// ---------------------------------------------------------------------------
// Fallback: Week Ending Label
// ---------------------------------------------------------------------------

/**
 * Calculates the "Week Ending MM/DD/YY" sheet title for the current calendar week.
 *
 * Used when no FY start date is configured. The "week ending" date is always
 * the upcoming Sunday from the reference date. This matches the naming
 * convention used by the original call log before fiscal period labels were added.
 *
 * @param {Date} referenceDate — The date to calculate from (typically today).
 * @returns {string} e.g. "Week Ending 10/19/25"
 */
function calculateWeekEndingLabel_(referenceDate) {
  const dayOfWeek = referenceDate.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

  // Days until the next Sunday. If today is Sunday (0), that is 7 days away
  // because we want the END of the current week, not the start of the next.
  const daysUntilSunday = 7 - dayOfWeek;

  const weekEndDate = new Date(referenceDate);
  weekEndDate.setDate(referenceDate.getDate() + daysUntilSunday);

  const month = weekEndDate.getMonth() + 1; // getMonth() is 0-indexed
  const day   = weekEndDate.getDate();
  const year  = String(weekEndDate.getFullYear()).slice(2); // "2025" → "25"

  return `Week Ending ${month}/${day}/${year}`;
}


// ---------------------------------------------------------------------------
// Config Sheet Reader
// ---------------------------------------------------------------------------

/**
 * Reads the fiscal year start date from the "Absence Config" sheet.
 *
 * The manager enters a date (e.g. "9/1/2025") into cell B2 of the config sheet.
 * This function retrieves that value and coerces it into a JavaScript Date.
 *
 * Returns null (without throwing) in any of these cases:
 *   - The config sheet does not exist in the workbook
 *   - Cell B2 is empty
 *   - Cell B2 contains a value that cannot be interpreted as a real date
 *
 * The caller (getActiveCallLogSheetName_) treats null as "use the fallback label".
 *
 * @returns {Date|null} The FY start date, or null if unavailable.
 */
function readFiscalYearStartDate_() {
  try {
    const workbook     = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet  = workbook.getSheetByName(CONFIG_SHEET_NAME);

    if (!configSheet) return null;

    const cellValue = configSheet.getRange(FISCAL_YEAR_START_CELL).getValue();

    if (!cellValue || cellValue === '') return null;

    // Sheets may deliver the value as a Date object directly if the cell is
    // date-formatted, or as a string if it was typed in plain text.
    if (cellValue instanceof Date) {
      return isNaN(cellValue.getTime()) ? null : cellValue;
    }

    if (typeof cellValue === 'string') {
      const parsed = new Date(cellValue);
      return isNaN(parsed.getTime()) ? null : parsed;
    }

    // Number: Sheets serial date (days since Jan 1, 1900)
    if (typeof cellValue === 'number' && cellValue >= 1) {
      const epochMilliseconds = (cellValue - 25569) * 86400000;
      const parsed = new Date(epochMilliseconds);
      return isNaN(parsed.getTime()) ? null : parsed;
    }

    return null;
  } catch (error) {
    console.warn('sheetUtils: Could not read FY start date from config sheet —', error);
    return null;
  }
}
