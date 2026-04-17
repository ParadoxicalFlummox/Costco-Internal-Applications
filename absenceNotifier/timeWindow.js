/**
 * timeWindow.js — Time window calculations and time value parsing utilities.
 * VERSION: 1.0.0
 *
 * This file owns all logic related to time in the Absence Notifier. Its two
 * primary responsibilities are:
 *
 *   1. WINDOW CALCULATION: Determining the previous N-minute window boundary.
 *      The notifier fires every N minutes and looks back at the most recently
 *      completed window. For example, a trigger firing at 09:17 with a 15-minute
 *      window produces a window of (09:00, 09:15].
 *
 *   2. TIME VALUE PARSING: Google Sheets stores time in several formats depending
 *      on how the cell was formatted and how the data was entered. A time cell may
 *      arrive as a JavaScript Date object (a "time-only" Date anchored to
 *      December 30, 1899), a decimal fraction of a day (e.g., 0.375 = 09:00), or
 *      a string such as "9:00 AM". This file normalizes all of those formats into
 *      a single epoch millisecond value suitable for window comparison.
 *
 * WHY SEPARATE FROM dataIngestion.js:
 *   Time math is the most complex and testable part of the notifier. Isolating it
 *   here means it can be reasoned about and debugged independently of the sheet
 *   read logic. dataIngestion.js calls into this file but does not contain any
 *   time arithmetic itself.
 */


// ---------------------------------------------------------------------------
// Window Calculation
// ---------------------------------------------------------------------------

/**
 * Computes the start and end boundaries of the most recently completed
 * N-minute window, relative to the current clock time.
 *
 * The "previous window" is defined as the last complete slot that ended before
 * right now. For a 15-minute cadence:
 *   - A trigger firing at 09:17 returns { start: 09:00, end: 09:15 }
 *   - A trigger firing at 09:00 returns { start: 08:45, end: 09:00 }
 *
 * The window is expressed as a half-open interval (start, end], meaning the
 * end boundary is included and the start boundary is excluded. This matches how
 * the filter in dataIngestion.js evaluates time values:
 *   callTimeMs > windowStartMs  AND  callTimeMs <= windowEndMs
 *
 * @param {number} windowMinutes — The length of each window in minutes.
 *   Must be a positive integer that divides evenly into 60 (e.g., 15 or 30).
 * @returns {{ start: Date, end: Date }} The start and end of the previous window.
 */
function getPreviousWindow_(windowMinutes) {
  const now = new Date();
  const slotLengthMinutes = windowMinutes | 0; // Coerce to integer; guards against float input

  // Find the most recent window boundary by snapping the current minute down
  // to the nearest multiple of slotLengthMinutes.
  // Example: now = 09:17, slotLength = 15 → floor(17/15)*15 = 15 → boundary = 09:15
  const currentMinutes = now.getMinutes();
  const boundaryMinute = Math.floor(currentMinutes / slotLengthMinutes) * slotLengthMinutes;

  const windowEndBoundary = new Date(now);
  windowEndBoundary.setMinutes(boundaryMinute, 0, 0); // zero out seconds and milliseconds

  // The window start is exactly one slot length before the end boundary.
  const windowStartBoundary = new Date(windowEndBoundary.getTime() - slotLengthMinutes * 60 * 1000);

  return {
    start: windowStartBoundary,
    end: windowEndBoundary,
  };
}


// ---------------------------------------------------------------------------
// Time Value Parsing
// ---------------------------------------------------------------------------

/**
 * Converts a raw cell value from the "Time Called" column (column H) into an
 * epoch millisecond timestamp for window comparison.
 *
 * Google Sheets delivers time values in three possible formats depending on cell
 * formatting and input method:
 *
 *   1. Date object with year 1899 or 1900 — a "time-only" Date. Sheets stores
 *      time-only values as a Date anchored to December 30, 1899. The year will
 *      be 1899 or 1900 (effectively < 1971). We extract h/m/s and anchor them
 *      to the window's calendar date.
 *
 *   2. Date object with year >= 1971 — a real datetime. We use it as-is.
 *
 *   3. Number < 1 — a fractional day (e.g., 0.375 = 09:00). We convert to
 *      total seconds, extract h/m/s, and anchor to the window date.
 *
 *   4. Number >= 1 — an Excel/Sheets serial date (days since Jan 1, 1900).
 *      We convert to epoch milliseconds directly.
 *
 *   5. String — a human-typed time such as "9:00 AM" or "14:30:00". We try
 *      a strict HH:MM[:SS] AM/PM regex first, then fall back to Date.parse().
 *
 * @param {Date|number|string|null} cellValue — The raw value from the time cell.
 * @param {{ start: Date, end: Date }} window — The current time window, used to
 *   anchor time-only values (h/m/s) to the correct calendar day.
 * @returns {number|null} Epoch milliseconds, or null if the value cannot be parsed.
 */
function parseTimeToMilliseconds_(cellValue, window) {
  if (cellValue == null || cellValue === '') return null;

  // --- Format 1 & 2: Date object ---
  if (cellValue instanceof Date) {
    const parsedTime = cellValue.getTime();
    if (isNaN(parsedTime)) return null;

    // A year >= 1971 means this is a real datetime with a known calendar date.
    // We trust it directly without anchoring.
    if (cellValue.getFullYear() >= 1971) return parsedTime;

    // A year < 1971 means Sheets is using the 1899 epoch as a time-only placeholder.
    // We extract h/m/s and anchor them to the window's calendar date.
    return anchorTimePartsToWindow_(
      cellValue.getHours(),
      cellValue.getMinutes(),
      cellValue.getSeconds(),
      window
    );
  }

  // --- Format 3 & 4: Number ---
  if (typeof cellValue === 'number') {
    if (cellValue < 1) {
      // Fractional day: convert to total seconds then extract h/m/s
      const totalSeconds = Math.round(cellValue * 86400);
      const hours = Math.floor(totalSeconds / 3600) % 24;
      const minutes = Math.floor((totalSeconds % 3600) / 60);
      const seconds = totalSeconds % 60;
      return anchorTimePartsToWindow_(hours, minutes, seconds, window);
    } else {
      // Excel/Sheets serial date: 1 = January 1, 1900
      // SHEETS_EPOCH_OFFSET (25569) is the Unix epoch in Sheets serial days.
      const epochMilliseconds = (cellValue - SHEETS_EPOCH_OFFSET) * 86400000;
      const parsed = new Date(epochMilliseconds);
      return isNaN(parsed.getTime()) ? null : parsed.getTime();
    }
  }

  // --- Format 5: String ---
  if (typeof cellValue === 'string') {
    const trimmed = cellValue.trim();

    // Try a strict HH:MM[:SS] AM/PM pattern before falling back to Date.parse(),
    // because Date.parse() is locale-dependent and can silently produce wrong results.
    const twelveHourPattern = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)$/i;
    const regexMatch = trimmed.match(twelveHourPattern);

    if (regexMatch) {
      let hours = parseInt(regexMatch[1], 10);
      const minutes = parseInt(regexMatch[2], 10);
      const seconds = regexMatch[3] ? parseInt(regexMatch[3], 10) : 0;
      const meridiem = regexMatch[4].toUpperCase();

      // Standard 12-hour to 24-hour conversion:
      //   12:xx AM → 0:xx (midnight hour)
      //   12:xx PM → 12:xx (noon hour, no change)
      //    1-11 AM → unchanged
      //    1-11 PM → add 12
      if (hours === 12) hours = 0;
      if (meridiem === 'PM') hours += 12;

      return anchorTimePartsToWindow_(hours % 24, minutes % 60, seconds % 60, window);
    }

    // Last resort: let JavaScript try to parse the string as a date/time.
    const fallbackDate = new Date(trimmed);
    return isNaN(fallbackDate.getTime()) ? null : fallbackDate.getTime();
  }

  return null;
}

/**
 * Given a time expressed as hours, minutes, and seconds, anchors it to the
 * calendar date of the window's end boundary and returns epoch milliseconds.
 *
 * This is necessary because "time-only" values from Sheets carry no calendar date.
 * By anchoring to the window's end date, we ensure the resulting timestamp falls
 * on the same calendar day as the window we are scanning — which is always the
 * intent when an employee calls in and we record just the clock time.
 *
 * MIDNIGHT ROLLOVER GUARD:
 *   If the anchored timestamp falls AFTER the window's end boundary, we subtract
 *   24 hours. This handles the edge case where a time-only value (e.g., "11:50 PM")
 *   was logged just before midnight but the window end crossed into the next day.
 *
 * @param {number} hours   — 24-hour hour (0–23)
 * @param {number} minutes — Minutes (0–59)
 * @param {number} seconds — Seconds (0–59)
 * @param {{ start: Date, end: Date }} window — The current scanning window.
 * @returns {number} Epoch milliseconds for this time on the window's calendar date.
 */
function anchorTimePartsToWindow_(hours, minutes, seconds, window) {
  const anchoredDate = new Date(window.end); // clone the end boundary to inherit its calendar date
  anchoredDate.setHours(hours, minutes, seconds, 0);

  let anchoredMilliseconds = anchoredDate.getTime();

  // If anchoring produced a timestamp that is strictly after the window end,
  // the time-only value belongs to the previous calendar day — roll it back.
  if (anchoredMilliseconds > window.end.getTime()) {
    anchoredMilliseconds -= 24 * 60 * 60 * 1000;
  }

  return anchoredMilliseconds;
}


// ---------------------------------------------------------------------------
// Date Validation Helpers
// ---------------------------------------------------------------------------

/**
 * Attempts to extract a real calendar date from a raw cell value in the
 * "Date" column (column A).
 *
 * Column A is used to verify that a row belongs to the same calendar day as
 * the window currently being scanned. Without this check, a time-only value
 * in column H could produce a false positive for a row that belongs to a
 * different day.
 *
 * A "real" calendar date is any value that resolves to a Date whose year
 * is 1971 or later. Values that resolve to the 1899/1900 Sheets epoch are
 * considered time-only placeholders and are rejected (returned as null).
 *
 * @param {Date|number|string|null} cellValue — The raw value from the date cell.
 * @returns {Date|null} A Date object if a real calendar date could be extracted,
 *   or null if the value is empty, a time-only placeholder, or unparseable.
 */
function coerceToCalendarDate_(cellValue) {
  if (cellValue == null || cellValue === '') return null;

  if (cellValue instanceof Date) {
    if (isNaN(cellValue.getTime())) return null;
    // Pre-1971 years indicate Sheets' time-only epoch placeholder — not a real date.
    return cellValue.getFullYear() >= 1971 ? cellValue : null;
  }

  if (typeof cellValue === 'number') {
    // A fractional number (< 1) represents a time-only value with no calendar component.
    if (cellValue < 1) return null;
    // Convert from Sheets/Excel serial date to a JavaScript Date.
    const epochMilliseconds = (cellValue - SHEETS_EPOCH_OFFSET) * 86400000;
    const converted = new Date(epochMilliseconds);
    return isNaN(converted.getTime()) ? null : converted;
  }

  if (typeof cellValue === 'string') {
    const parsed = new Date(cellValue);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  return null;
}

/**
 * Formats a Date into a "yyyy-MM-dd" string in the given time zone.
 *
 * This is used to compare whether two dates (the row's call date and the
 * window's end date) fall on the same calendar day. A simple string comparison
 * of "yyyy-MM-dd" keys is reliable across time zones because Utilities.formatDate
 * performs the conversion correctly, whereas comparing raw epoch milliseconds
 * would produce wrong results near midnight.
 *
 * @param {Date}   date     — The date to format.
 * @param {string} timeZone — The time zone for the conversion (e.g., "America/Los_Angeles").
 * @returns {string} A date key in "yyyy-MM-dd" format, e.g. "2025-10-19".
 */
function getLocalDateKey_(date, timeZone) {
  return Utilities.formatDate(
    date,
    timeZone || Session.getScriptTimeZone() || 'Etc/GMT',
    'yyyy-MM-dd'
  );
}
