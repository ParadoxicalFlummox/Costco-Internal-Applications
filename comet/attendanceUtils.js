/**
 * attendanceUtils.js — Shared utilities for attendance controller parsing and generation.
 * VERSION: 0.1.0
 *
 * This module contains pure utility functions used by both the legacy attendance grid
 * scanner (for migration tools) and the new JSON-native reader. Functions are geometry-agnostic
 * and focus on normalization, lookup, and date/column conversions.
 */


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
 * Matching is case-insensitive. Returns null for blank or unrecognized strings.
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
// Year Extraction
// ---------------------------------------------------------------------------

/**
 * Extracts a four-digit calendar year from the year title string.
 *
 * e.g. "2026 Attendance Controller" → 2026
 * Returns null if no four-digit year is found.
 *
 * @param {string} titleString — The raw string to parse.
 * @returns {number|null}
 */
function parseYearFromTitle_(titleString) {
  if (!titleString) return null;
  const match = String(titleString).match(/\b(19|20)\d{2}\b/);
  return match ? parseInt(match[0], 10) : null;
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
