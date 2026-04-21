/**
 * performance.gs — Performance monitoring and profiling utilities for COMET.
 * VERSION: 0.1.0
 *
 * Provides helpers for:
 *   - Execution time logging
 *   - Batch sheet read/write operations
 *   - Performance-aware caching
 *
 * Usage:
 *   const {result, elapsed} = logExecutionTime_('Label', () => expensiveFunction());
 *   console.log('Took ' + elapsed + 'ms');
 */


// ---------------------------------------------------------------------------
// Execution Time Logging
// ---------------------------------------------------------------------------

/**
 * Wraps a function, measures its execution time, logs to console, and returns the result.
 *
 * @param {string}   label    — Human-readable label for the operation (e.g., "Phase 1")
 * @param {Function} fn       — The function to time
 * @returns {{result: any, elapsed: number}} — {result, elapsed in milliseconds}
 */
function logExecutionTime_(label, fn) {
  const start = Date.now();
  let result;
  try {
    result = fn();
  } catch (error) {
    const elapsed = Date.now() - start;
    console.error('[PERF] ' + label + ' FAILED: ' + error.message + ' (' + elapsed + 'ms)');
    throw error;
  }
  const elapsed = Date.now() - start;
  console.log('[PERF] ' + label + ': ' + elapsed + 'ms');
  return { result, elapsed };
}


// ---------------------------------------------------------------------------
// Batch Sheet Operations
// ---------------------------------------------------------------------------

/**
 * Reads a rectangular range from a sheet in one API call.
 *
 * @param {Sheet}  sheet    — The sheet to read from
 * @param {number} startRow — 1-indexed row number (1 = first row)
 * @param {number} numRows  — Number of rows to read
 * @param {number} numCols  — Number of columns to read (starting from column A)
 * @returns {Array<Array>}  — 2D array of cell values
 */
function batchSheetRead_(sheet, startRow, numRows, numCols) {
  return sheet.getRange(startRow, 1, numRows, numCols).getValues();
}

/**
 * Writes a 2D array to a sheet in one API call.
 *
 * @param {Sheet}       sheet    — The sheet to write to
 * @param {number}      startRow — 1-indexed row number
 * @param {Array<Array>} data    — 2D array of values to write
 */
function batchSheetWrite_(sheet, startRow, data) {
  const numRows = data.length;
  const numCols = numRows > 0 ? data[0].length : 1;
  const range = sheet.getRange(startRow, 1, numRows, numCols);
  range.setValues(data);
}

/**
 * Writes a 2D array of colors to a sheet in one API call.
 *
 * @param {Sheet}       sheet    — The sheet to write to
 * @param {number}      startRow — 1-indexed row number
 * @param {Array<Array>} colors  — 2D array of color hex strings (e.g., "#FF0000")
 */
function batchSheetColorize_(sheet, startRow, colors) {
  const numRows = colors.length;
  const numCols = numRows > 0 ? colors[0].length : 1;
  const range = sheet.getRange(startRow, 1, numRows, numCols);
  range.setBackgrounds(colors);
}


// ---------------------------------------------------------------------------
// Performance-Aware Caching
// ---------------------------------------------------------------------------

/**
 * Simple cache using PropertiesService with TTL support.
 *
 * Usage:
 *   setCacheValue_('myKey', expensiveData, 30); // 30-minute TTL
 *   const data = getCacheValue_('myKey');
 *
 * @param {string} key       — Cache key
 * @param {*}      value     — Value to cache (will be JSON serialized)
 * @param {number} ttlMinutes — Time-to-live in minutes (default 30)
 */
function setCacheValue_(key, value, ttlMinutes) {
  ttlMinutes = ttlMinutes || 30;
  const properties = PropertiesService.getUserProperties();
  const cacheEntry = {
    value: value,
    expiresAt: Date.now() + (ttlMinutes * 60 * 1000),
  };
  properties.setProperty(key, JSON.stringify(cacheEntry));
}

/**
 * Retrieves a cached value if it exists and hasn't expired.
 *
 * @param {string} key — Cache key
 * @returns {*|null}   — Cached value, or null if not found or expired
 */
function getCacheValue_(key) {
  const properties = PropertiesService.getUserProperties();
  const stored = properties.getProperty(key);
  if (!stored) return null;

  let cacheEntry;
  try {
    cacheEntry = JSON.parse(stored);
  } catch (error) {
    // Corrupted cache entry — delete it
    properties.deleteProperty(key);
    return null;
  }

  if (Date.now() > cacheEntry.expiresAt) {
    // Expired — delete and return null
    properties.deleteProperty(key);
    return null;
  }

  return cacheEntry.value;
}

/**
 * Clears a single cache entry.
 *
 * @param {string} key — Cache key
 */
function clearCacheValue_(key) {
  PropertiesService.getUserProperties().deleteProperty(key);
}

/**
 * Clears all COMET cache entries (those starting with "COMET_").
 */
function clearAllCache_() {
  const properties = PropertiesService.getUserProperties();
  const keys = properties.getKeys();
  keys.forEach(function(key) {
    if (key.indexOf('COMET_') === 0) {
      properties.deleteProperty(key);
    }
  });
}


// ---------------------------------------------------------------------------
// Profiling Configuration
// ---------------------------------------------------------------------------

/**
 * Returns true if profiling is enabled (from config.js PROFILING_ENABLED flag).
 * If PROFILING_ENABLED is not defined, returns false.
 *
 * @returns {boolean}
 */
function isProfilingEnabled_() {
  return (typeof PROFILING_ENABLED !== 'undefined') && PROFILING_ENABLED;
}

/**
 * Conditional logging — only logs if profiling is enabled.
 *
 * @param {string} message — Log message
 */
function profileLog_(message) {
  if (isProfilingEnabled_()) {
    console.log('[PROFILE] ' + message);
  }
}
