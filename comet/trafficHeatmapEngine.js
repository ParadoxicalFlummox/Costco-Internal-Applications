/**
 * trafficHeatmapEngine.js — Traffic heatmap classification, stagger scheduling, and pool management.
 * VERSION: 0.5.1
 *
 * This module owns all logic related to the unified traffic heatmap system:
 *   1. Traffic classification (Low / Moderate / High based on door counts)
 *   2. Peak window identification from traffic curves
 *   3. Stagger map pre-computation (assigns specific start times per shift per day)
 *   4. Pool member partitioning and selection
 *   5. Pool member scheduling with staggered start times
 *
 * Dependency: config.js, settingsManager.js, scheduleEngine.js (for grid/employee structures)
 */


// ---------------------------------------------------------------------------
// Traffic Classification
// ---------------------------------------------------------------------------

/**
 * Classifies each day into Low / Moderate / High traffic level based on the heatmap config.
 *
 * Algorithm:
 *   1. For each hour in the traffic curve, classify as Low/Moderate/High per thresholds
 *   2. Count hours in each category
 *   3. Day level = dominant category (most hours)
 *
 * @param {Object} heatmapConfig — { thresholds, trafficCurves, weeklyOverrides }
 * @param {Date} weekStartDate — Monday of the week (for weekly override lookup)
 * @returns {Object} { Monday: "Low", Tuesday: "Moderate", ... }
 */
function classifyDayTrafficLevels_(heatmapConfig, weekStartDate) {
  const thresholds = heatmapConfig.thresholds || HEATMAP_DEFAULTS.thresholds;
  const lowThreshold = thresholds.low || 100;
  const highThreshold = thresholds.high || 300;

  // Resolve curves: baseline + weekly override
  const curves = resolveTrafficCurves_(heatmapConfig, weekStartDate);

  const dayLevels = {};

  DAY_NAMES_IN_ORDER.forEach(function(dayName) {
    const curve = curves[dayName] || [];
    if (curve.length === 0) {
      dayLevels[dayName] = 'Low';
      return;
    }

    // Classify each hour
    let lowCount = 0, moderateCount = 0, highCount = 0;
    curve.forEach(function(doorCount) {
      if (doorCount >= highThreshold) {
        highCount++;
      } else if (doorCount >= lowThreshold) {
        moderateCount++;
      } else {
        lowCount++;
      }
    });

    // Determine dominant level
    if (highCount > moderateCount && highCount > lowCount) {
      dayLevels[dayName] = 'High';
    } else if (moderateCount > lowCount) {
      dayLevels[dayName] = 'Moderate';
    } else {
      dayLevels[dayName] = 'Low';
    }
  });

  return dayLevels;
}

/**
 * Identifies the pre-peak window for each day (the contiguous block before/at peak traffic).
 *
 * Algorithm:
 *   1. Find the highest door count in the day's curve
 *   2. Extend backward to include hours above the low threshold
 *   3. Return { startHour, endHour } for the pre-peak window
 *
 * @param {Object} heatmapConfig
 * @param {Date} weekStartDate
 * @returns {Object} { Monday: { startHour, endHour }, ... }
 */
function identifyPrePeakWindows_(heatmapConfig, weekStartDate) {
  const thresholds = heatmapConfig.thresholds || HEATMAP_DEFAULTS.thresholds;
  const lowThreshold = thresholds.low || 100;

  const curves = resolveTrafficCurves_(heatmapConfig, weekStartDate);
  const peakWindows = {};

  DAY_NAMES_IN_ORDER.forEach(function(dayName) {
    const curve = curves[dayName] || [];
    if (curve.length === 0) {
      peakWindows[dayName] = { startHour: 8, endHour: 17 };  // Default 8 AM – 5 PM
      return;
    }

    // Find the hour with maximum traffic
    let maxIdx = 0, maxValue = 0;
    curve.forEach(function(doorCount, idx) {
      if (doorCount > maxValue) {
        maxValue = doorCount;
        maxIdx = idx;
      }
    });

    // Extend backward from max to find contiguous block above low threshold
    let startHour = maxIdx;
    while (startHour > 0 && curve[startHour - 1] >= lowThreshold) {
      startHour--;
    }

    // Extend forward from max
    let endHour = maxIdx;
    while (endHour < curve.length - 1 && curve[endHour + 1] >= lowThreshold) {
      endHour++;
    }

    peakWindows[dayName] = {
      startHour: startHour,
      endHour: endHour + 1,  // Exclusive end
    };
  });

  return peakWindows;
}

/**
 * Resolves traffic curves: baseline + weekly override for the given week.
 *
 * @param {Object} heatmapConfig
 * @param {Date} weekStartDate
 * @returns {Object} { Monday: [...], Tuesday: [...], ... }
 */
function resolveTrafficCurves_(heatmapConfig, weekStartDate) {
  const baseline = heatmapConfig.trafficCurves || HEATMAP_DEFAULTS.defaultTrafficCurves;
  const weekKey = Utilities.formatDate(weekStartDate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const weekOverride = (heatmapConfig.weeklyOverrides || {})[weekKey] || {};

  const resolved = {};
  DAY_NAMES_IN_ORDER.forEach(function(dayName) {
    resolved[dayName] = weekOverride[dayName] || baseline[dayName] || [];
  });

  return resolved;
}


// ---------------------------------------------------------------------------
// Stagger Map Pre-Computation (Core Algorithm)
// ---------------------------------------------------------------------------

/**
 * Builds the stagger map: a pre-computed list of start times per shift type per day.
 *
 * Returns a plain object: { dayName: { "ShiftName|Status": [startTime, startTime, ...], ... }, ... }
 *
 * Algorithm:
 *   1. For each shift type with flexEnabled=true, determine its flex window and day-specific constraints
 *   2. Generate stagger positions within the window at the configured increment (15 or 30 min)
 *   3. Bias toward early on High traffic days, spread evenly on Low days
 *   4. Respect hard constraints (opener=4 AM fixed, closer <= store closing time)
 *   5. Return a map where each shift has a list of available start times
 *
 * @param {Object} shiftTimingMap — { "ShiftName|Status": ShiftDefinition, ... }
 * @param {Object} dayTrafficLevels — { Monday: "Low", ... }
 * @param {Date} weekStartDate — For determining day-of-week (Sat/Sun closer constraints)
 * @returns {Object} { Monday: { "ShiftName|Status": [startTimes...], ... }, ... }
 */
function buildStaggeredStartMap_(shiftTimingMap, dayTrafficLevels, weekStartDate) {
  const staggerIncrement = HEATMAP_DEFAULTS.staggerIncrement || 15;  // minutes
  const staggerMap = {};

  // For each day
  DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
    const trafficLevel = dayTrafficLevels[dayName] || 'Moderate';
    staggerMap[dayName] = {};

    // For each shift type
    Object.keys(shiftTimingMap).forEach(function(shiftKey) {
      const shiftDef = shiftTimingMap[shiftKey];

      // Check if flex is enabled for this shift
      if (shiftDef.flexEnabled === false) {
        // Fixed shift: use the day-specific anchor start time (sat/sun may differ)
        const fixedStart = getStartMinutesForDay_(shiftDef, dayName); // settingsManager.js
        staggerMap[dayName][shiftKey] = [formatMinutesAsTimeString(fixedStart)];
        return;
      }

      // Flex-enabled shift: compute stagger positions within flex window
      const startTimes = computeStaggerPositions_(shiftDef, trafficLevel, dayName, dayIndex, weekStartDate, staggerIncrement);
      staggerMap[dayName][shiftKey] = startTimes;
    });
  });

  return staggerMap;
}

/**
 * Computes staggered start times for a flex-enabled shift on a given day.
 *
 * @param {Object} shiftDef — { name, startMinutes, flexWindowEarliest, flexWindowLatest, endMinutes, ... }
 * @param {string} trafficLevel — "Low", "Moderate", or "High"
 * @param {string} dayName — e.g., "Friday"
 * @param {number} dayIndex — 0=Monday, 6=Sunday (for day-of-week closing time checks)
 * @param {Date} weekStartDate
 * @param {number} staggerIncrement — minutes between stagger positions (15 or 30)
 * @returns {Array<string>} ["HH:MM", "HH:MM", ...]
 */
function computeStaggerPositions_(shiftDef, trafficLevel, dayName, dayIndex, weekStartDate, staggerIncrement) {
  // Use the day-specific anchor for fallback calculations
  const dayAnchorMinutes = getStartMinutesForDay_(shiftDef, dayName); // settingsManager.js

  // Parse flex window from shift definition
  let flexEarliestMin = timeStringToMinutes_(shiftDef.flexWindowEarliest);
  let flexLatestMin = timeStringToMinutes_(shiftDef.flexWindowLatest);

  // Handle missing flex window (should not happen if flexEnabled=true, but be defensive)
  if (!flexEarliestMin || !flexLatestMin || flexLatestMin <= flexEarliestMin) {
    // Fallback: day anchor +/- 30 minutes
    flexEarliestMin = Math.max(240, dayAnchorMinutes - 30);  // 240 = 4 AM (earliest safe)
    flexLatestMin = dayAnchorMinutes + 30;
  }

  // Apply hard constraints for openers and closers
  if (shiftDef.name && shiftDef.name.toLowerCase().includes('open')) {
    // Opener: fixed at 4:00 AM (240 minutes)
    flexEarliestMin = 240;
    flexLatestMin = 240;
  } else if (shiftDef.name && shiftDef.name.toLowerCase().includes('clos')) {
    // Closer: cap latest start time so shift ends by store closing time
    const closingTime = getStoreClosingTimeMinutes_(dayIndex);
    const maxStartMin = closingTime - shiftDef.blockMinutes;  // blockMinutes = paidHours*60 + (hasLunch?30:0)
    flexLatestMin = Math.min(flexLatestMin, maxStartMin);
  }

  // Generate stagger positions within [flexEarliest, flexLatest]
  const positions = [];
  for (let min = flexEarliestMin; min <= flexLatestMin; min += staggerIncrement) {
    positions.push(min);
  }

  if (positions.length === 0) {
    // Fallback: just use the earliest time
    positions.push(flexEarliestMin);
  }

  // Bias the positions based on traffic level
  const biased = biasStaggerPositions_(positions, trafficLevel);

  // Convert to "HH:MM" strings
  return biased.map(function(minSinceMidnight) {
    return formatMinutesAsTimeString(minSinceMidnight);
  });
}

/**
 * Biases a list of stagger positions (minutes) toward early or late based on traffic level.
 *
 * Algorithm:
 *   - Low: spread evenly across the list
 *   - Moderate: duplicate some early positions
 *   - High: heavily weight early positions
 *
 * @param {Array<number>} positions — e.g., [480, 495, 510, 525, 540] (minutes)
 * @param {string} trafficLevel — "Low", "Moderate", or "High"
 * @returns {Array<number>} biased positions (may have duplicates)
 */
function biasStaggerPositions_(positions, trafficLevel) {
  if (positions.length <= 1) return positions;

  if (trafficLevel === 'Low') {
    // Return as-is: spread evenly
    return positions;
  } else if (trafficLevel === 'Moderate') {
    // Add one extra early position
    const biased = [positions[0]].concat(positions);
    return biased;
  } else {  // High
    // Add multiple early positions
    const biased = [positions[0], positions[0], positions[0]].concat(positions);
    return biased;
  }
}

/**
 * Returns the store closing time (in minutes since midnight) for a given day of week.
 *
 * @param {number} dayIndex — 0=Monday, 1=Tuesday, ..., 6=Sunday
 * @returns {number} minutes since midnight
 */
function getStoreClosingTimeMinutes_(dayIndex) {
  const closingTime = dayIndex === 5 ? STORE_CLOSING_TIMES.saturday :
                     dayIndex === 6 ? STORE_CLOSING_TIMES.sunday :
                     STORE_CLOSING_TIMES.weekday;
  return timeStringToMinutes_(closingTime);
}


// ---------------------------------------------------------------------------
// Pool Member Partitioning & Selection
// ---------------------------------------------------------------------------

/**
 * Partitions employees into pool members and regular employees.
 *
 * Pool members are those listed in heatmapConfig.pool (tier 1 + tier 2).
 * Tier 2 members are cross-trained employees from other departments with
 * redirectable secondary time.
 *
 * @param {Array} employeeList
 * @param {Object} heatmapConfig
 * @returns {{ poolMembers: Array, regularEmployees: Array }}
 */
function partitionPoolMembers_(employeeList, heatmapConfig) {
  const poolConfig = heatmapConfig.pool || {};
  const tier1Ids = new Set(poolConfig.tier1EmployeeIds || []);
  const tier2Ids = new Set(poolConfig.tier2EmployeeIds || []);

  const poolMembers = [];
  const regularEmployees = [];

  employeeList.forEach(function(emp) {
    const empId = emp.id || '';
    if (tier1Ids.has(empId) || tier2Ids.has(empId)) {
      poolMembers.push(emp);
    } else {
      regularEmployees.push(emp);
    }
  });

  return { poolMembers: poolMembers, regularEmployees: regularEmployees };
}

/**
 * Selects which pool members should be scheduled for a given day based on traffic level.
 *
 * Selection is by seniority: highest seniority first.
 *
 * @param {Array} poolMembers
 * @param {string} trafficLevel — "Low", "Moderate", or "High"
 * @param {Object} heatmapConfig
 * @returns {Array} selected pool members (subset)
 */
function selectPoolMembers_(poolMembers, trafficLevel, heatmapConfig) {
  const poolConfig = heatmapConfig.pool || {};
  const schedulingCounts = poolConfig.schedulingCounts || HEATMAP_DEFAULTS.poolSchedulingCounts;
  const targetCount = schedulingCounts[trafficLevel] || 0;

  // Sort by seniority rank (lower rank = higher seniority)
  const sorted = poolMembers.slice().sort(function(a, b) {
    return (a.seniorityRank || 999) - (b.seniorityRank || 999);
  });

  return sorted.slice(0, targetCount);
}


// ---------------------------------------------------------------------------
// Pool Member Scheduling
// ---------------------------------------------------------------------------

/**
 * Schedules selected pool members into the week grid with staggered start times.
 *
 * Pool members appear in their own section above regular employees in the final sheet.
 * This function assigns them to shifts respecting their constraints.
 *
 * @param {Array} selectedPoolMembers
 * @param {Array} weekGrid
 * @param {Object} staggerMap — { dayName: { "ShiftName|Status": [startTimes...], ... }, ... }
 * @param {Object} shiftTimingMap
 * @param {Object} dayTrafficLevels
 * @returns {void} mutates weekGrid
 */
function schedulePoolMembers_(selectedPoolMembers, weekGrid, staggerMap, shiftTimingMap, dayTrafficLevels) {
  // Placeholder for full pool scheduling logic
  // This will assign pool members to shifts, respecting their primary role
  // and using staggered start times from the stagger map

  // For now, log that pool scheduling was called; full implementation follows
  console.log('schedulePoolMembers_: ' + selectedPoolMembers.length + ' pool members to schedule');
}


// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Converts minutes since midnight to a 12-hour "h:mm AM/PM" string.
 * Reuses the function from settingsManager.js.
 *
 * @param {number} totalMinutes
 * @returns {string}
 */
function formatMinutesAsTimeString(totalMinutes) {
  const totalHours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  const period = totalHours >= 12 ? 'PM' : 'AM';
  const twelve = totalHours % 12 === 0 ? 12 : totalHours % 12;
  return twelve + ':' + String(minutes).padStart(2, '0') + ' ' + period;
}

/**
 * Converts "HH:MM" string to minutes since midnight.
 * Reuses the function from scheduleSettings.js.
 *
 * @param {string} timeString
 * @returns {number}
 */
function timeStringToMinutes_(timeString) {
  if (!timeString) return 0;
  const parts = String(timeString).split(':');
  if (parts.length < 2) return 0;
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}
