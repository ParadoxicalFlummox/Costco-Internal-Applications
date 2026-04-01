/**
 *       _____             _                                _            _____        _                _         _
 *      / ____|           | |                  /\          | |          / ____|      | |              | |       | |
 *     | |      ___   ___ | |_  ___  ___      /  \   _   _ | |_  ___   | (___    ___ | |__    ___   __| | _   _ | |  ___  _ __
 *     | |     / _ \ / __|| __|/ __|/ _ \    / /\ \ | | | || __|/ _ \   \___ \  / __|| '_ \  / _ \ / _` || | | || | / _ \| '__|
 *     | |____| (_) |\__ \| |_| (__| (_) |  / ____ \| |_| || |_| (_) |  ____) || (__ | | | ||  __/| (_| || |_| || ||  __/| |
 *      \_____|\___/ |___/ \__|\___|\___/  /_/    \_\\__,_| \__|\___/  |_____/  \___||_| |_| \___| \__,_| \__,_||_| \___||_|
 *       _____        _                _         _          _____                ______                _
 *      / ____|      | |              | |       | |        / ____|              |  ____|              (_)
 *     | (___    ___ | |__    ___   __| | _   _ | |  ___  | |  __   ___  _ __   | |__    _ __    __ _  _  _ __    ___
 *      \___ \  / __|| '_ \  / _ \ / _` || | | || | / _ \ | | |_ | / _ \| '_ \  |  __|  | '_ \  / _` || || '_ \  / _ \
 *      ____) || (__ | | | ||  __/| (_| || |_| || ||  __/ | |__| ||  __/| | | | | |____ | | | || (_| || || | | ||  __/
 *     |_____/  \___||_| |_| \___| \__,_| \__,_||_| \___|  \_____| \___||_| |_| |______||_| |_| \__, ||_||_| |_| \___|
 *                                                                                               __/ |
 *                                                                                              |___/
 * Built by: Adam Roy
 * Branch: shift-window-with-minimums
 * Version 0.2.0
 */

// ─── Orchestration ───────────────────────────────────────────────────────────

/**
 * Entry point. Creates a new Weekly Schedule, auto-generating the template sheet if needed.
 */
function generateWeeklySchedule() {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    ensureTemplateExists(activeSpreadsheet);

    const configSheet   = activeSpreadsheet.getSheetByName(CONFIGURATION_SHEET_NAME);
    const templateSheet = activeSpreadsheet.getSheetByName(TEMPLATE_SHEET_NAME);
    const weeklyLabel   = generateWeeklyDateLabel();

    const scheduleSheet = activeSpreadsheet.getSheetByName(weeklyLabel) ||
        templateSheet.copyTo(activeSpreadsheet).setName(weeklyLabel);
    scheduleSheet.showSheet();

    const rosterData = getRosterSortedBySeniority(configSheet);
    populateEmployeeRows(scheduleSheet, rosterData);

    const lastEmployeeRow = 6 + (rosterData.length * 3) - 1;
    attachStaffingSummary(scheduleSheet, lastEmployeeRow);
    applyScheduleFormatting(scheduleSheet, rosterData.length);
    resolveEntireWeek(scheduleSheet);

    activeSpreadsheet.setActiveSheet(scheduleSheet);
}

// ─── Template Bootstrap ───────────────────────────────────────────────────────

/**
 * Creates the Grid Template sheet if it does not already exist.
 * Sets up headers, column widths, and formatting so copies are ready to use.
 */
function ensureTemplateExists(activeSpreadsheet) {
    if (activeSpreadsheet.getSheetByName(TEMPLATE_SHEET_NAME)) return;

    const templateSheet = activeSpreadsheet.insertSheet(TEMPLATE_SHEET_NAME);
    const weekDayNames  = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

    templateSheet.getRange(5, 2).setValue("Employee");
    weekDayNames.forEach((dayName, dayColumnOffset) => templateSheet.getRange(5, 3 + dayColumnOffset).setValue(dayName));
    templateSheet.getRange(5, 10).setValue("Total Hrs");

    templateSheet.getRange(5, 1, 1, 10)
        .setBackground("#263238")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    templateSheet.setRowHeight(5, 30);

    templateSheet.setColumnWidth(1, 55);
    templateSheet.setColumnWidth(2, 160);
    for (let columnNumber = 3; columnNumber <= 9; columnNumber++) templateSheet.setColumnWidth(columnNumber, 95);
    templateSheet.setColumnWidth(10, 85);

    templateSheet.hideSheet();
    activeSpreadsheet.toast("Grid Template sheet was created automatically.", "First Run");
}

// ─── Sheet Population ─────────────────────────────────────────────────────────

/**
 * Reads the CONFIG sheet and returns employees sorted by seniority (highest first).
 */
function getRosterSortedBySeniority(configSheet) {
    const rosterData = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 9).getValues();
    return rosterData.sort((employeeA, employeeB) => employeeB[COLUMN_INDEX_SENIORITY_RANK] - employeeA[COLUMN_INDEX_SENIORITY_RANK]);
}

/**
 * Writes the VAC / RDO / SHIFT row structure for each employee onto the schedule sheet.
 */
function populateEmployeeRows(scheduleSheet, rosterData) {
    const weekDayNames = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
    let currentRow = 6;

    rosterData.forEach(employeeRecord => {
        scheduleSheet.getRange(currentRow, 1, 3, 1).setValues([["VAC"], ["RDO"], ["SHIFT"]]);
        scheduleSheet.getRange(currentRow, 2).setValue(employeeRecord[COLUMN_INDEX_NAME]);
        scheduleSheet.getRange(currentRow, 2, 3, 1).merge().setVerticalAlignment("middle");
        scheduleSheet.getRange(currentRow, 3, 2, 7).insertCheckboxes();

        [employeeRecord[COLUMN_INDEX_PREFERENCE_ONE], employeeRecord[COLUMN_INDEX_PREFERENCE_TWO]].forEach(preferredDay => {
            const dayIndex = weekDayNames.indexOf(preferredDay);
            if (dayIndex !== -1) scheduleSheet.getRange(currentRow + 1, 3 + dayIndex).setValue(true);
        });

        currentRow += 3;
    });
}

/**
 * Adds the Staffing Summary block (REQUIRED / ACTUAL / STATUS) below the employee rows.
 */
function attachStaffingSummary(scheduleSheet, lastEmployeeRow) {
    const summaryHeaderRow = lastEmployeeRow + 2;
    const weekDayNames     = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

    scheduleSheet.getRange(summaryHeaderRow, 2, 3, 8).clearContent();
    scheduleSheet.getRange(summaryHeaderRow, 2, 3, 1)
        .setValues([["REQUIRED"], ["ACTUAL"], ["STATUS"]])
        .setFontWeight("bold");

    for (let dayIndex = 0; dayIndex < 7; dayIndex++) {
        const columnNumber    = 3 + dayIndex;
        const columnLetter    = String.fromCharCode(67 + dayIndex);
        const minimumRequired = getMinimumStaffRequiredForDay(weekDayNames[dayIndex]);

        scheduleSheet.getRange(summaryHeaderRow, columnNumber).setValue(minimumRequired);

        // Count cells containing ":" — only time-formatted shift strings match, not OFF/VAC/NO SHIFT
        const actualStaffFormula = `=COUNTIF(${columnLetter}6:${columnLetter}${lastEmployeeRow},"*:*")`;
        scheduleSheet.getRange(summaryHeaderRow + 1, columnNumber)
            .setFormula(actualStaffFormula)
            .setFontWeight("bold")
            .setHorizontalAlignment("center");

        const statusFormula = `=IF(${columnLetter}${summaryHeaderRow + 1}>=${columnLetter}${summaryHeaderRow}, "OK", "UNDER")`;
        scheduleSheet.getRange(summaryHeaderRow + 2, columnNumber)
            .setFormula(statusFormula)
            .setHorizontalAlignment("center");
    }
}

// ─── Formatting ───────────────────────────────────────────────────────────────

/**
 * Applies visual formatting: colors, borders, freeze panes, and column widths.
 */
function applyScheduleFormatting(scheduleSheet, employeeCount) {
    scheduleSheet.setFrozenRows(5);
    scheduleSheet.setFrozenColumns(2);

    for (let employeeIndex = 0; employeeIndex < employeeCount; employeeIndex++) {
        const baseRow            = 6 + (employeeIndex * 3);
        const rowBackgroundColor = (employeeIndex % 2 === 0) ? "#FFFFFF" : "#F5F7F8";

        scheduleSheet.getRange(baseRow, 2, 3, 9).setBackground(rowBackgroundColor);

        scheduleSheet.getRange(baseRow, 1)
            .setBackground("#FFCDD2").setFontWeight("bold").setHorizontalAlignment("center");
        scheduleSheet.getRange(baseRow + 1, 1)
            .setBackground("#FFF9C4").setFontWeight("bold").setHorizontalAlignment("center");
        scheduleSheet.getRange(baseRow + 2, 1)
            .setBackground("#C8E6C9").setFontWeight("bold").setHorizontalAlignment("center");

        scheduleSheet.getRange(baseRow, 2).setFontWeight("bold");

        scheduleSheet.getRange(baseRow + 2, 1, 1, 10)
            .setBorder(null, null, true, null, null, null, "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

        scheduleSheet.getRange(baseRow + 2, 10)
            .setBackground("#E8F5E9").setHorizontalAlignment("center").setFontWeight("bold");
    }

    applyConditionalFormattingToSummary(scheduleSheet, employeeCount);
}

/**
 * Adds OK/UNDER conditional formatting rules to the STATUS row in the summary block.
 */
function applyConditionalFormattingToSummary(scheduleSheet, employeeCount) {
    const statusRowNumber = 6 + (employeeCount * 3) + 3; // 2-row gap + 3rd summary row
    const statusRange     = scheduleSheet.getRange(statusRowNumber, 3, 1, 7);

    const okRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("OK")
        .setBackground("#B7E1CD")
        .setRanges([statusRange])
        .build();

    const underRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("UNDER")
        .setBackground("#F4C7C3")
        .setRanges([statusRange])
        .build();

    const existingRules = scheduleSheet.getConditionalFormatRules();
    scheduleSheet.setConditionalFormatRules([...existingRules, okRule, underRule]);
}

// ─── Schedule Resolution ──────────────────────────────────────────────────────

/**
 * Orchestrates full schedule resolution across 4 phases.
 * Called on initial generation and on every checkbox edit via onEdit.
 */
function resolveEntireWeek(scheduleSheet) {
    const lastRow       = scheduleSheet.getLastRow();
    const allNameRows   = scheduleSheet.getRange("B6:B" + lastRow).getValues();
    const employeeNames = allNameRows
        .filter(nameRow => nameRow[0] !== "" && !["REQUIRED", "ACTUAL", "STATUS"].includes(nameRow[0].toString().toUpperCase()))
        .map(nameRow => nameRow[0]);

    if (employeeNames.length === 0) return;

    const gridRange      = scheduleSheet.getRange(6, 3, employeeNames.length * 3, 8);
    const gridValues     = gridRange.getValues();
    const rosterMap      = loadRosterMapFromConfig();
    const shiftTimingMap = buildShiftTimingMap(); // Phase 0: build timing map once for all phases

    assignShiftsByPreferenceAndSeniority(gridValues, employeeNames, rosterMap, shiftTimingMap); // Phase 1
    enforceMinimumWeeklyHours(gridValues, employeeNames, rosterMap, shiftTimingMap);            // Phase 2
    detectAndResolveCoverageGaps(gridValues, employeeNames, rosterMap, shiftTimingMap);          // Phase 3

    // Write SHIFT rows only — never overwrite VAC or RDO checkboxes set by the manager
    for (let employeeIndex = 0; employeeIndex < employeeNames.length; employeeIndex++) {
        const shiftRowInSheet = 6 + (employeeIndex * 3) + 2;
        scheduleSheet.getRange(shiftRowInSheet, 3, 1, 8).setValues([gridValues[employeeIndex * 3 + 2]]);
    }
}

/**
 * Builds a lookup map from CONFIG sheet: employee name → { status, pref, qualifiedShifts }.
 * qualifiedShifts is parsed from column I (comma-separated). Defaults to [pref] if blank.
 */
function loadRosterMapFromConfig() {
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIGURATION_SHEET_NAME);
    const rosterRows  = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 9).getValues();
    const rosterMap   = {};

    rosterRows.forEach(employeeRow => {
        if (!employeeRow[COLUMN_INDEX_NAME]) return;

        const rawQualifiedShifts = employeeRow[COLUMN_INDEX_QUALIFIED_SHIFTS].toString().trim();
        const qualifiedShifts    = rawQualifiedShifts
            ? rawQualifiedShifts.split(",").map(shiftName => shiftName.trim()).filter(shiftName => shiftName !== "")
            : [employeeRow[COLUMN_INDEX_SHIFT_PREFERENCE]];

        rosterMap[employeeRow[COLUMN_INDEX_NAME]] = {
            status:          employeeRow[COLUMN_INDEX_EMPLOYMENT_STATUS],
            pref:            employeeRow[COLUMN_INDEX_SHIFT_PREFERENCE],
            qualifiedShifts: qualifiedShifts
        };
    });

    return rosterMap;
}

// ─── Phase 1 ──────────────────────────────────────────────────────────────────

/**
 * Assigns shifts by seniority order, granting RDO/VAC requests within staffing floors.
 * Uses coverage-aware shift selection so flexible employees fill gaps left by senior staff.
 * Mutates gridValues in place. Never writes to VAC or RDO rows.
 */
function assignShiftsByPreferenceAndSeniority(gridValues, employeeNames, rosterMap, shiftTimingMap) {
    const weekDayNames   = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
    const totalEmployees = employeeNames.length;

    for (let dayIndex = 0; dayIndex < 7; dayIndex++) {
        const maximumAllowedOff = totalEmployees - getMinimumStaffRequiredForDay(weekDayNames[dayIndex]);
        let currentOffCount     = 0;
        const dayCoverageSlots  = new Array(39).fill(0); // Running 30-min coverage map for this day

        for (let employeeIndex = 0; employeeIndex < totalEmployees; employeeIndex++) {
            const employeeGridOffset = employeeIndex * 3;
            const employeeSettings   = rosterMap[employeeNames[employeeIndex]] ||
                { status: "PT", pref: "Morning", qualifiedShifts: ["Morning"] };

            if (dayIndex === 0) gridValues[employeeGridOffset + 2][7] = 0; // Initialize total hours on Monday pass

            const isOnVacation  = gridValues[employeeGridOffset][dayIndex];
            const hasRdoRequest = gridValues[employeeGridOffset + 1][dayIndex];

            if (isOnVacation === true) {
                currentOffCount++;
                gridValues[employeeGridOffset + 2][dayIndex] = "VAC";

            } else if (hasRdoRequest === true && currentOffCount < maximumAllowedOff) {
                currentOffCount++;
                gridValues[employeeGridOffset + 2][dayIndex] = "OFF";

            } else {
                const assignedShift    = selectBestCoverageShift(employeeSettings.qualifiedShifts, employeeSettings.status, dayCoverageSlots, shiftTimingMap);
                const maximumHours     = (employeeSettings.status === "FT") ? FT_MAXIMUM_WEEKLY_HOURS : PT_MAXIMUM_WEEKLY_HOURS;
                const accumulatedHours = parseFloat(gridValues[employeeGridOffset + 2][7]) || 0;
                const wouldExceedCap   = (accumulatedHours + assignedShift.hours) > maximumHours;

                if (wouldExceedCap && currentOffCount < maximumAllowedOff) {
                    currentOffCount++;
                    gridValues[employeeGridOffset + 2][dayIndex] = "OFF";
                } else {
                    gridValues[employeeGridOffset + 2][dayIndex] = assignedShift.text;
                    gridValues[employeeGridOffset + 2][7]        = accumulatedHours + assignedShift.hours;

                    // Update the running coverage map so lower-seniority employees see the current state
                    const startSlot = Math.max(0, Math.floor((assignedShift.startMinutes - 240) / 30));
                    const endSlot   = Math.min(39, Math.floor((assignedShift.endMinutes - 240) / 30));
                    for (let slotIndex = startSlot; slotIndex < endSlot; slotIndex++) {
                        dayCoverageSlots[slotIndex]++;
                    }
                }
            }
        }
    }
}

// ─── Phase 2 ──────────────────────────────────────────────────────────────────

/**
 * Converts OFF days to working shifts for employees below their weekly hour minimum.
 * Uses the employee's preference shift for top-off (not coverage-aware — Phase 3 handles gaps).
 * Mutates gridValues in place.
 */
function enforceMinimumWeeklyHours(gridValues, employeeNames, rosterMap, shiftTimingMap) {
    for (let employeeIndex = 0; employeeIndex < employeeNames.length; employeeIndex++) {
        const employeeGridOffset = employeeIndex * 3;
        const employeeSettings   = rosterMap[employeeNames[employeeIndex]] ||
            { status: "PT", pref: "Morning", qualifiedShifts: ["Morning"] };
        const weeklyHourTarget   = (employeeSettings.status === "FT") ? FT_MINIMUM_WEEKLY_HOURS : PT_MINIMUM_WEEKLY_HOURS;
        const maximumHours       = (employeeSettings.status === "FT") ? FT_MAXIMUM_WEEKLY_HOURS : PT_MAXIMUM_WEEKLY_HOURS;

        let currentWeeklyTotal = parseFloat(gridValues[employeeGridOffset + 2][7]) || 0;
        if (currentWeeklyTotal >= weeklyHourTarget) continue;

        const preferenceShift = shiftTimingMap[employeeSettings.pref + "|" + employeeSettings.status];
        if (!preferenceShift || preferenceShift.hours <= 0) continue;

        for (let dayIndex = 0; dayIndex < 7; dayIndex++) {
            if (gridValues[employeeGridOffset + 2][dayIndex] !== "OFF") continue;
            if ((currentWeeklyTotal + preferenceShift.hours) > maximumHours) continue;

            gridValues[employeeGridOffset + 2][dayIndex] = preferenceShift.text;
            currentWeeklyTotal += preferenceShift.hours;
            gridValues[employeeGridOffset + 2][7] = currentWeeklyTotal;

            if (currentWeeklyTotal >= weeklyHourTarget) break;
        }
    }
}

// ─── Phase 3 ──────────────────────────────────────────────────────────────────

/**
 * Detects gaps (30-min slots with zero coverage) and resolves them via a two-step cascade:
 *   Cascade A — Reassign a working employee to a different qualified shift that covers the gap
 *   Cascade B — Pull in an OFF employee if gaps survive Cascade A
 * VAC days are never touched. Mutates gridValues in place.
 */
function detectAndResolveCoverageGaps(gridValues, employeeNames, rosterMap, shiftTimingMap) {
    const totalEmployees = employeeNames.length;

    for (let dayIndex = 0; dayIndex < 7; dayIndex++) {
        let dayCoverageSlots = buildDayCoverageSlots(gridValues, employeeNames, dayIndex, shiftTimingMap);
        if (!dayCoverageSlots.some(slotCount => slotCount === 0)) continue;

        // Cascade A: Reassign a working employee to a better shift (junior first)
        for (let employeeIndex = totalEmployees - 1; employeeIndex >= 0; employeeIndex--) {
            const employeeGridOffset = employeeIndex * 3;
            const currentShiftText   = gridValues[employeeGridOffset + 2][dayIndex];

            if (!currentShiftText || !currentShiftText.toString().includes(":")) continue;

            const employeeSettings = rosterMap[employeeNames[employeeIndex]] ||
                { status: "PT", pref: "Morning", qualifiedShifts: ["Morning"] };
            if (employeeSettings.qualifiedShifts.length <= 1) continue;

            // Find the hours for the current shift so we can calculate the net hour change
            let currentShiftHours = 0;
            for (const shiftKey in shiftTimingMap) {
                if (shiftTimingMap[shiftKey].text === currentShiftText) {
                    currentShiftHours = shiftTimingMap[shiftKey].hours;
                    break;
                }
            }

            // Score alternative shifts against the coverage gaps (excluding this employee's current contribution)
            const coverageWithoutThisEmployee = buildDayCoverageSlots(gridValues, employeeNames, dayIndex, shiftTimingMap, employeeIndex);
            const alternativeShifts           = employeeSettings.qualifiedShifts.filter(shiftName => {
                const mapKey = shiftName.trim() + "|" + employeeSettings.status;
                return shiftTimingMap[mapKey] && shiftTimingMap[mapKey].text !== currentShiftText;
            });

            const bestAlternativeShift = selectBestCoverageShift(alternativeShifts, employeeSettings.status, coverageWithoutThisEmployee, shiftTimingMap);
            if (!bestAlternativeShift || bestAlternativeShift.hours === 0) continue;

            const accumulatedHours    = parseFloat(gridValues[employeeGridOffset + 2][7]) || 0;
            const newAccumulatedHours = (accumulatedHours - currentShiftHours) + bestAlternativeShift.hours;
            const maximumHours        = (employeeSettings.status === "FT") ? FT_MAXIMUM_WEEKLY_HOURS : PT_MAXIMUM_WEEKLY_HOURS;
            if (newAccumulatedHours > maximumHours) continue;

            gridValues[employeeGridOffset + 2][dayIndex] = bestAlternativeShift.text;
            gridValues[employeeGridOffset + 2][7]        = newAccumulatedHours;

            dayCoverageSlots = buildDayCoverageSlots(gridValues, employeeNames, dayIndex, shiftTimingMap);
            if (!dayCoverageSlots.some(slotCount => slotCount === 0)) break;
        }

        if (!dayCoverageSlots.some(slotCount => slotCount === 0)) continue;

        // Cascade B: Pull in an OFF employee to cover remaining gaps (junior first)
        for (let employeeIndex = totalEmployees - 1; employeeIndex >= 0; employeeIndex--) {
            const employeeGridOffset = employeeIndex * 3;
            if (gridValues[employeeGridOffset + 2][dayIndex] !== "OFF") continue;

            const employeeSettings = rosterMap[employeeNames[employeeIndex]] ||
                { status: "PT", pref: "Morning", qualifiedShifts: ["Morning"] };
            const accumulatedHours = parseFloat(gridValues[employeeGridOffset + 2][7]) || 0;
            const maximumHours     = (employeeSettings.status === "FT") ? FT_MAXIMUM_WEEKLY_HOURS : PT_MAXIMUM_WEEKLY_HOURS;

            const gapFillingShift = selectBestCoverageShift(employeeSettings.qualifiedShifts, employeeSettings.status, dayCoverageSlots, shiftTimingMap);
            if (!gapFillingShift || gapFillingShift.hours === 0) continue;
            if ((accumulatedHours + gapFillingShift.hours) > maximumHours) continue;

            gridValues[employeeGridOffset + 2][dayIndex] = gapFillingShift.text;
            gridValues[employeeGridOffset + 2][7]        = accumulatedHours + gapFillingShift.hours;

            dayCoverageSlots = buildDayCoverageSlots(gridValues, employeeNames, dayIndex, shiftTimingMap);
            if (!dayCoverageSlots.some(slotCount => slotCount === 0)) break;
        }
    }
}

// ─── Coverage Helpers ─────────────────────────────────────────────────────────

/**
 * Builds a 39-element coverage array for a single day (30-min slots from 04:00 to 23:30).
 * Each value = number of employees covering that slot.
 * Pass excludeEmployeeIndex to calculate coverage as if that employee is not scheduled.
 */
function buildDayCoverageSlots(gridValues, employeeNames, dayIndex, shiftTimingMap, excludeEmployeeIndex) {
    const coverageSlots = new Array(39).fill(0);

    for (let employeeIndex = 0; employeeIndex < employeeNames.length; employeeIndex++) {
        if (employeeIndex === excludeEmployeeIndex) continue;

        const shiftText = gridValues[employeeIndex * 3 + 2][dayIndex];
        if (!shiftText || !shiftText.toString().includes(":")) continue;

        for (const shiftKey in shiftTimingMap) {
            const shiftTiming = shiftTimingMap[shiftKey];
            if (shiftTiming.text !== shiftText) continue;

            const startSlot = Math.max(0, Math.floor((shiftTiming.startMinutes - 240) / 30));
            const endSlot   = Math.min(39, Math.floor((shiftTiming.endMinutes - 240) / 30));
            for (let slotIndex = startSlot; slotIndex < endSlot; slotIndex++) {
                coverageSlots[slotIndex]++;
            }
            break;
        }
    }

    return coverageSlots;
}

/**
 * Given a list of qualified shift names, returns the shift that best fills the current coverage gaps.
 * Scoring: sum of (1 / (currentCoverage + 1)) per slot covered — lower existing coverage = higher score.
 */
function selectBestCoverageShift(qualifiedShifts, employmentStatus, dayCoverageSlots, shiftTimingMap) {
    let bestShiftTiming = null;
    let bestScore       = -1;

    qualifiedShifts.forEach(shiftName => {
        const mapKey      = shiftName.toString().trim() + "|" + employmentStatus;
        const shiftTiming = shiftTimingMap[mapKey];
        if (!shiftTiming) return;

        const startSlot = Math.max(0, Math.floor((shiftTiming.startMinutes - 240) / 30));
        const endSlot   = Math.min(39, Math.floor((shiftTiming.endMinutes - 240) / 30));
        let shiftScore  = 0;

        for (let slotIndex = startSlot; slotIndex < endSlot; slotIndex++) {
            shiftScore += 1 / (dayCoverageSlots[slotIndex] + 1);
        }

        if (shiftScore > bestScore) {
            bestScore       = shiftScore;
            bestShiftTiming = shiftTiming;
        }
    });

    return bestShiftTiming || { text: "NO SHIFT", hours: 0, startMinutes: 0, endMinutes: 0 };
}

// ─── Utilities ────────────────────────────────────────────────────────────────

/**
 * Generates a standardized label for the weekly schedule tab (e.g., "Week_03_23_26").
 * Uses the Monday of the current week for consistent naming.
 */
function generateWeeklyDateLabel() {
    const todayDate    = new Date();
    const daysToMonday = (todayDate.getDay() === 0) ? 6 : todayDate.getDay() - 1;
    const mondayDate   = new Date(todayDate);
    mondayDate.setDate(todayDate.getDate() - daysToMonday);

    const calendarMonth = (mondayDate.getMonth() + 1).toString().padStart(2, '0');
    const calendarDay   = mondayDate.getDate().toString().padStart(2, '0');
    const shortYear     = mondayDate.getFullYear().toString().slice(-2);

    return `Week_${calendarMonth}_${calendarDay}_${shortYear}`;
}
