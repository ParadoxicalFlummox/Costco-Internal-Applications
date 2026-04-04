# Auto Schedule Generator
**Version 0.3.3**

---

## PART ONE — Manager's Guide

*Plain-language instructions for day-to-day use. No technical knowledge required.*

---

### What This Tool Does

The Auto Schedule Generator automatically builds a weekly work schedule for your department. You tell it who your employees are, what shifts exist, and how many people you need each day — it handles the rest.

In a single click it generates three weeks of schedules at once (the current week plus the next two weeks), giving you roughly a month of forward visibility. Each week appears as its own color-coded tab in the spreadsheet, making it easy to review, adjust, and share with your team.

Schedules are built fairly: more senior employees get first pick of their preferred days off, and full-time employees are prioritized over part-time employees with the same hire date. The tool automatically ensures everyone hits their required weekly hours.

---

### How Schedules Are Built

The tool follows a clear priority order every time it generates a schedule:

1. **Seniority** — Employees who have worked at the company longer get first choice of their preferred days off and preferred shifts. Full-time employees outrank part-time employees with the same hire date.

2. **Employment Type (FT/PT)** — Full-time employees must work exactly 40 hours per week. Part-time employees must work between 24 and 35 hours per week.

3. **Vacation Days** — Any dates entered in the Vacation Dates column of the Roster are treated as immovable. No shift will ever be placed on a vacation day.

4. **Preferred Days Off** — Each employee can have up to two preferred days off per week. The tool honors these requests in seniority order, as long as doing so does not drop the team below the minimum staffing level you have set.

5. **Preferred Shift** — The tool assigns each working employee their preferred shift (e.g., Morning, Mid, or Closing). When two employees with different preferences are competing, the tool picks whichever shift better fills a coverage gap.

**Example:** If Jane has 15 years of service and Bob has 3 years, and both request Monday off, Jane's request is granted first. If granting Bob's request would leave the department understaffed, Bob is scheduled to work on Monday.

After honoring preferences, the tool checks that everyone meets their weekly minimum hours and fills any remaining coverage gaps automatically.

---

### What You Need Before You Start

Before running the tool for the first time you will need:

1. **A master employee list** in a separate Google Sheet. This sheet must have the following columns:
   - Column A: Employee Name
   - Column B: Employee ID
   - Column C: Department
   - Column F: Hire Date

   The tool reads these specific column positions. If your source sheet has a different layout, see the technical section on adapting `fetchEmployeesFromSource()`.

2. **The Spreadsheet ID** of that master employee list. You can find it in the URL of the spreadsheet — it is the long string of characters between `/d/` and `/edit`.
   - Example URL: `https://docs.google.com/spreadsheets/d/`**`1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms`**`/edit`
   - The bold portion is the spreadsheet ID.

3. **Edit access** to both this scheduling spreadsheet and the master employee list.

---

### Step-by-Step Usage Guide

#### Step 1: First-Time Setup

The very first time you open this spreadsheet, click:

> **Schedule Admin → Setup Sheets (First Run Only)**

This creates three support sheets — **Ingestion**, **Roster**, and **Settings** — with example content to get you started. You only need to do this once.

---

#### Step 2: Configure Your Shifts and Staffing Requirements

Open the **Settings** sheet. You will see two tables:

**Staffing Requirements Table (columns A–B):**
Fill in how many employees you need working each day. The tool will not schedule fewer than this number.

| Day       | Min Staff Required |
|-----------|--------------------|
| Monday    | 6                  |
| Tuesday   | 6                  |
| ...       | ...                |
| Sunday    | 4                  |

**Shift Definitions Table (columns D–I):**
Each row defines one shift. Full-time and part-time versions of the same shift are separate rows with the same name. The first-run template includes Early, Morning, Mid, and Closing shifts as a starting point.

| Shift Name | Status | Start Time | End Time  | Paid Hours | Has Lunch |
|------------|--------|------------|-----------|------------|-----------|
| Early      | FT     | 4:00 AM    | 12:30 PM  | 8.0        | TRUE      |
| Early      | PT     | 4:00 AM    | 9:00 AM   | 5.0        | FALSE     |
| Early+     | PT     | 4:00 AM    | 9:30 AM   | 5.0        | TRUE      |
| Morning    | FT     | 8:00 AM    | 4:30 PM   | 8.0        | TRUE      |
| Morning    | PT     | 8:00 AM    | 1:00 PM   | 5.0        | FALSE     |
| Morning+   | PT     | 8:00 AM    | 1:30 PM   | 5.0        | TRUE      |
| Mid        | FT     | 11:00 AM   | 7:30 PM   | 8.0        | TRUE      |
| Mid        | PT     | 11:00 AM   | 4:00 PM   | 5.0        | FALSE     |
| Mid+       | PT     | 11:00 AM   | 4:30 PM   | 5.0        | TRUE      |
| Closing    | FT     | 2:00 PM    | 10:30 PM  | 8.0        | TRUE      |
| Closing    | PT     | 2:00 PM    | 7:00 PM   | 5.0        | FALSE     |
| Closing+   | PT     | 2:00 PM    | 7:30 PM   | 5.0        | TRUE      |

- **FT shifts** are always 8 paid hours + 30 minutes unpaid lunch = 8.5 hours on the clock.
- **PT shifts** are 5 paid hours, with or without a 30-minute unpaid lunch.
- Use a "+" suffix to distinguish the lunch variant of a PT shift (e.g., "Morning" vs "Morning+").

The template created in Step 1 includes all rows above. Modify times or remove shifts to match your department's actual hours.

---

#### Step 3: Connect Your Employee List

1. Open the **Ingestion** sheet.
2. In cell B3, type or paste your master spreadsheet ID.
3. Wait a moment — the Department dropdown in cell B4 will automatically populate.
4. Select your department from the B4 dropdown.
5. Click: **Schedule Admin → Sync Roster**

The tool will copy all employees from that department into the **Roster** sheet. New employees will appear with their name, ID, and hire date filled in, and their status set to PT by default.

*After syncing, update each employee's status (FT or PT), preferred days off, preferred shift, and qualified shifts in the Roster sheet.*

---

#### Step 4: Fill In Employee Preferences

In the **Roster** sheet, complete the following columns for each employee:

| Column | What to Enter |
|--------|---------------|
| **D — Status** | FT (full-time, 40 hrs/week) or PT (part-time, 24–35 hrs/week) |
| **E — Day Off Pref 1** | First preferred day off (e.g., Monday) |
| **F — Day Off Pref 2** | Second preferred day off (e.g., Tuesday) |
| **G — Preferred Shift** | The shift they prefer most (e.g., Morning) |
| **H — Qualified Shifts** | All shifts they are trained to work, comma-separated (e.g., Morning, Mid) |
| **I — Vacation Dates** | Dates they will be absent, comma-separated (e.g., 2026-04-14, 2026-04-15) |

The **Seniority Rank** column (J) is calculated automatically — do not edit it.

---

#### Step 5: Generate a Schedule Draft

Once your Settings and Roster are ready, you have two generation options:

**Option A — Three weeks (recommended for regular use):**
> **Schedule Admin → GENERATE SCHEDULE (3 weeks)**

A dialog will appear asking you to confirm the Monday of the starting week. Press OK to accept the current week, or type a different date.

Three new sheet tabs will be created:
- **Week_MM_DD_YY** — Current week
- **Week_MM_DD_YY** — Following week
- **Week_MM_DD_YY** — Week after that

A confirmation message will appear listing the names of all three sheets.

**Option B — Single week (useful for re-generating one specific week):**
> **Schedule Admin → GENERATE SCHEDULE (1 week)**

This generates only the current week's schedule tab. Useful when you need to quickly regenerate after a roster change without overwriting the other two weeks.

---

#### Step 6: Reading the Generated Schedule

Each schedule sheet has the same structure:

**Header rows (1–4):** Week date range, generation timestamp, department name.

**Column headers (row 5):** Label | Employee | Mon | Tue | Wed | Thu | Fri | Sat | Sun | Total Hrs

**Employee rows:** Each employee has three rows:
- **VAC row:** Check a box to mark that day as vacation (locks the day — no shift will be assigned).
- **RDO row:** Check a box to mark a requested day off (honored if staffing allows).
- **SHIFT row:** Shows the assigned shift time (e.g., "8:00 AM - 4:30 PM") or "OFF". Total hours appear in the last column.

**Summary rows:** At the bottom of the sheet:
- **REQUIRED:** Minimum staff needed each day (from Settings).
- **ACTUAL:** Count of employees scheduled that day (auto-calculated).
- **STATUS:** "OK" (green) if staffed, "UNDER" (red) if below minimum.

**Color Reference:**

| Color  | Meaning |
|--------|---------|
| Blue   | Full-time shift |
| Green  | Part-time shift |
| Yellow | Vacation day |
| Gray   | Day off (RDO or unscheduled) |
| Red name | Employee is below their weekly hour minimum |

---

#### Step 7: Making Manual Adjustments

**To mark a vacation day:** Check the VAC checkbox for that employee on that day. The schedule will recalculate automatically.

**To mark a requested day off:** Check the RDO checkbox. The schedule will recalculate and honor the request if staffing allows.

**To change a shift assignment:** Directly edit the text in the SHIFT row cell (e.g., change "8:00 AM - 4:30 PM" to "10:00 AM - 6:30 PM").

**To add or remove employees:** Update the Roster sheet and re-generate the schedule.

---

#### What to Do When the Schedule Shows Red Names

A red employee name means the tool could not schedule enough shifts to meet that employee's weekly minimum. Common causes:

1. **Too many vacation days** — The employee has so many vacation days this week that there are not enough working days to reach the minimum. This is expected and requires no action unless you want to override it.
2. **No valid shifts configured** — Check that the employee's "Preferred Shift" and "Qualified Shifts" match shift names in the Settings sheet exactly (spelling and capitalization matter).
3. **Weekly maximum reached early** — Unlikely, but possible if a PT employee's shift durations do not divide evenly into their maximum hours.

---

### Limitations

- **Midnight-crossing shifts are not supported.** Shifts that start before midnight and end after midnight (e.g., 10:00 PM – 6:00 AM) cannot be used. All shifts must start and end within the same calendar day.
- **Simultaneous editing is not recommended.** If two people edit VAC or RDO checkboxes at the same time, the last change wins. Coordinate with anyone who has access to the schedule before making changes.
- **The tool generates draft schedules.** It automates the time-consuming first pass, but manager review and manual adjustments are always expected before the schedule is published.
- **Vacation dates must be entered before generating.** The tool reads vacation dates at generation time. If you add a vacation date after generating, re-run generation for that week or manually check the VAC checkbox.

---

---

## PART TWO — Technical Reference

*For developers, technically-minded managers, or anyone who wants to understand how the tool works internally or adapt it for other use cases.*

---

### File Map

The tool is split into six JavaScript files, each with a single responsibility. This means a bug in one area of the system is isolated to one file.

```
autoScheduleGenerator/
├── config.js           All constants, colors, column positions, hour rules.
│                       No functions — only constant declarations. Every other file
│                       reads from this file and nothing else needs changing when a
│                       business rule changes.
│
├── settingsManager.js  Reads the Settings sheet (shift definitions and staffing
│                       requirements). The only file that reads the Settings sheet.
│                       Returns clean JavaScript objects that all other files use.
│                       Uses Utilities.formatDate() with the spreadsheet's own
│                       timezone when parsing time cell values to avoid offset
│                       errors caused by mismatches between the script execution
│                       timezone and the spreadsheet's timezone.
│
├── rosterIngestion.js  Handles the "bring employees in from an external spreadsheet"
│                       workflow. The only file that reads the Ingestion sheet and
│                       writes to the Roster sheet during sync.
│
├── scheduleEngine.js   Contains the 4-phase scheduling algorithm. Takes a roster
│                       and shift definitions as input and produces a WeekGrid as
│                       output. Pure computation — reads the Roster sheet once but
│                       never writes to any sheet.
│
├── formatter.js        Translates a WeekGrid (from scheduleEngine.js) into what
│                       managers see on screen. The only file that writes to Week
│                       schedule sheets.
│
└── ui.js               Entry point for all user actions. Contains onOpen(), onEdit(),
                        menu handlers, and thin orchestrators that connect the engine
                        and formatter. Contains no scheduling or formatting logic.
```

---

### Data Flow Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│  MANAGER enters source spreadsheet ID in Ingestion sheet        │
└─────────────────────┬───────────────────────────────────────────┘
                      │ onEdit() → populateDepartmentDropdown()
                      ▼
┌─────────────────────────────────────────────────────────────────┐
│  INGESTION SHEET                                                 │
│  Source Spreadsheet ID  →  Department dropdown                   │
└─────────────────────┬───────────────────────────────────────────┘
                      │ menuSyncRoster() → syncRosterFromSource()
                      │   fetchEmployeesFromSource()
                      │   deduplicateAgainstRoster()
                      │   writeNewEmployeesToRoster()
                      ▼
┌─────────────────────────────────────────────────────────────────┐
│  ROSTER SHEET                                                    │
│  Name | ID | Hire Date | FT/PT | Day Off Prefs | Shifts | Vac   │
└─────────────────────┬───────────────────────────────────────────┘
                      │ menuGenerateScheduleDraft()
                      │   loadRosterSortedBySeniority()
                      │
┌─────────────────────┴───────────────────────────────────────────┐
│  SETTINGS SHEET                                                  │
│  Staffing requirements + Shift definitions                       │
└─────────────────────┬───────────────────────────────────────────┘
                      │   buildShiftTimingMap()
                      │   loadStaffingRequirements()
                      ▼
┌─────────────────────────────────────────────────────────────────┐
│  SCHEDULE ENGINE (in-memory)                                     │
│  Phase 0: Bootstrap       → WeekGrid (all OFF + VAC locks)      │
│  Phase 1: Preferences     → WeekGrid (RDO + SHIFT assigned)     │
│  Phase 2: Hour Minimum    → WeekGrid (more SHIFTs added)        │
│  Phase 3: Gap Resolution  → WeekGrid (gaps filled)              │
└─────────────────────┬───────────────────────────────────────────┘
                      │   writeAndFormatSchedule()
                      ▼
┌─────────────────────────────────────────────────────────────────┐
│  WEEK_MM_DD_YY SHEET                                             │
│  Color-coded schedule with VAC/RDO/SHIFT rows per employee      │
│  REQUIRED / ACTUAL / STATUS summary footer                      │
└─────────────────────────────────────────────────────────────────┘
```

---

### The Seniority Rank Formula

The seniority rank is a single integer that encodes both employment status and length of service. The formula is:

```
statusBase = 200,000,000  (for FT)  or  100,000,000  (for PT)
referenceDate = January 1, 2050
daysFromHireToReference = floor((referenceDate − hireDate) ÷ 86,400,000 ms)
seniorityRank = statusBase + daysFromHireToReference
```

**Why this works:**

- An employee hired in 2005 has a larger `daysFromHireToReference` than one hired in 2015 (because 2005 is further from 2050). So earlier hire dates produce higher ranks — no conditional logic needed.
- The 200M/100M split between FT and PT is large enough that no realistic `daysFromHireToReference` value can close the gap (you would need to be hired ~273,000 years ago for PT to ever beat FT). This ensures FT always outranks PT at the same hire date.
- Employees are sorted by descending rank, so the highest number = most senior.

**To change the FT/PT weighting**, edit `SENIORITY.FT_BASE` and `SENIORITY.PT_BASE` in `config.js`. Keep the gap large enough (at least 200,000) to ensure no real hire date can cross the boundary.

**To change the reference date**, edit `SENIORITY.REFERENCE_DATE_STRING` in `config.js`. Any future date works — 2050 was chosen arbitrarily. Changing this will recompute all seniority ranks, so run "Refresh Seniority" from the menu afterward.

---

### The 4-Phase Generation Algorithm

#### Phase 0 — Bootstrap

**What it does:** Reads the Roster sheet, sorts employees by seniority, and creates an empty WeekGrid. Vacation dates from the Roster are checked against the current week's date range; matching dates are stamped as locked VAC cells.

**Inputs:** Roster sheet, week start date.
**Output:** Sorted employee list, WeekGrid (all cells OFF or VAC).
**If skipped:** No other phase has data to work with — the schedule cannot be built.

#### Phase 1 — Preference Assignment

**What it does:** Processes employees in seniority order (most senior first). For each day, it grants Requested Days Off (RDO) up to the staffing floor, then assigns each working employee their preferred shift. A running coverage map tracks which 30-minute windows are covered so that employees with multiple qualified shifts are steered toward whichever shift fills a gap.

**Inputs:** WeekGrid (with VAC locks), employee list, shift timing map, staffing requirements.
**Output:** WeekGrid with RDO and SHIFT cells assigned.
**If skipped:** Employees would receive no preferred-day-off consideration and would all be assigned shifts without fairness or coverage awareness.

Phase 1 is split into three sub-functions to keep each concern isolated:
1. `grantRequestedDaysOff()` — decides which RDO requests are honored.
2. `assignPreferredShifts()` — assigns shift text to remaining working days.
3. `runPhaseOnePreferenceAssignment()` — thin orchestrator that calls both.

#### Phase 2 — Minimum Hour Enforcement

**What it does:** Checks each employee's total paid hours after Phase 1. Any employee below their minimum (40 for FT, 24 for PT) gets additional shifts added to their remaining OFF days until the minimum is reached or no OFF days remain. The weekly maximum (40 for FT, 35 for PT) is respected as a hard cap.

**Inputs:** WeekGrid (with Phase 1 assignments), employee list, shift timing map.
**Output:** WeekGrid with additional SHIFT cells where needed.
**If skipped:** Some employees would finish the week below their required minimum hours. The red name highlight would flag them but no correction would be attempted.

#### Phase 3 — Gap Resolution

**What it does:** For each day, builds a 30-minute slot coverage map and checks for zero-coverage windows (gaps). Runs two cascades to fill gaps:

- **Cascade A** — Already-working employees are considered for shift reassignment. If switching an employee to an alternative qualified shift would better cover uncovered slots without pushing them over their weekly maximum, the reassignment is made.
- **Cascade B** — If gaps remain after Cascade A, employees currently scheduled as OFF are pulled in and assigned the best gap-filling shift from their qualified list.

Both cascades process employees from most junior to most senior, so schedule disruptions fall on junior employees first.

**Inputs:** WeekGrid (after Phase 2), employee list, shift timing map, staffing requirements.
**Output:** WeekGrid with coverage gaps minimized.
**If skipped:** The schedule may have uncovered time windows. The STATUS row would show "UNDER" for those days but no automatic gap-filling would occur.

---

### The Coverage Slot Map

The coverage map is a 39-element array representing the operating day in 30-minute windows from **04:00 to 23:30**.

```
Slot index:  0       1       2     ...   38
Time window: 04:00   04:30   05:00  ...  23:00–23:30
```

Each element counts how many employees are physically present during that window. A value of `0` means a gap.

**Slot index formula:**
```
slotIndex = floor((minutesSinceMidnight − 240) ÷ 30)
```
Where 240 = 4 × 60 = 04:00 in minutes since midnight.

**Coverage vs. paid hours:** The slot map uses the shift's block hours (start to end including unpaid lunch), not paid hours. A full-time employee on an 8.5-hour shift physically covers 17 slots even though they are only paid for 16 slots (8 hours).

**Weekend coverage windows differ from weekday windows.** The tool enforces separate closing times for Saturday and Sunday because those days have shorter operating hours:

| Day            | Coverage Window         |
|----------------|-------------------------|
| Monday–Friday  | 4:00 AM – 11:30 PM      |
| Saturday       | 4:00 AM – 10:00 PM      |
| Sunday         | 4:00 AM – 9:00 PM       |

Phase 3 (gap resolution) uses these windows when deciding whether a time slot needs coverage. Slots outside the day's window are never treated as gaps. These windows are defined in the `COVERAGE_WINDOW` constant in `config.js` — update them there if your department's hours change.

**To extend the weekday coverage window** (e.g., for a 3:00 AM opening shift):
1. Change `COVERAGE.COVERAGE_START_MINUTE` in `config.js` (e.g., 180 for 03:00).
2. Recalculate `COVERAGE.SLOT_COUNT`: `(endMinute − startMinute) ÷ 30`.
3. Update `COVERAGE_WINDOW` entries to match the new start minute.
4. Midnight-crossing shifts are still not supported — the end time must be before midnight.

---

### How to Adapt for a Different Department

The tool is designed to support multiple departments by deploying separate copies of the spreadsheet. Here is the step-by-step process:

1. **Create a new Google Spreadsheet** and paste all six `.js` files into its Apps Script editor.
2. Run **Setup Sheets (First Run Only)** to create the Ingestion, Roster, and Settings sheets.
3. In the **Settings** sheet, replace the example shifts with the shifts used by the new department.
4. Update the staffing requirements table with the correct minimum staff counts for the new department.
5. In the **Ingestion** sheet, enter the source spreadsheet ID and select the new department.
6. Run **Sync Roster** to pull in the department's employees.
7. Fill in employee preferences in the Roster sheet.
8. Generate a schedule draft.

No code changes are required to support a different department. Everything is driven by the Settings and Roster sheets.

---

### How to Add a New Shift Type

1. Open the **Settings** sheet.
2. Add a new row to the Shift Definitions table (columns D–I) for the FT version:
   - Shift Name: the new shift's name (e.g., "Early")
   - Status: FT
   - Start Time, End Time, Paid Hours: appropriate values
   - Has Lunch: TRUE (for FT)
3. Add another row for the PT version with the same name, Status = PT, and PT hours.
4. Optionally add a PT+ version with a lunch break.
5. Run **Schedule Admin → Sync Shift Dropdowns** to update the Preferred Shift dropdown in the Roster sheet.
6. For each employee who should work this shift, update their Preferred Shift and/or Qualified Shifts columns.

---

### Known Limitations and Why They Exist

**Midnight-crossing shifts not supported.**
The coverage slot array ends at 23:30 (slot index 38). A shift that ends at, say, 02:00 the next day would need to wrap around to the beginning of the next day's array — this would require tracking coverage across a day boundary, which significantly complicates the algorithm. For departments that need overnight shifts, a future version could extend the slot array to 47 slots (04:00–05:30+24hr) and handle the wrap-around.

**Simultaneous editing is not recommended.**
Google Apps Script is single-threaded per execution. Two simultaneous onEdit() triggers can interleave unpredictably, and the last one to finish wins. This is a platform limitation, not something the tool can solve. The workaround is for managers to coordinate edits — or to use the "Generate Schedule Draft" menu item to do a full re-generation after all edits are complete.

**VAC and RDO checkbox edits recalculate one week at a time.**
Editing a VAC or RDO checkbox triggers an automatic `resolveEntireWeek()` on that sheet only — the other week tabs are not affected. To regenerate from scratch, use one of the Generate Schedule menu items (which preserve existing VAC/RDO checkboxes).

**Source spreadsheet must match the expected column layout.**
The sync function reads: Column A = Employee Name, Column B = Employee ID, Column C = Department, Column F = Hire Date. If your source sheet uses a different column order, edit `fetchEmployeesFromSource()` in `rosterIngestion.js` to use the correct column indices.

---

### Versioning and Change Log

This project follows `v0.MINOR.PATCH` versioning:

| Version | Description |
|---------|-------------|
| v0.3.1  | Weekend coverage windows (Saturday 4 AM–10 PM, Sunday 4 AM–9 PM). Single-week generation menu option. 12-hour AM/PM shift time display. Expanded shift template (Early, Morning, Mid, Closing with FT/PT/PT+ variants). Timezone fix for shift time parsing (`Utilities.formatDate()` with spreadsheet timezone). Corrected source spreadsheet column mapping (Department = col C, Hire Date = col F). |
| v0.3.0  | Initial rewrite. External roster ingestion via Ingestion sheet, 3-week generation, FT/PT color coding, seniority rank formula, 4-phase generation algorithm, department-agnostic configuration, comprehensive comments. |

Previous versions (v0.2.x and earlier) are archived in the git history of the `refactor-to-improve-core-logic` branch.

To report a bug or request a feature, file an issue in the project repository.
