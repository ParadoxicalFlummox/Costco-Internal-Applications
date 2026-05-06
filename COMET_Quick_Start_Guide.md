# COMET Operations Guide

**What is COMET?**  
A unified warehouse management system for scheduling, attendance tracking, absence logging, and infraction management. Designed to reduce manual spreadsheet work and keep everyone in sync across all departments.

---

## Table of Contents
1. [Manager Day-to-Day Workflows](#manager-day-to-day)
2. [Manager Scheduling: Hourly Employees](#manager-scheduling)
3. [Admin/Staff Setup & Maintenance](#admin-setup)
4. [Payroll & Infraction Workflows](#payroll-workflows)
5. [Naming Conventions & Best Practices](#naming-conventions)
6. [Troubleshooting](#troubleshooting)

---

## Manager Day-to-Day Workflows {#manager-day-to-day}

### 1. **Log Employee Absences** (First Thing You See)
- **When:** Employee calls out, no-shows, FMLA, or any unscheduled absence
- **How:** 
  1. Click **"New Entry"** button on the Absence tab
  2. Modal opens with form fields:
     - Employee (search by name or ID)
     - Date (select from calendar)
     - Absence Type (Call-Out, No-Show, FMLA, Other)
     - Notes (optional)
  3. Click **"Send"** to log immediately, OR check **"Auto-send"** toggle to queue for batch send
  4. If auto-send is enabled, click **"Send All Emails"** button to send queued absences at once
- **Result:** Absence logged; notification sent to payroll/relevant staff; logged in call history

**Note on LOA vs LPT:**
- **LOA (Leave of Absence):** Extended unpaid time off; managed in Admin tab (archive button to remove active status)
- **LPT (Limited Part Time):** Temporary reduction in hours; set in Admin → Edit Employee → FT/PT dropdown

---

## Manager Scheduling: Hourly Employees {#manager-scheduling}

### 2. **Generate & Manage Weekly Schedules**
- **When:** Creating or adjusting weekly employee schedules
- **How:**
  1. Go to **Schedule** tab
  2. Select your department and week
  3. Click **"Generate Schedule"** (system creates optimal schedule in ~6 seconds)
  4. Review: employee names, assigned shifts, coverage hours, any conflicts/gaps
  5. To edit a shift:
     - Click on the shift in the schedule
     - Modal opens with options to:
       - Select a **different pre-defined shift** from the list
       - Enter a **custom time range** (start/end time)
       - Mark VAC (Vacation) or RDO (Regular Day Off)
  6. Make all needed edits, then click **"Publish"** when satisfied
- **Result:** Schedule locked and visible to team; conflicts highlighted in red

---

### 3. **Schedule Hybrid Employees** (Multiple Departments)
**What is a hybrid employee?**  
An employee who works in multiple departments per their shift (e.g., Front End part-time, Merch full-time).

- **Before scheduling:**
  1. Ensure employee has a **secondary department** set in their profile (Admin → Edit Employee → Secondary Department field)
  2. Ensure their **preferred shifts** list includes both departments (e.g., "080-night" and "083-day")

- **How to schedule:**
  1. When generating the schedule for primary department (e.g., 080 Front End):
     - System shows hybrid employee with suggested shifts
     - Edit shifts using the modal picker (click shift, select time or custom range)
  2. When scheduling the secondary department (e.g., 083 Merch):
     - System recognizes they already have hours assigned elsewhere
     - Adjusts suggestions to avoid exceeding FT (40hrs) or PT maximums
  3. Manually verify no conflicts:
     - Same person scheduled for overlapping times = error (must fix before publishing)
     - Person exceeds max hours = warning (you can override if intentional)

- **Result:** Cohesive schedule across departments; no double-bookings

---

## Admin/Staff Setup & Maintenance {#admin-setup}

### 4. **Import Employee List from CSV**
**Who does this:** Staff or Admin (not Managers)  
**When:** Initial setup or periodic refresh to capture new hires, separations, role changes

- **How:**
  1. Go to **Admin** tab (password protected; ask your GM if locked)
  2. Click **"Import Employees"** or find the import button
  3. Select a CSV file from your computer (download from UKG payroll system)
  4. Click **"Upload"**
  5. System parses CSV and updates the Employees sheet
  6. Review the import summary: new employees, updates, any errors
- **What it imports:**
  - Employee name, ID, hire date, status (active/inactive)
  - Department assignment
  - FT/PT designation
  - Seniority rank (calculated automatically)
- **What you need to add manually after import** (see section 5)

- **Result:** Employee list current with payroll system

---

### 5. **Enrich Employee Data via Spreadsheet**
**Why:** CSV import only covers basic payroll data. You need to add shift preferences, days off, roles, and department assignments to make scheduling work properly.

**What to add (use the spreadsheet, not the UI—it's faster):**
1. **Preferred Days Off (full/part)** — Which days does each employee prefer off?
2. **Preferred Shifts** — What shifts/times do they prefer? (See naming conventions below)
3. **Qualified Shifts** — What shifts are they *trained* to work?
4. **Roles** — Department-specific roles (Cashier, SCO, Supervisor, Merch Team, etc.)
5. **Secondary Department** — For hybrid employees, their second department

**How:**
1. Go to **Admin** tab → **Employees** section
2. Click **"Edit as Spreadsheet"** (or find the Employees sheet in Google Drive)
3. Add/update columns as needed:
   - Column F: Preferred Days Off
   - Column G: Preferred Shifts
   - Column H: Qualified Shifts
   - Column I: Roles
   - Column J: Secondary Department
4. **Spell check your entries** (especially shift names—they must match your naming convention exactly)
5. Save; COMET will sync automatically within minutes
- **Result:** Scheduling engine has complete employee profiles and can generate smart schedules

---

## Naming Conventions & Best Practices {#naming-conventions}

**Shift Naming:** Consistency is critical for the scheduling engine to recognize shift preferences.

### Standard Format
```
[DEPT_NUMBER]-[SECONDARY_DEPT_IF_HYBRID]-[TIME_OR_SHIFT_NAME]
```

**Examples:**
- **Front End, night shift:** `080-night` or `080-2300` (you choose the time format)
- **Merch, day shift:** `083-day` or `083-1400`
- **Hybrid employee (Front End primary):** `080-083-night` (work FE during night)
- **Hybrid employee (Merch primary):** `083-080-day` (work Merch during day)

**Rules:**
- Use your department numbers (not names—easier to parse)
- Be consistent within your warehouse (don't mix `080-night` and `080-1100pm`)
- For hybrid shifts, list primary dept first, secondary dept second
- Spell check shift names in employee profiles (typos break scheduling)

**Real-world example:**
- You work Front End (080) primary and Merch (083) secondary
- Your shifts: `080-083-night` (FE at night) and `083-080-day` (Merch during day)
- This makes it clear which department is primary for each shift variant

---

## Payroll & Infraction Workflows {#payroll-workflows}

### 6. **Add or Update Attendance Codes**
- **When:** Keying in tardies (TD), no-shows (NS), swipe errors (SE), or other codes
- **How:**
  1. Navigate to the **Attendance Calendar** for the employee
  2. Click on a specific day cell (the day number and codes area)
  3. Modal opens with:
     - **Tardy counters** (TD-1, TD-2, TD-3) — use +/− buttons to add multiple instances
     - **Other codes** (NS, SE, MP, SZ, BL, LP, CN, FH, H, JD, etc.) — toggle buttons to select
  4. Adjust counts/selections as needed
  5. Click **"Save"** to update; calendar refreshes with new codes
- **Tardy tiers:**
  - TD-1: 4–29 minutes (light orange)
  - TD-2: 30–119 minutes (medium orange)
  - TD-3: 120+ minutes (dark orange)
- **Result:** Codes stored in employee record; visible on calendar; logged for infraction tracking

---

### 7. **Review Infractions & Send Notices**
- **When:** Policy violation confirmed (tardy threshold reached, no-show, etc.)
- **How:**
  1. Go to **Infractions** tab
  2. Review active infractions and check if policy thresholds are met (based on CONFIG rules)
  3. Create a new Counseling Notice (CN):
     - Click **"New CN"**
     - Enter employee, infraction code, and violation window
     - System auto-calculates if threshold is met
  4. Send notification:
     - Click **"Send Email"** to notify employee
     - Optional: customize email before sending
     - Confirm; email sent and logged automatically
- **Result:** CN filed, employee notified, full audit trail preserved

---

## Quick Tips

| Task | Location | Notes |
|------|----------|-------|
| Find employee | Top search bar | Start typing name or ID |
| Navigate weeks | Schedule tab | Use ← → arrows for prev/next week |
| Edit employee | Admin → Employees | Use spreadsheet view for bulk edits |
| Import CSV | Admin tab | Download from UKG, upload here |
| Add codes | Click day cell | Click the actual day cell, not day number |
| Send absences | Absence tab | Use "Send All Emails" button if auto-send enabled |

---

## Troubleshooting {#troubleshooting}

**"The schedule won't generate"**  
→ Make sure your staffing requirements are set (Admin → Staffing Requirements). Also check that employees have preferred shifts in their profiles.

**"Hybrid employee is showing in both department schedules with no hours balance"**  
→ Check that their secondary department is set correctly in Admin → Edit Employee.

**"I added a shift name but the employee still doesn't see it"**  
→ Spell check your entry (must match exactly, including spaces/hyphens). Also refresh your browser (Ctrl+R).

**"I logged an absence but don't see it in the call history"**  
→ Refresh the page (Ctrl+R). New entries appear instantly if sent; check auto-send status if queued.

**"CSV import says there are new employees but they don't appear"**  
→ They're in the system but not yet enriched. Go to Admin → Employees and add their preferred shifts, roles, and department info via spreadsheet.

**"The attendance code modal won't open"**  
→ Make sure you're clicking directly on a day cell. Also check that the calendar tab loaded correctly (refresh if needed).

**"I can't edit a published schedule"**  
→ Go back to the Schedule tab, select the week, and click "Generate Schedule" again. Edit in the new interface, then publish to replace the old one.

---

## Feedback & Support

**Have a bug report, feature request, or a question?**

📧 **Email:** adamr814@protonmail.com

When reporting issues, please include:
- What you were trying to do
- What went wrong (error message, unexpected behavior)
- Screenshots if helpful
- Your warehouse number and department

Also, you can click the **COMET logo 5 times** in the top-left corner to open the Version & Feedback modal, which displays the feedback email address.

---

**Version:** 0.7.3 | **Last Updated:** May 5, 2026 | **For Support:** Email adamr814@protonmail.com or contact your General Manager
