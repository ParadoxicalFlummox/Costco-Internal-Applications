# Maintenance Auto Scheduler
The purpose of this script is to test the feasibility of an automatic schedule generator on a small department before moving to others

The script would also include weighing the following in order to generate preferred schedules rather than completely random ones

## Weighted values
- Full Time or Part Time employee
- Years Served
- Preferred days off/two days together or split

## High Level Technical Overview
1. Data ingestion: The script scrapes the master employee data sheet and creates a database of employees
2. Processing (the weighted engine): Convert the dates and employment status into a seniority score into an objective math problem or value
3. Schedule simulation: The script checks each slot with questions "is this person available", "do they have rank over another person to have this spot", "will this create a shortage or gap in the schedule"
4. Output: Push the best fitting arrangement back into the spreadsheet or a new "generated sheet"

## Manager Guide:
This section is desiged to explain the logic to someone who doesn't care about code, but cares about results.

The **Goal**: To create a fair, full schedule that meets daily staffing needs while honoring employee seniority and requested days off and vacation/personal time

1. The brains of the operation (Seniority Logic)
The tool does not pick names at random. It calculates a seniority rank for every person based on:
- Length of service: How long they have been with the company.
- Employment status: Full time roles are priorized for hours over part time roles.
- Why does this matter: High seniority employees get the first "pick" for their preferred days off or consecutive shifts.

2. The "rules" (Staffing requirements)
You tell the tool how many people you need each day (e.g., 6 people scheduled monday 2 off, 4 people on sunday and 4 off).
The tool then works backward to fill those slots, ensuring that there are no "gaps" and building in any overlap to cover potential call-outs

3. The process
- Input: You as a manager update the Master Employee list and enter the department, 