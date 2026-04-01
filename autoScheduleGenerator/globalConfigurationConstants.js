/**
 *       _____             _                                _            _____        _                _         _                    
 *      / ____|           | |                  /\          | |          / ____|      | |              | |       | |                   
 *     | |      ___   ___ | |_  ___  ___      /  \   _   _ | |_  ___   | (___    ___ | |__    ___   __| | _   _ | |  ___  _ __        
 *     | |     / _ \ / __|| __|/ __|/ _ \    / /\ \ | | | || __|/ _ \   \___ \  / __|| '_ \  / _ \ / _` || | | || | / _ \| '__|       
 *     | |____| (_) |\__ \| |_| (__| (_) |  / ____ \| |_| || |_| (_) |  ____) || (__ | | | ||  __/| (_| || |_| || ||  __/| |          
 *      \_____|\___/ |___/ \__|\___|\___/  /_/    \_\\__,_| \__|\___/  |_____/  \___||_| |_| \___| \__,_| \__,_||_| \___||_|          
 *       _____  _         _             _    _____                __  _        
 *      / ____|| |       | |           | |  / ____|              / _|(_)       
 *     | |  __ | |  ___  | |__    __ _ | | | |      ___   _ __  | |_  _   __ _ 
 *     | | |_ || | / _ \ | '_ \  / _` || | | |     / _ \ | '_ \ |  _|| | / _` |
 *     | |__| || || (_) || |_) || (_| || | | |____| (_) || | | || |  | || (_| |
 *      \_____||_| \___/ |_.__/  \__,_||_|  \_____|\___/ |_| |_||_|  |_| \__, |
 *                                                                        __/ |
 *                                                                       |___/  
 * 
 * Built by: Adam Roy
 * Branch: shift-window-with-minimum
 * Version 0.2.1
 */

/* --- Application Environment Settings --- */
const TARGET_DEPARTMENT_NAME = "Maintenance";
const TEMPLATE_SHEET_NAME = "Grid Template";
const CONFIGURATION_SHEET_NAME = "CONFIG";
const SETTINGS_SHEET_NAME = "SETTINGS";

/* --- Configuration Sheet Column Mapping (0-indexed) --- */
const COLUMN_INDEX_NAME = 0; // Column A
const COLUMN_INDEX_EMPLOYEE_ID = 1; // Column B
const COLUMN_INDEX_HIRE_DATE = 2; // Column C
const COLUMN_INDEX_EMPLOYMENT_STATUS = 3; // Column D
const COLUMN_INDEX_PREFERENCE_ONE = 4; // Column E
const COLUMN_INDEX_PREFERENCE_TWO = 5; // Column F
const COLUMN_INDEX_SENIORITY_RANK = 6; // Column G
const COLUMN_INDEX_SHIFT_PREFERENCE = 7; // Column H
const COLUMN_INDEX_QUALIFIED_SHIFTS = 8; // Column I — comma-separated list of shifts the employee is cleared to work

/* --- Master Input Data Mapping (0-indexed) --- */
const MASTER_COLUMN_NAME = 0;
const MASTER_COLUMN_ID = 1;
const MASTER_COLUMN_DEPARTMENT = 2;
const MASTER_COLUMN_HIRE_DATE = 5;

/* --- Weekly Schedule Column Mapping --- */
const COLUMN_INDEX_TOTAL_HOURS = 9; // Column J (0-indexed absolute); offset by 2 when used in grid arrays starting at Column C
const TIME_FORMAT_STRING = "HH:mm"; // Make this "HH:mm a" to use 12hr time
const TIME_ZONE = "America/New_York";

/* --- Weekly Hour Minimums and Maximums --- */
const FT_MINIMUM_WEEKLY_HOURS = 40;
const PT_MINIMUM_WEEKLY_HOURS = 24;
const FT_MAXIMUM_WEEKLY_HOURS = 40;
const PT_MAXIMUM_WEEKLY_HOURS = 30;
