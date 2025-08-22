/**
 * This is the main function to update calculated data in the Master Directory.
 * It calculates full names, ages, membership date details, and looks up the last visit date.
 */
function runDailyUpdate() {
  // --- Configuration ---
  const TARGET_SHEET_NAME = getSetting('Master Directory Tab');
  const ATTENDANCE_URL = getSetting('Attendance Tracker URL');
  const VISIT_SHEET_NAME = getSetting('Attendance Stats Tab');

  // Column definitions
  const LAST_NAME_COL = 3,
    FIRST_NAME_COL = 4,
    FULL_NAME_COL = 2;
  const DOB_COL = 11,
    AGE_COL = 12,
    AGE_GROUP_COL = 13;
  const MEMBERSHIP_DATE_COL = 22,
    MEMBERSHIP_YEAR_COL = 25,
    MEMBERSHIP_QUARTER_COL = 26,
    MEMBERSHIP_MONTH_COL = 27,
    MEMBERSHIP_WEEK_COL = 28;
  const LAST_VISIT_COL = 23; // This is Column W

  const VISIT_DATA_RANGE_A1 = 'B2:H1001';
  const VISIT_KEY_COL_IDX_IN_DATA = 0;
  const VISIT_RETURN_COL_IDX_IN_DATA = 6;

  const HEADER_ROWS = 1;
  // --- End Configuration ---

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) {
    Logger.log(`Error: Sheet named "${TARGET_SHEET_NAME}" not found.`);
    return;
  }
  // CORRECTED: The error message now accurately reflects the variables being checked.
  if (!ATTENDANCE_URL || !VISIT_SHEET_NAME) {
    Logger.log(`Error: Please ensure 'Attendance Tracker URL' and 'Attendance Stats Tab' are set in your Config sheet.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROWS) return;

  const range = sheet.getRange(HEADER_ROWS + 1, 1, lastRow - HEADER_ROWS, sheet.getLastColumn());
  const data = range.getValues();

  const visitData = getVisitDataArray(ATTENDANCE_URL, VISIT_SHEET_NAME, VISIT_DATA_RANGE_A1);
  if (!visitData || visitData.length === 0) {
    Logger.log(`ERROR: Could not fetch any data from the attendance sheet. Please check permissions and configuration.`);
    return;
  } else {
    Logger.log(`SUCCESS: Fetched ${visitData.length} rows from the '${VISIT_SHEET_NAME}' tab.`);
  }

  let matchesFound = 0;

  for (let i = 0; i < data.length; i++) {
    const rowData = data[i];

    // --- UPDATED: Full Name, Age, and Membership calculations are now included ---
    // Full Name Calculation
    const firstName = rowData[FIRST_NAME_COL - 1];
    const lastName = rowData[LAST_NAME_COL - 1];
    if (firstName || lastName) {
     rowData[FULL_NAME_COL - 1] = [firstName, lastName].filter(Boolean).join(' ');
    }

    // Age and Age Group Calculation
    const dob = rowData[DOB_COL - 1];
    if (dob instanceof Date && !isNaN(dob)) {
      const age = calculateAge(dob);
      rowData[AGE_COL - 1] = age;
      rowData[AGE_GROUP_COL - 1] = getAgeGroup(age);
    } else {
      rowData[AGE_COL - 1] = "";
      rowData[AGE_GROUP_COL - 1] = "";
    }

    // Membership Date Calculations
    const membershipDate = rowData[MEMBERSHIP_DATE_COL - 1];
    if (membershipDate instanceof Date && !isNaN(membershipDate)) {
      rowData[MEMBERSHIP_YEAR_COL - 1] = membershipDate.getFullYear();
      rowData[MEMBERSHIP_MONTH_COL - 1] = membershipDate.getMonth() + 1; // getMonth() is 0-indexed
      rowData[MEMBERSHIP_QUARTER_COL - 1] = Math.floor(membershipDate.getMonth() / 3) + 1;
      rowData[MEMBERSHIP_WEEK_COL - 1] = getIsoWeek(membershipDate);
    } else {
      rowData[MEMBERSHIP_YEAR_COL - 1] = "";
      rowData[MEMBERSHIP_MONTH_COL - 1] = "";
      rowData[MEMBERSHIP_QUARTER_COL - 1] = "";
      rowData[MEMBERSHIP_WEEK_COL - 1] = "";
    }
    // --- End of new calculation block ---

    // --- VLOOKUP for "Last Visit" ---
    const existingFullName = String(rowData[FULL_NAME_COL - 1] || '').trim();
    if (existingFullName) { // Only search if there is a name
      const lookupResult = performScriptVlookup(existingFullName, visitData, VISIT_KEY_COL_IDX_IN_DATA, VISIT_RETURN_COL_IDX_IN_DATA);

      if (lookupResult) {
        matchesFound++;
      } else {
        Logger.log(`  ❌ NO MATCH FOUND for "${existingFullName}"`);
      }

      rowData[LAST_VISIT_COL - 1] = lookupResult;
    }
  }

  range.setValues(data);
  Logger.log(`--- SUMMARY ---`);
  Logger.log(`Found last visit dates for ${matchesFound} out of ${data.length} people.`);
  Logger.log(`Successfully updated calculated data.`);
}


// --- HELPER FUNCTIONS ---

/**
 * Retrieves data from a specified range in an external spreadsheet.
 */
function getVisitDataArray(spreadsheetUrl, sheetName, rangeA1) {
  try {
    const ss = SpreadsheetApp.openByUrl(spreadsheetUrl);
    const sheet = ss.getSheetByName(sheetName);
    return sheet ? sheet.getRange(rangeA1).getValues() : null;
  } catch (e) {
    Logger.log(`Error accessing external sheet: ${e.message}`);
    return null;
  }
}

/**
 * Calculates age based on a birth date.
 */
function calculateAge(birthDate) {
  if (!(birthDate instanceof Date) || isNaN(birthDate)) return "";
  const today = new Date();
  if (birthDate > today) return "";
  let age = today.getFullYear() - birthDate.getFullYear();
  const m = today.getMonth() - birthDate.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }
  return age >= 0 ? age : "";
}

/**
 * Determines the age group based on age.
 */
function getAgeGroup(age) {
  if (typeof age !== 'number' || age < 0) return "";
  if (age < 19) return "0-18";
  if (age < 40) return "19-39";
  if (age < 60) return "40-59";
  return "60+";
}

/**
 * Calculates the ISO week number for a given date.
 */
function getIsoWeek(date) {
  if (!(date instanceof Date) || isNaN(date)) return "";
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

/**
 * Performs a VLOOKUP-like operation on a 2D array, cleaning whitespace before comparing.
 */
function performScriptVlookup(searchValue, dataArray, searchColIndex, returnColIndex) {
  if (!dataArray || dataArray.length === 0) return "";

  const cleanedSearchValue = searchValue.replace(/\s+/g, ' ').trim();

  for (let i = 0; i < dataArray.length; i++) {
    const valueFromSheet = String(dataArray[i][searchColIndex]);
    const cleanedValueFromSheet = valueFromSheet.replace(/\s+/g, ' ').trim();

    if (cleanedValueFromSheet === cleanedSearchValue) {
      return dataArray[i][returnColIndex];
    }
  }
  return "";
}

/**
 * Retrieves a setting value from the 'Config for Urls' sheet.
 */
function getSetting(settingName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config for Urls");
  if (!configSheet) return null;
  const data = configSheet.getRange("A:B").getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === settingName) {
      return data[i][1];
    }
  }
  return null;
}

/**
 * A temporary helper function to force the permission pop-up for external sheet access.
 */
function forcePermissionRequest() {
  const url = getSetting('Directory 2 URL');
  if (url) {
    try {
      SpreadsheetApp.openByUrl(url);
      Logger.log("SUCCESS: The script was able to access the external URL. Permissions should now be granted.");
      Logger.log("You can now select 'runDailyUpdate' from the dropdown and run the main script.");
    } catch (e) {
      Logger.log(`ERROR: Could not access the URL. Please double-check it in your Config sheet. Details: ${e.message}`);
    }
  } else {
    Logger.log('ERROR: The "Directory 2 URL" setting is missing from your Config sheet.');
  }
}
