/**
 * Combines names, calculates age, age group, membership details, and last visit
 * when relevant data is edited in the "Directory" sheet.
 * All outputs are written as static values.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object from the edit.
 */
function onEdit(e) {
  // --- Configuration ---
  
  // -- ⭐️ CHANGED LINES START HERE ⭐️ --
  const TARGET_SHEET_NAME = getSetting('Master Directory Tab'); // Was "Directory"
  const VISIT_SHEET_NAME = getSetting('Visit Worksheet Tab');   // Was 'Visit Wksht'
  // -- ⭐️ CHANGED LINES END HERE ⭐️ --

  // Name Combination Columns
  const LAST_NAME_COL = 3;   // Column C for Last Name
  const FIRST_NAME_COL = 4;  // Column D for First Name
  const FULL_NAME_COL = 2;   // Column B for Full Name (output)

  // Date of Birth & Age Columns
  const DOB_COL = 11;        // Column K for Date of Birth (input)
  const AGE_COL = 12;        // Column L for Age (output)
  const AGE_GROUP_COL = 13;  // Column M for Age Group (output)

  // Membership Date Columns
  const MEMBERSHIP_DATE_COL = 22; // Column V for Membership Date (input)
  const MEMBERSHIP_YEAR_COL = 25; // Column Y for Membership Year (output)
  const MEMBERSHIP_QUARTER_COL = 26;// Column Z for Membership Quarter (output)
  const MEMBERSHIP_MONTH_COL = 27; // Column AA for Membership Month (output)
  const MEMBERSHIP_WEEK_COL = 28;  // Column AB for Membership Week (ISO) (output)

  // Last Visit Lookup Columns
  const LAST_VISIT_COL = 23;            // Column W for Last 2025 Visit (output)
  const VISIT_DATA_RANGE_A1 = 'B4:D1001'; // Range in Visit Wksht (e.g., B4:D for key in B, value in D)
  const VISIT_KEY_COL_IDX_IN_DATA = 0;   // 0-indexed: Col B is the 1st col in B:D, so index 0
  const VISIT_RETURN_COL_IDX_IN_DATA = 2; // 0-indexed: Col D is the 3rd col in B:D, so index 2

  const HEADER_ROWS = 1;      // Number of header rows to skip for data processing

  const HEADERS_CONFIG = [
    { column: FULL_NAME_COL, text: "Full Name" },
    { column: AGE_COL, text: "Age" },
    { column: AGE_GROUP_COL, text: "Age Group" },
    { column: LAST_VISIT_COL, text: "Last 2025 Visit" },
    { column: MEMBERSHIP_YEAR_COL, text: "Membership Year" },
    { column: MEMBERSHIP_QUARTER_COL, text: "Membership Quarter" },
    { column: MEMBERSHIP_MONTH_COL, text: "Membership Month" },
    { column: MEMBERSHIP_WEEK_COL, text: "Membership Week" }
  ];
  // --- End Configuration ---

  // Check if the event object and range are valid
  if (!e || !e.range) {
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();

  // 1. Check if the edit is on the target sheet
  if (sheet.getName() !== TARGET_SHEET_NAME) {
    return;
  }

  // Ensure headers are present
  ensureHeaders(sheet, HEADERS_CONFIG, HEADER_ROWS);

  const firstEditedRow = range.getRow();
  const lastEditedRow = range.getLastRow();
  const firstEditedCol = range.getColumn();
  const lastEditedCol = range.getLastColumn();

  // Cache for VLOOKUP data for this onEdit run. Cleared for each event.
  let visitDataCacheForThisRun = null;

  for (let r = firstEditedRow; r <= lastEditedRow; r++) {
    if (r <= HEADER_ROWS) {
      continue;
    }

    const nameColsAffectedForRow = (firstEditedCol <= Math.max(LAST_NAME_COL, FIRST_NAME_COL) && lastEditedCol >= Math.min(LAST_NAME_COL, FIRST_NAME_COL));
    const dobColAffectedForRow = (firstEditedCol <= DOB_COL && lastEditedCol >= DOB_COL);
    const membershipDateColAffectedForRow = (firstEditedCol <= MEMBERSHIP_DATE_COL && lastEditedCol >= MEMBERSHIP_DATE_COL);
    const fullNameColDirectlyAffectedForRow = (firstEditedCol <= FULL_NAME_COL && lastEditedCol >= FULL_NAME_COL);
    
    let fullNameWasUpdatedByScriptThisRow = false;

    // --- 1. Full Name Combination ---
    if (nameColsAffectedForRow) {
      const lastNameValue = sheet.getRange(r, LAST_NAME_COL).getValue();
      const firstNameValue = sheet.getRange(r, FIRST_NAME_COL).getValue();
      const lastNameTrimmed = String(lastNameValue).trim();
      const firstNameTrimmed = String(firstNameValue).trim();
      const currentFullNameCell = sheet.getRange(r, FULL_NAME_COL);
      let newFullName = "";

      if (firstNameTrimmed !== "" && lastNameTrimmed !== "") {
        newFullName = `${firstNameTrimmed} ${lastNameTrimmed}`;
      }

      if (currentFullNameCell.getValue() !== newFullName) {
        currentFullNameCell.setValue(newFullName);
        fullNameWasUpdatedByScriptThisRow = true;
      }
    }

    // --- 2. Age and Age Group Calculation ---
    if (dobColAffectedForRow) {
      const dobValue = sheet.getRange(r, DOB_COL).getValue();
      if (dobValue instanceof Date && !isNaN(dobValue)) {
        const age = calculateAge(dobValue);
        sheet.getRange(r, AGE_COL).setValue(age);
        if (age !== "" && typeof age === 'number' && age >= 0) {
          const ageGroup = getAgeGroup(age);
          sheet.getRange(r, AGE_GROUP_COL).setValue(ageGroup);
        } else {
          sheet.getRange(r, AGE_GROUP_COL).setValue("");
        }
      } else {
        sheet.getRange(r, AGE_COL).setValue("");
        sheet.getRange(r, AGE_GROUP_COL).setValue("");
      }
    }

    // --- 3. Membership Date Parts Calculation ---
    if (membershipDateColAffectedForRow) {
      const memDateValue = sheet.getRange(r, MEMBERSHIP_DATE_COL).getValue();
      if (memDateValue instanceof Date && !isNaN(memDateValue)) {
        sheet.getRange(r, MEMBERSHIP_YEAR_COL).setValue(memDateValue.getFullYear());
        sheet.getRange(r, MEMBERSHIP_QUARTER_COL).setValue(Math.floor(memDateValue.getMonth() / 3) + 1);
        sheet.getRange(r, MEMBERSHIP_MONTH_COL).setValue(Utilities.formatDate(memDateValue, Session.getScriptTimeZone(), "MMMM"));
        sheet.getRange(r, MEMBERSHIP_WEEK_COL).setValue(getIsoWeek(memDateValue));
      } else {
        sheet.getRange(r, MEMBERSHIP_YEAR_COL).setValue("");
        sheet.getRange(r, MEMBERSHIP_QUARTER_COL).setValue("");
        sheet.getRange(r, MEMBERSHIP_MONTH_COL).setValue("");
        sheet.getRange(r, MEMBERSHIP_WEEK_COL).setValue("");
      }
    }
    
    // --- 4. VLOOKUP for "Last 2025 Visit" ---
    if (fullNameWasUpdatedByScriptThisRow || fullNameColDirectlyAffectedForRow) {
      const fullNameForLookup = sheet.getRange(r, FULL_NAME_COL).getValue();
      if (String(fullNameForLookup).trim() !== "") {
        if (visitDataCacheForThisRun === null) { 
          visitDataCacheForThisRun = getVisitDataArray(VISIT_SHEET_NAME, VISIT_DATA_RANGE_A1);
        }
        if (visitDataCacheForThisRun && visitDataCacheForThisRun.length > 0) {
          const lookupResult = performScriptVlookup(
            fullNameForLookup, 
            visitDataCacheForThisRun, 
            VISIT_KEY_COL_IDX_IN_DATA, 
            VISIT_RETURN_COL_IDX_IN_DATA
          );
          sheet.getRange(r, LAST_VISIT_COL).setValue(lookupResult);
        } else {
          sheet.getRange(r, LAST_VISIT_COL).setValue(""); 
        }
      } else {
        sheet.getRange(r, LAST_VISIT_COL).setValue("");
      }
    }
  } 
}

// --- Helper Functions (These do not need any changes) ---

function ensureHeaders(sheet, headerConfig, headerRow) {
  if (headerRow < 1) headerRow = 1;
  headerConfig.forEach(conf => {
    try {
      const headerCell = sheet.getRange(headerRow, conf.column);
      if (headerCell.getValue() !== conf.text) {
        headerCell.setValue(conf.text);
      }
    } catch (e) {
      // Error logging can be added here if needed
    }
  });
}

function calculateAge(birthDate) {
  if (!(birthDate instanceof Date) || isNaN(birthDate)) return "";
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  birthDate.setHours(0, 0, 0, 0);

  if (birthDate > today) return "";

  let age = today.getFullYear() - birthDate.getFullYear();
  const monthDiff = today.getMonth() - birthDate.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }
  return age >= 0 ? age : "";
}

function getAgeGroup(age) {
  if (typeof age !== 'number' || age < 0) return "";
  if (age < 19) return "0-18";
  if (age < 40) return "19-39";
  if (age < 60) return "40-59";
  if (age < 120) return "60+";
  return "";
}

function getIsoWeek(date) {
  if (!(date instanceof Date) || isNaN(date)) return "";
  const d = new Date(date.valueOf());
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 3 - (d.getDay() + 6) % 7);
  const week1 = new Date(d.getFullYear(), 0, 4);
  return 1 + Math.round(((d.getTime() - week1.getTime()) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
}

function getVisitDataArray(sheetName, rangeA1) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
      return sheet.getRange(rangeA1).getValues();
    } else {
      return null;
    }
  } catch (e) {
    return null;
  }
}

function performScriptVlookup(searchValue, dataArray, searchColIndex, returnColIndex) {
  if (!dataArray || dataArray.length === 0) return "";
  for (let i = 0; i < dataArray.length; i++) {
    if (dataArray[i][searchColIndex] == searchValue) {
      return dataArray[i][returnColIndex];
    }
  }
  return "";
}