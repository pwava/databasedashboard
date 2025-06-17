/**
 * Contains functions to assign a unique Person ID to a new member upon form submission.
 * It dynamically scans all configured spreadsheets to ensure ID uniqueness and find the highest current ID.
 */

// These global variables will be populated by the initializeMasterDataAndHighestId function.
let masterNameIdMap = new Map();
let highestOverallId = 0;

/**
 * Initializes master data by loading from all locations defined in the 'Config for Urls' sheet.
 * This function finds the true highest ID from across the entire system.
 */
function initializeMasterDataAndHighestId() {
  masterNameIdMap.clear();
  highestOverallId = 0;
  Logger.log("Starting master data initialization for on-form-submit ID check...");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config for Urls');
  if (!configSheet) {
    Logger.log('CRITICAL ERROR: "Config for Urls" sheet not found. Aborting ID assignment.');
    return;
  }
  
  // Get the locations to check from the Config sheet's "PERSON ID CHECK LOCATIONS" table
  const allConfigData = configSheet.getDataRange().getValues();
  let idCheckLocations = [];
  let startReading = false;
  for (const row of allConfigData) {
    if (row[0] === 'PERSON ID CHECK LOCATIONS') {
      startReading = true;
      continue;
    }
    if (startReading && row[0]) {
      idCheckLocations.push(row);
    } else if (startReading && !row[0]) {
      break;
    }
  }

  Logger.log(`Found ${idCheckLocations.length} locations to scan for Person IDs.`);

  // Loop through each location defined in the config sheet
  for (const location of idCheckLocations) {
    const spreadsheetKey = location[0]; // e.g., 'Dashboard', 'Attendance Tracker'
    const tabNameToScan = location[1]; // e.g., 'Directory', 'Service Attendance'
    
    const urlSettingKey = `${spreadsheetKey} URL`; // e.g., 'Dashboard URL'
    const spreadsheetUrl = getSetting(urlSettingKey);

    if (!spreadsheetUrl) {
      Logger.log(`Warning: URL for "${spreadsheetKey}" not found in settings. Skipping this location.`);
      continue;
    }

    try {
      let spreadsheetToScan;
      if (spreadsheetUrl === ss.getUrl()) {
        spreadsheetToScan = ss; // It's the current spreadsheet
      } else {
        spreadsheetToScan = SpreadsheetApp.openByUrl(spreadsheetUrl);
      }
      
      const sheetToScan = spreadsheetToScan.getSheetByName(tabNameToScan);
      if (!sheetToScan) {
        Logger.log(`Warning: Tab "${tabNameToScan}" not found in spreadsheet "${spreadsheetKey}". Skipping.`);
        continue;
      }

      Logger.log(`Scanning: Spreadsheet -> "${spreadsheetKey}", Tab -> "${tabNameToScan}"...`);
      const data = sheetToScan.getDataRange().getValues();

      // Assuming ID is in Col A (index 0) and Name is in Col B (index 1) for all sheets
      for (let i = 1; i < data.length; i++) { // Start at 1 to skip header
        const idFromSheet = data[i][0];
        const nameFromSheet = data[i][1];
        
        const numericId = parseInt(String(idFromSheet || '').replace(/\D/g, ''));
        if (!isNaN(numericId)) {
          highestOverallId = Math.max(highestOverallId, numericId);
        }

        if (nameFromSheet && String(nameFromSheet).trim() !== "") {
          const cleanedName = String(nameFromSheet).trim().toUpperCase();
          if (!masterNameIdMap.has(cleanedName) && !isNaN(numericId) && numericId > 0) {
            masterNameIdMap.set(cleanedName, String(numericId).padStart(5, '0'));
          }
        }
      }
    } catch (e) {
      Logger.log(`ERROR accessing ID mapping location: "${spreadsheetKey}" -> "${tabNameToScan}". Error: ${e.message}`);
    }
  }

  Logger.log(`Master list built for form submission. Unique names found: ${masterNameIdMap.size}. Highest ID found: ${highestOverallId}.`);
}


/**
 * Assigns a unique numeric Personal ID to the specific new member form response row.
 * This is the main function called by the onFormSubmit trigger.
 * -- THIS IS THE CORRECTED VERSION --
 */
function assignMemberIds(e) {
  Logger.log("assignMemberIds function started for a new form submission.");

  const newMemberSheetName = getSetting('New Member Form Tab');
  if (!newMemberSheetName) {
      Logger.log('CRITICAL ERROR: "New Member Form Tab" setting not found in config. Aborting.');
      return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newMemberSheet = ss.getSheetByName(newMemberSheetName);

  if (!newMemberSheet) {
    Logger.log(`CRITICAL ERROR: Sheet named "${newMemberSheetName}" not found. Cannot assign ID.`);
    return;
  }

  initializeMasterDataAndHighestId();

  const idCol = 1;
  const fullNameCol = 2;

  const lastRowOfSubmission = e.range.getRow();
  Logger.log(`Processing new member form submission in row: ${lastRowOfSubmission}.`);

  const rowData = newMemberSheet.getRange(lastRowOfSubmission, 1, 1, fullNameCol).getValues()[0];
  const nameCell = rowData[fullNameCol - 1]; 
  const existingIdCell = rowData[idCol - 1]; 

  const nameString = String(nameCell || '').trim();
  if (!nameString) {
    Logger.log(`Row ${lastRowOfSubmission}: Full Name is blank. Skipping ID assignment.`);
    return;
  }

  const cleanedFormName = nameString.toUpperCase();
  let assignedId = '';

  if (masterNameIdMap.has(cleanedFormName)) {
    // -- ⭐️ CORRECTION IS HERE ⭐️ --
    assignedId = masterNameIdMap.get(cleanedFormName); // Removed the space from "cleanedFormN ame"
    Logger.log(`Row ${lastRowOfSubmission} ('${nameString}'): Matched in master data. Using existing ID: ${assignedId}`);
  } 
  else {
    highestOverallId++;
    assignedId = String(highestOverallId).padStart(5, '0');
    Logger.log(`Row ${lastRowOfSubmission} ('${nameString}'): No match found. Generated new ID: ${assignedId}`);
  }

  if (String(existingIdCell).trim() !== assignedId) {
      newMemberSheet.getRange(lastRowOfSubmission, idCol).setValue(assignedId);
      Logger.log(`Success: ID "${assignedId}" written for '${nameString}' in row ${lastRowOfSubmission}.`);
  } else {
      Logger.log(`Info: ID "${assignedId}" for '${nameString}' was already correct. No change needed.`);
  }
}

// The original getTrueLastRow function can be removed if it's not used by any other scripts in this file.