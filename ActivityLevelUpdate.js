
/**
 * Contains functions to sync data (like Activity Level and Giving Level) from other spreadsheets
 * into the main Directory sheet. This script now gets all its settings from the central config.
 */

function updateActivityLevels() {
  Logger.log("Starting updateActivityLevels with permanent fix...");

  // Get all settings directly from PropertiesService, bypassing the getSetting() function.
  const props = PropertiesService.getScriptProperties();
  const attendanceTrackerUrl = props.getProperty('Attendance Tracker URL');
  const directorySheetName = props.getProperty('Master Directory Tab');
  const attendanceStatsSheetName = props.getProperty('Attendance Stats Tab');

  // Check if the essential settings were retrieved correctly.
  if (!attendanceTrackerUrl) {
    Logger.log('Error: Could not retrieve Attendance Tracker URL from properties.');
    return;
  }
  if (!directorySheetName) {
    Logger.log('Error: Could not retrieve Master Directory Tab from properties.');
    return;
  }
   if (!attendanceStatsSheetName) {
    Logger.log('Error: Could not retrieve Attendance Stats Tab from properties.');
    return;
  }

  const directorySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const directorySheet = directorySpreadsheet.getSheetByName(directorySheetName);

  if (!directorySheet) {
    Logger.log(`Error: Sheet named "${directorySheetName}" not found in this spreadsheet.`);
    return;
  }

  let sourceSpreadsheet;
  try {
    sourceSpreadsheet = SpreadsheetApp.openByUrl(attendanceTrackerUrl);
  } catch (e) {
    Logger.log(`Error: Could not open Attendance Tracker spreadsheet. URL: ${attendanceTrackerUrl}. Error: ${e.message}`);
    return;
  }

  const attendanceSheet = sourceSpreadsheet.getSheetByName(attendanceStatsSheetName);
  if (!attendanceSheet) {
    Logger.log(`Error: Sheet named "${attendanceStatsSheetName}" not found in the Attendance Tracker spreadsheet.`);
    return;
  }

  const directoryValues = directorySheet.getDataRange().getValues();
  const attendanceValues = attendanceSheet.getDataRange().getValues();

  if (attendanceValues.length < 2) {
    Logger.log(`Info: No data (beyond headers) found in "${attendanceStatsSheetName}".`);
    return;
  }

  // --- The rest of the original function continues unchanged ---
  const attendanceMap = new Map();
  for (let i = 1; i < attendanceValues.length; i++) {
    const personId = String(attendanceValues[i][0] || "").trim();
    const fullName = String(attendanceValues[i][1] || "").trim().toLowerCase();
    const activityLevel = attendanceValues[i][11];
    if (personId && fullName) {
      attendanceMap.set(`${personId}_${fullName}`, activityLevel);
    }
  }

  if (attendanceMap.size === 0) {
    Logger.log('Info: No valid data to map in "Attendance Stats" sheet.');
    return;
  }

  let updatesMade = 0;
  const newActivityLevelsForDirectory = [];
  const backgroundsForDirectory = [];
  const numColumns = directoryValues[0].length;
  const LIGHT_GRAY = '#efefef';

  for (let i = 0; i < directoryValues.length; i++) {
    if (i === 0) {
      newActivityLevelsForDirectory.push([directoryValues[i][19]]);
      backgroundsForDirectory.push(Array(numColumns).fill(null));
      continue;
    }

    const dirRow = directoryValues[i];
    const personId = String(dirRow[0] || "").trim();
    const fullName = String(dirRow[1] || "").trim().toLowerCase();
    const currentActivityValueInDir = dirRow[19];
    let newActivityValue = currentActivityValueInDir;
    let backgroundRow = Array(numColumns).fill(null);

    if (personId && fullName) {
      const lookupKey = `${personId}_${fullName}`;
      if (attendanceMap.has(lookupKey)) {
        newActivityValue = attendanceMap.get(lookupKey);
      } else {
        newActivityValue = 'Archive';
        backgroundRow = Array(numColumns).fill(LIGHT_GRAY);
      }
    }

    if (newActivityValue !== currentActivityValueInDir) {
      updatesMade++;
    }

    newActivityLevelsForDirectory.push([newActivityValue]);
    backgroundsForDirectory.push(backgroundRow);
  }

  if (updatesMade > 0) {
    directorySheet.getRange(1, 20, newActivityLevelsForDirectory.length, 1).setValues(newActivityLevelsForDirectory);
    directorySheet.getRange(1, 1, backgroundsForDirectory.length, numColumns).setBackgrounds(backgroundsForDirectory);
    Logger.log(`Success: ${updatesMade} activity levels and/or backgrounds updated in "${directorySheetName}".`);
  } else {
    Logger.log(`Info: No matching records required an update for activity levels.`);
  }
}

function updateGivingLevelsFromDonorStats() {
  Logger.log("Starting updateGivingLevelsFromDonorStats...");

  // Get URLs and Tab Names from the central config
  const donorDataUrl = getSetting('Donation Data URL');
  const directorySheetName = getSetting('Master Directory Tab');
  const donorStatsSheetName = getSetting('Donor Stats Tab');

  if (!donorDataUrl) {
    Logger.log('Error: updateGivingLevels - Donation Data URL not set in Config.');
    return;
  }

  const directorySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const directorySheet = directorySpreadsheet.getSheetByName(directorySheetName);

  if (!directorySheet) {
    Logger.log(`Error: updateGivingLevels - Sheet named "${directorySheetName}" not found in this spreadsheet.`);
    return;
  }

  let sourceDonorStatsSpreadsheet;
  try {
    sourceDonorStatsSpreadsheet = SpreadsheetApp.openByUrl(donorDataUrl);
  } catch (e) {
    Logger.log(`Error: updateGivingLevels - Could not open Donation Data spreadsheet. URL: ${donorDataUrl}. Error: ${e.message}`);
    return;
  }

  const donorStatsSheet = sourceDonorStatsSpreadsheet.getSheetByName(donorStatsSheetName);
  if (!donorStatsSheet) {
    Logger.log(`Error: updateGivingLevels - Sheet named "${donorStatsSheetName}" not found in the Donation Data spreadsheet.`);
    return;
  }

  const directoryValues = directorySheet.getDataRange().getValues();
  const donorStatsValues = donorStatsSheet.getDataRange().getValues();

  if (donorStatsValues.length < 2) {
    Logger.log(`Info: updateGivingLevels - "${donorStatsSheetName}" sheet has no data beyond header.`);
    return;
  }
  
  // Create a map for quick lookups: "PersonID||firstname||lastname" -> "Giving Level"
  const donorStatsMap = new Map();
  for (let i = 1; i < donorStatsValues.length; i++) {
    const personId = String(donorStatsValues[i][0] || "").trim(); // DS_MATCH_PERSON_ID_COL_IDX = 0
    const firstName = String(donorStatsValues[i][1] || "").trim().toLowerCase(); // DS_MATCH_FIRST_NAME_COL_IDX = 1
    const lastName = String(donorStatsValues[i][2] || "").trim().toLowerCase(); // DS_MATCH_LAST_NAME_COL_IDX = 2
    const givingLevel = donorStatsValues[i][9]; // DS_GIVING_LEVEL_COL_J_IDX = 9
    if (personId && firstName && lastName) {
      donorStatsMap.set(`${personId}||${firstName}||${lastName}`, givingLevel);
    }
  }

  if (donorStatsMap.size === 0) {
    Logger.log(`Error: updateGivingLevels - No valid data in "${donorStatsSheetName}" to create lookup map.`);
    return;
  }
  
  let updatesMade = 0;
  const newGivingLevelsForDirectory = directoryValues.map((dirRow, i) => {
      if (i === 0) return [dirRow[18]]; // Return header as-is. DIR_GIVING_LEVEL_COL_S_IDX = 18

      const personIdDir = String(dirRow[0] || "").trim(); // DIR_FOR_DONOR_MATCH_PERSON_ID_COL_IDX = 0
      const lastNameDir = String(dirRow[2] || "").trim().toLowerCase(); // DIR_FOR_DONOR_MATCH_LAST_NAME_COL_IDX = 2
      const firstNameDir = String(dirRow[3] || "").trim().toLowerCase(); // DIR_FOR_DONOR_MATCH_FIRST_NAME_COL_IDX = 3
      const currentGivingLevelInDir = dirRow[18]; // DIR_GIVING_LEVEL_COL_S_IDX = 18

      if (personIdDir && firstNameDir && lastNameDir) {
          const lookupKey = `${personIdDir}||${firstNameDir}||${lastNameDir}`;
          if (donorStatsMap.has(lookupKey)) {
              const givingLevelFromStats = donorStatsMap.get(lookupKey);
              if (givingLevelFromStats !== currentGivingLevelInDir) {
                  updatesMade++;
              }
              return [givingLevelFromStats]; // Return the new value
          }
      }
      return [currentGivingLevelInDir]; // Return original value if no match
  });

  if (updatesMade > 0) {
    directorySheet.getRange(1, 19, newGivingLevelsForDirectory.length, 1).setValues(newGivingLevelsForDirectory); // DIR_GIVING_LEVEL_COL_S_IDX + 1 = 19
    Logger.log(`Success: updateGivingLevels - ${updatesMade} giving levels updated in "${directorySheetName}".`);
  } else {
    Logger.log(`Info: updateGivingLevels - No matching records required an update for giving levels.`);
  }
}
/**
 * A temporary debug function to see all saved properties.
 */
function checkScriptProperties() {
  const properties = PropertiesService.getScriptProperties().getProperties();
  Logger.log('--- START OF SAVED PROPERTIES ---');
  Logger.log(properties);
  Logger.log('--- END OF SAVED PROPERTIES ---');
}
