/**
 * Assigns/Updates IDs in the 'Directory' sheet by building a master list from all configured locations.
 * This function is designed to be run on an automated trigger.
 */
function assignIdsInDirectory() {
  Logger.log("Starting automated ID assignment in 'Directory' sheet...");

  // Get the main Directory sheet from the current spreadsheet
  const directorySheetName = getSetting('Master Directory Tab');
  if (!directorySheetName) {
    Logger.log('CRITICAL ERROR: "Master Directory Tab" setting not found in config. Aborting.');
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const directorySheet = ss.getSheetByName(directorySheetName);

  if (!directorySheet) {
    Logger.log(`CRITICAL ERROR: Sheet named "${directorySheetName}" not found in this spreadsheet. Aborting.`);
    return;
  }

  // --- 1. Build Master Name/ID Map from all configured locations ---
  
  const nameToIdMap = new Map(); // Stores: cleanedName.toUpperCase() -> formattedId
  let highestId = 0;

  // Get the locations to check from the Config sheet
  const configSheet = ss.getSheetByName('Config for Urls');
  if (!configSheet) {
    Logger.log('CRITICAL ERROR: "Config for Urls" sheet not found. Aborting.');
    return;
  }
  
  // Find the start of the PERSON ID CHECK LOCATIONS table
  const allConfigData = configSheet.getDataRange().getValues();
  let idCheckLocations = [];
  let startReading = false;
  for (const row of allConfigData) {
    if (row[0] === 'PERSON ID CHECK LOCATIONS') {
      startReading = true;
      continue; // Skip the header row itself
    }
    if (startReading && row[0]) { // If we've started and the key column isn't blank
      idCheckLocations.push(row);
    } else if (startReading && !row[0]) {
      break; // Stop at the first blank row after the section starts
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
        spreadsheetToScan = ss; // It's the current spreadsheet, no need to open by URL
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
          highestId = Math.max(highestId, numericId);
        }

        if (nameFromSheet && String(nameFromSheet).trim() !== "") {
          const cleanedName = String(nameFromSheet).trim().toUpperCase();
          if (!nameToIdMap.has(cleanedName) && !isNaN(numericId) && numericId > 0) {
            nameToIdMap.set(cleanedName, String(numericId).padStart(5, '0'));
          }
        }
      }
    } catch (e) {
      Logger.log(`ERROR accessing location: "${spreadsheetKey}" -> "${tabNameToScan}". Error: ${e.message}`);
    }
  }

  Logger.log(`Master list built. Total unique names with IDs found: ${nameToIdMap.size}. Highest ID found: ${highestId}.`);

  // --- 2. Process the Directory Sheet to Assign/Update IDs ---
  
  const headerRows = directorySheet.getFrozenRows() || 1;
  const lastRow = directorySheet.getLastRow();
  if (lastRow <= headerRows) {
    Logger.log("Directory sheet has no data to process. Exiting.");
    return;
  }

  const directoryRange = directorySheet.getRange(headerRows + 1, 1, lastRow - headerRows, 2);
  const directoryData = directoryRange.getValues();
  const outputIds = [];

  let newIdsAssignedCount = 0;
  let idsUpdatedCount = 0;

  for(let i = 0; i < directoryData.length; i++) {
    const existingId = String(directoryData[i][0] || '').trim();
    const name = String(directoryData[i][1] || '').trim();
    let idToSet = existingId;

    if (name) {
      const cleanedName = name.toUpperCase();
      if (nameToIdMap.has(cleanedName)) {
        idToSet = nameToIdMap.get(cleanedName); // Use existing ID from master list
      } else {
        // Name not found anywhere, generate a brand new ID
        highestId++;
        idToSet = String(highestId).padStart(5, '0');
        nameToIdMap.set(cleanedName, idToSet); // Add to map to avoid re-assigning in this run
        newIdsAssignedCount++;
      }
    }

    if (idToSet !== existingId) {
      idsUpdatedCount++;
    }
    outputIds.push([idToSet]);
  }

  // Write all the updated IDs back to the sheet in one go
  if (idsUpdatedCount > 0) {
    directoryRange.offset(0, 0, outputIds.length, 1).setValues(outputIds);
    Logger.log(`ID processing complete. IDs updated/written: ${idsUpdatedCount}. New IDs generated: ${newIdsAssignedCount}.`);
  } else {
    Logger.log("ID processing complete. No updates were needed.");
  }
}

// NOTE: The getTrueLastRow function is not used in this version but can be kept if other scripts need it.
// NOTE: The onOpenDirectoryIdAssigner function has been removed as requested.