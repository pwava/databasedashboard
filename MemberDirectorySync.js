/**
 * Main handler for the onFormSubmit trigger. Processes new member data,
 * assigns IDs, and syncs to the master Directory if the member is not a duplicate.
 */
function processNewMemberForm(e) {
  Logger.log("MASTER onFormSubmit - processNewMemberForm triggered.");

  // --- Configuration ---
  // -- ⭐️ CHANGED LINES START HERE ⭐️ --
  const directorySheetName = getSetting('Master Directory Tab'); // Was 'Directory'
  const newMemberFormSheetName = getSetting('New Member Form Tab'); // Was 'New Member Form'
  // -- ⭐️ CHANGED LINES END HERE ⭐️ --


  // Incoming e.values indices from the FORM SUBMISSION
  const formLastNameIndex = 0; // Corresponds to the form question for Last Name
  const formFirstNameIndex = 1; // Corresponds to the form question for First Name
  const formTimestampIndex = 10; // Corresponds to the form question for Timestamp

  // New Member Form Sheet Column Numbers (1-based, as seen on the actual sheet)
  const newMemberFormSheetIdCol = 1; // Column A
  const newMemberFormSheetFullNameCol = 2; // Column B
  const newMemberFormSheetLastNameCol = 3; // Column C
  const newMemberFormSheetFirstNameCol = 4; // Column D
  const newMemberFormSheetTimestampCol = 21; // Column U

  // Directory Sheet Column Numbers (1-based)
  const directoryIdCol = 1; // Column A
  const directoryFullNameCol = 2; // Column B
  const directoryLastNameCol = 3; // Column C
  const directoryFirstNameCol = 4; // Column D
  const directoryTimestampCol = 22; // Column V
  const directoryKeyColumnForEmptinessCheck = 3; // Using Last Name in Directory to find next blank row

  const startRowForDirectorySearchAndWrite = 2; // Assuming Row 1 is header in 'Directory'
  const expectedFormResponseLength = 11;
  // --- End Configuration ---

  if (!e || !e.values || !e.range) {
    Logger.log("Error: Script was triggered without complete form data (missing e.values or e.range).");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newMemberFormSheet = ss.getSheetByName(newMemberFormSheetName);
  const directorySheet = ss.getSheetByName(directorySheetName);

  if (!newMemberFormSheet) {
    Logger.log(`Error: Sheet named '${newMemberFormSheetName}' not found based on settings. Please check 'Config for Urls'.`);
    return;
  }
  if (!directorySheet) {
    Logger.log(`Error: Sheet named '${directorySheetName}' not found based on settings. Please check 'Config for Urls'.`);
    return;
  }

  const formResponses = e.values;
  const lastRowOfSubmission = e.range.getRow();

  if (formResponses.length < expectedFormResponseLength) {
    Logger.log(`Error: Form response does not contain enough data. Expected ${expectedFormResponseLength}, received ${formResponses.length}.`);
    return;
  }

  // Step 1: Combine names (assuming combineNamesOnFormSubmit exists and works on the event object)
  combineNamesOnFormSubmit(e);

  // Step 2: Assign Member ID (assuming assignMemberIds exists and works on the event object)
  assignMemberIds(e);

  // Step 3: Duplicate Check and Copy to Directory
  const updatedNewMemberFormRowData = newMemberFormSheet.getRange(lastRowOfSubmission, 1, 1, newMemberFormSheetFirstNameCol).getValues()[0];
  const newMemberId = updatedNewMemberFormRowData[newMemberFormSheetIdCol - 1];
  const newMemberFullName = updatedNewMemberFormRowData[newMemberFormSheetFullNameCol - 1];
  const newMemberLastNameFromSheet = updatedNewMemberFormRowData[newMemberFormSheetLastNameCol - 1];
  const newMemberFirstNameFromSheet = updatedNewMemberFormRowData[newMemberFormSheetFirstNameCol - 1];

  const newMemberLastNameLower = String(newMemberLastNameFromSheet || "").toLowerCase().trim();
  const newMemberFirstNameLower = String(newMemberFirstNameFromSheet || "").toLowerCase().trim();

  if (!newMemberLastNameLower || !newMemberFirstNameLower) {
    Logger.log(`Warning: Last Name or First Name is empty for row ${lastRowOfSubmission}. Skipping Directory copy.`);
    return;
  }

  const lastRowInDirectoryWithContent = directorySheet.getLastRow();
  let matchFound = false;

  if (lastRowInDirectoryWithContent >= startRowForDirectorySearchAndWrite) {
    const directoryDataRange = directorySheet.getRange(startRowForDirectorySearchAndWrite, directoryLastNameCol, lastRowInDirectoryWithContent - (startRowForDirectorySearchAndWrite - 1), 2);
    const directoryData = directoryDataRange.getValues();

    for (let i = 0; i < directoryData.length; i++) {
      const directoryLastName = directoryData[i][0];
      const directoryFirstName = directoryData[i][1];
      const processedDirectoryLastName = directoryLastName ? String(directoryLastName).toLowerCase().trim() : '';
      const processedDirectoryFirstName = directoryFirstName ? String(directoryFirstName).toLowerCase().trim() : '';

      if (processedDirectoryLastName === newMemberLastNameLower && processedDirectoryFirstName === newMemberFirstNameLower) {
        matchFound = true;
        break;
      }
    }
  }

  if (matchFound) {
    Logger.log(`Result: A combined match for "${newMemberFirstNameFromSheet} ${newMemberLastNameFromSheet}" already exists. No new entry in Directory.`);
    return;
  }

  Logger.log(`Result: No combined match found. Proceeding to copy data to Directory.`);

  let nextRowInDirectory = startRowForDirectorySearchAndWrite;
  while (directorySheet.getRange(nextRowInDirectory, directoryKeyColumnForEmptinessCheck).getValue() !== "") {
    nextRowInDirectory++;
    if (nextRowInDirectory > directorySheet.getMaxRows() + 5) {
      Logger.log("Error: Exceeded maximum row search limit for finding an empty row in Directory sheet.");
      return;
    }
  }

  const userProvidedTimestampValue = formResponses[formTimestampIndex];
  let formattedDate = "";
  if (userProvidedTimestampValue) {
    try {
      const dateObject = new Date(userProvidedTimestampValue);
      if (!isNaN(dateObject.getTime())) {
        formattedDate = Utilities.formatDate(dateObject, Session.getScriptTimeZone(), "MM/dd/yyyy");
      } else {
        formattedDate = userProvidedTimestampValue;
      }
    } catch (error) {
      formattedDate = userProvidedTimestampValue;
    }
  }

  try {
    const fullNewMemberFormRow = newMemberFormSheet.getRange(lastRowOfSubmission, 1, 1, newMemberFormSheetTimestampCol).getValues()[0];
    const directoryNewRow = new Array(22).fill("");

    const safeGetString = (value) => (value !== undefined && value !== null ? String(value).trim() : "");
    const safeGetStringAndUpper = (value) => (value !== undefined && value !== null ? String(value).trim().toUpperCase() : "");

    directoryNewRow[directoryIdCol - 1] = fullNewMemberFormRow[newMemberFormSheetIdCol - 1];
    directoryNewRow[directoryFullNameCol - 1] = fullNewMemberFormRow[newMemberFormSheetFullNameCol - 1];
    directoryNewRow[directoryLastNameCol - 1] = fullNewMemberFormRow[newMemberFormSheetLastNameCol - 1];
    directoryNewRow[directoryFirstNameCol - 1] = fullNewMemberFormRow[newMemberFormSheetFirstNameCol - 1];
    directoryNewRow[4] = safeGetString(formResponses[2]); // Street
    directoryNewRow[5] = safeGetString(formResponses[3]); // City
    directoryNewRow[6] = safeGetStringAndUpper(formResponses[4]); // State
    directoryNewRow[7] = safeGetString(formResponses[5]); // ZIP
    directoryNewRow[9] = safeGetString(formResponses[6]); // Phone
    directoryNewRow[10] = safeGetString(formResponses[7]); // Email
    directoryNewRow[13] = safeGetString(formResponses[8]); // Notes
    directoryNewRow[14] = safeGetString(formResponses[9]); // Status
    directoryNewRow[directoryTimestampCol - 1] = formattedDate; // Timestamp

    directorySheet.getRange(nextRowInDirectory, 1, 1, directoryNewRow.length).setValues([directoryNewRow]);
    Logger.log(`Success: New member data copied to '${directorySheetName}' sheet, row ${nextRowInDirectory}.`);
  } catch (error) {
    Logger.log(`Error: Error writing data to '${directorySheetName}'. Error: ${error.message}`);
  }
}

/**
 * This is a test function for processNewMemberForm and is not affected by the config changes,
 * as it operates on the same sheet. It can remain as is.
 */
function testProcessNewMemberForm() {
    // Test function code remains unchanged...
}