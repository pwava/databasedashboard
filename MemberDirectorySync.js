/**
 * Main handler for the onFormSubmit trigger. Processes new member data,
 * assigns IDs, and syncs to the master Directory. If configured, it also
 * syncs the data to a second, external Directory spreadsheet with a potentially
 * different column layout.
 */
function processNewMemberForm(e) {
  Logger.log("MASTER onFormSubmit - processNewMemberForm triggered.");

  // --- Configuration ---
  const directorySheetName = getSetting('Master Directory Tab');
  const newMemberFormSheetName = getSetting('New Member Form Tab');
  const directory2Url = getSetting('Directory 2 URL');
  const directory2TabName = getSetting('Directory 2 Tab Name');

  // Incoming e.values indices from the FORM SUBMISSION
  const formTimestampIndex = 10;
  // Other form indices for reference:
  // 0:LastName, 1:FirstName, 2:Street, 3:City, 4:State, 5:ZIP, 6:Gender, 7:DOB, 8:Phone, 10:Timestamp, 11:Email
  const formLineageIndex = 12; // New Lineage question added to the form

  // New Member Form Sheet Column Numbers (1-based)
  const newMemberFormSheetIdCol = 1;
  const newMemberFormSheetFullNameCol = 2;
  const newMemberFormSheetLastNameCol = 3;
  const newMemberFormSheetFirstNameCol = 4;
  const newMemberFormSheetTimestampCol = 21;

  // Directory 1 Sheet Column Numbers (1-based)
  const directoryLastNameCol = 3;
  const directoryKeyColumnForEmptinessCheck = 3;

  const startRowForDirectorySearchAndWrite = 2;
  const expectedFormResponseLength = 11;
  // --- End Configuration ---

  if (!e || !e.values || !e.range) {
    Logger.log("Error: Script was triggered without complete form data.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newMemberFormSheet = ss.getSheetByName(newMemberFormSheetName);
  const directorySheet = ss.getSheetByName(directorySheetName);

  if (!newMemberFormSheet || !directorySheet) {
    Logger.log(`Error: A required sheet was not found. Please check your settings.`);
    return;
  }

  const formResponses = e.values;
  const lastRowOfSubmission = e.range.getRow();

  // Steps 1 & 2: Combine names and Assign IDs
  combineNamesOnFormSubmit(e);
  assignMemberIds(e);

  // Step 3: Duplicate Check in Directory 1
  const updatedNewMemberFormRowData = newMemberFormSheet.getRange(lastRowOfSubmission, 1, 1, newMemberFormSheetFirstNameCol).getValues()[0];
  const newMemberLastNameLower = String(updatedNewMemberFormRowData[newMemberFormSheetLastNameCol - 1] || "").toLowerCase().trim();
  const newMemberFirstNameLower = String(updatedNewMemberFormRowData[newMemberFormSheetFirstNameCol - 1] || "").toLowerCase().trim();

  const lastRowInDirectoryWithContent = directorySheet.getLastRow();
  let matchFound = false;
  if (lastRowInDirectoryWithContent >= startRowForDirectorySearchAndWrite) {
    const directoryDataRange = directorySheet.getRange(startRowForDirectorySearchAndWrite, directoryLastNameCol, lastRowInDirectoryWithContent - startRowForDirectorySearchAndWrite + 1, 2);
    const directoryData = directoryDataRange.getValues();
    for (const row of directoryData) {
      const directoryLastName = String(row[0] || "").toLowerCase().trim();
      const directoryFirstName = String(row[1] || "").toLowerCase().trim();
      if (directoryLastName === newMemberLastNameLower && directoryFirstName === newMemberFirstNameLower) {
        matchFound = true;
        break;
      }
    }
  }

  if (matchFound) {
    Logger.log(`Result: A combined match for "${newMemberFirstNameFromSheet} ${newMemberLastNameFromSheet}" already exists. No new entry in Directory.`);
    return;
  }
  Logger.log(`Result: No combined match found. Proceeding to copy data.`);

  // --- Prepare Data for ALL Syncs ---
  const fullNewMemberFormRow = newMemberFormSheet.getRange(lastRowOfSubmission, 1, 1, newMemberFormSheetTimestampCol).getValues()[0];
  const safeGetString = (value) => (value != null ? String(value).trim() : "");
  const safeGetStringAndUpper = (value) => (value != null ? String(value).trim().toUpperCase() : "");
  const formattedDate = Utilities.formatDate(new Date(formResponses[formTimestampIndex]), Session.getScriptTimeZone(), "MM/dd/yyyy");

  // --- SYNC 1: Write to Primary Directory ---
  try {
    const nextRowInDirectory = directorySheet.getLastRow() + 1;
    const directory1NewRow = new Array(22).fill("");
    directory1NewRow[0] = fullNewMemberFormRow[newMemberFormSheetIdCol - 1]; // ID (Col A)
    directory1NewRow[1] = fullNewMemberFormRow[newMemberFormSheetFullNameCol - 1]; // Full Name (Col B)
    directory1NewRow[2] = fullNewMemberFormRow[newMemberFormSheetLastNameCol - 1]; // Last Name (Col C)
    directory1NewRow[3] = fullNewMemberFormRow[newMemberFormSheetFirstNameCol - 1]; // First Name (Col D)
    directory1NewRow[4] = safeGetString(formResponses[2]); // Street (Col E)
    directory1NewRow[5] = safeGetString(formResponses[3]); // City (Col F)
    directory1NewRow[6] = safeGetStringAndUpper(formResponses[4]); // State (Col G)
    directory1NewRow[7] = safeGetString(formResponses[5]); // ZIP (Col H)
    directory1NewRow[8] = safeGetString(formResponses[formLineageIndex]); // Lineage (Col I)
    directory1NewRow[9] = safeGetString(formResponses[6]); // Gender (Col J)
    directory1NewRow[10] = safeGetString(formResponses[7]); // date of birth (Col K)
    directory1NewRow[13] = safeGetString(formResponses[8]); // Phone (Col N)
    directory1NewRow[14] = safeGetString(formResponses[11]); // email (Col O)
    directory1NewRow[21] = formattedDate; // Timestamp (Col V)

    directorySheet.getRange(nextRowInDirectory, 1, 1, directory1NewRow.length).setValues([directory1NewRow]);
    Logger.log(`Success (Sync 1): New member data copied to '${directorySheetName}' sheet, row ${nextRowInDirectory}.`);
  } catch (error) {
    Logger.log(`Error (Sync 1): Error writing data to '${directorySheetName}'. Error: ${error.message}`);
    return;
  }

  // --- SYNC 2: Write to Second Directory (with different layout) ---
  if (directory2Url && directory2TabName) {
    Logger.log(`Info: Proceeding with sync to second directory.`);
    try {
      const secondSpreadsheet = SpreadsheetApp.openByUrl(directory2Url);
      const secondDirectorySheet = secondSpreadsheet.getSheetByName(directory2TabName);

      if (!secondDirectorySheet) {
        Logger.log(`Error (Sync 2): Sheet named '${directory2TabName}' was not found in the second spreadsheet.`);
        return;
      }

      // RELIABLY find the next empty row in the "Last Name" column (Column C)
      const lastNameColumn = secondDirectorySheet.getRange("C:C").getValues();
      let lastRowWithData = 0;
      for (let i = lastNameColumn.length - 1; i >= 0; i--) {
        if (lastNameColumn[i][0]) { // If the cell is not empty
          lastRowWithData = i + 1; // i is 0-based, row is 1-based
          break; // Exit loop once the last data row is found
        }
      }
      // If sheet is empty, start writing at row 2, otherwise write after the last data row
      const nextRowInSecondDirectory = lastRowWithData === 0 ? startRowForDirectorySearchAndWrite : lastRowWithData + 1;
      Logger.log(`(Sync 2) Last row with data in Column C is ${lastRowWithData}. Next available row is ${nextRowInSecondDirectory}.`);


      // --- Using Your Final Provided Mapping ---
      const dir2_TotalColumns = 21; // The highest column used is U (21)
      const directory2NewRow = new Array(dir2_TotalColumns).fill("");

      directory2NewRow[1] = fullNewMemberFormRow[newMemberFormSheetFullNameCol - 1]; // Col B: Full Name
      directory2NewRow[2] = safeGetString(formResponses[0]); // Col C: Last Name (q1)
      directory2NewRow[3] = safeGetString(formResponses[1]); // Col D: First Name (q2)
      directory2NewRow[4] = safeGetString(formResponses[6]); // Col E: Gender (q7)
      directory2NewRow[5] = safeGetString(formResponses[formLineageIndex]); // Lineage (Col F)
      directory2NewRow[6] = safeGetString(formResponses[7]); // Col G: DOB (q8)
      directory2NewRow[9] = "New Member"; // **NEW** Status (Col J)
      directory2NewRow[14] = safeGetString(formResponses[8]); // Phone (Col O)
      directory2NewRow[15] = safeGetString(formResponses[11]); // Col P: Email
      directory2NewRow[16] = safeGetString(formResponses[2]); // Col Q: Street (q3)
      directory2NewRow[17] = safeGetString(formResponses[3]); // Col R: City (q4)
      directory2NewRow[18] = safeGetStringAndUpper(formResponses[4]); // Col S: State (q5)
      directory2NewRow[19] = safeGetString(formResponses[5]); // Col T: ZIP (q6)
      directory2NewRow[20] = formattedDate; // Col U: Timestamp (q11)

      secondDirectorySheet.getRange(nextRowInSecondDirectory, 1, 1, directory2NewRow.length).setValues([directory2NewRow]);
      Logger.log(`✅ Success (Sync 2): New member data copied to '${directory2TabName}' in the second directory, row ${nextRowInSecondDirectory}.`);

    } catch (error) {
      Logger.log(`Error (Sync 2): Could not write to the second directory. Check URL, permissions, and column configuration. Error: ${error.message}`);
    }
  } else {
    Logger.log("Info: 'Directory 2 URL' not set in Config. Skipping second directory sync.");
  }
}

/**
 * This is a test function for processNewMemberForm and is not affected by the config changes,
 * as it operates on the same sheet. It can remain as is.
 */
function testProcessNewMemberForm() {
    // Test function code remains unchanged...
}
