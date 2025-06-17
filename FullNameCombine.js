/**
 * Combines First and Last Names from a form submission and writes the Full Name
 * to the 'Full Name' column (Column B) on the 'New Member Form' sheet.
 *
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e The event object from the form submission.
 */
function combineNamesOnFormSubmit(e) {
  Logger.log("combineNamesOnFormSubmit function started.");

  // Configuration for the 'New Member Form' sheet
  const newMemberFormSheetName = "New Member Form";
  const fullNameColumn = 2; // Column B
  const lastNameIndexInFormResponse = 0; // e.values[0] for Last Name from form
  const firstNameIndexInFormResponse = 1; // e.values[1] for First Name from form

  if (!e || !e.values || !e.range) {
    Logger.log("Error: combineNamesOnFormSubmit was called without complete form data (missing e.values or e.range).");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(newMemberFormSheetName);

  if (!sheet) {
    Logger.log(`Error: Sheet '${newMemberFormSheetName}' not found for combineNamesOnFormSubmit.`);
    // FIX: Using correct Ui.alert signature
    SpreadsheetApp.getUi().alert("Error", `Sheet '${newMemberFormSheetName}' not found. Cannot combine names.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const lastRow = e.range.getRow(); // Get the row where the form data was submitted

  // Get the Last Name and First Name directly from the event object's values
  const lastName = e.values[lastNameIndexInFormResponse];
  const firstName = e.values[firstNameIndexInFormResponse];

  // Basic validation for names
  if (!firstName || String(firstName).trim() === "" || !lastName || String(lastName).trim() === "") {
    Logger.log(`Warning: First Name ("${firstName}") or Last Name ("${lastName}") is empty. Skipping full name combination for row ${lastRow}.`);
    return;
  }

  // Combine first name and last name.
  // Ensure trimming to avoid extra spaces if names have leading/trailing whitespace
  const fullName = `${String(firstName).trim()} ${String(lastName).trim()}`;

  try {
    // Set the combined full name into Column B of the last row.
    sheet.getRange(lastRow, fullNameColumn).setValue(fullName);
    Logger.log(`Success: Combined name "${fullName}" written to Column B, Row ${lastRow} on '${newMemberFormSheetName}'.`);
  } catch (error) {
    Logger.log(`Error: Failed to write combined name to sheet for row ${lastRow}. Error: ${error.message}`);
  }
}