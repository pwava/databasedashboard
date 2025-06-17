/**
 * Counts unique attendees for each Sunday date in 'Attend Wksht'
 * based on data from the 'Sunday Service Attend' tab.
 */
function updateAttendanceCounts() {
  // 1. Get the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 2. Get the sheets by name using settings from the Config sheet
  // -- ⭐️ CHANGED LINES START HERE ⭐️ --
  const attendSheet = spreadsheet.getSheetByName(getSetting('Attendance Worksheet Tab'));     // Was 'Attend Wksht'
  const attendServiceSheet = spreadsheet.getSheetByName(getSetting('Sunday Service Attend Tab')); // Was 'Sunday Service Attend'
  // -- ⭐️ CHANGED LINES END HERE ⭐️ --

  // Check if the sheets exist
  if (!attendSheet) {
    Logger.log('Error: Sheet for Attendance Worksheet not found based on settings.');
    SpreadsheetApp.getUi().alert('Error: Sheet for Attendance Worksheet not found based on settings.');
    return; // Stop execution if sheet is not found
  }
  if (!attendServiceSheet) {
    Logger.log('Error: Sheet for Sunday Service Attend not found based on settings.');
    SpreadsheetApp.getUi().alert('Error: Sheet for Sunday Service Attend not found based on settings.');
    return; // Stop execution if sheet is not found
  }

  // 3. Read all data from the 'Sunday Service Attend' tab
  const attendanceDataRange = attendServiceSheet.getDataRange();
  const attendanceValues = attendanceDataRange.getValues();

  // Map to store unique attendees per date.
  const attendeesByDate = new Map();

  // Process attendance data (skip header row)
  const dateColumnIndex = 4; // Column E
  const personalIdColumnIndex = 0; // Column A

  for (let i = 1; i < attendanceValues.length; i++) {
    const row = attendanceValues[i];
    const date = row[dateColumnIndex];
    const personalId = row[personalIdColumnIndex];

    if (date && personalId && date instanceof Date) {
      const dateString = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      if (!attendeesByDate.has(dateString)) {
        attendeesByDate.set(dateString, new Set());
      }
      attendeesByDate.get(dateString).add(personalId);
    }
  }

  // 4. Read the Sunday dates from the 'Attend Wksht' tab
  const startRow = 4;
  const dateColumnAttendIndex = 1; // Column B
  const countColumnAttendIndex = 2; // Column C

  // Make sure the sheet has enough rows before getting the range
  if (attendSheet.getLastRow() < startRow) {
      Logger.log('No dates found in "Attend Wksht" to process.');
      return;
  }
  
  const attendDatesRange = attendSheet.getRange(startRow, dateColumnAttendIndex + 1, attendSheet.getLastRow() - startRow + 1, 1);
  const attendDatesValues = attendDatesRange.getValues();

  const calculatedCounts = [];

  for (let i = 0; i < attendDatesValues.length; i++) {
    const dateValue = attendDatesValues[i][0];
    let uniqueCount = 0;

    if (dateValue && dateValue instanceof Date) {
      const dateString = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (attendeesByDate.has(dateString)) {
        uniqueCount = attendeesByDate.get(dateString).size;
      }
    }
    calculatedCounts.push([uniqueCount]);
  }

  // 5. Write the calculated counts back to column C in 'Attend Wksht'
  if (calculatedCounts.length > 0) {
    const targetRange = attendSheet.getRange(startRow, countColumnAttendIndex + 1, calculatedCounts.length, 1);
    targetRange.setValues(calculatedCounts);
  } else {
    Logger.log('No dates with attendance found to update in "Attend Wksht".');
  }

  Logger.log('Attendance counts updated successfully.');
}