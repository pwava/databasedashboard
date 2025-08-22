/**
 * Counts unique attendees for each Sunday date in 'Attend Wksht'
 * based on data from the 'Sunday Service Attend' tab and an external sheet.
 */
function updateAttendanceCounts() {
  Logger.log('Script execution started.');
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetTimeZone = spreadsheet.getSpreadsheetTimeZone();

  const attendSheet = spreadsheet.getSheetByName(getSetting('Attendance Worksheet Tab'));
  const attendServiceSheet = spreadsheet.getSheetByName(getSetting('Sunday Service Attend Tab'));

  if (!attendSheet || !attendServiceSheet) {
    const missingSheet = !attendSheet ? "Attendance Worksheet Tab" : "Sunday Service Attend Tab";
    Logger.log(`Error: Sheet "${missingSheet}" not found based on settings.`);
    SpreadsheetApp.getUi().alert(`Error: Sheet "${missingSheet}" not found based on settings.`);
    return;
  }

  const attendanceDataRange = attendServiceSheet.getDataRange();
  const attendanceValues = attendanceDataRange.getValues();
  const attendeesByDate = new Map();

  const dateColumnIndex_main = 4; // Column E
  const personalIdColumnIndex_main = 1; // Column B

  for (let i = 1; i < attendanceValues.length; i++) {
    const row = attendanceValues[i];
    const dateValue = row[dateColumnIndex_main];
    const personalId = row[personalIdColumnIndex_main];

    if (dateValue && personalId && personalId.toString().trim() !== '') {
      try {
        const serviceDate = new Date(dateValue);
        if (isNaN(serviceDate.getTime())) continue;

        const dateString = Utilities.formatDate(serviceDate, spreadsheetTimeZone, 'yyyy-MM-dd');

        if (!attendeesByDate.has(dateString)) {
          attendeesByDate.set(dateString, new Set());
        }
        
        // --- NAME CLEANING REVERTED TO EXACTLY MATCH THE ORIGINAL SCRIPT ---
        attendeesByDate.get(dateString).add(personalId.toString().trim());

      } catch (e) {
        Logger.log(`Could not process row ${i + 1}. Date: ${dateValue}, Name: ${personalId}`);
      }
    }
  }
  Logger.log('Finished processing main attendance data.');

  // 4. Process data from the external spreadsheet
  const externalDataByDate_D = new Map();
  const externalDataByDate_E = new Map();
  const directory2Url = getSetting('Directory 2 URL');
  
  if (directory2Url) {
    try {
      const externalSpreadsheet = SpreadsheetApp.openByUrl(directory2Url);
      const externalSheet = externalSpreadsheet.getSheetByName(getSetting('Attendance Worksheet Tab'));
      
      if (externalSheet) {
        const externalDataRange = externalSheet.getDataRange();
        const externalValues = externalDataRange.getValues();
        
        const externalDateColumnIndex = 1; // Column B
        const externalDataColumnIndex_D = 3; // Column D
        const externalDataColumnIndex_E = 4; // Column E
        
        for (let i = 1; i < externalValues.length; i++) {
          const row = externalValues[i];
          const date = row[externalDateColumnIndex];
          const dataFromD = row[externalDataColumnIndex_D];
          const dataFromE = row[externalDataColumnIndex_E];
          
          if (date && date instanceof Date) {
            const dateString = Utilities.formatDate(date, spreadsheetTimeZone, 'yyyy-MM-dd');
            externalDataByDate_D.set(dateString, dataFromD);
            externalDataByDate_E.set(dateString, dataFromE);
          }
        }
      }
    } catch (e) {
      Logger.log('Error opening or accessing the external spreadsheet: ' + e.message);
    }
  }

  // 5. Read the Sunday dates from the 'Attend Wksht' tab of the main spreadsheet
  const startRow = 4;
  const dateColumnAttendIndex = 1; // Column B
  const countColumnAttendIndex = 2; // Column C
  const externalDataColumnIndex_D_main = 3; // Column D
  const externalDataColumnIndex_E_main = 4; // Column E

  if (attendSheet.getLastRow() < startRow) {
      Logger.log('No dates found in "Attend Wksht" to process.');
      return;
  }
  
  const attendDatesRange = attendSheet.getRange(startRow, dateColumnAttendIndex + 1, attendSheet.getLastRow() - startRow + 1, 1);
  const attendDatesValues = attendDatesRange.getValues();

  const uniqueCounts = [];
  const calculatedExternalData_D = [];
  const calculatedExternalData_E = [];

  for (let i = 0; i < attendDatesValues.length; i++) {
    const dateValue = attendDatesValues[i][0];
    let uniqueCount = 0;
    let externalData_D = '';
    let externalData_E = '';

    if (dateValue && dateValue instanceof Date) {
      const dateString = Utilities.formatDate(dateValue, spreadsheetTimeZone, 'yyyy-MM-dd');
      
      if (attendeesByDate.has(dateString)) {
        uniqueCount = attendeesByDate.get(dateString).size;
      }
      
      if (externalDataByDate_D.has(dateString)) {
        externalData_D = externalDataByDate_D.get(dateString);
      }
      if (externalDataByDate_E.has(dateString)) {
        externalData_E = externalDataByDate_E.get(dateString);
      }
    }
    
    uniqueCounts.push([uniqueCount]);
    calculatedExternalData_D.push([externalData_D]);
    calculatedExternalData_E.push([externalData_E]);
  }

  // 6. Write the calculated data back to the main sheet
  if (uniqueCounts.length > 0) {
    attendSheet.getRange(startRow, countColumnAttendIndex + 1, uniqueCounts.length, 1).setValues(uniqueCounts);
    attendSheet.getRange(startRow, externalDataColumnIndex_D_main + 1, calculatedExternalData_D.length, 1).setValues(calculatedExternalData_D);
    attendSheet.getRange(startRow, externalDataColumnIndex_E_main + 1, calculatedExternalData_E.length, 1).setValues(calculatedExternalData_E);
    
    if (directory2Url) {
        try {
            const externalSpreadsheet = SpreadsheetApp.openByUrl(directory2Url);
            const externalSheet = externalSpreadsheet.getSheetByName(getSetting('Attendance Worksheet Tab'));
            if (externalSheet) {
                externalSheet.getRange(startRow, countColumnAttendIndex + 1, uniqueCounts.length, 1).setValues(uniqueCounts);
            }
        } catch(e) {
            Logger.log('Error writing to the external spreadsheet: ' + e.message);
        }
    }
  }
  Logger.log('Script execution completed.');
}

function getSetting(settingName) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config for Urls');
  if (!configSheet) {
    Logger.log('Error: Config sheet "Config for Urls" not found.');
    SpreadsheetApp.getUi().alert('Error: Config sheet "Config for Urls" not found.');
    return null;
  }
  
  const data = configSheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === settingName) {
      return data[i][1];
    }
  }
  return null;
}
