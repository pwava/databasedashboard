/**
 * @OnlyCurrentDoc
 * This script manages the central configuration for the multi-sheet church system.
 */

/// This function runs automatically when the spreadsheet is opened.
// It creates the custom menu for the setup process.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ System Setup')
    .addItem('🎨 1. Create Config Sheet', 'createConfigSheet')
    .addSeparator()
    .addItem('🚀 2. Initialize System', 'initializeSystem')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔄 Data Sync')
        .addItem('Update Activity Levels', 'updateActivityLevels')
        .addItem('Update Giving Levels', 'updateGivingLevelsFromDonorStats'))
    .addToUi();
}

/**
 * Creates and formats the "Config for Urls" sheet with all necessary headers and fields.
 * -- THIS IS THE CORRECTED VERSION --
 */
function createConfigSheet() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Config for Urls';

  if (spreadsheet.getSheetByName(sheetName)) {
    ui.alert(`The "${sheetName}" sheet already exists. No changes were made.`);
    return;
  }
  
  const sheet = spreadsheet.insertSheet(sheetName, 0);
  spreadsheet.setActiveSheet(sheet);

  ui.alert(`A "${sheetName}" sheet has been created. Now formatting the sheet...`);

  // --- Formatting ---
  sheet.getRange('A:C').setWrap(true);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 300);
  
  const headers = [['Setting Name', 'Value (Paste URL or Name Here)', 'Description']];
  const headerRange = sheet.getRange('A1:C1');
  headerRange.setValues(headers);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9d2e9');
  headerRange.setFontColor('#000000');
  
  // --- Data Population ---
  const configData = [
    ['SPREADSHEET URLs', '', 'Paste the full URL from your browser\'s address bar.'],
    ['Dashboard URL', '(This field auto-fills during initialization)', 'The URL of this main dashboard spreadsheet.'],
    ['Attendance Tracker URL', '', 'Paste the URL of the "Attendance Tracker" sheet.'],
    ['Event Management URL', '', 'Paste the URL of the "Event Management" sheet.'],
    ['Donation Data URL', '', 'Paste the URL of the "Donation Data" sheet.'],
    ['Central Response URL', '', 'Paste the URL of the "Central Response" sheet.'],
    ['Tools URL', '', 'Paste the URL of the "Tools" sheet.'],
    ['', '', ''],
    ['CRITICAL TAB NAMES', '', 'Only change these if you rename the tabs in your sheets.'],
    ['Master Directory Tab', 'Directory', 'The name of the tab containing the main member list.'],
    ['Activity Level Column', 'Activity Level', 'The exact name of the column in the Directory to be updated.'],
    ['Event Check-in Tab', 'Check-in Management', 'The tab in Event Mgt that lists newly created forms.'],
    ['Event Attendance Tab', 'Event Attendance', 'The tab in Central Response that logs event attendees.'],
    ['', '', ''],
    ['PERSON ID CHECK LOCATIONS', '', 'List of all tabs to scan for duplicate names & max ID.'],
    ['Dashboard', 'Directory', ''],
    ['Attendance Tracker', 'Service Attendance', ''],
    ['Attendance Tracker', 'Event Attendance', ''],
    ['Central Response', 'Event Attendance', '']
  ];

  sheet.getRange(2, 1, configData.length, 3).setValues(configData);
  
  // -- FIXED LINES START HERE --
  // Use getRangeList for multiple, non-contiguous ranges.
  sheet.getRangeList(['A2:A20', 'C2:C20']).setBackground('#f3f3f3');
  sheet.getRangeList(['A2', 'A10', 'A16']).setFontWeight('bold');
  // -- FIXED LINES END HERE --
  
  ui.alert(`The "${sheetName}" sheet has been created and formatted. Please fill in the required URLs in Column B, then run "System Setup > 2. Initialize System".`);
}

/**
 * Reads settings from the "Config for Urls" sheet and saves them as script properties.
 */
function initializeSystem() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Config for Urls'; // Defined the new name here
  const configSheet = spreadsheet.getSheetByName(sheetName);

  if (!configSheet) {
    ui.alert('Error', `A sheet named "${sheetName}" was not found. Please run "1. Create Config Sheet" first.`, ui.ButtonSet.OK);
    return;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  
  const dashboardUrl = spreadsheet.getUrl();
  configSheet.getRange('B3').setValue(dashboardUrl);
  scriptProperties.setProperty('Dashboard URL', dashboardUrl);
  
  const settingsRange = configSheet.getRange('A4:B' + configSheet.getLastRow());
  const settings = settingsRange.getValues();

  let settingsCount = 0;
  for (let i = 0; i < settings.length; i++) {
    const key = settings[i][0];
    const value = settings[i][1];
    if (key && value) {
      scriptProperties.setProperty(key, value);
      settingsCount++;
    }
  }
  ui.alert('Setup Complete!', `Successfully saved ${settingsCount + 1} settings. The system is now configured.`, ui.ButtonSet.OK);
}

/**
 * A helper function to easily retrieve a saved setting.
 */
function getSetting(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}