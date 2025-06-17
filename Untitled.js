/********* CONFIGURATION *********/
const PROFILE_SHEET      = 'Profile';      // sheet that holds the community code
const COMMUNITY_ID_CELL  = 'E4';           // where that code lives
const DIRECTORY_SHEET    = 'Directory';    // sheet with the people
const HEADER_ROW_INDEX   = 2;              // 1-based row number of the header row
const ID_HEADER_TEXT     = 'Person ID';    // exact text of the column header
const ID_PAD_LENGTH      = 5;              // digits after the dash (00001)

/**
 * Add a custom menu every time the file opens
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Directory Tools')
    .addItem('Assign Missing Person IDs', 'assignPersonIds')
    .addToUi();
}

/**
 * Main routine – fills blank Person-ID cells
 */
function assignPersonIds() {
  const ss          = SpreadsheetApp.getActive();
  const profile     = ss.getSheetByName(PROFILE_SHEET);
  const directory   = ss.getSheetByName(DIRECTORY_SHEET);

  if (!profile || !directory) {
    SpreadsheetApp.getUi().alert('Could not find the Profile or Directory sheet.');
    return;
  }

  // 1. Fetch community code (e.g. BEL)
  const communityCode = String(profile.getRange(COMMUNITY_ID_CELL).getValue()).trim().toUpperCase();
  if (!communityCode) {
    SpreadsheetApp.getUi().alert(`No community code found in ${PROFILE_SHEET}!${COMMUNITY_ID_CELL}`);
    return;
  }

  // 2. Locate Person-ID column from the header row
  const headerRow = directory.getRange(HEADER_ROW_INDEX, 1, 1, directory.getLastColumn()).getValues()[0];
  const idColIdx  = headerRow.indexOf(ID_HEADER_TEXT);
  if (idColIdx === -1) {
    SpreadsheetApp.getUi().alert(`Header "${ID_HEADER_TEXT}" not found in the Directory sheet.`);
    return;
  }

  // 3. Read all data once
  const firstDataRow = HEADER_ROW_INDEX + 1;
  const numRows      = directory.getLastRow() - HEADER_ROW_INDEX;
  if (numRows <= 0) { SpreadsheetApp.getUi().alert('No data rows to process.'); return; }

  const data = directory.getRange(firstDataRow, 1, numRows, directory.getLastColumn()).getValues();

  // 4. Determine the current highest sequence for this community
  let maxSeq = 0;
  data.forEach(r => {
    const id = r[idColIdx];
    if (id && typeof id === 'string' && id.startsWith(communityCode + '-')) {
      const num = parseInt(id.split('-')[1], 10);
      if (!isNaN(num)) maxSeq = Math.max(maxSeq, num);
    }
  });

  // 5. Assign new IDs where blank
  let updates = 0;
  data.forEach((r, i) => {
    if (!r[idColIdx]) {
      maxSeq++;
      r[idColIdx] = `${communityCode}-${String(maxSeq).padStart(ID_PAD_LENGTH, '0')}`;
      updates++;
    }
  });

  // 6. Push the updated ID column back to the sheet (one batch write)
  if (updates > 0) {
    const idRange = directory.getRange(firstDataRow, idColIdx + 1, numRows, 1);
    idRange.setValues(data.map(r => [r[idColIdx]]));
    SpreadsheetApp.getUi().alert(`Assigned ${updates} new Person IDs (latest = ${communityCode}-${String(maxSeq).padStart(ID_PAD_LENGTH, '0')}).`);
  } else {
    SpreadsheetApp.getUi().alert('All rows already have a Person ID – nothing to do.');
  }
}
