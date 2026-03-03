/**
 * Converts a regular student into a Compulsory student by appending them to the setup table.
 */
function makeCompulsory() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  if (sheet.getName() !== CONFIG.SESSION_SHEET_NAME) return ui.alert("Please use this function on the 'sessions' sheet.");
  
  const activeRow = sheet.getActiveCell().getRow();
  if (activeRow < CONFIG.ROW_DATA_START) return ui.alert("Please select a valid student row.");

  const star = sheet.getRange(activeRow, CONFIG.COL_STAR).getValue();
  if (star === CONFIG.STAR_SYMBOL) return ui.alert("This student is already marked as compulsory.");

  const rawStudent = sheet.getRange(activeRow, CONFIG.COL_RAW_STUDENT).getValue();
  if (!rawStudent || rawStudent.toString().trim() === "") return ui.alert("No student name found on this row.");

  try {
    const mapping = getTableMapping(CONFIG.TABLE_NAME_COMPULSORY);
    const setupSheet = mapping.sheet;
    const today = Utilities.formatDate(new Date(), 'Europe/London', 'dd/MM/yy');

    // Find the first blank row in the table, or add to the very bottom
    const students = setupSheet.getRange(mapping.dataStartRow, mapping.cols[CONFIG.TABLE_COL_STUDENT], mapping.numRows, 1).getValues();
    let targetRowOffset = -1;
    for (let i = 0; i < students.length; i++) {
        if (students[i][0] === "") {
            targetRowOffset = i;
            break;
        }
    }

    let targetRow;
    if (targetRowOffset !== -1) {
        targetRow = mapping.dataStartRow + targetRowOffset; // Fill an empty slot
    } else {
        targetRow = mapping.dataStartRow + mapping.numRows; // Expand the table
        setupSheet.insertRowAfter(targetRow - 1); 
    }

    // Write the new start record
    setupSheet.getRange(targetRow, mapping.cols[CONFIG.TABLE_COL_STUDENT]).setValue(rawStudent);
    setupSheet.getRange(targetRow, mapping.cols[CONFIG.TABLE_COL_START_DATE]).setValue(today);
    setupSheet.getRange(targetRow, mapping.cols[CONFIG.TABLE_COL_END_DATE]).setValue("");

    ui.alert("Success", `${rawStudent} has been added to the Compulsory register. The star will appear momentarily.`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert("Error", error.message, ui.ButtonSet.OK);
  }
}

/**
 * Converts a Compulsory student into a regular student by closing their open record in the setup table.
 */
function makeNonCompulsory() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  if (sheet.getName() !== CONFIG.SESSION_SHEET_NAME) return ui.alert("Please use this function on the 'sessions' sheet.");

  const activeRow = sheet.getActiveCell().getRow();
  if (activeRow < CONFIG.ROW_DATA_START) return ui.alert("Please select a valid student row.");

  const star = sheet.getRange(activeRow, CONFIG.COL_STAR).getValue();
  if (star !== CONFIG.STAR_SYMBOL) return ui.alert("This student is already non-compulsory.");

  const rawStudent = sheet.getRange(activeRow, CONFIG.COL_RAW_STUDENT).getValue();
  if (!rawStudent) return ui.alert("No student name found on this row.");

  try {
    const mapping = getTableMapping(CONFIG.TABLE_NAME_COMPULSORY);
    const setupSheet = mapping.sheet;
    const today = Utilities.formatDate(new Date(), 'Europe/London', 'dd/MM/yy');

    const studentCol = mapping.cols[CONFIG.TABLE_COL_STUDENT];
    const endCol = mapping.cols[CONFIG.TABLE_COL_END_DATE];

    const students = setupSheet.getRange(mapping.dataStartRow, studentCol, mapping.numRows, 1).getValues();
    const endDates = setupSheet.getRange(mapping.dataStartRow, endCol, mapping.numRows, 1).getValues();

    // BOTTOM-UP SEARCH: Find the most recent open record for this specific student
    let recordClosed = false;
    for (let i = students.length - 1; i >= 0; i--) {
        if (students[i][0] === rawStudent && endDates[i][0] === "") {
            setupSheet.getRange(mapping.dataStartRow + i, endCol).setValue(today);
            recordClosed = true;
            break;
        }
    }

    if (recordClosed) {
       ui.alert("Success", `${rawStudent}'s compulsory status has ended. The star will disappear momentarily.`, ui.ButtonSet.OK);
    } else {
       ui.alert("Error", `Could not find an open (blank end-date) record for ${rawStudent} in the setup table.`, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert("Error", error.message, ui.ButtonSet.OK);
  }
}

/**
 * 1. Pulls missing active Compulsory students into the sessions sheet.
 * 2. Sorts the entire sheet (Stars grouped at top, Alphabetical by name).
 */
function syncAndSortRegister() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SESSION_SHEET_NAME);

  if (!sheet) return ui.alert("Sessions sheet not found.");

  try {
    // 1. Get Master List (Compulsory Students without end dates)
    const mapping = getTableMapping(CONFIG.TABLE_NAME_COMPULSORY);
    const students = mapping.sheet.getRange(mapping.dataStartRow, mapping.cols[CONFIG.TABLE_COL_STUDENT], mapping.numRows, 1).getValues();
    const endDates = mapping.sheet.getRange(mapping.dataStartRow, mapping.cols[CONFIG.TABLE_COL_END_DATE], mapping.numRows, 1).getValues();

    const activeCompulsory = new Set();
    for(let i=0; i < students.length; i++) {
        if (students[i][0] !== "" && endDates[i][0] === "") {
            activeCompulsory.add(students[i][0].toString().trim());
        }
    }

    // 2. Get Current List from Sessions sheet
    const lastRow = sheet.getDataRange().getLastRow();
    let currentStudents = [];
    if (lastRow >= CONFIG.ROW_DATA_START) {
        currentStudents = sheet.getRange(CONFIG.ROW_DATA_START, CONFIG.COL_RAW_STUDENT, lastRow - CONFIG.ROW_DATA_START + 1, 1).getValues().map(r => r[0].toString().trim());
    }
    const currentSet = new Set(currentStudents);

    // 3. Find Missing Students
    const missing = [];
    activeCompulsory.forEach(student => {
        if (!currentSet.has(student)) missing.push([student]);
    });

    // 4. Append Missing Students to the bottom of Col A
    if (missing.length > 0) {
        const appendRow = lastRow < CONFIG.ROW_DATA_START ? CONFIG.ROW_DATA_START : lastRow + 1;
        sheet.getRange(appendRow, CONFIG.COL_RAW_STUDENT, missing.length, 1).setValues(missing);
    }

    // 5. Multi-Level Sort
    // Force the BYROW formula to calculate the new stars BEFORE we sort them
    SpreadsheetApp.flush(); 

    const newLastRow = sheet.getDataRange().getLastRow();
    const newLastCol = sheet.getDataRange().getLastColumn();

    if (newLastRow >= CONFIG.ROW_DATA_START) {
        const sortRange = sheet.getRange(CONFIG.ROW_DATA_START, 1, newLastRow - CONFIG.ROW_DATA_START + 1, newLastCol);
        sortRange.sort([
            {column: CONFIG.COL_STAR, ascending: false}, // Sort 1: ★ Descending (Z to A groups stars at top)
            {column: CONFIG.COL_RAW_STUDENT, ascending: true} // Sort 2: Raw Student Alphabetical
        ]);
    }

    ui.alert("Success", `Register synced and sorted!\nAdded ${missing.length} missing compulsory students to the sheet.`, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert("Error during Sync & Sort", error.message, ui.ButtonSet.OK);
  }
}