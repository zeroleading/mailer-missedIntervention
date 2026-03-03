/**
 * Orchestrates the creation of the new sessions sheet from the Compulsory table.
 */
function setupSessions() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (ss.getSheetByName(CONFIG.SESSION_SHEET_NAME)) {
    ui.alert("Sheet Exists", `A sheet named '${CONFIG.SESSION_SHEET_NAME}' already exists.\n\nPlease rename the current sessions sheet (e.g., to 'sessions_term1') and retry.`, ui.ButtonSet.OK);
    return;
  }
  
  const templateSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET_NAME);
  if (!templateSheet) {
    ui.alert("Error", `Could not find the template sheet named '${CONFIG.TEMPLATE_SHEET_NAME}'.`, ui.ButtonSet.OK);
    return;
  }

  try {
    const mapping = getTableMapping(CONFIG.TABLE_NAME_COMPULSORY);
    const studentCol = mapping.cols[CONFIG.TABLE_COL_STUDENT];
    const endCol = mapping.cols[CONFIG.TABLE_COL_END_DATE];
    
    // Read the students and end dates from the table
    let studentsRaw = [];
    let endDatesRaw = [];
    if (mapping.numRows > 0) {
      studentsRaw = mapping.sheet.getRange(mapping.dataStartRow, studentCol, mapping.numRows, 1).getValues();
      endDatesRaw = mapping.sheet.getRange(mapping.dataStartRow, endCol, mapping.numRows, 1).getValues();
    }
    
    // Filter out blank rows and students who have an end date
    // Using a Set naturally prevents duplicates if a student was added twice without an end date
    const activeCompulsory = new Set();
    for (let i = 0; i < studentsRaw.length; i++) {
      const studentName = studentsRaw[i][0];
      const endDate = endDatesRaw[i][0];
      if (studentName !== "" && endDate === "") {
        activeCompulsory.add(studentName.toString().trim());
      }
    }
    
    let studentList = Array.from(activeCompulsory);
    
    const newSheet = templateSheet.copyTo(ss);
    newSheet.setName(CONFIG.SESSION_SHEET_NAME);
    newSheet.showSheet(); 
    ss.setActiveSheet(newSheet);
    
    if (studentList.length > 0) {
      studentList.sort((a, b) => a.localeCompare(b));
      const pasteData = studentList.map(student => [student]);
      newSheet.getRange(CONFIG.ROW_DATA_START, CONFIG.COL_RAW_STUDENT, pasteData.length, 1).setValues(pasteData);
      ui.alert("Success", `Setup complete. Added and sorted ${studentList.length} active compulsory students to the new sessions sheet.`, ui.ButtonSet.OK);
    } else {
      ui.alert("Success", `Setup complete. A blank sessions sheet was created (no active compulsory students found).`, ui.ButtonSet.OK);
    }

    // Stamp initial start dates for any students in the table who don't have one
    stampInitialStartDates(mapping);
    
  } catch (error) {
    ui.alert("Setup Failed", error.message, ui.ButtonSet.OK);
  }
}

/**
 * Stamps today's date into the startDate column for any table row that is missing it.
 */
function stampInitialStartDates(mapping) {
  if (mapping.numRows <= 0) return;
  const studentCol = mapping.cols[CONFIG.TABLE_COL_STUDENT];
  const startCol = mapping.cols[CONFIG.TABLE_COL_START_DATE];
  if (!studentCol || !startCol) return;

  const students = mapping.sheet.getRange(mapping.dataStartRow, studentCol, mapping.numRows, 1).getValues();
  const startDates = mapping.sheet.getRange(mapping.dataStartRow, startCol, mapping.numRows, 1).getValues();
  const today = Utilities.formatDate(new Date(), 'Europe/London', 'dd/MM/yy');
  let changed = false;

  for (let i = 0; i < students.length; i++) {
    if (students[i][0] !== "" && startDates[i][0] === "") {
      startDates[i][0] = today;
      changed = true;
    }
  }

  if (changed) {
    mapping.sheet.getRange(mapping.dataStartRow, startCol, mapping.numRows, 1).setValues(startDates);
  }
}

/**
 * Super-Helper: Returns exact grid coordinates and column mappings for a named Table.
 */
function getTableMapping(tableName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssData = Sheets.Spreadsheets.get(ss.getId());

  const sheetData = ssData.sheets.find(s => s.tables?.some(t => t.name === tableName));
  if (!sheetData) throw new Error(`Sheet containing table '${tableName}' not found.`);
  const table = sheetData.tables.find(t => t.name === tableName);
  if (!table) throw new Error(`Table '${tableName}' not found.`);

  const sheet = ss.getSheetByName(sheetData.properties.title);
  const startRow = table.range.startRowIndex + 1; // 1-based Header row
  const dataStartRow = startRow + 1;              // First row of data
  const numRows = table.range.endRowIndex - table.range.startRowIndex - 1; 
  const startCol = table.range.startColumnIndex + 1;

  const colMapping = {};
  table.columnProperties.forEach((col, idx) => {
     colMapping[col.columnName] = startCol + idx;
  });

  return {
     sheet: sheet,
     dataStartRow: dataStartRow,
     numRows: numRows,
     cols: colMapping
  };
}