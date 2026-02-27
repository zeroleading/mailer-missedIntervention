/**
 * Orchestrates the creation of the new sessions sheet from the Compulsory table.
 */
function setupSessions() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Safety Check: Does 'sessions' already exist?
  if (ss.getSheetByName(CONFIG.SESSION_SHEET_NAME)) {
    ui.alert("Sheet Exists", `A sheet named '${CONFIG.SESSION_SHEET_NAME}' already exists.\n\nPlease rename the current sessions sheet (e.g., to 'sessions_term1') and retry.`, ui.ButtonSet.OK);
    return;
  }
  
  // 2. Fetch the cleanSheet template
  const templateSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET_NAME);
  if (!templateSheet) {
    ui.alert("Error", `Could not find the template sheet named '${CONFIG.TEMPLATE_SHEET_NAME}'.`, ui.ButtonSet.OK);
    return;
  }

  try {
    // 3. Get students from the table (filtering blanks)
    const studentList = getTableColumn(CONFIG.TABLE_NAME_COMPULSORY, CONFIG.TABLE_COL_STUDENT, true);
    
    // 4. Duplicate the template sheet regardless of whether there are students
    const newSheet = templateSheet.copyTo(ss);
    newSheet.setName(CONFIG.SESSION_SHEET_NAME);
    newSheet.showSheet(); // Ensure it is visible
    ss.setActiveSheet(newSheet);
    
    // 5. Inject the data ONLY if compulsory students exist
    if (studentList.length > 0) {
      // Convert 1D array to 2D array for pasting: ["A", "B"] -> [["A"], ["B"]]
      const pasteData = studentList.map(student => [student]);
      newSheet.getRange(CONFIG.ROW_DATA_START, CONFIG.COL_RAW_STUDENT, pasteData.length, 1).setValues(pasteData);
      
      ui.alert("Success", `Setup complete. Added ${studentList.length} compulsory students to the new sessions sheet.`, ui.ButtonSet.OK);
    } else {
      ui.alert("Success", `Setup complete. A blank sessions sheet was created (no compulsory students found).`, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert("Setup Failed", error.message, ui.ButtonSet.OK);
  }
}

/**
 * Retrieves values from a specific column in a Table via the Advanced Sheets API.
 */
function getTableColumn(tableName, colName, excludeBlanks = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssData = Sheets.Spreadsheets.get(ss.getId());

  const sheet = ssData.sheets.find(s => s.tables?.some(t => t.name === tableName));
  const table = sheet?.tables.find(t => t.name === tableName);
  if (!table) throw new Error(`Table '${tableName}' not found.`);

  const col = table.columnProperties?.find(c => c.columnName === colName);
  if (!col) throw new Error(`Column '${colName}' not found.`);

  const startRow = table.range.startRowIndex + 2; 
  const numRows = table.range.endRowIndex - table.range.startRowIndex - 1; 

  if (numRows <= 0) return [];

  let data = ss.getSheetByName(sheet.properties.title)
    .getRange(startRow, col.columnIndex + 1, numRows, 1)
    .getValues()
    .flat(); 

  if (excludeBlanks) {
    data = data.filter(val => val !== ""); 
  }

  return data;
}