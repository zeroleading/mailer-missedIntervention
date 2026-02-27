/**
 * Opens the Preview Modal dialog in the spreadsheet.
 */
function openPreviewModal() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Basic validation: Ensure they are on a tracking sheet and on a valid data row
  const activeRow = sheet.getActiveCell().getRow();
  
  if (sheet.getName() === CONFIG.TABLE_SETUP_SHEET || sheet.getName() === CONFIG.MESSAGE_SHEET_NAME) {
    ui.alert("Navigation Error", "Please navigate to your 'sessions' sheet and select a student row to preview.", ui.ButtonSet.OK);
    return;
  }
  
  if (activeRow < CONFIG.ROW_DATA_START) {
    ui.alert("Selection Error", `Please select a row containing student data (Row ${CONFIG.ROW_DATA_START} or below).`, ui.ButtonSet.OK);
    return;
  }

  // Open the HTML modal
  const htmlOutput = HtmlService.createHtmlOutputFromFile('PreviewUI')
      .setWidth(650)
      .setHeight(600)
      .setTitle('Email Preview');
  ui.showModalDialog(htmlOutput, 'Email Preview Tool');
}

/**
 * Called by the HTML Modal to fetch the compiled HTML strings for the active row.
 * @returns {Object} { absentHtml, attendedHtml, studentName, subjectLine }
 */
function getPreviewData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRow = sheet.getActiveCell().getRow();
  
  // 1. Get Student Name from active row
  const studentName = sheet.getRange(activeRow, CONFIG.COL_STUDENT_NAME).getValue() || "[Student Name]";
  
  // 2. Try to get the most recent session date for a realistic subject line
  const lastColumn = sheet.getRange(1, 1).getDataRegion().getLastColumn();
  let sessionDate = "[Session Date]";
  if (lastColumn > CONFIG.COL_STAR) {
    const sessionRecent = sheet.getRange(1, lastColumn - 1).getValue();
    const wheresTheAt = sessionRecent.indexOf('@');
    if (wheresTheAt !== -1) {
      sessionDate = sessionRecent.substring(wheresTheAt + 12).trim();
    }
  }

  // 3. Fetch Base Templates
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const msgSheet = ss.getSheetByName(CONFIG.MESSAGE_SHEET_NAME);
  const absentMsgRaw = msgSheet.getRange('D2').getValue();
  const attendedMsgRaw = msgSheet.getRange('D3').getValue();
  const closingText = msgSheet.getRange('D4').getValue();
  const signature = msgSheet.getRange('D5').getValue();

  // Helper to compile a specific message type
  const compilePreview = (baseMsg) => {
    if (!baseMsg || baseMsg.toString().trim() === "") {
      return `<div style="padding:20px; font-family:sans-serif; color:#666; text-align:center;">
                <i>This template is currently blank. <br><br>Emails of this type will be safely skipped during processing.</i>
              </div>`;
    }
    
    let body = `${baseMsg}${closingText}${signature}`;
    body = body.replace(/\{\{name\}\}/g, studentName);
    body = body.replace(/\n/g, '<br>');
    return buildHtmlEmail(body); // Reuses our existing function from EmailService.gs!
  };

  // 4. Generate Subject Line
  const subjectLine = CONFIG.SUBJECT_TEMPLATE
    .replace(/\{\{studentName\}\}/g, studentName)
    .replace(/\{\{subjectName\}\}/g, CONFIG.SUBJECT_NAME)
    .replace(/\{\{sessionType\}\}/g, CONFIG.SESSION_TYPE)
    .replace(/\{\{sessionDate\}\}/g, sessionDate);

  // 5. Return data to the frontend
  return {
    absentHtml: compilePreview(absentMsgRaw),
    attendedHtml: compilePreview(attendedMsgRaw),
    studentName: studentName,
    subjectLine: subjectLine
  };
}