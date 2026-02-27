/**
 * Adds two new columns to the active sheet for tracking a new session.
 */
function addSession() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Guard clause to ensure they are on a tracking sheet
  if (sheet.getName() === CONFIG.TABLE_SETUP_SHEET || sheet.getName() === CONFIG.MESSAGE_SHEET_NAME) {
    ui.alert("Please navigate to your 'sessions' sheet before adding a session.");
    return;
  }

  const lastColumn = sheet.getRange(1, 1).getDataRegion().getLastColumn();
  const lastRow = Math.max(sheet.getRange(1, 1).getDataRegion().getLastRow(), CONFIG.ROW_DATA_START);
  const user = Session.getActiveUser().getEmail();
  const formattedDate = Utilities.formatDate(new Date(), 'Europe/London', 'dd/MM/yy');

  const sessionDatePrompt = ui.prompt(
    'Revision session date',
    `Enter date below, today's date is ${formattedDate}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (sessionDatePrompt.getSelectedButton() == ui.Button.OK) {
    const sessionDate = sessionDatePrompt.getResponseText() || formattedDate;

    // Insert columns and format
    sheet.insertColumnsAfter(lastColumn, 2);
    sheet.setColumnWidths(lastColumn + 1, 2, 160);
    
    const sessionHeader = `Logged by: ${user}\n${sessionDate}`;
    sheet.getRange(1, lastColumn + 1).setValue(sessionHeader);
    sheet.getRange(1, lastColumn + 2).setValue('Outcome');

    // Random pastel colors for header
    const r = Math.floor(Math.random() * 50) + 205;
    const g = Math.floor(Math.random() * 50) + 205;
    const b = Math.floor(Math.random() * 50) + 205;

    const headerRange = sheet.getRange(1, lastColumn + 1, 1, 2);
    headerRange.setBackgroundRGB(r, g, b)
               .setFontWeight('bold')
               .setWrap(true);

    // Apply font and Data Validation starting from the data row
    sheet.getRange(1, lastColumn + 1, lastRow, 2).setFontFamily('Proxima Nova');

    const dvList = ['Attended', 'Absent', 'Unable to attend'];
    const dvRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dvList, true)
      .setAllowInvalid(false)
      .build();
    
    // Apply validation rules to the new data column
    const numRowsToValidate = lastRow - CONFIG.ROW_DATA_START + 1;
    if (numRowsToValidate > 0) {
      const dataRange = sheet.getRange(CONFIG.ROW_DATA_START, lastColumn + 1, numRowsToValidate, 1);
      dataRange.setDataValidation(dvRule).setValue('');
    }
  } else {
    ui.alert('Session not added, try again');
  }
}

/**
 * Identifies the most recently added session column.
 * @returns {Array|null} [sessionDateText, outcomeColumnIndex]
 */
function sessionDetails() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastColumn = sheet.getRange(1, 1).getDataRegion().getLastColumn();

  let sessionRecent = sheet.getRange(1, lastColumn - 1).getValue();
  const wheresTheAt = sessionRecent.indexOf('@'); // Looks for email in "Logged by:"

  if (wheresTheAt == -1) {
    ui.alert('Session has not been set correctly, please set at least one session before continuing');
    return null;
  } else {
    sessionRecent = sessionRecent.substring(wheresTheAt + 12).trim(); // Extracts the date

    const sessionIdentified = ui.alert(
      `Confirm Batch Process`,
      `You are about to process emails for the session:\n\n${sessionRecent}`,
      ui.ButtonSet.OK_CANCEL
    );

    if (sessionIdentified == ui.Button.OK) {
      return [sessionRecent, lastColumn]; // Returns the date string and the OUTCOME column index
    }
  }
  return null;
}