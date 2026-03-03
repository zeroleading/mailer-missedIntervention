/**
 * Opens the modal for drafting emails.
 */
function openDraftSelector() {
  openSessionModal('draft');
}

/**
 * Opens the modal for sending emails.
 */
function openSendSelector() {
  openSessionModal('send');
}

/**
 * Core function to open the modal dialog.
 * @param {string} actionType - 'draft' or 'send'
 */
function openSessionModal(actionType) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  if (sheet.getName() !== CONFIG.SESSION_SHEET_NAME) {
    ui.alert(`Please navigate to your '${CONFIG.SESSION_SHEET_NAME}' sheet.`);
    return;
  }

  const template = HtmlService.createTemplateFromFile('-11- SessionSelectorUI');
  template.actionType = actionType;
  
  const htmlOutput = template.evaluate()
      .setWidth(450)
      .setHeight(300)
      .setTitle(actionType === 'draft' ? 'Select Session (Drafts)' : 'Select Session (Send Live)');
      
  ui.showModalDialog(htmlOutput, actionType === 'draft' ? 'Create Email Drafts' : 'Send Live Emails');
}

/**
 * Scans the sheet to find all sessions and groups them by processed state.
 * Called by the HTML Modal.
 * @returns {Object} { pending: [], processed: [] }
 */
function getSessionList() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastColumn = sheet.getRange(1, 1).getDataRegion().getLastColumn();
  
  const pending = [];
  const processed = [];

  if (lastColumn <= CONFIG.COL_STAR) return { pending, processed };

  // Fetch all headers (Row 1) from the star column onwards
  const headers = sheet.getRange(1, CONFIG.COL_STAR + 1, 1, lastColumn - CONFIG.COL_STAR).getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    const headerText = headers[i] ? headers[i].toString() : "";
    const absoluteColIndex = CONFIG.COL_STAR + 1 + i;
    
    // Check if this is a "Logged by" header
    if (headerText.indexOf('@') !== -1) {
      const sessionDate = headerText.substring(headerText.indexOf('@') + 12).trim();
      
      // The outcome header is the very next column
      const outcomeColIndex = absoluteColIndex + 1;
      const outcomeHeader = (outcomeColIndex <= lastColumn) ? sheet.getRange(1, outcomeColIndex).getValue().toString() : "";
      
      const sessionData = {
        date: sessionDate,
        outcomeCol: outcomeColIndex
      };

      if (outcomeHeader.includes('Processed') || outcomeHeader.includes('✅')) {
        processed.push(sessionData);
      } else if (outcomeHeader.includes('Outcome')) {
        pending.push(sessionData);
      }
    }
  }

  return { pending, processed };
}

/**
 * Triggered by the HTML modal to begin processing.
 * @param {number} outcomeColIndex - The column to process.
 * @param {string} actionType - 'draft' or 'send'.
 */
function executeSessionProcess(outcomeColIndex, actionType) {
  const isDraft = (actionType === 'draft');
  processCommunications(isDraft, outcomeColIndex);
}
