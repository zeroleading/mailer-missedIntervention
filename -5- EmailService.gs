/**
 * Processes the active sheet to CREATE DRAFTS for students.
 */
function setDrafts() {
  processCommunications(true);
}

/**
 * Processes the active sheet to SEND EMAILS for students.
 */
function sendEmails() {
  processCommunications(false);
}

/**
 * Core engine to read sheet data in batch, evaluate logic, and send/draft emails.
 * @param {boolean} isDraft - True if creating drafts, false if sending emails directly.
 */
function processCommunications(isDraft) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  const details = sessionDetails();
  if (!details) return; // User cancelled or error
  
  const [sessionDate, outcomeColIndex] = details;
  const attendanceColIndex = outcomeColIndex - 1; 

  // Batch Read: Get all sheet data at once to improve speed
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // Prep array to hold the updated statuses (to write back in batch)
  const updatedStatuses = [];

  // Fetch base templates from the message sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const msgSheet = ss.getSheetByName(CONFIG.MESSAGE_SHEET_NAME);
  const absentMsgRaw = msgSheet.getRange('D2').getValue();
  const attendedMsgRaw = msgSheet.getRange('D3').getValue();
  const closingText = msgSheet.getRange('D4').getValue();
  const signature = msgSheet.getRange('D5').getValue();

  // Loop through data rows (arrays are 0-indexed, so Row 3 is index 2)
  for (let i = CONFIG.ROW_DATA_START - 1; i < values.length; i++) {
    const row = values[i];
    const studentName = row[CONFIG.COL_STUDENT_NAME - 1];
    const studentEmail = row[CONFIG.COL_STUDENT_EMAIL - 1];
    const parentEmail = row[CONFIG.COL_PARENT_EMAIL - 1];
    const starStatus = row[CONFIG.COL_STAR - 1];
    const attendance = row[attendanceColIndex - 1];
    let currentOutcome = row[outcomeColIndex - 1];

    // Check Logic: Did they attend, OR are they absent WITH a star?
    if (attendance === 'Attended' || (starStatus === CONFIG.STAR_SYMBOL && attendance === 'Absent')) {
      
      // Select base text
      const baseText = (attendance === 'Absent') ? absentMsgRaw : attendedMsgRaw;
      
      // BLANK TEMPLATE CHECK: Skip if template text is missing (e.g. "Absence Only" usage)
      if (!baseText || baseText.toString().trim() === "") {
        updatedStatuses.push([currentOutcome]); // Keep existing outcome
        continue; // Skip this student
      }
      
      // Replace placeholders
      let personalisedBody = `${baseText}${closingText}${signature}`;
      personalisedBody = personalisedBody.replace(/\{\{name\}\}/g, studentName);
      
      // Preserve line breaks for HTML
      personalisedBody = personalisedBody.replace(/\n/g, '<br>');

      // Generate final HTML using the Shell
      const htmlBody = buildHtmlEmail(personalisedBody);

      // Generate Subject line
      let subject = CONFIG.SUBJECT_TEMPLATE
        .replace(/\{\{studentName\}\}/g, studentName)
        .replace(/\{\{subjectName\}\}/g, CONFIG.SUBJECT_NAME)
        .replace(/\{\{sessionType\}\}/g, CONFIG.SESSION_TYPE)
        .replace(/\{\{sessionDate\}\}/g, sessionDate);

      // Execute Communication
      if (isDraft) {
        GmailApp.createDraft(parentEmail, subject, "", {
          htmlBody: htmlBody,
          cc: studentEmail
        });
        currentOutcome = 'Email drafted';
      } else {
        MailApp.sendEmail(parentEmail, subject, "", {
          htmlBody: htmlBody,
          cc: studentEmail
        });
        currentOutcome = 'Email sent';
      }
    }
    
    // Store the outcome state for this row
    updatedStatuses.push([currentOutcome]);
  }

  // Batch Write: Put all statuses back in the outcome column in one go
  if (updatedStatuses.length > 0) {
    sheet.getRange(CONFIG.ROW_DATA_START, outcomeColIndex, updatedStatuses.length, 1).setValues(updatedStatuses);
  }
}

/**
 * Wraps the text message inside the HTML template shell.
 * @param {string} bodyContent - The personalised text to inject.
 * @returns {string} The fully compiled HTML string.
 */
function buildHtmlEmail(bodyContent) {
  const template = HtmlService.createTemplateFromFile('-6- EmailTemplate');
  template.messageBody = bodyContent;
  
  // FLATTEN the config object to remove getters before passing to HtmlService
  template.config = {
    SUBJECT_NAME: CONFIG.SUBJECT_NAME,
    SESSION_TYPE: CONFIG.SESSION_TYPE
  };
  
  return template.evaluate().getContent();
}