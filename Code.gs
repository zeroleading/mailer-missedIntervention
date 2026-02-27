const ui = SpreadsheetApp.getUi();
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getActiveSheet();

function onOpen() {
  ui.createMenu("Revision Follow-Up Functions")
    .addItem("Add new session", "addSession")
    .addItem("Set drafts for attendance/absence", "setDrafts")
    .addItem("Send emails for attendance/absence", "sendEmails")
    .addToUi();
}

const addSession = () => {
  const lastColumn = sheet.getRange(1, 1).getDataRegion().getLastColumn();
  const lastRow = sheet.getRange(1, 1).getDataRegion().getLastRow();
  const user = Session.getActiveUser();
  const date = new Date();
  const formattedDate = Utilities.formatDate(date, 'Europe/London', 'dd/MM/yy');

  const sessionDatePrompt = ui.prompt(
    'Revision session date',
    `Enter date below, today's date is ${formattedDate}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (sessionDatePrompt.getSelectedButton() == ui.Button.OK) {
    sheet.insertColumnsAfter(lastColumn, 2);

    const sessionDate = sessionDatePrompt.getResponseText() || formattedDate;

    sheet.setColumnWidths(lastColumn + 1, 2, 160);
    const sessionHeader = `Logged by: ${user}\n${sessionDate}`;
    sheet.getRange(1, lastColumn + 1).setValue(sessionHeader);
    sheet.getRange(1, lastColumn + 2).setValue('Outcome');

    const red = Math.floor(Math.random() * 50) + 205;
    const green = Math.floor(Math.random() * 50) + 205;
    const blue = Math.floor(Math.random() * 50) + 205;

    sheet.getRange(1, lastColumn + 1, 1, 2).setBackgroundRGB(red, green, blue);
    sheet.getRange(1, lastColumn + 1, 1, 2).setFontWeight('bold');
    sheet.getRange(1, lastColumn + 1, 1, 2).setWrap(true);

    sheet.getRange(1, lastColumn + 1, lastRow, 2).setFontFamily('Proxima Nova');

    const dvList = ['Attended', 'Absent', 'Unable to attend'];
    const dvRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dvList, true)
      .setAllowInvalid(false)
      .build();
    for (let i = 2; i <= lastRow; i++) {
      sheet.getRange(i, lastColumn + 1).setDataValidation(dvRule).setValue('');
    }

  } else {
    ui.alert('Session not added, try again');
  }
};

const setDrafts = () => {

  const user = Session.getActiveUser();
  const lastRow = sheet.getRange(1, 1).getDataRegion().getLastRow();

  const sessionElements = sessionDetails();

  if (sessionElements != null) {
    const sessionDate = sessionElements[0];
    const useColumn = sessionElements[1];
    const forScope = GmailApp.getInboxUnreadCount();

    for (let i = 2; i <= lastRow; i++) {
      if (
        sheet.getRange(i, useColumn - 1).getValue() === 'Attended' ||
        (sheet.getRange(i, 5).getValue() === '★' && sheet.getRange(i, useColumn - 1).getValue() === 'Absent')
      ) {
        const studentFirstName = sheet.getRange(i, 2).getValue();
        const attendance = sheet.getRange(i, useColumn - 1).getValue();
        const message = getMessage(studentFirstName, attendance);

        const parentEmailAddress = sheet.getRange(i, 4).getValue();
        const raw = `From: ${user}\r\n` +
          `To: ${parentEmailAddress}\r\n` +
          `Subject: ${studentFirstName.toUpperCase()}: GEOGRAPHY REVISION SESSION ${sessionDate}\r\n` +
          'Content-Type: text/html; charset=UTF-8\r\n' +
          '\r\n' +
          message;

        const draftBody = Utilities.base64Encode(raw, Utilities.Charset.UTF_8).replace(/\//g, '_').replace(/\+/g, '-');

        const params = {
          method: 'post',
          contentType: 'application/json',
          headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
          muteHttpExceptions: true,
          payload: JSON.stringify({
            message: {raw: draftBody},
          }),
        };

        const resp = UrlFetchApp.fetch('https://www.googleapis.com/gmail/v1/users/me/drafts', params);
        Logger.log(resp.getContentText());

        sheet.getRange(i, useColumn).setValue('Email drafted');
      }
    }

  } else {
    ui.alert('No drafts set, try again');
  }
};

const sendEmails = () => {
  const user = Session.getActiveUser();
  const lastRow = sheet.getRange(1, 1).getDataRegion().getLastRow();

  const sessionElements = sessionDetails();

  if (sessionElements != null) {
    const sessionDate = sessionElements[0];
    const useColumn = sessionElements[1];

    for (let i = 2; i <= lastRow; i++) {
      if (
        sheet.getRange(i, useColumn - 1).getValue() === 'Attended' ||
        (sheet.getRange(i, 5).getValue() === '★' && sheet.getRange(i, useColumn - 1).getValue() === 'Absent')
      ) {
        const studentFirstName = sheet.getRange(i, 2).getValue();
        const parentEmailAddress = sheet.getRange(i, 4).getValue();
        const attendance = sheet.getRange(i, useColumn - 1).getValue();
        const message = getMessage(studentFirstName, attendance);

        const subject = `${studentFirstName.toUpperCase()}: GEOGRAPHY INTERVENTION ${sessionDate}`;

        MailApp.sendEmail({
          to: parentEmailAddress,
          subject: subject,
          htmlBody: message,
        });

        sheet.getRange(i, useColumn).setValue('Email sent');
      }
    }

  } else {
    ui.alert('No emails sent, try again');
  }
};

/**
 * Gets the email message from a Google Sheet and personalizes it.
 *
 * @param {string} name - The name of the recipient.
 * @param {string} attendance - The attendance status ('Present' or 'Absent').
 * @returns {string} The complete, personalized email message as HTML.
 */

const getMessage = (name, attendance) => {

  const messageSheet = ss.getSheetByName('message');

  // Get the message parts from the specified cells.
  const absentMessage = messageSheet.getRange('D2').getValue();
  const attendedMessage = messageSheet.getRange('D3').getValue();
  const closingText = messageSheet.getRange('D4').getValue();
  const signature = messageSheet.getRange('D5').getValue();

  let selectedMessage;

  // Select the correct message based on attendance.
  if (attendance === 'Absent') {
    selectedMessage = absentMessage;
  } else {
    // If not 'Absent', assume they attended.
    selectedMessage = attendedMessage;
  }

  // Replace the name placeholder in the selected message and the closing text.
  selectedMessage = selectedMessage.replace(/\{\{name\}\}/g, name);
  const personalisedClosingText = closingText.replace(/\{\{name\}\}/g, name);

  // Combine the selected message, the closing text and signature.
  const finalMessage = `${selectedMessage}${personalisedClosingText}${signature}`;

  return finalMessage;
};

const sessionDetails = () => {
  const lastColumn = sheet.getRange(1, 1).getDataRegion().getLastColumn();

  let sessionRecent = sheet.getRange(1, lastColumn - 1).getValue();
  const wheresTheAt = sessionRecent.indexOf('@');

  if (wheresTheAt == -1) {
    ui.alert('Session has not been set correctly, please set at least one session before continuing');
  } else {
    sessionRecent = sessionRecent.substr(wheresTheAt + 12);

    const sessionIdentified = ui.alert(
      `You are about to send emails/set drafts relating to the session\n${sessionRecent}`,
      ui.ButtonSet.OK_CANCEL
    );

    if (sessionIdentified == ui.Button.OK) {
      return [sessionRecent, lastColumn];
    }
  }
};
