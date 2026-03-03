/**
 * Creates the custom menu when the spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Revision Follow-Up Functions")
    .addItem("1. Setup Sessions (From Table)", "setupSessions")
    .addSeparator()
    .addItem("2. Add new session column", "addSession")
    .addItem("3. Set drafts for attendance/absence", "setDrafts")
    .addItem("4. Send emails for attendance/absence", "sendEmails")
    .addSeparator()
    .addItem("⭐ Convert to Compulsory (Active Row)", "makeCompulsory")
    .addItem("❌ Convert to Non-Compulsory (Active Row)", "makeNonCompulsory")
    .addItem("🔄 Sync & Sort Register", "syncAndSortRegister")
    .addSeparator()
    .addItem("🔍 Preview Email (Active Row)", "openPreviewModal")
    .addToUi();
}