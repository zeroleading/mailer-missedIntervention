/**
 * GLOBAL CONFIGURATION
 * Update these variables to change the behavior of the script without altering the core code.
 */

const CONFIG = {
  // --- Email Subject Line Settings ---
  _cachedSubject: null,
  get SUBJECT_NAME() {
    if (!this._cachedSubject) {
      const range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('selectedSubject');
      this._cachedSubject = range ? range.getValue() : "Subject Not Set"; 
    }
    return this._cachedSubject;
  },
  
  _cachedSession: null,
  get SESSION_TYPE() {
    if (!this._cachedSession) {
      const range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('selectedSessionType');
      this._cachedSession = range ? range.getValue() : "Session Not Set";
    }
    return this._cachedSession;
  },

  SUBJECT_TEMPLATE: "{{studentName}}: {{subjectName}} {{sessionType}} {{sessionDate}}",

  // --- Sheet Settings ---
  MESSAGE_SHEET_NAME: "message",
  TEMPLATE_SHEET_NAME: "cleanSheet", 
  SESSION_SHEET_NAME: "sessions",    
  STAR_SYMBOL: "★", 
  
  // --- Table Settings (for Setup) ---
  TABLE_SETUP_SHEET: "setup",
  TABLE_NAME_COMPULSORY: "Compulsory",
  TABLE_COL_STUDENT: "Student",
  TABLE_COL_START_DATE: "startDate", // NEW
  TABLE_COL_END_DATE: "endDate",     // NEW
  
  // --- Grid/Layout Indices (1-based) ---
  COL_RAW_STUDENT: 1,  
  COL_STUDENT_NAME: 2, 
  COL_STUDENT_EMAIL: 3, 
  COL_PARENT_EMAIL: 4, 
  COL_STAR: 5,         
  ROW_DATA_START: 3    
};