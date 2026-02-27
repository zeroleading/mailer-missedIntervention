/**
 * GLOBAL CONFIGURATION
 * Update these variables to change the behavior of the script without altering the core code.
 */

const CONFIG = {
  // --- Email Subject Line Settings ---
  SUBJECT_NAME: "Geography", 
  SESSION_TYPE: "Intervention", 
  SUBJECT_TEMPLATE: "{{studentName}}: {{subjectName}} {{sessionType}} {{sessionDate}}",

  // --- Sheet Settings ---
  MESSAGE_SHEET_NAME: "message",
  TEMPLATE_SHEET_NAME: "cleanSheet", // The hidden template
  SESSION_SHEET_NAME: "sessions",    // The newly created sheet
  STAR_SYMBOL: "★", 
  
  // --- Table Settings (for Setup) ---
  TABLE_SETUP_SHEET: "setup",
  TABLE_NAME_COMPULSORY: "Compulsory",
  TABLE_COL_STUDENT: "Student",
  
  // --- Grid/Layout Indices (1-based) ---
  COL_RAW_STUDENT: 1,  // The raw data pasted from the Table (Col A)
  COL_STUDENT_NAME: 2, // The calculated First Name via BYROW (Col B)
  COL_STUDENT_EMAIL: 3, // Student email for CC (Col C)
  COL_PARENT_EMAIL: 4, // Parent Email (Col D)
  COL_STAR: 5,         // Star Tracker (Col E)
  ROW_DATA_START: 3    // Row 1: Headers, Row 2: Hidden BYROWs, Row 3: Data start
};