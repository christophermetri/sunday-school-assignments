/**
 * Configuration for Sunday School Attendance Automation
 */

const CONFIG = {
  // Google Drive
  DRIVE_FOLDER_ID: '1Un-scNhfDC_QRVKDb9DvD49rEB1TkEwY', // Church Teams Attendance folder

  // Google Sheets
  SHEETS_ID: '1CgfXd2B_dlAMDBqBkhzzU9jSCYh7mxFpXxcUS97u4q4', // Attendance Tracking sheet
  SHEETS_RANGE: 'Absences!A:F', // Range for tracking data

  // Email
  EMAIL_FROM: 'stabraammain@gmail.com',
  EMAIL_TO: 'christophermetri@gmail.com',
  EMAIL_SUBJECT_PREFIX: 'Attendance for',

  // Absence tracking thresholds
  ABSENCE_THRESHOLDS: {
    CRITICAL: 4, // 4+ weeks = "hasn't attended in 4 weeks"
  },

  // Class names (used for formatting)
  CLASSES: [
    'Nursery',
    'PreK',
    'Kindergarten',
    'Grade 1',
    'Grade 2',
    'Grade 3',
    'Grade 4-5',
    'Grade 6-7',
    'Middle School',
    '8th Grade',
    'High School',
    'Special Needs'
  ]
};

module.exports = CONFIG;
