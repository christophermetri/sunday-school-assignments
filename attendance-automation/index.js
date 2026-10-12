#!/usr/bin/env node

/**
 * Sunday School Attendance Automation - Main Orchestrator
 *
 * Workflow:
 * 1. Initialize Google APIs
 * 2. Fetch latest PDF from Google Drive
 * 3. Parse PDF to extract attendance data
 * 4. Get absence history from Google Sheet
 * 5. Calculate consecutive absences
 * 6. Format email summary
 * 7. Send email to coordinator
 * 8. Update Google Sheet with new data
 */

require('dotenv').config();

const CONFIG = require('./config');
const googleApis = require('./googleApis');
const { parseAttendancePdf, formatAttendanceData } = require('./pdfParser');
const { formatAttendanceEmail, calculateConsecutiveWeeks } = require('./emailFormatter');

/**
 * Main execution function
 */
async function main() {
  try {
    console.log('🚀 Starting Sunday School Attendance Automation...\n');

    // Step 1: Initialize Google APIs
    console.log('Step 1: Initializing Google APIs...');
    await googleApis.initializeAuth();

    // Step 2: Fetch latest PDF from Drive
    console.log('\nStep 2: Fetching latest attendance PDF...');
    const latestPdf = await googleApis.getLatestPdfFromDrive(CONFIG.DRIVE_FOLDER_ID);

    if (!latestPdf) {
      console.log('No PDFs found. Exiting.');
      return;
    }

    // Step 3: Download and parse PDF
    console.log('\nStep 3: Downloading and parsing PDF...');
    const pdfStream = await googleApis.downloadPdfFromDrive(latestPdf.id);
    const pdfBuffer = await streamToBuffer(pdfStream);
    const attendanceData = await parseAttendancePdf(pdfBuffer);

    console.log('Parsed attendance data:');
    console.log(formatAttendanceData(attendanceData));

    // Step 4: Get absence history from sheet
    console.log('\nStep 4: Retrieving absence history...');
    const absenceHistory = await googleApis.getAbsenceHistory(CONFIG.SHEETS_ID);

    // Step 5: Calculate consecutive absences
    console.log('\nStep 5: Calculating consecutive absences...');
    const enrichedAbsenceData = enrichAbsenceData(
      attendanceData,
      absenceHistory,
      attendanceData.date
    );

    // Step 6: Format email
    console.log('\nStep 6: Formatting email...');
    const emailBody = formatAttendanceEmail(enrichedAbsenceData, enrichedAbsenceData.absenceHistoryEnriched);

    // Step 7: Send email
    console.log('\nStep 7: Sending email...');
    const subject = `${CONFIG.EMAIL_SUBJECT_PREFIX} ${attendanceData.date}`;
    await googleApis.sendEmail(CONFIG.EMAIL_TO, subject, emailBody);

    // Step 8: Update Google Sheet
    console.log('\nStep 8: Updating Google Sheet...');
    const sheetData = prepareSheetData(enrichedAbsenceData, attendanceData.date);
    await googleApis.updateAttendanceSheet(CONFIG.SHEETS_ID, sheetData);

    console.log('\n✅ Automation complete!');

  } catch (error) {
    console.error('\n❌ Error during automation:', error.message);
    console.error(error);
    process.exit(1);
  }
}

/**
 * Convert stream to buffer
 */
async function streamToBuffer(stream) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    stream.on('data', chunk => chunks.push(chunk));
    stream.on('end', () => resolve(Buffer.concat(chunks)));
    stream.on('error', reject);
  });
}

/**
 * Enrich absence data with consecutive week calculations
 */
function enrichAbsenceData(attendanceData, absenceHistory, currentDate) {
  const enriched = JSON.parse(JSON.stringify(attendanceData));
  enriched.absenceHistoryEnriched = {};

  for (const [className, servants] of Object.entries(enriched.servantAbsences)) {
    servants.forEach(servantName => {
      const consecutiveWeeks = calculateConsecutiveWeeks(servantName, absenceHistory, currentDate);
      if (!enriched.absenceHistoryEnriched[servantName]) {
        enriched.absenceHistoryEnriched[servantName] = {
          className,
          consecutiveWeeks
        };
      }
    });
  }

  return enriched;
}

/**
 * Prepare data for Google Sheet update
 */
function prepareSheetData(enrichedData, dateStr) {
  const sheetData = [];

  for (const [className, servants] of Object.entries(enrichedData.servantAbsences)) {
    servants.forEach(servantName => {
      const historyInfo = enrichedData.absenceHistoryEnriched[servantName] || {};
      sheetData.push({
        date: dateStr,
        className,
        servantName,
        wasAbsent: true,
        consecutiveWeeks: historyInfo.consecutiveWeeks || 1,
        notes: ''
      });
    });
  }

  return sheetData;
}

// Run the automation
if (require.main === module) {
  main();
}

module.exports = { main };
