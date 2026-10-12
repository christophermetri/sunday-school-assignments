/**
 * Google APIs utilities for Drive, Sheets, and Gmail integration
 */

const { google } = require('googleapis');
const { auth } = require('google-auth-library');
const fs = require('fs');
const path = require('path');

let authClient = null;

/**
 * Initialize Google Auth (uses Application Default Credentials)
 */
async function initializeAuth() {
  try {
    // Try to use Application Default Credentials
    const auth = new (require('google-auth-library')).GoogleAuth({
      scopes: [
        'https://www.googleapis.com/auth/drive.readonly',
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/gmail.send'
      ]
    });

    authClient = await auth.getClient();
    console.log('✓ Google Auth initialized (ADC)');
    return authClient;
  } catch (error) {
    console.error('Error initializing Google Auth:', error.message);
    throw new Error('Failed to initialize Google authentication. Ensure GOOGLE_APPLICATION_CREDENTIALS is set.');
  }
}

/**
 * Get the latest PDF from Google Drive folder
 */
async function getLatestPdfFromDrive(folderId) {
  try {
    const drive = google.drive({ version: 'v3', auth: authClient });

    const result = await drive.files.list({
      q: `'${folderId}' in parents and mimeType='application/pdf' and trashed=false`,
      pageSize: 10,
      fields: 'files(id, name, modifiedTime, webContentLink)',
      orderBy: 'modifiedTime desc'
    });

    if (!result.data.files || result.data.files.length === 0) {
      console.log('No PDFs found in folder');
      return null;
    }

    const latestFile = result.data.files[0];
    console.log(`✓ Found latest PDF: ${latestFile.name}`);
    return latestFile;
  } catch (error) {
    console.error('Error accessing Google Drive:', error.message);
    throw error;
  }
}

/**
 * Download PDF file from Google Drive
 */
async function downloadPdfFromDrive(fileId) {
  try {
    const drive = google.drive({ version: 'v3', auth: authClient });

    const res = await drive.files.get(
      { fileId: fileId, alt: 'media' },
      { responseType: 'stream' }
    );

    console.log(`✓ Downloaded PDF content for file ID: ${fileId}`);
    return res.data;
  } catch (error) {
    console.error('Error downloading PDF:', error.message);
    throw error;
  }
}

/**
 * Update Google Sheet with absence data
 */
async function updateAttendanceSheet(sheetId, data) {
  try {
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    // Prepare values for sheet update
    const values = data.map(row => [
      row.date,
      row.className,
      row.servantName,
      row.wasAbsent ? 'YES' : 'NO',
      row.consecutiveWeeks || 0,
      row.notes || ''
    ]);

    const result = await sheets.spreadsheets.values.append({
      spreadsheetId: sheetId,
      range: 'Absences!A:F',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: values
      }
    });

    console.log(`✓ Updated sheet: ${result.data.updates.updatedRows} rows added`);
    return result.data;
  } catch (error) {
    console.error('Error updating sheet:', error.message);
    throw error;
  }
}

/**
 * Get current absence data from Google Sheet
 */
async function getAbsenceHistory(sheetId) {
  try {
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: 'Absences!A:F'
    });

    const rows = result.data.values || [];
    console.log(`✓ Retrieved ${rows.length} rows from sheet`);

    // Parse into usable format
    const absenceMap = {};
    for (let i = 1; i < rows.length; i++) { // Skip header
      const [date, className, servantName, wasAbsent, consecutiveWeeks] = rows[i];
      if (servantName) {
        absenceMap[servantName] = {
          date,
          className,
          wasAbsent: wasAbsent === 'YES',
          consecutiveWeeks: parseInt(consecutiveWeeks) || 0
        };
      }
    }

    return absenceMap;
  } catch (error) {
    console.error('Error retrieving absence history:', error.message);
    throw error;
  }
}

/**
 * Send email via Gmail
 */
async function sendEmail(to, subject, htmlBody) {
  try {
    const gmail = google.gmail({ version: 'v1', auth: authClient });

    const message = {
      raw: Buffer.from(
        `From: stabraammain@gmail.com\r\n` +
        `To: ${to}\r\n` +
        `Subject: ${subject}\r\n` +
        `Content-Type: text/html; charset="UTF-8"\r\n\r\n` +
        htmlBody
      ).toString('base64')
    };

    const result = await gmail.users.messages.send({
      userId: 'me',
      resource: message
    });

    console.log(`✓ Email sent successfully (Message ID: ${result.data.id})`);
    return result.data;
  } catch (error) {
    console.error('Error sending email:', error.message);
    throw error;
  }
}

module.exports = {
  initializeAuth,
  getLatestPdfFromDrive,
  downloadPdfFromDrive,
  updateAttendanceSheet,
  getAbsenceHistory,
  sendEmail
};
