const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SCOPES = [
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/spreadsheets',
  'https://mail.google.com/'
];

const TOKEN_PATH = path.join(__dirname, 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');

let auth = null;

async function initializeAuth() {
  try {
    // Load client credentials from OAuth setup
    const credentialsContent = fs.readFileSync(CREDENTIALS_PATH, 'utf8');
    const credentials = JSON.parse(credentialsContent);
    
    const { client_id, client_secret, redirect_uris } = credentials.installed;
    
    const oauth2Client = new google.auth.OAuth2(
      client_id,
      client_secret,
      redirect_uris[0]
    );

    // Check if we have saved tokens
    if (fs.existsSync(TOKEN_PATH)) {
      const token = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
      oauth2Client.setCredentials(token);
      auth = oauth2Client;
      console.log('✅ Using saved OAuth tokens');
      return oauth2Client;
    }

    // If no token, need to authenticate
    console.log('⚠️ No saved tokens. Need to authenticate first.');
    console.log('Run: node authenticate.js');
    throw new Error('Please run authenticate.js first to authorize the app');
  } catch (error) {
    console.error('❌ Auth Error:', error.message);
    throw error;
  }
}

async function getLatestPdfFromDrive() {
  try {
    const drive = google.drive({ version: 'v3', auth });
    const FOLDER_ID = require('./config').DRIVE_FOLDER_ID;

    const response = await drive.files.list({
      q: `'${FOLDER_ID}' in parents and mimeType='application/pdf' and trashed=false`,
      spaces: 'drive',
      fields: 'files(id, name, createdTime)',
      orderBy: 'createdTime desc',
      pageSize: 1,
    });

    if (response.data.files.length === 0) {
      return null;
    }

    return response.data.files[0];
  } catch (error) {
    console.error('❌ Error fetching PDF from Drive:', error.message);
    throw error;
  }
}

async function downloadPdfFromDrive(fileId) {
  try {
    const drive = google.drive({ version: 'v3', auth });
    const response = await drive.files.get(
      { fileId, alt: 'media' },
      { responseType: 'arraybuffer' }
    );
    return Buffer.from(response.data);
  } catch (error) {
    console.error('❌ Error downloading PDF:', error.message);
    throw error;
  }
}

async function updateAttendanceSheet(data) {
  try {
    const sheets = google.sheets({ version: 'v4', auth });
    const SHEETS_ID = require('./config').SHEETS_ID;

    const values = [
      [new Date().toISOString(), JSON.stringify(data)]
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEETS_ID,
      range: 'Sheet1!A:B',
      valueInputOption: 'RAW',
      resource: { values }
    });

    console.log('✅ Data saved to Google Sheet');
  } catch (error) {
    console.error('❌ Error updating sheet:', error.message);
    throw error;
  }
}

async function getAbsenceHistory() {
  try {
    const sheets = google.sheets({ version: 'v4', auth });
    const SHEETS_ID = require('./config').SHEETS_ID;

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEETS_ID,
      range: 'Sheet1!A:B'
    });

    return response.data.values || [];
  } catch (error) {
    console.error('❌ Error reading sheet:', error.message);
    throw error;
  }
}

async function sendEmail(to, subject, htmlContent) {
  try {
    const gmail = google.gmail({ version: 'v1', auth });
    const { EMAIL_FROM } = require('./config');

    const emailMessage = [
      `From: ${EMAIL_FROM}`,
      `To: ${to}`,
      `Subject: ${subject}`,
      'Content-Type: text/html; charset=utf-8',
      '',
      htmlContent
    ].join('\n');

    const encodedMessage = Buffer.from(emailMessage).toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '');

    await gmail.users.messages.send({
      userId: 'me',
      resource: { raw: encodedMessage }
    });

    console.log('✅ Email sent');
  } catch (error) {
    console.error('❌ Error sending email:', error.message);
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
