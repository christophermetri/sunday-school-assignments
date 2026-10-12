# Sunday School Attendance Automation

Automated attendance tracking and email summary generator for St. Abraam Sunday School.

## Overview

This system automates the Sunday School attendance workflow:

1. **PDF Capture** (via Google Apps Script)
   - Church Teams emails with attendance PDFs are automatically saved to Google Drive

2. **PDF Processing** (this script)
   - Extracts attendance data from the PDF
   - Identifies absent servants and classes that didn't take attendance
   - Calculates consecutive absence weeks
   - Sends formatted email summary to coordinator

3. **Data Tracking** (via Google Sheets)
   - Records weekly attendance data
   - Tracks consecutive absences for pattern identification

## Setup Instructions

### 1. Google Cloud Project Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable these APIs:
   - Google Drive API
   - Google Sheets API
   - Gmail API

4. Create a Service Account:
   - Go to Credentials → Create Credentials → Service Account
   - Download the JSON key file
   - Save as `service-account-key.json` in this directory

5. Share Google Resources with the Service Account:
   - **Google Drive folder**: Grant Editor access to the service account email
   - **Google Sheets**: Grant Editor access to the service account email
   - **Gmail account**: The service account needs sending permissions (may need OAuth setup)

### 2. Install Dependencies

```bash
cd attendance-automation
npm install
```

### 3. Configuration

Copy the environment template:
```bash
cp .env.example .env
```

Update `.env` with your Google service account credentials.

### 4. Update Config

Edit `config.js` with your specific IDs:
- `DRIVE_FOLDER_ID`: Your Google Drive folder for attendance PDFs
- `SHEETS_ID`: Your Google Sheets ID for tracking
- `EMAIL_FROM`: Church email (stabraammain@gmail.com)
- `EMAIL_TO`: Coordinator email (christophermetri@gmail.com)

## Running the Automation

### Manual Run (Testing)
```bash
npm start
```

### Automated Scheduling (via Google Apps Script)

The Google Apps Script trigger (already configured in the main repo) will call this automation daily:

**Apps Script function** calls this automation when a new PDF arrives:
```javascript
// In Google Apps Script:
// After saveChurchTeamsAttendance() saves a PDF,
// trigger this Node.js script via Cloud Function or external API
```

## File Structure

```
attendance-automation/
├── index.js                 # Main orchestrator
├── config.js               # Configuration constants
├── googleApis.js           # Google APIs utilities
├── pdfParser.js            # PDF extraction logic
├── emailFormatter.js       # Email formatting
├── package.json            # Node.js dependencies
├── .env.example            # Environment template
└── README.md              # This file
```

## Email Format

The system generates emails with this structure:

```
Attendance for 10-12-2026

Classes that didn't take attendance:
• Nursery
• Grade 1

Absent from PreK:
• Sarah Johnson (2nd week in a row)
• Michael Brown

Absent from Grade 2:
• Chris Metri (hasn't attended in 4 weeks)
```

## Consecutive Absence Tracking

- **1st absence**: No notation
- **2nd-3rd consecutive**: "(2nd week in a row)" or "(3rd week in a row)"
- **4+ consecutive**: "(hasn't attended in X weeks)"

## Error Handling

The script logs detailed error messages if:
- Google authentication fails
- PDF parsing encounters issues
- Sheet updates fail
- Email sending fails

All errors are logged to console and the process exits with status code 1.

## Future Enhancements

- [ ] Servant name matching improvements (handle variations in names)
- [ ] PDF format learning (auto-adapt to different Church Teams exports)
- [ ] Slack integration for real-time notifications
- [ ] Web dashboard for viewing historical attendance
- [ ] Servant list sync from GitHub

## Support

For issues or questions, contact the Sunday School coordinator.
