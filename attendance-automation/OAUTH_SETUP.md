# OAuth 2.0 Setup Guide

## Step 1: Create OAuth Credentials in Google Cloud

1. Go to https://console.cloud.google.com/apis/credentials
2. Make sure "Sunday School Attendance" project is selected
3. Click **"+ Create Credentials"** → **"OAuth client ID"**
4. Choose **"Desktop application"**
5. Download the JSON file and save it as `credentials.json` in this folder

## Step 2: Run Authentication

```bash
node authenticate.js
```

This will:
- Open a browser window asking you to authorize the app
- Authorize it with stabraammain@gmail.com
- Save the tokens to `token.json` (auto-generated)

## Step 3: Done!

The automation will now use the saved OAuth tokens. Everything works automatically from here.

## What permissions are granted?

- **Google Drive**: Read-only access to the attendance PDF folder
- **Google Sheets**: Read/Write access to the attendance tracking sheet
- **Gmail**: Send-only access to send attendance summaries

All through the church's own account (stabraammain@gmail.com).
