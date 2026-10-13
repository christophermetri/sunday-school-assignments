# Sunday School Servant Attendance (Google Apps Script)

Runs entirely on Google's servers under the **stabraammain@gmail.com** account. No computer needs to be on.

## What it does
Every hour it checks the church inbox for the Church Teams scheduled report
("Scheduled Report - Sunday School Attendance"). When a new one arrives it:

1. Saves the PDF to the **Church Teams Attendance** Drive folder.
2. Reads which group members were marked present (X).
3. Checks each servant listed on the servants page (`index.html` in this repo, read live from GitHub).
4. Counts consecutive absences using the **History** tab of the *Sunday School Attendance Tracking* sheet.
5. Emails the summary to christophermetri@gmail.com and logs the week to **History**.

## Sheet tabs
- **History** – one row per servant per week (Present / Absent / Not found).
- **Name Fixes** – when a servant's name on the website doesn't match Church Teams, add a row:
  website name → Church Teams name exactly as written (`Last, First`). Write `SKIP` to ignore someone who doesn't check in.
- **Servants** – a copy of the roster from the website (refreshed every run, used as a backup).

## Changing who gets the email
Edit `EMAIL_TO` at the top of `Code.gs` in the Apps Script editor (script.google.com, signed in as stabraammain).
Separate several addresses with commas.

## Manual runs (Apps Script editor)
- `setup()` – one-time: creates the tabs and the hourly trigger, and logs any reports already in the inbox (without emailing).
- `runLatestReport()` – re-processes the most recent report and re-sends the email.

`Code.gs` here is the source of truth; paste it into the editor after changes.
