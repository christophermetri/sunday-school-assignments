/**
 * Email Formatter for Attendance Summary
 * Formats attendance data into HTML email
 */

/**
 * Format attendance summary into HTML email body
 *
 * Format:
 * - Header: "Attendance for [Date]"
 * - Classes that didn't take attendance
 * - For each class: Absent servants with (consecutive weeks) notation
 */
function formatAttendanceEmail(attendanceData, absenceHistory) {
  const { date, classesNotTakingAttendance, servantAbsences } = attendanceData;

  let htmlContent = `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
          h1 { color: #1a4d8a; border-bottom: 2px solid #1a4d8a; padding-bottom: 10px; }
          h2 { color: #2e5f9a; margin-top: 20px; font-size: 16px; }
          .no-attendance { background-color: #f5f5f5; padding: 15px; border-left: 4px solid #ff9800; }
          .absent-list { background-color: #fafafa; padding: 15px; border-left: 4px solid #d32f2f; }
          .servant-name { margin: 8px 0; }
          .note { font-size: 12px; color: #666; }
          .date-note { color: #666; font-size: 13px; margin-top: 30px; }
        </style>
      </head>
      <body>
        <h1>Attendance for ${date}</h1>
  `;

  // Classes that didn't take attendance
  if (classesNotTakingAttendance.length > 0) {
    htmlContent += `
      <div class="no-attendance">
        <h2>Classes that didn't take attendance:</h2>
        <ul>
    `;

    classesNotTakingAttendance.forEach(className => {
      htmlContent += `<li>${className}</li>`;
    });

    htmlContent += `
        </ul>
      </div>
    `;
  }

  // Absent servants by class
  for (const [className, servants] of Object.entries(servantAbsences)) {
    if (servants.length > 0) {
      htmlContent += `
        <div class="absent-list">
          <h2>Absent from ${className}:</h2>
          <ul>
      `;

      servants.forEach(servantName => {
        const history = absenceHistory[servantName];
        let servantEntry = `<li class="servant-name"><strong>${servantName}</strong>`;

        if (history) {
          if (history.consecutiveWeeks >= 4) {
            servantEntry += ` <span class="note">(hasn't attended in ${history.consecutiveWeeks} weeks)</span>`;
          } else if (history.consecutiveWeeks > 1) {
            servantEntry += ` <span class="note">(${history.consecutiveWeeks}${getWeekSuffix(history.consecutiveWeeks)} week in a row)</span>`;
          }
        }

        servantEntry += '</li>';
        htmlContent += servantEntry;
      });

      htmlContent += `
          </ul>
        </div>
      `;
    }
  }

  htmlContent += `
        <p class="date-note">Report generated automatically on ${new Date().toLocaleString()}</p>
      </body>
    </html>
  `;

  return htmlContent;
}

/**
 * Get ordinal suffix for week count (st, nd, rd, th)
 */
function getWeekSuffix(num) {
  const j = num % 10;
  const k = num % 100;

  if (j === 1 && k !== 11) return 'st';
  if (j === 2 && k !== 12) return 'nd';
  if (j === 3 && k !== 13) return 'rd';
  return 'th';
}

/**
 * Calculate consecutive weeks for a servant
 */
function calculateConsecutiveWeeks(servantName, absenceHistory, currentDate) {
  if (!absenceHistory[servantName]) {
    return 1; // First absence this week
  }

  const history = absenceHistory[servantName];

  // Check if absence history is recent (within last 7 days)
  const historyDate = new Date(history.date);
  const currentDateObj = new Date(currentDate);
  const daysDiff = Math.floor((currentDateObj - historyDate) / (1000 * 60 * 60 * 24));

  if (daysDiff <= 7) {
    // Same week or consecutive
    return (history.consecutiveWeeks || 1) + 1;
  }

  return 1; // Reset if gap in weeks
}

module.exports = {
  formatAttendanceEmail,
  calculateConsecutiveWeeks
};
