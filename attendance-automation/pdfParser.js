/**
 * PDF Parser for Church Teams Attendance Reports
 * Extracts class attendance and servant absence data
 */

const pdfParse = require('pdf-parse');

/**
 * Extract attendance data from PDF buffer
 *
 * PDF Structure (expected format):
 * - Header with date
 * - Table with columns: Class Name, Servants, Attendance Status
 * - "Absent" or marked individuals
 */
async function parseAttendancePdf(pdfBuffer) {
  try {
    // Parse PDF
    const pdfData = await pdfParse(pdfBuffer);
    const text = pdfData.text;

    console.log('✓ PDF parsed successfully');

    // Extract date from PDF (usually in header)
    const dateMatch = text.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    let attendanceDate = null;
    if (dateMatch) {
      const [, month, day, year] = dateMatch;
      attendanceDate = `${month.padStart(2, '0')}-${day.padStart(2, '0')}-${year}`;
    }

    // Parse attendance data
    const attendanceData = {
      date: attendanceDate,
      classesNotTakingAttendance: [],
      servantAbsences: {} // className -> [servants]
    };

    // Parse text content
    // This is a placeholder - actual parsing depends on PDF structure
    const lines = text.split('\n');
    let currentClass = null;

    for (const line of lines) {
      const trimmed = line.trim();

      // Look for class headers
      if (isClassHeader(trimmed)) {
        currentClass = trimmed;
        attendanceData.servantAbsences[currentClass] = [];
      }

      // Look for "Absent" markers
      if (trimmed.includes('Absent') && currentClass) {
        // Extract servant name from context
        const servantName = extractServantName(trimmed, lines);
        if (servantName) {
          attendanceData.servantAbsences[currentClass].push(servantName);
        }
      }

      // Look for "Did not take attendance"
      if (trimmed.includes('Did not take attendance') || trimmed.includes('No attendance')) {
        attendanceData.classesNotTakingAttendance.push(currentClass);
      }
    }

    return attendanceData;
  } catch (error) {
    console.error('Error parsing PDF:', error.message);
    throw error;
  }
}

/**
 * Check if line appears to be a class header
 */
function isClassHeader(line) {
  const classKeywords = [
    'Nursery', 'PreK', 'Kindergarten', 'Grade',
    'Middle School', 'High School', 'Special Needs',
    'Class 1', 'Class 2', 'Class 3'
  ];

  return classKeywords.some(keyword => line.includes(keyword));
}

/**
 * Extract servant name from context around "Absent" marker
 */
function extractServantName(line, allLines) {
  // Try to extract name from line containing "Absent"
  const namePattern = /^([A-Z][a-z]+ [A-Z][a-z]+)/;
  const match = line.match(namePattern);

  if (match) {
    return match[1];
  }

  // If not found in same line, check surrounding lines
  // This is a basic approach; may need refinement based on actual PDF format
  return null;
}

/**
 * Format parsed attendance data for readability
 */
function formatAttendanceData(data) {
  let output = `Attendance Report for ${data.date}\n`;
  output += '='.repeat(50) + '\n\n';

  if (data.classesNotTakingAttendance.length > 0) {
    output += 'Classes that did not take attendance:\n';
    data.classesNotTakingAttendance.forEach(cls => {
      output += `  • ${cls}\n`;
    });
    output += '\n';
  }

  for (const [className, servants] of Object.entries(data.servantAbsences)) {
    if (servants.length > 0) {
      output += `Absent from ${className}:\n`;
      servants.forEach(servant => {
        output += `  • ${servant}\n`;
      });
      output += '\n';
    }
  }

  return output;
}

module.exports = {
  parseAttendancePdf,
  formatAttendanceData
};
