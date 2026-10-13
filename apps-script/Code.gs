/**
 * St. Abraam Sunday School — Weekly Servant Attendance
 *
 * Runs inside Google Apps Script on the stabraammain@gmail.com account.
 * Every hour it looks for a new "Scheduled Report - Sunday School Attendance"
 * email from Church Teams. When one arrives it:
 *   1. saves the PDF to the "Church Teams Attendance" Drive folder,
 *   2. reads who was marked present (X) in each group,
 *   3. checks every servant listed on the GitHub servants page,
 *   4. works out consecutive absences from the History tab,
 *   5. emails the summary and logs the week to the History tab.
 *
 * Functions you can run by hand from the Apps Script editor:
 *   setup()             – creates the sheet tabs and the hourly trigger (run once)
 *   runLatestReport()   – re-processes the most recent report and re-sends the email
 *   checkForNewReports() – what the hourly trigger runs
 */

const CONFIG = {
  EMAIL_TO: 'christophermetri@gmail.com, amirloka@gmail.com', // weekly summary (comma-separated)
  ERROR_EMAIL_TO: 'christophermetri@gmail.com',                // problems with the automation
  SENDER_NAME: 'St. Abraam Sunday School',
  SHEET_ID: '1CgfXd2B_dlAMDBqBkhzzU9jSCYh7mxFpXxcUS97u4q4',
  PDF_FOLDER_ID: '1Un-scNhfDC_QRVKDb9DvD49rEB1TkEwY',
  ROSTER_URL: 'https://raw.githubusercontent.com/christophermetri/sunday-school-assignments/main/index.html',
  GMAIL_QUERY: 'from:groupfinder@churchteams.com subject:"Sunday School Attendance" has:attachment',
  PROCESSED_LABEL: 'Attendance Processed',
  MAROON: '#8B2323',
  GOLD: '#DAA520'
};

// Website class name -> Church Teams group name(s). Order here is the order in the email.
const CLASS_GROUPS = [
  ['Nursery', ['Nursery']],
  ['PreK', ['Pre-K']],
  ['Kindergarten', ['KG']],
  ['1st Grade', ['1st Grade']],
  ['2nd/3rd Grade', ['2nd Grade', '3rd Grade']],
  ['4th Grade', ['4th Grade']],
  ['5th Grade', ['5th Grade']],
  ['6th Grade', ['6th Grade']],
  ['7th Grade', ['7th Grade']],
  ['8th Grade', ['8th Grade']],
  ['Preservants 1', ['Pre-servants Group 1']],
  ['Preservants 2', ['Pre-servants Group 2']],
  ['Preservants 3', ['Pre-servants Group 3']]
];

const TAB_HISTORY = 'History';
const TAB_FIXES = 'Name Fixes';
const TAB_SERVANTS = 'Servants';

/* ============================== Entry points ============================== */

function setup() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  ensureTab_(ss, TAB_HISTORY, ['Date', 'Class', 'Servant', 'Status', 'Church Teams name']);
  ensureTab_(ss, TAB_FIXES, ['Servant (as on website)', 'Church Teams name (Last, First) — or SKIP']);
  ensureTab_(ss, TAB_SERVANTS, ['Class', 'Servant']);
  seedNameFixes_(ss);
  loadRoster_(); // fills the Servants tab

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkForNewReports')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkForNewReports').timeBased().everyHours(1).create();

  // Log any reports already in the inbox to History (no email), so streaks start counting from them.
  const label = getOrCreateLabel_();
  const threads = unprocessedThreads_();
  threads.forEach(t => t.getMessages().forEach(m => processMessage_(m, { sendEmail: false })));
  threads.forEach(t => t.addLabel(label));
  Logger.log('Setup complete. Hourly trigger installed. Back-filled ' + threads.length + ' report(s).');
}

/** Church Teams report threads from the last 30 days that haven't been processed yet. */
function unprocessedThreads_() {
  return GmailApp.search(CONFIG.GMAIL_QUERY + ' newer_than:30d')
    .filter(t => !t.getLabels().some(l => l.getName() === CONFIG.PROCESSED_LABEL));
}

function checkForNewReports() {
  const label = getOrCreateLabel_();
  const threads = unprocessedThreads_();
  const messages = [];
  threads.forEach(t => t.getMessages().forEach(m => messages.push(m)));
  messages.sort((a, b) => a.getDate() - b.getDate());
  messages.forEach(m => {
    try {
      processMessage_(m);
    } catch (err) {
      notifyError_(err, m);
    }
  });
  threads.forEach(t => t.addLabel(label));
}

function runLatestReport() {
  const threads = GmailApp.search(CONFIG.GMAIL_QUERY, 0, 5);
  if (!threads.length) throw new Error('No Church Teams attendance emails found.');
  let latest = null;
  threads.forEach(t => t.getMessages().forEach(m => { if (!latest || m.getDate() > latest.getDate()) latest = m; }));
  processMessage_(latest);
  latest.getThread().addLabel(getOrCreateLabel_());
}

/* ============================== Main pipeline ============================== */

function processMessage_(message, options) {
  const sendEmail = !(options && options.sendEmail === false);
  const body = message.getPlainBody();
  const reportDate = parseReportDate(body) || message.getDate();
  const dateKey = Utilities.formatDate(reportDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const requestedGroups = parseRequestedGroups(body);

  const pdf = message.getAttachments().filter(a => /pdf/i.test(a.getContentType()) || /\.pdf$/i.test(a.getName()))[0];
  if (!pdf) throw new Error('The Church Teams email has no PDF attached.');
  savePdf_(pdf, dateKey);

  const lines = pdfToLines_(pdf);
  const report = parseReport(lines);
  if (!Object.keys(report).length) throw new Error('Could not read any groups from the attendance PDF.');

  const roster = loadRoster_();
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const fixes = loadNameFixes_(ss);
  const result = analyze(report, requestedGroups, roster, fixes);

  const historySheet = ss.getSheetByName(TAB_HISTORY) || ensureTab_(ss, TAB_HISTORY, ['Date', 'Class', 'Servant', 'Status', 'Church Teams name']);
  const history = readHistory_(historySheet, dateKey);
  addStreaks(result, history, dateKey);
  writeHistory_(historySheet, dateKey, result);

  if (!sendEmail) return;
  const html = buildEmailHtml(result, reportDate, 'https://docs.google.com/spreadsheets/d/' + CONFIG.SHEET_ID + '/edit');
  GmailApp.sendEmail(CONFIG.EMAIL_TO, 'Attendance for ' + formatLongDate_(reportDate), buildEmailText(result, reportDate), {
    htmlBody: html,
    name: CONFIG.SENDER_NAME
  });
}

/* ============================== Pure logic (tested) ============================== */

function parseReportDate(body) {
  const m = /(\d{1,2})\/(\d{1,2})\/(\d{4})\s+to/i.exec(body || '');
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[1]) - 1, Number(m[2]), 12);
}

/**
 * Groups listed in the email ("Groups: '6th Grade ','3rd Grade',..."), plus every group the
 * website classes map to — so a class that didn't take attendance is always reported.
 */
function parseRequestedGroups(body) {
  const text = String(body || '').replace(/&#39;|&apos;|&#8216;|&#8217;/g, "'").replace(/[‘’]/g, "'");
  // The groups line is often wrapped across several lines in the plain-text body
  const m = /Groups:\s*([\s\S]*?)(?:\n\s*\n|$)/i.exec(text);
  const line = m ? m[1].replace(/\s*[\r\n]+\s*/g, ' ') : '';
  const found = line.split(/'\s*,\s*'|,/).map(s => cleanGroup(s.replace(/['"]/g, ''))).filter(Boolean);
  const all = found.slice();
  CLASS_GROUPS.forEach(([, groups]) => groups.forEach(g => {
    if (!all.some(x => groupKey(x) === groupKey(g))) all.push(g);
  }));
  return all;
}

function cleanGroup(name) {
  return String(name || '').replace(/\s+/g, ' ').trim();
}

function groupKey(name) {
  return cleanGroup(name).toLowerCase();
}

/**
 * lines: text lines in reading order. Table rows arrive as "cell | cell | cell".
 * Returns { groupKey: { name, members: [{ name, present }], presentCount } }
 */
function parseReport(lines) {
  const groups = {};
  let current = null;
  const HEADER = /^([^,|]+?)\s*\|?\s*\d{1,2}\/\d{1,2}\s*-\s*\d{1,2}\/\d{1,2}/;
  lines.forEach(raw => {
    String(raw).split(/[\n\r\u000b]+/).forEach(piece => {
      const line = piece.replace(/ /g, ' ').trim();
      if (!line) return;
      const h = HEADER.exec(line);
      if (h) {
        const name = cleanGroup(h[1]);
        const key = groupKey(name);
        current = groups[key] || (groups[key] = { name: name, members: [], presentCount: 0 });
        return;
      }
      if (!current) return;
      if (/group members|overall percentage|%/i.test(line)) return;
      let name, present;
      if (line.indexOf('|') !== -1) {
        const cells = line.split('|').map(c => c.trim());
        name = cells[0];
        present = cells.slice(1).some(c => /^x$/i.test(c));
      } else {
        const m = /^(.*?)(?:\s+(X))?$/.exec(line);
        name = m[1];
        present = !!m[2];
      }
      if (name.indexOf(',') === -1) return;
      current.members.push({ name: name.replace(/\s+/g, ' ').trim(), present: present });
      if (present) current.presentCount++;
    });
  });
  return groups;
}

/** Website roster: { className: [servant, ...] } for the classes in CLASS_GROUPS. */
function parseRoster(html) {
  const cut = html.search(/Hymns Classes/i);
  if (cut > 0) html = html.slice(0, cut);
  const wanted = {};
  CLASS_GROUPS.forEach(([cls]) => { wanted[cls.toLowerCase()] = cls; });
  const roster = {};
  html.split(/<div class="card">/).slice(1).forEach(card => {
    const h = /<h3>([\s\S]*?)<\/h3>/.exec(card);
    if (!h) return;
    const cls = wanted[decodeHtml(h[1]).toLowerCase()];
    if (!cls) return;
    const names = [];
    const re = /<div class="servant(?: [^"]*)?">([\s\S]*?)<\/div>/g;
    let m;
    while ((m = re.exec(card))) {
      const n = decodeHtml(m[1]);
      if (n && !/^others$/i.test(n)) names.push(n);
    }
    roster[cls] = names;
  });
  return roster;
}

function decodeHtml(s) {
  return String(s).replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ').replace(/&#39;|&apos;/g, "'")
    .replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ').trim();
}

function normName(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/\(.*?\)/g, ' ').replace(/[^a-z,]+/g, ' ').replace(/\s*,\s*/g, ', ').replace(/\s+/g, ' ').trim();
}

function lev(a, b) {
  const d = [];
  for (let i = 0; i <= a.length; i++) d[i] = [i];
  for (let j = 0; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
  }
  return d[a.length][b.length];
}

function similarWord(a, b) {
  if (!a || !b) return false;
  if (a === b) return true;
  const n = Math.min(a.length, b.length);
  if (n >= 4 && lev(a, b) <= 1) return true;
  if (n >= 6 && lev(a, b) <= 2) return true;
  return false;
}

function similarFirst(a, b) {
  if (similarWord(a, b)) return true;
  if (!a || !b) return false;
  let p = 0;
  while (p < a.length && p < b.length && a[p] === b[p]) p++;
  if (p >= 4) return true;                                   // Raphy / Raphael
  if (Math.min(a.length, b.length) >= 4 && (a.endsWith(b) || b.endsWith(a))) return true; // Tony / Antony
  return false;
}

function splitNames_(siteName, ctName) {
  const site = normName(siteName).replace(/,/g, ' ').split(' ').filter(Boolean);
  const parts = normName(ctName).split(',');
  if (site.length < 2 || parts.length < 2) return null;
  return {
    siteFirst: site[0],
    siteRest: site.slice(1),
    ctLast: parts[0].trim().split(' ').filter(Boolean),
    ctFirst: parts.slice(1).join(' ').trim().split(' ').filter(Boolean)
  };
}

/**
 * How well website name "First [Middle] Last" matches Church Teams "Last, First [Middle]".
 * 0 = no match, 1 = close spelling, 2 = exact first and last name.
 */
function matchScore(siteName, ctName) {
  const p = splitNames_(siteName, ctName);
  if (!p || !p.ctFirst.length) return 0;
  const lastExact = p.ctLast.some(l => p.siteRest.indexOf(l) !== -1);
  const lastOk = lastExact || p.ctLast.some(l => p.siteRest.some(r => similarWord(l, r)));
  const firstExact = p.siteFirst === p.ctFirst[0];
  const firstOk = firstExact || similarFirst(p.siteFirst, p.ctFirst[0]);
  if (!lastOk || !firstOk) return 0;
  return lastExact && firstExact ? 2 : 1;
}

function namesMatch(siteName, ctName) {
  return matchScore(siteName, ctName) > 0;
}

/** Church Teams names that share a first or last name — shown as hints for unmatched servants. */
function suggestCtNames(siteName, members) {
  const site = normName(siteName).replace(/,/g, ' ').split(' ').filter(Boolean);
  if (!site.length) return [];
  const out = {};
  members.forEach(m => {
    const parts = normName(m.name).split(',');
    if (parts.length < 2) return;
    const last = parts[0].trim().split(' ');
    const first = parts[1].trim().split(' ')[0];
    if (similarFirst(site[0], first) || last.some(l => site.slice(1).some(r => similarWord(l, r)))) out[normName(m.name)] = m.name;
  });
  return Object.keys(out).map(k => out[k]).slice(0, 3);
}

function sameCtName(a, b) {
  return normName(a) === normName(b);
}

/**
 * Returns {
 *   notTaken: [names of classes / groups with no attendance taken],
 *   classes: [{ cls, servants: [{ servant, status: 'Absent'|'Present'|'Not found', ctName }] }]
 * }
 */
function analyze(report, requestedGroups, roster, fixes) {
  const taken = key => report[key] && report[key].presentCount > 0;
  const requested = {};
  requestedGroups.forEach(g => { requested[groupKey(g)] = cleanGroup(g); });
  const allMembers = [];
  Object.keys(report).forEach(k => report[k].members.forEach(m => allMembers.push(Object.assign({ group: k }, m))));
  const presentAnywhere = ct => allMembers.some(m => m.present && sameCtName(m.name, ct));

  const notTaken = [];
  const classes = [];
  const mapped = {};

  CLASS_GROUPS.forEach(([cls, groupNames]) => {
    const keys = groupNames.map(groupKey);
    keys.forEach(k => { mapped[k] = true; });
    const takenKeys = keys.filter(taken);
    if (!takenKeys.length) {
      if (keys.some(k => requested[k] || report[k])) notTaken.push(cls);
      return;
    }
    keys.filter(k => !taken(k) && (requested[k] || report[k]))
      .forEach(k => notTaken.push(requested[k] || (report[k] && report[k].name) || k));

    const servants = (roster[cls] || []).map(servant => {
      const fix = lookupFix(fixes, servant);
      if (fix && /^skip$/i.test(fix)) return null;
      let ctName = null;
      if (fix) {
        const hit = allMembers.find(m => sameCtName(m.name, fix));
        ctName = hit ? hit.name : null;
      } else {
        ctName = findCtName(servant, allMembers.filter(m => takenKeys.indexOf(m.group) !== -1))
          || findCtName(servant, allMembers);
      }
      if (!ctName) return { servant: servant, status: 'Not found', ctName: '' };
      return { servant: servant, status: presentAnywhere(ctName) ? 'Present' : 'Absent', ctName: ctName };
    }).filter(Boolean);
    const used = servants.map(s => normName(s.ctName)).filter(Boolean);
    const classMembers = allMembers.filter(m => takenKeys.indexOf(m.group) !== -1 && used.indexOf(normName(m.name)) === -1);
    servants.filter(s => s.status === 'Not found').forEach(s => { s.suggestions = suggestCtNames(s.servant, classMembers); });
    classes.push({ cls: cls, servants: servants });
  });

  // Groups Church Teams sent that aren't mapped to a website class
  Object.keys(requested).forEach(k => { if (!mapped[k] && !taken(k)) notTaken.push(requested[k]); });
  return { notTaken: notTaken, classes: classes };
}

function lookupFix(fixes, servant) {
  const k = normName(servant);
  for (const name in fixes) if (normName(name) === k) return fixes[name];
  return null;
}

function findCtName(servant, members) {
  const hits = {};
  let best = 0;
  members.forEach(m => {
    const s = matchScore(servant, m.name);
    if (!s) return;
    if (s > best) { best = s; Object.keys(hits).forEach(k => delete hits[k]); }
    if (s === best) hits[normName(m.name)] = m.name;
  });
  const keys = Object.keys(hits);
  return keys.length === 1 ? hits[keys[0]] : null;
}

/** history: rows [{date:'yyyy-MM-dd', cls, servant, status}] from previous weeks. */
function addStreaks(result, history, dateKey) {
  result.classes.forEach(c => c.servants.forEach(s => {
    if (s.status !== 'Absent') return;
    const past = history.filter(h => h.date < dateKey && h.cls === c.cls && normName(h.servant) === normName(s.servant))
      .sort((a, b) => (a.date < b.date ? 1 : -1));
    let streak = 1;
    for (const h of past) {
      if (h.status === 'Absent') streak++;
      else break;
    }
    s.streak = streak;
  }));
}

function ordinal(n) {
  const j = n % 10, k = n % 100;
  if (j === 1 && k !== 11) return n + 'st';
  if (j === 2 && k !== 12) return n + 'nd';
  if (j === 3 && k !== 13) return n + 'rd';
  return n + 'th';
}

function streakNote(streak) {
  if (!streak || streak < 2) return '';
  if (streak >= 4) return "(hasn't attended in " + streak + ' weeks)';
  return '(' + ordinal(streak) + ' week in a row)';
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function formatLongDate_(d) {
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  return months[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
}

function buildEmailHtml(result, reportDate, sheetUrl) {
  // Built with tables, explicit backgrounds and a "light only" color scheme so phone mail apps
  // in dark mode (e.g. iPhone Mail) show the real maroon/gold design instead of recoloring it.
  const M = CONFIG.MAROON, G = CONFIG.GOLD, TEXT = '#222222', CARD = '#ffffff', PAGE = '#f4f1ea';
  const section = (title, items) =>
    '<tr><td style="padding:20px 24px 0 24px;background-color:' + CARD + ';">' +
    '<div style="font-size:16px;font-weight:bold;color:' + M + ';border-bottom:2px solid ' + G + ';padding-bottom:5px;">' + title + '</div>' +
    '<ul style="margin:10px 0 0 0;padding-left:22px;color:' + TEXT + ';">' +
    items.map(i => '<li style="margin:4px 0;color:' + TEXT + ';">' + i + '</li>').join('') + '</ul></td></tr>';

  let rows = '';
  if (result.notTaken.length) rows += section("Classes that didn't take attendance:", result.notTaken.map(escapeHtml));
  result.classes.forEach(c => {
    const absent = c.servants.filter(s => s.status === 'Absent');
    if (!absent.length) {
      rows += section(escapeHtml(c.cls) + ':', ['All servants attended']);
      return;
    }
    rows += section('Absent from ' + escapeHtml(c.cls) + ':', absent.map(s => {
      const note = streakNote(s.streak);
      return escapeHtml(s.servant) + (note ? ' <span style="color:' + M + ';font-weight:bold;">' + escapeHtml(note) + '</span>' : '');
    }));
  });

  const missing = [];
  result.classes.forEach(c => c.servants.filter(s => s.status === 'Not found').forEach(s => missing.push(
    escapeHtml(s.servant) + ' — ' + escapeHtml(c.cls) +
    (s.suggestions && s.suggestions.length ? ' <span style="color:#777777;">(Church Teams has: ' + s.suggestions.map(escapeHtml).join('; ') + ')</span>' : ''))));
  if (missing.length) {
    rows += '<tr><td style="padding:24px 24px 0 24px;background-color:' + CARD + ';">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
      '<td bgcolor="#f7f3e8" style="background-color:#f7f3e8;border-left:4px solid ' + G + ';padding:10px 14px;font-size:13px;color:#555555;">' +
      '<b style="color:#333333;">Names to check</b> (not counted above): these servants are on the website but couldn\'t be found on their class\'s Church Teams list. ' +
      'Add their Church Teams spelling on the <a href="' + sheetUrl + '" style="color:' + M + ';">Name Fixes tab</a>, or write SKIP if they don\'t check in.' +
      '<ul style="margin:8px 0 0 0;padding-left:20px;color:#555555;">' + missing.map(i => '<li style="margin:3px 0;">' + i + '</li>').join('') + '</ul>' +
      '</td></tr></table></td></tr>';
  }

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<meta name="color-scheme" content="light only"><meta name="supported-color-schemes" content="light only">' +
    '<style>:root{color-scheme:light only;supported-color-schemes:light only;}</style></head>' +
    '<body style="margin:0;padding:0;background-color:' + PAGE + ';">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="' + PAGE + '" style="background-color:' + PAGE + ';">' +
    '<tr><td align="center" style="padding:16px 8px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="' + CARD + '" ' +
    'style="max-width:620px;background-color:' + CARD + ';font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.4;color:' + TEXT + ';">' +
    '<tr><td bgcolor="' + M + '" style="background-color:' + M + ';padding:18px 24px;border-bottom:4px solid ' + G + ';">' +
    '<div style="font-size:12px;letter-spacing:1.5px;font-weight:bold;color:' + G + ';text-transform:uppercase;">St. Abraam Sunday School</div>' +
    '<div style="font-size:22px;font-weight:bold;color:#ffffff;margin-top:4px;">Attendance for ' + formatLongDate_(reportDate) + '</div>' +
    '</td></tr>' + rows +
    '<tr><td style="padding:24px;background-color:' + CARD + ';"></td></tr>' +
    '</table></td></tr></table></body></html>';
}

function buildEmailText(result, reportDate) {
  const lines = ['Attendance for ' + formatLongDate_(reportDate), ''];
  if (result.notTaken.length) {
    lines.push("Classes that didn't take attendance:");
    result.notTaken.forEach(n => lines.push('  * ' + n));
    lines.push('');
  }
  result.classes.forEach(c => {
    const absent = c.servants.filter(s => s.status === 'Absent');
    if (!absent.length) {
      lines.push(c.cls + ':', '  * All servants attended', '');
      return;
    }
    lines.push('Absent from ' + c.cls + ':');
    absent.forEach(s => { const n = streakNote(s.streak); lines.push('  * ' + s.servant + (n ? ' ' + n : '')); });
    lines.push('');
  });
  return lines.join('\n');
}

/* ============================== Google services ============================== */

function getOrCreateLabel_() {
  return GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL) || GmailApp.createLabel(CONFIG.PROCESSED_LABEL);
}

function savePdf_(blob, dateKey) {
  const folder = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);
  const name = 'Attendance ' + dateKey + '.pdf';
  if (!folder.getFilesByName(name).hasNext()) folder.createFile(blob.copyBlob().setName(name));
}

/** Converts the PDF to a temporary Google Doc and returns its text lines in reading order. */
function pdfToLines_(pdfBlob) {
  const file = Drive.Files.create(
    { name: 'tmp-attendance-' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
    pdfBlob.copyBlob().setContentType('application/pdf')
  );
  try {
    const body = DocumentApp.openById(file.id).getBody();
    const lines = [];
    for (let i = 0; i < body.getNumChildren(); i++) {
      const el = body.getChild(i);
      const type = el.getType();
      if (type === DocumentApp.ElementType.TABLE) {
        const table = el.asTable();
        for (let r = 0; r < table.getNumRows(); r++) {
          const row = table.getRow(r);
          const cells = [];
          for (let c = 0; c < row.getNumCells(); c++) cells.push(row.getCell(c).getText().replace(/\s+/g, ' ').trim());
          lines.push(cells.join(' | '));
        }
      } else if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.LIST_ITEM) {
        lines.push(el.asText().getText());
      }
    }
    return lines;
  } finally {
    DriveApp.getFileById(file.id).setTrashed(true);
  }
}

function loadRoster_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(TAB_SERVANTS) || ensureTab_(ss, TAB_SERVANTS, ['Class', 'Servant']);
  try {
    const html = UrlFetchApp.fetch(CONFIG.ROSTER_URL, { muteHttpExceptions: false }).getContentText();
    const roster = parseRoster(html);
    if (Object.keys(roster).length < 5) throw new Error('Roster looked incomplete');
    const rows = [];
    CLASS_GROUPS.forEach(([cls]) => (roster[cls] || []).forEach(s => rows.push([cls, s])));
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
    if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    return roster;
  } catch (err) {
    // Fall back to the last copy saved in the Servants tab
    const roster = {};
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(([cls, s]) => {
        if (cls && s) (roster[cls] = roster[cls] || []).push(String(s));
      });
    }
    if (!Object.keys(roster).length) throw err;
    return roster;
  }
}

function loadNameFixes_(ss) {
  const sheet = ss.getSheetByName(TAB_FIXES);
  const fixes = {};
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(([site, ct]) => {
      if (String(site).trim() && String(ct).trim()) fixes[String(site).trim()] = String(ct).trim();
    });
  }
  return fixes;
}

function readHistory_(sheet, dateKey) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getDisplayValues()
    .filter(r => r[0] && r[0] !== dateKey)
    .map(r => ({ date: r[0], cls: r[1], servant: r[2], status: r[3] }));
}

function writeHistory_(sheet, dateKey, result) {
  // Remove rows for this date (so re-running a week doesn't double count)
  if (sheet.getLastRow() > 1) {
    const dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    for (let i = dates.length - 1; i >= 0; i--) if (dates[i][0] === dateKey) sheet.deleteRow(i + 2);
  }
  const rows = [];
  result.classes.forEach(c => c.servants.forEach(s => rows.push([dateKey, c.cls, s.servant, s.status, s.ctName])));
  if (!rows.length) return;
  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, 1).setNumberFormat('@');
  sheet.getRange(start, 1, rows.length, 5).setValues(rows);
}

function ensureTab_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold')
      .setBackground(CONFIG.MAROON).setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Known spelling differences between the website and Church Teams (from the 10/11/2026 report).
// Confirmed by Chris on 10/13/2026.
const SEED_FIXES = [
  ['Nermin Mouris', 'Abdelsayed, Nermin'],
  ['Sana Hosny', 'Gerges, Sanaa'],
  ['Amira Beshay', 'Gerges, Amira'],
  ['Joly Fawzy', 'Hanna, Joly'],
  ['Kerolous Haded', 'Farag, Kerolous'],
  ['Joanne Roufail', 'Habib, Joann'],
  ['Maro Fouad', 'Fouad, Marquilla'],
  ['Kevin Boktor', 'Botros, Kevin'],
  ['Nardeen Loka', 'Mickail, Nardeen']
];

function seedNameFixes_(ss) {
  const sheet = ss.getSheetByName(TAB_FIXES);
  if (!SEED_FIXES.length || sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, SEED_FIXES.length, 2).setValues(SEED_FIXES);
}

function notifyError_(err, message) {
  GmailApp.sendEmail(CONFIG.ERROR_EMAIL_TO, 'Sunday School attendance automation hit a problem',
    'The attendance email for "' + (message ? message.getSubject() + ' (' + message.getDate() + ')' : 'unknown') +
    '" could not be processed.\n\nError: ' + (err && err.stack || err) +
    '\n\nYou can re-run it from the Apps Script editor with runLatestReport().', { name: CONFIG.SENDER_NAME });
}
