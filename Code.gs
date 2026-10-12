const SPREADSHEET_ID = '1q7U2KHKFW-DU8ncANE-XXQ-vgdxZDLG2q2uqfxcAuZ4';
const LOG_SHEET = 'Attendance_Log';
const SETTINGS_SHEET = 'Attendance_Settings';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ZIET Mumbai | Workshop Attendance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const excluded = [LOG_SHEET, SETTINGS_SHEET];
  const tabs = ss.getSheets()
    .map(s => s.getName())
    .filter(n => !excluded.includes(n));
  return { spreadsheetName: ss.getName(), tabs: tabs, settings: getAttendanceSettings_() };
}

function getParticipants(sheetName) {
  if (!sheetName) throw new Error('Select a participant list.');
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh || [LOG_SHEET, SETTINGS_SHEET].includes(sheetName)) {
    throw new Error('The selected participant sheet was not found.');
  }

  const values = sh.getDataRange().getDisplayValues();
  if (!values.length) return [];

  // Some participant sheets have a title or merged heading above the column headers.
  // Scan the first 10 rows and use the row that best identifies participant names.
  let headerRowIndex = -1;
  let nameIndex = -1;
  let schoolIndex = -1;
  let bestScore = -1;
  const nameHeaders = [
    'participant name', 'name of participant', 'name of the participant',
    'teacher name', 'name of teacher', 'name of the teacher',
    'teacher', 'participant', 'name', 'full name', 'employee name',
    'name of employee', 'staff name', 'teacher\'s name',
    'प्रशिक्षार्थी का नाम', 'प्रतिभागी का नाम', 'शिक्षक का नाम', 'नाम'
  ];
  const schoolHeaders = [
    'school', 'school name', 'name of school', 'name of the school',
    'kv name', 'name of kv', 'name of the kv', 'kendriya vidyalaya',
    'vidyalaya', 'विद्यालय', 'विद्यालय का नाम', 'केन्द्रीय विद्यालय'
  ];

  const scanLimit = Math.min(values.length, 10);
  for (let r = 0; r < scanLimit; r++) {
    const headers = values[r].map(v => normalizeHeader_(v));
    const n = findHeader_(headers, nameHeaders);
    const s = findHeader_(headers, schoolHeaders);
    let score = n >= 0 ? 10 : 0;
    if (s >= 0) score += 3;
    if (n >= 0 && headers.filter(Boolean).length >= 2) score += 1;
    if (score > bestScore) {
      bestScore = score;
      headerRowIndex = r;
      nameIndex = n;
      schoolIndex = s;
    }
  }

  if (nameIndex < 0) {
    const preview = values.slice(0, Math.min(values.length, 5))
      .map((row, i) => 'Row ' + (i + 1) + ': ' + row.filter(String).join(' | '))
      .join(' ; ');
    throw new Error('Could not identify the participant-name column in the first 10 rows of "' +
      sheetName + '". Check the header row. Sample: ' + preview);
  }

  return values.slice(headerRowIndex + 1).map((row, offset) => ({
    // Actual row number in the sheet, used as a stable participant row identifier.
    id: String(headerRowIndex + offset + 2),
    name: String(row[nameIndex] || '').trim(),
    school: schoolIndex >= 0 ? String(row[schoolIndex] || '').trim() : ''
  })).filter(p => p.name);
}

function normalizeHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[：:]/g, '')
    .replace(/[_-]+/g, ' ')
    .replace(/\\s+/g, ' ')
    .trim();
}

function findHeader_(headers, candidates) {
  const normalizedCandidates = candidates.map(normalizeHeader_);
  for (let i = 0; i < headers.length; i++) {
    const h = normalizeHeader_(headers[i]);
    if (normalizedCandidates.includes(h)) return i;
  }
  // Allow variants such as "Name of the Teacher (as per KVS records)".
  for (let i = 0; i < headers.length; i++) {
    const h = normalizeHeader_(headers[i]);
    if (h && normalizedCandidates.some(c => c.length >= 4 && (h.includes(c) || c.includes(h)))) return i;
  }
  return -1;
}

function getAttendanceSettings_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  const headers = ['Day', 'Session', 'Date (YYYY-MM-DD)', 'Opening time (HH:MM)', 'Closing time (HH:MM)'];
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_SHEET);
    sh.getRange(1, 1, 11, 5).setValues([headers].concat(makeDefaultSettingsRows_()));
    sh.setFrozenRows(1);
  } else {
    const firstRow = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getDisplayValues()[0]
      .map(v => String(v).trim().toLowerCase());
    if (firstRow.indexOf('session') === -1 || firstRow.indexOf('date (yyyy-mm-dd)') === -1) {
      // Upgrade the earlier five-day settings format without losing its configured dates/times.
      const oldRows = sh.getLastRow() > 1
        ? sh.getRange(2, 1, Math.min(5, sh.getLastRow() - 1), Math.min(4, sh.getLastColumn())).getDisplayValues()
        : [];
      const oldByDay = {};
      oldRows.forEach((r, i) => { oldByDay[i + 1] = r; });
      const newRows = makeDefaultSettingsRows_().map((r, i) => {
        const day = Math.floor(i / 2) + 1;
        const old = oldByDay[day] || [];
        return [r[0], r[1], old[1] || '', old[2] || '09:00', old[3] || '17:00'];
      });
      sh.clearContents();
      sh.getRange(1, 1, 11, 5).setValues([headers].concat(newRows));
      sh.setFrozenRows(1);
    } else if (sh.getLastRow() < 11) {
      const existing = sh.getLastRow() > 1
        ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues() : [];
      const defaults = makeDefaultSettingsRows_();
      const byKey = {};
      existing.forEach(r => byKey[String(r[0]) + '|' + String(r[1])] = r);
      const rows = defaults.map(r => byKey[r[0] + '|' + r[1]] || r);
      sh.getRange(1, 1, 11, 5).setValues([headers].concat(rows));
    }
  }
  // Read underlying cell values so display formats like 12.10.2026 remain unchanged.
  const range = sh.getRange(2, 1, 10, 5);
  const rawRows = range.getValues();
  const displayRows = range.getDisplayValues();
  return rawRows.map((r, i) => ({
    day: String(displayRows[i][0] || r[0] || ('Day ' + (Math.floor(i / 2) + 1))).trim(),
    session: String(displayRows[i][1] || r[1] || (i % 2 === 0 ? 'Forenoon' : 'Afternoon')).trim(),
    date: normalizeWorkshopDate_(r[2], displayRows[i][2]),
    open: normalizeTime_(r[3], displayRows[i][3]) || (i % 2 === 0 ? '09:00' : '13:00'),
    close: normalizeTime_(r[4], displayRows[i][4]) || (i % 2 === 0 ? '12:30' : '17:00')
  }));
}

function normalizeWorkshopDate_(rawValue, displayedValue) {
  // Preserve the sheet's visible date format; normalize only inside the app.
  if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
    return Utilities.formatDate(rawValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(displayedValue || rawValue || '').trim();
  if (!s) return '';
  let m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return s;
}

function normalizeTime_(rawValue, displayedValue) {
  // Prefer the visible value when it contains a clear clock time (e.g. 2:00:00 PM).
  // Return a 24-hour HH:mm string for reliable string comparisons.
  const s = String(displayedValue || '').trim();
  const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i);
  if (m) {
    let hour = Number(m[1]), minute = Number(m[2]);
    const ap = (m[4] || '').toUpperCase();
    if (ap === 'PM' && hour < 12) hour += 12;
    if (ap === 'AM' && hour === 12) hour = 0;
    return ('0' + hour).slice(-2) + ':' + ('0' + minute).slice(-2);
  }
  if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
    return Utilities.formatDate(rawValue, 'Asia/Kolkata', 'HH:mm');
  }
  const fallback = String(rawValue || '').trim();
  const mf = fallback.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i);
  if (mf) {
    let hour = Number(mf[1]), minute = Number(mf[2]);
    const ap = (mf[4] || '').toUpperCase();
    if (ap === 'PM' && hour < 12) hour += 12;
    if (ap === 'AM' && hour === 12) hour = 0;
    return ('0' + hour).slice(-2) + ':' + ('0' + minute).slice(-2);
  }
  return fallback;
}

function makeDefaultSettingsRows_() {
  const rows = [];
  for (let day = 1; day <= 5; day++) {
    rows.push(['Day ' + day, 'Forenoon (पूर्वाह्न)', '', '09:00', '12:30']);
    rows.push(['Day ' + day, 'Afternoon (अपराह्न)', '', '13:00', '17:00']);
  }
  return rows;
}

function getCurrentAttendanceSession() {
  const settings = getAttendanceSettings_();
  const now = new Date();
  const tz = 'Asia/Kolkata';
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const currentTime = Utilities.formatDate(now, tz, 'HH:mm');
  const matching = [];
  settings.forEach((s, i) => {
    if (s.date === today && currentTime >= s.open && currentTime <= s.close) {
      matching.push({
        dayNumber: Math.floor(i / 2) + 1,
        day: s.day,
        sessionCode: i % 2 === 0 ? 'FN' : 'AN',
        session: s.session,
        date: s.date,
        open: s.open,
        close: s.close
      });
    }
  });
  if (matching.length) return { available: true, sessions: matching };
  const todaySettings = settings.map((s, i) => ({
    dayNumber: Math.floor(i / 2) + 1, sessionCode: i % 2 === 0 ? 'FN' : 'AN',
    session: s.session, date: s.date, open: s.open, close: s.close
  })).filter(s => s.date === today);
  return {
    available: false,
    sessions: [],
    todaySessions: todaySettings,
    message: todaySettings.length
      ? 'Attendance is not open right now. Today\'s configured session times are shown below.'
      : 'No workshop attendance session is configured for today. Please check Attendance_Settings.'
  };
}

function submitAttendance(payload) {
  if (!payload || !payload.sheetName || !payload.participantId ||
      !payload.dayNumber || !payload.sessionCode || !['Present', 'Absent'].includes(payload.status)) {
    throw new Error('Select your participant and mark Present or Absent for the current session.');
  }
  const dayNum = Number(payload.dayNumber);
  const sessionCode = String(payload.sessionCode);
  if (dayNum < 1 || dayNum > 5 || !Number.isInteger(dayNum) || !['FN', 'AN'].includes(sessionCode)) {
    throw new Error('Invalid day or session.');
  }

  const participantList = getParticipants(payload.sheetName);
  const participant = participantList.find(p => p.id === String(payload.participantId));
  if (!participant) throw new Error('Participant not found. Refresh the list and try again.');

  const settings = getAttendanceSettings_();
  const index = (dayNum - 1) * 2 + (sessionCode === 'FN' ? 0 : 1);
  const setting = settings[index];
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const currentTime = Utilities.formatDate(now, tz, 'HH:mm');
  const sessionLabel = sessionCode === 'FN' ? 'Forenoon (पूर्वाह्न)' : 'Afternoon (अपराह्न)';

  if (!setting || !setting.date) throw new Error('The date for Day ' + dayNum + ' ' + sessionLabel + ' is not configured.');
  if (today !== setting.date) throw new Error('This session is configured for ' + setting.date + ', not today.');
  if (currentTime < setting.open || currentTime > setting.close) {
    throw new Error('Attendance is closed for ' + sessionLabel + '. Permitted time: ' + setting.open + '–' + setting.close + '.');
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let log = ss.getSheetByName(LOG_SHEET);
  const logHeaders = ['Timestamp', 'Participant Sheet', 'Participant Row', 'Participant Name', 'School', 'Day', 'Session', 'Workshop Date', 'Status', 'Submitted At'];
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    log.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
    log.setFrozenRows(1);
  } else if (String(log.getRange(1, 7).getDisplayValue()).trim() !== 'Session') {
    const existing = log.getDataRange().getValues();
    log.clearContents();
    log.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
    const migrated = existing.slice(1).map(r => [r[0], r[1], r[2], r[3], r[4], r[5], 'Legacy (session unspecified)', r[6], r[7], r[8]]);
    if (migrated.length) log.getRange(2, 1, migrated.length, logHeaders.length).setValues(migrated);
    log.setFrozenRows(1);
  }

  const data = log.getDataRange().getValues();
  let existingRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === payload.sheetName &&
        String(data[i][2]) === String(payload.participantId) &&
        String(data[i][5]) === ('Day ' + dayNum) &&
        String(data[i][6]) === sessionLabel &&
        String(data[i][7]) === setting.date) {
      existingRow = i + 1;
      break;
    }
  }
  const row = [now, payload.sheetName, String(payload.participantId), participant.name, participant.school,
    'Day ' + dayNum, sessionLabel, setting.date, payload.status, now];
  if (existingRow > 0) log.getRange(existingRow, 1, 1, row.length).setValues([row]);
  else log.appendRow(row);

  return {
    success: true,
    message: 'Attendance saved: Day ' + dayNum + ' — ' + sessionLabel + ' — ' + payload.status + '.',
    dayNumber: dayNum,
    sessionCode: sessionCode,
    status: payload.status
  };
}
