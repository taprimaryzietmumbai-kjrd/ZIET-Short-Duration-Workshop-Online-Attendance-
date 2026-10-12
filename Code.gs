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
  const rows = sh.getRange(2, 1, 10, 5).getDisplayValues();
  return rows.map((r, i) => ({
    day: r[0] || ('Day ' + (Math.floor(i / 2) + 1)),
    session: r[1] || (i % 2 === 0 ? 'Forenoon' : 'Afternoon'),
    date: String(r[2] || '').trim(),
    open: String(r[3] || '').trim() || (i % 2 === 0 ? '09:00' : '13:00'),
    close: String(r[4] || '').trim() || (i % 2 === 0 ? '12:30' : '17:00')
  }));
}

function makeDefaultSettingsRows_() {
  const rows = [];
  for (let day = 1; day <= 5; day++) {
    rows.push(['Day ' + day, 'Forenoon (पूर्वाह्न)', '', '09:00', '12:30']);
    rows.push(['Day ' + day, 'Afternoon (अपराह्न)', '', '13:00', '17:00']);
  }
  return rows;
}

function submitAttendance(payload) {
  if (!payload || !payload.sheetName || !payload.participantId || !payload.statuses) {
    throw new Error('Participant details and session-wise attendance statuses are required.');
  }
  const statuses = payload.statuses;
  const sessionKeys = [];
  for (let day = 1; day <= 5; day++) {
    ['FN', 'AN'].forEach(session => {
      const key = day + '_' + session;
      sessionKeys.push(key);
      if (!['Present', 'Absent'].includes(statuses[key])) {
        throw new Error('Please choose Present or Absent for all forenoon and afternoon sessions.');
      }
    });
  }

  const participantList = getParticipants(payload.sheetName);
  const participant = participantList.find(p => p.id === String(payload.participantId));
  if (!participant) throw new Error('Participant not found. Refresh the list and try again.');

  const settings = getAttendanceSettings_();
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let log = ss.getSheetByName(LOG_SHEET);
  const logHeaders = ['Timestamp', 'Participant Sheet', 'Participant Row', 'Participant Name', 'School', 'Day', 'Session', 'Workshop Date', 'Status', 'Submitted At'];
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    log.appendRow(logHeaders);
    log.setFrozenRows(1);
  } else if (String(log.getRange(1, 7).getDisplayValue()).trim() !== 'Session') {
    // Upgrade the old log layout by adding a Session column while preserving prior records.
    const existing = log.getDataRange().getValues();
    log.clearContents();
    log.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
    const migrated = existing.slice(1).map(r => [r[0], r[1], r[2], r[3], r[4], r[5], 'Legacy (session unspecified)', r[6], r[7], r[8]]);
    if (migrated.length) log.getRange(2, 1, migrated.length, logHeaders.length).setValues(migrated);
    log.setFrozenRows(1);
  }

  const accepted = [];
  const rejected = [];
  settings.forEach((setting, index) => {
    const dayNum = Math.floor(index / 2) + 1;
    const sessionCode = index % 2 === 0 ? 'FN' : 'AN';
    const sessionLabel = index % 2 === 0 ? 'Forenoon (पूर्वाह्न)' : 'Afternoon (अपराह्न)';
    const status = statuses[dayNum + '_' + sessionCode];
    if (!setting.date) {
      rejected.push('Day ' + dayNum + ' ' + sessionLabel + ': date not configured');
      return;
    }
    const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    if (today !== setting.date) {
      rejected.push('Day ' + dayNum + ' ' + sessionLabel + ': attendance available only on ' + setting.date);
      return;
    }
    const currentTime = Utilities.formatDate(now, tz, 'HH:mm');
    if (currentTime < setting.open || currentTime > setting.close) {
      rejected.push('Day ' + dayNum + ' ' + sessionLabel + ': window is ' + setting.open + '–' + setting.close);
      return;
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
      'Day ' + dayNum, sessionLabel, setting.date, status, now];
    if (existingRow > 0) log.getRange(existingRow, 1, 1, row.length).setValues([row]);
    else log.appendRow(row);
    accepted.push('Day ' + dayNum + ' ' + sessionLabel);
  });

  if (!accepted.length) {
    throw new Error('No attendance was saved. ' + rejected.join('; ') +
      '. Configure dates and time windows in the "' + SETTINGS_SHEET + '" tab.');
  }
  return {
    success: true,
    message: 'Attendance saved for: ' + accepted.join(', ') +
      (rejected.length ? '. Not saved: ' + rejected.join('; ') : '.'),
    accepted: accepted,
    rejected: rejected
  };
}
