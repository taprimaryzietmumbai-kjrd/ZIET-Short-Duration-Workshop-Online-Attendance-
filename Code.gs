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
  if (values.length < 2) return [];
  const headers = values[0].map(v => String(v).trim().toLowerCase());
  const nameIndex = findHeader_(headers, ['participant name','name of participant','teacher name','name','participant']);
  const schoolIndex = findHeader_(headers, ['school','school name','kv name','name of kv','kendriya vidyalaya','vidyalaya']);
  if (nameIndex < 0) throw new Error('Could not identify the participant-name column in row 1 of "' + sheetName + '".');
  return values.slice(1).map((row, i) => ({
    id: String(i + 2),
    name: String(row[nameIndex] || '').trim(),
    school: schoolIndex >= 0 ? String(row[schoolIndex] || '').trim() : ''
  })).filter(p => p.name);
}

function findHeader_(headers, candidates) {
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i].replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim();
    if (candidates.includes(h)) return i;
  }
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i].replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim();
    if (candidates.some(c => h.includes(c))) return i;
  }
  return -1;
}

function getAttendanceSettings_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_SHEET);
    sh.getRange(1, 1, 6, 4).setValues([
      ['Day', 'Date (YYYY-MM-DD)', 'Opening time (HH:MM)', 'Closing time (HH:MM)'],
      ['Day 1', '', '09:00', '17:00'],
      ['Day 2', '', '09:00', '17:00'],
      ['Day 3', '', '09:00', '17:00'],
      ['Day 4', '', '09:00', '17:00'],
      ['Day 5', '', '09:00', '17:00']
    ]);
    sh.setFrozenRows(1);
  }
  const rows = sh.getRange(2, 1, 5, 4).getDisplayValues();
  return rows.map((r, i) => ({
    day: r[0] || ('Day ' + (i + 1)),
    date: r[1].trim(),
    open: r[2].trim() || '09:00',
    close: r[3].trim() || '17:00'
  }));
}

function submitAttendance(payload) {
  if (!payload || !payload.sheetName || !payload.participantId || !payload.statuses) {
    throw new Error('Participant details and attendance statuses are required.');
  }
  const statuses = payload.statuses;
  for (let d = 1; d <= 5; d++) {
    if (!['Present', 'Absent'].includes(statuses[String(d)])) {
      throw new Error('Please select Present or Absent for all five days.');
    }
  }

  const participantList = getParticipants(payload.sheetName);
  const participant = participantList.find(p => p.id === String(payload.participantId));
  if (!participant) throw new Error('Participant not found. Refresh the list and try again.');

  const settings = getAttendanceSettings_();
  const now = new Date();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    log.appendRow(['Timestamp', 'Participant Sheet', 'Participant Row', 'Participant Name', 'School', 'Day', 'Workshop Date', 'Status', 'Submitted At']);
    log.setFrozenRows(1);
  }

  const accepted = [];
  const rejected = [];
  settings.forEach((setting, index) => {
    const dayNum = index + 1;
    const status = statuses[String(dayNum)];
    if (!setting.date) {
      rejected.push('Day ' + dayNum + ': date not configured');
      return;
    }
    const dateParts = setting.date.split('-').map(Number);
    if (dateParts.length !== 3 || !dateParts[0] || !dateParts[1] || !dateParts[2]) {
      rejected.push('Day ' + dayNum + ': invalid date setting');
      return;
    }
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (today !== setting.date) {
      rejected.push('Day ' + dayNum + ': attendance is available only on ' + setting.date);
      return;
    }
    const currentTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');
    if (currentTime < setting.open || currentTime > setting.close) {
      rejected.push('Day ' + dayNum + ': attendance window is ' + setting.open + '–' + setting.close);
      return;
    }

    const data = log.getDataRange().getValues();
    let existingRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === payload.sheetName &&
          String(data[i][2]) === String(payload.participantId) &&
          String(data[i][5]) === ('Day ' + dayNum) &&
          String(data[i][6]) === setting.date) {
        existingRow = i + 1;
        break;
      }
    }
    const row = [now, payload.sheetName, String(payload.participantId), participant.name, participant.school,
      'Day ' + dayNum, setting.date, status, now];
    if (existingRow > 0) log.getRange(existingRow, 1, 1, row.length).setValues([row]);
    else log.appendRow(row);
    accepted.push('Day ' + dayNum);
  });

  if (!accepted.length) {
    throw new Error('No attendance was saved. ' + rejected.join('; ') +
      ' Configure dates and time windows in the "' + SETTINGS_SHEET + '" tab.');
  }
  return {
    success: true,
    message: 'Attendance saved for ' + accepted.join(', ') +
      (rejected.length ? '. Not saved: ' + rejected.join('; ') : '.'),
    accepted: accepted,
    rejected: rejected
  };
}
