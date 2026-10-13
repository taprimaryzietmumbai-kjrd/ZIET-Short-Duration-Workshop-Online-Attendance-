const SPREADSHEET_ID = '1q7U2KHKFW-DU8ncANE-XXQ-vgdxZDLG2q2uqfxcAuZ4';
const LOG_SHEET = 'Attendance_Log';
const SETTINGS_SHEET = 'Attendance_Settings';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ZIET Mumbai | Workshop Attendance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
  // Keep the attendance application aligned with the organiser's approved timetable.
  // Dates are ISO (yyyy-MM-dd); times are 24-hour local India time.
  sh.getRange(2, 1, 10, 5).setValues(makeDefaultSettingsRows_());

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
  if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
    return Utilities.formatDate(rawValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(displayedValue || rawValue || '').trim();
  if (!s) return '';
  let m = s.match(/^(\\d{1,2})[.\\/-](\\d{1,2})[.\\/-](\\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = s.match(/^(\\d{4})[.\\/-](\\d{1,2})[.\\/-](\\d{1,2})$/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return s;
}

function normalizeTime_(rawValue, displayedValue) {
  const s = String(displayedValue || '').trim();
  const m = s.match(/^(\\d{1,2}):(\\d{2})(?::(\\d{2}))?\\s*(AM|PM)?$/i);
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
  const mf = fallback.match(/^(\\d{1,2}):(\\d{2})(?::(\\d{2}))?\\s*(AM|PM)?$/i);
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
  const schedule = [
    ['2026-10-12', '2026-10-12'],
    ['2026-10-13', '2026-10-13'],
    ['2026-10-14', '2026-10-14'],
    ['2026-10-15', '2026-10-15'],
    ['2026-10-16', '2026-10-16']
  ];
  const rows = [];
  for (let day = 1; day <= 5; day++) {
    const date = schedule[day - 1][0];
    rows.push(['Day ' + day, 'Forenoon (पूर्वाह्न)', date, '09:00', '11:30']);
    rows.push(['Day ' + day, 'Afternoon (अपराह्न)', date, '14:00', '17:30']);
  }
  return rows;
}

function getCurrentAttendanceSession() {
  const settings = getAttendanceSettings_();
  const now = new Date();
  const tz = 'Asia/Kolkata';
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const currentTime = Utilities.formatDate(now, tz, 'HH:mm');
  const sessions = settings.map((s, i) => ({
    dayNumber: Math.floor(i / 2) + 1,
    day: s.day,
    sessionCode: i % 2 === 0 ? 'FN' : 'AN',
    session: s.session,
    date: s.date,
    open: s.open,
    close: s.close,
    currentlyWithinWindow: s.date === today && currentTime >= s.open && currentTime <= s.close
  })).filter(s => s.date === today);
  if (sessions.length) return { available: true, sessions: sessions, today: today };
  return {
    available: false,
    sessions: [],
    today: today,
    message: 'No workshop attendance session is configured for today. Check that today\'s date is present in Attendance_Settings.'
  };
}

function submitAttendance(payload) {
  if (!payload || !payload.sheetName || !payload.participantId ||
      !payload.dayNumber || !payload.sessionCode || !['Present', 'Absent'].includes(payload.status)) {
    throw new Error('Select your participant, attendance status, and session.');
  }
  if (!payload.signatureDataUrl || !/^data:image\\/(png|jpeg|jpg);base64,/.test(payload.signatureDataUrl)) {
    throw new Error('Please draw or upload your signature before submitting attendance.');
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
  const tz = 'Asia/Kolkata';
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const timeStampText = Utilities.formatDate(now, tz, 'dd.MM.yyyy HH:mm:ss');
  const sessionLabel = sessionCode === 'FN' ? 'Forenoon (पूर्वाह्न)' : 'Afternoon (अपराह्न)';

  if (!setting || !setting.date) throw new Error('The date for Day ' + dayNum + ' ' + sessionLabel + ' is not configured.');
  if (today !== setting.date) throw new Error('This session is configured for ' + setting.date + ', not today.');
  const currentTime = Utilities.formatDate(now, tz, 'HH:mm');
  if (currentTime < setting.open || currentTime > setting.close) {
    throw new Error('Attendance is accepted only between ' + setting.open + ' and ' + setting.close +
      ' IST for Day ' + dayNum + ' ' + sessionLabel + '. Current time: ' + currentTime + ' IST.');
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let log = ss.getSheetByName(LOG_SHEET);
  const logHeaders = ['Timestamp', 'Participant Sheet', 'Participant Row', 'Participant Name', 'School', 'Day', 'Session', 'Workshop Date', 'Status', 'Submitted At', 'Signature Cell'];
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    log.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
    log.setFrozenRows(1);
  } else if (String(log.getRange(1, 7).getDisplayValue()).trim() !== 'Session') {
    const existing = log.getDataRange().getValues();
    log.clearContents();
    log.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
    const migrated = existing.slice(1).map(r => [r[0], r[1], r[2], r[3], r[4], r[5], 'Legacy (session unspecified)', r[6], r[7], r[8], '']);
    if (migrated.length) log.getRange(2, 1, migrated.length, logHeaders.length).setValues(migrated);
    log.setFrozenRows(1);
  } else if (log.getLastColumn() < logHeaders.length) {
    log.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
  }

  const participantSheet = ss.getSheetByName(payload.sheetName);
  const participantRow = Number(payload.participantId);
  if (!participantSheet || participantRow < 1 || participantRow > participantSheet.getMaxRows()) {
    throw new Error('Could not find the participant row in the selected worksheet.');
  }

  const signatureColumn = 4 + index;
  const signatureCell = participantSheet.getRange(participantRow, signatureColumn);
  const signatureCellA1 = signatureCell.getA1Notation();

  const dataUrlParts = payload.signatureDataUrl.split(',');
  const mime = dataUrlParts[0].match(/^data:(image\\/(?:png|jpeg|jpg));base64$/i);
  if (!mime || !dataUrlParts[1]) throw new Error('The signature image could not be read.');
  const imageBytes = Utilities.base64Decode(dataUrlParts[1]);
  if (imageBytes.length > 1500000) throw new Error('Please use a smaller signature image (under 1.5 MB).');
  const blob = Utilities.newBlob(imageBytes, mime[1].toLowerCase().replace('image/jpg','image/jpeg'),
    'signature_' + payload.participantId + '_day' + dayNum + '_' + sessionCode + '.png');

  participantSheet.setColumnWidth(signatureColumn, 110);
  participantSheet.setRowHeight(participantRow, Math.max(participantSheet.getRowHeight(participantRow), 55));
  signatureCell.setNote('Attendance: ' + payload.status + '\\nSession: ' + sessionLabel + '\\nSubmitted: ' + timeStampText);

  participantSheet.getImages().forEach(image => {
    const anchor = image.getAnchorCell();
    if (anchor.getRow() === participantRow && anchor.getColumn() === signatureColumn) image.remove();
  });
  const insertedImage = participantSheet.insertImage(blob, signatureColumn, participantRow);
  insertedImage.setWidth(100).setHeight(45).setAnchorCell(signatureCell);

  const data = log.getDataRange().getValues();
  let existingRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === payload.sheetName &&
        String(data[i][2]) === String(payload.participantId) &&
        String(data[i][5]) === ('Day ' + dayNum) &&
        String(data[i][6]) === sessionLabel &&
        normalizeWorkshopDate_(data[i][7], data[i][7] instanceof Date
          ? Utilities.formatDate(data[i][7], tz, 'yyyy-MM-dd')
          : String(data[i][7])) === setting.date) {
      existingRow = i + 1;
      break;
    }
  }
  const row = [now, payload.sheetName, String(payload.participantId), participant.name, participant.school,
    'Day ' + dayNum, sessionLabel, setting.date, payload.status, timeStampText, signatureCellA1];
  if (existingRow > 0) log.getRange(existingRow, 1, 1, row.length).setValues([row]);
  else log.appendRow(row);

  signatureCell.setNote('Attendance: ' + payload.status + '\\nSession: ' + sessionLabel + '\\nSubmitted: ' + timeStampText);

  return {
    success: true,
    message: 'Attendance and signature saved for Day ' + dayNum + ' — ' + sessionLabel + ' at ' + timeStampText + '.',
    submittedAt: timeStampText,
    signatureCell: signatureCellA1
  };
}
