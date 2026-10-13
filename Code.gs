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
  const timetableSheet = findTimetableSheet_(ss);
  const excluded = [LOG_SHEET, SETTINGS_SHEET];
  if (timetableSheet) excluded.push(timetableSheet.getName());
  const tabs = ss.getSheets()
    .map(s => s.getName())
    .filter(n => !excluded.includes(n));
  return {
    spreadsheetName: ss.getName(),
    tabs: tabs,
    settings: getAttendanceSettings_()
  };
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
    .replace(/\s+/g, ' ')
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

function findTimetableSheet_(ss) {
  const excluded = [LOG_SHEET];
  return ss.getSheets().find(sh => {
    if (excluded.includes(sh.getName())) return false;
    const lastRow = Math.min(Math.max(sh.getLastRow(), 1), 10);
    const lastCol = Math.min(Math.max(sh.getLastColumn(), 1), 10);
    const rows = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    return rows.some(row => {
      const headers = row.map(normalizeHeader_);
      const hasDay = headers.some(h => h === 'day' || /^day label$/.test(h));
      const hasSession = headers.some(h => h === 'session');
      const hasDate = headers.some(h => h.indexOf('date') >= 0);
      const hasOpen = headers.some(h => h.indexOf('opening time') >= 0 || h === 'open' || h === 'opening');
      const hasClose = headers.some(h => h.indexOf('closing time') >= 0 || h === 'close' || h === 'closing');
      return hasDay && hasSession && hasDate && hasOpen && hasClose;
    });
  }) || null;
}

function getAttendanceSettings_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = findTimetableSheet_(ss);
  let headerRow = -1;
  let headerMap = null;

  if (sh) {
    const rowsToScan = Math.min(sh.getLastRow(), 10);
    const colsToScan = Math.min(Math.max(sh.getLastColumn(), 1), 10);
    const rows = sh.getRange(1, 1, Math.max(rowsToScan, 1), colsToScan).getDisplayValues();
    for (let r = 0; r < rows.length; r++) {
      const headers = rows[r].map(normalizeHeader_);
      const map = {
        day: headers.findIndex(h => h === 'day' || h === 'day label'),
        session: headers.findIndex(h => h === 'session'),
        date: headers.findIndex(h => h.indexOf('date') >= 0),
        open: headers.findIndex(h => h.indexOf('opening time') >= 0 || h === 'open' || h === 'opening'),
        close: headers.findIndex(h => h.indexOf('closing time') >= 0 || h === 'close' || h === 'closing')
      };
      if (Object.values(map).every(i => i >= 0)) {
        headerRow = r + 1;
        headerMap = map;
        break;
      }
    }
  }

  // Compatibility fallback: use an existing Attendance_Settings tab, but never
  // overwrite it with hard-coded dates or times.
  if (!sh) {
    sh = ss.getSheetByName(SETTINGS_SHEET);
    if (sh) {
      const rowsToScan = Math.min(sh.getLastRow(), 10);
      const colsToScan = Math.min(Math.max(sh.getLastColumn(), 1), 10);
      const rows = sh.getRange(1, 1, Math.max(rowsToScan, 1), colsToScan).getDisplayValues();
      for (let r = 0; r < rows.length; r++) {
        const headers = rows[r].map(normalizeHeader_);
        const map = {
          day: headers.findIndex(h => h === 'day' || h === 'day label'),
          session: headers.findIndex(h => h === 'session'),
          date: headers.findIndex(h => h.indexOf('date') >= 0),
          open: headers.findIndex(h => h.indexOf('opening time') >= 0 || h === 'open' || h === 'opening'),
          close: headers.findIndex(h => h.indexOf('closing time') >= 0 || h === 'close' || h === 'closing')
        };
        if (Object.values(map).every(i => i >= 0)) {
          headerRow = r + 1;
          headerMap = map;
          break;
        }
      }
    }
  }

  if (!sh || headerRow < 1 || !headerMap) {
    throw new Error('Timetable not found. In the connected spreadsheet, provide a worksheet with columns Day, Session, Date, Opening time, and Closing time.');
  }

  const lastRow = sh.getLastRow();
  if (lastRow <= headerRow) {
    throw new Error('The timetable worksheet has headers but no session rows.');
  }
  const rawRows = sh.getRange(headerRow + 1, 1, lastRow - headerRow, sh.getLastColumn()).getValues();
  const displayRows = sh.getRange(headerRow + 1, 1, lastRow - headerRow, sh.getLastColumn()).getDisplayValues();

  const settings = [];
  rawRows.forEach((row, i) => {
    const display = displayRows[i];
    const dayText = String(display[headerMap.day] || row[headerMap.day] || '').trim();
    const sessionText = String(display[headerMap.session] || row[headerMap.session] || '').trim();
    const dateValue = normalizeWorkshopDate_(row[headerMap.date], display[headerMap.date]);
    const openValue = normalizeTime_(row[headerMap.open], display[headerMap.open]);
    const closeValue = normalizeTime_(row[headerMap.close], display[headerMap.close]);
    if (!dayText && !sessionText && !dateValue && !openValue && !closeValue) return;
    if (!dayText || !sessionText || !dateValue || !openValue || !closeValue) {
      throw new Error('Incomplete timetable row ' + (headerRow + i + 1) + ' in worksheet "' + sh.getName() + '". Fill Day, Session, Date, Opening time, and Closing time.');
    }
    settings.push({
      day: dayText,
      session: sessionText,
      date: dateValue,
      open: openValue,
      close: closeValue
    });
  });

  settings.sort((a, b) => {
    const dayA = Number((a.day.match(/\\d+/) || [0])[0]);
    const dayB = Number((b.day.match(/\\d+/) || [0])[0]);
    if (dayA !== dayB) return dayA - dayB;
    const codeA = /afternoon|अपराह्न|\\bAN\\b/i.test(a.session) ? 1 : 0;
    const codeB = /afternoon|अपराह्न|\\bAN\\b/i.test(b.session) ? 1 : 0;
    return codeA - codeB;
  });
  if (settings.length === 0) throw new Error('No timetable session rows were found in worksheet "' + sh.getName() + '".');
  return settings;
}

function normalizeWorkshopDate_(rawValue, displayedValue) {
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
  const signatureText = String(payload.signatureDataUrl || '');
  if (!['data:image/png;base64,', 'data:image/jpeg;base64,', 'data:image/jpg;base64,'].some(prefix => signatureText.startsWith(prefix))) {
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
  const mimeHeader = String(dataUrlParts[0] || '').toLowerCase();
  const mimeType = ['data:image/png;base64', 'data:image/jpeg;base64', 'data:image/jpg;base64'].includes(mimeHeader)
    ? mimeHeader.slice(5, mimeHeader.indexOf(';')) : '';
  if (!mimeType || !dataUrlParts[1]) throw new Error('The signature image could not be read.');
  const imageBytes = Utilities.base64Decode(dataUrlParts[1]);
  if (imageBytes.length > 1500000) throw new Error('Please use a smaller signature image (under 1.5 MB).');
  const blob = Utilities.newBlob(imageBytes, mimeType.replace('image/jpg','image/jpeg'),
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
