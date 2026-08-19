/**
 * Duty Roster Dashboard — backend
 * Bound to (or pointed at) the "Update schedule" spreadsheet.
 * Sheet ID is hardcoded below — change SPREADSHEET_ID if you copy this
 * to a different file, or delete the line to use the bound sheet instead.
 */

const SPREADSHEET_ID = '10AzEohHpnaus7ToYKTMprIpWxmI-Gg9XaoeSwpKPCgc';

const SHEET_DOCTORS = 'Doctors';
const SHEET_SCHEDULE = 'Schedule';
const SHEET_HISTORY = 'History';

// Schedule columns, in order, must match row 1 of the Schedule sheet.
const SCHEDULE_COLS = [
  'Date', 'Day', 'SPV_Day1', 'SPV_Day2', 'SPV_Day3', 'SPV_Night',
  'SPV_Off', 'HMV_Day', 'HMV_Night', 'UpdatedAt', 'UpdatedBy'
];

// Which schedule columns are editable duty slots (vs Date/Day/UpdatedAt/UpdatedBy)
const EDITABLE_FIELDS = [
  'SPV_Day1', 'SPV_Day2', 'SPV_Day3', 'SPV_Night', 'SPV_Off', 'HMV_Day', 'HMV_Night'
];

function ss_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  // JSON API mode — used by the static GitHub Pages frontend via fetch()
  if (action === 'data') {
    return jsonOut_(getDashboardData());
  }
  if (action === 'history') {
    const limit = e.parameter.limit ? Number(e.parameter.limit) : 20;
    return jsonOut_(getHistory(limit));
  }

  // Default: serve the bound HTML dashboard (Apps Script UI mode)
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Duty Roster Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * POST endpoint for the static frontend. Expects JSON body:
 * { action: 'updateAssignment', date, field, newName }
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'updateAssignment') {
      const result = updateAssignment(body.date, body.field, body.newName);
      return jsonOut_(result);
    }
    return jsonOut_({ error: 'Unknown action: ' + body.action });
  } catch (err) {
    return jsonOut_({ error: err.message });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Returns { doctors: [...], schedule: [...] } as plain objects for the frontend. */
function getDashboardData() {
  const ss = ss_();

  const docSheet = ss.getSheetByName(SHEET_DOCTORS);
  const docRows = docSheet.getDataRange().getValues();
  const docHeader = docRows.shift();
  const doctors = docRows
    .filter(r => r.join('').trim() !== '')
    .map(r => {
      const o = {};
      docHeader.forEach((h, i) => o[h] = r[i]);
      return o;
    });

  const schSheet = ss.getSheetByName(SHEET_SCHEDULE);
  const schRows = schSheet.getDataRange().getValues();
  schRows.shift(); // header
  const schedule = schRows
    .filter(r => r.join('').trim() !== '')
    .map(r => {
      const o = {};
      SCHEDULE_COLS.forEach((h, i) => o[h] = r[i]);
      // Normalize Date to yyyy-mm-dd string for the frontend
      if (o.Date instanceof Date) {
        o.Date = Utilities.formatDate(o.Date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      return o;
    });

  return { doctors: doctors, schedule: schedule };
}

/** Returns the most recent history rows, newest first. */
function getHistory(limit) {
  const ss = ss_();
  const hSheet = ss.getSheetByName(SHEET_HISTORY);
  const rows = hSheet.getDataRange().getValues();
  const header = rows.shift();
  const entries = rows
    .filter(r => r.join('').trim() !== '')
    .map(r => {
      const o = {};
      header.forEach((h, i) => o[h] = r[i]);
      if (o[header[0]] instanceof Date) {
        o[header[0]] = Utilities.formatDate(o[header[0]], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
      }
      return o;
    })
    .reverse();
  return limit ? entries.slice(0, limit) : entries;
}

/**
 * Update one duty slot for one date, and append a row to History.
 * date: 'yyyy-MM-dd', field: one of EDITABLE_FIELDS, newName: string
 */
function updateAssignment(date, field, newName) {
  if (EDITABLE_FIELDS.indexOf(field) === -1) {
    throw new Error('Field not editable: ' + field);
  }

  const ss = ss_();
  const schSheet = ss.getSheetByName(SHEET_SCHEDULE);
  const data = schSheet.getDataRange().getValues();
  const header = data[0];

  const dateColIdx = header.indexOf('Date');
  const fieldColIdx = header.indexOf(field);
  const updatedAtIdx = header.indexOf('UpdatedAt');
  const updatedByIdx = header.indexOf('UpdatedBy');

  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    let cellDate = data[i][dateColIdx];
    let cellDateStr = cellDate instanceof Date
      ? Utilities.formatDate(cellDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(cellDate);
    if (cellDateStr === date) {
      targetRow = i;
      break;
    }
  }

  if (targetRow === -1) {
    throw new Error('Date not found in schedule: ' + date);
  }

  const oldName = data[targetRow][fieldColIdx];
  const user = Session.getActiveUser().getEmail() || 'unknown';
  const now = new Date();
  const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

  // Write the new value + metadata (rows are 1-indexed in Sheets API)
  const sheetRow = targetRow + 1;
  schSheet.getRange(sheetRow, fieldColIdx + 1).setValue(newName);
  if (updatedAtIdx > -1) schSheet.getRange(sheetRow, updatedAtIdx + 1).setValue(nowStr);
  if (updatedByIdx > -1) schSheet.getRange(sheetRow, updatedByIdx + 1).setValue(user);

  // Append to History: ពេលវេលា, កាលបរិច្ឆេទ, វេន (Field), ឈ្មោះចាស់, ឈ្មោះថ្មី, អ្នកប្រើ
  const hSheet = ss.getSheetByName(SHEET_HISTORY);
  hSheet.appendRow([nowStr, date, field, oldName, newName, user]);

  return { success: true, oldName: oldName, newName: newName, updatedAt: nowStr, updatedBy: user };
}
