/**
 * តារាងវេន SPV + HMV — Duty Schedule (matches real paper/Excel layout)
 * Columns: Date | Day | SPV Day1 | SPV Day2 | SPV Day3 | SPV Night | SPV Off Today | HMV Day | HMV Night
 */

var SS = SpreadsheetApp.getActiveSpreadsheet();
var SHEET_DOCTORS = 'Doctors';
var SHEET_SCHEDULE = 'Schedule';
var SHEET_HISTORY = 'History';

var SCHEDULE_HEADERS = ['Date', 'Day', 'SPV_Day1', 'SPV_Day2', 'SPV_Day3', 'SPV_Night', 'SPV_Off', 'HMV_Day', 'HMV_Night', 'UpdatedAt', 'UpdatedBy'];
var SLOT_FIELDS = ['SPV_Day1', 'SPV_Day2', 'SPV_Day3', 'SPV_Night', 'HMV_Day', 'HMV_Night'];
var SPV_SLOTS = ['SPV_Day1', 'SPV_Day2', 'SPV_Day3', 'SPV_Night'];
var HMV_SLOTS = ['HMV_Day', 'HMV_Night'];
var DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/* ---------------- MENU ---------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('តារាងវេន')
    .addItem('បើក Web App URL', 'showSidebarUrl')
    .addItem('ត្រៀមទិន្នន័យ (Setup)', 'initializeSheets')
    .addItem('Reset ទាំងស្រុង (លុបចាស់ + ដាក់ថ្មី + ទិន្នន័យ Aug 2026)', 'resetAndSeed')
    .addToUi();
}

function showSidebarUrl() {
  var url = ScriptApp.getService().getUrl();
  var html = HtmlService.createHtmlOutput('<p>Web App URL:</p><p><a href="' + url + '" target="_blank">' + url + '</a></p>').setWidth(420).setHeight(110);
  SpreadsheetApp.getUi().showModalDialog(html, 'Web App Link');
}

/* ---------------- SETUP ---------------- */

function initializeSheets() {
  var doctorsSh = SS.getSheetByName(SHEET_DOCTORS) || SS.insertSheet(SHEET_DOCTORS);
  if (doctorsSh.getLastRow() === 0) {
    doctorsSh.appendRow(['ID', 'NameSPV', 'NameHMV', 'ClinicPool', 'isLinda', 'Active']);
    doctorsSh.setFrozenRows(1);
    seedDoctors(doctorsSh);
  }
  var scheduleSh = SS.getSheetByName(SHEET_SCHEDULE) || SS.insertSheet(SHEET_SCHEDULE);
  if (scheduleSh.getLastRow() === 0) {
    scheduleSh.appendRow(SCHEDULE_HEADERS);
    scheduleSh.setFrozenRows(1);
  }
  var historySh = SS.getSheetByName(SHEET_HISTORY) || SS.insertSheet(SHEET_HISTORY);
  if (historySh.getLastRow() === 0) {
    historySh.appendRow(['ពេលវេលា', 'កាលបរិច្ឆេទ', 'វេន (Field)', 'ឈ្មោះចាស់', 'ឈ្មោះថ្មី', 'អ្នកប្រើ']);
    historySh.setFrozenRows(1);
  }
  return 'OK';
}

function seedDoctors(sh) {
  sh.appendRow(['S1', 'Dr.Monyrachana', '', 'SPV', false, true]);
  sh.appendRow(['S2', 'Dr.Khean Daline', '', 'SPV', false, true]);
  sh.appendRow(['S3', 'Dr.Ka Seyla', '', 'SPV', false, true]);
  sh.appendRow(['S4', 'Dr.Phalla Daline', '', 'SPV', false, true]);
  sh.appendRow(['S5', 'Dr.Nann Vanna', '', 'SPV', false, true]);
  sh.appendRow(['S6', 'Dr.Khim Piseth', '', 'SPV', false, true]);
  sh.appendRow(['H1', '', 'វេជ្ជ.ឈន ចំរឿន', 'HMV', false, true]);
  sh.appendRow(['H2', '', 'វេជ្ជ.អ៊ុងគីមឡេង', 'HMV', false, true]);
  sh.appendRow(['H3', '', 'ហុក ប៊ុនណារ៉ា', 'HMV', false, true]);
  sh.appendRow(['LINDA', 'Dr.Ben Linda', 'វេជ្ជ.ប៊ិន លីនដា', 'Both', true, true]);
}

function resetAndSeed() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert('Reset ទាំងស្រុង?', 'នេះនឹងលុប sheet Doctors/Schedule/History ចាស់ ហើយបង្កើតថ្មីព្រមទាំងទិន្នន័យគំរូខែសីហា 2026។ បន្តទេ?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  // Google Sheets won't allow deleting every visible sheet — create a temporary
  // placeholder first so there's always at least one sheet left during cleanup.
  var placeholder = SS.insertSheet('__temp_reset__');

  [SHEET_DOCTORS, SHEET_SCHEDULE, SHEET_HISTORY].forEach(function (name) {
    var sh = SS.getSheetByName(name);
    if (sh) SS.deleteSheet(sh);
  });
  initializeSheets();
  seedAugust2026();

  SS.deleteSheet(placeholder);
  ui.alert('រួចរាល់! ទិន្នន័យខែសីហា 2026 ត្រូវបានបញ្ចូល (រួមទាំងប៉ះទង្គិច ២ថ្ងៃសម្រាប់សាកល្បង)។');
}

function seedAugust2026() {
  // [date, spvDay1, spvDay2, spvDay3, spvNight, spvOff, hmvDay, hmvNight]
  var rows = [
    ['2026-08-01', 'Dr.Monyrachana', 'Dr.Khean Daline', 'Dr.Ka Seyla', 'Dr.Monyrachana', '', 'វេជ្ជ.ប៊ិន លីនដា', 'វេជ្ជ.ប៊ិន លីនដា'],
    ['2026-08-02', 'Dr.Phalla Daline', 'Dr.Khean Daline', 'Dr.Ka Seyla', 'Dr.Monyrachana', '', 'វេជ្ជ.ប៊ិន លីនដា', 'វេជ្ជ.ប៊ិន លីនដា'],
    ['2026-08-03', 'Dr.Nann Vanna', 'Dr.Khean Daline', '', 'Dr.Khim Piseth', 'DAY OFF', 'វេជ្ជ.ឈន ចំរឿន', 'វេជ្ជ.ឈន ចំរឿន'],
    ['2026-08-04', 'Dr.Nann Vanna', 'Dr.Khean Daline', '', 'Dr.Nann Vanna', 'DAY OFF', 'វេជ្ជ.អ៊ុងគីមឡេង', 'វេជ្ជ.អ៊ុងគីមឡេង'],
    ['2026-08-05', 'Dr.Phalla Daline', 'Dr.Khean Daline', '', 'Dr.Nann Vanna', '', 'វេជ្ជ.ប៊ិន លីនដា', 'វេជ្ជ.ប៊ិន លីនដា'],
    ['2026-08-06', 'Dr.Ben Linda', 'Dr.Khean Daline', '', 'Dr.Monyrachana', '', 'វេជ្ជ.ឈន ចំរឿន', 'វេជ្ជ.ឈន ចំរឿន'],
    ['2026-08-07', 'Dr.Phalla Daline', 'Dr.Ben Linda', '', 'Dr.Phalla Daline', '', 'វេជ្ជ.ឈន ចំរឿន', 'វេជ្ជ.អ៊ុងគីមឡេង'],
    ['2026-08-08', 'Dr.Phalla Daline', 'Dr.Khean Daline', 'Dr.Ben Linda', 'Dr.Phalla Daline', '', 'វេជ្ជ.អ៊ុងគីមឡេង', 'វេជ្ជ.អ៊ុងគីមឡេង'],
    ['2026-08-09', 'Dr.Nann Vanna', 'Dr.Khean Daline', 'Dr.Ka Seyla', 'Dr.Nann Vanna', '', 'វេជ្ជ.ប៊ិន លីនដា', 'វេជ្ជ.ប៊ិន លីនដា'],
    ['2026-08-10', 'Dr.Ben Linda', 'Dr.Khean Daline', '', 'Dr.Monyrachana', '', 'វេជ្ជ.អ៊ុងគីមឡេង', 'វេជ្ជ.អ៊ុងគីមឡេង'],
    ['2026-08-11', 'Dr.Phalla Daline', 'Dr.Khean Daline', '', 'Dr.Khim Piseth', '', 'វេជ្ជ.ប៊ិន លីនដា', 'វេជ្ជ.ប៊ិន លីនដា'],
    ['2026-08-12', 'Dr.Ben Linda', 'Dr.Phalla Daline', '', 'Dr.Ben Linda', '', 'ហុក ប៊ុនណារ៉ា', 'ហុក ប៊ុនណារ៉ា'],
    ['2026-08-13', 'Dr.Nann Vanna', 'Dr.Khean Daline', '', 'Dr.Nann Vanna', '', 'ហុក ប៊ុនណារ៉ា', 'វេជ្ជ.ប៊ិន លីនដា'],
    ['2026-08-14', 'Dr.Ben Linda', 'Dr.Khean Daline', '', 'Dr.Monyrachana', '', 'វេជ្ជ.ឈន ចំរឿន', 'វេជ្ជ.ឈន ចំរឿន'],
    ['2026-08-15', 'Dr.Khim Piseth', 'Dr.Ben Linda', 'Dr.Phalla Daline', 'Dr.Khim Piseth', '', 'វេជ្ជ.អ៊ុងគីមឡេង', 'វេជ្ជ.អ៊ុងគីមឡេង'],
    ['2026-08-16', 'Dr.Khim Piseth', 'Dr.Khean Daline', 'Dr.Ben Linda', 'Dr.Khim Piseth', '', 'វេជ្ជ.ឈន ចំរឿន', 'វេជ្ជ.ឈន ចំរឿន'],
    ['2026-08-17', 'Dr.Phalla Daline', 'Dr.Ben Linda', '', 'Dr.Phalla Daline', '', 'វេជ្ជ.ឈន ចំរឿន', 'វេជ្ជ.ឈន ចំរឿន'],
    ['2026-08-18', 'Dr.Nann Vanna', 'Dr.Phalla Daline', '', 'Dr.Nann Vanna', 'DAY OFF', 'វេជ្ជ.អ៊ុងគីមឡេង', 'វេជ្ជ.អ៊ុងគីមឡេង'],
    ['2026-08-19', 'Dr.Nann Vanna', 'Dr.Khean Daline', '', 'Dr.Khim Piseth', '', 'វេជ្ជ.ប៊ិន លីនដា', 'វេជ្ជ.ប៊ិន លីនដា'],
    ['2026-08-20', 'Dr.Ben Linda', 'Dr.Khean Daline', '', 'Dr.Ben Linda', '', 'វេជ្ជ.ឈន ចំរឿន', 'វេជ្ជ.ឈន ចំរឿន'],
    ['2026-08-21', 'Dr.Monyrachana', 'Dr.Khean Daline', '', 'Dr.Monyrachana', '', 'វេជ្ជ.ប៊ិន លីនដា', 'វេជ្ជ.អ៊ុងគីមឡេង'],
    ['2026-08-22', 'Dr.Phalla Daline', 'Dr.Khean Daline', 'Dr.Ka Seyla', 'Dr.Phalla Daline', '', 'វេជ្ជ.ប៊ិន លីនដា', 'វេជ្ជ.ប៊ិន លីនដា'],
    ['2026-08-23', 'Dr.Khim Piseth', 'Dr.Nann Vanna', 'Dr.Ka Seyla', 'Dr.Khim Piseth', 'DAY OFF', 'វេជ្ជ.អ៊ុងគីមឡេង', 'វេជ្ជ.អ៊ុងគីមឡេង'],
    ['2026-08-24', 'Dr.Nann Vanna', 'Dr.Khean Daline', '', 'Dr.Nann Vanna', '', 'វេជ្ជ.ប៊ិន លីនដា', 'វេជ្ជ.អ៊ុងគីមឡេង'],
    ['2026-08-25', 'Dr.Phalla Daline', 'Dr.Khean Daline', '', 'Dr.Nann Vanna', 'DAY OFF', 'ហុក ប៊ុនណារ៉ា', 'ហុក ប៊ុនណារ៉ា'],
    ['2026-08-26', 'Dr.Phalla Daline', 'Dr.Khean Daline', '', 'Dr.Monyrachana', 'DAY OFF', 'ហុក ប៊ុនណារ៉ា', 'ហុក ប៊ុនណារ៉ា'],
    ['2026-08-27', 'Dr.Nann Vanna', 'Dr.Khean Daline', '', 'Dr.Khim Piseth', 'DAY OFF', 'វេជ្ជ.ឈន ចំរឿន', 'វេជ្ជ.ឈន ចំរឿន'],
    ['2026-08-28', 'Dr.Nann Vanna', 'Dr.Ben Linda', '', 'Dr.Nann Vanna', '', 'ហុក ប៊ុនណារ៉ា', 'វេជ្ជ.ប៊ិន លីនដា'],
    ['2026-08-29', 'Dr.Monyrachana', 'Dr.Nann Vanna', '', 'Dr.Monyrachana', 'DAY OFF', 'វេជ្ជ.អ៊ុងគីមឡេង', 'វេជ្ជ.អ៊ុងគីមឡេង'],
    ['2026-08-30', 'Dr.Phalla Daline', 'Dr.Khean Daline', 'Dr.Ben Linda', 'Dr.Phalla Daline', '', 'វេជ្ជ.អ៊ុងគីមឡេង', 'វេជ្ជ.អ៊ុងគីមឡេង'],
    ['2026-08-31', 'Dr.Ben Linda', 'Dr.Khean Daline', '', 'Dr.Nann Vanna', '', 'ហុក ប៊ុនណារ៉ា', 'វេជ្ជ.ប៊ិន លីនដា']
  ];
  var sh = SS.getSheetByName(SHEET_SCHEDULE);
  var now = new Date();
  rows.forEach(function (r) {
    var dow = DAY_NAMES[new Date(r[0] + 'T00:00:00').getDay()];
    sh.appendRow([r[0], dow, r[1], r[2], r[3], r[4], r[5], r[6], r[7], now, 'seed']);
  });
}

/* ---------------- WEB APP ---------------- */

// Two modes on the same /exec URL:
//   1) No "action" param  -> serves the Apps Script HTML app (kept as a fallback/testing view)
//   2) ?action=... present -> serves a JSON API response for the static GitHub Pages website
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter);
  }
  initializeSheets();
  var tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.initialDataJson = JSON.stringify(getAllData());
  return tmpl.evaluate()
    .setTitle('តារាងវេន SPV + HMV')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// JSON/JSONP API for the static website (GitHub Pages). Apps Script Web Apps
// do NOT reliably send CORS headers for cross-origin fetch(), so when a
// "callback" param is present we respond as JSONP (a <script> tag load,
// which browsers never block via CORS) instead of raw JSON.
function handleApiRequest(params) {
  initializeSheets();
  var action = params.action;
  var result;
  try {
    if (action === 'getAllData') {
      result = getAllData();
    } else if (action === 'saveCell') {
      result = saveCell(params.date, params.field, params.value || '');
    } else if (action === 'addDoctor') {
      result = addDoctor({
        nameSPV: params.nameSPV || '',
        nameHMV: params.nameHMV || '',
        clinicPool: params.clinicPool || 'SPV',
        isLinda: params.isLinda === 'true'
      });
    } else {
      result = { ok: false, error: 'unknown action: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }

  if (params.callback) {
    var js = params.callback + '(' + JSON.stringify(result) + ');';
    return ContentService.createTextOutput(js).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// Receives a fire-and-forget backup from the offline website (localStorage is
// its real source of truth — this just "stocks" a copy into the Sheet as
// history). Posted as a hidden-iframe form submit, so no CORS issues and no
// response is read back by the browser.
function doPost(e) {
  try {
    var payload = JSON.parse((e && e.parameter && e.parameter.payload) || '{}');
    bulkBackup(payload);
  } catch (err) {
    // best-effort only — never surfaces to the website
  }
  return ContentService.createTextOutput('ok');
}

function bulkBackup(payload) {
  initializeSheets();
  if (payload.doctors && payload.doctors.length) upsertDoctors(payload.doctors);
  if (payload.schedule && payload.schedule.length) upsertSchedule(payload.schedule);
}

function upsertDoctors(doctors) {
  var sh = SS.getSheetByName(SHEET_DOCTORS);
  var existing = sheetToObjects(sh);
  var rowById = {};
  existing.forEach(function (d) { rowById[d.ID] = d._row; });

  doctors.forEach(function (d) {
    var vals = [d.ID, d.NameSPV || '', d.NameHMV || '', d.ClinicPool || '', !!d.isLinda, d.Active !== false];
    if (rowById[d.ID]) {
      sh.getRange(rowById[d.ID], 1, 1, vals.length).setValues([vals]);
    } else {
      sh.appendRow(vals);
    }
  });
}

function upsertSchedule(schedule) {
  var sh = SS.getSheetByName(SHEET_SCHEDULE);
  var existing = sheetToObjects(sh).map(function (r) { r.Date = fmtDate(r.Date); return r; });
  var rowByDate = {};
  existing.forEach(function (r) { rowByDate[r.Date] = r._row; });
  var headers = SCHEDULE_HEADERS;

  schedule.forEach(function (row) {
    var vals = headers.map(function (h) {
      if (h === 'UpdatedAt') return new Date();
      if (h === 'UpdatedBy') return 'website-backup';
      return row[h] || '';
    });
    if (rowByDate[row.Date]) {
      sh.getRange(rowByDate[row.Date], 1, 1, vals.length).setValues([vals]);
    } else {
      sh.appendRow(vals);
    }
  });
}

/* ---------------- HELPERS ---------------- */

function sheetToObjects(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (row.join('') === '') continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
    obj._row = r + 1;
    out.push(obj);
  }
  return out;
}

function fmtDate(d) {
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return d;
}

/* ---------------- DOCTORS ---------------- */

function getDoctors() {
  var sh = SS.getSheetByName(SHEET_DOCTORS);
  return sheetToObjects(sh).filter(function (d) { return d.Active !== false; });
}

function addDoctor(doc) {
  var sh = SS.getSheetByName(SHEET_DOCTORS);
  var id = (doc.clinicPool === 'HMV' ? 'H' : doc.clinicPool === 'SPV' ? 'S' : 'B') + new Date().getTime();
  var isLinda = !!doc.isLinda;
  sh.appendRow([id, doc.nameSPV || '', doc.nameHMV || '', doc.clinicPool, isLinda, true]);
  return {
    ok: true,
    doctor: { ID: id, NameSPV: doc.nameSPV || '', NameHMV: doc.nameHMV || '', ClinicPool: doc.clinicPool, isLinda: isLinda, Active: true }
  };
}

function findDoctorByName(doctors, name) {
  if (!name) return null;
  return doctors.filter(function (d) { return d.NameSPV === name || d.NameHMV === name; })[0] || null;
}

/* ---------------- SCHEDULE ---------------- */

function getMonthSchedule(year, month) {
  var sh = SS.getSheetByName(SHEET_SCHEDULE);
  var rows = sheetToObjects(sh);
  var prefix = year + '-' + ('0' + month).slice(-2);
  return rows
    .map(function (r) { r.Date = fmtDate(r.Date); return r; })
    .filter(function (r) { return String(r.Date).indexOf(prefix) === 0; })
    .sort(function (a, b) { return a.Date.localeCompare(b.Date); });
}

function getOrCreateRow(sh, date) {
  var data = sh.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (fmtDate(data[r][0]) === date) return r + 1;
  }
  var dow = DAY_NAMES[new Date(date + 'T00:00:00').getDay()];
  sh.appendRow([date, dow, '', '', '', '', '', '', '', new Date(), '']);
  return sh.getLastRow();
}

// Only same time-of-day counts as a clash: Day-slot vs Day-slot, or Night-slot vs Night-slot.
// A doctor working SPV day then HMV night (or vice versa) is a normal back-to-back schedule, not a clash.
var SPV_DAY_SLOTS = ['SPV_Day1', 'SPV_Day2', 'SPV_Day3'];
var SPV_NIGHT_SLOTS = ['SPV_Night'];
var HMV_DAY_SLOTS = ['HMV_Day'];
var HMV_NIGHT_SLOTS = ['HMV_Night'];

function namesForSlots(rowObj, slots) {
  return slots.map(function (f) { return rowObj[f]; }).filter(Boolean);
}

function sameDoctorOverlap(namesA, namesB, doctors) {
  var clashes = [];
  namesA.forEach(function (a) {
    var docA = findDoctorByName(doctors, a);
    if (!docA) return;
    namesB.forEach(function (b) {
      var docB = findDoctorByName(doctors, b);
      if (docB && docB.ID === docA.ID) clashes.push(docA.NameSPV || docA.NameHMV);
    });
  });
  return clashes;
}

function computeClashesForRow(rowObj, doctors) {
  var dayClash = sameDoctorOverlap(namesForSlots(rowObj, SPV_DAY_SLOTS), namesForSlots(rowObj, HMV_DAY_SLOTS), doctors);
  var nightClash = sameDoctorOverlap(namesForSlots(rowObj, SPV_NIGHT_SLOTS), namesForSlots(rowObj, HMV_NIGHT_SLOTS), doctors);
  return dayClash.concat(nightClash);
}

var AUTO_OFF_TEXT = 'DAY OFF';

function saveCell(date, field, value) {
  if (SLOT_FIELDS.indexOf(field) === -1 && field !== 'SPV_Off') {
    return { ok: false, error: 'invalid field' };
  }
  var sh = SS.getSheetByName(SHEET_SCHEDULE);
  var rowIdx = getOrCreateRow(sh, date);
  var headers = SCHEDULE_HEADERS;
  var colIdx = headers.indexOf(field) + 1;
  var oldValue = sh.getRange(rowIdx, colIdx).getValue();
  var user = Session.getActiveUser().getEmail() || 'unknown';

  sh.getRange(rowIdx, colIdx).setValue(value);
  sh.getRange(rowIdx, headers.indexOf('UpdatedAt') + 1).setValue(new Date());
  sh.getRange(rowIdx, headers.indexOf('UpdatedBy') + 1).setValue(user);

  var historyEntry = null;
  if (String(oldValue) !== String(value)) {
    var histSh = SS.getSheetByName(SHEET_HISTORY);
    var now = new Date();
    histSh.appendRow([now, date, field, oldValue || '', value || '', user]);
    historyEntry = {
      'ពេលវេលា': Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      'កាលបរិច្ឆេទ': date,
      'វេន (Field)': field,
      'ឈ្មោះចាស់': oldValue || '',
      'ឈ្មោះថ្មី': value || '',
      'អ្នកប្រើ': user
    };
  }

  // Auto DAY OFF: if every SPV/HMV slot is now empty, write "DAY OFF" into SPV_Off
  // (only if that cell is currently blank, so a manual note is never overwritten).
  // If a slot just got filled in and SPV_Off still holds the auto-written text,
  // clear it back out automatically.
  if (field !== 'SPV_Off') {
    var offColIdx = headers.indexOf('SPV_Off') + 1;
    var offValue = sh.getRange(rowIdx, offColIdx).getValue();
    var rowNow = sh.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
    var allEmpty = SLOT_FIELDS.every(function (f) { return !rowNow[headers.indexOf(f)]; });
    if (allEmpty && !offValue) {
      sh.getRange(rowIdx, offColIdx).setValue(AUTO_OFF_TEXT);
    } else if (!allEmpty && offValue === AUTO_OFF_TEXT) {
      sh.getRange(rowIdx, offColIdx).setValue('');
    }
  }

  var rowValues = sh.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
  var rowObj = {};
  headers.forEach(function (h, i) { rowObj[h] = rowValues[i]; });
  rowObj.Date = fmtDate(rowObj.Date);
  var clashes = computeClashesForRow(rowObj, getDoctors());

  return { ok: true, clashes: clashes, historyEntry: historyEntry, offValue: rowObj.SPV_Off };
}

/* ---------------- HISTORY ---------------- */

function getHistory(filters) {
  var sh = SS.getSheetByName(SHEET_HISTORY);
  var rows = sheetToObjects(sh);
  rows.forEach(function (r) {
    r['ពេលវេលា'] = Utilities.formatDate(new Date(r['ពេលវេលា']), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  });
  rows.sort(function (a, b) { return b['ពេលវេលា'].localeCompare(a['ពេលវេលា']); });
  if (filters && filters.field) rows = rows.filter(function (r) { return r['វេន (Field)'] === filters.field; });
  return rows.slice(0, 300);
}

/* ---------------- PAYROLL (rate-based, split by clinic) ---------------- */

// SPV clinic: any 12H shift (Day1/Day2/Day3/Night) = $25 / 100,000 KHR
// HMV clinic: Day 12H = $24, Night 12H = $20
var RATES_USD = { SPV_Day1: 25, SPV_Day2: 25, SPV_Day3: 25, SPV_Night: 25, HMV_Day: 24, HMV_Night: 20 };
var KHR_PER_USD = 4000; // 25 USD = 100,000 KHR

function getPayrollSummary(year, month) {
  var rows = getMonthSchedule(year, month);
  var doctors = getDoctors();
  var spvCounts = {};
  var hmvCounts = {};

  rows.forEach(function (r) {
    SLOT_FIELDS.forEach(function (f) {
      var name = r[f];
      if (!name) return;
      var doc = findDoctorByName(doctors, name);
      var key = doc ? doc.ID : name;
      var isSpv = f.indexOf('SPV') === 0;
      var displayName = doc ? (isSpv ? (doc.NameSPV || doc.NameHMV) : (doc.NameHMV || doc.NameSPV)) : name;
      var bucket = isSpv ? spvCounts : hmvCounts;
      if (!bucket[key]) bucket[key] = { name: displayName, shifts: 0, usd: 0 };
      bucket[key].shifts++;
      bucket[key].usd += RATES_USD[f] || 0;
    });
  });

  function toList(bucket) {
    return Object.keys(bucket).map(function (k) {
      var c = bucket[k];
      c.khr = c.usd * KHR_PER_USD;
      return c;
    }).sort(function (a, b) { return b.usd - a.usd; });
  }

  var lindaDoc = doctors.filter(function (d) { return d.isLinda; })[0];
  var linda = { usd: 0, khr: 0, spvShifts: 0, hmvShifts: 0, shifts: 0 };
  if (lindaDoc) {
    var spvEntry = spvCounts[lindaDoc.ID];
    var hmvEntry = hmvCounts[lindaDoc.ID];
    if (spvEntry) { linda.usd += spvEntry.usd; linda.spvShifts = spvEntry.shifts; }
    if (hmvEntry) { linda.usd += hmvEntry.usd; linda.hmvShifts = hmvEntry.shifts; }
    linda.shifts = linda.spvShifts + linda.hmvShifts;
    linda.khr = linda.usd * KHR_PER_USD;
  }

  return { spv: toList(spvCounts), hmv: toList(hmvCounts), linda: linda };
}

function getDayPay(rowObj) {
  var usd = 0;
  SLOT_FIELDS.forEach(function (f) { if (rowObj[f]) usd += RATES_USD[f] || 0; });
  return { usd: usd, khr: usd * KHR_PER_USD };
}

/* ---------------- INIT DATA FOR CLIENT ---------------- */

function getInitialData() {
  var today = new Date();
  return {
    doctors: getDoctors(), year: today.getFullYear(), month: today.getMonth() + 1,
    rates: RATES_USD, khrPerUsd: KHR_PER_USD
  };
}

// Loads everything once — used by doGet() (embedded) and the "ទាញទិន្នន័យចុងក្រោយ" refresh link.
// Returns the FULL schedule history (all months ever entered) so the client can page
// between months locally without extra server round-trips.
function getAllData() {
  var today = new Date();
  var scheduleSh = SS.getSheetByName(SHEET_SCHEDULE);
  var scheduleRows = sheetToObjects(scheduleSh).map(function (r) { r.Date = fmtDate(r.Date); return r; });

  var histSh = SS.getSheetByName(SHEET_HISTORY);
  var historyRows = sheetToObjects(histSh);
  historyRows.forEach(function (r) {
    r['ពេលវេលា'] = Utilities.formatDate(new Date(r['ពេលវេលា']), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  });
  historyRows.sort(function (a, b) { return b['ពេលវេលា'].localeCompare(a['ពេលវេលា']); });
  historyRows = historyRows.slice(0, 300);

  return {
    doctors: getDoctors(),
    schedule: scheduleRows,
    history: historyRows,
    rates: RATES_USD,
    khrPerUsd: KHR_PER_USD,
    year: today.getFullYear(),
    month: today.getMonth() + 1
  };
}
