var DATA_SHEET_ID = '1rct6063UB1M3gXuuU6NmTYqqnvurn6pQ';
var USERS_SHEET_ID = '16m9UdjL2QQIC5l-Nivi4Zfp8RueRyBZ1';
var MAX_SESSION_MS = 30 * 60 * 1000;

function doPost(e) {
  try {
    var body = {};
    try { body = JSON.parse(e.postData.contents || '{}'); } catch(ex) {}
    var token = (body.token || e.parameter.token || '').trim();
    if (!token) return forbidden();
    if (!validateToken(token)) return forbidden();
    return serveData();
  } catch(err) {
    return error('Server error');
  }
}

function doGet(e) {
  return forbidden();
}

function validateToken(token) {
  var usersSheet = SpreadsheetApp.openById(USERS_SHEET_ID).getSheetByName('Users');
  var data = usersSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim() !== token) continue;
    var expiry = new Date(data[i][4]);
    if (isNaN(expiry.getTime()) || expiry <= new Date()) return false;
    var createdAt = data[i][7] ? new Date(data[i][7]) : null;
    if (createdAt && (new Date().getTime() - createdAt.getTime()) > MAX_SESSION_MS) {
      usersSheet.getRange(i + 1, 4).setValue('');
      usersSheet.getRange(i + 1, 5).setValue('');
      usersSheet.getRange(i + 1, 7).setValue('');
      return false;
    }
    usersSheet.getRange(i + 1, 5).setValue(new Date(Date.now() + 30 * 60 * 1000).toISOString());
    return true;
  }
  return false;
}

function serveData() {
  var ss = SpreadsheetApp.openById(DATA_SHEET_ID);
  var sheet = ss.getSheetByName('Vessel Checking');
  if (!sheet) return notFound('Sheet not found');
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return notFound('No data');
  var csv = '';
  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var vals = [];
    for (var c = 0; c < row.length; c++) {
      var v = row[c];
      if (v === null || v === undefined) v = '';
      v = String(v).replace(/"/g, '""');
      if (v.indexOf(',') !== -1 || v.indexOf('"') !== -1 || v.indexOf('\n') !== -1) v = '"' + v + '"';
      vals.push(v);
    }
    csv += vals.join(',') + '\n';
  }
  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
}

function forbidden() {
  return ContentService.createTextOutput('Forbidden').setMimeType(ContentService.MimeType.TEXT);
}

function notFound(msg) {
  return ContentService.createTextOutput('Not found').setMimeType(ContentService.MimeType.TEXT);
}

function error(msg) {
  return ContentService.createTextOutput('Error').setMimeType(ContentService.MimeType.TEXT);
}
