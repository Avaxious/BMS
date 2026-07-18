function doGet(e) {
  var action = e.parameter.action || '';

  if (action === 'ping') {
    return json({ok:true, v:'v3'});
  }

  if (action === 'validate') {
    var token = (e.parameter.token || '').trim();
    if (!token) return json({ok:false, error:'no token'});
    var sheet = SpreadsheetApp.openById('16m9UdjL2QQIC5l-Nivi4Zfp8RueRyBZ1').getSheetByName('Users');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() !== token) continue;
      var exp = new Date(data[i][4]);
      if (isNaN(exp.getTime()) || exp <= new Date()) {
        return json({ok:false, error:'expired'});
      }
      sheet.getRange(i + 1, 5).setValue(new Date(Date.now() + 1800000).toISOString());
      return json({ok:true});
    }
    return json({ok:false, error:'invalid'});
  }

  if (action === 'logout') {
    var token = (e.parameter.token || '').trim();
    if (token) invalidateToken(token);
    return json({ok:true});
  }

  return json({ok:false, error:'use POST for login'});
}

function doPost(e) {
  try {
    var body = {};
    try { body = JSON.parse(e.postData.contents || '{}'); } catch(ex) {}
    var action = body.action || e.parameter.action || 'login';

    if (action === 'login') {
      return handleLogin(body);
    }

    if (action === 'logout') {
      var token = (body.token || e.parameter.token || '').trim();
      if (token) invalidateToken(token);
      return json({ok:true});
    }

    return json({ok:false, error:'unknown action'});
  } catch(err) {
    return json({ok:false, error:'server error'});
  }
}

function handleLogin(body) {
  var user = (body.user || '').trim();
  var hash = (body.hash || '').trim().toLowerCase();
  var pass = (body.pass || '').trim();

  if (!user || !hash) {
    return json({ok:false, error:'missing credentials'});
  }

  var sheet = SpreadsheetApp.openById('16m9UdjL2QQIC5l-Nivi4Zfp8RueRyBZ1').getSheetByName('Users');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== user) continue;

    var storedHash = data[i][2] ? String(data[i][2]).trim().toLowerCase() : '';

    if (storedHash && hash === storedHash) {
      var token = generateToken();
      sheet.getRange(i + 1, 4).setValue(token);
      sheet.getRange(i + 1, 5).setValue(new Date(Date.now() + 1800000).toISOString());
      return json({ok:true, token:token});
    }

    if (!storedHash && pass && String(data[i][1]).trim() === pass) {
      sheet.getRange(i + 1, 3).setValue(hash);
      var token = generateToken();
      sheet.getRange(i + 1, 4).setValue(token);
      sheet.getRange(i + 1, 5).setValue(new Date(Date.now() + 1800000).toISOString());
      return json({ok:true, token:token});
    }
  }

  return json({ok:false, error:'invalid credentials'});
}

function invalidateToken(token) {
  try {
    var sheet = SpreadsheetApp.openById('16m9UdjL2QQIC5l-Nivi4Zfp8RueRyBZ1').getSheetByName('Users');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() === token) {
        sheet.getRange(i + 1, 4).setValue('');
        sheet.getRange(i + 1, 5).setValue('');
        break;
      }
    }
  } catch(e) {}
}

function generateToken() {
  var hex = '';
  var chars = '0123456789abcdef';
  for (var i = 0; i < 48; i++) {
    hex += chars[Math.floor(Math.random() * 16)];
  }
  return 'tok_' + hex;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
