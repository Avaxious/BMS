/**
 * Auto-Hash Trigger for Users Sheet
 * 
 * وقتی رمز عبور جدیدی در ستون B نوشته بشه، خودکار hash میشه در ستون C
 * 
 * نصب:
 *   1. این کد رو در Login Proxy Apps Script بذار
 *   2. تابع onEdit رو اجرا کن (یک بار کافیه)
 *   3. trigger ساخته میشه
 */

function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  
  if (sheet.getName() !== 'Users') return;
  if (range.getColumn() !== 2) return;
  
  var row = range.getRow();
  if (row <= 1) return;
  
  var pass = String(range.getValue()).trim();
  if (!pass) return;
  
  var hash = sha256(pass);
  sheet.getRange(row, 3).setValue(hash);
}

function sha256(input) {
  var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input);
  var hex = '';
  for (var i = 0; i < hash.length; i++) {
    var byte = hash[i];
    if (byte < 0) byte += 256;
    hex += ('0' + byte.toString(16)).slice(-2);
  }
  return hex;
}

function setupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet('16m9UdjL2QQIC5l-Nivi4Zfp8RueRyBZ1')
    .onEdit()
    .create();
  
  Logger.log('Trigger created!');
}
