/**
 * Password Migration Script
 * 
 * این اسکریپت رمزهای عبور خام (ستون B) رو hash می‌کنه و در ستون C می‌ذاره.
 * 
 * نحوه استفاده:
 *   1. این کد رو در Apps Script بذار
 *   2. تابع migratePasswords() رو اجرا کن
 *   3. بعد از اتمام، اسکریپت رو پاک کن
 */

function migratePasswords() {
  var SHEET_ID = '16m9UdjL2QQIC5l-Nivi4Zfp8RueRyBZ1';
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  
  var migrated = 0;
  var skipped = 0;
  
  for (var i = 1; i < data.length; i++) {
    var user = String(data[i][0]).trim();
    var pass = String(data[i][1]).trim();
    var existingHash = String(data[i][2]).trim();
    
    if (!user || !pass) continue;
    
    // اگه hash از قبل وجود داره، رد شو
    if (existingHash) {
      skipped++;
      continue;
    }
    
    // hash کردن با SHA-256
    var hash = sha256(pass);
    
    // نوشتن hash در ستون C
    sheet.getRange(i + 1, 3).setValue(hash);
    migrated++;
    
    Logger.log('Migrated: ' + user);
  }
  
  Logger.log('Done! Migrated: ' + migrated + ', Skipped: ' + skipped);
  SpreadsheetApp.getUi().alert('Migration complete!\nMigrated: ' + migrated + '\nSkipped (already had hash): ' + skipped);
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
