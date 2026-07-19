/**
 * Enhanced Login Proxy with PBKDF2, Rate Limiting, and Audit Logging
 * Load Security-Config.gs before this script
 */

function doGet(e) {
  const action = e.parameter.action || '';

  if (action === 'ping') {
    return json({ok: true, v: 'v4-secure'});
  }

  if (action === 'validate') {
    const token = (e.parameter.token || '').trim();
    if (!token) return json({ok: false, error: 'no_token'});
    
    const validation = validateTokenWithExpiry(token);
    if (!validation.valid) {
      return json({ok: false, error: validation.reason});
    }
    
    // Refresh token expiry
    const usersSheet = SpreadsheetApp.openById(SECURITY_CONFIG.USERS_SHEET_ID)
      .getSheetByName('Users');
    const data = usersSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === validation.username) {
        const newExpiry = new Date(Date.now() + SECURITY_CONFIG.SESSION_TIMEOUT_MS).toISOString();
        usersSheet.getRange(i + 1, 5).setValue(newExpiry);
        logSecurityEvent('TOKEN_REFRESHED', validation.username, 'Token refreshed');
        break;
      }
    }
    
    return json({ok: true, username: validation.username});
  }

  if (action === 'logout') {
    const token = (e.parameter.token || '').trim();
    if (token) invalidateToken(token);
    return json({ok: true});
  }

  return json({ok: false, error: 'use POST for login'});
}

function doPost(e) {
  try {
    let body = {};
    try {
      body = JSON.parse(e.postData.contents || '{}');
    } catch (ex) {
      Logger.log('JSON parse error: ' + ex);
    }
    
    const action = body.action || e.parameter.action || 'login';

    if (action === 'login') {
      return handleLoginEnhanced(body);
    }

    if (action === 'logout') {
      const token = (body.token || e.parameter.token || '').trim();
      if (token) invalidateToken(token);
      return json({ok: true});
    }

    return json({ok: false, error: 'unknown_action'});
  } catch (err) {
    logSecurityEvent('SERVER_ERROR', 'SYSTEM', err.message);
    return json({ok: false, error: 'server_error'});
  }
}

/**
 * Enhanced login handler with PBKDF2 hashing and rate limiting
 */
function handleLoginEnhanced(body) {
  const user = (body.user || '').trim();
  const passwordHash = (body.hash || '').trim().toLowerCase();

  if (!user || !passwordHash) {
    logSecurityEvent('INVALID_CREDENTIALS', user || 'UNKNOWN', 'Missing credentials');
    return json({ok: false, error: 'missing_credentials'});
  }

  // Check if user is locked out
  if (isUserLockedOut(user)) {
    logSecurityEvent('LOCKOUT_ATTEMPT', user, 'User attempted login while locked out');
    return json({ok: false, error: 'account_locked', lockout_duration_min: Math.ceil(SECURITY_CONFIG.LOCKOUT_DURATION_MS / 60000)});
  }

  const usersSheet = SpreadsheetApp.openById(SECURITY_CONFIG.USERS_SHEET_ID)
    .getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== user) continue;

    // Get stored hash and salt
    const storedHashWithSalt = data[i][2] ? String(data[i][2]).trim() : '';
    
    // Format: hash:salt
    const [storedHash, salt] = storedHashWithSalt.split(':');
    
    if (!storedHash || !salt) {
      logSecurityEvent('PASSWORD_CONFIG_ERROR', user, 'Hash or salt missing');
      recordFailedLoginAttempt(user);
      return json({ok: false, error: 'invalid_credentials'});
    }

    // Hash provided password with stored salt
    const computedHash = hashPasswordPBKDF2(passwordHash, salt);

    if (computedHash.toLowerCase() === storedHash.toLowerCase()) {
      // Clear failed attempts
      clearFailedLoginAttempts(user);
      
      // Generate new token
      const token = generateSecureToken();
      const expiryTime = new Date(Date.now() + SECURITY_CONFIG.SESSION_TIMEOUT_MS).toISOString();
      const createdAt = new Date().toISOString();
      
      // Update sheet
      usersSheet.getRange(i + 1, 4).setValue(token);
      usersSheet.getRange(i + 1, 5).setValue(expiryTime);
      usersSheet.getRange(i + 1, 8).setValue(createdAt);
      
      logSecurityEvent('LOGIN_SUCCESS', user, 'User logged in successfully');
      
      return json({
        ok: true,
        token: token,
        expires_in: SECURITY_CONFIG.SESSION_TIMEOUT_MS,
        username: user
      });
    }
  }

  // User not found or password mismatch
  recordFailedLoginAttempt(user);
  logSecurityEvent('LOGIN_FAILED', user, 'Invalid password or user not found');
  return json({ok: false, error: 'invalid_credentials'});
}

/**
 * Invalidate token (logout)
 */
function invalidateToken(token) {
  try {
    const usersSheet = SpreadsheetApp.openById(SECURITY_CONFIG.USERS_SHEET_ID)
      .getSheetByName('Users');
    const data = usersSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() === token) {
        usersSheet.getRange(i + 1, 4).setValue('');
        usersSheet.getRange(i + 1, 5).setValue('');
        logSecurityEvent('LOGOUT', data[i][0], 'User logged out');
        break;
      }
    }
  } catch (e) {
    Logger.log('Logout error: ' + e.message);
  }
}

/**
 * Setup auto-hash trigger for password changes
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  
  if (sheet.getName() !== 'Users') return;
  if (range.getColumn() !== 2) return; // Column B = Password
  
  const row = range.getRow();
  if (row <= 1) return;
  
  const plainPassword = String(range.getValue()).trim();
  if (!plainPassword) return;

  // Generate salt
  const salt = generateSalt(SECURITY_CONFIG.SALT_LENGTH);
  
  // Hash password with PBKDF2
  const hash = hashPasswordPBKDF2(plainPassword, salt);
  
  // Store as hash:salt
  sheet.getRange(row, 3).setValue(hash + ':' + salt);
  
  // Clear password field (don't store plain text)
  sheet.getRange(row, 2).setValue('');
  
  logSecurityEvent('PASSWORD_CHANGED', sheet.getRange(row, 1).getValue(), 'Password updated with PBKDF2 hashing');
}

/**
 * Helper function
 */
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
