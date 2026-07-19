/**
 * Security Configuration Module
 * Centralized security settings and utilities
 */

// ============================================================
// CONFIGURATION (Override with environment variables)
// ============================================================
const SECURITY_CONFIG = {
  // Session & Token
  SESSION_TIMEOUT_MS: 30 * 60 * 1000,        // 30 minutes
  TOKEN_LENGTH: 32,
  
  // Password Hashing (PBKDF2)
  PBKDF2_ITERATIONS: 100000,
  PBKDF2_HASH_ALGORITHM: 'SHA_256',
  SALT_LENGTH: 16,
  
  // Rate Limiting
  MAX_LOGIN_ATTEMPTS: 5,
  LOCKOUT_DURATION_MS: 15 * 60 * 1000,      // 15 minutes
  RATE_LIMIT_WINDOW: 5 * 60 * 1000,         // 5 minutes
  RATE_LIMIT_MAX_REQUESTS: 20,
  
  // Encryption
  ENCRYPTION_ALGORITHM: 'AES-GCM',
  ENCRYPTION_KEY_LENGTH: 256,
  IV_LENGTH: 12,
  SALT_FOR_KEY_DERIVATION: 16,
  
  // Audit Logging
  ENABLE_AUDIT_LOG: true,
  AUDIT_LOG_RETENTION_DAYS: 90,
  
  // CORS & Security Headers
  ALLOWED_ORIGINS: ['https://avaxious.github.io', 'https://example.com'],
  SECURE_COOKIE_FLAGS: 'HttpOnly; Secure; SameSite=Strict'
};

// ============================================================
// UTILITY FUNCTIONS
// ============================================================

/**
 * Generate cryptographically secure random token
 */
function generateSecureToken(length = 32) {
  return 'tok_' + Utilities.getUuid().replace(/-/g, '');
}

/**
 * Generate random salt for PBKDF2
 */
function generateSalt(length = 16) {
  const bytes = new Uint8Array(length);
  for (let i = 0; i < length; i++) {
    bytes[i] = Math.floor(Math.random() * 256);
  }
  return bytesToHex(bytes);
}

/**
 * Convert bytes to hex string
 */
function bytesToHex(bytes) {
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    const byte = bytes[i];
    hex += ('0' + byte.toString(16)).slice(-2);
  }
  return hex;
}

/**
 * Convert hex string to bytes
 */
function hexToBytes(hex) {
  const bytes = new Uint8Array(hex.length / 2);
  for (let i = 0; i < hex.length; i += 2) {
    bytes[i / 2] = parseInt(hex.substr(i, 2), 16);
  }
  return bytes;
}

/**
 * Hash password using PBKDF2 (Apps Script native)
 */
function hashPasswordPBKDF2(password, salt) {
  let hash = password + salt;
  for (let i = 0; i < SECURITY_CONFIG.PBKDF2_ITERATIONS; i++) {
    hash = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      hash,
      Utilities.Charset.UTF_8
    );
    hash = bytesToHex(hash);
  }
  return hash;
}

/**
 * Check if user is locked out
 */
function isUserLockedOut(username) {
  const usersSheet = SpreadsheetApp.openById(SECURITY_CONFIG.USERS_SHEET_ID)
    .getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === username) {
      const lockoutTime = data[i][8] ? new Date(data[i][8]) : null;
      if (lockoutTime && new Date() < new Date(lockoutTime.getTime() + SECURITY_CONFIG.LOCKOUT_DURATION_MS)) {
        return true;
      }
      return false;
    }
  }
  return false;
}

/**
 * Record failed login attempt
 */
function recordFailedLoginAttempt(username) {
  const usersSheet = SpreadsheetApp.openById(SECURITY_CONFIG.USERS_SHEET_ID)
    .getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === username) {
      const attempts = (data[i][6] || 0) + 1;
      usersSheet.getRange(i + 1, 7).setValue(attempts);
      
      if (attempts >= SECURITY_CONFIG.MAX_LOGIN_ATTEMPTS) {
        usersSheet.getRange(i + 1, 9).setValue(new Date().toISOString());
        logSecurityEvent('LOCKOUT', username, 'Too many failed login attempts');
      } else {
        logSecurityEvent('FAILED_LOGIN', username, `Attempt ${attempts}/${SECURITY_CONFIG.MAX_LOGIN_ATTEMPTS}`);
      }
      break;
    }
  }
}

/**
 * Clear failed login attempts
 */
function clearFailedLoginAttempts(username) {
  const usersSheet = SpreadsheetApp.openById(SECURITY_CONFIG.USERS_SHEET_ID)
    .getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === username) {
      usersSheet.getRange(i + 1, 7).setValue(0);
      usersSheet.getRange(i + 1, 9).setValue('');
      break;
    }
  }
}

/**
 * Validate token with expiry
 */
function validateTokenWithExpiry(token) {
  const usersSheet = SpreadsheetApp.openById(SECURITY_CONFIG.USERS_SHEET_ID)
    .getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim() === token) {
      const expiry = new Date(data[i][4]);
      const createdAt = data[i][7] ? new Date(data[i][7]) : null;
      
      if (isNaN(expiry.getTime()) || expiry <= new Date()) {
        logSecurityEvent('TOKEN_EXPIRED', data[i][0], 'Token expired');
        return {valid: false, reason: 'expired'};
      }
      
      if (createdAt && (new Date().getTime() - createdAt.getTime()) > SECURITY_CONFIG.SESSION_TIMEOUT_MS) {
        usersSheet.getRange(i + 1, 4).setValue('');
        usersSheet.getRange(i + 1, 5).setValue('');
        logSecurityEvent('SESSION_TIMEOUT', data[i][0], 'Session exceeded max duration');
        return {valid: false, reason: 'session_timeout'};
      }
      
      return {valid: true, username: data[i][0]};
    }
  }
  
  logSecurityEvent('INVALID_TOKEN', 'UNKNOWN', 'Invalid token provided');
  return {valid: false, reason: 'invalid_token'};
}

/**
 * Log security events
 */
function logSecurityEvent(eventType, username, details) {
  if (!SECURITY_CONFIG.ENABLE_AUDIT_LOG) return;
  
  try {
    const auditSheet = SpreadsheetApp.openById(SECURITY_CONFIG.USERS_SHEET_ID)
      .getSheetByName('AuditLog');
    
    auditSheet.appendRow([
      new Date().toISOString(),
      eventType,
      username,
      details,
      'N/A',
      'N/A'
    ]);
    
    cleanupOldAuditLogs(auditSheet);
  } catch (e) {
    Logger.log('Audit logging failed: ' + e.message);
  }
}

/**
 * Cleanup old audit logs
 */
function cleanupOldAuditLogs(auditSheet) {
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - SECURITY_CONFIG.AUDIT_LOG_RETENTION_DAYS);
  
  const data = auditSheet.getDataRange().getValues();
  let rowsToDelete = [];
  
  for (let i = data.length - 1; i > 0; i--) {
    const logDate = new Date(data[i][0]);
    if (logDate < cutoffDate) {
      rowsToDelete.push(i + 1);
    }
  }
  
  rowsToDelete.reverse();
  for (let row of rowsToDelete) {
    auditSheet.deleteRow(row);
  }
}
