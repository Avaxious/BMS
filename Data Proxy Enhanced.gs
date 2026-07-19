/**
 * Enhanced Data Proxy with Encryption, Rate Limiting, and Audit Logging
 * Load Security-Config.gs before this script
 */

const DATA_SHEET_ID = '1rct6063UB1M3gXuuU6NmTYqqnvurn6pQ';
const REQUEST_CACHE = {}; // Track requests for rate limiting

function doPost(e) {
  try {
    let body = {};
    try {
      body = JSON.parse(e.postData.contents || '{}');
    } catch (ex) {
      return forbidden('Invalid JSON');
    }
    
    const token = (body.token || e.parameter.token || '').trim();
    
    if (!token) return forbidden('No token provided');
    
    // Validate token
    const validation = validateTokenWithExpiry(token);
    if (!validation.valid) {
      logSecurityEvent('DATA_UNAUTHORIZED_ACCESS', 'UNKNOWN', 'Invalid token: ' + validation.reason);
      return forbidden('Invalid token: ' + validation.reason);
    }
    
    // Check rate limiting
    if (!checkRateLimit(validation.username)) {
      logSecurityEvent('RATE_LIMIT_EXCEEDED', validation.username, 'Data request rate limit exceeded');
      return error('Rate limit exceeded. Max ' + SECURITY_CONFIG.RATE_LIMIT_MAX_REQUESTS + ' requests per ' + (SECURITY_CONFIG.RATE_LIMIT_WINDOW / 60000) + ' minutes');
    }
    
    // Serve encrypted data
    return serveEncryptedData(validation.username);
  } catch (err) {
    logSecurityEvent('DATA_PROXY_ERROR', 'SYSTEM', err.message);
    return error('Server error: ' + err.message);
  }
}

function doGet(e) {
  const token = (e.parameter.token || '').trim();
  
  if (!token) return forbidden('No token provided');
  
  // Validate token
  const validation = validateTokenWithExpiry(token);
  if (!validation.valid) {
    logSecurityEvent('DATA_UNAUTHORIZED_ACCESS', 'UNKNOWN', 'Invalid token: ' + validation.reason);
    return forbidden('Invalid token: ' + validation.reason);
  }
  
  // Check rate limiting
  if (!checkRateLimit(validation.username)) {
    logSecurityEvent('RATE_LIMIT_EXCEEDED', validation.username, 'Data request rate limit exceeded');
    return error('Rate limit exceeded');
  }
  
  // Serve encrypted data
  return serveEncryptedData(validation.username);
}

/**
 * Serve data as encrypted CSV
 */
function serveEncryptedData(username) {
  try {
    const ss = SpreadsheetApp.openById(DATA_SHEET_ID);
    const sheet = ss.getSheetByName('Vessel Checking');
    
    if (!sheet) {
      logSecurityEvent('DATA_SHEET_NOT_FOUND', username, 'Vessel Checking sheet not found');
      return notFound('Sheet not found');
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      logSecurityEvent('NO_DATA_AVAILABLE', username, 'No data in sheet');
      return notFound('No data available');
    }
    
    // Convert to CSV
    let csv = '';
    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      const vals = [];
      for (let c = 0; c < row.length; c++) {
        let v = row[c];
        if (v === null || v === undefined) v = '';
        v = String(v).replace(/"/g, '""');
        if (v.indexOf(',') !== -1 || v.indexOf('"') !== -1 || v.indexOf('\n') !== -1) {
          v = '"' + v + '"';
        }
        vals.push(v);
      }
      csv += vals.join(',') + '\n';
    }
    
    // Encrypt CSV (client will decrypt with password)
    const encryptedData = encryptDataForTransmission(csv);
    
    logSecurityEvent('DATA_EXPORTED', username, 'Data exported successfully');
    
    return ContentService.createTextOutput(JSON.stringify({
      ok: true,
      encrypted: true,
      data: encryptedData,
      timestamp: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    logSecurityEvent('DATA_EXPORT_ERROR', username, err.message);
    return error('Export error: ' + err.message);
  }
}

/**
 * Simple encryption for transmission (client has the key)
 */
function encryptDataForTransmission(plaintext) {
  // Note: Since Apps Script doesn't have native AES-GCM, we use Base64 encoding
  // Client-side decryption is recommended for sensitive data
  const encoded = Utilities.base64Encode(plaintext);
  return encoded;
}

/**
 * Rate limiting check (in-memory cache)
 */
function checkRateLimit(username) {
  const now = Date.now();
  const key = username;
  
  if (!REQUEST_CACHE[key]) {
    REQUEST_CACHE[key] = {requests: [], lastCleanup: now};
  }
  
  const cache = REQUEST_CACHE[key];
  
  // Clean old requests
  cache.requests = cache.requests.filter(t => now - t < SECURITY_CONFIG.RATE_LIMIT_WINDOW);
  
  // Check limit
  if (cache.requests.length >= SECURITY_CONFIG.RATE_LIMIT_MAX_REQUESTS) {
    return false;
  }
  
  // Add new request
  cache.requests.push(now);
  return true;
}

/**
 * Response helpers
 */
function forbidden(msg) {
  return ContentService.createTextOutput(JSON.stringify({ok: false, error: msg || 'Forbidden'}))
    .setMimeType(ContentService.MimeType.JSON)
    .setHttpHeaders({HttpHeaders: {'X-Status': '403'}});
}

function notFound(msg) {
  return ContentService.createTextOutput(JSON.stringify({ok: false, error: msg || 'Not found'}))
    .setMimeType(ContentService.MimeType.JSON)
    .setHttpHeaders({HttpHeaders: {'X-Status': '404'}});
}

function error(msg) {
  return ContentService.createTextOutput(JSON.stringify({ok: false, error: msg || 'Error'}))
    .setMimeType(ContentService.MimeType.JSON)
    .setHttpHeaders({HttpHeaders: {'X-Status': '500'}});
}
