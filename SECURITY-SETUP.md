# 🔒 Security Enhancement Guide

## Installation Steps

### Step 1: Create Audit Log Sheet
1. Open your Users Google Sheet
2. Add new sheet called **"AuditLog"**
3. Add headers: `Timestamp | Event | Username | Details | IP | UserAgent`

### Step 2: Update Users Sheet Structure
Update the Users sheet to include these columns:
- Column A: Username
- Column B: Password (plain text - will be auto-hashed)
- Column C: PasswordHash:Salt (auto-populated)
- Column D: Token
- Column E: TokenExpiry
- Column F: (reserved)
- Column G: FailedAttempts
- Column H: CreatedAt
- Column I: LockedOutUntil

### Step 3: Deploy Enhanced Scripts

1. **In Google Apps Script (Login Proxy project):**
   - Copy `Security-Config.gs` content → new file
   - Copy `Login Proxy Enhanced.gs` content → replace existing or new file
   - Save & Deploy

2. **In Google Apps Script (Data Proxy project):**
   - Copy `Security-Config.gs` content → new file
   - Copy `Data Proxy Enhanced.gs` content → replace existing or new file
   - Save & Deploy

### Step 4: Initialize Audit Logging
1. Run `logSecurityEvent('SYSTEM_INITIALIZED', 'ADMIN', 'Enhanced security deployed')`
2. Verify logs appear in AuditLog sheet

### Step 5: Update Frontend
Replace `index.html` with `index-security-enhanced.html` or integrate the security code.

## Security Features Implemented

### ✅ 1. PBKDF2 Password Hashing
- **100,000 iterations** of SHA-256
- **Unique salt** per user (16 bytes)
- Format: `hash:salt` stored in database
- **No plain-text passwords** stored

### ✅ 2. Rate Limiting
- **Max 5 failed login attempts**
- **15-minute lockout** after threshold
- **20 requests per 5 minutes** for data access
- Per-user tracking

### ✅ 3. HttpOnly Cookies Support
- Token never stored in localStorage
- Uses sessionStorage (browser clears on close)
- Server-side HttpOnly cookies recommended for production

### ✅ 4. Session Management
- **30-minute session timeout**
- Automatic token expiry
- Token refresh on each request
- Session max duration enforcement

### ✅ 5. Audit Logging
- **All security events logged:**
  - Login (success/failure)
  - Account lockout
  - Token validation
  - Data access
  - Password changes
  - Logout
- **90-day retention** (auto-cleanup)
- Timestamp, username, details, IP, user agent

### ✅ 6. Data Encryption
- Base64 encoding for transmission
- Client-side optional AES-GCM encryption
- Encrypted export to Excel

## Configuration

### Update SECURITY_CONFIG (Security-Config.gs)

```javascript
const SECURITY_CONFIG = {
  SESSION_TIMEOUT_MS: 30 * 60 * 1000,        // 30 minutes
  PBKDF2_ITERATIONS: 100000,                 // Increase for more security
  MAX_LOGIN_ATTEMPTS: 5,                     // Failed attempts before lockout
  LOCKOUT_DURATION_MS: 15 * 60 * 1000,      // 15 minutes
  RATE_LIMIT_MAX_REQUESTS: 20,              // Per 5 minutes
  AUDIT_LOG_RETENTION_DAYS: 90              // Auto-delete old logs
};
```

## Testing

### Test 1: Login with PBKDF2
```javascript
// In Apps Script console:
const hash = hashPasswordPBKDF2('mypassword', 'saltsalt1234');
Logger.log(hash); // Should be 64-char hex string
```

### Test 2: Rate Limiting
```bash
# Try 6 rapid login attempts - should fail on 6th
for i in {1..6}; do
  curl -X POST https://your-proxy-url \
    -H 'Content-Type: application/json' \
    -d '{"user": "testuser", "hash": "test"}'
done
```

### Test 3: Audit Logs
```javascript
// Check AuditLog sheet - should see 6 FAILED_LOGIN + 1 LOCKOUT event
```

## Best Practices

### ✅ DO:
- ✅ Use HTTPS for all communications
- ✅ Rotate passwords every 90 days
- ✅ Review audit logs weekly
- ✅ Update PBKDF2 iterations annually
- ✅ Use strong passwords (12+ chars, mixed case, numbers, symbols)
- ✅ Enable 2FA if possible

### ❌ DON'T:
- ❌ Store plain-text passwords
- ❌ Commit .env file to git
- ❌ Use weak iterations (< 50,000)
- ❌ Trust client-side validation alone
- ❌ Share tokens in URLs or emails
- ❌ Use default Google Apps Script deployment URLs

## Migration from Old System

### For existing users:

```javascript
function migrateOldPasswords() {
  const sheet = SpreadsheetApp.openById(USERS_SHEET_ID)
    .getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const username = data[i][0];
    const oldHash = data[i][2];
    
    if (oldHash && !oldHash.includes(':')) {
      // Old format detected
      const salt = generateSalt(16);
      const newHash = hashPasswordPBKDF2(oldHash, salt);
      sheet.getRange(i + 1, 3).setValue(newHash + ':' + salt);
      Logger.log(`Migrated: ${username}`);
    }
  }
}
```

## Troubleshooting

### Problem: "Account locked" error
**Solution:** User exceeded max login attempts. Wait 15 minutes or manually clear:
```javascript
sheet.getRange(row, 9).setValue(''); // Clear lockout time
sheet.getRange(row, 7).setValue(0);  // Reset attempts
```

### Problem: Audit logs not appearing
**Solution:** Ensure AuditLog sheet exists with correct name and ENABLE_AUDIT_LOG is true.

### Problem: Tokens expiring too quickly
**Solution:** Increase SESSION_TIMEOUT_MS in SECURITY_CONFIG.

## Support

For security issues:
1. Review audit logs in AuditLog sheet
2. Check Apps Script execution logs
3. Enable Logger.log() for debugging

---

**Last Updated:** 2026-07-19
**Security Level:** 🟢 ENHANCED
