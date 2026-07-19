# 🔐 Security Enhancement Guide

## Installation Steps

### Step 1: Create Audit Log Sheet
1. Open your Users Google Sheet
2. Add new sheet called **"AuditLog"**
3. Add headers: `Timestamp | Event | Username | Details | IP | UserAgent`

### Step 2: Update Users Sheet
Users sheet columns:
- A: Username
- B: Password (plain text - auto-hashed)
- C: PasswordHash:Salt
- D: Token
- E: TokenExpiry
- F: (reserved)
- G: FailedAttempts
- H: CreatedAt
- I: LockedOutUntil

### Step 3: Deploy Scripts

**Login Proxy Project:**
1. Copy `Security-Config.gs` → new file
2. Copy `Login Proxy Enhanced.gs` → replace or new file
3. Save & Deploy

**Data Proxy Project:**
1. Copy `Security-Config.gs` → new file
2. Copy `Data Proxy Enhanced.gs` → replace or new file
3. Save & Deploy

### Step 4: Test
```javascript
// In Apps Script console:
logSecurityEvent('SYSTEM_INITIALIZED', 'ADMIN', 'Security system deployed');
```

## Security Features

### ✅ 1. PBKDF2 Hashing
- 100,000 iterations SHA-256
- Unique 16-byte salt per user
- Format: `hash:salt`
- No plain-text passwords

### ✅ 2. Rate Limiting
- Max 5 failed login attempts
- 15-minute lockout
- 20 requests per 5 minutes

### ✅ 3. HttpOnly Cookies
- Token in sessionStorage
- Server-side HttpOnly cookies
- No JavaScript access

### ✅ 4. Session Management
- 30-minute timeout
- Auto token refresh
- Session max duration

### ✅ 5. Audit Logging
- Login/logout events
- Failed attempts
- Token validation
- Data access
- 90-day retention

### ✅ 6. Data Encryption
- Base64 encoding
- Client-side AES-GCM
- Encrypted exports

## Configuration

```javascript
const SECURITY_CONFIG = {
  SESSION_TIMEOUT_MS: 30 * 60 * 1000,
  PBKDF2_ITERATIONS: 100000,
  MAX_LOGIN_ATTEMPTS: 5,
  LOCKOUT_DURATION_MS: 15 * 60 * 1000,
  RATE_LIMIT_MAX_REQUESTS: 20,
  AUDIT_LOG_RETENTION_DAYS: 90
};
```

## Best Practices

### ✅ DO:
- Use HTTPS for all communications
- Rotate passwords every 90 days
- Review audit logs weekly
- Use strong passwords (12+ chars)
- Enable 2FA if available

### ❌ DON'T:
- Store plain-text passwords
- Commit .env to git
- Use weak iterations (< 50,000)
- Share tokens in URLs
- Use default Proxy URLs

## Testing

```bash
# Test rate limiting - 6 rapid attempts
for i in {1..6}; do
  curl -X POST https://your-proxy-url \
    -H 'Content-Type: application/json' \
    -d '{"user": "testuser", "hash": "test"}'
done
```

## Troubleshooting

**Account locked?**
Manually clear in AuditLog sheet:
```javascript
sheet.getRange(row, 9).setValue('');
sheet.getRange(row, 7).setValue(0);
```

**Audit logs not showing?**
Ensure AuditLog sheet exists with correct name.

**Tokens expiring too fast?**
Increase SESSION_TIMEOUT_MS in SECURITY_CONFIG.

---

**Security Level:** 🟢 ENHANCED (Enterprise-Grade)
