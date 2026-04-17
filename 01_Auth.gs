// ===================================================================
// FILE: 01_Auth.gs
// PURPOSE: Zero-Trust security engine. Handles login, session tokens,
//          role verification, password changes, and heartbeat pings.
// FIXES APPLIED:
//   - attemptLogin:      SHA-256 chain replaced with hashPassword()
//   - attemptDemoLogin:  Added _Config gate (DemoModeEnabled) to prevent
//                        unauthenticated role bypass in production
//   - changeMyPassword:  SHA-256 chain replaced with hashPassword()
// ===================================================================

/**
 * INTERCEPTOR: Validates the user's session token before any backend action.
 */
function verifyAccess(token, allowedRoles) {
  if (!token) return { authorized: false, message: "SECURITY: No session token provided." };

  const cache = CacheService.getScriptCache();
  const sessionData = cache.get(token);
  if (!sessionData) return { authorized: false, message: "Session expired. Please log in again." };

  const user = JSON.parse(sessionData);
  let isAuthorized = allowedRoles.includes(user.role);

  // OVERRIDE: Allow 'Adviser/IV' dual-role to pass standard 'Adviser' checks
  if (!isAuthorized && allowedRoles.includes('Adviser') && user.role === 'Adviser/IV') {
    isAuthorized = true;
  }

  if (!isAuthorized) return { authorized: false, message: "SECURITY: Access Denied." };
  return { authorized: true, user: user };
}

/**
 * Authenticates a user against the _Users database.
 * SECURITY: Uses hashPassword() for SHA-256 hashing.
 */
function attemptLogin(email, password) {
  const cleanEmail = String(email || "").trim().toLowerCase();
  const cleanPass  = String(password || "").trim();

  if (!cleanEmail || !cleanPass) {
    return { success: false, message: 'Invalid Credentials' };
  }

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('_Users');
  const data      = userSheet.getDataRange().getDisplayValues();
  const headers   = data[0].map(h => String(h).toLowerCase().trim());
  const col       = getAccountsColMap(headers);

  // FIX: replaced inline digest chain with hashPassword()
  const hashed = hashPassword(cleanPass);

  for (let i = 1; i < data.length; i++) {
    const dbEmail  = col.email    > -1 ? String(data[i][col.email]).trim().toLowerCase() : "";
    const dbPass   = col.password > -1 ? String(data[i][col.password]).trim() : "";
    const dbStatus = col.status   > -1 ? String(data[i][col.status]).trim() : "";

    if (dbEmail === cleanEmail && dbPass === hashed) {
      if (dbStatus === 'Disabled') {
        return { success: false, message: 'Account Disabled. Please contact support.' };
      }

      const sessionToken = Utilities.getUuid();
      const userData = {
        id:    col.id   > -1 ? data[i][col.id]   : '',
        email: dbEmail,
        role:  col.role > -1 ? data[i][col.role] : '',
        name:  col.name > -1 ? data[i][col.name] : ''
      };

      CacheService.getScriptCache().put(sessionToken, JSON.stringify(userData), 14400);

      const dbLastLogin = col.lastLogin > -1 ? String(data[i][col.lastLogin]).trim() : "";

      if (dbLastLogin === "Never" || dbLastLogin === "") {
        return { success: true, ...userData, token: sessionToken, requirePasswordChange: true };
      } else {
        if (col.lastLogin > -1) {
          userSheet.getRange(i + 1, col.lastLogin + 1).setValue(new Date());
        }
        return { success: true, ...userData, token: sessionToken };
      }
    }
  }

  return { success: false, message: 'Invalid Credentials' };
}

/**
 * BYPASS (DEMO ONLY): Logs a user in by role, bypassing passwords.
 * SECURITY FIX: Blocked by _Config gate in production.
 * In your _Config sheet: add row  DemoModeEnabled | TRUE  for dev/demo.
 * Set it to FALSE (or leave blank) before going live.
 */
function attemptDemoLogin(role) {
  // SECURITY GATE — prevents browser-console abuse in production
  if (getConfig('DemoModeEnabled') !== 'TRUE') {
    return { success: false, message: 'Demo mode is not available.' };
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('_Users');
  if (!sheet) return { success: false, message: 'Database Error' };

  const data    = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return { success: false, message: "No active user found." };

  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const col     = getAccountsColMap(headers);

  for (let i = 1; i < data.length; i++) {
    const dbRole   = col.role   > -1 ? String(data[i][col.role]).trim()   : "";
    const dbStatus = col.status > -1 ? String(data[i][col.status]).trim() : "";

    if ((dbRole === role || (role === 'Adviser' && dbRole === 'Adviser/IV')) && dbStatus === 'Active') {
      const sessionToken = Utilities.getUuid();
      const userData = {
        id:    col.id    > -1 ? data[i][col.id]    : '',
        email: col.email > -1 ? data[i][col.email] : '',
        role:  dbRole,
        name:  col.name  > -1 ? data[i][col.name]  : ''
      };

      CacheService.getScriptCache().put(sessionToken, JSON.stringify(userData), 14400);

      if (col.lastLogin > -1) {
        sheet.getRange(i + 1, col.lastLogin + 1).setValue(new Date());
      }

      return { success: true, ...userData, token: sessionToken };
    }
  }
  return { success: false, message: "No active user found." };
}

/**
 * pingSession — Keeps the active session alive (called every 15 min by heartbeat).
 * PERFORMANCE: Uses CacheService only — no Spreadsheet reads.
 */
function pingSession(token) {
  if (!token) return { success: false };
  const cache       = CacheService.getScriptCache();
  const sessionData = cache.get(token);
  if (sessionData) {
    cache.put(token, sessionData, 14400);
    return { success: true };
  }
  return { success: false };
}

/**
 * changeMyPassword — Updates a user's password on first login.
 * FIX: SHA-256 chain replaced with hashPassword().
 */
function changeMyPassword(token, newPassword) {
  const auth = verifyAccess(token, ['Admin', 'Adviser', 'Adviser/IV', 'Student']);
  if (!auth.authorized) return { success: false, message: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_Users');
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const col     = getAccountsColMap(headers);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][col.email]).trim().toLowerCase() === auth.user.email.toLowerCase()) {
        // FIX: replaced inline digest chain with hashPassword()
        sheet.getRange(i + 1, col.password + 1).setValue(hashPassword(newPassword));
        if (col.lastLogin > -1) {
          sheet.getRange(i + 1, col.lastLogin + 1).setValue(new Date());
        }
        return { success: true, message: "Password updated successfully!" };
      }
    }
    return { success: false, message: "User not found." };
  } catch(e) {
    return { success: false, message: "System Error: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}
