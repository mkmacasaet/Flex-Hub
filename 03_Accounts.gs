// ===================================================================
// FILE: 03_Accounts.gs
// PURPOSE: Manages Hub user accounts (staff and students).
// FIXES APPLIED:
//   - updateAccountAdmin: SHA-256 chain replaced with hashPassword()
//   - createNewAccount:   SHA-256 chain replaced with hashPassword()
//   - runUniversalNameSync: NOT included — dead code, never called
// ===================================================================

/**
 * Retrieves all registered Hub users for the Admin table.
 */
function getAllUsersAdmin(token) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('_Users');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1)
    .map((r, i) => ({ rowIndex: i + 2, id: r[0], email: r[1], role: r[3], name: r[4], status: r[5] }))
    .filter(u => u.id !== "");
}

/**
 * Updates a user's role, name, email, or status.
 * FIX: SHA-256 chain replaced with hashPassword().
 */
function updateAccountAdmin(token, data) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheet   = ss.getSheetByName('_Users');
    const rowRange = sheet.getRange(data.rowIndex, 1, 1, sheet.getLastColumn());
    const rowData  = rowRange.getValues()[0];

    if (data.email)  rowData[1] = data.email;
    // FIX: replaced inline digest chain with hashPassword()
    if (data.password) rowData[2] = hashPassword(data.password);
    if (data.role)   rowData[3] = data.role;
    if (data.name)   rowData[4] = data.name;
    if (data.status) rowData[5] = data.status;

    rowRange.setValues([rowData]);
    return { success: true, message: "Account updated successfully!" };
  } catch (e) {
    return { success: false, message: "System error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Permanently deletes a user account from the Hub.
 */
function deleteAccountAdmin(token, rowIndex) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_Users').deleteRow(rowIndex);
    return { success: true, message: "Account successfully deleted." };
  } catch(e) {
    return { success: false, message: "Error deleting account: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Provisions a new Hub account.
 * Automatically called by onboardNewStudent() for new students.
 * FIX: SHA-256 chain replaced with hashPassword().
 */
function createNewAccount(token, data) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { success: false, message: auth.message };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_Users');
  if (!sheet) return { success: false, message: "Users sheet not found." };

  const dataRange = sheet.getDataRange().getValues();
  const headers   = dataRange[0].map(h => String(h).toLowerCase().trim());
  const col       = getAccountsColMap(headers);

  if (col.email > -1) {
    const existingEmails = dataRange.map(r => String(r[col.email]).toLowerCase().trim());
    if (existingEmails.includes(String(data.email).toLowerCase().trim())) {
      return { success: false, duplicate: true, message: "Account already exists." };
    }
  }

  const prefix    = data.role === "Admin" ? "ADM-" : "STU-";
  const newId     = prefix + Math.floor(Math.random() * 9000 + 1000);
  // FIX: replaced inline digest chain with hashPassword()
  const hashedPw  = hashPassword(data.password);

  let newRow = new Array(headers.length).fill("");
  if (col.id       > -1) newRow[col.id]       = newId;
  if (col.email    > -1) newRow[col.email]    = data.email;
  if (col.password > -1) newRow[col.password] = hashedPw;
  if (col.role     > -1) newRow[col.role]     = data.role;
  if (col.name     > -1) newRow[col.name]     = data.name;
  if (col.status   > -1) newRow[col.status]   = "Active";
  if (col.lastLogin > -1) newRow[col.lastLogin] = "Never";

  sheet.appendRow(newRow);
  return { success: true, duplicate: false, message: `Account created for ${data.name}` };
}
