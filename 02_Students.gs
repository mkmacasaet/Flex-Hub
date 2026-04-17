// ===================================================================
// FILE: 02_Students.gs
// PURPOSE: Master Student Directory — column mappers and all student
//          CRUD operations.
// FIXES APPLIED:
//   - getMasterlistColMap / getAccountsColMap / getLoaColMap:
//       Local `const getIdx` removed; replaced with shared colIdx()
//       from 00_Config.gs
//   - getMasterStudentList, onboardNewStudent, deleteStudentFromMaster:
//       Inline sheet-open replaced with getMasterSheet()
//   - updateMasterStudentDetails: getMasterSheet() used; existing
//       optimistic-locking / concurrency check kept intact
//   - runUniversalNameSync: REMOVED — was dead code (never called)
// ===================================================================

// ===================================================================
// COLUMN MAPPERS
// ===================================================================

/**
 * DYNAMIC COLUMN MAPPER for the Master Student List sheet.
 * FIX: Uses shared colIdx() — local getIdx declaration removed.
 */
function getMasterlistColMap(headers) {
  return {
    id:             colIdx(headers, ["school id", "student id", "id"]),
    lastName:       colIdx(headers, ["last name", "surname"]),
    firstName:      colIdx(headers, ["first name"]),
    middleName:     colIdx(headers, ["middle name"]),
    middleInitial:  colIdx(headers, ["middle initial", "m.i."]),
    nickname:       colIdx(headers, ["nickname", "nick name"]),
    admissionDate:  colIdx(headers, ["admission date", "date of admission"]),
    programme:      colIdx(headers, ["programme & pathway", "programme", "program"]),
    specification:  colIdx(headers, ["specification", "spec"]),
    gender:         colIdx(headers, ["gender", "sex"]),
    personalEmail:  colIdx(headers, ["personal email"]),
    sgenEmail:      colIdx(headers, ["sgen email"]),
    contactNo:      colIdx(headers, ["contact no.", "contact number", "mobile"]),
    landline:       colIdx(headers, ["landline number", "landline"]),
    birthDate:      colIdx(headers, ["birthdate", "birth date", "dob"]),
    age:            colIdx(headers, ["age"]),
    nationality:    colIdx(headers, ["nationality"]),
    currentAddress: colIdx(headers, ["current address", "address"]),
    city:           colIdx(headers, ["city/ province", "city", "province"]),
    country:        colIdx(headers, ["country"]),
    folderLink:     colIdx(headers, ["folder link", "folder", "gdrive"]),
    status:         colIdx(headers, ["enrollment status", "status"]),
    statusRemarks:  colIdx(headers, ["status remarks", "remarks"]),
    specialNeed:    colIdx(headers, ["special needs", "special need"]),
    lastUpdated:    colIdx(headers, ["last updated", "timestamp"]),
    editedBy:       colIdx(headers, ["updated by", "edited by"])
  };
}

/**
 * DYNAMIC COLUMN MAPPER for the _Users sheet.
 * FIX: Uses shared colIdx() — local getIdx declaration removed.
 */
function getAccountsColMap(headers) {
  return {
    id:        colIdx(headers, ["id", "user id"]),
    email:     colIdx(headers, ["email", "sgen email", "user email"]),
    password:  colIdx(headers, ["password", "hash", "pass"]),
    role:      colIdx(headers, ["role", "access", "type"]),
    name:      colIdx(headers, ["name", "full name"]),
    status:    colIdx(headers, ["status", "account status"]),
    lastLogin: colIdx(headers, ["last login", "timestamp"])
  };
}

/**
 * DYNAMIC COLUMN MAPPER for the LOA & Withdrawals sheet.
 * FIX: Uses shared colIdx() — local getIdx declaration removed.
 */
function getLoaColMap(headers) {
  return {
    recordId:       colIdx(headers, ["record id", "loa id"]),
    studentName:    colIdx(headers, ["student name", "name"]),
    studentId:      colIdx(headers, ["school id", "student id", "id"]),
    program:        colIdx(headers, ["program", "programme", "pathway"]),
    separationType: colIdx(headers, ["separation type", "type"]),
    effectiveDate:  colIdx(headers, ["effective date", "effective"]),
    returnDate:     colIdx(headers, ["expected return date", "return date"]),
    reason:         colIdx(headers, ["reason"]),
    attachment:     colIdx(headers, ["attachment", "document", "file"]),
    loaStatus:      colIdx(headers, ["loa status", "status"]),
    lastUpdated:    colIdx(headers, ["last updated", "timestamp", "date encoded"]),
    updatedBy:      colIdx(headers, ["updated by", "encoded by"])
  };
}

// ===================================================================
// STUDENT CRUD OPERATIONS
// ===================================================================

/**
 * Fetches the entire Master Student List.
 * FIX: Replaced inline sheet-open with getMasterSheet().
 */
function getMasterStudentList(token) {
  const auth = verifyAccess(token, ['Admin', 'Adviser', 'Adviser/IV']);
  if (!auth.authorized) return { error: auth.message };

  try {
    const sheet = getMasterSheet();
    const data  = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const col     = getMasterlistColMap(headers);

    return data.slice(1).map((r, i) => {
      const val = (idx) => (idx > -1 && r[idx] !== undefined && r[idx] !== null) ? String(r[idx]).trim() : "";
      const displayName = formatUniversalName(val(col.firstName), val(col.middleInitial), val(col.lastName));

      return {
        rowIndex: i + 2,
        id: val(col.id), lastName: val(col.lastName), firstName: val(col.firstName),
        middleName: val(col.middleName), middleInitial: val(col.middleInitial),
        nickname: val(col.nickname), admissionDate: val(col.admissionDate),
        programme: val(col.programme), specification: val(col.specification),
        gender: val(col.gender), personalEmail: val(col.personalEmail),
        sgenEmail: val(col.sgenEmail), contactNo: val(col.contactNo),
        landline: val(col.landline), birthDate: val(col.birthDate),
        age: val(col.age), nationality: val(col.nationality),
        currentAddress: val(col.currentAddress), city: val(col.city),
        country: val(col.country), folderLink: val(col.folderLink),
        status: val(col.status), statusRemarks: val(col.statusRemarks),
        specialNeed: val(col.specialNeed), lastUpdated: val(col.lastUpdated),
        editedBy: val(col.editedBy), fullName: displayName
      };
    }).filter(s => s.id !== "");
  } catch (e) { return { error: "Connection Error: " + e.toString() }; }
}

/**
 * Enrolls a new student and provisions their Hub login account.
 * FIX: Replaced inline sheet-open with getMasterSheet().
 */
function onboardNewStudent(token, form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const auth = verifyAccess(token, ['Admin']);
    if (!auth.authorized) return { success: false, message: auth.message };

    const sheet   = getMasterSheet();
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const col     = getMasterlistColMap(headers);

    if (col.id === -1 || col.lastName === -1) {
      return { success: false, message: "CRITICAL: Cannot find Headers in Row 1." };
    }

    if (form.id) form.sgenEmail = String(form.id).trim() + "@sgen.edu.ph";

    if (form.birthDate) {
      const dob  = new Date(form.birthDate);
      const today = new Date();
      let calculatedAge = today.getFullYear() - dob.getFullYear();
      const m = today.getMonth() - dob.getMonth();
      if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) calculatedAge--;
      form.age = calculatedAge;
    }

    let newRow = new Array(headers.length).fill("");
    for (const key in col) {
      const colIndex = col[key];
      if (colIndex > -1) {
        if (key === 'status') newRow[colIndex] = "Active";
        else if (form[key] !== undefined) newRow[colIndex] = form[key];
      }
    }

    if (col.lastUpdated > -1) newRow[col.lastUpdated] = new Date();
    if (col.editedBy   > -1) newRow[col.editedBy]    = "Enrolled by: " + auth.user.name;

    sheet.appendRow(newRow);

    const universalName = formatUniversalName(form.firstName, form.middleInitial, form.lastName);
    const accountStatus = createNewAccount(token, {
      name: universalName, email: form.sgenEmail,
      password: form.password, role: "Student"
    });

    if (accountStatus && accountStatus.duplicate) {
      return { success: true, message: "Student Enrolled! (Note: Account creation skipped — email already exists)." };
    }

    return { success: true, message: "Student Enrolled and Hub Account Created!" };
  } catch (err) {
    Logger.log("Onboarding Error: " + err.toString());
    return { success: false, message: "System Error: " + err.message };
  } finally { lock.releaseLock(); }
}

/**
 * Surgically updates a student's profile.
 * Concurrency / optimistic-locking check is kept intact.
 * FIX: Replaced inline sheet-open with getMasterSheet().
 */
function updateMasterStudentDetails(token, payload) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getMasterSheet();
    const data  = sheet.getDataRange().getValues();

    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const col     = getMasterlistColMap(headers);
    if (col.id === -1) return { success: false, message: "CRITICAL: ID column missing in Masterlist." };

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][col.id]) === String(payload.id)) { rowIndex = i + 1; break; }
    }
    if (rowIndex === -1) return { success: false, message: "Student ID not found." };

    const rowRange = sheet.getRange(rowIndex, 1, 1, Math.max(sheet.getLastColumn(), headers.length));
    const rowData  = rowRange.getValues()[0];

    // CONCURRENCY CHECK (optimistic locking — already existed, kept intact)
    if (payload.lastUpdated && col.lastUpdated > -1 && rowData[col.lastUpdated]) {
      const clientTime = new Date(payload.lastUpdated).getTime();
      const sheetTime  = new Date(rowData[col.lastUpdated]).getTime();
      if (!isNaN(clientTime) && !isNaN(sheetTime) && Math.abs(clientTime - sheetTime) > 60000) {
        return { success: false, message: "❌ SAVE REJECTED: Another Admin updated this profile." };
      }
    }

    const fieldMap = {
      'e-lname': col.lastName, 'e-fname': col.firstName, 'e-mname': col.middleName,
      'e-mi': col.middleInitial, 'e-nick': col.nickname, 'e-admission': col.admissionDate,
      'e-programme': col.programme, 'e-specification': col.specification, 'e-gender': col.gender,
      'e-pers-email': col.personalEmail, 'e-sgen-email': col.sgenEmail, 'e-contact': col.contactNo,
      'e-landline': col.landline, 'e-dob': col.birthDate, 'e-age': col.age,
      'e-nation': col.nationality, 'e-address': col.currentAddress, 'e-city': col.city,
      'e-country': col.country, 'e-fold-link': col.folderLink, 'e-enrollmentStatus': col.status,
      'e-stat-remarks': col.statusRemarks, 'e-special': col.specialNeed,
      'status': col.status, 'lastName': col.lastName, 'firstName': col.firstName,
      'programme': col.programme, 'specification': col.specification, 'sgenEmail': col.sgenEmail,
      'specialNeed': col.specialNeed, 'city': col.city, 'country': col.country,
      'gender': col.gender, 'age': col.age
    };

    let fieldsUpdated = 0;
    let keysReceived  = [];
    const updatesToApply = payload.updates || payload;

    for (const key in updatesToApply) {
      keysReceived.push(key);
      if (fieldMap[key] !== undefined && fieldMap[key] > -1) {
        rowData[fieldMap[key]] = updatesToApply[key];
        fieldsUpdated++;
      }
    }

    if (fieldsUpdated === 0) {
      return { success: false, message: `No matching fields found. Received: [${keysReceived.join(', ')}]` };
    }

    if (col.lastUpdated > -1) rowData[col.lastUpdated] = new Date();
    if (col.editedBy   > -1) rowData[col.editedBy]    = "Edited by: " + (payload.editedBy || "Unknown Admin");

    rowRange.setValues([rowData]);
    return { success: true, message: `Student details surgically updated! (${fieldsUpdated} fields changed)` };

  } catch (e) {
    return { success: false, message: "System busy. Please try again. " + e.toString() };
  } finally { lock.releaseLock(); }
}

/**
 * Permanently removes a student from the Master Directory.
 * FIX: Replaced inline sheet-open with getMasterSheet().
 */
function deleteStudentFromMaster(token, studentId) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  const sheet = getMasterSheet();
  const data  = sheet.getDataRange().getValues();

  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const col     = getMasterlistColMap(headers);
  if (col.id === -1) return { success: false, message: "CRITICAL: Cannot find ID column!" };

  let accountStatus = "";
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col.id]) === String(studentId)) {
      const emailToDelete = (col.sgenEmail > -1) ? data[i][col.sgenEmail] : null;
      sheet.deleteRow(i + 1);
      if (emailToDelete) {
        const deleted = autoDeleteAccountByEmail(emailToDelete);
        accountStatus = deleted ? " and associated Hub Account." : " (Note: No matching Hub account found).";
      }
      return { success: true, message: "Student record deleted" + accountStatus };
    }
  }
  return { success: false, message: "Student ID not found." };
}

/**
 * HELPER: Searches _Users for an email and deletes that row.
 */
function autoDeleteAccountByEmail(email) {
  if (!email) return false;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_Users');
  if (!sheet) return false;

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const col     = getAccountsColMap(headers);
  if (col.email === -1) return false;

  let found = false;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][col.email]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      sheet.deleteRow(i + 1);
      found = true;
    }
  }
  return found;
}
