// ===================================================================
// FILE: 10_LOA.gs
// PURPOSE: Leave of Absence and Withdrawal management.
// FIXES APPLIED:
//   - All functions: Replaced inline LOA sheet open with getLOASheet()
//   - getMasterSheet() used wherever Master List is opened
// ===================================================================

function getLOAModalData(token) {
  try {
    const auth = verifyAccess(token, ['Admin']);
    if (!auth.authorized) return { error: auth.message };

    let payload = { students: [], types: [] };

    const sheet = getMasterSheet();
    const data  = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const idCol     = headers.findIndex(h => h.includes("school id") || h.includes("student id") || h === "id");
    const progCol   = headers.findIndex(h => h.includes("programme") || h.includes("program") || h.includes("pathway"));
    const gdriveCol = headers.findIndex(h => h.includes("folder link") || h.includes("gdrive"));
    const statusCol = headers.findIndex(h => h.includes("status"));

    if (statusCol === -1) return { error: "CRITICAL: Could not find 'Status' column in the Masterlist!" };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][statusCol]).trim() === "Active") {
        const fullName = formatUniversalName(data[i][2], data[i][4], data[i][1]);
        payload.students.push({
          id:      idCol     !== -1 ? data[i][idCol]     : '',
          name:    fullName,
          program: progCol   !== -1 ? data[i][progCol]   : '',
          gdrive:  gdriveCol !== -1 ? data[i][gdriveCol] : ''
        });
      }
    }
    payload.students.sort((a, b) => a.name.localeCompare(b.name));

    const confSS = SpreadsheetApp.getActiveSpreadsheet();
    const ddSheet = confSS.getSheetByName('_Dropdowns');
    if (!ddSheet) return { error: "CRITICAL: Could not find the '_Dropdowns' sheet." };

    const ddData   = ddSheet.getDataRange().getDisplayValues();
    const headersDD = ddData[0].map(h => String(h).toLowerCase().trim());
    const typeCol    = headersDD.findIndex(h => h.includes('enrollment status') || h === 'status');
    const loaStatusCol = headersDD.findIndex(h => h === 'loa status');

    if (typeCol > -1) {
      payload.types = [...new Set(ddData.slice(1).map(r => String(r[typeCol]).trim()).filter(String))];
    } else { return { error: "CRITICAL: Could not find 'Enrollment Status' column in _Dropdowns!" }; }

    payload.loaStatuses = loaStatusCol > -1
      ? [...new Set(ddData.slice(1).map(r => String(r[loaStatusCol]).trim()).filter(String))]
      : ['Active', 'Pending Approval', 'Returned'];

    return payload;
  } catch (e) { return { error: "Backend Crash: " + e.toString() }; }
}

function saveLOARecord(token, payload) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { success: false, message: auth.message };

  let attachmentUrl = "";
  if (payload.attachmentBase64) {
    try {
      const folderId = getConfig('LOA/WithdrawalFolderID');
      if (!folderId) return { success: false, message: 'LOA/WithdrawalFolderID missing in Config.' };
      const blob = Utilities.newBlob(Utilities.base64Decode(payload.attachmentBase64), payload.attachmentMimeType, payload.attachmentName);
      attachmentUrl = DriveApp.getFolderById(folderId).createFile(blob).getUrl();
    } catch(e) { return { success: false, message: 'File Upload Failed: ' + e.toString() }; }
  }

  // FIX: replaced inline open with getLOASheet()
  const loaSheet = getLOASheet();
  const data     = loaSheet.getDataRange().getValues();
  const headers  = data[0].map(h => String(h).trim().toLowerCase());
  const loaMap   = getLoaColMap(headers);

  const recordId  = 'LOA-' + Math.floor(100000 + Math.random() * 900000);
  const timestamp = new Date();
  let rowData     = new Array(headers.length).fill('');

  if (loaMap.recordId      > -1) rowData[loaMap.recordId]       = recordId;
  if (loaMap.studentName   > -1) rowData[loaMap.studentName]    = payload.name;
  if (loaMap.studentId     > -1) rowData[loaMap.studentId]      = payload.studentId;
  if (loaMap.program       > -1) rowData[loaMap.program]        = payload.program;
  if (loaMap.separationType > -1) rowData[loaMap.separationType] = payload.type;
  if (loaMap.effectiveDate > -1) rowData[loaMap.effectiveDate]  = payload.effectiveDate;
  if (loaMap.returnDate    > -1) rowData[loaMap.returnDate]     = payload.returnDate || 'N/A';
  if (loaMap.reason        > -1) rowData[loaMap.reason]         = payload.reason;
  if (loaMap.attachment    > -1) rowData[loaMap.attachment]     = attachmentUrl;
  if (loaMap.loaStatus     > -1) rowData[loaMap.loaStatus]      = payload.loaStatus || 'Active';
  if (loaMap.lastUpdated   > -1) rowData[loaMap.lastUpdated]    = timestamp;
  if (loaMap.updatedBy     > -1) rowData[loaMap.updatedBy]      = auth.user.name;

  loaSheet.appendRow(rowData);

  const masterSheet = getMasterSheet();
  const mData       = masterSheet.getDataRange().getDisplayValues();
  const mHeaders    = mData[0].map(h => String(h).trim().toLowerCase());
  const masterMap   = getMasterlistColMap(mHeaders);

  let targetStatus = payload.type;
  if (payload.loaStatus === 'LOA Completed' || payload.loaStatus === 'Returned') targetStatus = 'Active';

  if (masterMap.id !== -1 && masterMap.status !== -1) {
    let studentFound = false;
    for (let i = 1; i < mData.length; i++) {
      if (String(mData[i][masterMap.id]).trim() === String(payload.studentId).trim()) {
        masterSheet.getRange(i + 1, masterMap.status + 1).setValue(targetStatus);
        studentFound = true; break;
      }
    }
    if (!studentFound) return { success: false, message: 'LOA Saved, but Student ID was not found in Masterlist.' };
  }

  return { success: true, message: 'Record saved, File Uploaded, and Masterlist status updated!' };
}

function getLOARecords(token) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  // FIX: replaced inline open with getLOASheet()
  const sheet = getLOASheet();
  if (sheet.getLastRow() < 2) return [];

  const data    = sheet.getDataRange().getDisplayValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const loaMap  = getLoaColMap(headers);

  let records = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    records.push({
      id:         loaMap.recordId       > -1 ? data[i][loaMap.recordId]       : '',
      studentId:  loaMap.studentId      > -1 ? data[i][loaMap.studentId]      : '',
      date:       loaMap.lastUpdated    > -1 ? data[i][loaMap.lastUpdated]    : '',
      name:       loaMap.studentName    > -1 ? data[i][loaMap.studentName]    : '',
      program:    loaMap.program        > -1 ? data[i][loaMap.program]        : 'N/A',
      type:       loaMap.separationType > -1 ? data[i][loaMap.separationType] : '',
      effective:  loaMap.effectiveDate  > -1 ? data[i][loaMap.effectiveDate]  : '',
      returnDate: loaMap.returnDate     > -1 ? data[i][loaMap.returnDate]     : '',
      reason:     loaMap.reason         > -1 ? data[i][loaMap.reason]         : '',
      attachment: loaMap.attachment     > -1 ? data[i][loaMap.attachment]     : '',
      loaStatus:  loaMap.loaStatus      > -1 ? data[i][loaMap.loaStatus]      : '',
      updatedBy:  loaMap.updatedBy      > -1 ? data[i][loaMap.updatedBy]      : ''
    });
  }
  return records.reverse();
}

function updateLOARecord(token, payload) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { success: false, message: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    // FIX: replaced inline open with getLOASheet()
    const sheet   = getLOASheet();
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const loaMap  = getLoaColMap(headers);

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][loaMap.recordId]) === String(payload.id)) { rowIndex = i + 1; break; }
    }
    if (rowIndex === -1) return { success: false, message: "Record not found." };

    const rowRange = sheet.getRange(rowIndex, 1, 1, Math.max(sheet.getLastColumn(), headers.length));
    const rowData  = rowRange.getValues()[0];

    let fieldsUpdated = 0, syncMasterlist = false;

    for (const key in payload.updates) {
      if (loaMap[key] !== undefined && loaMap[key] > -1) {
        rowData[loaMap[key]] = payload.updates[key];
        fieldsUpdated++;
        if (key === 'separationType' || key === 'loaStatus') syncMasterlist = true;
      }
    }

    if (payload.attachmentBase64) {
      const folderId = getConfig('LOA/WithdrawalFolderID');
      const folder   = DriveApp.getFolderById(folderId);
      const blob     = Utilities.newBlob(Utilities.base64Decode(payload.attachmentBase64), payload.attachmentMimeType, payload.attachmentName);
      rowData[loaMap.attachment] = folder.createFile(blob).getUrl();
      fieldsUpdated++;
    }

    if (fieldsUpdated === 0) return { success: false, message: "No changes detected." };

    if (loaMap.lastUpdated > -1) rowData[loaMap.lastUpdated] = new Date();
    if (loaMap.updatedBy   > -1) rowData[loaMap.updatedBy]   = auth.user.name;

    rowRange.setValues([rowData]);

    if (syncMasterlist && rowData[loaMap.studentId]) {
      const mSheet  = getMasterSheet();
      const mData   = mSheet.getDataRange().getDisplayValues();
      const mHeaders = mData[0].map(h => String(h).trim().toLowerCase());
      const mIdCol  = mHeaders.findIndex(h => h.includes("school id") || h.includes("student id") || h === "id");
      const mStatCol = mHeaders.findIndex(h => h.includes("status"));

      let targetStatus = rowData[loaMap.separationType];
      const currentLoaStatus = String(rowData[loaMap.loaStatus]).trim();
      if (currentLoaStatus === 'LOA Completed' || currentLoaStatus === 'Returned') targetStatus = 'Active';

      if (mIdCol !== -1 && mStatCol !== -1) {
        for (let i = 1; i < mData.length; i++) {
          if (String(mData[i][mIdCol]).trim() === String(rowData[loaMap.studentId]).trim()) {
            mSheet.getRange(i + 1, mStatCol + 1).setValue(targetStatus); break;
          }
        }
      }
    }

    return { success: true, message: `Record surgically updated! (${fieldsUpdated} fields)` };
  } catch (e) {
    return { success: false, message: "System Error: " + e.toString() };
  } finally { lock.releaseLock(); }
}

function deleteLOARecord(token, recordId) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { success: false, message: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    // FIX: replaced inline open with getLOASheet()
    const sheet   = getLOASheet();
    const data    = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const loaMap  = getLoaColMap(headers);

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][loaMap.recordId]) === String(recordId)) { rowIndex = i + 1; break; }
    }
    if (rowIndex === -1) return { success: false, message: "Record not found." };

    sheet.deleteRow(rowIndex);
    return { success: true, message: "Record deleted permanently." };
  } catch (e) {
    return { success: false, message: "System Error: " + e.toString() };
  } finally { lock.releaseLock(); }
}
