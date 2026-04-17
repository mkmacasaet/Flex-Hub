// ===================================================================
// FILE: 05_Deferrals.gs
// PURPOSE: Deferral Engine — handles student deadline extension requests.
// FIXES APPLIED:
//   - All functions: Replaced repeated getConfig/openById/_Deferrals
//     boilerplate with getDeferralSheet() from 00_Config.gs
// ===================================================================

function getEligibleDeferralUnits(token, studentName) {
  const auth = verifyAccess(token, ['Admin', 'Student']);
  if (!auth.authorized) return { error: auth.message };

  try {
    let targetAkIds = [];

    if (auth.user.role === 'Student') {
      const studentId = String(auth.user.email).toLowerCase().replace('@sgen.edu.ph', '').trim();
      const indexSheet = getStudentIndexSheet();
      const indexData  = indexSheet.getDataRange().getValues();
      for (let i = 1; i < indexData.length; i++) {
        if (String(indexData[i][0]).trim() === studentId) {
          targetAkIds.push(String(indexData[i][3]).trim());
        }
      }
    }

    const akId = getConfig('AccessKitGeneratorID');
    if (!akId) return [];

    const akSheet = SpreadsheetApp.openById(akId).getSheets()[0];
    const akData  = akSheet.getDataRange().getDisplayValues();
    let eligibleUnits = [];

    for (let i = 1; i < akData.length; i++) {
      const rowAkId       = String(akData[i][0]).trim();
      const rowStudentName = String(akData[i][1]).trim().toLowerCase();
      const isMatch = (auth.user.role === 'Student')
        ? targetAkIds.includes(rowAkId)
        : rowStudentName === String(studentName).trim().toLowerCase();

      if (isMatch) {
        const status = String(akData[i][29] || '').toLowerCase();
        if (status.includes('complete') || status.includes('drop')) continue;
        eligibleUnits.push({
          akId: rowAkId, unitName: akData[i][8],
          adviser: akData[i][10], finalDeadline: akData[i][14]
        });
      }
    }
    return eligibleUnits;
  } catch(e) { return { error: e.toString() }; }
}

function getAdviserDeferrals(userToken, requestedName, requestedEmail) {
  const auth = verifyAccess(userToken, ['Adviser', 'Adviser/IV', 'Admin']);
  if (!auth.authorized) return { error: auth.message };

  const adviserName = auth.user.role.includes('Adviser') ? auth.user.name : requestedName;

  try {
    // FIX: replaced 3-line open pattern with getDeferralSheet()
    const sheet = getDeferralSheet();
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getDisplayValues();
    return data.map((r, i) => ({
      rowIndex: i + 1, defId: r[0], dateReq: r[1], name: r[2], unit: r[3], adviser: r[4],
      origDeadline: r[5], extDays: r[6], reason: r[7], docLink: r[8],
      status: r[9] || 'Pending', approvedDate: r[10], remarks: r[11], email: r[12]
    })).slice(1)
    .filter(row => {
      const matchAdv = String(row.adviser).trim().toLowerCase() === String(adviserName).trim().toLowerCase();
      return matchAdv && row.status !== 'Deleted';
    }).reverse();
  } catch(e) { return []; }
}

function updateDeferralAdmin(token, payload) {
  const auth = verifyAccess(token, ['Admin', 'Adviser', 'Adviser/IV']);
  if (!auth.authorized) return { error: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    // FIX: replaced 3-line open pattern with getDeferralSheet()
    const sheet = getDeferralSheet();

    let newFileUrl = null;
    if (payload.fileData && payload.fileName) {
      const folder = DriveApp.getFolderById(getConfig('DeferralDocsID'));
      const file   = folder.createFile(Utilities.newBlob(
        Utilities.base64Decode(payload.fileData), payload.mimeType,
        payload.defId + "_" + payload.fileName
      ));
      newFileUrl = file.getUrl();
    }

    if (payload.rowIndex) {
      const rowIdx   = parseInt(payload.rowIndex);
      const rowRange = sheet.getRange(rowIdx, 1, 1, 13);
      const rowData  = rowRange.getValues()[0];

      if (String(rowData[0]).trim() !== String(payload.defId).trim()) {
        return { success: false, message: "CRITICAL ERROR: Row Mismatch!" };
      }

      const fieldMap = {
        'def-name': 2, 'def-unit': 3, 'def-adviser': 4, 'def-origDeadline': 5,
        'def-extDays': 6, 'def-reason': 7, 'def-status': 9, 'def-approvedDate': 10,
        'def-remarks': 11, 'def-email': 12
      };

      let updatedKeys = [];
      if (payload.updates) {
        for (const key in payload.updates) {
          if (fieldMap[key] !== undefined) {
            rowData[fieldMap[key]] = payload.updates[key];
            updatedKeys.push(key);
          }
        }
      }
      if (newFileUrl) { rowData[8] = newFileUrl; updatedKeys.push("Document Attached"); }
      rowRange.setValues([rowData]);
      return { success: true, message: "Successfully updated: " + updatedKeys.join(", ") };

    } else {
      const d = payload.fullData;
      const fileUrlToSave = newFileUrl || payload.existingDoc || "";
      const rowData = [
        payload.defId, payload.dateReq, d['def-name'], d['def-unit'], d['def-adviser'],
        d['def-origDeadline'], d['def-extDays'], d['def-reason'], fileUrlToSave,
        d['def-status'], d['def-approvedDate'], d['def-remarks'], d['def-email'] || ""
      ];
      sheet.appendRow(rowData);
      return { success: true, message: "New deferral logged successfully!" };
    }
  } catch(e) { return { success: false, message: "Backend Error: " + e.toString() }; }
  finally { lock.releaseLock(); }
}

function getDeferralsAdmin(token) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  try {
    // FIX: replaced 3-line open pattern with getDeferralSheet()
    const sheet = getDeferralSheet();
    if (!sheet || sheet.getLastRow() < 2) {
      return { stats: { pending: 0, approved: 0, rejected: 0, avgDays: "0 Days" }, records: [] };
    }

    const data = sheet.getDataRange().getDisplayValues();
    let pending = 0, approved = 0, rejected = 0, totalExtDays = 0;

    const records = data.slice(1).map((r, i) => {
      const status  = r[9] || 'Pending';
      const extDays = parseInt(r[6]) || 0;
      if (status !== 'Deleted') {
        if (status === 'Pending')      pending++;
        if (status === 'Approved')     { approved++; totalExtDays += extDays; }
        if (status === 'Not Approved') rejected++;
      }
      return {
        rowIndex: i + 2, defId: r[0], dateReq: r[1], name: r[2], unit: r[3], adviser: r[4],
        origDeadline: r[5], extDays: r[6], reason: r[7], docLink: r[8],
        status: status, approvedDate: r[10], remarks: r[11]
      };
    }).filter(d => d.status !== 'Deleted').reverse();

    const avgDays = approved > 0 ? Math.round(totalExtDays / approved) : 0;
    return { stats: { pending, approved, rejected, avgDays: avgDays + " Days" }, records };
  } catch(e) { return { error: "Failed to load Deferrals: " + e.toString() }; }
}

function getStudentDeferrals(userToken, requestedName, requestedEmail) {
  const auth = verifyAccess(userToken, ['Student', 'Adviser', 'Adviser/IV', 'Admin']);
  if (!auth.authorized) return { error: auth.message };

  const studentName  = (auth.user.role === 'Student') ? auth.user.name  : requestedName;
  const studentEmail = (auth.user.role === 'Student') ? auth.user.email : requestedEmail;

  try {
    // FIX: replaced 3-line open pattern with getDeferralSheet()
    const sheet = getDeferralSheet();
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getDisplayValues();
    return data.slice(1).filter(r => {
      const matchName  = String(r[2]).toLowerCase() === String(studentName).toLowerCase();
      const matchEmail = studentEmail && String(r[12]).toLowerCase() === String(studentEmail).toLowerCase();
      return (matchName || matchEmail) && (r[9] !== 'Deleted');
    }).map(r => ({
      dateReq: r[1], unit: r[3], adviser: r[4], origDeadline: r[5],
      extDays: r[6], reason: r[7], status: r[9] || 'Pending', remarks: r[11]
    })).reverse();
  } catch(e) { return []; }
}

function createStudentDeferral(token, payload) {
  const auth = verifyAccess(token, ['Student', 'Admin']);
  if (!auth.authorized) return { error: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    // FIX: replaced 3-line open pattern with getDeferralSheet()
    const sheet = getDeferralSheet();

    let fileUrl = "";
    if (payload.fileData && payload.fileName) {
      const folderId = getConfig('DeferralDocsID');
      if (folderId) {
        const folder = DriveApp.getFolderById(folderId);
        const file   = folder.createFile(Utilities.newBlob(
          Utilities.base64Decode(payload.fileData), payload.mimeType,
          payload.defId + "_" + payload.fileName
        ));
        fileUrl = file.getUrl();
      }
    }

    sheet.appendRow([
      payload.defId, payload.dateReq, payload.name, payload.unit, payload.adviser,
      payload.origDeadline, payload.extDays, payload.reason, fileUrl,
      "Pending", "", "", payload.email || "", payload.akId || ""
    ]);

    return { success: true, message: "Deferral request submitted successfully!" };
  } catch(e) { return { success: false, message: "Error: " + e.toString() }; }
  finally { lock.releaseLock(); }
}
