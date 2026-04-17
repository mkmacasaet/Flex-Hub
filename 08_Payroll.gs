// ===================================================================
// FILE: 08_Payroll.gs
// PURPOSE: Adviser payroll and synchronous meeting logs.
// ===================================================================

function getAllMeetings(userToken) {
  const auth = verifyAccess(userToken, ['Adviser', 'Adviser/IV', 'Admin']);
  if (!auth.authorized) return { error: auth.message };

  try {
    const id = getConfig('MeetingLogsID');
    if (!id) return { error: "MeetingLogsID missing in _Config sheet." };
    const sheet = SpreadsheetApp.openById(id).getSheets()[0];
    const data  = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    return data.slice(1).map((r, i) => ({
      rowIndex: i + 2, mtgId: r[0], adviser: r[1], student: r[2], unit: r[3], date: r[4],
      timeStart: r[5], timeEnd: r[6], duration: r[7], topic: r[8], proof: r[9],
      payrollStatus: r[10], payrollRemarks: r[11], timestamp: r[12]
    })).reverse();
  } catch(e) { return { error: "Failed to load Meeting Logs: " + e.toString() }; }
}

function saveMeetingLog(token, form) {
  const auth = verifyAccess(token, ['Adviser', 'Adviser/IV', 'Admin']);
  if (!auth.authorized) return { error: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const id = getConfig('MeetingLogsID');
    if (!id) return { success: false, message: "MeetingLogsID missing in _Config sheet." };
    const sheet     = SpreadsheetApp.openById(id).getSheets()[0];
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
    const mtgId     = `MTG-${new Date().getFullYear()}-${Math.floor(Math.random() * 9000 + 1000)}`;

    sheet.appendRow([mtgId, form.adviser, form.student, form.unit, form.date,
      form.timeStart, form.timeEnd, form.duration, form.topic, form.proof, "Pending", "", timestamp]);

    return { success: true, message: "Meeting log added successfully!" };
  } catch(e) {
    return { success: false, message: "Error saving meeting: " + e.toString() };
  } finally { lock.releaseLock(); }
}

function updateMeetingPayroll(token, form) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const id    = getConfig('MeetingLogsID');
    const sheet = SpreadsheetApp.openById(id).getSheets()[0];
    sheet.getRange(parseInt(form.rowIndex), 11, 1, 2).setValues([[form.payrollStatus, form.payrollRemarks]]);
    return { success: true, message: "Payroll status updated successfully!" };
  } catch(e) {
    return { success: false, message: "Error updating: " + e.toString() };
  } finally { lock.releaseLock(); }
}

function getUnitPayrollForAY(token, sheetId) {
  const auth = verifyAccess(token, ['Admin', 'Adviser', 'Adviser/IV']);
  if (!auth.authorized) return { error: auth.message };

  try {
    if (!sheetId) return [];

    const sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
    if (sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getDisplayValues();
    const rows = data.slice(1).map((r, i) => ({
      sheetId, rowIndex: i + 2, akId: r[16], studentName: r[1], unit: r[2],
      assessorGrade: r[10], unitStatus: r[23],
      advPayrollStatus: r[17], advPayrollRemarks: r[18],
      assPayrollStatus: r[19], assPayrollRemarks: r[20],
      isHidden: String(r[24] || "").trim()
    })).filter(record => record.isHidden !== 'Yes');

    const akIdConf = getConfig('AccessKitGeneratorID');
    let akMap = {};
    if (akIdConf) {
      const akSS   = SpreadsheetApp.openById(akIdConf);
      const akData = akSS.getSheets()[0].getDataRange().getDisplayValues();
      const detData = akSS.getSheetByName("_Details").getDataRange().getDisplayValues();
      let detMap = {};
      for (let d = 1; d < detData.length; d++) {
        detMap[String(detData[d][1]).trim() + "|" + String(detData[d][10]).trim()] = {
          assessor: String(detData[d][12]).trim(), folder: String(detData[d][9]).trim()
        };
      }
      for (let a = 1; a < akData.length; a++) {
        const adv   = String(akData[a][10]).trim();
        const match = detMap[String(akData[a][8]).trim() + "|" + adv] || {};
        akMap[String(akData[a][0]).trim()] = { adviser: adv, assessor: match.assessor || adv, folder: match.folder || "" };
      }
    }

    return rows.map(g => ({
      ...g,
      adviserName:  akMap[String(g.akId).trim()]?.adviser  || "Unknown",
      assessorName: akMap[String(g.akId).trim()]?.assessor || "Unknown",
      gDriveLink:   akMap[String(g.akId).trim()]?.folder   || ""
    })).reverse();
  } catch(e) { return { error: "Failed to load Unit Payroll." }; }
}

function updateUnitPayroll(token, form) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(form.sheetId).getSheets()[0];
    if (!form.rowIndex) return { success: false, message: "Missing Row Index!" };

    if (form.payType === 'Adviser') {
      sheet.getRange(parseInt(form.rowIndex), 18, 1, 2).setValues([[form.payrollDate, form.payrollRemarks]]);
    } else {
      sheet.getRange(parseInt(form.rowIndex), 20, 1, 2).setValues([[form.payrollDate, form.payrollRemarks]]);
    }

    return { success: true, message: form.payType + " payment saved successfully!" };
  } catch(e) {
    return { success: false, message: "System Error: " + e.toString() };
  } finally { lock.releaseLock(); }
}
