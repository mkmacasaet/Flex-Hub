// ===================================================================
// FILE: 12_Risk.gs
// PURPOSE: Early Warning System & Risk Management.
//          Contains the main report engine, all fetch* helpers,
//          and the Adviser wrapper.
// FIXES APPLIED:
//   - getDeferralSheet() used instead of inline Deferral sheet open
// ===================================================================

function getRiskAssessmentReport(token) {
  const auth = verifyAccess(token, ['Admin', 'Adviser', 'Adviser/IV']);
  if (!auth.authorized) return { success: false, error: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const { excludedIds, flaggedCases } = fetchRiskMetadata();
    const targetStudents = fetchActiveStudentsForScan(excludedIds);
    const meetingSet     = fetchMeetingLogsSet();
    const akDateMap      = fetchAccessKitDates();
    const academicYears  = getAcademicYears(token);
    const akInfoMap      = fetchAccessKitInfoMap();
    const userEmailMap   = fetchUserEmailMap();

    let potentialRisks  = [];
    let studentRiskMap  = {};
    let studentAdviserMap = {};

    targetStudents.forEach(s => {
      studentRiskMap[s.id] = { name: s.name, factors: new Set(), details: [], enrollmentStatus: s.status };
    });

    academicYears.forEach(ay => {
      if (!ay.id) return;
      try {
        const pmSheet = SpreadsheetApp.openById(ay.id).getSheets()[0];
        const pmData  = pmSheet.getDataRange().getDisplayValues();

        for (let i = 1; i < pmData.length; i++) {
          const sId = String(pmData[i][0]).trim();
          if (!studentRiskMap[sId]) continue;

          const student       = studentRiskMap[sId];
          const unit          = pmData[i][2];
          const akId          = String(pmData[i][16]).trim();
          const assessorGrade = pmData[i][10].toUpperCase();
          const finalGrade    = pmData[i][12].toUpperCase();
          const akInfo        = akDateMap[akId] || {};
          const startDate     = akInfo.startDate    ? new Date(akInfo.startDate)    : null;
          const finalDeadline = akInfo.finalDeadline ? new Date(akInfo.finalDeadline) : null;

          if (akInfo.adviserName) {
            const rawName      = String(akInfo.adviserName).trim();
            const advEmail     = userEmailMap[rawName.toLowerCase()];
            const contactString = advEmail ? `${advEmail} | ${rawName}` : rawName;
            if (!studentAdviserMap[sId]) studentAdviserMap[sId] = new Set();
            studentAdviserMap[sId].add(contactString);
          }

          if (finalGrade === 'REFERRED' || assessorGrade === 'REFERRED') {
            student.factors.add("Academic");
            student.details.push(`<b>${unit}</b>: Grade is Referred`);
          }
          if (!finalGrade && !assessorGrade && finalDeadline && finalDeadline < today) {
            student.factors.add("Academic");
            student.details.push(`<b>${unit}</b>: Overdue (No grade submitted)`);
          }
          if (startDate && !isNaN(startDate.getTime()) && !finalGrade) {
            const daysSinceStart = Math.floor((today - startDate) / (1000 * 60 * 60 * 24));
            const hasMeeting     = meetingSet.has(`${student.name.toLowerCase()}|${unit.toLowerCase()}`);
            if (daysSinceStart > 21 && !hasMeeting) {
              student.factors.add("Engagement");
              student.details.push(`<b>${unit}</b>: No meeting logged in ${daysSinceStart} days since Access Kit sent`);
            }
          }
        }
      } catch (e) { Logger.log("Error scanning AY: " + ay.id); }
    });

    // FIX: getDeferralSheet() instead of inline open
    const defSheet = getDeferralSheet();
    const defData  = defSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < defData.length; i++) {
      const sName    = defData[i][2].toLowerCase();
      const status   = defData[i][9];
      const deadline = new Date(defData[i][10]);
      if (status === 'Approved' && !isNaN(deadline.getTime()) && deadline < today) {
        const studentId = Object.keys(studentRiskMap).find(id => studentRiskMap[id].name.toLowerCase() === sName);
        if (studentId) {
          studentRiskMap[studentId].factors.add("Pacing");
          studentRiskMap[studentId].details.push(`<b>${defData[i][3]}</b>: Deferral deadline passed`);
        }
      }
    }

    const loaId    = getConfig('WithdrawalLOAID');
    const loaSheet = SpreadsheetApp.openById(loaId).getSheets()[0];
    const loaData  = loaSheet.getDataRange().getDisplayValues();
    const loaHeaders = loaData[0].map(h => String(h).toLowerCase().trim());
    const loaMap   = getLoaColMap(loaHeaders);

    for (let i = 1; i < loaData.length; i++) {
      const sId        = String(loaData[i][loaMap.studentId]).trim();
      const sType      = String(loaData[i][loaMap.separationType]).trim();
      const returnDate = new Date(loaData[i][loaMap.returnDate]);
      if (!isNaN(returnDate.getTime()) && returnDate < today) {
        if (sType !== "LOA Completed" && studentRiskMap[sId]) {
          studentRiskMap[sId].factors.add("LOA");
          studentRiskMap[sId].details.push(`LOA Return Date passed (${loaData[i][loaMap.returnDate]}) but status is still <b>${sType}</b>`);
        }
      }
    }

    for (const sId in studentRiskMap) {
      const s = studentRiskMap[sId];
      if (s.factors.size > 0) {
        let assignedEmails = [];
        if (s.factors.size === 1 && s.factors.has("LOA")) {
          assignedEmails = ["Flex Admin"];
        } else {
          assignedEmails = studentAdviserMap[sId] ? Array.from(studentAdviserMap[sId]) : ["Flex Admin"];
        }
        potentialRisks.push({
          id: sId, name: s.name, factors: Array.from(s.factors),
          details: s.details.join('<br>'), adviserEmails: assignedEmails.join(', ')
        });
      }
    }

    return { success: true, warnings: potentialRisks, flaggedCases: flaggedCases.reverse() };
  } catch(e) {
    return { success: false, error: e.toString() };
  } finally { lock.releaseLock(); }
}

// ─── HELPER FUNCTIONS ───────────────────────────────────────────────

function fetchAccessKitDates() {
  const akId  = getConfig('AccessKitGeneratorID');
  const sheet = SpreadsheetApp.openById(akId).getSheets()[0];
  const data  = sheet.getDataRange().getDisplayValues();
  let map = {};
  for (let i = 1; i < data.length; i++) {
    map[String(data[i][0]).trim()] = { startDate: data[i][12], finalDeadline: data[i][14] };
  }
  return map;
}

function fetchMeetingLogsSet() {
  const mtgId = getConfig('MeetingLogsID');
  const data  = SpreadsheetApp.openById(mtgId).getSheets()[0].getDataRange().getDisplayValues();
  let set = new Set();
  for (let i = 1; i < data.length; i++) {
    set.add(`${data[i][2].trim().toLowerCase()}|${data[i][3].trim().toLowerCase()}`);
  }
  return set;
}

function fetchRiskMetadata() {
  const riskId = getConfig('RiskMonitorID');
  const sheet  = SpreadsheetApp.openById(riskId).getSheets()[0];
  const data   = sheet.getDataRange().getDisplayValues();
  let flaggedCases = [], excludedIds = [];
  for (let i = 1; i < data.length; i++) {
    const status = data[i][4];
    const sId    = String(data[i][1]).trim();
    const savedEmails  = data[i][9]  || '';
    const savedDetails = data[i][10] || '';
    if (status === 'Flagged') {
      excludedIds.push(sId);
      flaggedCases.push({
        recordId: data[i][0], studentId: sId, studentName: data[i][2],
        factors: data[i][3], status: status, logs: data[i][5],
        dateFlagged: data[i][6] ? new Date(data[i][6]).toLocaleDateString() : '',
        adviserEmails: savedEmails, details: savedDetails
      });
    } else if (status === 'Dismissed') {
      excludedIds.push(sId);
    } else if (status === 'Resolved') {
      flaggedCases.push({
        recordId: data[i][0], studentId: sId, studentName: data[i][2],
        factors: data[i][3], status: status, logs: data[i][5],
        dateFlagged: data[i][6] ? new Date(data[i][6]).toLocaleDateString() : '',
        adviserEmails: savedEmails, details: savedDetails
      });
    }
  }
  return { excludedIds, flaggedCases };
}

function fetchActiveStudentsForScan(excludedIds) {
  const masterId = getConfig('MasterStudentListId');
  const sheet    = SpreadsheetApp.openById(masterId).getSheets()[0];
  const data     = sheet.getDataRange().getDisplayValues();
  const col      = getMasterlistColMap(data[0].map(h => String(h).toLowerCase().trim()));
  let list = [];
  for (let i = 1; i < data.length; i++) {
    const sId = String(data[i][col.id]).trim();
    const stat = String(data[i][col.status]).trim();
    if (!excludedIds.includes(sId)) {
      list.push({
        id: sId,
        name: formatUniversalName(data[i][col.firstName], data[i][col.middleInitial], data[i][col.lastName]),
        status: stat
      });
    }
  }
  return list;
}

function fetchAcademicReferredMap(token) {
  let referredMap = {};
  const ays = getAcademicYears(token);
  for (const ay of ays) {
    if (ay.id) {
      try {
        const pmData = SpreadsheetApp.openById(ay.id).getSheets()[0].getDataRange().getDisplayValues();
        for (let i = 1; i < pmData.length; i++) {
          const sId          = String(pmData[i][0]).trim().toLowerCase();
          const assessorGrade = String(pmData[i][10]).trim().toUpperCase();
          const finalGrade    = String(pmData[i][12]).trim().toUpperCase();
          if (assessorGrade === 'REFERRED' || finalGrade === 'REFERRED') {
            referredMap[sId] = (referredMap[sId] || 0) + 1;
          }
        }
      } catch(e) { Logger.log("Skipping AY sheet: " + ay.id + " — " + e.toString()); }
    }
  }
  return referredMap;
}

function fetchDeferralCountMap() {
  // FIX: getDeferralSheet() instead of inline open
  const sheet = getDeferralSheet();
  const data  = sheet.getDataRange().getDisplayValues();
  let map = {};
  for (let i = 1; i < data.length; i++) {
    const sName = String(data[i][2]).trim().toLowerCase();
    map[sName] = (map[sName] || 0) + 1;
  }
  return map;
}

function fetchLastMeetingMap() {
  const mtgId = getConfig('MeetingLogsID');
  const sheet = SpreadsheetApp.openById(mtgId).getSheets()[0];
  const data  = sheet.getDataRange().getDisplayValues();
  let map = {};
  for (let i = 1; i < data.length; i++) {
    const sName = String(data[i][2]).trim().toLowerCase();
    const mDate = new Date(data[i][4]);
    if (!isNaN(mDate.getTime())) {
      if (!map[sName] || mDate > map[sName]) map[sName] = mDate;
    }
  }
  return map;
}

function fetchAccessKitInfoMap() {
  const akId  = getConfig('AccessKitGeneratorID');
  const sheet = SpreadsheetApp.openById(akId).getSheets()[0];
  const data  = sheet.getDataRange().getDisplayValues();
  let map = {};
  for (let i = 1; i < data.length; i++) {
    map[String(data[i][0]).trim()] = {
      adviserName:   String(data[i][10]).trim(),
      startDate:     data[i][12],
      finalDeadline: data[i][14]
    };
  }
  return map;
}

function fetchUserEmailMap() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_Users');
    const data  = sheet.getDataRange().getDisplayValues();
    let map = {};
    for (let i = 1; i < data.length; i++) {
      map[String(data[i][1]).trim().toLowerCase()] = String(data[i][2]).trim().toLowerCase();
    }
    return map;
  } catch(e) { return {}; }
}

function handleRiskWarning(token, payload) {
  const auth = verifyAccess(token, ['Admin', 'Adviser', 'Adviser/IV']);
  if (!auth.authorized) return { success: false, message: auth.message };

  try {
    const riskId    = getConfig('RiskMonitorID');
    const sheet     = SpreadsheetApp.openById(riskId).getSheets()[0];
    const recordId  = "RSK-" + Math.floor(Math.random() * 900000 + 100000);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
    const logEntry  = `[${timestamp}] System Warning processed as: ${payload.action} by ${auth.user.name}`;

    sheet.appendRow([
      recordId, payload.studentId, payload.name, payload.factors,
      payload.action, logEntry, timestamp, timestamp, auth.user.name
    ]);

    return { success: true, message: `Student successfully ${payload.action}!` };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function updateRiskCase(token, payload) {
  const auth = verifyAccess(token, ['Admin', 'Adviser', 'Adviser/IV']);
  if (!auth.authorized) return { success: false, message: auth.message };

  try {
    const riskId = getConfig('RiskMonitorID');
    const sheet  = SpreadsheetApp.openById(riskId).getSheets()[0];
    const data   = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.recordId)) { rowIndex = i + 1; break; }
    }
    if (rowIndex === -1) return { success: false, message: "Record not found." };

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
    let currentLogs = data[rowIndex - 1][5];

    if (payload.newLog) {
      currentLogs = `[${timestamp} - ${auth.user.name}]\n${payload.newLog}\n\n` + currentLogs;
    }
    if (payload.status === 'Resolved') {
      currentLogs = `[${timestamp} - SYSTEM]\nCase officially marked as Resolved by ${auth.user.name}.\n\n` + currentLogs;
    }

    sheet.getRange(rowIndex, 5, 1, 5).setValues([[
      payload.status, currentLogs, data[rowIndex - 1][6], timestamp, auth.user.name
    ]]);

    return { success: true, message: "Intervention Log updated!" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function handleBatchRiskAction(token, payload) {
  const auth = verifyAccess(token, ['Admin', 'Adviser', 'Adviser/IV']);
  if (!auth.authorized) return { success: false, message: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const riskId    = getConfig('RiskMonitorID');
    const sheet     = SpreadsheetApp.openById(riskId).getSheets()[0];
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
    let rowsToAppend = [];

    payload.items.forEach(item => {
      const recordId     = "RSK-" + Math.floor(Math.random() * 900000 + 100000);
      const logEntry     = `[${timestamp}] System Warning processed as: ${payload.action} by ${auth.user.name}`;
      const emailsToSave = item.adviserEmails || 'Flex Admin';
      const detailsToSave = item.details || '';
      rowsToAppend.push([
        recordId, item.id, item.name, item.factors.join(', '),
        payload.action, logEntry, timestamp, timestamp, auth.user.name,
        emailsToSave, detailsToSave
      ]);
    });

    if (rowsToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }

    return { success: true, message: `Successfully processed ${payload.items.length} records.` };
  } catch (e) {
    return { success: false, message: "System Error: " + e.toString() };
  } finally { lock.releaseLock(); }
}

function getAdviserRiskAssessmentReport(token) {
  const auth = verifyAccess(token, ['Adviser', 'Adviser/IV', 'Admin']);
  if (!auth.authorized) return { success: false, error: auth.message };

  try {
    const myEmail      = auth.user.email.toLowerCase();
    const emailMap     = fetchUserEmailMap();
    let mySystemNames  = [auth.user.name.toLowerCase()];
    for (const [name, email] of Object.entries(emailMap)) {
      if (email === myEmail) mySystemNames.push(name);
    }

    const akId    = getConfig('AccessKitGeneratorID');
    const akData  = SpreadsheetApp.openById(akId).getSheets()[0].getDataRange().getDisplayValues();
    const myAkCodes = new Set();
    for (let i = 1; i < akData.length; i++) {
      const code    = String(akData[i][0]).trim().toUpperCase();
      const advName = String(akData[i][10]).trim().toLowerCase();
      if (mySystemNames.includes(advName)) myAkCodes.add(code);
    }

    const ays          = getAcademicYears(token);
    const myStudentIds = new Set();
    ays.forEach(ay => {
      if (!ay.id) return;
      try {
        const pmData = SpreadsheetApp.openById(ay.id).getSheets()[0].getDataRange().getDisplayValues();
        for (let i = 1; i < pmData.length; i++) {
          const sId    = String(pmData[i][0]).trim();
          const akCode = String(pmData[i][16]).trim().toUpperCase();
          if (myAkCodes.has(akCode)) myStudentIds.add(sId);
        }
      } catch(e) {}
    });

    const report = getRiskAssessmentReport(token);
    if (!report.success) return report;

    const myIdsArray = Array.from(myStudentIds);
    report.warnings = report.warnings.filter(w => myIdsArray.includes(w.id)).map(w => {
      w.adviserEmails = `${myEmail}`;
      return w;
    });
    report.flaggedCases = report.flaggedCases.filter(c => myIdsArray.includes(c.studentId));

    return report;
  } catch(e) { return { success: false, error: e.toString() }; }
}
