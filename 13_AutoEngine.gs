// ===================================================================
// FILE: 13_AutoEngine.gs
// PURPOSE: Automatic Student Status Engine.
//          Run via a time-based trigger (e.g., daily at midnight).
//          NOT called by the frontend.
// FIXES APPLIED:
//   - getDeferralSheet() used instead of inline Deferral open
//   - getAYSheet() used instead of inline AY open
// ===================================================================

function runAutoStatusEngine() {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    // PHASE 1: GATHER SUPPORTING DATA
    // FIX: getAYSheet() instead of inline open
    const aySheet = getAYSheet();
    const ayData  = aySheet.getDataRange().getValues();
    let academicYears = [];
    for (let i = 1; i < ayData.length; i++) {
      if (ayData[i][0]) {
        academicYears.push({
          name:  String(ayData[i][0]).trim(),
          start: new Date(ayData[i][2]).getTime(),
          end:   new Date(ayData[i][3]).getTime()
        });
      }
    }

    // FIX: getDeferralSheet() instead of inline open
    let deferralMap = {};
    try {
      const defSheet = getDeferralSheet();
      const defData  = defSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < defData.length; i++) {
        const status      = defData[i][9];
        const newDeadline = new Date(defData[i][10]).getTime();
        const akId        = String(defData[i][13]).trim();
        if (status === 'Approved' && akId && !isNaN(newDeadline)) {
          deferralMap[akId] = newDeadline;
        }
      }
    } catch(e) { Logger.log("Error reading deferrals in auto-status"); }

    const loaId = getConfig('WithdrawalLOAID');
    let activeLoaMap = {};
    if (loaId) {
      try {
        const loaSheet  = SpreadsheetApp.openById(loaId).getSheets()[0];
        const loaData   = loaSheet.getDataRange().getDisplayValues();
        const loaHeaders = loaData[0].map(h => String(h).toLowerCase().trim());
        const loaMap    = getLoaColMap(loaHeaders);

        for (let i = 1; i < loaData.length; i++) {
          const sId   = String(loaData[i][loaMap.studentId]).trim();
          const sType = String(loaData[i][loaMap.separationType]).trim();
          if (sType && sType !== "LOA Completed" && sType !== "Returned") {
            activeLoaMap[sId] = sType;
          }
        }
      } catch(e) { Logger.log("Error reading LOA data"); }
    }

    // PHASE 2: SCAN ACCESS KITS
    const akId    = getConfig('AccessKitGeneratorID');
    const akSheet = SpreadsheetApp.openById(akId).getSheets()[0];
    const akData  = akSheet.getDataRange().getValues();

    let studentKits = {};
    for (let i = 1; i < akData.length; i++) {
      const rowAkId = String(akData[i][0]).trim();
      const sName   = String(akData[i][1]).trim();
      if (!sName) continue;

      if (!studentKits[sName]) studentKits[sName] = { kits: [], hasSpillover: false };

      const assignedAyName = String(akData[i][7]).trim();
      const status         = String(akData[i][29] || '').trim().toLowerCase();
      const isComplete     = status.includes('complete') || status.includes('pass') ||
                             status.includes('merit') || status.includes('distinction');

      let effectiveDeadline = deferralMap[rowAkId];
      if (!effectiveDeadline) {
        const finalDate = new Date(akData[i][14]);
        effectiveDeadline = finalDate.getTime();
      }

      studentKits[sName].kits.push({ isComplete });

      if (!isComplete && effectiveDeadline && !isNaN(effectiveDeadline)) {
        let fallingAyName = null;
        for (const ay of academicYears) {
          if (effectiveDeadline >= ay.start && effectiveDeadline <= ay.end) {
            fallingAyName = ay.name; break;
          }
        }
        if (fallingAyName && fallingAyName !== assignedAyName) {
          studentKits[sName].hasSpillover = true;
        }
      }
    }

    // PHASE 3: APPLY PRIORITY HIERARCHY TO MASTER LIST
    const mSheet = getMasterSheet();
    const mData  = mSheet.getDataRange().getValues();
    let updatesCount = 0;

    for (let i = 1; i < mData.length; i++) {
      const studentId    = String(mData[i][0]).trim();
      const lName        = String(mData[i][1] || '').trim();
      const fName        = String(mData[i][2] || '').trim();
      const mi           = mData[i][4] ? String(mData[i][4]).trim() + "." : "";
      const universalName = `${lName}, ${fName} ${mi}`.trim();
      const currentStatus = String(mData[i][23]).trim();

      let calculatedStatus = "Active";
      const studentData    = studentKits[universalName];

      // PRIORITY HIERARCHY
      if (activeLoaMap[studentId]) {
        calculatedStatus = activeLoaMap[studentId];
      } else if (studentData && studentData.kits.length >= 15 && studentData.kits.every(k => k.isComplete)) {
        calculatedStatus = "Completed";
      } else if (studentData && studentData.hasSpillover) {
        calculatedStatus = "Spillover";
      }

      if (calculatedStatus !== currentStatus) {
        mSheet.getRange(i + 1, 24).setValue(calculatedStatus);
        updatesCount++;
      }
    }

    return `SUCCESS: Auto-Status Engine updated ${updatesCount} students.`;

  } catch(e) {
    Logger.log("Engine Error: " + e.message);
    return `ERROR: ${e.message}`;
  } finally {
    lock.releaseLock();
  }
}
