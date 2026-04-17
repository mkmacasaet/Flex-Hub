// ===================================================================
// FILE: 04_AccessKit.gs
// PURPOSE: Access Kit Generator — assigns students to units,
//          routes data to AY shards, and syncs the Performance Monitor.
// ===================================================================

function getAccessKitFormOptions() {
  try {
    const masterId = getConfig('MasterStudentListId');
    const akId     = getConfig('AccessKitGeneratorID');
    const confSS   = SpreadsheetApp.getActiveSpreadsheet();

    let students = [], terms = [], specs = [], units = [], advisers = [],
        detailsMatrix = [], academicYears = [], unitStatuses = [];

    try {
      const aySheet = confSS.getSheetByName('_AcademicYears');
      if (aySheet && aySheet.getLastRow() > 1) {
        academicYears = aySheet.getRange(2, 1, aySheet.getLastRow() - 1, 1)
          .getValues().map(r => String(r[0])).filter(String);
      }
    } catch(e) {}

    if (masterId) {
      try {
        const mSheet  = SpreadsheetApp.openById(masterId).getSheets()[0];
        if (mSheet.getLastRow() > 1) {
          const rawData = mSheet.getDataRange().getDisplayValues();
          const headers = rawData[0].map(h => String(h).toLowerCase().trim());
          const col     = getMasterlistColMap(headers);

          students = rawData.slice(1).map(r => {
            const mi = r[col.middleInitial] ? r[col.middleInitial] + '.' : '';
            return {
              id: r[col.id] || '',
              firstName: r[col.firstName] || '',
              middleInitial: r[col.middleInitial] || '',
              lastName: r[col.lastName] || '',
              fullName: `${r[col.lastName] || ''}, ${r[col.firstName] || ''} ${mi}`.trim(),
              email: r[col.sgenEmail] || '',
              pathway: r[col.programme] || '',
              status: r[col.status] || ''
            };
          }).filter(s => s.fullName && s.fullName.length > 2 &&
            (s.status === 'Active' || s.status === 'Spillover'));

          students.sort((a, b) => a.fullName.localeCompare(b.fullName));
        }
      } catch(e) { Logger.log("Error fetching students for Access Kit: " + e.toString()); }
    }

    try {
      const ddSheet = confSS.getSheetByName('_Dropdowns');
      if (ddSheet && ddSheet.getLastRow() > 1) {
        const ddData  = ddSheet.getDataRange().getDisplayValues();
        const headers = ddData[0].map(h => String(h).trim().toLowerCase());
        const termCol = headers.indexOf('term');
        if (termCol > -1) terms = [...new Set(ddData.slice(1).map(r => r[termCol]).filter(String))];
        unitStatuses = [...new Set(ddData.slice(1).map(r => r[3]).filter(String))];
      }
    } catch(e) {}

    if (akId) {
      try {
        const detSheet = SpreadsheetApp.openById(akId).getSheetByName("_Details");
        if (detSheet && detSheet.getLastRow() > 1) {
          const detData = detSheet.getRange(2, 1, detSheet.getLastRow() - 1,
            Math.max(detSheet.getLastColumn(), 15)).getValues();
          units    = [...new Set(detData.map(r => String(r[1] || '').trim()).filter(String))];
          specs    = [...new Set(detData.map(r => String(r[2] || '').trim()).filter(String))];
          advisers = [...new Set(detData.map(r => String(r[10] || '').trim()).filter(String))];
          detailsMatrix = detData.map(r => ({
            akCode: String(r[0]||'').trim(), unit: String(r[1]||'').trim(), spec: String(r[2]||'').trim(),
            level: String(r[3]||'').trim(), credit: String(r[4]||'').trim(), moodleName: r[5]||'',
            moodleLink: r[6]||'', key: r[7]||'', space: r[8]||'', folder: r[9]||'',
            adviser: String(r[10]||'').trim(), advEmail: r[11]||'', ivName: r[12]||'',
            ivEmail: r[13]||'', remarks: r[14]||''
          }));
        }
      } catch(e) {}
    }

    return { students, terms, specs, units, advisers, detailsMatrix, academicYears, unitStatuses };
  } catch(e) { return null; }
}

function getAccessKitList(token, targetAy) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  try {
    const aySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_AcademicYears');
    const ayData  = aySheet.getDataRange().getValues();
    let dbIdsToSweep = [];

    for (let i = 1; i < ayData.length; i++) {
      const yearName = String(ayData[i][0]).trim();
      const akDbId   = String(ayData[i][5]).trim();
      if (akDbId && targetAy === yearName) {
        dbIdsToSweep.push({ ayName: yearName, id: akDbId });
        break;
      }
    }

    if (dbIdsToSweep.length === 0) return { error: "No databases found. Check Column F of _AcademicYears." };

    let allRecords = [];
    dbIdsToSweep.forEach(db => {
      try {
        const shardSS    = SpreadsheetApp.openById(db.id);
        const shardSheet = shardSS.getSheets()[0];
        const data       = shardSheet.getDataRange().getDisplayValues();
        for (let r = 1; r < data.length; r++) {
          allRecords.push({
            rowIndex: r + 1, ayName: db.ayName, akId: data[r][0],
            studentName: data[r][1], unit: data[r][8], adviser: data[r][10],
            start: data[r][12], draft: data[r][13], final: data[r][14],
            statusAk: data[r][29] || "", rowData: data[r]
          });
        }
      } catch (err) { Logger.log("Failed to sweep database ID: " + db.id); }
    });

    return allRecords.reverse();
  } catch (e) { return { error: "System Error: " + e.toString() }; }
}

function getAyInfrastructure(ayName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_AcademicYears');
  if (!sheet) return null;
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(ayName).trim()) {
      return { pmSheetId: String(data[i][1]).trim(), akDbId: String(data[i][5]).trim() };
    }
  }
  return null;
}

function saveAccessKit(token, form) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);

    const targetAy = form.targetAy;
    if (!targetAy) return { success: false, message: "Target Academic Year is required." };

    const infra = getAyInfrastructure(targetAy);
    if (!infra || !infra.akDbId)   return { success: false, message: "Access Kit Database ID not found for " + targetAy + "." };
    if (!infra.pmSheetId)          return { success: false, message: "Performance Monitor Sheet ID not found for " + targetAy + "." };

    const ss        = SpreadsheetApp.openById(infra.akDbId);
    const sheet     = ss.getSheets()[0];
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
    const editorNote = "Updated by: " + (form.editedBy || "Unknown") + " at " + timestamp;

    let finalAKData = [];

    if (form.rowIndex && form.rowIndex !== "") {
      const rowIdx   = parseInt(form.rowIndex);
      const rowRange = sheet.getRange(rowIdx, 1, 1, 28);
      finalAKData    = rowRange.getValues()[0];
      for (const key in form.updates) {
        const colIndex = parseInt(key.split('-')[1]);
        if (!isNaN(colIndex) && colIndex < 28) finalAKData[colIndex] = form.updates[key];
      }
      finalAKData[27] = editorNote;
      rowRange.setValues([finalAKData]);
    } else {
      for (let i = 0; i < 28; i++) { finalAKData.push(form.data[i] || ""); }
      finalAKData[27] = editorNote;
      sheet.appendRow(finalAKData);
    }

    let unitLevel = "", unitCredit = "";
    let detSheet = ss.getSheetByName("_Details");
    if (!detSheet) {
      const oldAkId = getConfig('AccessKitGeneratorID');
      if (oldAkId) detSheet = SpreadsheetApp.openById(oldAkId).getSheetByName("_Details");
    }
    if (detSheet) {
      const detData   = detSheet.getDataRange().getValues();
      const targetUnit = String(finalAKData[8]).trim();
      for (let d = 1; d < detData.length; d++) {
        if (String(detData[d][1]).trim() === targetUnit) {
          unitLevel = detData[d][3]; unitCredit = detData[d][4]; break;
        }
      }
    }

    const akId        = finalAKData[0];
    const studentName = finalAKData[1];
    const pmSheet     = SpreadsheetApp.openById(infra.pmSheetId).getSheets()[0];
    const pmData      = pmSheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < pmData.length; i++) { if (String(pmData[i][16]) === String(akId)) { foundRow = i + 1; break; } }

    const pmRow = [
      form.studentId || "PENDING", studentName, finalAKData[8], unitLevel, finalAKData[6], unitCredit,
      "", "", "", "", "", "", "", false, "", "", akId, "Pending", "", "Pending", "",
      form.editedBy || "Unknown Admin", timestamp, finalAKData[2], "No"
    ];

    if (foundRow > -1) pmSheet.getRange(foundRow, 1, 1, pmRow.length).setValues([pmRow]);
    else pmSheet.appendRow(pmRow);

    updateSurgicalIndex(form.studentId, targetAy, infra.pmSheetId);
    return { success: true, message: "Access Kit saved to Shard, and Grades Synced!" };
  } catch(e) {
    return { success: false, message: "Error: " + e.toString() };
  } finally { lock.releaseLock(); }
}

function deleteAccessKitRow(token, rowIndex, targetAy) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (!targetAy) return { success: false, message: "Target Academic Year is required for deletion." };

    const infra = getAyInfrastructure(targetAy);
    if (!infra || !infra.akDbId) return { success: false, message: "Access Kit Database ID not found." };

    const sheet   = SpreadsheetApp.openById(infra.akDbId).getSheets()[0];
    const rowData = sheet.getRange(rowIndex, 1, 1, 28).getValues()[0];
    const akId    = rowData[0];
    sheet.deleteRow(rowIndex);

    let targetStudentId = null, remainingUnits = 0;
    if (infra.pmSheetId && akId) {
      const pmSheet = SpreadsheetApp.openById(infra.pmSheetId).getSheets()[0];
      const pmData  = pmSheet.getDataRange().getValues();
      for (let i = pmData.length - 1; i >= 1; i--) {
        if (String(pmData[i][16]).trim() === String(akId).trim()) {
          targetStudentId = String(pmData[i][0]).trim();
          pmSheet.getRange(i + 1, 25).setValue("Yes");
          break;
        }
      }
      if (targetStudentId && targetStudentId !== "PENDING") {
        for (let i = 1; i < pmData.length; i++) {
          const rStuId  = String(pmData[i][0]).trim();
          const rAkId   = String(pmData[i][16]).trim();
          const rHidden = String(pmData[i][24]).trim();
          if (rStuId === targetStudentId && rAkId !== String(akId).trim() && rHidden !== "Yes") remainingUnits++;
        }
      }
    }

    if (targetStudentId && remainingUnits === 0) {
      const indexSheet = getStudentIndexSheet();
      const indexData  = indexSheet.getDataRange().getValues();
      for (let i = indexData.length - 1; i > 0; i--) {
        if (String(indexData[i][0]).trim() === targetStudentId &&
            String(indexData[i][1]).trim() === String(targetAy).trim()) {
          indexSheet.deleteRow(i + 1); break;
        }
      }
    }

    return { success: true, message: "Access Kit deleted. Database optimized." };
  } catch (e) {
    return { success: false, message: "Error deleting: " + e.toString() };
  } finally { lock.releaseLock(); }
}
