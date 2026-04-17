// ===================================================================
// FILE: 06_Grades.gs
// PURPOSE: Performance Monitor & Grading — academic year management,
//          grade entry, surgical transcript lookup, and PDF generation.
// FIXES APPLIED:
//   - updateSurgicalIndex / getStudentProgressionSurgical:
//       Replaced StudentIndex open pattern with getStudentIndexSheet()
//   - getDeferralSheet() used inside getGradesForAY instead of
//       inline open pattern
//   - batchPublishFinalGrades and batchUpdateRegistrarDates MOVED HERE
//       from the end of old Code.txt (they logically belong in Grades)
// ===================================================================

function getAcademicYears(userToken) {
  const auth = verifyAccess(userToken, ['Adviser', 'Adviser/IV', 'Admin', 'Student']);
  if (!auth.authorized) return { error: auth.message };

  try {
    const sheet = getAYSheet();
    if (!sheet || sheet.getLastRow() < 2) return [];
    const rawData = sheet.getDataRange().getDisplayValues();
    return rawData.slice(1).map(r => ({
      year: r[0], id: r[1], status: r[4], akDbId: r[5] || ""
    })).filter(r => r.year !== "");
  } catch(e) { return { error: "Server error reading Academic Years: " + e.toString() }; }
}

function getGradesForAY(userToken, sheetId) {
  const auth = verifyAccess(userToken, ['Adviser', 'Adviser/IV', 'Admin']);
  if (!auth.authorized) return { error: auth.message };

  try {
    const ss    = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheets()[0];
    if (sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getDisplayValues();
    let akMap  = {};

    try {
      const akIdConfig = getConfig('AccessKitGeneratorID');
      const akSS   = SpreadsheetApp.openById(akIdConfig);
      const akData = akSS.getSheets()[0].getDataRange().getDisplayValues();
      const detData = akSS.getSheetByName("_Details").getDataRange().getDisplayValues();

      let detMap = {};
      for (let d = 1; d < detData.length; d++) {
        const key = String(detData[d][1]).trim() + "|" + String(detData[d][2]).trim() + "|" + String(detData[d][10]).trim();
        detMap[key] = {
          folder: String(detData[d][9] || "").trim(), ivEmail: String(detData[d][13] || "").trim(),
          moodleName: String(detData[d][5] || "").trim(), moodleLink: String(detData[d][6] || "").trim()
        };
      }
      for (let a = 1; a < akData.length; a++) {
        const unit  = String(akData[a][8]).trim();
        const spec  = String(akData[a][9]).trim();
        const adv   = String(akData[a][10]).trim();
        const match = detMap[unit + "|" + spec + "|" + adv] || {};
        akMap[String(akData[a][0]).trim()] = {
          folder: match.folder || "", adviser: adv, ivEmail: match.ivEmail || "",
          moodleName: match.moodleName || "", moodleLink: match.moodleLink || "",
          startDate: String(akData[a][12] || "").trim(),
          draftDate: String(akData[a][13] || "").trim(),
          finalDate: String(akData[a][14] || "").trim()
        };
      }
    } catch(e) { Logger.log("Lookup Error: " + e.toString()); }

    // FIX: getDeferralSheet() instead of inline open
    let defMap = {};
    try {
      const defSheet = getDeferralSheet();
      const defData  = defSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < defData.length; i++) {
        const defAkId = String(defData[i][13] || "").trim();
        if (defAkId) defMap[defAkId] = { status: defData[i][9], approvedDate: defData[i][10] };
      }
    } catch(e) { Logger.log("Deferral Lookup Error: " + e.toString()); }

    let secureData = data.slice(1).map((r, i) => ({
      rowIndex: i + 2, studentId: r[0], studentName: r[1], unit: r[2], level: r[3],
      term: r[4], credits: r[5], lo1: r[6], lo2: r[7], lo3: r[8], lo4: r[9],
      assessor: r[10], ivGrade: r[11], finalGrade: r[12], ivSampling: r[13],
      ivComments: r[14], facultyRemarks: r[15], akId: r[16],
      advPayrollStatus: r[17], advPayrollRemarks: r[18],
      assPayrollStatus: r[19], assPayrollRemarks: r[20],
      updatedBy: r[21], timestamp: r[22], unitStatus: r[23],
      isHidden: String(r[24] || "").trim(),
      adviserName: akMap[String(r[16]).trim()]?.adviser || "Not Assigned",
      gDriveLink:  akMap[String(r[16]).trim()]?.folder || "",
      ivEmail:     akMap[String(r[16]).trim()]?.ivEmail || "",
      moodleName:  akMap[String(r[16]).trim()]?.moodleName || "",
      moodleLink:  akMap[String(r[16]).trim()]?.moodleLink || "",
      startDate:   akMap[String(r[16]).trim()]?.startDate || "",
      draftDate:   akMap[String(r[16]).trim()]?.draftDate || "",
      finalDate:   akMap[String(r[16]).trim()]?.finalDate || "",
      deferralStatus: defMap[String(r[16]).trim()]?.status || "",
      deferralDate:   defMap[String(r[16]).trim()]?.approvedDate || ""
    })).filter(record => record.isHidden !== 'Yes');

    if (auth.user.role === 'Adviser' || auth.user.role === 'Adviser/IV') {
      secureData = secureData.filter(row => {
        const isMyStudent     = String(row.adviserName).trim().toLowerCase() === String(auth.user.name).trim().toLowerCase();
        const isMyIvAssignment = String(row.ivEmail).trim().toLowerCase() === String(auth.user.email).trim().toLowerCase();
        return isMyStudent || isMyIvAssignment;
      });
    }

    return secureData;
  } catch(e) { return { error: "Failed to load grades." }; }
}

function updateGradeForAY(token, form) {
  const auth = verifyAccess(token, ['Adviser', 'Adviser/IV', 'Admin']);
  if (!auth.authorized) return { error: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss        = SpreadsheetApp.openById(form.sheetId);
    const sheet     = ss.getSheets()[0];
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
    const rowIdx    = parseInt(form.rowIndex);
    const rowRange  = sheet.getRange(rowIdx, 1, 1, 23);
    const rowData   = rowRange.getValues()[0];

    const fieldMap = {
      'ge-lo1': 6, 'ge-lo2': 7, 'ge-lo3': 8, 'ge-lo4': 9,
      'ge-assessor': 10, 'ge-ivGrade': 11, 'ge-finalGrade': 12,
      'ge-ivSampling': 13, 'ge-ivComments': 14, 'ge-facultyRemarks': 15
    };

    for (const key in form.updates) {
      if (fieldMap[key] !== undefined) {
        rowData[fieldMap[key]] = (key === 'ge-ivSampling') ? (form.updates[key] === 'true') : form.updates[key];
      }
    }

    rowData[21] = form.editedBy || "Unknown";
    rowData[22] = timestamp;
    rowRange.setValues([rowData]);
    return { success: true, message: "Record surgically updated successfully!" };
  } catch(e) {
    return { success: false, message: e.toString() };
  } finally { lock.releaseLock(); }
}

function getStudentProgressionSurgical(token, requestedId) {
  const auth = verifyAccess(token, ['Admin', 'Adviser', 'Adviser/IV', 'Student']);
  if (!auth.authorized) return { error: auth.message };

  const studentId = (auth.user.role === 'Student') ? auth.user.id : requestedId;

  try {
    // FIX: replaced StudentIndex open pattern with getStudentIndexSheet()
    const indexSheet = getStudentIndexSheet();
    if (!indexSheet || indexSheet.getLastRow() < 2) return { records: [] };

    const indexData  = indexSheet.getDataRange().getDisplayValues();
    let activeYears  = [];

    for (let i = 1; i < indexData.length; i++) {
      if (String(indexData[i][0]).trim() === String(studentId).trim()) {
        activeYears.push(String(indexData[i][1]).trim());
      }
    }

    if (activeYears.length === 0) return { records: [] };

    let records = [];

    activeYears.forEach(ayName => {
      const infra = getAyInfrastructure(ayName);
      if (!infra || !infra.pmSheetId || !infra.akDbId) return;

      try {
        let akMap = {};
        try {
          const akSS    = SpreadsheetApp.openById(infra.akDbId);
          const akData  = akSS.getSheets()[0].getDataRange().getDisplayValues();
          const detSheet = akSS.getSheetByName("_Details");
          if (detSheet) {
            const detData = detSheet.getDataRange().getDisplayValues();
            let detMap = {};
            for (let d = 1; d < detData.length; d++) {
              const key = String(detData[d][1]).trim() + "|" + String(detData[d][2]).trim() + "|" + String(detData[d][10]).trim();
              detMap[key] = String(detData[d][9] || "").trim();
            }
            for (let a = 1; a < akData.length; a++) {
              const unit = String(akData[a][8]).trim();
              const spec = String(akData[a][9]).trim();
              const adv  = String(akData[a][10]).trim();
              akMap[String(akData[a][0]).trim()] = detMap[unit + "|" + spec + "|" + adv] || "";
            }
          }
        } catch(err) { Logger.log("Lookup Error in AK Shard for " + ayName + ": " + err.toString()); }

        const pmSheet = SpreadsheetApp.openById(infra.pmSheetId).getSheets()[0];
        const pmData  = pmSheet.getDataRange().getDisplayValues();

        for (let i = 1; i < pmData.length; i++) {
          if (String(pmData[i][0]).trim() === String(studentId).trim()) {
            if (String(pmData[i][24]).trim() === "Yes") continue;
            const akId = String(pmData[i][16]).trim();
            records.push({
              unit: pmData[i][2], level: pmData[i][3], term: pmData[i][4],
              assessorGrade: pmData[i][10], finalGrade: pmData[i][12],
              gDriveLink: akMap[akId] || "", akId: akId,
              dateSubmitted: pmData[i][26] || "", sheetId: infra.pmSheetId
            });
          }
        }
      } catch(err) { Logger.log("Error sweeping year " + ayName + ": " + err.toString()); }
    });

    return { records: records.reverse() };
  } catch(e) { return { error: e.toString() }; }
}

function updateSurgicalIndex(studentId, ayName, sheetId) {
  if (!studentId || studentId === "PENDING") return;
  // FIX: replaced inline open pattern with getStudentIndexSheet()
  const sheet = getStudentIndexSheet();
  const data  = sheet.getDataRange().getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(studentId).trim() &&
        String(data[i][1]).trim() === String(ayName).trim()) return;
  }

  sheet.appendRow([studentId, ayName, sheetId]);
}

function getStudentDeadlines(userToken, requestedName) {
  const auth = verifyAccess(userToken, ['Student', 'Adviser', 'Adviser/IV', 'Admin']);
  if (!auth.authorized) return { error: auth.message };

  try {
    let targetAkIds = [];

    if (auth.user.role === 'Student') {
      const studentId  = String(auth.user.email).toLowerCase().replace('@sgen.edu.ph', '').trim();
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

    const ss    = SpreadsheetApp.openById(akId);
    const sheet = ss.getSheets()[0];
    const data  = sheet.getDataRange().getDisplayValues();
    let deadlines = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    for (let i = 1; i < data.length; i++) {
      const rowAkId       = String(data[i][0]).trim();
      const rowStudentName = String(data[i][1]).trim().toLowerCase();
      const isMatch = (auth.user.role === 'Student')
        ? targetAkIds.includes(rowAkId)
        : rowStudentName === String(requestedName).trim().toLowerCase();

      if (isMatch) {
        const unit   = data[i][8];
        const draftStr = data[i][13];
        const finalStr = data[i][14];
        const status = String(data[i][29] || '').toLowerCase();

        if (status.includes('complete') || status.includes('drop')) continue;

        if (draftStr) {
          const dDate = new Date(draftStr);
          if (!isNaN(dDate.getTime()) && dDate >= today)
            deadlines.push({ unit, type: 'Draft', date: draftStr, timestamp: dDate.getTime() });
        }
        if (finalStr) {
          const fDate = new Date(finalStr);
          if (!isNaN(fDate.getTime()) && fDate >= today)
            deadlines.push({ unit, type: 'Final', date: finalStr, timestamp: fDate.getTime() });
        }
      }
    }

    deadlines.sort((a, b) => a.timestamp - b.timestamp);
    return deadlines.slice(0, 4);
  } catch(e) { return { error: e.toString() }; }
}

function generateInterimTranscriptPDF(token, requestedId) {
  const auth = verifyAccess(token, ['Admin', 'Adviser', 'Adviser/IV']);
  if (!auth.authorized) return { error: auth.message };

  try {
    const tempId = getConfig('InterimTranscriptTempID');
    if (!tempId) return { error: "InterimTranscriptTempID is missing in _Config." };

    const mSheet = getMasterSheet();
    const mData  = mSheet.getDataRange().getDisplayValues();

    let sName = "Unknown Student", sProg = "Unknown Programme";
    for (let i = 1; i < mData.length; i++) {
      if (String(mData[i][0]).trim() === String(requestedId).trim()) {
        const mi = mData[i][4] ? String(mData[i][4]).trim() + "." : "";
        sName = `${mData[i][1]}, ${mData[i][2]} ${mi}`.trim();
        sProg = mData[i][7] || "N/A";
        break;
      }
    }

    const gradeData = getStudentProgressionSurgical(token, requestedId);
    const records   = gradeData.records || [];

    const tempDocFile = DriveApp.getFileById(tempId).makeCopy('TEMP_Transcript_' + requestedId);
    const doc  = DocumentApp.openById(tempDocFile.getId());
    const body = doc.getBody();

    body.replaceText('<<STUDENT_NAME>>', sName);
    body.replaceText('<<STUDENT_ID>>', requestedId);
    body.replaceText('<<PROGRAMME>>', sProg);
    body.replaceText('<<DATE_PRINTED>>', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy"));

    const tables = body.getTables();
    let targetTable = null, templateRow = null, templateRowIndex = -1;

    for (let t = 0; t < tables.length; t++) {
      const tbl = tables[t];
      for (let r = 0; r < tbl.getNumRows(); r++) {
        if (tbl.getRow(r).getText().includes('<<T_UNIT>>')) {
          targetTable = tbl; templateRow = tbl.getRow(r); templateRowIndex = r; break;
        }
      }
      if (targetTable) break;
    }

    if (targetTable && templateRow) {
      records.forEach(rec => {
        let newRow = templateRow.copy();
        newRow.replaceText('<<T_UNIT>>',    String(rec.unit || '-'));
        newRow.replaceText('<<T_LVL>>',     String(rec.level || '-'));
        newRow.replaceText('<<T_TERM>>',    String(rec.term || '-'));
        newRow.replaceText('<<T_AGRADE>>', String(rec.assessorGrade || '-'));
        newRow.replaceText('<<T_FGRADE>>', String(rec.finalGrade || '-'));
        targetTable.insertTableRow(templateRowIndex, newRow);
        templateRowIndex++;
      });
      templateRow.removeFromParent();
    }

    doc.saveAndClose();
    const pdfBlob = tempDocFile.getAs('application/pdf');
    tempDocFile.setTrashed(true);

    return {
      success: true,
      fileName: `${sName.replace(/[^a-zA-Z0-9]/g, '_')}_Transcript.pdf`,
      base64: Utilities.base64Encode(pdfBlob.getBytes())
    };
  } catch(e) { return { error: "PDF Generation Failed: " + e.toString() }; }
}

/**
 * Bulk copies Assessor/IV grades to the Final Grade column.
 * MOVED HERE from end of old Code.txt — logically belongs in Grades.
 */
function batchPublishFinalGrades(token, payload) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { success: false, message: auth.message };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const sheetId      = payload.sheetId;
    const rowsToUpdate = payload.rowIndices;

    if (!sheetId || !rowsToUpdate || rowsToUpdate.length === 0) {
      return { success: false, message: "No sheet ID or rows provided." };
    }

    const sheet     = SpreadsheetApp.openById(sheetId).getSheets()[0];
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
    const editorTag = auth.user.name || "Unknown Admin";
    const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), 23);
    const data      = dataRange.getValues();
    let updateCount = 0;

    rowsToUpdate.forEach(rowIndex => {
      const arrIndex = rowIndex - 1;
      if (arrIndex > 0 && arrIndex < data.length) {
        const assessorGrade  = String(data[arrIndex][10] || "").trim();
        const ivGrade        = String(data[arrIndex][11] || "").trim();
        const gradeToPublish = ivGrade !== "" ? ivGrade : assessorGrade;
        if (gradeToPublish !== "") {
          data[arrIndex][12] = gradeToPublish;
          data[arrIndex][21] = editorTag;
          data[arrIndex][22] = timestamp;
          updateCount++;
        }
      }
    });

    if (updateCount > 0) {
      dataRange.setValues(data);
      return { success: true, message: `Successfully published Final Grades for ${updateCount} students!` };
    } else {
      return { success: false, message: "No valid grades were found in the selected rows." };
    }
  } catch (e) {
    return { success: false, message: "System Error: " + e.toString() };
  } finally { lock.releaseLock(); }
}

/**
 * Batch-updates submission dates from the Academic Progression modal.
 * MOVED HERE from end of old Code.txt — logically belongs in Grades.
 */
function batchUpdateRegistrarDates(token, updates) {
  try {
    const auth = verifyAccess(token, ['Admin']);
    if (!auth.authorized) return { success: false, message: auth.message };

    const DATE_SUBMITTED_COL = 27;
    const AK_ID_COL          = 17;

    const sheetGroups = {};
    updates.forEach(u => {
      if (!sheetGroups[u.sheetId]) sheetGroups[u.sheetId] = [];
      sheetGroups[u.sheetId].push(u);
    });

    let matchCount = 0;
    for (const sheetId in sheetGroups) {
      const targetSheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
      const dataRange   = targetSheet.getDataRange();
      const data        = dataRange.getDisplayValues();
      const sheetUpdates = sheetGroups[sheetId];

      sheetUpdates.forEach(update => {
        const targetAkId = String(update.akId).trim();
        for (let r = 1; r < data.length; r++) {
          if (String(data[r][AK_ID_COL - 1]).trim() === targetAkId) {
            targetSheet.getRange(r + 1, DATE_SUBMITTED_COL).setValue(update.dateSubmitted);
            matchCount++;
            break;
          }
        }
      });
    }

    return { success: true, message: `Successfully updated ${matchCount} submission date(s)!` };
  } catch (error) {
    return { success: false, message: "System Error: " + error.toString() };
  }
}
