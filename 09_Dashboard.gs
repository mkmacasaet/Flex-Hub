// ===================================================================
// FILE: 09_Dashboard.gs
// PURPOSE: Dashboard aggregators, news, and grade summary for all roles.
// ===================================================================

function getAdviserDashboardData(userToken, requestedName, requestedEmail, requestedRole) {
  const auth = verifyAccess(userToken, ['Adviser', 'Adviser/IV', 'Admin']);
  if (!auth.authorized) return { error: auth.message };

  const name  = auth.user.role.includes('Adviser') ? auth.user.name  : requestedName;
  const email = auth.user.role.includes('Adviser') ? auth.user.email : requestedEmail;

  try {
    let payload = { news: [], quickLinks: [], actionCenter: { deferrals: 0, ungraded: 0 }, events: [] };

    const dashboard = getDashboardData(userToken);
    if (dashboard && dashboard.news) payload.news = dashboard.news;

    const myDeferrals = getAdviserDeferrals(userToken, name, email);
    if (myDeferrals && !myDeferrals.error) {
      payload.actionCenter.deferrals = myDeferrals.filter(d => d.status === 'Pending').length;
    }

    const ays      = getAcademicYears(userToken);
    const activeAy = ays.find(ay => ay.status === 'Active');
    if (activeAy && activeAy.id) {
      const grades = getGradesForAY(userToken, activeAy.id);
      if (grades && !grades.error) {
        grades.forEach(r => {
          if (r.adviserName === name && (!r.assessor || r.assessor.trim() === '')) {
            payload.actionCenter.ungraded++;
          }
        });
      }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let linksSheet = ss.getSheetByName('_AdviserLinks');
    if (!linksSheet) {
      linksSheet = ss.insertSheet('_AdviserLinks');
      linksSheet.appendRow(['Title', 'URL', 'Icon']);
      linksSheet.appendRow(['Google Drive', 'https://drive.google.com', 'folder_shared']);
      linksSheet.appendRow(['Faculty Handbook', 'https://docs.google.com', 'menu_book']);
      linksSheet.appendRow(['IT Support', 'mailto:it@school.edu', 'contact_support']);
      linksSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#f8fafc');
    }
    if (linksSheet.getLastRow() > 1) {
      const linkData = linksSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < linkData.length; i++) {
        if (linkData[i][0] !== '') {
          payload.quickLinks.push({ title: linkData[i][0], url: linkData[i][1], icon: linkData[i][2] || 'link' });
        }
      }
    }

    try {
      const calId = getConfig('CalendarID');
      if (calId) {
        const calSheet = SpreadsheetApp.openById(calId).getSheets()[0];
        if (calSheet.getLastRow() > 1) {
          const calData = calSheet.getDataRange().getDisplayValues();
          for (let i = 1; i < calData.length; i++) {
            if (calData[i][0]) {
              payload.events.push({ id: calData[i][0], date: calData[i][1], time: calData[i][2], title: calData[i][3] });
            }
          }
        }
      }
    } catch(e) { Logger.log("Adviser Calendar fetch error: " + e.message); }

    return payload;
  } catch(e) { return { error: "Failed to aggregate dashboard: " + e.toString() }; }
}

function getDashboardData(userToken) {
  const auth = verifyAccess(userToken, ['Student', 'Adviser', 'Adviser/IV', 'Admin']);
  if (!auth.authorized) return { error: auth.message };

  const newsId = getConfig('NewsDatabaseId');
  if (!newsId) return { news: [] };

  try {
    const ss       = SpreadsheetApp.openById(newsId);
    const allNews  = ss.getSheets()[0].getDataRange().getValues().slice(1).reverse().slice(0, 10);
    const myNews   = allNews
      .filter(r => r[4] === 'All' || r[4] === 'Students')
      .map(r => ({
        date: new Date(r[0]).toLocaleDateString(), author: r[1],
        title: r[2], content: r[3], attachment: r[5]
      }));
    return { news: myNews };
  } catch(e) { return { news: [] }; }
}

function getMyGrades(userToken, requestedId) {
  const auth = verifyAccess(userToken, ['Student', 'Adviser', 'Adviser/IV', 'Admin']);
  if (!auth.authorized) return { error: auth.message };

  let targetId = requestedId;
  if (auth.user.role === 'Student') {
    targetId = String(auth.user.email).toLowerCase().replace('@sgen.edu.ph', '').trim();
  }

  try {
    const aySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_AcademicYears');
    if (!aySheet) return { error: "Directory missing." };
    const ays = aySheet.getRange(2, 1, aySheet.getLastRow() - 1, 5).getValues();

    let records = [], completedUnits = 0, totalCredits = 0;

    ays.forEach(ay => {
      if (ay[1]) {
        try {
          const sheet = SpreadsheetApp.openById(ay[1]).getSheets()[0];
          const data  = sheet.getDataRange().getValues();
          for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim() === String(targetId).trim()) {
              const fg   = data[i][10] || data[i][12] || "";
              const cred = parseInt(data[i][5]) || 0;
              records.push({ unit: data[i][2], level: data[i][3], term: data[i][4], credits: data[i][5], grade: fg, comments: data[i][15] });
              if (fg.charAt(0) === 'P' || fg.charAt(0) === 'M' || fg.charAt(0) === 'D') {
                completedUnits++; totalCredits += cred;
              }
            }
          }
        } catch(e) {}
      }
    });
    return { records, completedUnits, totalCredits };
  } catch(e) { return { error: e.toString() }; }
}

function getAllNewsAdmin() {
  const newsId = getConfig('NewsDatabaseId');
  if (!newsId) return [];
  try {
    const sheet = SpreadsheetApp.openById(newsId).getSheets()[0];
    const data  = sheet.getDataRange().getValues();
    return data.slice(1).map((r, i) => ({
      rowIndex: i + 2, date: new Date(r[0]).toLocaleDateString(),
      author: r[1], title: r[2], content: r[3], audience: r[4]
    })).reverse();
  } catch(e) { return []; }
}

function adminPostNews(form) {
  const newsId = getConfig('NewsDatabaseId');
  if (!newsId) return { success: false, message: "Configuration Error: NewsDatabaseId missing." };

  let fileUrl = "";
  try {
    if (form.fileData && form.fileName) {
      const folder = DriveApp.getFolderById(getConfig('UploadFolderId').trim());
      const file   = folder.createFile(Utilities.newBlob(Utilities.base64Decode(form.fileData), form.mimeType, form.fileName));
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }
  } catch (e) { return { success: false, message: "File upload failed: " + e.toString() }; }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(newsId).getSheets()[0];
    if (form.rowIndex && form.rowIndex !== "") {
      const rowRange = sheet.getRange(parseInt(form.rowIndex), 1, 1, Math.max(sheet.getLastColumn(), 6));
      const rowData  = rowRange.getValues()[0];
      rowData[2] = form.title; rowData[3] = form.content; rowData[4] = form.audience;
      if (fileUrl !== "") rowData[5] = fileUrl;
      rowRange.setValues([rowData]);
      return { success: true, message: "Announcement Updated!" };
    } else {
      sheet.appendRow([new Date(), form.author, form.title, form.content, form.audience, fileUrl]);
      return { success: true, message: "Announcement Posted!" };
    }
  } catch (e) {
    return { success: false, message: "System busy. Please try again." };
  } finally { lock.releaseLock(); }
}

function deleteNewsAdmin(token, rowIndex) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };
  const newsId = getConfig('NewsDatabaseId');
  if (!newsId) return { success: false, message: "Configuration Error." };
  try {
    SpreadsheetApp.openById(newsId).getSheets()[0].deleteRow(rowIndex);
    return { success: true, message: "Announcement Deleted!" };
  } catch(e) { return { success: false, message: "Error deleting announcement." }; }
}

function getAdminDashboardData(userToken) {
  const auth = verifyAccess(userToken, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  const stats = {
    students: { total: 0, active: 0, spillOver: 0, graduates: 0, loa: 0, atRisk: 0, rawList: [] },
    helpdesk: { pending: 0, total: 0, avgMins: 0, complexity: {} },
    system: { topLogins: [], newsCount: 0 },
    act: { deferrals: 0, ungraded: 0, warnings: 0 },
    news: [], quickLinks: [], events: []
  };

  try {
    const masterId = getConfig('MasterStudentListId');
    const ticketId = getConfig('TicketingID');

    if (masterId) {
      try {
        const sheet = SpreadsheetApp.openById(masterId).getSheets()[0];
        if (sheet.getLastRow() > 1) {
          const data    = sheet.getDataRange().getDisplayValues();
          const headers = data[0].map(h => String(h).toLowerCase().trim());
          const cPath   = headers.indexOf("programme") > -1 ? headers.indexOf("programme") : 7;
          const cSpec   = headers.indexOf("specification") > -1 ? headers.indexOf("specification") : 8;
          const cGen    = headers.indexOf("gender") > -1 ? headers.indexOf("gender") : 9;
          const cAge    = headers.indexOf("age") > -1 ? headers.indexOf("age") : 15;
          const cNat    = headers.indexOf("nationality") > -1 ? headers.indexOf("nationality") : 16;
          const cCity   = headers.indexOf("city") > -1 ? headers.indexOf("city") : 18;
          const cCountry = headers.indexOf("country") > -1 ? headers.indexOf("country") : 19;
          const cStat   = headers.indexOf("enrollment status") > -1 ? headers.indexOf("enrollment status") : 23;
          const cSpecial = headers.indexOf("special need") > -1 ? headers.indexOf("special need") : 25;

          for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (!row[0]) continue;
            stats.students.total++;
            const currentStat = row[cStat] ? String(row[cStat]).trim() : "Unknown";
            const lowerStat   = currentStat.toLowerCase();
            if (lowerStat === 'active') stats.students.active++;
            else if (lowerStat.includes('spill')) stats.students.spillOver++;
            else if (lowerStat.includes('graduat') || lowerStat === 'completed') stats.students.graduates++;
            else if (lowerStat === 'official loa' || lowerStat.includes('unofficial loa')) stats.students.loa++;
            else if (lowerStat.includes('risk')) stats.students.atRisk++;
            const mi       = row[4] ? String(row[4]).trim() + "." : "";
            const fullName = `${row[1] || ''}, ${row[2] || ''} ${mi}`.trim();
            stats.students.rawList.push({
              id: String(row[0] || ''), name: fullName, email: String(row[11] || ''),
              pathway: row[cPath] ? String(row[cPath]).trim() : "Unknown",
              spec: row[cSpec] ? String(row[cSpec]).trim() : "Unknown",
              gender: row[cGen] ? String(row[cGen]).trim() : "Unknown",
              nationality: row[cNat] ? String(row[cNat]).trim() : "Unknown",
              country: row[cCountry] ? String(row[cCountry]).trim() : "Unknown",
              status: currentStat, city: row[cCity] ? String(row[cCity]).trim() : "Unknown",
              age: parseInt(row[cAge]) || 0,
              specialNeed: row[cSpecial] ? String(row[cSpecial]).trim() : "",
              progLink: row[21] ? String(row[21]).trim() : "",
              folderLink: row[22] ? String(row[22]).trim() : ""
            });
          }
        }
      } catch(e) {}

      try {
        const riskReport = getRiskAssessmentReport(userToken);
        if (riskReport && riskReport.success) {
          const pendingWarnings = riskReport.warnings    ? riskReport.warnings.length    : 0;
          const activeCases     = riskReport.flaggedCases ? riskReport.flaggedCases.length : 0;
          stats.act.warnings = pendingWarnings + activeCases;
        } else { stats.act.warnings = 0; }
      } catch(e) { Logger.log("Failed to load risk report for Dashboard: " + e.toString()); stats.act.warnings = 0; }
    }

    if (ticketId) {
      try {
        const sheet = SpreadsheetApp.openById(ticketId).getSheetByName("Monitoring");
        if (sheet && sheet.getLastRow() > 1) {
          let totalMins = 0, resolvedCount = 0;
          stats.helpdesk.total = sheet.getLastRow() - 1;
          sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues().forEach(row => {
            if (row[7] !== 'Resolved') stats.helpdesk.pending++;
            if (row[7] === 'Resolved' && String(row[12]).includes('mins')) { totalMins += parseInt(row[12]); resolvedCount++; }
          });
          stats.helpdesk.avgMins = resolvedCount > 0 ? Math.round(totalMins / resolvedCount) : 0;
        }
      } catch(e) {}
    }

    const dashboard = getDashboardData(userToken);
    if (dashboard && dashboard.news) stats.news = dashboard.news;

    const allDeferrals = getDeferralsAdmin(userToken);
    if (allDeferrals && allDeferrals.stats) stats.act.deferrals = allDeferrals.stats.pending;

    const ays      = getAcademicYears(userToken);
    const activeAy = ays.find(ay => ay.status === 'Active');
    if (activeAy && activeAy.id) {
      const grades = getGradesForAY(userToken, activeAy.id);
      if (grades && !grades.error) {
        grades.forEach(r => {
          if (!r.assessor || String(r.assessor).trim() === '') stats.act.ungraded++;
        });
      }
    }

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let linksSheet = ss.getSheetByName('_AdviserLinks');
      if (linksSheet && linksSheet.getLastRow() > 1) {
        const linkData = linksSheet.getDataRange().getDisplayValues();
        for (let i = 1; i < linkData.length; i++) {
          if (linkData[i][0] !== '') {
            stats.quickLinks.push({ title: linkData[i][0], url: linkData[i][1], icon: linkData[i][2] || 'link' });
          }
        }
      }
    } catch(e) {}

    try {
      const uS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_Users');
      if (uS && uS.getLastRow() > 1) {
        stats.system.topLogins = uS.getRange(2, 5, uS.getLastRow() - 1, 3).getValues()
          .filter(r => r[2]).sort((a, b) => new Date(b[2]) - new Date(a[2]))
          .slice(0, 5).map(l => ({ name: l[0], time: new Date(l[2]).toISOString() }));
      }
      const nId = getConfig('NewsDatabaseId');
      const nS  = SpreadsheetApp.openById(nId).getSheets()[0];
      if (nS) stats.system.newsCount = Math.max(0, nS.getLastRow() - 1);
    } catch(e) {}

    try {
      const calId = getConfig('CalendarID');
      if (calId) {
        const calSheet = SpreadsheetApp.openById(calId).getSheets()[0];
        if (calSheet.getLastRow() > 1) {
          const calData = calSheet.getDataRange().getDisplayValues();
          for (let i = 1; i < calData.length; i++) {
            if (calData[i][0]) {
              stats.events.push({ id: calData[i][0], date: calData[i][1], time: calData[i][2], title: calData[i][3], rowIndex: i + 1 });
            }
          }
        }
      }
    } catch(e) { Logger.log("Calendar fetch error: " + e.message); }

  } catch(e) { Logger.log("Admin Dashboard Error: " + e.message); }

  return stats;
}
