// ===================================================================
// FILE: 11_Settings.gs
// PURPOSE: System settings, dropdown lists, Academic Year management,
//          and calendar event operations.
// ===================================================================

function getProgramList() {
  try {
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_Dropdowns').getDataRange().getDisplayValues();
    const col  = data[0].findIndex(h => String(h).toLowerCase().includes('pathway'));
    return col > -1 ? [...new Set(data.slice(1).map(r => String(r[col]).trim()).filter(String))] : [];
  } catch(e) { return []; }
}

function getSpecificationList() {
  try {
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_Dropdowns').getDataRange().getDisplayValues();
    const col  = data[0].findIndex(h => String(h).toLowerCase().includes('spec'));
    return col > -1 ? [...new Set(data.slice(1).map(r => String(r[col]).trim()).filter(String))] : [];
  } catch(e) { return []; }
}

function getStatusList(token) {
  try {
    const ddSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_Dropdowns');
    if (!ddSheet) return [];
    const ddData  = ddSheet.getDataRange().getDisplayValues();
    const typeCol = ddData[0].findIndex(h => {
      const header = String(h).toLowerCase().trim();
      return header.includes('enrollment status') || header === 'status';
    });
    if (typeCol > -1) return [...new Set(ddData.slice(1).map(r => String(r[typeCol]).trim()).filter(String))];
    return [];
  } catch (e) { Logger.log("Error fetching Status list: " + e.toString()); return []; }
}

function refreshAcademicYearStatuses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_AcademicYears');
  if (!sheet || sheet.getLastRow() < 2) return;
  const data  = sheet.getDataRange().getValues();
  const today = new Date().getTime();

  for (let i = 1; i < data.length; i++) {
    const startDate = new Date(data[i][2]).getTime();
    const endDate   = new Date(data[i][3]).getTime();
    let status = "Inactive";
    if (today >= startDate && today <= endDate) status = "Active";
    else if (today < startDate) status = "Upcoming";
    else status = "Ended";
    sheet.getRange(i + 1, 5).setValue(status);
  }
}

function getSystemSettingsData(moduleName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const dropdownMaps = {
      'term': 'Term', 'specs': 'Specifications', 'pathway': 'Pathway',
      'unit-status': 'Unit Status', 'assignee': 'Ticket Assignee',
      'simplicity': 'Ticket Simplicity', 'enrollment': 'Enrollment Status'
    };

    if (dropdownMaps[moduleName]) {
      const sheet = ss.getSheetByName('_Dropdowns');
      if (!sheet) return { success: false, message: "Missing '_Dropdowns' sheet." };
      const data        = sheet.getDataRange().getDisplayValues();
      const headers     = data[0].map(h => String(h).trim().toLowerCase());
      const targetHeader = dropdownMaps[moduleName].toLowerCase();
      const colIndex    = headers.indexOf(targetHeader);
      if (colIndex === -1) return { success: false, message: `Could not find column named '${dropdownMaps[moduleName]}'` };
      let listData = [];
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][colIndex]).trim() !== "") listData.push({ value: data[i][colIndex] });
      }
      return { success: true, type: 'dropdown', data: listData };
    }

    if (moduleName === 'ay-pm') {
      const sheet = ss.getSheetByName('_AcademicYears');
      if (!sheet) return { success: false, message: "Missing '_AcademicYears' sheet." };
      const data     = sheet.getDataRange().getDisplayValues();
      let tableData  = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          tableData.push({
            ay: data[i][0], sheetId: data[i][1], startDate: data[i][2],
            endDate: data[i][3], status: data[i][4] || 'Active'
          });
        }
      }
      return { success: true, type: 'ay', data: tableData };
    }

    if (moduleName === 'config') {
      const sheet = ss.getSheetByName('_Config');
      if (!sheet) return { success: false, message: "Missing '_Config' sheet." };
      const data    = sheet.getDataRange().getDisplayValues();
      let tableData = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) tableData.push({ name: data[i][0], details: data[i][1] });
      }
      return { success: true, type: 'config', data: tableData };
    }

    return { success: false, message: "Unknown Module Requested." };
  } catch (err) { return { success: false, message: "System Error: " + err.message }; }
}

function addCalendarEventAdmin(userToken, payload) {
  const auth = verifyAccess(userToken, ['Admin']);
  if (!auth.authorized) return { success: false, message: auth.message };
  try {
    const calId  = getConfig('CalendarID');
    if (!calId) return { success: false, message: "CalendarID missing in config!" };
    const sheet  = SpreadsheetApp.openById(calId).getSheets()[0];
    const eventId = "EVT-" + new Date().getFullYear() + "-" + Math.floor(Math.random() * 9000 + 1000);
    sheet.appendRow([eventId, payload.date, payload.time, payload.title]);
    return { success: true, message: "Event added to calendar!" };
  } catch(e) { return { success: false, message: "Error: " + e.message }; }
}

function deleteCalendarEventAdmin(userToken, rowIndex) {
  const auth = verifyAccess(userToken, ['Admin']);
  if (!auth.authorized) return { success: false, message: auth.message };
  try {
    const calId = getConfig('CalendarID');
    SpreadsheetApp.openById(calId).getSheets()[0].deleteRow(rowIndex);
    return { success: true, message: "Event removed from calendar." };
  } catch(e) { return { success: false, message: "Error: " + e.message }; }
}
