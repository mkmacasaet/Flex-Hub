// ===================================================================
// FILE: 07_Helpdesk.gs
// PURPOSE: Support ticketing — create, read, and update tickets.
// ===================================================================

function getStudentTickets(userToken, requestedName, requestedEmail) {
  const auth = verifyAccess(userToken, ['Student', 'Adviser', 'Adviser/IV', 'Admin']);
  if (!auth.authorized) return { error: auth.message };

  const studentName  = (auth.user.role === 'Student') ? auth.user.name  : requestedName;
  const studentEmail = (auth.user.role === 'Student') ? auth.user.email : requestedEmail;

  try {
    const data = SpreadsheetApp.openById(getConfig('TicketingID')).getSheetByName("Monitoring").getDataRange().getValues();
    const myTickets = [];
    for (let i = 1; i < data.length; i++) {
      const matchName  = String(data[i][1]).toLowerCase() === String(studentName).toLowerCase();
      const matchEmail = studentEmail && String(data[i][13]).toLowerCase() === String(studentEmail).toLowerCase();
      if ((matchName || matchEmail) && data[i][7] !== 'Deleted') {
        myTickets.push({
          ticketId: data[i][0], summary: data[i][2], attachment: data[i][4],
          status: data[i][7],
          date: data[i][8] ? Utilities.formatDate(new Date(data[i][8]), Session.getScriptTimeZone(), "MM/dd/yyyy") : "",
          remarks: data[i][10]
        });
      }
    }
    return myTickets.reverse();
  } catch (err) { return []; }
}

function createStudentTicket(form) {
  try {
    const sheet = SpreadsheetApp.openById(getConfig('TicketingID')).getSheetByName("Monitoring");
    const ticketNumber = "TKT-" + new Date().getFullYear() + "-" + String(sheet.getLastRow()).padStart(3, '0');
    let attachmentUrl = "";
    if (form.fileData && form.fileName) {
      const file = DriveApp.getFolderById(getConfig('UploadFolderId'))
        .createFile(Utilities.newBlob(Utilities.base64Decode(form.fileData), form.mimeType, form.fileName));
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      attachmentUrl = file.getUrl();
    }
    sheet.appendRow([ticketNumber, form.requesterName, form.summary, form.notes, attachmentUrl,
      "", "", "Open", new Date(), "", "", "", "", form.requesterEmail || ""]);
    return { success: true, message: "Ticket " + ticketNumber + " created!" };
  } catch (error) { return { success: false, message: "Error: " + error.toString() }; }
}

function getAdminTickets(token) {
  const auth = verifyAccess(token, ['Admin']);
  if (!auth.authorized) return { error: auth.message };

  try {
    const id = getConfig('TicketingID');
    if (!id) return { error: "Configuration Error: TicketingID is missing in the _Config sheet." };
    const sheet = SpreadsheetApp.openById(id).getSheetByName("Monitoring");
    if (!sheet) return { error: "Database Error: Could not find 'Monitoring' tab." };
    if (sheet.getLastRow() < 2) return [];

    const data    = sheet.getDataRange().getDisplayValues();
    const tickets = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[7] !== 'Deleted') {
        tickets.push({
          rowIndex: i + 1, ticketId: row[0], requesitioner: row[1], summary: row[2], notes: row[3],
          attachment: row[4], simplicity: row[5], assigned: row[6], status: row[7], createdDate: row[8],
          dueDate: row[9], remarks: row[10], resolvedDate: row[11], duration: row[12],
          email: row[13], lastUpdated: row[14]
        });
      }
    }
    return tickets.reverse();
  } catch (e) { return { error: "Backend Error: " + e.toString() }; }
}

function saveAllTicket(token, data) {
  const auth = verifyAccess(token, ['Admin', 'Adviser', 'Adviser/IV', 'Student']);
  if (!auth.authorized) return { error: auth.message };

  if (data.rowIndex && data.rowIndex !== "") {
    if (auth.user.role !== 'Admin') {
      return { success: false, message: "SECURITY BREACH: You do not have permission to modify existing tickets." };
    }
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const id = getConfig('TicketingID');
    if (!id) return { success: false, message: "Ticketing ID missing." };
    const sheet        = SpreadsheetApp.openById(id).getSheetByName("Monitoring");
    const newTimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");

    let attachmentUrl = "";
    if (data.fileData && data.fileName) {
      const folderId = getConfig('UploadFolderId');
      if (folderId && folderId.trim() !== "") {
        const folder = DriveApp.getFolderById(folderId.trim());
        const file   = folder.createFile(Utilities.newBlob(Utilities.base64Decode(data.fileData), data.mimeType, data.fileName));
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        attachmentUrl = file.getUrl();
      }
    }

    if (!data.rowIndex || data.rowIndex === "") {
      const year         = new Date().getFullYear();
      const ticketNumber = "TKT-" + year + "-" + String(sheet.getLastRow() + 1).padStart(3, '0');
      sheet.appendRow([ticketNumber, data.requesitioner, data.summary, data.notes, attachmentUrl,
        data.simplicity || "", data.assigned || "", data.status || "Open", new Date(), data.dueDate || "",
        data.remarks || "", "", "", data.email || "", newTimestamp, "Created by: " + (data.editedBy || "Unknown")]);
    } else {
      const rowRange = sheet.getRange(parseInt(data.rowIndex), 1, 1, Math.max(sheet.getLastColumn(), 16));
      const rowData  = rowRange.getValues()[0];

      const sheetLastUpdated = String(rowData[14] || "");
      if (data.lastUpdated && sheetLastUpdated !== "" && String(data.lastUpdated) !== sheetLastUpdated) {
        return { success: false, message: "❌ SAVE REJECTED: Another Admin updated this ticket." };
      }

      rowData[1] = data.requesitioner; rowData[2] = data.summary; rowData[3] = data.notes;
      if (attachmentUrl !== "") rowData[4] = attachmentUrl;
      rowData[5] = data.simplicity || ""; rowData[6] = data.assigned;
      rowData[7] = data.status; rowData[9] = data.dueDate; rowData[10] = data.remarks;

      if (data.status === 'Resolved') {
        if (!rowData[11] || rowData[11] === "") rowData[11] = new Date();
        const createdDate  = new Date(rowData[8]);
        const resolvedDate = new Date(rowData[11]);
        if (!isNaN(createdDate.getTime()) && !isNaN(resolvedDate.getTime())) {
          rowData[12] = ((resolvedDate.getTime() - createdDate.getTime()) / (1000 * 60)).toFixed(0) + " mins";
        }
      } else { rowData[11] = ""; rowData[12] = ""; }

      if (data.email) rowData[13] = data.email;
      rowData[14] = newTimestamp;
      rowData[15] = "Edited by: " + (data.editedBy || "Unknown Admin");
      rowRange.setValues([rowData]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: "System busy. Please try again. " + e.toString() };
  } finally { lock.releaseLock(); }
}
