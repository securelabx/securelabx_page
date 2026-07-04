// ============================================================
//  Leave Application — Google Apps Script (Code.gs)
//  Paste this entire file into the Apps Script editor.
//  Also create a separate Approve.html file in the editor.
// ============================================================

// ▶ Replace this with YOUR Google Sheet ID.
var SHEET_ID      = "12qW8VhSJMGx2JbHWg6QPTO2BhOhtLMhUw2nSwZzJEF0";

// ▶ Replace with your GitHub Pages Approve.html URL (e.g. https://username.github.io/repo/Approve.html)
var APPROVE_PAGE_URL = "https://qasehnuurul0103.github.io/claim-form/Approve.html";

var SHEET_NAME        = "Leave Applications";
var BALANCE_NAME      = "Leave Balance";
var EMPLOYEE_NAME     = "Employee";
var PASSWORD_NAME     = "Passwords";
var SETTINGS_NAME     = "Settings";
var DRIVE_FOLDER_NAME = "SLX_Leave_Documents";
var DRIVE_FOLDER_ID   = "169Ys_XTWqAXwDgjYNjik9a_AxbU9L2CQ"; // ▶ Paste folder ID from Drive URL

// ============================================================
//  Drive helpers
// ============================================================
function getOrCreateDriveFolder() {
  if (DRIVE_FOLDER_ID && DRIVE_FOLDER_ID !== "YOUR_DRIVE_FOLDER_ID_HERE") {
    return DriveApp.getFolderById(DRIVE_FOLDER_ID);
  }
  // Fallback: search by name, create if not found (My Drive only)
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function getOrCreateYearFolder(parentFolder) {
  var year       = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy");
  var subFolders = parentFolder.getFoldersByName(year);
  if (subFolders.hasNext()) return subFolders.next();
  return parentFolder.createFolder(year);
}

function generateAppId(sheet) {
  var tz     = Session.getScriptTimeZone();
  var today  = Utilities.formatDate(new Date(), tz, "yyyyMMdd");
  var prefix = "LA" + today;
  var lastCol0 = sheet.getLastColumn();
  var hdr    = lastCol0 > 0 ? sheet.getRange(1, 1, 1, lastCol0).getValues()[0] : [];
  var idCol  = hdr.indexOf("Application ID");
  var data   = sheet.getDataRange().getValues();
  var seq    = 0;
  if (idCol >= 0) {
    for (var i = data.length - 1; i >= 1; i--) {
      var id = (data[i][idCol] || "").toString();
      if (id.indexOf(prefix) === 0) {
        var n = parseInt(id.slice(prefix.length), 10);
        if (!isNaN(n)) { seq = n; break; }
      }
    }
  }
  return prefix + ("000" + (seq + 1)).slice(-3);
}

function sanitizeForFilename(str) {
  return (str || "").replace(/[^a-zA-Z0-9]/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "");
}

function buildFilenameBase(appId, staffId, leaveType) {
  return appId + "_" + sanitizeForFilename(staffId) + "_" + sanitizeForFilename(leaveType);
}

function uploadSupportingDoc(base64Data, appId, staffId, leaveType, folder) {
  var filename = buildFilenameBase(appId, staffId, leaveType) + "_Supporting.pdf";
  var blob     = Utilities.newBlob(Utilities.base64Decode(base64Data), "application/pdf", filename);
  return folder.createFile(blob).getUrl();
}

function generateDecisionPdf(rowData, hdr, appId, status, reason, managerName, decisionDate, folder) {
  function get(col) { var idx = hdr.indexOf(col); return idx >= 0 ? (rowData[idx] || "").toString() : ""; }

  var filename = buildFilenameBase(appId, get("Staff ID"), get("Leave Type")) + "_" + status + ".pdf";

  var months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  var ds = new Date(get("Submitted At"));
  var niceSubmitted = ds.getDate() + " " + months[ds.getMonth()] + " " + ds.getFullYear();
  var dd = new Date(decisionDate);
  var niceDecision  = dd.getDate() + " " + months[dd.getMonth()] + " " + dd.getFullYear();

  var leaveRows = [
    ["Leave Type", get("Leave Type")],
    ["Start Date", get("Start Date")],
    ["End Date",   get("End Date")],
    ["Total Days", get("Total Leave Days") + " working day(s)"],
    ["Reason",     get("Reason")],
  ];
  if (get("Compassionate Type")) leaveRows.splice(1, 0, ["Family Relation", get("Compassionate Type")]);
  if (get("Half Day") === "Yes") leaveRows.splice(1, 0, ["Half Day", "Yes – " + get("Session")]);
  if (status === "Rejected" && reason) leaveRows.push(["Rejection Reason", reason]);

  function detailRow(label, value) {
    return '<tr><td style="width:160px;color:#555;font-weight:600;padding:5px 0;vertical-align:top;">' +
      label + '</td><td style="color:#222;padding:5px 0;vertical-align:top;">' + (value || "") + '</td></tr>';
  }

  var empHtml   = [
    ["Name",       get("Employee Name")],
    ["Staff ID",   get("Staff ID")],
    ["Department", get("Department")],
    ["Position",   get("Position")],
    ["Email",      get("Email")],
    ["Phone",      get("Phone")],
  ].map(function(r){ return detailRow(r[0], r[1]); }).join("");

  var leaveHtml = leaveRows.map(function(r){ return detailRow(r[0], r[1]); }).join("");

  var sectionTitle = function(t) {
    return '<div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;' +
      'color:#1a73e8;border-bottom:2px solid #e8f0fe;padding-bottom:5px;margin:20px 0 10px;">' + t + '</div>';
  };

  var statusBg     = status === "Approved" ? "#e6f4ea" : "#fdecea";
  var statusBorder = status === "Approved" ? "#0f9d58" : "#e53935";
  var statusColor  = status === "Approved" ? "#137333" : "#c62828";

  var html =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
    '<body style="font-family:Arial,Helvetica,sans-serif;margin:0;padding:0;color:#222;font-size:13px;">' +

    '<div style="background:#0d47a1;color:#fff;padding:26px 36px;text-align:center;">' +
      '<div style="font-size:20px;font-weight:700;letter-spacing:1px;">LEAVE APPLICATION — ' + status.toUpperCase() + '</div>' +
      '<div style="font-size:12px;margin-top:7px;opacity:0.88;">Ref No: <strong>' + appId + '</strong> &nbsp;|&nbsp; Submitted: ' + niceSubmitted + '</div>' +
    '</div>' +

    '<div style="padding:24px 36px;">' +

    sectionTitle("Employee Information") +
    '<table style="width:100%;border-collapse:collapse;">' + empHtml + '</table>' +

    sectionTitle("Leave Details") +
    '<table style="width:100%;border-collapse:collapse;">' + leaveHtml + '</table>' +

    '<div style="background:' + statusBg + ';border:2px solid ' + statusBorder + ';border-radius:8px;padding:10px;' +
      'text-align:center;color:' + statusColor + ';font-weight:700;font-size:14px;margin:22px 0;">' +
      'Status: ' + status +
    '</div>' +

    '<table style="width:100%;border-collapse:collapse;margin-top:20px;"><tr>' +

    '<td style="width:50%;vertical-align:top;padding-right:24px;">' +
      '<div style="font-size:10px;color:#555;margin-bottom:3px;">Applicant Name:</div>' +
      '<div style="font-size:12px;font-weight:700;border-bottom:1px solid #999;padding-bottom:4px;margin-bottom:10px;min-height:22px;">' + get("Employee Name") + '</div>' +
      '<div style="font-size:10px;color:#555;margin-bottom:3px;">Date Applied:</div>' +
      '<div style="font-size:12px;font-weight:700;border-bottom:1px solid #999;padding-bottom:4px;min-height:22px;">' + niceSubmitted + '</div>' +
    '</td>' +

    '<td style="width:50%;vertical-align:top;">' +
      '<div style="font-size:10px;color:#555;margin-bottom:3px;">' + (status === "Approved" ? "Approved by:" : "Rejected by:") + '</div>' +
      '<div style="font-size:12px;font-weight:700;border-bottom:1px solid #999;padding-bottom:4px;margin-bottom:10px;min-height:22px;">' + (managerName || "") + '</div>' +
      '<div style="font-size:10px;color:#555;margin-bottom:3px;">Date ' + status + ':</div>' +
      '<div style="font-size:12px;font-weight:700;border-bottom:1px solid #999;padding-bottom:4px;min-height:22px;">' + niceDecision + '</div>' +
    '</td>' +

    '</tr></table>' +
    '</div></body></html>';

  var htmlBlob = Utilities.newBlob(html, MimeType.HTML, "__temp_decision_" + appId + ".html");
  var tempFile = DriveApp.createFile(htmlBlob);
  var pdfBlob  = tempFile.getAs(MimeType.PDF);
  pdfBlob.setName(filename);
  var pdfFile  = folder.createFile(pdfBlob);
  tempFile.setTrashed(true);

  return { url: pdfFile.getUrl(), blob: pdfBlob };
}

function ensureAppColumns(sheet) {
  var blue  = "#1a73e8";
  var white = "#ffffff";
  var lastCol = sheet.getLastColumn();

  // Sheet exists but is completely empty — initialise all headers
  if (lastCol < 1) {
    sheet.appendRow([
      "Application ID","Submitted At","Employee Name","Staff ID","Department","Position",
      "Email","Phone","Leave Type","Compassionate Type","Half Day",
      "Session","Start Date","End Date","Total Leave Days","Reason",
      "Status","Rejection Reason","Decision By","Decision Date",
      "Application Document","Supporting Document",
    ]);
    var hRange = sheet.getRange(1, 1, 1, 22);
    hRange.setBackground(blue).setFontColor(white).setFontWeight("bold");
    sheet.setFrozenRows(1);
    return;
  }

  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  if (hdr.indexOf("Application ID") === -1) {
    sheet.insertColumnBefore(1);
    var r = sheet.getRange(1, 1);
    r.setValue("Application ID");
    r.setBackground(blue).setFontColor(white).setFontWeight("bold");
    sheet.setColumnWidth(1, 160);
  }

  lastCol = sheet.getLastColumn();
  hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (hdr.indexOf("Application Document") === -1) {
    var c = lastCol + 1;
    var r2 = sheet.getRange(1, c);
    r2.setValue("Application Document");
    r2.setBackground(blue).setFontColor(white).setFontWeight("bold");
    sheet.setColumnWidth(c, 280);
  }

  lastCol = sheet.getLastColumn();
  hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (hdr.indexOf("Supporting Document") === -1) {
    var c2 = lastCol + 1;
    var r3 = sheet.getRange(1, c2);
    r3.setValue("Supporting Document");
    r3.setBackground(blue).setFontColor(white).setFontWeight("bold");
    sheet.setColumnWidth(c2, 280);
  }
}

// ============================================================
//  doGet
// ============================================================
function doGet(e) {
  if (!e || !e.parameter) {
    return ContentService.createTextOutput("Leave Application API is running.").setMimeType(ContentService.MimeType.TEXT);
  }

  // Serve the approval page
  if (e.parameter.page === "approve") {
    var template = HtmlService.createTemplateFromFile("Approve");
    template.scriptUrl = ScriptApp.getService().getUrl();
    template.ref       = e.parameter.ref || "";
    return template.evaluate()
      .setTitle("Leave Approval")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (e.parameter.action === "login")              return loginManager(e.parameter.staffId, e.parameter.password);
  if (e.parameter.action === "getPending")         return getPendingApplications(e.parameter.staffId, e.parameter.password);
  if (e.parameter.action === "getApplication")     return getApplication(e.parameter.ref, e.parameter.staffId, e.parameter.password);
  if (e.parameter.action === "updateStatus")       return updateLeaveStatus(e.parameter.ref, e.parameter.staffId, e.parameter.password, e.parameter.status, e.parameter.reason || "");
  if (e.parameter.action === "getBalance")         return getLeaveBalance(e.parameter.staffId, e.parameter.leaveType);
  if (e.parameter.action === "getEmployee")        return getEmployee(e.parameter.staffId);

  return ContentService.createTextOutput("Leave Application API is running.").setMimeType(ContentService.MimeType.TEXT);
}

// ============================================================
//  Authentication helpers
// ============================================================
function validateManager(staffId, password) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ensurePasswordSheet(ss);
  var data  = sheet.getDataRange().getValues();
  var lower = staffId.toString().toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase().trim() === lower) {
      return data[i][1].toString() === password.toString();
    }
  }
  return false;
}

function loginManager(staffId, password) {
  try {
    if (!staffId || !password) return json({ status: "error", message: "Staff ID and password required." });
    if (!validateManager(staffId, password)) return json({ status: "error", message: "Invalid Staff ID or password." });

    var ss      = SpreadsheetApp.openById(SHEET_ID);
    var empSheet = ensureEmployeeSheet(ss);
    var empData  = empSheet.getDataRange().getValues();
    var headers  = empData[0];
    var lower    = staffId.toLowerCase().trim();
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][0].toString().toLowerCase().trim() === lower) {
        return json({ status: "success", name: empData[i][headers.indexOf("Employee Name")] || staffId });
      }
    }
    return json({ status: "success", name: staffId });
  } catch (err) {
    return json({ status: "error", message: err.toString() });
  }
}

// ============================================================
//  getPendingApplications — returns all Pending leaves for
//  employees whose manager email matches the logged-in manager
// ============================================================
function getPendingApplications(staffId, password) {
  try {
    if (!validateManager(staffId, password)) return json({ status: "unauthorized" });

    var ss       = SpreadsheetApp.openById(SHEET_ID);
    var appSheet = ss.getSheetByName(SHEET_NAME);
    if (!appSheet) return json({ status: "success", applications: [] });

    var empSheet = ensureEmployeeSheet(ss);
    var empData  = empSheet.getDataRange().getValues();
    var empHdr   = empData[0];

    // Build map: staffId → managerEmail
    var managerMap = {};
    for (var i = 1; i < empData.length; i++) {
      managerMap[empData[i][0].toString().toLowerCase().trim()] =
        (empData[i][empHdr.indexOf("Manager Email")] || "").toString().toLowerCase().trim();
    }

    // Get this manager's email from Employee sheet
    var managerEmail = "";
    var lower = staffId.toLowerCase().trim();
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][0].toString().toLowerCase().trim() === lower) {
        managerEmail = (empData[i][empHdr.indexOf("Email")] || "").toString().toLowerCase().trim();
        break;
      }
    }

    var appData = appSheet.getDataRange().getValues();
    var appHdr  = appData[0];
    var results = [];

    for (var r = 1; r < appData.length; r++) {
      var row      = appData[r];
      var empId    = row[appHdr.indexOf("Staff ID")].toString().toLowerCase().trim();
      var status   = row[appHdr.indexOf("Status")].toString();
      var empMgr   = managerMap[empId] || "";

      if (status === "Pending" && empMgr === managerEmail) {
        var obj = {};
        appHdr.forEach(function(h, idx) { obj[h] = formatCell(row[idx]); });
        results.push(obj);
      }
    }

    return json({ status: "success", applications: results });
  } catch (err) {
    return json({ status: "error", message: err.toString() });
  }
}

// ============================================================
//  getApplication — returns a single application by ref
// ============================================================
function getApplication(ref, staffId, password) {
  try {
    if (!validateManager(staffId, password)) return json({ status: "unauthorized" });

    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return json({ status: "not_found" });

    var data = sheet.getDataRange().getValues();
    var hdr  = data[0];
    var refTrimmed = ref.toString().trim();

    var appIdCol = hdr.indexOf("Application ID");
    var lookupCol = appIdCol >= 0 ? appIdCol : 0;
    for (var i = 1; i < data.length; i++) {
      if (formatCell(data[i][lookupCol]) === refTrimmed) {
        var obj = {};
        hdr.forEach(function(h, idx) { obj[h] = formatCell(data[i][idx]); });
        return json({ status: "success", application: obj });
      }
    }
    return json({ status: "not_found" });
  } catch (err) {
    return json({ status: "error", message: err.toString() });
  }
}

// ============================================================
//  updateLeaveStatus — approve or reject a leave application
// ============================================================
function updateLeaveStatus(ref, staffId, password, status, reason) {
  try {
    if (!validateManager(staffId, password)) return json({ status: "unauthorized" });
    if (status !== "Approved" && status !== "Rejected") return json({ status: "error", message: "Invalid status." });

    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return json({ status: "error", message: "Applications sheet not found." });

    var data = sheet.getDataRange().getValues();
    var hdr  = data[0];

    // Ensure decision columns exist
    var statusIdx    = hdr.indexOf("Status");
    var rejIdx       = hdr.indexOf("Rejection Reason");
    var decByIdx     = hdr.indexOf("Decision By");
    var decDateIdx   = hdr.indexOf("Decision Date");

    var refTrimmed = ref.toString().trim();
    var appIdCol2  = hdr.indexOf("Application ID");
    var lookupCol2 = appIdCol2 >= 0 ? appIdCol2 : 0;
    for (var i = 1; i < data.length; i++) {
      if (formatCell(data[i][lookupCol2]) === refTrimmed) {
        var rowNum = i + 1;
        var now    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

        if (statusIdx  >= 0) sheet.getRange(rowNum, statusIdx  + 1).setValue(status);
        if (rejIdx     >= 0) sheet.getRange(rowNum, rejIdx     + 1).setValue(status === "Rejected" ? reason : "");
        if (decByIdx   >= 0) sheet.getRange(rowNum, decByIdx   + 1).setValue(staffId);
        if (decDateIdx >= 0) sheet.getRange(rowNum, decDateIdx + 1).setValue(now);

        // Colour-code the status cell
        if (statusIdx >= 0) {
          var cell = sheet.getRange(rowNum, statusIdx + 1);
          cell.setBackground(status === "Approved" ? "#e6f4ea" : "#fdecea");
          cell.setFontColor(status === "Approved" ? "#137333" : "#c62828");
          cell.setFontWeight("bold");
        }

        // Look up manager's display name
        var empSheet2    = ensureEmployeeSheet(ss);
        var empData2     = empSheet2.getDataRange().getValues();
        var empHdr2      = empData2[0];
        var managerDisplayName = staffId;
        var lowerMgr = staffId.toLowerCase().trim();
        for (var m = 1; m < empData2.length; m++) {
          if (empData2[m][0].toString().toLowerCase().trim() === lowerMgr) {
            managerDisplayName = empData2[m][empHdr2.indexOf("Employee Name")] || staffId;
            break;
          }
        }

        // Generate decision PDF
        var decisionPdfBlob = null;
        try {
          var rootFolder2 = getOrCreateDriveFolder();
          var yearFolder2 = getOrCreateYearFolder(rootFolder2);
          var appId2      = formatCell(data[i][lookupCol2]);
          var decResult   = generateDecisionPdf(data[i], hdr, appId2, status, reason, managerDisplayName, now, yearFolder2);
          decisionPdfBlob = decResult.blob;
        } catch (pdfErr) {
          Logger.log("Decision PDF failed: " + pdfErr.toString());
        }

        // Notify employee
        var empEmail  = data[i][hdr.indexOf("Email")] || "";
        var empName   = data[i][hdr.indexOf("Employee Name")] || "";
        var leaveType = data[i][hdr.indexOf("Leave Type")] || "";
        var startDate = data[i][hdr.indexOf("Start Date")] || "";
        var endDate   = data[i][hdr.indexOf("End Date")]   || "";
        if (empEmail) {
          sendDecisionEmail(empEmail, empName, leaveType, startDate, endDate, status, reason, managerDisplayName, decisionPdfBlob);
        }

        return json({ status: "success" });
      }
    }
    return json({ status: "not_found" });
  } catch (err) {
    return json({ status: "error", message: err.toString() });
  }
}

// ============================================================
//  sendDecisionEmail — notify employee and HR of the decision
// ============================================================
function sendDecisionEmail(email, name, leaveType, startDate, endDate, status, reason, decidedBy, pdfBlob) {
  var hrEmail    = getSetting("HR Email");
  var hrName     = getSetting("HR Name") || "HR";
  var company    = getSetting("Company Name") || "";
  var subject    = (company ? "[" + company + "] " : "") + "Leave Application " + status + " — " + leaveType;

  var employeeBody =
    "Dear " + name + ",\n\n" +
    (status === "Approved"
      ? "We are pleased to inform you that your leave application has been approved."
      : "We regret to inform you that your leave application has been rejected.") +
    "\n\n--- Leave Details ---\n" +
    "Leave Type   : " + leaveType + "\n" +
    "Start Date   : " + startDate + "\n" +
    "End Date     : " + endDate + "\n" +
    "Decision     : " + status + "\n" +
    (status === "Rejected" && reason ? "Reason       : " + reason + "\n" : "") +
    "\nDecision by  : " + decidedBy + "\n" +
    (status === "Rejected"
      ? "\n---\nIf you believe this decision was made in error, or if you would like to\n" +
        "reapply with revised dates or reason, please submit a new leave application.\n" +
        "Your original application has been kept on record.\n"
      : "") +
    "\nThank you.";

  var hrBody =
    "Dear " + hrName + ",\n\n" +
    "Please be informed that the following leave application has been " + status.toLowerCase() + ".\n\n" +
    "--- Leave Details ---\n" +
    "Employee     : " + name + "\n" +
    "Leave Type   : " + leaveType + "\n" +
    "Start Date   : " + startDate + "\n" +
    "End Date     : " + endDate + "\n" +
    "Decision     : " + status + "\n" +
    (status === "Rejected" && reason ? "Reason       : " + reason + "\n" : "") +
    "\nDecision by  : " + decidedBy + "\n\n" +
    "Please update the leave records accordingly.\n\nThank you.";

  var mailOpts = pdfBlob ? { attachments: [pdfBlob] } : {};

  try {
    GmailApp.sendEmail(email, subject, employeeBody, mailOpts);
  } catch (err) {
    Logger.log("Employee decision email failed: " + err.toString());
  }

  if (hrEmail) {
    try {
      GmailApp.sendEmail(hrEmail, subject, hrBody, mailOpts);
    } catch (err) {
      Logger.log("HR notification email failed: " + err.toString());
    }
  }
}

// ============================================================
//  doPost — receives the form submission
// ============================================================
function doPost(e) {
  try {
    var data  = JSON.parse(e.postData.contents);
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "Application ID", "Submitted At", "Employee Name", "Staff ID", "Department", "Position",
        "Email", "Phone", "Leave Type", "Compassionate Type", "Half Day",
        "Session", "Start Date", "End Date", "Total Leave Days", "Reason",
        "Status", "Rejection Reason", "Decision By", "Decision Date",
        "Application Document", "Supporting Document",
      ]);
      var hRange = sheet.getRange(1, 1, 1, 22);
      hRange.setBackground("#1a73e8");
      hRange.setFontColor("#ffffff");
      hRange.setFontWeight("bold");
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1,  160); sheet.setColumnWidth(2,  160); sheet.setColumnWidth(3,  180);
      sheet.setColumnWidth(4,  110); sheet.setColumnWidth(5,  150); sheet.setColumnWidth(6,  170);
      sheet.setColumnWidth(7,  200); sheet.setColumnWidth(8,  130); sheet.setColumnWidth(9,  150);
      sheet.setColumnWidth(10, 160); sheet.setColumnWidth(11,  90); sheet.setColumnWidth(12, 100);
      sheet.setColumnWidth(13, 110); sheet.setColumnWidth(14, 110); sheet.setColumnWidth(15, 130);
      sheet.setColumnWidth(16, 280); sheet.setColumnWidth(17, 100); sheet.setColumnWidth(18, 240);
      sheet.setColumnWidth(19, 120); sheet.setColumnWidth(20, 160); sheet.setColumnWidth(21, 280);
      sheet.setColumnWidth(22, 280);
    } else {
      ensureAppColumns(sheet);
    }

    var now    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var appId  = generateAppId(sheet);

    // Upload supporting document if provided
    var suppDocUrl = "";
    if (data.fileBase64) {
      var rootFolder = getOrCreateDriveFolder();
      var yearFolder = getOrCreateYearFolder(rootFolder);
      suppDocUrl = uploadSupportingDoc(data.fileBase64, appId, data.employeeId || "", data.leaveType || "Leave", yearFolder);
    }

    sheet.appendRow([
      appId,
      now,
      data.employeeName      || "",
      data.employeeId        || "",
      data.department        || "",
      data.position          || "",
      data.email             || "",
      data.phone             || "",
      data.leaveType         || "",
      data.compassionateType || "",
      data.halfDay           || "No",
      data.session           || "",
      data.startDate         || "",
      data.endDate           || "",
      data.totalDays         || "",
      data.reason            || "",
      "Pending",
      "", "", "",   // Rejection Reason, Decision By, Decision Date
      "",            // Application Document (generated on approval)
      suppDocUrl,
    ]);

    if (data.email) sendConfirmationEmail(data, appId, now);

    // Look up manager and send approval email
    var empSheet   = ensureEmployeeSheet(ss);
    var empData    = empSheet.getDataRange().getValues();
    var empHdr     = empData[0];
    var lowerEmpId = (data.employeeId || "").toString().toLowerCase().trim();
    for (var j = 1; j < empData.length; j++) {
      if (empData[j][0].toString().toLowerCase().trim() === lowerEmpId) {
        var managerName  = empData[j][empHdr.indexOf("Manager Name")]  || "";
        var managerEmail = empData[j][empHdr.indexOf("Manager Email")] || "";
        if (managerEmail) sendManagerApprovalEmail(data, appId, managerName, managerEmail);
        break;
      }
    }

    return json({ status: "success", appId: appId, appDocUrl: "", suppDocUrl: suppDocUrl });
  } catch (err) {
    return json({ status: "error", message: err.toString() });
  }
}

// ============================================================
//  sendManagerApprovalEmail
// ============================================================
function sendManagerApprovalEmail(data, appId, managerName, managerEmail) {
  var approvalUrl = APPROVE_PAGE_URL + "?ref=" + encodeURIComponent(appId);

  var subject = "[Action Required] Leave Approval — " + data.employeeName + " (" + data.leaveType + ")";
  var body =
    "Dear " + (managerName || "Manager") + ",\n\n" +
    data.employeeName + " has submitted a leave application that requires your approval.\n\n" +
    "--- Leave Details ---\n" +
    "Employee     : " + data.employeeName + " (" + data.employeeId + ")\n" +
    "Department   : " + data.department + "\n" +
    "Leave Type   : " + data.leaveType + "\n" +
    (data.compassionateType ? "Family       : " + data.compassionateType + "\n" : "") +
    (data.halfDay === "Yes"  ? "Half Day     : Yes (" + data.session + ")\n" : "") +
    "Start Date   : " + data.startDate + "\n" +
    "End Date     : " + data.endDate + "\n" +
    "Total Days   : " + data.totalDays + "\n" +
    "Reason       : " + data.reason + "\n\n" +
    "Click the link below to review and approve or reject this application:\n\n" +
    approvalUrl + "\n\n" +
    "Thank you.";
  try {
    GmailApp.sendEmail(managerEmail, subject, body);
  } catch (err) {
    Logger.log("Manager email failed: " + err.toString());
  }
}

// ============================================================
//  sendConfirmationEmail
// ============================================================
function sendConfirmationEmail(data, appId, submittedAt) {
  var subject = "Leave Application Received — " + data.leaveType;
  var body =
    "Dear " + data.employeeName + ",\n\n" +
    "Your leave application has been received and is pending approval.\n\n" +
    "--- Application Details ---\n" +
    "Application ID   : " + appId + "\n" +
    "Submitted At     : " + submittedAt + "\n" +
    "Staff ID         : " + data.employeeId + "\n" +
    "Department       : " + data.department + "\n" +
    "Leave Type       : " + data.leaveType + "\n" +
    "Start Date       : " + data.startDate + "\n" +
    "End Date         : " + data.endDate + "\n" +
    "Total Leave Days : " + data.totalDays + "\n" +
    "Reason           : " + data.reason + "\n\n" +
    "You will be notified once your manager has reviewed your application.\n\n" +
    "Thank you.";
  try {
    GmailApp.sendEmail(data.email, subject, body);
  } catch (err) {
    Logger.log("Email send failed: " + err.toString());
  }
}

// ============================================================
//  getEmployee
// ============================================================
function getEmployee(staffId) {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ensureEmployeeSheet(ss);
    var data  = sheet.getDataRange().getValues();
    var hdr   = data[0];
    var lower = staffId.toString().toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase().trim() === lower) {
        var row = data[i];
        return json({
          status:         "success",
          employeeName:   row[hdr.indexOf("Employee Name")]       || "",
          department:     row[hdr.indexOf("Department")]          || "",
          position:       row[hdr.indexOf("Position")]            || "",
          email:          row[hdr.indexOf("Email")]               || "",
          phone:          row[hdr.indexOf("Phone")]               || "",
          internDuration: row[hdr.indexOf("Internship Duration")] || "",
          managerName:    row[hdr.indexOf("Manager Name")]        || "",
          managerEmail:   row[hdr.indexOf("Manager Email")]       || "",
        });
      }
    }
    return json({ status: "not_found" });
  } catch (err) {
    return json({ status: "error", message: err.toString() });
  }
}

// ============================================================
//  getLeaveBalance
// ============================================================
function getLeaveBalance(staffId, leaveType) {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ensureBalanceSheet(ss);
    var data  = sheet.getDataRange().getValues();
    var hdr   = data[0];
    var leaveCol    = hdr.indexOf(leaveType);
    var durationCol = hdr.indexOf("Internship Duration");
    if (leaveCol === -1) return json({ status: "error", message: "Column not found: " + leaveType });

    var lower = staffId.toString().toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase().trim() === lower) {
        return json({ status: "success", balance: data[i][leaveCol], internDuration: durationCol !== -1 ? data[i][durationCol] : "" });
      }
    }
    return json({ status: "not_found" });
  } catch (err) {
    return json({ status: "error", message: err.toString() });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function formatCell(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  }
  return val === null || val === undefined ? "" : val.toString();
}

// ============================================================
//  Sheet creation helpers
// ============================================================
function ensureEmployeeSheet(ss) {
  var sheet = ss.getSheetByName(EMPLOYEE_NAME);
  if (sheet) return sheet;
  sheet = ss.insertSheet(EMPLOYEE_NAME);
  sheet.appendRow(["Staff ID","Employee Name","Department","Position","Email","Phone","Manager Name","Manager Email","Internship Duration"]);
  var h = sheet.getRange(1,1,1,9); h.setBackground("#0d47a1"); h.setFontColor("#fff"); h.setFontWeight("bold");
  sheet.setFrozenRows(1);
  [120,180,150,160,200,130,160,200,160].forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  return sheet;
}

function ensureBalanceSheet(ss) {
  var sheet = ss.getSheetByName(BALANCE_NAME);
  if (sheet) return sheet;
  sheet = ss.insertSheet(BALANCE_NAME);
  sheet.appendRow(["Staff ID","Employee Name","Annual Leave","Medical Leave","Maternity Leave","Paternity Leave","Compassionate Leave","Internship Duration"]);
  var h = sheet.getRange(1,1,1,8); h.setBackground("#0d47a1"); h.setFontColor("#fff"); h.setFontWeight("bold");
  sheet.setFrozenRows(1);
  [120,180,130,130,140,140,160,160].forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  return sheet;
}

// ============================================================
//  Settings helpers — read/write key-value config from sheet
// ============================================================
function getSetting(key) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ensureSettingsSheet(ss);
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toLowerCase() === key.toLowerCase()) {
      return data[i][1].toString().trim();
    }
  }
  return "";
}

function ensureSettingsSheet(ss) {
  var sheet = ss.getSheetByName(SETTINGS_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SETTINGS_NAME);
  sheet.appendRow(["Setting", "Value", "Description"]);
  sheet.appendRow(["HR Email", "", "Email address of the HR person who receives leave notifications"]);
  sheet.appendRow(["HR Name",  "", "Name of the HR person (used in email greeting)"]);
  sheet.appendRow(["Company Name", "", "Company name shown in email subjects"]);

  var hdr = sheet.getRange(1, 1, 1, 3);
  hdr.setBackground("#6a1b9a");
  hdr.setFontColor("#ffffff");
  hdr.setFontWeight("bold");
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 360);

  // Highlight the value column so it's easy to find
  sheet.getRange(2, 2, 10, 1).setBackground("#f3e5f5");

  return sheet;
}

function ensurePasswordSheet(ss) {
  var sheet = ss.getSheetByName(PASSWORD_NAME);
  if (sheet) return sheet;
  sheet = ss.insertSheet(PASSWORD_NAME);
  sheet.appendRow(["Staff ID","Password"]);
  var h = sheet.getRange(1,1,1,2); h.setBackground("#37474f"); h.setFontColor("#fff"); h.setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,120); sheet.setColumnWidth(2,160);
  return sheet;
}
