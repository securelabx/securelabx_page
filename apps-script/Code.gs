// SecureLabX form backend.
// Deploy this in a Google Sheet's Apps Script editor (Extensions > Apps Script),
// then deploy as a Web App (see README.md for full steps).
//
// It receives POST data from the site's forms and appends a row to the
// matching sheet ("Leave Request" or "Claim Request"), creating the sheet
// with headers on first use.

function doPost(e) {
  var params = e.parameter;
  var formType = params.formType || "Unknown";

  var sheet = getOrCreateSheet(formType, params);
  appendRow(sheet, formType, params);

  return ContentService
    .createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(formType, params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(formType);

  if (!sheet) {
    sheet = ss.insertSheet(formType);
    var headers = ["Timestamp"].concat(fieldNames(params));
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function fieldNames(params) {
  var names = [];
  for (var key in params) {
    if (key !== "formType") names.push(key);
  }
  return names;
}

function appendRow(sheet, formType, params) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = headers.map(function (header) {
    if (header === "Timestamp") return new Date();
    return params[header] || "";
  });
  sheet.appendRow(row);
}
