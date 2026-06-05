// Desert Salt Delivery — Google Apps Script
// Deploy as: Web App → Execute as: Me → Who has access: Anyone
// After editing, go to Deploy → Manage Deployments → edit → set to "New version" → Update

function doGet(e) {
  const ss = SpreadsheetApp.openById("1JlsEhO5Zvk_VcDBVkUsvfChsLNi4i7OdliVHtGe25tg");

  // Handle write operations sent as GET params
  if (e && e.parameter && e.parameter.write === "1") {
    try {
      const action  = e.parameter.action;
      const sheet   = e.parameter.sheet;
      const payload = JSON.parse(e.parameter.data || "[]");
      if (action === "set")    setSheet(ss, sheet, payload);
      if (action === "append") appendRow(ss, sheet, payload);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Normal read — return all sheets as JSON
  try {
    const data = {};
    ss.getSheets().forEach(sheet => {
      const name = sheet.getName();
      const rows = sheet.getDataRange().getValues();
      if (rows.length < 2) { data[name] = []; return; }
      const headers = rows[0];
      data[name] = rows.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      });
    });
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function setSheet(ss, sheetName, rows) {
  if (!Array.isArray(rows) || rows.length === 0) return;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clearContents();
  const headers = Object.keys(rows[0]);
  const values  = rows.map(row => headers.map(h => {
    const v = row[h];
    if (v === null || v === undefined) return "";
    if (typeof v === "object") return JSON.stringify(v);
    return v;
  }));
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (values.length > 0) {
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1a73e8");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function appendRow(ss, sheetName, row) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = Object.keys(row);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#1a73e8");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values  = headers.map(h => {
    const v = row[h];
    if (v === null || v === undefined) return "";
    if (typeof v === "object") return JSON.stringify(v);
    return v;
  });
  sheet.appendRow(values);
}
