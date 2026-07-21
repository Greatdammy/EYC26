/**
 * EYC26 registration backend — paste this into the Apps Script editor
 * attached to your Google Sheet (Extensions > Apps Script), then deploy
 * as a Web App. See DEPLOY.md for step-by-step instructions.
 */

var SHEET_NAME = "Registrations";
var HEADERS = ["Timestamp","Name","Email","Phone","Age","Stage","Church","Source","Expectation","Gadgets","Extra","Serial"];
var TOTAL_SEATS = 120;
// Must match CFG.GAS_TOKEN in index.html — a soft barrier against direct API spam, not a real secret.
var SHARED_TOKEN = "eyc26_063bca0bcda4e9439e9e92c849b5faa9";

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var sheet = getSheet_();
  var action = e.parameter.action;

  if (action === "count") {
    var count = Math.max(sheet.getLastRow() - 1, 0);
    return jsonOut_({ count: count });
  }

  if (action === "check") {
    var email = String(e.parameter.email || "").trim().toLowerCase();
    var data = sheet.getDataRange().getValues();
    var emailCol = HEADERS.indexOf("Email");
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emailCol]).trim().toLowerCase() === email) {
        return jsonOut_({ exists: true });
      }
    }
    return jsonOut_({ exists: false });
  }

  if (action === "verify") {
    var serial = String(e.parameter.serial || "").trim();
    if (!serial) {
      return jsonOut_({ valid: false });
    }
    var vData = sheet.getDataRange().getValues();
    var serialCol = HEADERS.indexOf("Serial");
    var nameCol = HEADERS.indexOf("Name");
    var churchCol = HEADERS.indexOf("Church");
    var tsCol = HEADERS.indexOf("Timestamp");
    for (var j = 1; j < vData.length; j++) {
      if (String(vData[j][serialCol]).trim() === serial) {
        return jsonOut_({
          valid: true,
          name: vData[j][nameCol],
          church: vData[j][churchCol],
          serial: vData[j][serialCol],
          spot: j,
          total: TOTAL_SEATS,
          timestamp: vData[j][tsCol]
        });
      }
    }
    return jsonOut_({ valid: false });
  }

  return jsonOut_({ error: "unknown action" });
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);

  if (body.Token !== SHARED_TOKEN) {
    return jsonOut_({ ok: false, error: "unauthorized" });
  }

  var email = String(body.Email || "").trim().toLowerCase();
  if (!email) {
    return jsonOut_({ ok: false, error: "invalid" });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = getSheet_();

    var count = Math.max(sheet.getLastRow() - 1, 0);
    if (count >= TOTAL_SEATS) {
      return jsonOut_({ ok: false, error: "full", count: count });
    }

    var data = sheet.getDataRange().getValues();
    var emailCol = HEADERS.indexOf("Email");
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emailCol]).trim().toLowerCase() === email) {
        return jsonOut_({ ok: false, error: "duplicate" });
      }
    }

    sheet.appendRow([
      body.Timestamp || new Date().toISOString(),
      body.Name || "",
      body.Email || "",
      body.Phone || "",
      body.Age || "",
      body.Stage || "",
      body.Church || "",
      body.Source || "",
      body.Expectation || "",
      body.Gadgets || "",
      body.Extra || "",
      body.Serial || ""
    ]);

    return jsonOut_({ ok: true, count: count + 1 });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Maintenance only — NOT web-exposed. Run manually from the Apps Script
 * editor (select this function in the dropdown, then click Run) to wipe
 * all data rows and start the seat count back at 0 / 120 remaining.
 */
function resetTestData() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}
