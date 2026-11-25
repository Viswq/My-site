// Google Apps Script - stores blessings in a Google Sheet so EVERY guest sees them.
// Setup: see steps in chat. Paste this whole file in Extensions > Apps Script of a new Google Sheet.

const SHEET = 'Wishes';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(['ts', 'name', 'city', 'message']); }
  return sh;
}

// Page loads wishes: returns [{name, city, message, ts}, ...]
function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1)
    .map(r => ({ ts: Number(r[0]), name: String(r[1]), city: String(r[2]), message: String(r[3]) }))
    .filter(r => r.name && r.message)
    .slice(-300);
  return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
}

// Page sends a wish
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const name = String(d.name || '').slice(0, 60).trim();
    const city = String(d.city || '').slice(0, 60).trim();
    const message = String(d.message || '').slice(0, 300).trim();
    if (name && message) {
      const lock = LockService.getScriptLock(); lock.waitLock(10000);
      sheet_().appendRow([Date.now(), name, city, message]);
      lock.releaseLock();
    }
  } catch (err) {}
  return ContentService.createTextOutput('ok');
}
