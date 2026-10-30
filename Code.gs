/**
 * Wishes guestbook backend for the Anjana & Dipil wedding invitation.
 * Bound to a Google Sheet. Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
 * Columns: Timestamp | Name | Message | Show (checkbox). Untick Show to hide a wish on the site.
 */
const SHEET_NAME = 'Wishes';
const MAX_NAME = 80;
const MAX_MESSAGE = 700;
const MAX_ROWS_RETURNED = 300;
const CHECKBOX_ROWS = 1000;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    const sheets = ss.getSheets();
    sh = (sheets.length === 1 && sheets[0].getLastRow() === 0) ? sheets[0].setName(SHEET_NAME) : ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Name', 'Message', 'Show']]).setFontWeight('bold');
    sh.getRange(2, 4, CHECKBOX_ROWS, 1).insertCheckboxes();
    sh.setFrozenRows(1);
  }
  return sh;
}

// Stop spreadsheet formula injection (=, +, -, @ at the start of a cell).
function safe_(v, max) {
  v = String(v || '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET ?action=list -> newest first. Only rows with Show ticked are returned.
function doGet(e) {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return json_({ wishes: [] });
  const rows = sh.getRange(2, 1, last - 1, 4).getValues();
  const wishes = rows
    .filter(r => r[1] && r[2] && r[3] === true)
    .map(r => ({ time: r[0], name: String(r[1]).replace(/^'/, ''), message: String(r[2]).replace(/^'/, '') }))
    .reverse()
    .slice(0, MAX_ROWS_RETURNED);
  return json_({ wishes: wishes });
}

// POST {name, message}
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    const name = safe_(body.name, MAX_NAME);
    const message = safe_(body.message, MAX_MESSAGE);
    if (!name || !message) return json_({ ok: false, error: 'missing' });
    const sh = getSheet_();
    // Checkbox cells count as content, so find the first row with an empty Name instead of using appendRow.
    const names = sh.getRange(1, 2, Math.max(sh.getMaxRows(), 2), 1).getValues();
    let row = names.length + 1;
    for (let i = 1; i < names.length; i++) { if (!names[i][0]) { row = i + 1; break; } }
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
    sh.getRange(row, 1, 1, 3).setValues([[new Date(), name, message]]);
    sh.getRange(row, 4).insertCheckboxes().setValue(true);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
