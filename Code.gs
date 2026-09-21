const SPREADSHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
const SHEET_NAME = 'Schedule';

function doGet(e) {
  const key = String(e.parameter.key || '');
  const result = { state: readState_(key) };
  const output = JSON.stringify(result);

  if (e.parameter.callback) {
    return ContentService
      .createTextOutput(e.parameter.callback + '(' + output + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  const key = String(payload.key || '');
  const state = normalizeState_(payload.state);

  if (!key) throw new Error('A week key is required.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();
    const rowIndex = values.findIndex(row => row[0] === key);
    const row = [key, JSON.stringify(state.morning), JSON.stringify(state.evening), new Date()];

    if (rowIndex === -1) {
      sheet.appendRow(row);
    } else {
      sheet.getRange(rowIndex + 1, 1, 1, row.length).setValues([row]);
    }
  } finally {
    lock.releaseLock();
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['weekKey', 'morning', 'evening', 'updatedAt']);
  }
  return sheet;
}

function readState_(key) {
  if (!key) return null;

  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const row = values.find(candidate => candidate[0] === key);
  if (!row) return null;

  return normalizeState_({
    morning: JSON.parse(row[1]),
    evening: JSON.parse(row[2])
  });
}

function normalizeState_(state) {
  const blank = () => Array(7).fill('');
  return {
    morning: Array.isArray(state && state.morning) && state.morning.length === 7 ? state.morning : blank(),
    evening: Array.isArray(state && state.evening) && state.evening.length === 7 ? state.evening : blank()
  };
}