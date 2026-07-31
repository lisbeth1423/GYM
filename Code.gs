const SPREADSHEET_ID = '1LpMD3SWI2TKDJCA3Q2x4oHEFU4V7xsFt1GoED6IVDTI';
const SHEET_NAME = 'Hoja 1';

function doGet(e) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift() || [];
  const rows = values
    .filter(row => row[0])
    .map(row => headers.reduce((item, header, index) => {
      item[header] = row[index];
      return item;
    }, {}));

  const result = JSON.stringify({ ok: true, records: rows });
  const callback = String((e && e.parameter && e.parameter.callback) || 'receiveGymHistory')
    .replace(/[^a-zA-Z0-9_.$]/g, '');

  return ContentService
    .createTextOutput(callback + '(' + result + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const payload = JSON.parse((e && e.parameter && e.parameter.payload) || '{}');
    const records = Array.isArray(payload.records) ? payload.records : [payload];
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const current = sheet.getDataRange().getValues();
    const rowById = {};

    for (let row = 1; row < current.length; row++) {
      if (current[row][0]) rowById[String(current[row][0])] = row + 1;
    }

    records.filter(record => record && record.id).forEach(record => {
      const values = [[
        String(record.id),
        record.timestamp || new Date().toISOString(),
        Number(record.week || 0),
        String(record.day || ''),
        String(record.exercise || ''),
        record.done === true,
        String(record.weight || ''),
        String(record.reps || '')
      ]];

      if (rowById[record.id]) {
        sheet.getRange(rowById[record.id], 1, 1, 8).setValues(values);
      } else {
        sheet.getRange(sheet.getLastRow() + 1, 1, 1, 8).setValues(values);
      }
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, saved: records.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
