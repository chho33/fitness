/**
 * Google Apps Script Web App backing the fitness diary page.
 *
 * Setup (one-time):
 *   1. Open the spreadsheet:
 *        https://docs.google.com/spreadsheets/d/13UlkWImCvyPmqpy3xwMpWBKyUIZnTqG56SkmIDmEZHY/edit
 *   2. Extensions -> Apps Script. Replace the default Code.gs with this file's contents.
 *   3. Deploy -> New deployment -> Type: Web app
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      Click Deploy, authorize when prompted, copy the Web app URL.
 *   4. Paste that URL into WEBAPP_URL in index.html, then commit & push.
 *
 * Re-deploy after editing: Deploy -> Manage deployments -> edit -> New version -> Deploy.
 *
 * Sheet schema (header row, auto-created on first append):
 *   who | date | name | type | detail | mood | timestamp
 */

const HEADER = ['who', 'date', 'name', 'type', 'detail', 'mood', 'timestamp'];

function doGet(e) {
  try {
    const action = (e.parameter.action || 'read').toLowerCase();
    const sheetName = e.parameter.sheet || '訓練記錄';
    if (action !== 'read') return json({ status: 'error', error: 'unknown action: ' + action });
    const sheet = getOrCreateSheet(sheetName);
    const values = sheet.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < values.length; i++) {
      const r = values[i];
      if (!r[0] && !r[1]) continue;
      rows.push({
        who: String(r[0] || ''),
        date: formatDate(r[1]),
        name: String(r[2] || ''),
        type: String(r[3] || ''),
        detail: String(r[4] || ''),
        mood: String(r[5] || '')
      });
    }
    return json({ status: 'ok', rows: rows });
  } catch (err) {
    return json({ status: 'error', error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = (body.action || '').toLowerCase();
    const sheetName = body.sheet || '訓練記錄';
    if (action !== 'append') return json({ status: 'error', error: 'unknown action: ' + action });
    if (!Array.isArray(body.row)) return json({ status: 'error', error: 'row must be an array' });
    const sheet = getOrCreateSheet(sheetName);
    sheet.appendRow(body.row);
    return json({ status: 'ok' });
  } catch (err) {
    return json({ status: 'error', error: String(err) });
  }
}

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(HEADER);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER);
  }
  return sheet;
}

function formatDate(v) {
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(v || '');
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
