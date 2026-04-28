// Mo摸 & 何帥 健身日記 — Google Apps Script
// 部署方式：擴充功能 → Apps Script → 貼上此程式碼 → 部署 → 新增部署 → 類型選「網頁應用程式」
// 執行身份：「我」、存取權：「任何人」

const SPREADSHEET_ID = '13UlkWImCvyPmqpy3xwMpWBKyUIZnTqG56SkmIDmEZHY';

function doGet(e) {
  const params = e.parameter;
  const action = params.action;
  const sheetName = params.sheet || '訓練記錄';

  try {
    if (action === 'read') {
      return handleRead(sheetName);
    }
    return jsonResponse({ status: 'error', error: 'Unknown GET action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const sheetName = body.sheet || '訓練記錄';

    if (action === 'append') {
      return handleAppend(sheetName, body.row);
    }
    if (action === 'delete') {
      return handleDelete(sheetName, body.rowIndex);
    }
    return jsonResponse({ status: 'error', error: 'Unknown POST action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', error: err.message });
  }
}

// ── READ ──────────────────────────────────────────────────────────────────────
// Returns all rows with their actual sheet row index (1-based, skipping header).
function handleRead(sheetName) {
  const sheet = getOrCreateSheet(sheetName);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    // Only header row or empty
    return jsonResponse({ status: 'ok', rows: [] });
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const rows = data
    .map((row, i) => ({
      rowIndex: i + 2, // actual sheet row number (header is row 1)
      who:    row[0] || '',
      date:   row[1] ? formatDate(row[1]) : '',
      name:   row[2] || '',
      type:   row[3] || '',
      detail: row[4] || '',
      mood:   row[5] || '',
    }))
    .filter(r => r.who && r.date); // skip blank rows

  return jsonResponse({ status: 'ok', rows });
}

// ── APPEND ────────────────────────────────────────────────────────────────────
function handleAppend(sheetName, row) {
  const sheet = getOrCreateSheet(sheetName);
  sheet.appendRow(row);
  return jsonResponse({ status: 'ok' });
}

// ── DELETE ────────────────────────────────────────────────────────────────────
// rowIndex is the actual 1-based sheet row number returned by handleRead.
function handleDelete(sheetName, rowIndex) {
  if (!rowIndex || rowIndex < 2) {
    return jsonResponse({ status: 'error', error: 'Invalid rowIndex: ' + rowIndex });
  }
  const sheet = getOrCreateSheet(sheetName);
  sheet.deleteRow(rowIndex);
  return jsonResponse({ status: 'ok' });
}

// ── HELPERS ───────────────────────────────────────────────────────────────────
function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['who', 'date', 'name', 'type', 'detail', 'mood', 'timestamp']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatDate(value) {
  // Handles both Date objects and strings
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(value);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
