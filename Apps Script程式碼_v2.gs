// ================================================
// 銘勁汽車 CS特案報告 + 協議書紀錄
// Google Apps Script - 更新版
// ================================================

const SHEET_CS   = 'CS紀錄';
const SHEET_AGR  = '協議書紀錄';

// CS特案報告欄位
const CS_COLUMNS = [
  'id', 'savedAt', 'unit', 'staff', 'customer',
  'vehicle', 'phone', 'delivery', 'subject',
  'content', 'amount', 'attachment', 'items', 'cosign', 'cosignUnit'
];

// 協議書紀錄欄位
const AGR_COLUMNS = [
  'id', 'date', 'branch', 'plate', 'name',
  'model', 'color', 'reason', 'createdAt'
];

// ── 處理 GET 請求 ──
function doGet(e) {
  try {
    const action = e.parameter.action;
    const type   = e.parameter.type || 'cs';

    if (action === 'getAll') {
      const data = type === 'agreement'
        ? getAllRecords(SHEET_AGR, AGR_COLUMNS)
        : getAllRecords(SHEET_CS, CS_COLUMNS);
      return respond({ ok: true, data });
    }
    return respond({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

// ── 處理 POST 請求 ──
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;
    const type    = payload.type || 'cs';
    const sheetName = type === 'agreement' ? SHEET_AGR : SHEET_CS;
    const columns   = type === 'agreement' ? AGR_COLUMNS : CS_COLUMNS;

    if (action === 'save') {
      saveRecord(sheetName, columns, payload.record);
      return respond({ ok: true, message: 'Saved' });
    }
    if (action === 'delete') {
      deleteRecord(sheetName, payload.id);
      return respond({ ok: true, message: 'Deleted' });
    }
    return respond({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

// ── 讀取所有紀錄 ──
function getAllRecords(sheetName, columns) {
  const sheet = getSheet(sheetName, columns);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  return rows.slice(1).map(row => {
    const record = {};
    columns.forEach((col, i) => {
      record[col] = row[i] !== undefined ? String(row[i]) : '';
    });
    if (record.items) {
      try { record.items = JSON.parse(record.items); }
      catch(e) { record.items = []; }
    }
    if ('cosign' in record) record.cosign = record.cosign === 'true';
    return record;
  }).filter(r => r.id);
}

// ── 寫入一筆紀錄 ──
function saveRecord(sheetName, columns, record) {
  const sheet = getSheet(sheetName, columns);
  const row = columns.map(col => {
    if (col === 'items') return JSON.stringify(record[col] || []);
    return record[col] !== undefined ? record[col] : '';
  });
  sheet.appendRow(row);
}

// ── 刪除一筆紀錄 ──
function deleteRecord(sheetName, id) {
  const sheet = getSheet(sheetName);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

// ── 取得或建立工作表 ──
function getSheet(sheetName, columns) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (columns) setupHeaders(sheet, sheetName, columns);
  }
  return sheet;
}

// ── 設定標題列 ──
function setupHeaders(sheet, sheetName, columns) {
  const csHeaders = [
    '編號','儲存日期','申請單位','負責業務','顧客姓名',
    '車號/引擎碼','聯絡電話','交車日期','主旨',
    '處理需求','相關金額','請附件','報告項目','需會簽','會簽單位'
  ];
  const agrHeaders = [
    '編號','日期','責任單位','車牌號碼','消費者姓名',
    '車型','車色','協議原因','建立時間'
  ];
  const headers = sheetName === SHEET_AGR ? agrHeaders : csHeaders;
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (!firstRow[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1a1a1c')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

// ── 回傳 JSON ──
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
