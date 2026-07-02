const INVENTORY_SUMMARY_SHEET_NAME = '工作表1';

function doPost(e) {
  const payload = JSON.parse(e.postData.contents || '{}');
  if (payload.action !== 'UPSERT_INVENTORY_SUMMARY') {
    return jsonResponse({ ok: false, error: 'Unsupported action' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_SUMMARY_SHEET_NAME);
  if (!sheet) {
    return jsonResponse({ ok: false, error: `Sheet not found: ${INVENTORY_SUMMARY_SHEET_NAME}` });
  }

  const month = String(payload.month || '').trim();
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  if (!month || rows.length === 0) {
    return jsonResponse({ ok: false, error: 'Missing month or rows' });
  }

  const values = sheet.getDataRange().getValues();
  const monthRowIndex = values.findIndex(row => String(row[0] || '').trim() === month);
  if (monthRowIndex === -1) {
    return jsonResponse({ ok: false, error: `Month block not found: ${month}` });
  }

  const firstDataRowIndex = monthRowIndex + 2;
  const rowByDrugName = new Map();
  for (let rowIndex = firstDataRowIndex; rowIndex < values.length; rowIndex++) {
    const drugName = String(values[rowIndex][1] || '').trim();
    if (!drugName) break;
    rowByDrugName.set(drugName, rowIndex + 1);
  }

  const written = [];
  rows.forEach(row => {
    const drugName = String(row['對比劑藥名'] || '').trim();
    const targetRow = rowByDrugName.get(drugName);
    if (!targetRow) return;

    sheet.getRange(targetRow, 3, 1, 5).setValues([[
      Number(row['本月使用人數 (人)'] || 0),
      Number(row['本月消耗瓶數 (瓶)'] || 0),
      Number(row['攝影室期末結存 (瓶)'] || 0),
      Number(row['申請量 (可直接修改)'] || 0),
      Number(row['鐵櫃結存 (可直接修改)'] || 0)
    ]]);
    written.push(drugName);
  });

  return jsonResponse({ ok: true, month, written });
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
