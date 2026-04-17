/**
 * Statement Import Module (Disabled in UI)
 * Keep for internal/future use.
 */

function handleStatementUpload(fileBase64, fileName) {
  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let rows;
  try {
    const json = Utilities.newBlob(Utilities.base64Decode(fileBase64)).getDataAsString();
    rows = JSON.parse(json).map(row => row.map(cell => (cell && cell.__type === 'date') ? new Date(cell.value) : cell));
  } catch (e) { throw new Error('Data decode failed'); }

  const colMap = detectStatementColumns(rows[0], rows.slice(1, 4), config.geminiKey);
  const parsedRows = rows.slice(1).map(r => parseStatementRow(r, colMap)).filter(p => p !== null);
  
  if (parsedRows.length === 0) return 'No valid rows found';

  let mainSheet = ss.getSheetByName(MAIN_SHEET_NAME) || initMainSheet();
  writeStatementRows(mainSheet, parsedRows, config);

  const uniqueDates = [...new Set(parsedRows.map(r => formatDaySheetName(r.date)))];
  uniqueDates.forEach(dStr => {
    const date = parseDaySheetName(dStr) || parsedRows.find(r => formatDaySheetName(r.date) === dStr).date;
    createOrUpdateDaySheet(ss, date, mainSheet);
  });

  return `Imported ${parsedRows.length} rows.`;
}

function detectStatementColumns(headerRow, sampleRows, geminiKey) {
  return DEFAULT_COL_MAP; // Hardcoded default to save tokens for now
}

function parseStatementRow(rawRow, colMap) {
  const date = parseDate(rawRow[colMap.date]);
  const amount = parseAmount(rawRow[colMap.amount]);
  if (!date || !amount) return null;
  return { date, amount, currency: (rawRow[colMap.currency] || '').toString().trim().toUpperCase(), merchant: (rawRow[colMap.merchant] || '').toString().trim(), rawRow };
}

function writeStatementRows(mainSheet, parsedRows, config) {
  const startRow = mainSheet.getLastRow() + 1;
  const data2d = parsedRows.map(pr => {
    const row = new Array(24).fill('');
    pr.rawRow.slice(0, 16).forEach((v, i) => row[i] = v);
    row[COL.IO_CODE - 1] = config.defaults.io;
    row[COL.COST_CENTER - 1] = config.defaults.costCenter;
    row[COL.TAX_ID - 1] = config.defaults.taxId;
    row[COL.MATCH_STATUS - 1] = STATUS.PENDING;
    return row;
  });
  mainSheet.getRange(startRow, 1, data2d.length, 24).setValues(data2d);
  writeVlookupFormulas(mainSheet, startRow, startRow + data2d.length - 1, config.accountCodeRange);
  applyMatchStatusFormatting(mainSheet, startRow, startRow + data2d.length - 1);
}
