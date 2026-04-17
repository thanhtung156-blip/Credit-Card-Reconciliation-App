/**
 * Create or update the main "Corporate card details" sheet.
 * Updated: only initializes columns Q-X (17-24). Columns A-P are preserved.
 */
function initMainSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MAIN_SHEET_NAME);
  }

  // Clear only the system/user columns Q-X (columns 17-24)
  const maxRows = Math.max(sheet.getMaxRows(), 100);
  sheet.getRange(1, 17, maxRows, 8).clearContent().clearFormat().setDataValidation(null);

  const headers = [
    'Account Code', 'Account Name', 'I/O Code', 'Cost Center',
    'Summary', 'Tax ID', 'Match Status', 'Invoice Ref'
  ];

  // Set headers for columns Q-X
  const headerRange = sheet.getRange(1, 17, 1, 8);
  headerRange.setValues([headers]);
  headerRange
    .setFontWeight('bold')
    .setBackground('#34495e')
    .setFontColor('#ffffff');

  // Column widths for Q-X
  const widths = [110, 200, 80, 120, 250, 120, 120, 120];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 17, w));

  sheet.setFrozenRows(1);
  
  // Re-apply conditional formatting and validation if there's data
  const lastRow = sheet.getLastRow();
  if (lastRow >= DATA_START_ROW) {
    const config = getConfig();
    if (config.accountCodeRange) {
      writeVlookupFormulas(sheet, DATA_START_ROW, lastRow, config.accountCodeRange);
    }
    if (config.accountCodes.length > 0) {
      const codes = config.accountCodes.map(ac => ac.code);
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(codes, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(DATA_START_ROW, COL.ACCOUNT_CODE, lastRow - DATA_START_ROW + 1, 1).setDataValidation(rule);
    }
    applyMatchStatusFormatting(sheet, DATA_START_ROW, lastRow);
  }

  logMsg(LOG.SHEET, 'INFO', 'Main sheet initialized (Q-X only)');
  SpreadsheetApp.getUi().alert('Main sheet columns Q-X have been initialized. Columns A-P were preserved.');
}

/**
 * Create the Config sheet with all table headers pre-filled.
 */
function initConfigTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  } else {
    sheet.clear();
  }

  const template = [
    ['TABLE 1: API CREDENTIALS & SETTINGS', ''],
    ['Key', 'Value'],
    ['Service Account JSON', ''],
    ['Gemini API Key', ''],
    ['Gemini Model', 'gemini-2.0-flash'],
    ['Drive Root Folder ID', ''],
    ['', ''],
    ['TABLE 2: ACCOUNT CODES', ''],
    ['Account Code', 'Account Name'],
    ['6200', 'Travel & Entertainment'],
    ['6201', 'Meals & Entertainment'],
    ['6202', 'Office Supplies'],
    ['', ''],
    ['TABLE 3: I/O CODES', ''],
    ['Code', ''],
    ['IO-001', ''],
    ['', ''],
    ['TABLE 4: COST CENTERS', ''],
    ['Cost Center', ''],
    ['CC-001', ''],
    ['', ''],
    ['TABLE 5: TAX IDS', ''],
    ['Tax ID', ''],
    ['TAX-001', ''],
    ['', ''],
    ['TABLE 6: DEFAULTS', ''],
    ['Key', 'Value'],
    ['Default I/O', 'IO-001'],
    ['Default Cost Center', 'CC-001'],
    ['Default Tax ID', 'TAX-001'],
  ];

  sheet.getRange(1, 1, template.length, 2).setValues(template);

  const tableTitleRows = [1, 7, 13, 17, 21, 25];
  tableTitleRows.forEach(r => {
    sheet.getRange(r, 1, 1, 2)
      .setFontWeight('bold')
      .setBackground('#2c3e50')
      .setFontColor('#ffffff');
  });

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 400);

  logMsg(LOG.SHEET, 'INFO', 'Config template initialized');
  SpreadsheetApp.getUi().alert('Config sheet template has been created. Please fill in your API credentials and account codes.');
}

/**
 * Create or update a DD-MM day sheet with transaction summary.
 */
function createOrUpdateDaySheet(wb, date, mainSheet) {
  const sheetName = formatDaySheetName(date);
  logMsg(LOG.SHEET, 'INFO', `Creating/updating day sheet: ${sheetName}`);

  let sheet = wb.getSheetByName(sheetName);
  if (!sheet) {
    sheet = wb.insertSheet(sheetName);
  }

  const lastRow = mainSheet.getLastRow();
  if (lastRow < DATA_START_ROW) return sheet;

  const allData = mainSheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 24).getValues();

  const dateStr = formatDaySheetName(date);
  const matchingRows = allData.filter(row => {
    const d = parseDate(row[COL.DOC_DATE - 1]);
    return d && formatDaySheetName(d) === dateStr;
  });

  // Create blank sheet as requested
  if (sheet.getLastRow() > 0) {
    sheet.clear();
  }
  return sheet;
}

/**
 * Write VLOOKUP formulas in Col R for given row range.
 */
function writeVlookupFormulas(sheet, startRow, endRow, accountCodeRange) {
  const rowCount = endRow - startRow + 1;
  if (rowCount <= 0) return;
  const formulas = [];
  for (let r = startRow; r <= endRow; r++) {
    // Return empty string if match is not found to avoid #N/A errors
    formulas.push([`=IFERROR(VLOOKUP(Q${r},Config!${accountCodeRange},2,0),"")`]);
  }
  sheet.getRange(startRow, COL.ACCOUNT_NAME, rowCount, 1).setFormulas(formulas);
}

/**
 * Apply conditional formatting on Col W for all status values.
 */
function applyMatchStatusFormatting(sheet, startRow, endRow) {
  const rowCount = endRow - startRow + 1;
  if (rowCount <= 0) return;
  const colW = sheet.getRange(startRow, COL.MATCH_STATUS, rowCount, 1);
  const rules = sheet.getConditionalFormatRules();

  const newRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(STATUS.MATCHED)
      .setBackground('#d4edda')
      .setFontColor('#155724')
      .setRanges([colW])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(STATUS.PENDING)
      .setBackground('#fff3cd')
      .setFontColor('#856404')
      .setRanges([colW])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(STATUS.AMBIGUOUS)
      .setBackground('#fff3cd')
      .setFontColor('#856404')
      .setRanges([colW])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(STATUS.NO_INVOICE)
      .setBackground('#f8d7da')
      .setFontColor('#721c24')
      .setRanges([colW])
      .build(),
  ];

  sheet.setConditionalFormatRules([...rules, ...newRules]);
}

/**
 * Log OCR and matching results to a dedicated Audit sheet.
 * Creates the sheet if it doesn't exist.
 */
function logToAuditSheet(fileName, status, parsedData, matchReason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(AUDIT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AUDIT_SHEET_NAME);
    const headers = [['Timestamp', 'File Name', 'Status', 'Match Reason', 'Gemini Parsed']];
    sheet.getRange(1, 1, 1, 5).setValues(headers).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 250);
    sheet.setColumnWidth(5, 400);
  }

  const row = [
    new Date(),
    fileName,
    status,
    matchReason || "",
    JSON.stringify(parsedData)
  ];
  sheet.appendRow(row);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1001) sheet.deleteRow(2);
}
