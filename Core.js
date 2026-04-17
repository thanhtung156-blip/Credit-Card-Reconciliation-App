// â”€â”€â”€ SHEET NAMES â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const CONFIG_SHEET_NAME   = 'Config';
const MAIN_SHEET_NAME     = 'Corporate card details';
const AUDIT_SHEET_NAME    = 'OCR Audit';

// â”€â”€â”€ PROCESSING PARAMS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const DATA_START_ROW       = 2;
const MATCH_DATE_TOLERANCE = 4;
const MIN_MATCH_SCORE      = 0.8;
const DONE_SUBFOLDER       = 'Unmatched';
const ERROR_SUBFOLDER      = 'Error';

// â”€â”€â”€ AI CONFIGURATION â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const VISION_SCOPE        = 'https://www.googleapis.com/auth/cloud-vision';
const VISION_ENDPOINT     = 'https://vision.googleapis.com/v1/images:annotate';
const GEMINI_BASE_URL     = 'https://generativelanguage.googleapis.com/v1beta/models';
const TOKEN_CACHE_KEY     = 'vision_access_token';
const TOKEN_EXPIRY_BUFFER = 60 * 1000;

// â”€â”€â”€ BANK FILE â€” DEFAULT COLUMN INDICES (0-based) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const DEFAULT_COL_MAP = {
  date:     2,
  amount:   9,
  currency: 7,
  merchant: 5,
};

// â”€â”€â”€ OUTPUT SHEET â€” COLUMN NUMBERS (1-based) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const COL = {
  STATUS_NAME:     1,
  CARD_NO:         2,
  DOC_DATE:        3,
  APPROVAL_TYPE:   4,
  APPROVAL_NO:     5,
  MERCHANT:        6,
  CURRENCY_A:      7,
  CURRENCY_LOCAL:  8,
  AMOUNT_APPROVED: 9,
  AMOUNT_PURCHASE: 10,
  CURRENCY_KRW:    11,
  SUPPLY_PRICE:    12,
  TAX_AMOUNT:      13,
  BILLING_AMOUNT:  14,
  DOC_NO:          15,
  RECIPIENT:       16,
  ACCOUNT_CODE:    17,
  ACCOUNT_NAME:    18,
  IO_CODE:         19,
  COST_CENTER:     20,
  SUMMARY:         21,
  TAX_ID:          22,
  MATCH_STATUS:    23,
  INVOICE_REF:     24,
};

// â”€â”€â”€ MATCH STATUS VALUES â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const STATUS = {
  PENDING:    'â³ Pending',
  MATCHED:    'âœ… Matched',
  NO_INVOICE: 'âŒ No Invoice',
  AMBIGUOUS:  'âš ï¸ Needs Review',
};

// â”€â”€â”€ LOG PREFIXES â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const LOG = {
  IMPORT:  '[importStatement]',
  VISION:  '[callCloudVision]',
  GEMINI:  '[callGeminiBatch]',
  MATCH:   '[matchToStatement]',
  SHEET:   '[sheetOps]',
  DRIVE:   '[driveOps]',
  CONFIG:  '[getConfig]',
  PROCESS: '[processImagesInSheet]',
};

// â”€â”€â”€ SCRIPT-SCOPED CONFIG CACHE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
let _configCache = null;
/**
 * Read and validate Config sheet. Caches result for duration of execution.
 * @returns {Object} config object
 */
function getConfig() {
  if (_configCache) return _configCache;

  logMsg(LOG.CONFIG, 'INFO', 'Reading Config sheet');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) throw new Error('Config sheet not found. Run "Initialize Config Template" first.');

  const allValues = sheet.getDataRange().getValues();

  let visionJson = null;
  let geminiKey = null;
  let geminiModel = 'gemini-2.0-flash'; // Fallback
  let folderId = null;
  const accountCodes = [];
  const ioCodes = [];
  const costCenters = [];
  const taxIds = [];
  const defaults = { io: '', costCenter: '', taxId: '' };
  let accountCodeRange = '';

  let currentTable = null;
  let accountCodeStartRow = -1;
  let accountCodeEndRow = -1;

  for (let i = 0; i < allValues.length; i++) {
    const row = allValues[i];
    const firstCell = (row[0] || '').toString().trim();

    // Detect table headers
    if (firstCell === 'TABLE 1: API CREDENTIALS & SETTINGS') { currentTable = 'credentials'; continue; }
    if (firstCell === 'TABLE 2: ACCOUNT CODES') { currentTable = 'accounts'; accountCodeStartRow = i + 2; continue; }
    if (firstCell === 'TABLE 3: I/O CODES') { currentTable = 'io'; continue; }
    if (firstCell === 'TABLE 4: COST CENTERS') { currentTable = 'costCenters'; continue; }
    if (firstCell === 'TABLE 5: TAX IDS') { currentTable = 'taxIds'; continue; }
    if (firstCell === 'TABLE 6: DEFAULTS') { currentTable = 'defaults'; continue; }
    if (firstCell.startsWith('TABLE') && currentTable === 'accounts') {
      accountCodeEndRow = i - 1;
    }

    if (!firstCell && currentTable !== 'accounts') continue;

    switch (currentTable) {
      case 'credentials':
        if (firstCell === 'Service Account JSON') visionJson = (row[1] || '').toString().trim();
        if (firstCell === 'Gemini API Key') geminiKey = (row[1] || '').toString().trim();
        if (firstCell === 'Gemini Model') geminiModel = (row[1] || '').toString().trim() || 'gemini-2.0-flash';
        if (firstCell === 'Drive Root Folder ID') folderId = (row[1] || '').toString().trim();
        break;
      case 'accounts':
        if (firstCell === 'Code' || firstCell === 'Account Code') break;
        if (firstCell) {
          accountCodes.push({ code: firstCell, name: (row[1] || '').toString().trim() });
          accountCodeEndRow = i;
        }
        break;
      case 'io':
        if (firstCell) ioCodes.push(firstCell);
        break;
      case 'costCenters':
        if (firstCell) costCenters.push(firstCell);
        break;
      case 'taxIds':
        if (firstCell) taxIds.push(firstCell);
        break;
      case 'defaults':
        if (firstCell === 'Default I/O') defaults.io = (row[1] || '').toString().trim();
        if (firstCell === 'Default Cost Center') defaults.costCenter = (row[1] || '').toString().trim();
        if (firstCell === 'Default Tax ID') defaults.taxId = (row[1] || '').toString().trim();
        break;
    }
  }

  if (!visionJson) throw new Error('[getConfig] Service Account JSON is missing in Config TABLE 1.');
  if (!geminiKey)  throw new Error('[getConfig] Gemini API Key is missing in Config TABLE 1.');
  if (!folderId)   throw new Error('[getConfig] Drive Root Folder ID is missing in Config TABLE 1.');

  let parsedJson;
  try {
    parsedJson = JSON.parse(visionJson);
  } catch (e) {
    throw new Error('[getConfig] Service Account JSON is not valid JSON: ' + e.message);
  }
  if (!parsedJson.private_key) throw new Error('[getConfig] Service Account JSON missing "private_key".');
  if (!parsedJson.client_email) throw new Error('[getConfig] Service Account JSON missing "client_email".');

  if (accountCodeStartRow > 0 && accountCodeEndRow >= accountCodeStartRow) {
    const startRowSheet = accountCodeStartRow + 1;
    const endRowSheet = accountCodeEndRow + 1;
    accountCodeRange = `$A$${startRowSheet}:$B$${endRowSheet}`;
  }

  _configCache = {
    visionJson,
    geminiKey,
    geminiModel,
    folderId,
    accountCodes,
    ioCodes,
    costCenters,
    taxIds,
    defaults,
    accountCodeRange,
  };

  logMsg(LOG.CONFIG, 'INFO', 'Config loaded',
    `accounts=${accountCodes.length} ioCodes=${ioCodes.length} costCenters=${costCenters.length}`);

  return _configCache;
}
/**
 * Log a message with prefix, level, message, and optional detail.
 */
function logMsg(prefix, level, message, detail) {
  const parts = [prefix, level + ':', message];
  if (detail !== undefined && detail !== null) parts.push('| ' + detail);
  const formatted = parts.join(' ');
  Logger.log(formatted);
  return formatted;
}

/**
 * Parse a date value from various formats.
 */
function parseDate(val) {
  if (val === null || val === undefined || val === '') return null;

  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }

  const num = typeof val === 'string' ? Number(val.trim()) : val;
  if (typeof val === 'number' || (typeof val === 'string' && !isNaN(num) && val.trim() !== '')) {
    const serial = typeof val === 'number' ? val : num;
    if (serial <= 0) return null;
    const msPerDay = 86400000;
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(excelEpoch.getTime() + serial * msPerDay);
    return isNaN(d.getTime()) ? null : d;
  }

  const s = val.toString().trim();
  if (!s) return null;

  const dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmy) {
    const d = new Date(parseInt(dmy[3]), parseInt(dmy[2]) - 1, parseInt(dmy[1]));
    return isNaN(d.getTime()) ? null : d;
  }

  const ymd = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (ymd) {
    const d = new Date(parseInt(ymd[1]), parseInt(ymd[2]) - 1, parseInt(ymd[3]));
    return isNaN(d.getTime()) ? null : d;
  }

  const mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdy) {
    const d = new Date(parseInt(mdy[3]), parseInt(mdy[1]) - 1, parseInt(mdy[2]));
    return isNaN(d.getTime()) ? null : d;
  }

  logMsg(LOG.IMPORT, 'WARN', 'parseDate: unparseable value', String(val));
  return null;
}

/**
 * Parse an amount value to integer.
 */
function parseAmount(val) {
  if (val === null || val === undefined) return null;

  let n;
  if (typeof val === 'number') {
    n = val;
  } else {
    let s = val.toString().trim();
    s = s.replace(/[â‚«$â‚¬Â£Â¥â‚©à¸¿]/g, '');
    s = s.replace(/\s*[A-Z]{2,4}\s*$/i, '');
    s = s.trim();

    const dotCount = (s.match(/\./g) || []).length;
    const commaCount = (s.match(/,/g) || []).length;

    if (dotCount > 1) {
      s = s.replace(/\./g, '');
    } else if (commaCount > 1) {
      s = s.replace(/,/g, '');
    } else if (dotCount === 1 && commaCount === 1) {
      const dotIdx = s.lastIndexOf('.');
      const commaIdx = s.lastIndexOf(',');
      if (dotIdx > commaIdx) {
        s = s.replace(/,/g, '');
      } else {
        s = s.replace(/\./g, '').replace(',', '.');
      }
    } else if (commaCount === 1) {
      const parts = s.split(',');
      if (parts[1].length === 3) {
        s = s.replace(',', '');
      } else {
        s = s.replace(',', '.');
      }
    }
    n = parseFloat(s);
  }

  if (isNaN(n) || n < 0) return null;
  // Keep decimals for non-KRW currencies, round to 2 places for safety
  return Math.round(n * 100) / 100;
}

/**
 * Return absolute difference in whole days between two dates.
 */
function dateDiff(a, b) {
  const msPerDay = 86400000;
  return Math.abs(Math.round((a.getTime() - b.getTime()) / msPerDay));
}

/**
 * Format a Date to "DD-MM" zero-padded string.
 */
function formatDaySheetName(date) {
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  return `${dd}-${mm}`;
}

/**
 * Parse a "DD-MM" sheet name to a Date.
 */
function parseDaySheetName(name) {
  if (!name || !/^\d{2}-\d{2}$/.test(name)) return null;
  const dd = parseInt(name.substring(0, 2), 10);
  const mm = parseInt(name.substring(3, 5), 10);
  const year = new Date().getFullYear();
  const d = new Date(year, mm - 1, dd);
  if (isNaN(d.getTime())) return null;
  if (d.getDate() !== dd || d.getMonth() + 1 !== mm) return null;
  return d;
}

/**
 * Returns true if the sheet name matches the DD-MM pattern.
 */
function isDaySheet(sheetName) {
  return /^\d{2}-\d{2}$/.test(sheetName);
}

/**
 * Validate all rows in main sheet before export.
 */
function validateBeforeExport(mainSheet) {
  const lastRow = mainSheet.getLastRow();
  const issues = [];
  if (lastRow < DATA_START_ROW) return issues;

  const data = mainSheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 24).getValues();
  data.forEach((row, i) => {
    const rowNum = i + DATA_START_ROW;
    if (!row[COL.ACCOUNT_CODE - 1] || row[COL.ACCOUNT_CODE - 1].toString().trim() === '') {
      issues.push({ rowNum, issue: 'Missing Account Code' });
    }
    if (row[COL.MATCH_STATUS - 1] === STATUS.PENDING) {
      issues.push({ rowNum, issue: 'Not yet processed' });
    }
  });

  if (issues.length > 0) {
    const lines = issues.map(i => `Row ${i.rowNum}: ${i.issue}`).join('\n');
    SpreadsheetApp.getUi().alert(`Validation Issues Found:\n\n${lines}`);
  } else {
    SpreadsheetApp.getUi().alert('Validation passed. All rows are ready for export.');
  }
  return issues;
}
