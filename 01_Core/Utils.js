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
    s = s.replace(/[₫$€£¥₩฿]/g, '');
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
