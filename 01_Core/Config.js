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
