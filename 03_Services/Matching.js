/**
 * Core Processing Module: Matches invoices to statement rows using AI.
 * Phases A (OCR) → B (Aggregate) → C (AI Parse) → D (Match & Store)
 */

/**
 * Entry Point: Scans the root Drive folder for images and processes them.
 * Matches against ALL PENDING rows in the main sheet.
 */
function processInvoicesFromDrive() {
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const sheetName = "Drive Scan";
  
  logMsg(LOG.PROCESS, 'INFO', 'Starting batch processing from Drive folder');

  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) throw new Error('Main sheet not found');

  clearSidebarLog(); // Auto-clear log at start
  showProcessingSidebar(sheetName, `🚀 Starting batch processing from Drive folder...`);

  // 1. Get images from Drive root folder
  showProcessingSidebar(sheetName, `Phase A: Scanning root Drive folder...`);

  const rootFolder = DriveApp.getFolderById(config.folderId);
  const files = rootFolder.getFiles();
  const images = [];
  let fileCounter = 1;
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType().startsWith('image/')) {
      const base64 = Utilities.base64Encode(file.getBlob().getBytes());
      images.push({ 
        imageIndex: fileCounter++, 
        driveId: file.getId(),
        base64: base64, 
        imageObj: file, 
        fileName: file.getName() 
      });
    }
  }

  if (images.length === 0) {
    showProcessingSidebar(sheetName, `❌ No new images found in root folder.`);
    return "No images found.";
  }
  showProcessingSidebar(sheetName, `Found ${images.length} image(s).`);

  // 2. Get ALL pending statement rows
  const lastRow = mainSheet.getLastRow();
  let statementRows = [];
  if (lastRow >= DATA_START_ROW) {
    const allData = mainSheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 24).getValues();
    statementRows = allData
      .map((row, idx) => ({
        rowNum: idx + DATA_START_ROW,
        date: parseDate(row[COL.DOC_DATE - 1]),
        amount: parseAmount(row[COL.AMOUNT_PURCHASE - 1]),
        currency: (row[COL.CURRENCY_LOCAL - 1] || '').toString().trim().toUpperCase(),
        merchant: (row[COL.MERCHANT - 1] || '').toString().trim(),
        status: (row[COL.MATCH_STATUS - 1] || '').toString().trim(),
      }))
      .filter(r => !r.status || r.status === STATUS.PENDING);
  }

  // 3. Phase A (OCR via Cloud Vision)
  showProcessingSidebar(sheetName, `Running OCR on ${images.length} files...`);
  const ocrResults = phaseA_ocrAllImages(images, config.visionJson);
  ocrResults.forEach(r => {
    showProcessingSidebar(sheetName, `  ${r.success ? '✅' : '❌'} OCR: ${images.find(img => img.imageIndex === r.imageIndex).fileName}`);
  });

  // 4. Phase B (Payload Building)
  const batchPayload = phaseB_buildBatchPayload(ocrResults, statementRows);

  // 5. Phase C (Gemini Model Call)
  showProcessingSidebar(sheetName, `Phase C: Calling Gemini for batch parsing...`);
  const geminiResults = phaseC_callGeminiBatch(batchPayload, config.geminiKey);
  
  if (geminiResults.error) {
    showProcessingSidebar(sheetName, `❌ Gemini Error: ${geminiResults.error}`);
    return "Process failed.";
  }
  showProcessingSidebar(sheetName, `Parsed ${geminiResults.length} invoices successfully.`);

  // 6. Phase D (Apply & Classify Files)
  showProcessingSidebar(sheetName, `Phase D: Matching and moving files...`);
  const summary = phaseD_applyResults(mainSheet, geminiResults, statementRows, ocrResults, rootFolder, sheetName);
  
  // Cleanup OCR failures
  ocrResults.filter(r => !r.success).forEach(r => {
    try {
      moveImageToError(r.fileObj, r.error || 'OCR empty', rootFolder);
    } catch(e) {}
  });

  showProcessingSidebar(sheetName, `\n=== BATCH COMPLETE ===\nMatched: ${summary.matched} | Unmatched: ${summary.unmatched} | Needs Review: ${summary.ambiguous} | Errors: ${summary.errors}`);
  
  return summary;
}

/**
 * Entry point: process all images in the active DD-MM sheet.
 * (Kept for backwards compatibility)
 */
function processImagesInSheet() {
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const sheetName = activeSheet.getName();

  // Optional check: still allowing processing in current sheet if it's a valid date sheet,
  // but most processing now happens from Drive scan.
  logMsg(LOG.PROCESS, 'INFO', `Processing sheet: ${sheetName}`);

  logMsg(LOG.PROCESS, 'INFO', `Starting processing for sheet: ${sheetName}`);
  let sidebarLog = `Process Invoices — ${sheetName}\n\n`;
  showProcessingSidebar(sheetName, sidebarLog);

  const config = getConfig();
  const sheetDate = parseDaySheetName(sheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);

  // Get statement rows for this date
  const lastRow = mainSheet.getLastRow();
  let statementRows = [];
  if (lastRow >= DATA_START_ROW) {
    const allData = mainSheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 24).getValues();
    statementRows = allData
      .map((row, idx) => ({
        rowNum: idx + DATA_START_ROW,
        date: parseDate(row[COL.DOC_DATE - 1]),
        amount: parseAmount(row[COL.AMOUNT_PURCHASE - 1]),
        currency: (row[COL.CURRENCY_LOCAL - 1] || '').toString().trim().toUpperCase(),
        merchant: (row[COL.MERCHANT - 1] || '').toString().trim(),
        status: (row[COL.MATCH_STATUS - 1] || '').toString().trim(),
      }))
      .filter(r => r.date && r.amount && (!r.status || r.status === STATUS.PENDING));
  }

  const images = extractEmbeddedImages(activeSheet);
  if (images.length === 0) return;

  const ocrResults = phaseA_ocrAllImages(images, config.visionJson);
  const batchPayload = phaseB_buildBatchPayload(ocrResults, statementRows);
  const geminiResults = phaseC_callGeminiBatch(batchPayload, config.geminiKey);
  
  let rootFolder = null;
  try { rootFolder = DriveApp.getFolderById(config.folderId); } catch (e) {}

  phaseD_applyResults(mainSheet, geminiResults, statementRows, ocrResults, rootFolder, sheetName);
}

/**
 * Phase A: OCR all images via Cloud Vision.
 */
function phaseA_ocrAllImages(images, visionJson) {
  const results = [];
  if (!images || !Array.isArray(images)) return results;
  
  images.forEach(img => {
    const ocrResult = callCloudVision(img.base64, visionJson);
    results.push({ 
      imageIndex: img.imageIndex, 
      rawText: ocrResult.text || '', 
      success: !ocrResult.error, 
      error: ocrResult.error,
      fileObj: img.imageObj 
    });
  });
  return results;
}

/**
 * Phase B: Aggregate all OCR texts into one batch payload.
 */
function phaseB_buildBatchPayload(ocrResults, statementRows) {
  const successfulOcr = ocrResults.filter(r => r.success);
  return {
    invoices: successfulOcr.map(r => ({ 
      index: r.imageIndex, 
      ocrText: r.rawText,
      fileObj: r.fileObj 
    })),
    statementRows: statementRows.map(r => ({
      rowNum: r.rowNum,
      date: r.date ? r.date.toISOString().substring(0, 10) : '',
      amount: r.amount,
      currency: r.currency,
      merchant: r.merchant,
    })),
  };
}

/**
 * Phase C: Send ONE Gemini batch call.
 */
function phaseC_callGeminiBatch(batchPayload, geminiKey) {
  return callGeminiBatch(batchPayload, geminiKey);
}

/**
 * Phase D: Apply Gemini results to main sheet and move Drive files.
 */
function phaseD_applyResults(mainSheet, geminiResults, statementRows, ocrResults, driveRootFolder, sheetName) {
  // 6. Phase D: Apply results by iterating over all OCR attempts to ensure nothing is missed
  const summary = { matched: 0, unmatched: 0, ambiguous: 0, errors: 0 };
  
  ocrResults.forEach(ocr => {
    try {
      const result = Array.isArray(geminiResults) ? geminiResults.find(r => (r.imageIndex || r.index) == ocr.imageIndex) : null;
      const fileObj = ocr.fileObj;
      const fileName = fileObj ? fileObj.getName() : `Image_${ocr.imageIndex}`;
      
      let finalStatus = STATUS.PENDING;
      let matchReason = "No AI result returned for this image";
      let matchResult = null;

      if (result) {
        matchResult = matchToStatement(result, statementRows);
        matchReason = matchResult ? matchResult.reason : 'No matching statement row found';

        if (matchResult === null) {
          summary.unmatched++;
          finalStatus = "Unmatched";
        } else if (matchResult.ambiguous) {
          applyAmbiguous(mainSheet, matchResult.candidates, sheetName, ocr.imageIndex);
          summary.ambiguous++;
          finalStatus = STATUS.AMBIGUOUS;
        } else {
          finalStatus = STATUS.MATCHED;
          const matchedRow = statementRows.find(r => r.rowNum === matchResult.row);
          if (matchedRow) {
            matchedRow.status = STATUS.MATCHED;
            if (fileObj && matchedRow.date && driveRootFolder) {
              showProcessingSidebar(sheetName, `  ✨ Matched ${fileName} to Row ${matchedRow.rowNum}`);
              renameFileWithPrefix(fileObj, matchedRow.date);
              moveImageToDoneFolder(fileObj.getBlob(), driveRootFolder.getId(), matchedRow.date);
              fileObj.setTrashed(true); 
              createOrUpdateDaySheet(SpreadsheetApp.getActiveSpreadsheet(), matchedRow.date, mainSheet);
            }
          }
          applyMatch(mainSheet, matchResult.row, sheetName, ocr.imageIndex, result);
          summary.matched++;
        }
      } else {
        summary.unmatched++;
        finalStatus = "AI Skip";
      }

      logToAuditSheet(fileName, finalStatus, result || { error: 'No AI data' }, matchReason);
    } catch (e) {
      summary.errors++;
      logMsg(LOG.MATCH, 'ERROR', 'Apply Result Loop Error', e.message);
    }
  });

  return summary;
}

/**
 * Match a parsed invoice to statement rows using scoring.
 */
function matchToStatement(parsedInvoice, statementRows) {
  const invoiceDate = parseDate(parsedInvoice.date);
  const invoiceAmount = parsedInvoice.amount;
  const invoiceCurrency = (parsedInvoice.currency || '').trim();
  const invoiceMerchant = (parsedInvoice.merchant || '').toLowerCase().trim();
  const aiSuggestedRow = parsedInvoice.matchedRowIndex;

  if (!invoiceDate || !invoiceAmount || !invoiceCurrency) return null;

  const pendingRows = statementRows.filter(r => !r.status || r.status === STATUS.PENDING);
  
  // 1. Check AI Suggestion (High Confidence)
  if (aiSuggestedRow) {
    const suggested = pendingRows.find(r => r.rowNum === aiSuggestedRow);
    if (suggested) {
      const dateDiffDays = dateDiff(invoiceDate, suggested.date);
      const amountMatch = Math.abs(suggested.amount - invoiceAmount) <= 1; // tight AI match
      if (dateDiffDays <= MATCH_DATE_TOLERANCE && amountMatch) {
         return { row: aiSuggestedRow, score: 1.0 };
      }
    }
  }

  // 2. Fallback to Score-based matching
  const scored = [];
  for (const row of pendingRows) {
    if (!row.date || !row.amount) continue;

    // Currency Check (Skip if statement currency is empty)
    const rowCurrency = (row.currency || '').toUpperCase().trim();
    let currencyMatch = true;
    if (rowCurrency) {
      currencyMatch = (rowCurrency === invoiceCurrency.toUpperCase());
      if (!currencyMatch && invoiceCurrency === '$' && ['AUD','USD','VND','SGD'].includes(rowCurrency)) {
        currencyMatch = true; 
      }
    }
    if (!currencyMatch) continue;

    let score = 0;
    // A. Amount Score (Weight: 0.6)
    const amtDiff = Math.abs(row.amount - invoiceAmount);
    if (amtDiff < 0.01) score += 0.6; // Perfect match
    else if (amtDiff <= 2.0) score += 0.4; // Close match
    else continue;

    // B. Date Score (Weight: 0.3)
    const diff = dateDiff(invoiceDate, row.date);
    if (diff > MATCH_DATE_TOLERANCE) continue;
    if (diff === 0) score += 0.3;
    else if (diff === 1) score += 0.2;
    else if (diff === 2) score += 0.1;

    // C. Merchant Score (Weight: 0.2)
    const statementMerchantClean = (row.merchant || '').toLowerCase().replace(/[^a-z]/g, '');
    const invoiceMerchantClean = invoiceMerchant.replace(/[^a-z]/g, '');
    
    let merchantReason = '';
    if (statementMerchantClean && invoiceMerchantClean && (statementMerchantClean.includes(invoiceMerchantClean) || invoiceMerchantClean.includes(statementMerchantClean))) {
      score += 0.2;
      merchantReason = 'Merchant matched';
    }

    const reason = `Amt: ${amtDiff.toFixed(2)}, Day: ${diff}, Merchant: ${merchantReason || 'No match'}`;
    if (score >= MIN_MATCH_SCORE) scored.push({ row: row.rowNum, score, reason });
  }

  if (scored.length === 0) return null;
  scored.sort((a, b) => b.score - a.score);
  
  if (scored.length === 1 || scored[0].score > scored[1].score + 0.1) return scored[0];
  return { ambiguous: true, candidates: scored.map(s => s.row), reason: 'Multiple high-score candidates' };
}

function applyMatch(sheet, rowNum, daySheetName, imageIndex, invoiceData) {
  const config = getConfig();
  
  // Find the account name from config to avoid relying on VLOOKUP formulas which might be overwritten
  let accountName = '';
  if (invoiceData.accountCode) {
    const foundAcc = config.accountCodes.find(ac => ac.code == invoiceData.accountCode);
    if (foundAcc) accountName = foundAcc.name;
  }

  const update = [[
    invoiceData.accountCode || '',        // Q (17) Account Code
    accountName,                          // R (18) Account Name (Static value)
    config.defaults.io || '',             // S (19) I/O Code
    config.defaults.costCenter || '',     // T (20) Cost Center
    invoiceData.summary || '',            // U (21) Summary
    config.defaults.taxId || '',          // V (22) Tax ID
    STATUS.MATCHED,                       // W (23) Match Status
    `${daySheetName} #${imageIndex}`      // X (24) Invoice Ref
  ]];
  sheet.getRange(rowNum, COL.ACCOUNT_CODE, 1, 8).setValues(update);
}

function applyAmbiguous(sheet, candidateRows, daySheetName, imageIndex) {
  candidateRows.forEach(rowNum => {
    sheet.getRange(rowNum, COL.MATCH_STATUS, 1, 1).setValue(STATUS.AMBIGUOUS);
    sheet.getRange(rowNum, COL.INVOICE_REF, 1, 1).setValue(`${daySheetName} #${imageIndex} (ambiguous)`);
  });
}

function extractEmbeddedImages(sheet) {
  const overCellImages = sheet.getImages();
  const results = [];
  overCellImages.forEach((img, idx) => {
    try {
      const imgBlob = img.getBlob ? img.getBlob() : null;
      if (imgBlob) {
        const base64 = Utilities.base64Encode(imgBlob.getBytes());
        results.push({ imageIndex: idx + 1, base64, imageObj: img });
      }
    } catch (e) {}
  });
  return results;
}

function callCloudVision(base64, visionJson) {
  try {
    const token = getServiceAccountToken(visionJson, VISION_SCOPE);
    const payload = {
      requests: [{ image: { content: base64 }, features: [{ type: 'TEXT_DETECTION' }] }]
    };
    const resp = UrlFetchApp.fetch(VISION_ENDPOINT, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return { error: `HTTP ${resp.getResponseCode()}` };
    const data = JSON.parse(resp.getContentText());
    return { text: data.responses?.[0]?.fullTextAnnotation?.text || '' };
  } catch (e) { return { error: e.message }; }
}

function callGeminiBatch(batchPayload, geminiKey) {
  const config = getConfig();
  const model = config.geminiModel || 'gemini-2.0-flash';
  const endpoint = `${GEMINI_BASE_URL}/${model}:generateContent?key=${geminiKey}`;
  
  const prompt = buildBatchInvoicePrompt(batchPayload);
  const maxRetries = 2;
  let attempt = 0;

  while (attempt <= maxRetries) {
    try {
      const resp = UrlFetchApp.fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
        muteHttpExceptions: true
      });
      
      const code = resp.getResponseCode();
      if (code === 429) {
        logMsg(LOG.GEMINI, 'WARN', `Rate limited (429). Waiting... (Attempt ${attempt+1})`);
        Utilities.sleep(5000 * (attempt + 1)); // Backoff
        attempt++;
        continue;
      }
      
      if (code !== 200) return { error: `HTTP ${code}: ${resp.getContentText().substring(0, 50)}` };
      
      const data = JSON.parse(resp.getContentText());
      return parseGeminiBatchResponse(data.candidates?.[0]?.content?.parts?.[0]?.text || '');
    } catch (e) { 
      attempt++;
      Utilities.sleep(2000);
    }
  }
  return { error: 'Exceeded retries for Gemini API' };
}

function buildBatchInvoicePrompt(batchPayload) {
  const invoiceTexts = batchPayload.invoices.map(inv => `=== INVOICE ${inv.index} ===\n${inv.ocrText}`).join('\n\n');
  const statementHints = batchPayload.statementRows.map(r => `Row ${r.rowNum}: ${r.date} | ${r.currency} ${r.amount} | Merchant: ${r.merchant}`).join('\n');
  
  const config = getConfig();
  const accountCodesHint = config.accountCodes.map(ac => `${ac.code}: ${ac.name}`).join(', ');

  return `You are a corporate expense reconciliation assistant. 
TASK: 
1. Parse each invoice to find the FINAL BILLING AMOUNT and Date.
2. Link to a statement row if possible.
3. Suggest the most likely Account Code from the list below based on the merchant and items.
4. Return ONLY a valid JSON array.

AVAILABLE ACCOUNT CODES:
${accountCodesHint}

PENDING STATEMENT ROWS:
${statementHints}

INVOICES TO PARSE:
${invoiceTexts}

Return format:
[{"imageIndex":1,"date":"YYYY-MM-DD","amount":123.45,"currency":"AUD","merchant":"Name","summary":"Short desc","matchedRowIndex": 5, "accountCode": "Code"}]`;
}

function parseGeminiBatchResponse(responseText) {
  const clean = responseText.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
  try {
    const parsed = JSON.parse(clean);
    return Array.isArray(parsed) ? parsed : { error: 'Expected array' };
  } catch (e) { return { error: e.message }; }
}
