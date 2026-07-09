function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💳 Corporate Card')
    .addItem('📷 Upload/Paste Invoices (to Drive)', 'showUploadDialog')
    .addItem('🔍 Scan Drive Folder & Reconcile', 'processInvoicesFromDrive')
    .addItem('📏 Arrange & Resize Images', 'arrangeActiveSheetImagesMenu')
    .addSeparator()
    .addItem('✅ Validate before export', 'validateBeforeExportMenu')
    .addItem('🔧 Check Configuration', 'testConfig')
    .addItem('🔑 Test OAuth2 Token', 'testServiceAccountToken')
    // ... existing items ...
    .addSeparator()
    .addItem('⚙️ Initialize Main Sheet', 'initMainSheet')
    .addItem('⚙️ Initialize Config Template', 'initConfigTemplate')
    .addItem('❓ Help & Instructions', 'showHelp')
    .addToUi();
}

function showUploadDialog() {
  const html = HtmlService.createHtmlOutputFromFile('UploadDialog')
    .setTitle('Upload/Paste Invoices')
    .setWidth(550)
    .setHeight(400); // Reduced height as scan button is removed
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload Invoices');
}

/**
 * Handle multiple image uploads from the dialog.
 * Saves files to the root Drive folder before processing.
 */
function handleMultiUpload(files) {
  const config = getConfig();
  const rootFolder = DriveApp.getFolderById(config.folderId);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  let count = 0;

  files.forEach((file, idx) => {
    try {
      // Clean filename and add timestamp/index suffix to ensure uniqueness
      const cleanName = file.name.replace(/[^a-zA-Z0-9.-]/g, '_');
      const uniqueName = `rec_${timestamp}_${idx}_${cleanName}`;
      
      const blob = Utilities.newBlob(Utilities.base64Decode(file.base64), file.type, uniqueName);
      rootFolder.createFile(blob);
      count++;
    } catch (e) {
      logMsg(LOG.DRIVE, 'ERROR', `Failed to upload ${file.name}`, e.message);
    }
  });

  return `Successfully uploaded ${count} image(s) to Drive with unique names. Click "Scan" to process.`;
}

function validateBeforeExportMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Main sheet not found. Run "Initialize Main Sheet" first.');
    return;
  }
  validateBeforeExport(sheet);
}

function showProcessingSidebar(sheetName, content) {
  const cache = CacheService.getScriptCache();
  let fullLog = cache.get('SIDEBAR_LOG') || '';
  
  // If content is provided, append it; otherwise just show existing
  if (content) {
    fullLog += content + '\n';
    cache.put('SIDEBAR_LOG', fullLog, 21600); // Save for 6 hours
  }

  const safeContent = fullLog.replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
  const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Consolas', 'Monaco', monospace; font-size: 12px; padding: 12px; background: #0f172a; color: #cbd5e1; line-height: 1.5; }
    h3 { color: #38bdf8; margin: 0 0 10px; border-bottom: 1px solid #1e293b; padding-bottom: 5px; display: flex; justify-content: space-between; align-items: center; }
    .log-entry { margin-bottom: 4px; }
    .btn-clear { background: #334155; color: #94a3b8; border: none; padding: 4px 8px; border-radius: 4px; font-size: 10px; cursor: pointer; }
    .btn-clear:hover { background: #ef4444; color: white; }
    #scroll-anchor { height: 1px; }
  </style>
</head>
<body>
  <h3>
    <span>SESSION LOG</span>
    <button class="btn-clear" onclick="google.script.run.clearSidebarLog()">🗑️ Clear</button>
  </h3>
  <div id="log-container">${safeContent}</div>
  <div id="scroll-anchor"></div>
  <script>
    window.scrollTo(0, document.body.scrollHeight);
    // Auto-scroll to bottom on load
    document.getElementById('scroll-anchor').scrollIntoView();
  </script>
</body>
</html>
`;

  const html = HtmlService.createHtmlOutput(htmlContent).setTitle('Vibecode Activity Log');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Clears the persistent sidebar log from cache.
 */
function clearSidebarLog() {
  CacheService.getScriptCache().remove('SIDEBAR_LOG');
}

function showHelp() {
  const helpText = `CORPORATE CARD EXPENSE RECONCILIATION — HELP & INSTRUCTIONS

=== SETUP (one-time) ===
1. Run "Initialize Config Template" to create the Config sheet.
2. Fill in Config TABLE 1: Service Account JSON, Gemini API Key, Drive Root Folder ID.
3. Fill in Config TABLE 2: Account Codes (code + name pairs).
4. Fill in Config TABLE 6: Default I/O code, Cost Center, Tax ID.
5. Run "Initialize Main Sheet" to create the main data sheet.
6. Run "Check Configuration" to verify all settings are correct.

=== MONTHLY WORKFLOW ===
Step 1 — Add Transactions:
  • Manually enter or paste bank transactions into the "Corporate card details" sheet.
  • Required columns for matching: Doc Date (C), Currency Local (H), Amount Purchase (J), Merchant (F).

Step 2 — Add Invoices:
  • Open a DD-MM sheet (e.g. "02-02" for February 2nd).
  • Insert receipt photos: Insert → Image → Insert image OVER cells.
  ⚠️ CRITICAL: Use "Insert image OVER cells" NOT "Place image IN cell".
    Images placed IN cells cannot be read by this system.
  • Repeat for all dates that have transactions.

Step 3 — Process Invoices:
  • Navigate to a DD-MM sheet.
  • Click "Process Invoices in This Sheet".
  • The system will:
      Phase A: OCR each image via Google Cloud Vision
      Phase B: Build batch payload from all OCR results
      Phase C: Send ONE Gemini request to parse all invoices
      Phase D: Match parsed invoices to statement rows and write results

Step 4 — Review:
  • Check Col W (Match Status) for any "⚠️ Needs Review" rows.
  • For ambiguous matches, manually assign the correct row.
  • For transactions with no receipt, set Col W to "❌ No Invoice".

Step 5 — Export:
  • Run "Validate Before Export" to check for missing Account Codes or unprocessed rows.
  • Copy Columns A–X to your company Excel report template.

=== IMAGE REQUIREMENTS ===
  REQUIRED: Insert → Image → Insert image OVER cells
  NOT SUPPORTED: Insert → Image → Place image IN cell
  Max image size: 5MB per image
  Supported formats: JPG, PNG, PDF-as-image

=== MATCH STATUS MEANINGS ===
  ⏳ Pending     — Imported/Entered, awaiting invoice match
  ✅ Matched     — Invoice found and matched automatically
  ⚠️ Needs Review — Two or more possible matches found; user must choose
  ❌ No Invoice  — Confirmed: no receipt exists for this transaction`;

  SpreadsheetApp.getUi().alert(helpText);
}

/**
 * Menu action to resize and arrange all images in the active day sheet.
 */
function arrangeActiveSheetImagesMenu() {
  try {
    arrangeAndResizeImagesInSheet();
    SpreadsheetApp.getUi().alert('Images have been arranged and resized successfully.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

