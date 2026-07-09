/**
 * Unit tests and diagnostics for the system.
 */

function testConfig() {
  try {
    _configCache = null;
    const config = getConfig();
    const msg = `Configuration OK\n\nGemini Key: ${config.geminiKey ? config.geminiKey.substring(0,8) + '...' : 'MISSING'}\nDrive Folder: ${config.folderId}\nAccount Codes: ${config.accountCodes.length}`;
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Config Error: ' + e.message);
  }
}

function testServiceAccountToken() {
  try {
    const config = getConfig();
    const token = getServiceAccountToken(config.visionJson, VISION_SCOPE);
    SpreadsheetApp.getUi().alert(`Token OK: ${token.substring(0, 20)}...`);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Token Error: ' + e.message);
  }
}

/**
 * Updated: Test OCR using the first image found in the Drive root folder
 * (Since images are no longer inserted into sheets)
 */
function testCloudVisionFirstImage() {
  try {
    const config = getConfig();
    const rootFolder = DriveApp.getFolderById(config.folderId);
    const files = rootFolder.getFiles();
    
    let targetFile = null;
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType().startsWith('image/')) {
        targetFile = file;
        break;
      }
    }

    if (!targetFile) {
      SpreadsheetApp.getUi().alert('No image files found in the root Drive folder to test.');
      return;
    }

    const base64 = Utilities.base64Encode(targetFile.getBlob().getBytes());
    const result = callCloudVision(base64, config.visionJson);
    
    if (result.error) {
      SpreadsheetApp.getUi().alert('OCR Error: ' + result.error);
    } else {
      SpreadsheetApp.getUi().alert(`OCR Result for [${targetFile.getName()}]:\n\n${result.text.substring(0, 1000)}`);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Debug OCR Error: ' + e.message);
  }
}

/**
 * Main entry point for unit tests
 */
function runAllTests() {
  const tests = [
    testParseDate, 
    testParseAmount, 
    testDateDiff, 
    testMatchToStatement,
    testParseGeminiBatchResponse, 
    testFormatAndParseDaySheetName,
    testImageAutoSpacingAndScaling
  ];
  
  const results = tests.map(fn => {
    try {
      return fn();
    } catch (e) {
      return { name: fn.name, passed: false, error: e.message };
    }
  });

  const passedCount = results.filter(r => r.passed).length;
  const summary = results.map(r => `${r.passed ? '✅' : '❌'} ${r.name}${r.passed ? '' : ': ' + r.error}`).join('\n');
  
  SpreadsheetApp.getUi().alert(`Test Results (${passedCount}/${tests.length}):\n\n${summary}`);
}

// --- UNIT TEST IMPLEMENTATIONS ---

function testParseDate() {
  const d = parseDate('2026-02-27');
  const pass = d instanceof Date && d.getFullYear() === 2026 && d.getMonth() === 1 && d.getDate() === 27;
  return { name: 'testParseDate', passed: pass, error: pass ? null : 'Failed to parse YYYY-MM-DD' };
}

function testParseAmount() {
  const amt = parseAmount('1.500.000');
  const pass = amt === 1500000;
  return { name: 'testParseAmount', passed: pass, error: pass ? null : 'Failed to parse 1.500.000' };
}

function testDateDiff() {
  const d1 = new Date(2026, 1, 27);
  const d2 = new Date(2026, 1, 28);
  const diff = dateDiff(d1, d2);
  const pass = diff === 1;
  return { name: 'testDateDiff', passed: pass, error: pass ? null : 'Date diff calculation wrong' };
}

function testMatchToStatement() {
  const mockRows = [{ rowNum: 2, date: new Date(2026,1,27), amount: 1000, currency: 'VND', status: STATUS.PENDING }];
  const invoice = { date: '2026-02-27', amount: 1000, currency: 'VND' };
  const res = matchToStatement(invoice, mockRows);
  const pass = res && res.row === 2;
  return { name: 'testMatchToStatement', passed: pass, error: pass ? null : 'Failed to match exact row' };
}

function testParseGeminiBatchResponse() {
  const json = '[{"imageIndex":1, "date":"2026-02-02", "amount":1000, "currency":"VND", "merchant":"X", "summary":"Y"}]';
  const res = parseGeminiBatchResponse(json);
  const pass = Array.isArray(res) && res.length === 1;
  return { name: 'testParseGeminiBatchResponse', passed: pass, error: pass ? null : 'Failed to parse Gemini array' };
}

function testFormatAndParseDaySheetName() {
  const name = formatDaySheetName(new Date(2026, 1, 27));
  const pass = name === '27-02';
  return { name: 'testFormatAndParseDaySheetName', passed: pass, error: pass ? null : 'Format failed' };
}

function testImageAutoSpacingAndScaling() {
  const mockImages = [];
  const mockSheet = {
    getName: () => '27-02',
    getImages: () => mockImages,
    getRange: (row, col) => ({
      getRow: () => row,
      getColumn: () => col
    }),
    insertImage: (blob, col, row) => {
      const mockImg = {
        width: (blob && blob.width) || 400,
        height: (blob && blob.height) || 600,
        origWidth: (blob && blob.width) || 400,
        origHeight: (blob && blob.height) || 600,
        col: col,
        row: row,
        xOffset: 0,
        yOffset: 0,
        anchor: null,
        getInherentWidth: () => mockImg.origWidth,
        getInherentHeight: () => mockImg.origHeight,
        getWidth: () => mockImg.width,
        getHeight: () => mockImg.height,
        setWidth: (w) => { mockImg.width = w; },
        setHeight: (h) => { mockImg.height = h; },
        setAnchorCell: (cell) => { mockImg.anchor = cell; },
        setAnchorCellXOffset: (x) => { mockImg.xOffset = x; },
        setAnchorCellYOffset: (y) => { mockImg.yOffset = y; },
        getAnchorCell: () => mockImg.anchor || {
          getColumn: () => mockImg.col,
          getRow: () => mockImg.row
        }
      };
      mockImages.push(mockImg);
      return mockImg;
    }
  };

  // Test 1: Insert first image (800x400 px) -> should resize to 500x250 px, offset 0
  insertInvoiceImageToDaySheet(mockSheet, { width: 800, height: 400 });
  const img1 = mockImages[0];
  if (img1.width !== 500 || img1.height !== 250) {
    return { name: 'testImageAutoSpacingAndScaling', passed: false, error: `Img 1 expected 500x250, got ${img1.width}x${img1.height}` };
  }
  if (img1.xOffset !== 0 || img1.yOffset !== 0) {
    return { name: 'testImageAutoSpacingAndScaling', passed: false, error: `Img 1 expected offset 0, got ${img1.xOffset}, ${img1.yOffset}` };
  }

  // Test 2: Insert second image (300x600 px) -> should resize to 250x500 px, offset 510
  insertInvoiceImageToDaySheet(mockSheet, { width: 300, height: 600 });
  const img2 = mockImages[1];
  if (img2.width !== 250 || img2.height !== 500) {
    return { name: 'testImageAutoSpacingAndScaling', passed: false, error: `Img 2 expected 250x500, got ${img2.width}x${img2.height}` };
  }
  if (img2.xOffset !== 510 || img2.yOffset !== 0) {
    return { name: 'testImageAutoSpacingAndScaling', passed: false, error: `Img 2 expected X offset 510, got ${img2.xOffset}` };
  }

  return { name: 'testImageAutoSpacingAndScaling', passed: true, error: null };
}

