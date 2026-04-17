// ─── SHEET NAMES ─────────────────────────────────────────────────────────────
const CONFIG_SHEET_NAME   = 'Config';
const MAIN_SHEET_NAME     = 'Corporate card details';
const AUDIT_SHEET_NAME    = 'OCR Audit';

// ─── PROCESSING PARAMS ────────────────────────────────────────────────────────
const DATA_START_ROW       = 2;
const MATCH_DATE_TOLERANCE = 4;
const MIN_MATCH_SCORE      = 0.8;
const DONE_SUBFOLDER       = 'Unmatched';
const ERROR_SUBFOLDER      = 'Error';

// ─── AI CONFIGURATION ─────────────────────────────────────────────────────────
const VISION_SCOPE        = 'https://www.googleapis.com/auth/cloud-vision';
const VISION_ENDPOINT     = 'https://vision.googleapis.com/v1/images:annotate';
const GEMINI_BASE_URL     = 'https://generativelanguage.googleapis.com/v1beta/models';
const TOKEN_CACHE_KEY     = 'vision_access_token';
const TOKEN_EXPIRY_BUFFER = 60 * 1000;

// ─── BANK FILE — DEFAULT COLUMN INDICES (0-based) ─────────────────────────────
const DEFAULT_COL_MAP = {
  date:     2,
  amount:   9,
  currency: 7,
  merchant: 5,
};

// ─── OUTPUT SHEET — COLUMN NUMBERS (1-based) ─────────────────────────────────
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

// ─── MATCH STATUS VALUES ──────────────────────────────────────────────────────
const STATUS = {
  PENDING:    '⏳ Pending',
  MATCHED:    '✅ Matched',
  NO_INVOICE: '❌ No Invoice',
  AMBIGUOUS:  '⚠️ Needs Review',
};

// ─── LOG PREFIXES ─────────────────────────────────────────────────────────────
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

// ─── SCRIPT-SCOPED CONFIG CACHE ───────────────────────────────────────────────
let _configCache = null;
