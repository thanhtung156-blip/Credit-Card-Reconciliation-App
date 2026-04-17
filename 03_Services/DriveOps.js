/**
 * Get or refresh a Service Account OAuth2 token, with caching.
 */
function getServiceAccountToken(visionJson, scope) {
  try {
    const cached = PropertiesService.getScriptProperties().getProperty(TOKEN_CACHE_KEY);
    if (cached) {
      const { token, expiresAt } = JSON.parse(cached);
      if (token && expiresAt && (expiresAt - Date.now()) > TOKEN_EXPIRY_BUFFER) {
        return token;
      }
    }
  } catch (e) {}

  let sa = JSON.parse(visionJson);
  const now = Math.floor(Date.now() / 1000);
  const header = Utilities.base64EncodeWebSafe(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const claim = Utilities.base64EncodeWebSafe(JSON.stringify({
    iss: sa.client_email,
    scope,
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now,
  }));
  const signInput = `${header}.${claim}`;
  const signature = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(signInput, sa.private_key));
  const jwt = `${signInput}.${signature}`;

  const resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    payload: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200) throw new Error('Token exchange failed');
  const tokenData = JSON.parse(resp.getContentText());
  const token = tokenData.access_token;
  const expiresAt = Date.now() + (tokenData.expires_in * 1000);

  PropertiesService.getScriptProperties().setProperty(TOKEN_CACHE_KEY, JSON.stringify({ token, expiresAt }));
  return token;
}

function getOrCreateMonthFolder(rootFolder, date) {
  const name = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  return getOrCreateSubFolder(rootFolder, name);
}

function getOrCreateDayFolder(monthFolder, date) {
  const name = String(date.getDate()).padStart(2, '0');
  return getOrCreateSubFolder(monthFolder, name);
}

function getOrCreateSubFolder(parentFolder, name) {
  const existing = parentFolder.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parentFolder.createFolder(name);
}

function getOrCreateDoneFolder(rootFolder, date) {
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "MM-yyyy");
  const name = `DONE-${dateStr}`;
  return getOrCreateSubFolder(rootFolder, name);
}

/**
 * Move an image file to the appropriate YYYY-MM/DD folder.
 */
function moveImageToDayFolder(blob, rootFolderId, date) {
  const root = DriveApp.getFolderById(rootFolderId);
  const monthFolder = getOrCreateMonthFolder(root, date);
  const dayFolder = getOrCreateDayFolder(monthFolder, date);
  return dayFolder.createFile(blob);
}

/**
 * Move an image file to the single DONE-MM folder.
 */
function moveImageToDoneFolder(blob, rootFolderId, date) {
  const root = DriveApp.getFolderById(rootFolderId);
  const doneFolder = getOrCreateDoneFolder(root, date);
  return doneFolder.createFile(blob);
}

/**
 * Rename a matched file to [Done-DD-MM] original_name
 */
function renameFileWithPrefix(fileObj, date) {
  if (!fileObj || !date) return;
  const prefix = `[Done-${Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd-MM')}] `;
  const oldName = fileObj.getName();
  if (!oldName.startsWith('[Done-')) {
    fileObj.setName(prefix + oldName);
  }
}

/**
 * Move an image file to the Unmatched folder.
 */
function moveImageToUnmatched(fileObj, rootFolder) {
  const unmatchedFolder = getOrCreateSubFolder(rootFolder, DONE_SUBFOLDER);
  fileObj.moveTo(unmatchedFolder);
}

/**
 * Move an image file to the Error folder.
 */
function moveImageToError(fileObj, errorMsg, rootFolder) {
  const errorFolder = getOrCreateSubFolder(rootFolder, ERROR_SUBFOLDER);
  fileObj.moveTo(errorFolder);
  logMsg(LOG.DRIVE, 'ERROR', `Moved ${fileObj.getName()} to Error: ${errorMsg}`);
}
