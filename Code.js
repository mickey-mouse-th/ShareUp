// ============================================================
// ShareUp - Expense Splitting App (Google Apps Script)
// ============================================================

var SPREADSHEET_NAME = 'ShareUp_Database';
var CACHE_EXPIRY = 21600; // 6 hours - hard max allowed by CacheService
var DEFAULT_SETTINGS = {
  sessionMinutes: 30 * 24 * 60, // 30 days, matches the old fixed SESSION_DAYS
  pwMinLength: 6,
  pwRequireUpper: false,
  pwRequireLower: false,
  pwRequireNumber: false,
  pwRequireSpecial: false
};
// Mirrors Shared_css.html's :root defaults, so the theme picker starts accurate.
// buildThemeCss() only overrides keys actually saved, so nothing changes until
// an admin explicitly saves.
var DEFAULT_THEME = {
  p: '#4F46E5', pLt: '#4338CA',
  bg: '#F1F5F9', s1: '#FFFFFF', s2: '#F8FAFC', bd: '#E2E8F0',
  t1: '#0F172A', t2: '#64748B', t3: '#94A3B8',
  g: '#15803D', r: '#EF4444', o: '#B45309', b: '#0F766E',
  hdrBg: '#FFFFFF', navBg: '#FFFFFF', overlayBg: '#0F172A'
};

// ----------------------------------------------------------------
// Database Setup
// ----------------------------------------------------------------

function getSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SPREADSHEET_ID');
  var ss;

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      ssId = null;
    }
  }

  if (!ssId) {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    props.setProperty('SPREADSHEET_ID', ss.getId());
    initSheets(ss);

    // Create default admin account
    var adminPassword = hashPassword('admin123');
    var adminId = Utilities.getUuid();
    var now = new Date().toISOString();
    var accountsSheet = ss.getSheetByName('Accounts');
    accountsSheet.appendRow([adminId, 'Admin', 'admin', adminPassword, now, now, 'admin', 'active', '', '']);

    // Create "Me" friend for admin
    var friendsSheet = ss.getSheetByName('Friends');
    friendsSheet.appendRow([Utilities.getUuid(), adminId, 'Me', 'true']);
  }

  return ss;
}

function initSheets(ss) {
  // Remove default sheet if needed
  var defaultSheet = ss.getSheetByName('Sheet1');

  var accountsSheet = ss.insertSheet('Accounts');
  accountsSheet.appendRow(['id', 'displayName', 'username', 'password', 'firstLogin', 'lastLogin', 'role', 'status', 'email', 'photo']);

  var friendsSheet = ss.insertSheet('Friends');
  friendsSheet.appendRow(['id', 'accountId', 'name', 'isSelf']);

  var eventsSheet = ss.insertSheet('Events');
  eventsSheet.appendRow(['id', 'name', 'accountId', 'createdAt', 'active', 'icon']);

  var detailsSheet = ss.insertSheet('Details');
  detailsSheet.appendRow(['transactionId', 'eventId', 'payId', 'totalAmount', 'description', 'createdAt', 'splits']);

  var eventFriendsSheet = ss.insertSheet('EventFriends');
  eventFriendsSheet.appendRow(['id', 'eventId', 'friendId', 'createdAt']);

  var eventSharesSheet = ss.insertSheet('EventShares');
  eventSharesSheet.appendRow(['eventId', 'token', 'createdAt', 'permission']);

  var sessionsSheet = ss.insertSheet('Sessions');
  sessionsSheet.appendRow(['token', 'accountId', 'userInfo', 'createdAt', 'expiresAt']);

  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }
}

// ----------------------------------------------------------------
// Utilities
// ----------------------------------------------------------------

function hashPassword(pw) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
  }).join('');
}

function generateToken() {
  return Utilities.getUuid() + '-' + Utilities.getUuid();
}

// The one error shape every RPC function returns - accepts either a caught
// exception or a plain string message (both have .toString()).
function _fail(err) {
  return { success: false, error: err.toString() };
}

// Row index of the first row whose column `col` equals val, or -1.
function _findRowByCol(data, col, val) {
  for (var i = 1; i < data.length; i++) if (data[i][col] === val) return i;
  return -1;
}

// Shorthand for the common case: match against column 0 (most sheets here -
// Accounts/Sessions/EventShares/... - are keyed by their first column).
function _findRow(data, id) {
  return _findRowByCol(data, 0, id);
}

function getCache() {
  return CacheService.getScriptCache();
}

function getEventFriendsSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('EventFriends');
  if (!sheet) {
    sheet = ss.insertSheet('EventFriends');
    sheet.appendRow(['id', 'eventId', 'friendId', 'createdAt']);
  }
  return sheet;
}

function getEventSharesSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('EventShares');
  if (!sheet) {
    sheet = ss.insertSheet('EventShares');
    sheet.appendRow(['eventId', 'token', 'createdAt', 'permission']);
  }
  return sheet;
}

// Rows written before 'permission' existed have a blank 4th cell - treat
// that the same as 'view' (the original, only behavior), no migration needed.
function _sharePermission(row) {
  return row[3] === 'edit' ? 'edit' : 'view';
}

function getSessionsSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Sessions');
  if (!sheet) {
    sheet = ss.insertSheet('Sessions');
    sheet.appendRow(['token', 'accountId', 'userInfo', 'createdAt', 'expiresAt']);
  }
  return sheet;
}

function getSettlementPaymentsSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('SettlementPayments');
  if (!sheet) {
    sheet = ss.insertSheet('SettlementPayments');
    sheet.appendRow(['id', 'eventId', 'fromId', 'toId', 'amount', 'markedAt']);
  }
  return sheet;
}

// A transaction marked paid here is excluded from settlement math entirely
// (as if it never happened) - separate from SettlementPayments above, which
// marks a net aggregated debt as paid rather than a single expense.
function getTransactionPaymentsSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('TransactionPayments');
  if (!sheet) {
    sheet = ss.insertSheet('TransactionPayments');
    sheet.appendRow(['transactionId', 'eventId', 'paidAt']);
  }
  return sheet;
}

// transactionId is a UUID unique across all events, so a flat set (no event
// scoping needed) is enough to check "is this transaction paid".
function _paidTransactionSet(dataOpt) {
  var data = dataOpt || getTransactionPaymentsSheet().getDataRange().getValues();
  var set = {};
  for (var i = 1; i < data.length; i++) set[data[i][0]] = true;
  return set;
}

// One row per PHOTO, not a column on Details (a transaction fans out into
// one Details row per participant, so a column would duplicate each image N
// times). Pre-Drive rows have a blank 'fileId' and keep their original
// base64 data-URI in 'slip'/'slipHi' - still renders fine, no migration needed.
function getTransactionSlipsSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('TransactionSlips');
  if (!sheet) {
    sheet = ss.insertSheet('TransactionSlips');
    sheet.appendRow(['transactionId', 'slip', 'updatedAt', 'id', 'slipHi', 'fileId']);
  }
  return sheet;
}

// Pre-multi-photo rows have a blank id - backfill lazily (like
// _findSelfFriendRow) so old photos stay individually deletable.
function _backfillSlipIds(sheet, data) {
  var changed = false;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][3]) { data[i][3] = Utilities.getUuid(); changed = true; }
  }
  if (changed) {
    sheet.getRange(2, 4, data.length - 1, 1).setValues(data.slice(1).map(function (r) { return [r[3]] }));
  }
  return data;
}

// ----------------------------------------------------------------
// Slip photo storage (Google Drive) - a Sheets cell caps out around 50,000
// characters, far too small for a real photo, so slips are uploaded as Drive
// files instead and only the resulting URLs/fileId are kept in the sheet.
// ----------------------------------------------------------------

function _getSlipsFolder() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('SLIPS_FOLDER_ID');
  if (folderId) {
    try { return DriveApp.getFolderById(folderId) } catch (e) { /* fall through and recreate */ }
  }
  var folder = DriveApp.createFolder('ShareUp_Slips');
  props.setProperty('SLIPS_FOLDER_ID', folder.getId());
  return folder;
}

// dataUri: "data:<mime>;base64,<data>" from the client's canvas compression
// step. Kept PRIVATE (no public Drive sharing) - direct Drive embed URLs
// proved unreliable when hotlinked from this app's sandboxed iframe, so
// bytes are instead fetched back through google.script.run (see
// getSlipImage/getSlipImageViaShare) and turned into a data: URI client-side.
function _uploadSlipToDrive(dataUri) {
  var m = /^data:([^;]+);base64,(.*)$/.exec(dataUri || '');
  if (!m) throw new Error('Invalid image data');
  var mimeType = m[1], base64 = m[2];
  if (base64.length > 2000000) throw new Error('Photo is too large - please try a smaller one');
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mimeType, 'slip-' + Utilities.getUuid());
  var file = _getSlipsFolder().createFile(blob);
  return { fileId: file.getId() };
}

// Reads a slip file's bytes back out for the client to display - fileId must
// already appear in TransactionSlips, otherwise this would be an open proxy
// for reading ANY file in the deploying account's Drive by id (doGet/RPC
// calls always execute as that account per executeAs: USER_DEPLOYING).
function _slipBase64(fileId) {
  var blob = DriveApp.getFileById(fileId).getBlob();
  return { mimeType: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes()) };
}

function _readSlipImage(fileId) {
  if (!_isKnownSlipFile(fileId)) return _fail('Photo not found');
  var d = _slipBase64(fileId);
  return { success: true, mimeType: d.mimeType, base64: d.base64 };
}

function getSlipImage(token, fileId) {
  try {
    requireAuth(token);
    return _readSlipImage(fileId);
  } catch (e) {
    return _fail(e);
  }
}

function getSlipImageViaShare(shareToken, fileId) {
  try {
    if (!_shareEventId(shareToken, false)) return _fail('Invalid link');
    return _readSlipImage(fileId);
  } catch (e) {
    return _fail(e);
  }
}

function _isKnownSlipFile(fileId) {
  var data = getTransactionSlipsSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] === fileId) return true;
  }
  return false;
}

function _deleteSlipFile(fileId) {
  if (!fileId) return; // pre-Drive rows have no fileId - nothing to trash
  try { DriveApp.getFileById(fileId).setTrashed(true) } catch (e) { /* already gone - ignore */ }
}

function _trashSlipFilesForTx(transactionId) {
  var data = getTransactionSlipsSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === transactionId && data[i][5]) _deleteSlipFile(data[i][5]);
  }
}

function _trashSlipFilesForTxSet(txIdSet) {
  var data = getTransactionSlipsSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (txIdSet[data[i][0]] && data[i][5]) _deleteSlipFile(data[i][5]);
  }
}

function _lookupSession(token) {
  var sheet = getSessionsSheet();
  var data = sheet.getDataRange().getValues();
  var row = _findRow(data, token);
  if (row === -1) return null;
  if (new Date() >= new Date(data[row][4])) { sheet.deleteRow(row + 1); return null }
  var userInfo = JSON.parse(data[row][2]);
  if (_isAccountDisabled(userInfo.id)) { sheet.deleteRow(row + 1); return null }
  return { row: row + 1, userInfo: userInfo, expiresAt: data[row][4] };
}

// Caps the cache entry so it never outlives the session's real expiry - matters
// when an admin configures a short sessionMinutes value (e.g. for testing).
function _cacheTtlFor(expiresAtIso) {
  var remainingSec = Math.floor((new Date(expiresAtIso).getTime() - Date.now()) / 1000);
  return Math.max(1, Math.min(CACHE_EXPIRY, remainingSec));
}

// Only consulted on session-cache misses (~every CACHE_EXPIRY), so a disabled
// account is locked out within a few hours without a per-request sheet read.
function _isAccountDisabled(accountId) {
  var data = getSpreadsheet().getSheetByName('Accounts').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === accountId) return data[i][7] === 'disabled';
  }
  return false;
}

function _cleanExpiredSessions() {
  try {
    var sheet = getSessionsSheet();
    var data = sheet.getDataRange().getValues();
    var now = new Date();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][4] && now > new Date(data[i][4])) {
        sheet.deleteRow(i + 1);
      }
    }
  } catch (e) {}
}

// ----------------------------------------------------------------
// Entry Point
// ----------------------------------------------------------------

// Looked up at render time (not via getSharedEventView) so doGet can skip
// sending the add/edit/delete transaction markup entirely for the common
// view-only case, instead of shipping it and hiding it with CSS.
function _sharePermissionByToken(shareToken) {
  if (!shareToken) return 'view';
  var data = getEventSharesSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === shareToken) return _sharePermission(data[i]);
  }
  return 'view';
}

function doGet(e) {
  var rawToken = e && e.parameter && e.parameter.share;
  // Strict allowlist so this can be embedded directly into the page's inline script safely.
  var shareToken = (rawToken && /^[a-zA-Z0-9-]{10,100}$/.test(rawToken)) ? rawToken : '';

  // Optional ?tk=<token> - set by Auth_js.html right after a "remember me"
  // login. Safari wipes localStorage for this app's sandboxed frame on close
  // but keeps the tab's last URL, so reopening it re-runs doGet with this
  // param still attached and login survives. Re-validated here so a
  // revoked/expired token just falls back to logged-out instead of erroring.
  var bootToken = '', bootUser = null;
  var rawTk = e && e.parameter && e.parameter.tk;
  if (!shareToken && rawTk && /^[a-zA-Z0-9-]{10,100}$/.test(rawTk)) {
    try {
      var found = _lookupSession(rawTk);
      if (found) { bootToken = rawTk; bootUser = found.userInfo; }
    } catch (err) { /* ignore - falls back to logged-out */ }
  }

  var tpl = HtmlService.createTemplateFromFile('Index');
  tpl.shareToken = shareToken;
  tpl.sharePermission = shareToken ? _sharePermissionByToken(shareToken) : '';
  tpl.bootToken = bootToken;
  tpl.bootUser = bootUser;
  return tpl.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle(shareToken ? 'ShareUp - Shared Event' : 'ShareUp - Expense Splitting');
}

// Must use createTemplateFromFile().evaluate(), not createHtmlOutputFromFile()
// - the latter returns raw content and never runs <?...?> scriptlets, which
// silently broke ThemeOverride.html's saved-theme CSS injection. Safe for
// every other included file too (no scriptlets = passes through unchanged).
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

// ----------------------------------------------------------------
// App Settings (session length, password policy) - singleton config
// stored in Script Properties, not a sheet, since it's a single object
// with no per-row semantics.
// ----------------------------------------------------------------

function getAppSettings() {
  var raw = PropertiesService.getScriptProperties().getProperty('APP_SETTINGS');
  var saved = raw ? JSON.parse(raw) : {};
  var merged = {};
  for (var k in DEFAULT_SETTINGS) {
    merged[k] = (saved[k] !== undefined) ? saved[k] : DEFAULT_SETTINGS[k];
  }
  return merged;
}

// No auth required: the register form (pre-login) and change-password form
// both need to show the current rules before the user has a session token.
function getPasswordPolicy() {
  var s = getAppSettings();
  return {
    pwMinLength: s.pwMinLength,
    pwRequireUpper: s.pwRequireUpper,
    pwRequireLower: s.pwRequireLower,
    pwRequireNumber: s.pwRequireNumber,
    pwRequireSpecial: s.pwRequireSpecial
  };
}

function _validatePassword(pw, settingsOpt) {
  var s = settingsOpt || getAppSettings();
  if (!pw || pw.length < s.pwMinLength) return 'Password must be at least ' + s.pwMinLength + ' characters';
  if (s.pwRequireUpper && !/[A-Z]/.test(pw)) return 'Password must include an uppercase letter';
  if (s.pwRequireLower && !/[a-z]/.test(pw)) return 'Password must include a lowercase letter';
  if (s.pwRequireNumber && !/[0-9]/.test(pw)) return 'Password must include a number';
  if (s.pwRequireSpecial && !/[^A-Za-z0-9]/.test(pw)) return 'Password must include a special character';
  return null;
}

function getSettings(token) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return _fail('Forbidden');
    return { success: true, settings: getAppSettings() };
  } catch (e) {
    return _fail(e);
  }
}

function updateSettings(token, settings) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return _fail('Forbidden');
    var minutes = parseInt(settings.sessionMinutes, 10);
    // Floor at 5 minutes - anything shorter makes it easy to accidentally lock
    // yourself (or every user) out via a mistyped value (e.g. minutes vs hours).
    if (!minutes || minutes < 5) return _fail('Session length must be at least 5 minutes');
    var minLen = parseInt(settings.pwMinLength, 10);
    if (!minLen || minLen < 1) minLen = 1;
    var merged = {
      sessionMinutes: minutes,
      pwMinLength: minLen,
      pwRequireUpper: settings.pwRequireUpper === true,
      pwRequireLower: settings.pwRequireLower === true,
      pwRequireNumber: settings.pwRequireNumber === true,
      pwRequireSpecial: settings.pwRequireSpecial === true
    };
    PropertiesService.getScriptProperties().setProperty('APP_SETTINGS', JSON.stringify(merged));
    return { success: true, settings: merged };
  } catch (e) {
    return _fail(e);
  }
}

// ----------------------------------------------------------------
// App Theme - global color palette, applied identically to every user.
// Stored in Script Properties (singleton, like APP_SETTINGS above).
// buildThemeCss() only emits overrides for keys actually saved, so an
// empty/never-saved theme leaves Shared_css.html's own colors untouched.
// ----------------------------------------------------------------

var HEX_COLOR_RE = /^#[0-9a-fA-F]{6}$/;

function _rawTheme() {
  var raw = PropertiesService.getScriptProperties().getProperty('APP_THEME');
  return raw ? JSON.parse(raw) : {};
}

function getAppTheme() {
  var saved = _rawTheme();
  var merged = {};
  for (var k in DEFAULT_THEME) {
    merged[k] = HEX_COLOR_RE.test(saved[k]) ? saved[k] : DEFAULT_THEME[k];
  }
  return merged;
}

// No auth required: needs to be embeddable server-side into Index.html
// before login (see buildThemeCss/ThemeOverride.html), same reasoning as
// getPasswordPolicy() above.
function getTheme() {
  return { success: true, theme: getAppTheme() };
}

function saveTheme(token, theme) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return _fail('Forbidden');
    var saved = _rawTheme();
    var merged = {};
    for (var k in DEFAULT_THEME) merged[k] = saved[k];
    for (var key in DEFAULT_THEME) {
      if (theme && HEX_COLOR_RE.test(theme[key])) merged[key] = theme[key];
    }
    PropertiesService.getScriptProperties().setProperty('APP_THEME', JSON.stringify(merged));
    return { success: true, theme: getAppTheme() };
  } catch (e) {
    return _fail(e);
  }
}

// :root override CSS for only the keys actually saved - included right after
// Shared_css.html so it wins the cascade with zero flash-of-unstyled-color.
function buildThemeCss() {
  var saved = _rawTheme();
  var decls = [];
  var cssVarName = {
    p: '--p', pLt: '--p-lt', bg: '--bg', s1: '--s1', s2: '--s2', bd: '--bd',
    t1: '--t1', t2: '--t2', t3: '--t3', g: '--g', r: '--r', o: '--o', b: '--b',
    hdrBg: '--hdr-bg', navBg: '--nav-bg', overlayBg: '--overlay-bg'
  };
  for (var k in DEFAULT_THEME) {
    if (HEX_COLOR_RE.test(saved[k])) decls.push(cssVarName[k] + ':' + saved[k]);
  }
  return decls.length ? (':root{' + decls.join(';') + '}') : '';
}

// ----------------------------------------------------------------
// Auth
// ----------------------------------------------------------------

function loginUser(username, password) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Accounts');
    var data = sheet.getDataRange().getValues();
    var hashed = hashPassword(password);

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[2].toLowerCase() === username.toLowerCase() && row[3] === hashed) {
        if (row[7] === 'disabled') return _fail('This account has been disabled');
        // Update lastLogin
        sheet.getRange(i + 1, 6).setValue(new Date().toISOString());

        var token = generateToken();
        var userInfo = {
          id: row[0],
          displayName: row[1],
          username: row[2],
          role: row[6]
        };
        var now = new Date();
        var sessionMinutes = getAppSettings().sessionMinutes;
        var expires = new Date(now.getTime() + sessionMinutes * 60000);
        var cacheTtl = Math.max(1, Math.min(CACHE_EXPIRY, sessionMinutes * 60));
        getCache().put('token_' + token, JSON.stringify(userInfo), cacheTtl);
        getSessionsSheet().appendRow([token, row[0], JSON.stringify(userInfo), now.toISOString(), expires.toISOString()]);
        _cleanExpiredSessions();
        return { success: true, token: token, user: userInfo, url: ScriptApp.getService().getUrl() + '?tk=' + encodeURIComponent(token) };
      }
    }
    return _fail('Invalid username or password');
  } catch (e) {
    return _fail(e);
  }
}

function registerUser(displayName, username, password) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Accounts');
    var data = sheet.getDataRange().getValues();

    // Check username uniqueness
    for (var i = 1; i < data.length; i++) {
      if (data[i][2].toLowerCase() === username.toLowerCase()) {
        return _fail('Username already taken');
      }
    }

    var pwErr = _validatePassword(password);
    if (pwErr) return _fail(pwErr);

    var now = new Date().toISOString();
    var id = Utilities.getUuid();
    var hashed = hashPassword(password);
    sheet.appendRow([id, displayName, username.toLowerCase(), hashed, now, now, 'user', 'active', '', '']);

    // Create "Me" friend
    var friendsSheet = ss.getSheetByName('Friends');
    friendsSheet.appendRow([Utilities.getUuid(), id, 'Me', 'true']);

    return { success: true };
  } catch (e) {
    return _fail(e);
  }
}

function logoutUser(token) {
  try {
    getCache().remove('token_' + token);
    var sheet = getSessionsSheet();
    var row = _findRow(sheet.getDataRange().getValues(), token);
    if (row !== -1) sheet.deleteRow(row + 1);
    return { success: true };
  } catch (e) {
    return _fail(e);
  }
}

function getSessionUser(token) {
  try {
    var cached = getCache().get('token_' + token);
    if (cached) return { success: true, user: JSON.parse(cached) };
    var found = _lookupSession(token);
    if (!found) return _fail('Session expired');
    getCache().put('token_' + token, JSON.stringify(found.userInfo), _cacheTtlFor(found.expiresAt));
    return { success: true, user: found.userInfo };
  } catch (e) {
    return _fail(e);
  }
}

function requireAuth(token) {
  var cached = getCache().get('token_' + token);
  if (cached) return JSON.parse(cached);
  var found = _lookupSession(token);
  if (!found) throw new Error('Unauthorized');
  getCache().put('token_' + token, JSON.stringify(found.userInfo), _cacheTtlFor(found.expiresAt));
  return found.userInfo;
}

// Refreshes the cached userInfo for the CURRENT session/device only (the one
// that made this request) so a profile edit shows up immediately without
// re-login. Other devices logged into the same account keep their own
// cached copy until it naturally expires or they log in again — same
// limitation that already exists for admin-driven role changes.
function _updateSessionUserInfo(token, userInfo) {
  var sheet = getSessionsSheet();
  var data = sheet.getDataRange().getValues();
  var row = _findRow(data, token);
  var ttl = CACHE_EXPIRY;
  if (row !== -1) {
    sheet.getRange(row + 1, 3).setValue(JSON.stringify(userInfo));
    ttl = _cacheTtlFor(data[row][4]);
  }
  getCache().put('token_' + token, JSON.stringify(userInfo), ttl);
}

// ----------------------------------------------------------------
// Profile (self-service account settings)
// ----------------------------------------------------------------

function getMyProfile(token) {
  try {
    var user = requireAuth(token);
    var data = getSpreadsheet().getSheetByName('Accounts').getDataRange().getValues();
    var row = _findRow(data, user.id);
    if (row === -1) return _fail('Account not found');
    return { success: true, displayName: data[row][1], username: data[row][2], photo: data[row][9] || '' };
  } catch (e) {
    return _fail(e);
  }
}

// photo: pass a string to set it ('' clears it); omit/null to leave unchanged.
function updateProfile(token, displayName, photo) {
  try {
    var user = requireAuth(token);
    if (!displayName || !displayName.trim()) return _fail('Name is required');
    var trimmed = displayName.trim();
    var ss = getSpreadsheet();

    var acSheet = ss.getSheetByName('Accounts');
    var acData = acSheet.getDataRange().getValues();
    var acRow = _findRow(acData, user.id);
    if (acRow === -1) return _fail('Account not found');
    acSheet.getRange(acRow + 1, 2).setValue(trimmed);
    if (typeof photo === 'string') acSheet.getRange(acRow + 1, 10).setValue(photo);

    // Keep the self-friend (shown as payer/participant in every event) in sync.
    var frSheet = ss.getSheetByName('Friends');
    var frData = frSheet.getDataRange().getValues();
    var selfRow = _findSelfFriendRow(frData, user.id);
    if (selfRow !== -1) {
      frSheet.getRange(selfRow + 1, 3).setValue(trimmed);
      if (frData[selfRow][3] !== 'true') frSheet.getRange(selfRow + 1, 4).setValue('true'); // backfill pre-migration rows
    }

    var userInfo = { id: user.id, displayName: trimmed, username: user.username, role: user.role };
    _updateSessionUserInfo(token, userInfo);
    return { success: true, user: userInfo };
  } catch (e) {
    return _fail(e);
  }
}

function changePassword(token, currentPassword, newPassword) {
  try {
    var user = requireAuth(token);
    var pwErr = _validatePassword(newPassword);
    if (pwErr) return _fail(pwErr);
    var sheet = getSpreadsheet().getSheetByName('Accounts');
    var data = sheet.getDataRange().getValues();
    var row = _findRow(data, user.id);
    if (row === -1) return _fail('Account not found');
    if (data[row][3] !== hashPassword(currentPassword || '')) return _fail('Current password is incorrect');
    sheet.getRange(row + 1, 4).setValue(hashPassword(newPassword));
    return { success: true };
  } catch (e) {
    return _fail(e);
  }
}

// ----------------------------------------------------------------
// Combined data fetchers (reduce round-trips)
// ----------------------------------------------------------------

function getHomeData(token) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var evData = ss.getSheetByName('Events').getDataRange().getValues();
    var frData = ss.getSheetByName('Friends').getDataRange().getValues();
    var events = [], friends = [], friendMap = {};
    for (var i = 1; i < evData.length; i++) {
      if (evData[i][2] === user.id)
        events.push({ id: evData[i][0], name: evData[i][1], accountId: evData[i][2], createdAt: evData[i][3], active: evData[i][4] !== false, icon: evData[i][5] || '' });
    }
    events.sort(function(a,b){ return b.createdAt > a.createdAt ? 1 : -1 });
    for (var i = 1; i < frData.length; i++) {
      if (frData[i][1] === user.id) {
        friends.push({ id: frData[i][0], accountId: frData[i][1], name: frData[i][2] });
        friendMap[frData[i][0]] = frData[i][2];
      }
    }

    // Settlement state per event for the Home filter tabs, via the same
    // engine as getSummary. Each sheet is read once and shared via *Opt params.
    var rowsByEvent = {};
    events.forEach(function (ev) { rowsByEvent[ev.id] = [] });
    var dtData = ss.getSheetByName('Details').getDataRange().getValues();
    var paidTxSet = _paidTransactionSet();
    for (var i = 1; i < dtData.length; i++) {
      if (rowsByEvent.hasOwnProperty(dtData[i][1]) && !paidTxSet[dtData[i][0]]) {
        var payId = dtData[i][2];
        _splitsToRows(dtData[i]).forEach(function (s) { rowsByEvent[dtData[i][1]].push({ payId: payId, friendId: s.friendId, amount: s.amount }) });
      }
    }
    var efData = getEventFriendsSheet().getDataRange().getValues();
    var spData = getSettlementPaymentsSheet().getDataRange().getValues();
    events.forEach(function (ev) {
      var rows = rowsByEvent[ev.id];
      if (!rows.length) { ev.settled = true; return }
      var evFriendMap = {};
      _getEventFriends(ss, ev.id, user.id, friendMap, efData, dtData).forEach(function (f) { evFriendMap[f.id] = f.name });
      ev.settled = _computeSettlementsWithPaid(rows, evFriendMap, ev.id, spData).every(function (s) { return s.paid });
    });

    return { success: true, events: events, friends: friends };
  } catch (e) { return _fail(e) }
}

// Shared core behind getDetailData (authenticated) and getSharedEventView
// (public share link) so both paths compute details/friends/settlements/slips
// identically instead of maintaining two parallel implementations.
function _buildDetailPayload(ss, eventId, accountId) {
  var dtData = ss.getSheetByName('Details').getDataRange().getValues();
  var paidTxSet = _paidTransactionSet();
  var details = [], rows = [];
  for (var i = 1; i < dtData.length; i++) {
    if (dtData[i][1] === eventId) {
      var txPaid = !!paidTxSet[dtData[i][0]];
      var payId = dtData[i][2];
      _expandDetailRow(dtData[i]).forEach(function (d) { d.paid = txPaid; details.push(d) });
      // Paid transactions are excluded from settlement math - see markTransactionPaid.
      if (!txPaid) _splitsToRows(dtData[i]).forEach(function (s) { rows.push({ payId: payId, friendId: s.friendId, amount: s.amount }) });
    }
  }
  var frRawData = ss.getSheetByName('Friends').getDataRange().getValues();
  var ownedFriendMap = {};
  for (var i = 1; i < frRawData.length; i++) {
    if (frRawData[i][1] === accountId) ownedFriendMap[frRawData[i][0]] = frRawData[i][2];
  }
  var selfRow = _findSelfFriendRow(frRawData, accountId);
  var selfFriendId = selfRow !== -1 ? frRawData[selfRow][0] : null;

  var friends = _getEventFriends(ss, eventId, accountId, ownedFriendMap, undefined, dtData);
  var friendMap = {};
  friends.forEach(function (f) { friendMap[f.id] = f.name });
  // _computeSettlementsWithPaid so share-link visitors see the same "paid"
  // checkmarks the owner does, bundled here to avoid a second round-trip.
  var settlements = _computeSettlementsWithPaid(rows, friendMap, eventId);

  // TransactionSlips isn't keyed by eventId - filter to this event's
  // transactionIds or every event's photos ship down on every load.
  var eventTxIds = {};
  details.forEach(function (d) { eventTxIds[d.transactionId] = true });
  var slipSheet = getTransactionSlipsSheet();
  var slipData = _backfillSlipIds(slipSheet, slipSheet.getDataRange().getValues());
  var slips = {};
  for (var i = 1; i < slipData.length; i++) {
    if (slipData[i][1] && eventTxIds[slipData[i][0]]) {
      var stid = slipData[i][0];
      if (!slips[stid]) slips[stid] = [];
      slips[stid].push({ id: slipData[i][3], slip: slipData[i][1], slipHi: slipData[i][4] || '' });
    }
  }

  // selfFriendId: lets the client show the account's own profile photo.
  return { details: details, friends: friends, settlements: settlements, selfFriendId: selfFriendId, slips: slips };
}

function getDetailData(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Event not found');
    var payload = _buildDetailPayload(ss, eventId, user.id);
    payload.success = true;
    return payload;
  } catch (e) { return _fail(e) }
}

// ----------------------------------------------------------------
// Event-scoped friend membership
// ----------------------------------------------------------------

function _eventOwnedBy(ss, eventId, accountId) {
  var data = ss.getSheetByName('Events').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === eventId && data[i][2] === accountId) return true;
  }
  return false;
}

// Row index of the account's own self-friend (auto-created at registration,
// auto-linked into every new event). Prefers the isSelf marker, falls back
// to the pre-migration "Me" name convention. -1 if none found.
function _findSelfFriendRow(frData, accountId) {
  for (var i = 1; i < frData.length; i++) {
    if (frData[i][1] === accountId && frData[i][3] === 'true') return i;
  }
  for (var i = 1; i < frData.length; i++) {
    if (frData[i][1] === accountId && frData[i][2] === 'Me') return i;
  }
  return -1;
}

// Removes every row where column `col` equals `val`, in one read + one write
// (vs. one deleteRow() per match). Pass dataOpt if already read this request.
function _removeRowsWhere(sheet, col, val, dataOpt) {
  var data = dataOpt || sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  var kept = [data[0]];
  for (var i = 1; i < data.length; i++) {
    if (data[i][col] !== val) kept.push(data[i]);
  }
  if (kept.length === data.length) return; // nothing matched
  sheet.clearContents();
  sheet.getRange(1, 1, kept.length, kept[0].length).setValues(kept);
}

// Like _removeRowsWhere but matches a set of ids at once (deleteEvent needs
// to drop TransactionSlips rows for every transactionId under the event).
function _removeRowsBySet(sheet, col, idSet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  var kept = [data[0]];
  for (var i = 1; i < data.length; i++) {
    if (!idSet[data[i][col]]) kept.push(data[i]);
  }
  if (kept.length === data.length) return;
  sheet.clearContents();
  sheet.getRange(1, 1, kept.length, kept[0].length).setValues(kept);
}

// Friends linked to an event. Self-healing: an event with transactions but
// no EventFriends rows yet gets membership derived from Details and persisted.
// Pass the *Opt params when the caller already read those sheets this request.
function _getEventFriends(ss, eventId, accountId, friendMapOpt, efDataOpt, dtDataOpt) {
  var friendMap = friendMapOpt;
  if (!friendMap) {
    friendMap = {};
    var frData = ss.getSheetByName('Friends').getDataRange().getValues();
    for (var i = 1; i < frData.length; i++) {
      if (frData[i][1] === accountId) friendMap[frData[i][0]] = frData[i][2];
    }
  }

  // Only opens the sheet when actually needed - the common case (efDataOpt
  // already read by the caller, e.g. getHomeData's per-event loop) never
  // touches SpreadsheetApp.openById() at all.
  var efData = efDataOpt || getEventFriendsSheet().getDataRange().getValues();
  var hasAnyLink = false;
  var linkedIds = [];
  for (var i = 1; i < efData.length; i++) {
    if (efData[i][1] === eventId) {
      hasAnyLink = true;
      if (friendMap.hasOwnProperty(efData[i][2])) linkedIds.push(efData[i][2]);
    }
  }

  if (!hasAnyLink) {
    var dtData = dtDataOpt || ss.getSheetByName('Details').getDataRange().getValues();
    var derived = {};
    for (var i = 1; i < dtData.length; i++) {
      if (dtData[i][1] === eventId) {
        if (friendMap.hasOwnProperty(dtData[i][2])) derived[dtData[i][2]] = true;
        _splitsToRows(dtData[i]).forEach(function (s) { if (friendMap.hasOwnProperty(s.friendId)) derived[s.friendId] = true });
      }
    }
    var derivedIds = Object.keys(derived);
    if (derivedIds.length) {
      var now = new Date().toISOString();
      var newRows = derivedIds.map(function (fid) { return [Utilities.getUuid(), eventId, fid, now]; });
      var efSheet = getEventFriendsSheet();
      efSheet.getRange(efSheet.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
      linkedIds = derivedIds;
    }
  }

  return linkedIds.map(function (fid) {
    return { id: fid, name: friendMap[fid] };
  });
}

// Combined fetch for the Add/Manage Friends sheet — one round trip instead of
// separate getFriends + getEventFriends calls.
function getEventFriendsData(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Event not found');

    var frData = ss.getSheetByName('Friends').getDataRange().getValues();
    var friendMap = {};
    var allFriends = [];
    for (var i = 1; i < frData.length; i++) {
      if (frData[i][1] === user.id) {
        friendMap[frData[i][0]] = frData[i][2];
        allFriends.push({ id: frData[i][0], name: frData[i][2] });
      }
    }

    var linkedFriends = _getEventFriends(ss, eventId, user.id, friendMap);
    return { success: true, allFriends: allFriends, linkedFriends: linkedFriends };
  } catch (e) {
    return _fail(e);
  }
}

function addFriendToEvent(token, eventId, name) {
  try {
    var user = requireAuth(token);
    if (!name || name.trim() === '') return _fail('Name is required');
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Event not found');
    var trimmed = name.trim();

    var frSheet = ss.getSheetByName('Friends');
    var frData = frSheet.getDataRange().getValues();
    var friendId = null;
    for (var i = 1; i < frData.length; i++) {
      if (frData[i][1] === user.id && frData[i][2].toLowerCase() === trimmed.toLowerCase()) {
        friendId = frData[i][0];
        break;
      }
    }
    if (!friendId) {
      friendId = Utilities.getUuid();
      frSheet.appendRow([friendId, user.id, trimmed]);
    }

    var efSheet = getEventFriendsSheet();
    var efData = efSheet.getDataRange().getValues();
    var alreadyLinked = false;
    for (var i = 1; i < efData.length; i++) {
      if (efData[i][1] === eventId && efData[i][2] === friendId) { alreadyLinked = true; break; }
    }
    if (!alreadyLinked) {
      efSheet.appendRow([Utilities.getUuid(), eventId, friendId, new Date().toISOString()]);
    }

    return { success: true, friend: { id: friendId, name: trimmed } };
  } catch (e) {
    return _fail(e);
  }
}

function setEventFriends(token, eventId, friendIds) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Event not found');

    // Read each sheet exactly once for this request.
    var frData = ss.getSheetByName('Friends').getDataRange().getValues();
    var ownedFriendMap = {};
    for (var i = 1; i < frData.length; i++) {
      if (frData[i][1] === user.id) ownedFriendMap[frData[i][0]] = frData[i][2];
    }
    var wantedIds = (friendIds || []).filter(function (fid) { return ownedFriendMap.hasOwnProperty(fid); });

    var dtData = ss.getSheetByName('Details').getDataRange().getValues();
    var usedInEvent = {};
    var derivedFromDetails = {};
    for (var i = 1; i < dtData.length; i++) {
      if (dtData[i][1] === eventId) {
        var rowPayId = dtData[i][2];
        usedInEvent[rowPayId] = true;
        if (ownedFriendMap.hasOwnProperty(rowPayId)) derivedFromDetails[rowPayId] = true;
        _splitsToRows(dtData[i]).forEach(function (s) {
          usedInEvent[s.friendId] = true;
          if (ownedFriendMap.hasOwnProperty(s.friendId)) derivedFromDetails[s.friendId] = true;
        });
      }
    }

    var efSheet = getEventFriendsSheet();
    var efData = efSheet.getDataRange().getValues();
    var hasAnyLink = false;
    var currentIdSet = {};
    for (var i = 1; i < efData.length; i++) {
      if (efData[i][1] === eventId) {
        hasAnyLink = true;
        if (ownedFriendMap.hasOwnProperty(efData[i][2])) currentIdSet[efData[i][2]] = true;
      }
    }
    // Same lazy migration as _getEventFriends, inlined to avoid re-reading Details/EventFriends.
    if (!hasAnyLink) {
      Object.keys(derivedFromDetails).forEach(function (fid) { currentIdSet[fid] = true; });
    }

    var blocked = [];
    var toAdd = wantedIds.filter(function (fid) { return !currentIdSet[fid]; });
    var toRemove = Object.keys(currentIdSet).filter(function (fid) {
      if (wantedIds.indexOf(fid) !== -1) return false;
      if (usedInEvent[fid]) { blocked.push({ id: fid, name: ownedFriendMap[fid] }); return false; }
      return true;
    });

    toRemove.forEach(function (fid) { delete currentIdSet[fid]; });

    var now = new Date().toISOString();
    var newRows = [];
    if (!hasAnyLink) {
      Object.keys(derivedFromDetails).forEach(function (fid) {
        if (toRemove.indexOf(fid) === -1) newRows.push([Utilities.getUuid(), eventId, fid, now]);
      });
    }
    toAdd.forEach(function (fid) {
      newRows.push([Utilities.getUuid(), eventId, fid, now]);
      currentIdSet[fid] = true;
    });

    // Single read + single write for the whole mutation, not one
    // deleteRow()/appendRow() per changed row.
    if (toRemove.length || newRows.length) {
      var keepRows = efData.filter(function (row, i) {
        return i > 0 && !(row[1] === eventId && toRemove.indexOf(row[2]) !== -1);
      });
      efSheet.clearContents();
      var allRows = [efData[0]].concat(keepRows).concat(newRows);
      efSheet.getRange(1, 1, allRows.length, 4).setValues(allRows);
    }

    var finalFriends = Object.keys(currentIdSet).map(function (fid) {
      return { id: fid, name: ownedFriendMap[fid] };
    });
    return { success: true, friends: finalFriends, blocked: blocked };
  } catch (e) {
    return _fail(e);
  }
}

// ----------------------------------------------------------------
// Events
// ----------------------------------------------------------------

function addEvent(token, name, icon) {
  try {
    var user = requireAuth(token);
    if (!name || name.trim() === '') return _fail('Event name is required');
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Events');
    var id = Utilities.getUuid();
    var now = new Date().toISOString();
    sheet.appendRow([id, name.trim(), user.id, now, '', icon || '']);

    // Auto-link the account's own self-friend so every event starts with yourself in it
    var frData = ss.getSheetByName('Friends').getDataRange().getValues();
    var selfRow = _findSelfFriendRow(frData, user.id);
    if (selfRow !== -1) {
      getEventFriendsSheet().appendRow([Utilities.getUuid(), id, frData[selfRow][0], now]);
    }

    return { success: true, event: { id: id, name: name.trim(), accountId: user.id, createdAt: now, icon: icon || '' } };
  } catch (e) {
    return _fail(e);
  }
}

function renameEvent(token, eventId, name, icon) {
  try {
    var user = requireAuth(token);
    if (!name || name.trim() === '') return _fail('Event name is required');
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Events');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId && data[i][2] === user.id) {
        sheet.getRange(i + 1, 2).setValue(name.trim());
        sheet.getRange(i + 1, 6).setValue(icon || '');
        return { success: true, name: name.trim(), icon: icon || '' };
      }
    }
    return _fail('Event not found');
  } catch (e) {
    return _fail(e);
  }
}

function setEventActive(token, eventId, active) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Events');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId && data[i][2] === user.id) {
        sheet.getRange(i + 1, 5).setValue(active === true);
        return { success: true, active: active === true };
      }
    }
    return _fail('Event not found');
  } catch (e) {
    return _fail(e);
  }
}

function deleteEvent(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    // Check ownership before deleting anything - don't wipe another
    // account's Details/EventFriends/EventShares rows first.
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Event not found');

    var detailsSheet = ss.getSheetByName('Details');
    var dtData = detailsSheet.getDataRange().getValues();
    var txIds = {};
    for (var i = 1; i < dtData.length; i++) {
      if (dtData[i][1] === eventId) txIds[dtData[i][0]] = true;
    }

    _trashSlipFilesForTxSet(txIds);
    _removeRowsWhere(detailsSheet, 1, eventId, dtData);
    _removeRowsWhere(getEventFriendsSheet(), 1, eventId);
    _removeRowsWhere(getEventSharesSheet(), 0, eventId);
    _removeRowsBySet(getTransactionSlipsSheet(), 0, txIds);

    var eventsSheet = ss.getSheetByName('Events');
    var eventsData = eventsSheet.getDataRange().getValues();
    for (var j = 1; j < eventsData.length; j++) {
      if (eventsData[j][0] === eventId && eventsData[j][2] === user.id) {
        eventsSheet.deleteRow(j + 1);
        return { success: true };
      }
    }
    return _fail('Event not found');
  } catch (e) {
    return _fail(e);
  }
}

// ----------------------------------------------------------------
// Event Sharing (public read-only link)
// ----------------------------------------------------------------

function getShareLink(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Event not found');
    var data = getEventSharesSheet().getDataRange().getValues();
    var row = _findRow(data, eventId);
    if (row === -1) return { success: true, shareToken: null };
    return {
      success: true, shareToken: data[row][1], permission: _sharePermission(data[row]),
      shareUrl: ScriptApp.getService().getUrl() + '?share=' + data[row][1]
    };
  } catch (e) {
    return _fail(e);
  }
}

// Creates the link on first call; on later calls with an existing link, just
// updates its permission in place so the same URL keeps working.
function enableEventShare(token, eventId, permission) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Event not found');
    var perm = permission === 'edit' ? 'edit' : 'view';

    var sheet = getEventSharesSheet();
    var data = sheet.getDataRange().getValues();
    var row = _findRow(data, eventId);
    var shareToken;
    if (row !== -1) {
      shareToken = data[row][1];
      sheet.getRange(row + 1, 4).setValue(perm);
    } else {
      shareToken = Utilities.getUuid();
      sheet.appendRow([eventId, shareToken, new Date().toISOString(), perm]);
    }
    return { success: true, shareToken: shareToken, permission: perm, shareUrl: ScriptApp.getService().getUrl() + '?share=' + shareToken };
  } catch (e) {
    return _fail(e);
  }
}

function disableEventShare(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Event not found');
    _removeRowsWhere(getEventSharesSheet(), 0, eventId);
    return { success: true };
  } catch (e) {
    return _fail(e);
  }
}

// Public — intentionally takes no auth token. Only ever returns the one
// event a valid, unguessable share token points to; never account data.
function getSharedEventView(shareToken) {
  try {
    if (!shareToken) return _fail('Invalid link');
    var ss = getSpreadsheet();

    var shData = getEventSharesSheet().getDataRange().getValues();
    var eventId = null, permission = 'view';
    for (var i = 1; i < shData.length; i++) {
      if (shData[i][1] === shareToken) { eventId = shData[i][0]; permission = _sharePermission(shData[i]); break; }
    }
    if (!eventId) return _fail('This share link is no longer active');

    var evData = ss.getSheetByName('Events').getDataRange().getValues();
    var eventRow = null;
    for (var i = 1; i < evData.length; i++) {
      if (evData[i][0] === eventId) { eventRow = evData[i]; break; }
    }
    if (!eventRow) return _fail('This share link is no longer active');
    var accountId = eventRow[2];

    var ownerPhoto = '';
    var acData = ss.getSheetByName('Accounts').getDataRange().getValues();
    for (var i = 1; i < acData.length; i++) {
      if (acData[i][0] === accountId) { ownerPhoto = acData[i][9] || ''; break; }
    }

    // Same core the authenticated getDetailData uses, so a share visitor sees
    // identical details/friends/settlements (including paid state)/slips
    // shapes — the client renders both through the exact same Detail code.
    var payload = _buildDetailPayload(ss, eventId, accountId);

    return {
      success: true,
      event: { name: eventRow[1], createdAt: eventRow[3], icon: eventRow[5] || '' },
      details: payload.details,
      friends: payload.friends,
      settlements: payload.settlements,
      selfFriendId: payload.selfFriendId,
      ownerPhoto: ownerPhoto,
      slips: payload.slips,
      permission: permission
    };
  } catch (e) {
    return _fail(e);
  }
}

// ----------------------------------------------------------------
// Details (Transactions)
// ----------------------------------------------------------------

// splits cell: JSON {friendId: amount}. Tolerant of corrupt/missing data so
// one bad cell degrades to "zero participants" instead of a thrown error.
function _parseSplits(json) {
  try {
    var o = JSON.parse(json || '{}');
    return (o && typeof o === 'object') ? o : {};
  } catch (e) {
    return {};
  }
}

// One Details row -> [{friendId, amount}, ...], the shape _computeSettlements expects.
function _splitsToRows(dtRow) {
  var splits = _parseSplits(dtRow[6]);
  return Object.keys(splits).map(function (fid) { return { friendId: fid, amount: splits[fid] }; });
}

// One Details row -> flat per-friend objects, matching the old one-row-per-split
// shape the client already expects (details[] items) - keeps the client
// contract unchanged even though storage is now one row per transaction.
function _expandDetailRow(dtRow) {
  var transactionId = dtRow[0], eventId = dtRow[1], payId = dtRow[2],
      totalAmount = dtRow[3], description = dtRow[4], createdAt = dtRow[5];
  return _splitsToRows(dtRow).map(function (s) {
    return {
      id: transactionId + '_' + s.friendId, eventId: eventId, transactionId: transactionId,
      payId: payId, friendId: s.friendId, amount: s.amount, totalAmount: totalAmount,
      description: description, createdAt: createdAt
    };
  });
}

// friendIds -> one Details row, splits collapsed into a single JSON cell.
function _buildDetailRow(eventId, transactionId, payId, friendIds, total, description, customAmounts, createdAt) {
  var perPerson = total / friendIds.length;
  var splits = {};
  friendIds.forEach(function (fid) {
    splits[fid] = (customAmounts && customAmounts[fid] !== undefined) ? parseFloat(customAmounts[fid]) : perPerson;
  });
  return [transactionId, eventId, payId, total, description, createdAt, JSON.stringify(splits)];
}

// Builds a row via _buildDetailRow and appends it - the common tail shared by
// add, both authenticated and share-link variants, below.
function _writeDetailRow(sheet, eventId, transactionId, payId, friendIds, totalAmount, description, customAmounts, createdAt) {
  sheet.appendRow(_buildDetailRow(eventId, transactionId, payId, friendIds, parseFloat(totalAmount), description, customAmounts, createdAt));
}

function addDetail(token, eventId, payId, friendIds, totalAmount, description, customAmounts) {
  try {
    var user = requireAuth(token);
    if (!friendIds || !friendIds.length) return _fail('At least one person is required');
    var sheet = getSpreadsheet().getSheetByName('Details');
    var transactionId = Utilities.getUuid();
    _writeDetailRow(sheet, eventId, transactionId, payId, friendIds, totalAmount, description, customAmounts, new Date().toISOString());
    return { success: true, transactionId: transactionId };
  } catch (e) {
    return _fail(e);
  }
}

function updateDetail(token, transactionId, payId, friendIds, totalAmount, description, customAmounts) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Details');
    var data = sheet.getDataRange().getValues();

    var row = _findRow(data, transactionId);
    if (row === -1) return _fail('Transaction not found');
    var eventId = data[row][1], createdAt = data[row][5];
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Transaction not found');
    if (!friendIds || !friendIds.length) return _fail('At least one person is required');

    var newRow = _buildDetailRow(eventId, transactionId, payId, friendIds, parseFloat(totalAmount), description, customAmounts, createdAt);
    sheet.getRange(row + 1, 1, 1, newRow.length).setValues([newRow]);
    return { success: true, transactionId: transactionId };
  } catch (e) {
    return _fail(e);
  }
}

function deleteDetail(token, transactionId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Details');
    var data = sheet.getDataRange().getValues();
    var row = _findRow(data, transactionId);
    if (row === -1) return _fail('Transaction not found');
    var eventId = data[row][1];
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Transaction not found');

    _trashSlipFilesForTx(transactionId);
    sheet.deleteRow(row + 1);
    _removeRowsWhere(getTransactionSlipsSheet(), 0, transactionId);
    return { success: true };
  } catch (e) {
    return _fail(e);
  }
}

// ----------------------------------------------------------------
// Details (Transactions) via an 'edit' share link - no account, so these
// check the share token's permission instead of requireAuth. Deliberately
// scoped to transaction CRUD only (no friends/event management), per the
// share-permission design.
// ----------------------------------------------------------------

function _shareEventId(shareToken, requireEdit) {
  if (!shareToken) return null;
  var data = getEventSharesSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === shareToken) {
      if (requireEdit && _sharePermission(data[i]) !== 'edit') return null;
      return data[i][0];
    }
  }
  return null;
}

function addDetailViaShare(shareToken, payId, friendIds, totalAmount, description, customAmounts) {
  try {
    var eventId = _shareEventId(shareToken, true);
    if (!eventId) return _fail('This share link cannot make changes');
    if (!friendIds || !friendIds.length) return _fail('At least one person is required');

    var sheet = getSpreadsheet().getSheetByName('Details');
    var transactionId = Utilities.getUuid();
    _writeDetailRow(sheet, eventId, transactionId, payId, friendIds, totalAmount, description, customAmounts, new Date().toISOString());
    return { success: true, transactionId: transactionId };
  } catch (e) {
    return _fail(e);
  }
}

function updateDetailViaShare(shareToken, transactionId, payId, friendIds, totalAmount, description, customAmounts) {
  try {
    var eventId = _shareEventId(shareToken, true);
    if (!eventId) return _fail('This share link cannot make changes');
    if (!friendIds || !friendIds.length) return _fail('At least one person is required');

    var sheet = getSpreadsheet().getSheetByName('Details');
    var data = sheet.getDataRange().getValues();
    var row = _findRow(data, transactionId);
    if (row === -1 || data[row][1] !== eventId) return _fail('Transaction not found');
    var createdAt = data[row][5];

    var newRow = _buildDetailRow(eventId, transactionId, payId, friendIds, parseFloat(totalAmount), description, customAmounts, createdAt);
    sheet.getRange(row + 1, 1, 1, newRow.length).setValues([newRow]);
    return { success: true, transactionId: transactionId };
  } catch (e) {
    return _fail(e);
  }
}

// deleteDetailViaShare / deleteTransactionSlipViaShare were removed on purpose
// (not merely hidden client-side): no share link, editable or not, may ever
// delete a transaction or a saved photo. Add/Edit stays available below.

// 'slip'/'slipHi' both hold the Drive fileId (client resolves it to a real
// image via getSlipImage/getSlipImageViaShare); 'fileId' is its own column
// too since that's what the delete-cleanup helpers key off of.
function _saveUploadedSlip(transactionId, slip) {
  var uploaded = _uploadSlipToDrive(slip);
  var id = Utilities.getUuid();
  getTransactionSlipsSheet().appendRow([transactionId, uploaded.fileId, new Date().toISOString(), id, uploaded.fileId, uploaded.fileId]);
  return { success: true, id: id, slip: uploaded.fileId, slipHi: uploaded.fileId };
}

function uploadTransactionSlipViaShare(shareToken, transactionId, slip) {
  try {
    var eventId = _shareEventId(shareToken, true);
    if (!eventId) return _fail('This share link cannot make changes');
    if (!slip) return _fail('No photo provided');
    var dtData = getSpreadsheet().getSheetByName('Details').getDataRange().getValues();
    if (_transactionEventId(dtData, transactionId) !== eventId) return _fail('Transaction not found');
    return _saveUploadedSlip(transactionId, slip);
  } catch (e) {
    return _fail(e);
  }
}

function _transactionEventId(dtData, transactionId) {
  var row = _findRow(dtData, transactionId);
  return row === -1 ? null : dtData[row][1];
}

// Always adds a new photo (a transaction can have several) - returns its id
// so the client can target it with deleteTransactionSlip later.
function uploadTransactionSlip(token, transactionId, slip) {
  try {
    var user = requireAuth(token);
    if (!slip) return _fail('No photo provided');
    var ss = getSpreadsheet();
    var dtData = ss.getSheetByName('Details').getDataRange().getValues();
    var eventId = _transactionEventId(dtData, transactionId);
    if (!eventId) return _fail('Transaction not found');
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Transaction not found');
    return _saveUploadedSlip(transactionId, slip);
  } catch (e) {
    return _fail(e);
  }
}

function deleteTransactionSlip(token, transactionId, slipId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var dtData = ss.getSheetByName('Details').getDataRange().getValues();
    var eventId = _transactionEventId(dtData, transactionId);
    if (!eventId) return _fail('Transaction not found');
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Transaction not found');

    var sheet = getTransactionSlipsSheet();
    var data = _backfillSlipIds(sheet, sheet.getDataRange().getValues());
    for (var j = 1; j < data.length; j++) {
      if (data[j][0] === transactionId && data[j][3] === slipId) {
        _deleteSlipFile(data[j][5]);
        sheet.deleteRow(j + 1);
        return { success: true };
      }
    }
    return _fail('Photo not found');
  } catch (e) {
    return _fail(e);
  }
}

// ----------------------------------------------------------------
// Summary / Settlement Calculation
// ----------------------------------------------------------------

// detailRows: array of {payId, friendId, amount}. friendMap: id -> name.
function _computeSettlements(detailRows, friendMap) {
  // debt[debtor][creditor] = amount debtor owes creditor
  var debt = {};
  detailRows.forEach(function (row) {
    var payerId = row.payId;
    var participantId = row.friendId;
    var amount = parseFloat(row.amount);
    if (participantId !== payerId) {
      if (!debt[participantId]) debt[participantId] = {};
      if (!debt[participantId][payerId]) debt[participantId][payerId] = 0;
      debt[participantId][payerId] += amount;
    }
  });

  // Net pairwise debts
  var netDebt = {};
  var processed = {};
  Object.keys(debt).forEach(function (debtor) {
    Object.keys(debt[debtor]).forEach(function (creditor) {
      var key1 = debtor + '_' + creditor;
      var key2 = creditor + '_' + debtor;
      if (processed[key1] || processed[key2]) return;
      processed[key1] = true;
      processed[key2] = true;

      var owes = debt[debtor][creditor] || 0;
      var oweBack = (debt[creditor] && debt[creditor][debtor]) ? debt[creditor][debtor] : 0;
      var net = owes - oweBack;

      if (Math.abs(net) < 0.01) return; // negligible

      if (net > 0) {
        if (!netDebt[debtor]) netDebt[debtor] = {};
        netDebt[debtor][creditor] = net;
      } else {
        if (!netDebt[creditor]) netDebt[creditor] = {};
        netDebt[creditor][debtor] = -net;
      }
    });
  });

  // Build settlements array
  var settlements = [];
  Object.keys(netDebt).forEach(function (from) {
    Object.keys(netDebt[from]).forEach(function (to) {
      settlements.push({
        from: from,
        fromName: friendMap[from] || from,
        to: to,
        toName: friendMap[to] || to,
        amount: Math.round(netDebt[from][to] * 100) / 100
      });
    });
  });
  return settlements;
}

// Payment confirmations are keyed on (from, to, amount) rather than stored on
// the settlement row itself, since settlements are recomputed fresh every
// time from Details — any change to the underlying debt (new/edited/deleted
// expense) naturally invalidates a stale confirmation because the amount
// no longer matches, with no separate cleanup step needed.
function _settleKey(from, to, amount) {
  return from + '|' + to + '|' + Math.round(parseFloat(amount) * 100);
}

// Pass spDataOpt when the caller already read the SettlementPayments sheet
// this request (e.g. a per-event loop) to avoid re-reading it for every event.
function _getPaidSet(eventId, spDataOpt) {
  var data = spDataOpt || getSettlementPaymentsSheet().getDataRange().getValues();
  var set = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === eventId) set[_settleKey(data[i][2], data[i][3], data[i][4])] = true;
  }
  return set;
}

function _computeSettlementsWithPaid(detailRows, friendMap, eventId, spDataOpt) {
  var settlements = _computeSettlements(detailRows, friendMap);
  var paidSet = _getPaidSet(eventId, spDataOpt);
  settlements.forEach(function (s) { s.paid = !!paidSet[_settleKey(s.from, s.to, s.amount)] });
  return settlements;
}

function markSettlementPaid(token, eventId, fromId, toId, amount, paid) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Event not found');
    var sheet = getSettlementPaymentsSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === eventId && data[i][2] === fromId && data[i][3] === toId) sheet.deleteRow(i + 1);
    }
    if (paid) sheet.appendRow([Utilities.getUuid(), eventId, fromId, toId, amount, new Date().toISOString()]);
    return { success: true };
  } catch (e) {
    return _fail(e);
  }
}

// Marks a single transaction as already settled - it's then excluded from
// settlement math app-wide (see the paidTxSet filters in getHomeData,
// _buildDetailPayload, getSummary) instead of just noting a net debt as paid.
function markTransactionPaid(token, eventId, transactionId, paid) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return _fail('Event not found');
    _removeRowsWhere(getTransactionPaymentsSheet(), 0, transactionId);
    if (paid) getTransactionPaymentsSheet().appendRow([transactionId, eventId, new Date().toISOString()]);
    return { success: true };
  } catch (e) {
    return _fail(e);
  }
}

function getSummary(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();

    var friendsData = ss.getSheetByName('Friends').getDataRange().getValues();
    var friendMap = {};
    for (var f = 1; f < friendsData.length; f++) {
      if (friendsData[f][1] === user.id) friendMap[friendsData[f][0]] = friendsData[f][2];
    }

    var detailsData = ss.getSheetByName('Details').getDataRange().getValues();
    var paidTxSet = _paidTransactionSet();
    var rows = [];
    for (var i = 1; i < detailsData.length; i++) {
      if (detailsData[i][1] === eventId && !paidTxSet[detailsData[i][0]]) {
        var payId = detailsData[i][2];
        _splitsToRows(detailsData[i]).forEach(function (s) { rows.push({ payId: payId, friendId: s.friendId, amount: s.amount }) });
      }
    }

    return { success: true, settlements: _computeSettlementsWithPaid(rows, friendMap, eventId) };
  } catch (e) {
    return _fail(e);
  }
}

// ----------------------------------------------------------------
// Admin
// ----------------------------------------------------------------

function getAllAccounts(token) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return _fail('Forbidden');
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Accounts');
    var data = sheet.getDataRange().getValues();
    var accounts = [];
    for (var i = 1; i < data.length; i++) {
      accounts.push({
        id: data[i][0],
        displayName: data[i][1],
        username: data[i][2],
        firstLogin: data[i][4],
        lastLogin: data[i][5],
        role: data[i][6],
        status: data[i][7] || 'active',
        photo: data[i][9] || ''
      });
    }
    return { success: true, accounts: accounts };
  } catch (e) {
    return _fail(e);
  }
}

function updateAccountStatus(token, accountId, status) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return _fail('Forbidden');
    if (user.id === accountId) return _fail('Cannot disable your own account');
    var sheet = getSpreadsheet().getSheetByName('Accounts');
    var row = _findRow(sheet.getDataRange().getValues(), accountId);
    if (row === -1) return _fail('Account not found');
    sheet.getRange(row + 1, 8).setValue(status);
    return { success: true };
  } catch (e) {
    return _fail(e);
  }
}

function updateAccountRole(token, accountId, role) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return _fail('Forbidden');
    if (user.id === accountId) return _fail('Cannot change your own role');
    var sheet = getSpreadsheet().getSheetByName('Accounts');
    var row = _findRow(sheet.getDataRange().getValues(), accountId);
    if (row === -1) return _fail('Account not found');
    sheet.getRange(row + 1, 7).setValue(role);
    return { success: true };
  } catch (e) {
    return _fail(e);
  }
}

function deleteAccount(token, accountId) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return _fail('Forbidden');
    if (user.id === accountId) return _fail('Cannot delete your own account');
    var sheet = getSpreadsheet().getSheetByName('Accounts');
    var row = _findRow(sheet.getDataRange().getValues(), accountId);
    if (row === -1) return _fail('Account not found');
    sheet.deleteRow(row + 1);
    return { success: true };
  } catch (e) {
    return _fail(e);
  }
}

// ----------------------------------------------------------------
// ONE-TIME MIGRATION: Details sheet, old shape (one row per friend-split) ->
// new shape (one row per transaction, splits as JSON). Not client-callable -
// run manually from the Apps Script editor, once, then verify before
// redeploying the rest of this file. Never deletes the old data - the
// original sheet is renamed aside, not overwritten.
// ----------------------------------------------------------------

function migrateDetailsToV2() {
  var ss = getSpreadsheet();
  var old = ss.getSheetByName('Details');
  var oldData = old.getDataRange().getValues();

  if (oldData[0][0] === 'transactionId') {
    Logger.log('Already migrated - Details header is already the new shape. Nothing to do.');
    return { success: true, alreadyMigrated: true };
  }

  // Old shape: id0, eventId1, transactionId2, payId3, friendId4, amount5, totalAmount6, description7, createdAt8
  var groups = {}; // transactionId -> { eventId, payId, totalAmount, description, createdAt, splits: {friendId: amount} }
  for (var i = 1; i < oldData.length; i++) {
    var r = oldData[i];
    var tid = r[2];
    if (!groups[tid]) {
      groups[tid] = { eventId: r[1], payId: r[3], totalAmount: r[6], description: r[7], createdAt: r[8], splits: {} };
    }
    groups[tid].splits[r[4]] = r[5];
  }

  var newRows = Object.keys(groups).map(function (tid) {
    var g = groups[tid];
    return [tid, g.eventId, g.payId, g.totalAmount, g.description, g.createdAt, JSON.stringify(g.splits)];
  });

  // Sanity check before touching anything: every old split row must be
  // accounted for in the regrouped splits.
  var totalSplits = 0;
  Object.keys(groups).forEach(function (tid) { totalSplits += Object.keys(groups[tid].splits).length });
  if (totalSplits !== oldData.length - 1) {
    throw new Error('Migration aborted: split count mismatch (old rows=' + (oldData.length - 1) + ', regrouped splits=' + totalSplits + '). Nothing was changed.');
  }

  var backupName = 'Details_v1_backup_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  old.setName(backupName);

  var fresh = ss.insertSheet('Details');
  fresh.appendRow(['transactionId', 'eventId', 'payId', 'totalAmount', 'description', 'createdAt', 'splits']);
  if (newRows.length) fresh.getRange(2, 1, newRows.length, 7).setValues(newRows);

  Logger.log('Migrated ' + newRows.length + ' transactions from ' + (oldData.length - 1) + ' split rows. Backup: ' + backupName);

  // Verify immediately so running this ONE function from the editor is enough
  // - no need to separately look up and pass the backup sheet name.
  var verify = verifyDetailsMigration(backupName);
  Logger.log(verify.problems.length ? 'MIGRATION HAS PROBLEMS - see above. Old data is untouched in ' + backupName + '.' : 'VERIFIED OK - safe to redeploy.');

  return { success: true, transactions: newRows.length, oldRows: oldData.length - 1, backupSheetName: backupName, verify: verify };
}

// Read-only cross-check: every transaction in the backup must appear exactly
// once in the new Details sheet with identical eventId/payId/totalAmount/
// description/createdAt/splits. Safe to re-run anytime. No arg = auto-picks
// the most recently created Details_v1_backup_* sheet.
function verifyDetailsMigration(backupSheetName) {
  var ss = getSpreadsheet();
  if (!backupSheetName) {
    var candidates = ss.getSheets().map(function (s) { return s.getName() }).filter(function (n) { return n.indexOf('Details_v1_backup_') === 0 }).sort();
    backupSheetName = candidates[candidates.length - 1];
    if (!backupSheetName) { Logger.log('No Details_v1_backup_* sheet found - run migrateDetailsToV2 first.'); return { success: false, error: 'No backup sheet found' }; }
  }
  var backup = ss.getSheetByName(backupSheetName);
  if (!backup) { Logger.log('Backup sheet not found: ' + backupSheetName); return { success: false, error: 'Backup sheet not found' }; }
  var fresh = ss.getSheetByName('Details');
  if (!fresh) { Logger.log('Details sheet not found'); return { success: false, error: 'Details sheet not found' }; }

  var oldData = backup.getDataRange().getValues();
  var groups = {};
  for (var i = 1; i < oldData.length; i++) {
    var r = oldData[i];
    var tid = r[2];
    if (!groups[tid]) groups[tid] = { eventId: r[1], payId: r[3], totalAmount: r[6], description: r[7], createdAt: r[8], splits: {} };
    groups[tid].splits[r[4]] = r[5];
  }

  var newData = fresh.getDataRange().getValues();
  var newMap = {};
  for (var i = 1; i < newData.length; i++) newMap[newData[i][0]] = newData[i];

  var problems = [];
  var numsMatch = function (a, b) { return Math.abs(parseFloat(a) - parseFloat(b)) < 1e-9 };

  Object.keys(groups).forEach(function (tid) {
    var g = groups[tid], row = newMap[tid];
    if (!row) { problems.push('missing transaction: ' + tid); return; }
    if (row[1] !== g.eventId) problems.push(tid + ': eventId mismatch');
    if (row[2] !== g.payId) problems.push(tid + ': payId mismatch');
    if (!numsMatch(row[3], g.totalAmount)) problems.push(tid + ': totalAmount mismatch');
    if (row[4] !== g.description) problems.push(tid + ': description mismatch');
    if (row[5] !== g.createdAt) problems.push(tid + ': createdAt mismatch');
    var newSplits = _parseSplits(row[6]);
    var oldKeys = Object.keys(g.splits), newKeys = Object.keys(newSplits);
    if (oldKeys.length !== newKeys.length) problems.push(tid + ': split participant count mismatch');
    oldKeys.forEach(function (fid) {
      if (!newSplits.hasOwnProperty(fid)) problems.push(tid + ': missing friend ' + fid + ' in splits');
      else if (!numsMatch(newSplits[fid], g.splits[fid])) problems.push(tid + ': amount mismatch for friend ' + fid);
    });
  });

  Object.keys(newMap).forEach(function (tid) {
    if (!groups[tid]) problems.push('extra transaction not present in backup: ' + tid);
  });

  Logger.log(problems.length ? ('PROBLEMS FOUND:\n' + problems.join('\n')) : ('OK - ' + Object.keys(groups).length + ' transactions verified, zero problems.'));
  return { success: true, problems: problems, transactionsChecked: Object.keys(groups).length };
}
