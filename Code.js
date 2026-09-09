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
// Mirrors the light indigo palette baked into Shared_css.html's :root, so opening
// the theme picker for the first time shows accurate starting colors.
// buildThemeCss() only emits overrides for keys actually saved (see below), so
// nothing visually changes until an admin explicitly saves via the picker.
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
  detailsSheet.appendRow(['id', 'eventId', 'transactionId', 'payId', 'friendId', 'amount', 'totalAmount', 'description', 'createdAt']);

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

// One row per PHOTO (a transaction can have several), not a column on
// Details, since a transaction fans out into one Details row per
// participant - a column there would duplicate every image N times.
// Column F ('fileId') was added when slips moved to Drive storage (see
// _uploadSlipToDrive below) - rows written before that have a blank fileId
// and keep their original base64 data-URI in 'slip'/'slipHi', which still
// renders fine in an <img>, so no migration of old rows is needed.
function getTransactionSlipsSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('TransactionSlips');
  if (!sheet) {
    sheet = ss.insertSheet('TransactionSlips');
    sheet.appendRow(['transactionId', 'slip', 'updatedAt', 'id', 'slipHi', 'fileId']);
  }
  return sheet;
}

// 'id' was added after this sheet already shipped (single-slip-per-transaction
// model) - rows written before that have a blank id. Backfill it lazily,
// same self-healing style as _findSelfFriendRow, so old photos stay
// individually deletable once multi-photo support lands.
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

// dataUri: a "data:<mime>;base64,<data>" string from the client's canvas
// compression step. Shares the file "Anyone with the link - Viewer" so it can
// be embedded in an <img> for anonymous share visitors too (view-only file
// permissions, not edit) - matches this app's existing "unguessable link"
// sharing model rather than requiring a login just to view a receipt photo.
function _uploadSlipToDrive(dataUri) {
  var m = /^data:([^;]+);base64,(.*)$/.exec(dataUri || '');
  if (!m) throw new Error('Invalid image data');
  var mimeType = m[1], base64 = m[2];
  if (base64.length > 2000000) throw new Error('Photo is too large - please try a smaller one');
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mimeType, 'slip-' + Utilities.getUuid());
  var file = _getSlipsFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var fileId = file.getId();
  return {
    fileId: fileId,
    // Drive's documented thumbnail endpoint - unlike the unofficial
    // lh3.googleusercontent.com/d/<id> trick, this one reliably works in a
    // plain <img src> for anyone with view access to the file.
    previewUrl: 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000',
    // Forces a real file download (original quality) for the "Download Original" button.
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + fileId
  };
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
  var now = new Date();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      if (now < new Date(data[i][4])) {
        var userInfo = JSON.parse(data[i][2]);
        if (_isAccountDisabled(userInfo.id)) { sheet.deleteRow(i + 1); return null }
        return { row: i + 1, userInfo: userInfo, expiresAt: data[i][4] };
      }
      sheet.deleteRow(i + 1);
      return null;
    }
  }
  return null;
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
  var tpl = HtmlService.createTemplateFromFile('Index');
  tpl.shareToken = shareToken;
  tpl.sharePermission = shareToken ? _sharePermissionByToken(shareToken) : '';
  return tpl.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle(shareToken ? 'ShareUp - Shared Event' : 'ShareUp - Expense Splitting');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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
    if (user.role !== 'admin') return { success: false, error: 'Forbidden' };
    return { success: true, settings: getAppSettings() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateSettings(token, settings) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return { success: false, error: 'Forbidden' };
    var minutes = parseInt(settings.sessionMinutes, 10);
    // Floor at 5 minutes - anything shorter makes it easy to accidentally lock
    // yourself (or every user) out via a mistyped value (e.g. minutes vs hours).
    if (!minutes || minutes < 5) return { success: false, error: 'Session length must be at least 5 minutes' };
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
    return { success: false, error: e.toString() };
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
    if (user.role !== 'admin') return { success: false, error: 'Forbidden' };
    var saved = _rawTheme();
    var merged = {};
    for (var k in DEFAULT_THEME) merged[k] = saved[k];
    for (var key in DEFAULT_THEME) {
      if (theme && HEX_COLOR_RE.test(theme[key])) merged[key] = theme[key];
    }
    PropertiesService.getScriptProperties().setProperty('APP_THEME', JSON.stringify(merged));
    return { success: true, theme: getAppTheme() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Builds a :root override CSS string from only the keys an admin has actually
// saved (raw property, NOT the DEFAULT_THEME-merged view) - included right
// after Shared_css.html in Index.html's <head> so it wins the cascade with
// zero client round-trip / zero flash-of-unstyled-color.
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
        if (row[7] === 'disabled') return { success: false, error: 'This account has been disabled' };
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
        return { success: true, token: token, user: userInfo };
      }
    }
    return { success: false, error: 'Invalid username or password' };
  } catch (e) {
    return { success: false, error: e.toString() };
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
        return { success: false, error: 'Username already taken' };
      }
    }

    var pwErr = _validatePassword(password);
    if (pwErr) return { success: false, error: pwErr };

    var now = new Date().toISOString();
    var id = Utilities.getUuid();
    var hashed = hashPassword(password);
    sheet.appendRow([id, displayName, username.toLowerCase(), hashed, now, now, 'user', 'active', '', '']);

    // Create "Me" friend
    var friendsSheet = ss.getSheetByName('Friends');
    friendsSheet.appendRow([Utilities.getUuid(), id, 'Me', 'true']);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function logoutUser(token) {
  try {
    getCache().remove('token_' + token);
    var sheet = getSessionsSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === token) { sheet.deleteRow(i + 1); break; }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getSessionUser(token) {
  try {
    var cached = getCache().get('token_' + token);
    if (cached) return { success: true, user: JSON.parse(cached) };
    var found = _lookupSession(token);
    if (!found) return { success: false, error: 'Session expired' };
    getCache().put('token_' + token, JSON.stringify(found.userInfo), _cacheTtlFor(found.expiresAt));
    return { success: true, user: found.userInfo };
  } catch (e) {
    return { success: false, error: e.toString() };
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
  var ttl = CACHE_EXPIRY;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      sheet.getRange(i + 1, 3).setValue(JSON.stringify(userInfo));
      ttl = _cacheTtlFor(data[i][4]);
      break;
    }
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
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === user.id) {
        return { success: true, displayName: data[i][1], username: data[i][2], photo: data[i][9] || '' };
      }
    }
    return { success: false, error: 'Account not found' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// photo: pass a string to set it ('' clears it); omit/null to leave unchanged.
function updateProfile(token, displayName, photo) {
  try {
    var user = requireAuth(token);
    if (!displayName || !displayName.trim()) return { success: false, error: 'Name is required' };
    var trimmed = displayName.trim();
    var ss = getSpreadsheet();

    var acSheet = ss.getSheetByName('Accounts');
    var acData = acSheet.getDataRange().getValues();
    var acRow = -1;
    for (var i = 1; i < acData.length; i++) {
      if (acData[i][0] === user.id) { acRow = i; break; }
    }
    if (acRow === -1) return { success: false, error: 'Account not found' };
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
    return { success: false, error: e.toString() };
  }
}

function changePassword(token, currentPassword, newPassword) {
  try {
    var user = requireAuth(token);
    var pwErr = _validatePassword(newPassword);
    if (pwErr) return { success: false, error: pwErr };
    var sheet = getSpreadsheet().getSheetByName('Accounts');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === user.id) {
        if (data[i][3] !== hashPassword(currentPassword || '')) return { success: false, error: 'Current password is incorrect' };
        sheet.getRange(i + 1, 4).setValue(hashPassword(newPassword));
        return { success: true };
      }
    }
    return { success: false, error: 'Account not found' };
  } catch (e) {
    return { success: false, error: e.toString() };
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

    // Settlement state per event, for the Home filter tabs — reuses the same
    // settlement engine as getSummary instead of a separate status field.
    // Every sheet this needs is read ONCE here (not per event) and shared
    // across the loop below via the *Opt params on the helpers.
    var rowsByEvent = {};
    events.forEach(function (ev) { rowsByEvent[ev.id] = [] });
    var dtData = ss.getSheetByName('Details').getDataRange().getValues();
    for (var i = 1; i < dtData.length; i++) {
      if (rowsByEvent.hasOwnProperty(dtData[i][1]))
        rowsByEvent[dtData[i][1]].push({ payId: dtData[i][3], friendId: dtData[i][4], amount: dtData[i][5] });
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
  } catch (e) { return { success: false, error: e.toString() } }
}

// Shared core behind getDetailData (authenticated) and getSharedEventView
// (public share link) so both paths compute details/friends/settlements/slips
// identically instead of maintaining two parallel implementations.
function _buildDetailPayload(ss, eventId, accountId) {
  var dtData = ss.getSheetByName('Details').getDataRange().getValues();
  var details = [], rows = [];
  for (var i = 1; i < dtData.length; i++) {
    if (dtData[i][1] === eventId) {
      details.push({ id: dtData[i][0], eventId: dtData[i][1], transactionId: dtData[i][2],
        payId: dtData[i][3], friendId: dtData[i][4], amount: dtData[i][5],
        totalAmount: dtData[i][6], description: dtData[i][7], createdAt: dtData[i][8] });
      rows.push({ payId: dtData[i][3], friendId: dtData[i][4], amount: dtData[i][5] });
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
  // Bundled here so opening the Summary tab or exporting a PDF right after
  // doesn't force a second round-trip that re-reads the same Details rows.
  // _computeSettlementsWithPaid (not the plain variant) so a share-link
  // visitor sees the exact same "paid" checkmarks the owner does.
  var settlements = _computeSettlementsWithPaid(rows, friendMap, eventId);

  // Only this event's transactionIds - TransactionSlips isn't keyed by
  // eventId, so without this filter every event's photos would be shipped
  // down on every load (slow, and a lot of wasted bandwidth as photos add up).
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

  // selfFriendId lets the client show the account's own profile photo (set
  // via My Profile) for its own avatar instead of the initials circle.
  return { details: details, friends: friends, settlements: settlements, selfFriendId: selfFriendId, slips: slips };
}

function getDetailData(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Event not found' };
    var payload = _buildDetailPayload(ss, eventId, user.id);
    payload.success = true;
    return payload;
  } catch (e) { return { success: false, error: e.toString() } }
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

// Finds the row index (in an already-read Friends data array) of the
// account's own self-friend — the one auto-created at registration and
// auto-linked into every new event. Prefers the isSelf marker; falls back
// to the pre-migration convention (named literally "Me") for rows created
// before that column existed. Returns -1 if none found.
function _findSelfFriendRow(frData, accountId) {
  for (var i = 1; i < frData.length; i++) {
    if (frData[i][1] === accountId && frData[i][3] === 'true') return i;
  }
  for (var i = 1; i < frData.length; i++) {
    if (frData[i][1] === accountId && frData[i][2] === 'Me') return i;
  }
  return -1;
}

// Removes every row whose column `col` (0-indexed) equals `val` in a single
// read + single write, instead of one deleteRow() API call per matching row
// — matters most for sheets that can accumulate many rows per event
// (Details, EventFriends). Pass dataOpt when the caller already read this
// sheet this request.
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

// Like _removeRowsWhere but matches against a set of ids (object used as a
// hash set) - for deleteEvent, which needs to drop TransactionSlips rows for
// every transactionId under the event, not a single value.
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

// Friends currently linked to an event. Self-healing migration: the first time
// an event with existing transactions but no EventFriends rows yet is read,
// membership is derived from who already appears in its Details and persisted.
// Pass friendMapOpt when the caller already read the Friends sheet this request.
// Pass efDataOpt/dtDataOpt when the caller already read those sheets this
// request (e.g. a per-event loop) to avoid re-reading them for every event.
function _getEventFriends(ss, eventId, accountId, friendMapOpt, efDataOpt, dtDataOpt) {
  var friendMap = friendMapOpt;
  if (!friendMap) {
    friendMap = {};
    var frData = ss.getSheetByName('Friends').getDataRange().getValues();
    for (var i = 1; i < frData.length; i++) {
      if (frData[i][1] === accountId) friendMap[frData[i][0]] = frData[i][2];
    }
  }

  var efSheet = getEventFriendsSheet();
  var efData = efDataOpt || efSheet.getDataRange().getValues();
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
        if (friendMap.hasOwnProperty(dtData[i][3])) derived[dtData[i][3]] = true;
        if (friendMap.hasOwnProperty(dtData[i][4])) derived[dtData[i][4]] = true;
      }
    }
    var derivedIds = Object.keys(derived);
    if (derivedIds.length) {
      var now = new Date().toISOString();
      var newRows = derivedIds.map(function (fid) { return [Utilities.getUuid(), eventId, fid, now]; });
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
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Event not found' };

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
    return { success: false, error: e.toString() };
  }
}

function addFriendToEvent(token, eventId, name) {
  try {
    var user = requireAuth(token);
    if (!name || name.trim() === '') return { success: false, error: 'Name is required' };
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Event not found' };
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
    return { success: false, error: e.toString() };
  }
}

function setEventFriends(token, eventId, friendIds) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Event not found' };

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
        usedInEvent[dtData[i][3]] = true;
        usedInEvent[dtData[i][4]] = true;
        if (ownedFriendMap.hasOwnProperty(dtData[i][3])) derivedFromDetails[dtData[i][3]] = true;
        if (ownedFriendMap.hasOwnProperty(dtData[i][4])) derivedFromDetails[dtData[i][4]] = true;
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

    // Single read (efData, above) + single write for the whole mutation,
    // instead of one deleteRow()/appendRow() call per changed row. Rows for
    // this event that aren't being removed are carried over byte-for-byte
    // (same id/createdAt) — only genuinely new rows get fresh ones.
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
    return { success: false, error: e.toString() };
  }
}

// ----------------------------------------------------------------
// Events
// ----------------------------------------------------------------

function addEvent(token, name, icon) {
  try {
    var user = requireAuth(token);
    if (!name || name.trim() === '') return { success: false, error: 'Event name is required' };
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
    return { success: false, error: e.toString() };
  }
}

function renameEvent(token, eventId, name, icon) {
  try {
    var user = requireAuth(token);
    if (!name || name.trim() === '') return { success: false, error: 'Event name is required' };
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
    return { success: false, error: 'Event not found' };
  } catch (e) {
    return { success: false, error: e.toString() };
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
    return { success: false, error: 'Event not found' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteEvent(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    // Ownership check moved before any deletion (it used to run only against
    // the final Events-row lookup below, after other users' rows in
    // Details/EventFriends/EventShares had already been wiped for a
    // not-yours eventId).
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Event not found' };

    var detailsSheet = ss.getSheetByName('Details');
    var dtData = detailsSheet.getDataRange().getValues();
    var txIds = {};
    for (var i = 1; i < dtData.length; i++) {
      if (dtData[i][1] === eventId) txIds[dtData[i][2]] = true;
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
    return { success: false, error: 'Event not found' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ----------------------------------------------------------------
// Event Sharing (public read-only link)
// ----------------------------------------------------------------

function getShareLink(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Event not found' };
    var data = getEventSharesSheet().getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) {
        return {
          success: true, shareToken: data[i][1], permission: _sharePermission(data[i]),
          shareUrl: ScriptApp.getService().getUrl() + '?share=' + data[i][1]
        };
      }
    }
    return { success: true, shareToken: null };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Creates the link on first call; on later calls with an existing link, just
// updates its permission in place so the same URL keeps working.
function enableEventShare(token, eventId, permission) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Event not found' };
    var perm = permission === 'edit' ? 'edit' : 'view';

    var sheet = getEventSharesSheet();
    var data = sheet.getDataRange().getValues();
    var shareToken = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) {
        shareToken = data[i][1];
        sheet.getRange(i + 1, 4).setValue(perm);
        break;
      }
    }
    if (!shareToken) {
      shareToken = Utilities.getUuid();
      sheet.appendRow([eventId, shareToken, new Date().toISOString(), perm]);
    }
    return { success: true, shareToken: shareToken, permission: perm, shareUrl: ScriptApp.getService().getUrl() + '?share=' + shareToken };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function disableEventShare(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Event not found' };
    var sheet = getEventSharesSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === eventId) sheet.deleteRow(i + 1);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Public — intentionally takes no auth token. Only ever returns the one
// event a valid, unguessable share token points to; never account data.
function getSharedEventView(shareToken) {
  try {
    if (!shareToken) return { success: false, error: 'Invalid link' };
    var ss = getSpreadsheet();

    var shData = getEventSharesSheet().getDataRange().getValues();
    var eventId = null, permission = 'view';
    for (var i = 1; i < shData.length; i++) {
      if (shData[i][1] === shareToken) { eventId = shData[i][0]; permission = _sharePermission(shData[i]); break; }
    }
    if (!eventId) return { success: false, error: 'This share link is no longer active' };

    var evData = ss.getSheetByName('Events').getDataRange().getValues();
    var eventRow = null;
    for (var i = 1; i < evData.length; i++) {
      if (evData[i][0] === eventId) { eventRow = evData[i]; break; }
    }
    if (!eventRow) return { success: false, error: 'This share link is no longer active' };
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
    return { success: false, error: e.toString() };
  }
}

// ----------------------------------------------------------------
// Details (Transactions)
// ----------------------------------------------------------------

// friendIds -> one Details row each, sharing totalAmount/description/createdAt.
function _buildDetailRows(eventId, transactionId, payId, friendIds, total, description, customAmounts, createdAt) {
  var perPerson = total / friendIds.length;
  return friendIds.map(function (fid) {
    var amount = (customAmounts && customAmounts[fid] !== undefined)
      ? parseFloat(customAmounts[fid])
      : perPerson;
    return [Utilities.getUuid(), eventId, transactionId, payId, fid, amount, total, description, createdAt];
  });
}

function addDetail(token, eventId, payId, friendIds, totalAmount, description, customAmounts) {
  try {
    var user = requireAuth(token);
    if (!friendIds || !friendIds.length) return { success: false, error: 'At least one person is required' };
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Details');
    var transactionId = Utilities.getUuid();
    var total = parseFloat(totalAmount);
    var rows = _buildDetailRows(eventId, transactionId, payId, friendIds, total, description, customAmounts, new Date().toISOString());
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    return { success: true, transactionId: transactionId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateDetail(token, transactionId, payId, friendIds, totalAmount, description, customAmounts) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Details');
    var data = sheet.getDataRange().getValues();

    var eventId = null, createdAt = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === transactionId) { eventId = data[i][1]; createdAt = data[i][8]; break; }
    }
    if (!eventId) return { success: false, error: 'Transaction not found' };
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Transaction not found' };
    if (!friendIds || !friendIds.length) return { success: false, error: 'At least one person is required' };

    _removeRowsWhere(sheet, 2, transactionId, data);

    var total = parseFloat(totalAmount);
    var rows = _buildDetailRows(eventId, transactionId, payId, friendIds, total, description, customAmounts, createdAt);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    return { success: true, transactionId: transactionId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteDetail(token, transactionId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Details');
    var data = sheet.getDataRange().getValues();
    var eventId = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === transactionId) { eventId = data[i][1]; break; }
    }
    if (!eventId) return { success: false, error: 'Transaction not found' };
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Transaction not found' };

    _trashSlipFilesForTx(transactionId);
    _removeRowsWhere(sheet, 2, transactionId, data);
    _removeRowsWhere(getTransactionSlipsSheet(), 0, transactionId);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
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
    if (!eventId) return { success: false, error: 'This share link cannot make changes' };
    if (!friendIds || !friendIds.length) return { success: false, error: 'At least one person is required' };

    var sheet = getSpreadsheet().getSheetByName('Details');
    var transactionId = Utilities.getUuid();
    var total = parseFloat(totalAmount);
    var rows = _buildDetailRows(eventId, transactionId, payId, friendIds, total, description, customAmounts, new Date().toISOString());
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { success: true, transactionId: transactionId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateDetailViaShare(shareToken, transactionId, payId, friendIds, totalAmount, description, customAmounts) {
  try {
    var eventId = _shareEventId(shareToken, true);
    if (!eventId) return { success: false, error: 'This share link cannot make changes' };
    if (!friendIds || !friendIds.length) return { success: false, error: 'At least one person is required' };

    var sheet = getSpreadsheet().getSheetByName('Details');
    var data = sheet.getDataRange().getValues();
    var createdAt = null, txEventId = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === transactionId) { txEventId = data[i][1]; createdAt = data[i][8]; break; }
    }
    if (txEventId !== eventId) return { success: false, error: 'Transaction not found' };

    _removeRowsWhere(sheet, 2, transactionId, data);
    var total = parseFloat(totalAmount);
    var rows = _buildDetailRows(eventId, transactionId, payId, friendIds, total, description, customAmounts, createdAt);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { success: true, transactionId: transactionId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// deleteDetailViaShare / deleteTransactionSlipViaShare were removed on purpose
// (not merely hidden client-side): no share link, editable or not, may ever
// delete a transaction or a saved photo. Add/Edit stays available below.

function uploadTransactionSlipViaShare(shareToken, transactionId, slip) {
  try {
    var eventId = _shareEventId(shareToken, true);
    if (!eventId) return { success: false, error: 'This share link cannot make changes' };
    if (!slip) return { success: false, error: 'No photo provided' };

    var dtData = getSpreadsheet().getSheetByName('Details').getDataRange().getValues();
    if (_transactionEventId(dtData, transactionId) !== eventId) return { success: false, error: 'Transaction not found' };

    var uploaded = _uploadSlipToDrive(slip);
    var id = Utilities.getUuid();
    getTransactionSlipsSheet().appendRow([transactionId, uploaded.previewUrl, new Date().toISOString(), id, uploaded.downloadUrl, uploaded.fileId]);
    return { success: true, id: id, slip: uploaded.previewUrl, slipHi: uploaded.downloadUrl };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// slip: pass a data-URI string to set it, or '' to remove it.
function _transactionEventId(dtData, transactionId) {
  for (var i = 1; i < dtData.length; i++) {
    if (dtData[i][2] === transactionId) return dtData[i][1];
  }
  return null;
}

// Always adds a new photo (a transaction can have several) - returns its id
// so the client can target it with deleteTransactionSlip later. The photo is
// stored as a Drive file (see _uploadSlipToDrive); the returned slip/slipHi
// are the preview/download URLs, not the raw upload the client sent.
function uploadTransactionSlip(token, transactionId, slip) {
  try {
    var user = requireAuth(token);
    if (!slip) return { success: false, error: 'No photo provided' };
    var ss = getSpreadsheet();
    var dtData = ss.getSheetByName('Details').getDataRange().getValues();
    var eventId = _transactionEventId(dtData, transactionId);
    if (!eventId) return { success: false, error: 'Transaction not found' };
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Transaction not found' };

    var uploaded = _uploadSlipToDrive(slip);
    var id = Utilities.getUuid();
    getTransactionSlipsSheet().appendRow([transactionId, uploaded.previewUrl, new Date().toISOString(), id, uploaded.downloadUrl, uploaded.fileId]);
    return { success: true, id: id, slip: uploaded.previewUrl, slipHi: uploaded.downloadUrl };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteTransactionSlip(token, transactionId, slipId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var dtData = ss.getSheetByName('Details').getDataRange().getValues();
    var eventId = _transactionEventId(dtData, transactionId);
    if (!eventId) return { success: false, error: 'Transaction not found' };
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Transaction not found' };

    var sheet = getTransactionSlipsSheet();
    var data = _backfillSlipIds(sheet, sheet.getDataRange().getValues());
    for (var j = 1; j < data.length; j++) {
      if (data[j][0] === transactionId && data[j][3] === slipId) {
        _deleteSlipFile(data[j][5]);
        sheet.deleteRow(j + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Photo not found' };
  } catch (e) {
    return { success: false, error: e.toString() };
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
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Event not found' };
    var sheet = getSettlementPaymentsSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === eventId && data[i][2] === fromId && data[i][3] === toId) sheet.deleteRow(i + 1);
    }
    if (paid) sheet.appendRow([Utilities.getUuid(), eventId, fromId, toId, amount, new Date().toISOString()]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
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
    var rows = [];
    for (var i = 1; i < detailsData.length; i++) {
      if (detailsData[i][1] === eventId) {
        rows.push({ payId: detailsData[i][3], friendId: detailsData[i][4], amount: detailsData[i][5] });
      }
    }

    return { success: true, settlements: _computeSettlementsWithPaid(rows, friendMap, eventId) };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ----------------------------------------------------------------
// Admin
// ----------------------------------------------------------------

function getAllAccounts(token) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return { success: false, error: 'Forbidden' };
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
    return { success: false, error: e.toString() };
  }
}

function updateAccountStatus(token, accountId, status) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return { success: false, error: 'Forbidden' };
    if (user.id === accountId) return { success: false, error: 'Cannot disable your own account' };
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Accounts');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === accountId) {
        sheet.getRange(i + 1, 8).setValue(status);
        return { success: true };
      }
    }
    return { success: false, error: 'Account not found' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateAccountRole(token, accountId, role) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return { success: false, error: 'Forbidden' };
    if (user.id === accountId) return { success: false, error: 'Cannot change your own role' };
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Accounts');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === accountId) {
        sheet.getRange(i + 1, 7).setValue(role);
        return { success: true };
      }
    }
    return { success: false, error: 'Account not found' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteAccount(token, accountId) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return { success: false, error: 'Forbidden' };
    if (user.id === accountId) return { success: false, error: 'Cannot delete your own account' };
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Accounts');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === accountId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Account not found' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
