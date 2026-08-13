// ============================================================
// ShareUp - Expense Splitting App (Google Apps Script)
// ============================================================

var SPREADSHEET_NAME = 'ShareUp_Database';
var CACHE_EXPIRY = 21600; // 6 hours
var SESSION_DAYS = 30;   // persistent session lifetime

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
    accountsSheet.appendRow([adminId, 'Admin', 'admin', adminPassword, now, now, 'admin', 'active', '']);

    // Create "Me" friend for admin
    var friendsSheet = ss.getSheetByName('Friends');
    friendsSheet.appendRow([Utilities.getUuid(), adminId, 'Me']);
  }

  return ss;
}

function initSheets(ss) {
  // Remove default sheet if needed
  var defaultSheet = ss.getSheetByName('Sheet1');

  var accountsSheet = ss.insertSheet('Accounts');
  accountsSheet.appendRow(['id', 'displayName', 'username', 'password', 'firstLogin', 'lastLogin', 'role', 'status', 'email']);

  var friendsSheet = ss.insertSheet('Friends');
  friendsSheet.appendRow(['id', 'accountId', 'name']);

  var eventsSheet = ss.insertSheet('Events');
  eventsSheet.appendRow(['id', 'name', 'accountId', 'createdAt']);

  var detailsSheet = ss.insertSheet('Details');
  detailsSheet.appendRow(['id', 'eventId', 'transactionId', 'payId', 'friendId', 'amount', 'totalAmount', 'description', 'createdAt']);

  var eventFriendsSheet = ss.insertSheet('EventFriends');
  eventFriendsSheet.appendRow(['id', 'eventId', 'friendId', 'createdAt']);

  var eventSharesSheet = ss.insertSheet('EventShares');
  eventSharesSheet.appendRow(['eventId', 'token', 'createdAt']);

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
    sheet.appendRow(['eventId', 'token', 'createdAt']);
  }
  return sheet;
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

function _lookupSession(token) {
  var sheet = getSessionsSheet();
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      if (now < new Date(data[i][4])) {
        var userInfo = JSON.parse(data[i][2]);
        if (_isAccountDisabled(userInfo.id)) { sheet.deleteRow(i + 1); return null }
        return { row: i + 1, userInfo: userInfo };
      }
      sheet.deleteRow(i + 1);
      return null;
    }
  }
  return null;
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

function doGet(e) {
  var rawToken = e && e.parameter && e.parameter.share;
  if (rawToken) {
    // Strict allowlist so this can be embedded directly into the page's inline script safely.
    var shareToken = /^[a-zA-Z0-9-]{10,100}$/.test(rawToken) ? rawToken : '';
    var tpl = HtmlService.createTemplateFromFile('SharedView');
    tpl.shareToken = shareToken;
    return tpl.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('ShareUp - Shared Event');
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('ShareUp - Expense Splitting');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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
        var expires = new Date(now.getTime() + SESSION_DAYS * 86400000);
        getCache().put('token_' + token, JSON.stringify(userInfo), CACHE_EXPIRY);
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

    var now = new Date().toISOString();
    var id = Utilities.getUuid();
    var hashed = hashPassword(password);
    sheet.appendRow([id, displayName, username.toLowerCase(), hashed, now, now, 'user', 'active', '']);

    // Create "Me" friend
    var friendsSheet = ss.getSheetByName('Friends');
    friendsSheet.appendRow([Utilities.getUuid(), id, 'Me']);

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
    getCache().put('token_' + token, JSON.stringify(found.userInfo), CACHE_EXPIRY);
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
  getCache().put('token_' + token, JSON.stringify(found.userInfo), CACHE_EXPIRY);
  return found.userInfo;
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
        events.push({ id: evData[i][0], name: evData[i][1], accountId: evData[i][2], createdAt: evData[i][3] });
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

function getDetailData(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
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
    var friends = _getEventFriends(ss, eventId, user.id);
    var friendMap = {};
    friends.forEach(function (f) { friendMap[f.id] = f.name });
    // Bundled here so opening the Summary tab or exporting a PDF right after
    // doesn't force a second round-trip that re-reads the same Details rows.
    var settlements = _computeSettlementsWithPaid(rows, friendMap, eventId);
    return { success: true, details: details, friends: friends, settlements: settlements };
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

function addEvent(token, name) {
  try {
    var user = requireAuth(token);
    if (!name || name.trim() === '') return { success: false, error: 'Event name is required' };
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Events');
    var id = Utilities.getUuid();
    var now = new Date().toISOString();
    sheet.appendRow([id, name.trim(), user.id, now]);

    // Auto-link the account's "Me" friend so every event starts with yourself in it
    var frData = ss.getSheetByName('Friends').getDataRange().getValues();
    for (var i = 1; i < frData.length; i++) {
      if (frData[i][1] === user.id && frData[i][2] === 'Me') {
        getEventFriendsSheet().appendRow([Utilities.getUuid(), id, frData[i][0], now]);
        break;
      }
    }

    return { success: true, event: { id: id, name: name.trim(), accountId: user.id, createdAt: now } };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function renameEvent(token, eventId, name) {
  try {
    var user = requireAuth(token);
    if (!name || name.trim() === '') return { success: false, error: 'Event name is required' };
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Events');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId && data[i][2] === user.id) {
        sheet.getRange(i + 1, 2).setValue(name.trim());
        return { success: true, name: name.trim() };
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

    _removeRowsWhere(ss.getSheetByName('Details'), 1, eventId);
    _removeRowsWhere(getEventFriendsSheet(), 1, eventId);
    _removeRowsWhere(getEventSharesSheet(), 0, eventId);

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
        return { success: true, shareToken: data[i][1], shareUrl: ScriptApp.getService().getUrl() + '?share=' + data[i][1] };
      }
    }
    return { success: true, shareToken: null };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function enableEventShare(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Event not found' };

    var sheet = getEventSharesSheet();
    var data = sheet.getDataRange().getValues();
    var shareToken = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) { shareToken = data[i][1]; break; }
    }
    if (!shareToken) {
      shareToken = Utilities.getUuid();
      sheet.appendRow([eventId, shareToken, new Date().toISOString()]);
    }
    return { success: true, shareToken: shareToken, shareUrl: ScriptApp.getService().getUrl() + '?share=' + shareToken };
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
    var eventId = null;
    for (var i = 1; i < shData.length; i++) {
      if (shData[i][1] === shareToken) { eventId = shData[i][0]; break; }
    }
    if (!eventId) return { success: false, error: 'This share link is no longer active' };

    var evData = ss.getSheetByName('Events').getDataRange().getValues();
    var eventRow = null;
    for (var i = 1; i < evData.length; i++) {
      if (evData[i][0] === eventId) { eventRow = evData[i]; break; }
    }
    if (!eventRow) return { success: false, error: 'This share link is no longer active' };
    var accountId = eventRow[2];

    var dtData = ss.getSheetByName('Details').getDataRange().getValues();
    var rawDetails = [];
    for (var i = 1; i < dtData.length; i++) {
      if (dtData[i][1] === eventId) {
        rawDetails.push({ transactionId: dtData[i][2], payId: dtData[i][3], friendId: dtData[i][4],
          amount: dtData[i][5], totalAmount: dtData[i][6], description: dtData[i][7], createdAt: dtData[i][8] });
      }
    }

    var friends = _getEventFriends(ss, eventId, accountId);
    var friendMap = {};
    friends.forEach(function (f) { friendMap[f.id] = f.name });

    var details = rawDetails.map(function (d) {
      return {
        transactionId: d.transactionId,
        payerName: friendMap[d.payId] || d.payId,
        friendName: friendMap[d.friendId] || d.friendId,
        amount: d.amount,
        totalAmount: d.totalAmount,
        description: d.description,
        createdAt: d.createdAt
      };
    });

    return {
      success: true,
      event: { name: eventRow[1], createdAt: eventRow[3] },
      details: details,
      settlements: _computeSettlements(rawDetails, friendMap)
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
    _removeRowsWhere(ss.getSheetByName('Details'), 2, transactionId);
    return { success: true };
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
        status: data[i][7] || 'active'
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
