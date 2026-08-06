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
    accountsSheet.appendRow([adminId, 'Admin', 'admin', adminPassword, now, now, 'admin']);

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
  accountsSheet.appendRow(['id', 'displayName', 'username', 'password', 'firstLogin', 'lastLogin', 'role']);

  var friendsSheet = ss.insertSheet('Friends');
  friendsSheet.appendRow(['id', 'accountId', 'name']);

  var eventsSheet = ss.insertSheet('Events');
  eventsSheet.appendRow(['id', 'name', 'accountId', 'createdAt']);

  var detailsSheet = ss.insertSheet('Details');
  detailsSheet.appendRow(['id', 'eventId', 'transactionId', 'payId', 'friendId', 'amount', 'totalAmount', 'description', 'createdAt']);

  var eventFriendsSheet = ss.insertSheet('EventFriends');
  eventFriendsSheet.appendRow(['id', 'eventId', 'friendId', 'createdAt']);

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

function getSessionsSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Sessions');
  if (!sheet) {
    sheet = ss.insertSheet('Sessions');
    sheet.appendRow(['token', 'accountId', 'userInfo', 'createdAt', 'expiresAt']);
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
        return { row: i + 1, userInfo: JSON.parse(data[i][2]) };
      }
      sheet.deleteRow(i + 1);
      return null;
    }
  }
  return null;
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

function doGet() {
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
    sheet.appendRow([id, displayName, username.toLowerCase(), hashed, now, now, 'user']);

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
    var events = [], friends = [];
    for (var i = 1; i < evData.length; i++) {
      if (evData[i][2] === user.id)
        events.push({ id: evData[i][0], name: evData[i][1], accountId: evData[i][2], createdAt: evData[i][3] });
    }
    events.sort(function(a,b){ return b.createdAt > a.createdAt ? 1 : -1 });
    for (var i = 1; i < frData.length; i++) {
      if (frData[i][1] === user.id)
        friends.push({ id: frData[i][0], accountId: frData[i][1], name: frData[i][2] });
    }
    return { success: true, events: events, friends: friends };
  } catch (e) { return { success: false, error: e.toString() } }
}

function getDetailData(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var dtData = ss.getSheetByName('Details').getDataRange().getValues();
    var details = [];
    for (var i = 1; i < dtData.length; i++) {
      if (dtData[i][1] === eventId)
        details.push({ id: dtData[i][0], eventId: dtData[i][1], transactionId: dtData[i][2],
          payId: dtData[i][3], friendId: dtData[i][4], amount: dtData[i][5],
          totalAmount: dtData[i][6], description: dtData[i][7], createdAt: dtData[i][8] });
    }
    var friends = _getEventFriends(ss, eventId, user.id);
    return { success: true, details: details, friends: friends };
  } catch (e) { return { success: false, error: e.toString() } }
}

// ----------------------------------------------------------------
// Friends
// ----------------------------------------------------------------

function getFriends(token) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Friends');
    var data = sheet.getDataRange().getValues();
    var friends = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === user.id) {
        friends.push({ id: data[i][0], accountId: data[i][1], name: data[i][2] });
      }
    }
    return { success: true, friends: friends };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
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

// Friends currently linked to an event. Self-healing migration: the first time
// an event with existing transactions but no EventFriends rows yet is read,
// membership is derived from who already appears in its Details and persisted.
function _getEventFriends(ss, eventId, accountId) {
  var frData = ss.getSheetByName('Friends').getDataRange().getValues();
  var friendMap = {};
  for (var i = 1; i < frData.length; i++) {
    if (frData[i][1] === accountId) friendMap[frData[i][0]] = frData[i][2];
  }

  var efSheet = getEventFriendsSheet();
  var efData = efSheet.getDataRange().getValues();
  var hasAnyLink = false;
  var linkedIds = [];
  for (var i = 1; i < efData.length; i++) {
    if (efData[i][1] === eventId) {
      hasAnyLink = true;
      if (friendMap.hasOwnProperty(efData[i][2])) linkedIds.push(efData[i][2]);
    }
  }

  if (!hasAnyLink) {
    var dtData = ss.getSheetByName('Details').getDataRange().getValues();
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
      derivedIds.forEach(function (fid) {
        efSheet.appendRow([Utilities.getUuid(), eventId, fid, now]);
      });
      linkedIds = derivedIds;
    }
  }

  return linkedIds.map(function (fid) {
    return { id: fid, name: friendMap[fid] };
  });
}

function getEventFriends(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    if (!_eventOwnedBy(ss, eventId, user.id)) return { success: false, error: 'Event not found' };
    return { success: true, friends: _getEventFriends(ss, eventId, user.id) };
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

    var frData = ss.getSheetByName('Friends').getDataRange().getValues();
    var ownedFriendMap = {};
    for (var i = 1; i < frData.length; i++) {
      if (frData[i][1] === user.id) ownedFriendMap[frData[i][0]] = frData[i][2];
    }
    var wantedIds = (friendIds || []).filter(function (fid) { return ownedFriendMap.hasOwnProperty(fid); });

    var currentIds = _getEventFriends(ss, eventId, user.id).map(function (f) { return f.id; });

    var dtData = ss.getSheetByName('Details').getDataRange().getValues();
    var usedInEvent = {};
    for (var i = 1; i < dtData.length; i++) {
      if (dtData[i][1] === eventId) {
        usedInEvent[dtData[i][3]] = true;
        usedInEvent[dtData[i][4]] = true;
      }
    }

    var blocked = [];
    var toAdd = wantedIds.filter(function (fid) { return currentIds.indexOf(fid) === -1; });
    var toRemove = currentIds.filter(function (fid) {
      if (wantedIds.indexOf(fid) !== -1) return false;
      if (usedInEvent[fid]) { blocked.push({ id: fid, name: ownedFriendMap[fid] }); return false; }
      return true;
    });

    var efSheet = getEventFriendsSheet();
    if (toAdd.length) {
      var now = new Date().toISOString();
      toAdd.forEach(function (fid) { efSheet.appendRow([Utilities.getUuid(), eventId, fid, now]); });
    }
    if (toRemove.length) {
      var efData = efSheet.getDataRange().getValues();
      for (var i = efData.length - 1; i >= 1; i--) {
        if (efData[i][1] === eventId && toRemove.indexOf(efData[i][2]) !== -1) {
          efSheet.deleteRow(i + 1);
        }
      }
    }

    return { success: true, friends: _getEventFriends(ss, eventId, user.id), blocked: blocked };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ----------------------------------------------------------------
// Events
// ----------------------------------------------------------------

function getEvents(token) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Events');
    var data = sheet.getDataRange().getValues();
    var events = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === user.id) {
        events.push({ id: data[i][0], name: data[i][1], accountId: data[i][2], createdAt: data[i][3] });
      }
    }
    // Newest first
    events.sort(function(a, b) { return b.createdAt > a.createdAt ? 1 : -1; });
    return { success: true, events: events };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

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

    // Delete all details for this event
    var detailsSheet = ss.getSheetByName('Details');
    var detailsData = detailsSheet.getDataRange().getValues();
    for (var i = detailsData.length - 1; i >= 1; i--) {
      if (detailsData[i][1] === eventId) {
        detailsSheet.deleteRow(i + 1);
      }
    }

    // Delete event-friend links
    var efSheet = getEventFriendsSheet();
    var efData = efSheet.getDataRange().getValues();
    for (var i = efData.length - 1; i >= 1; i--) {
      if (efData[i][1] === eventId) {
        efSheet.deleteRow(i + 1);
      }
    }

    // Delete event
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
// Details (Transactions)
// ----------------------------------------------------------------

function getDetails(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Details');
    var data = sheet.getDataRange().getValues();
    var details = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === eventId) {
        details.push({
          id: data[i][0],
          eventId: data[i][1],
          transactionId: data[i][2],
          payId: data[i][3],
          friendId: data[i][4],
          amount: data[i][5],
          totalAmount: data[i][6],
          description: data[i][7],
          createdAt: data[i][8]
        });
      }
    }
    return { success: true, details: details };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function addDetail(token, eventId, payId, friendIds, totalAmount, description, customAmounts) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Details');
    var transactionId = Utilities.getUuid();
    var now = new Date().toISOString();
    var total = parseFloat(totalAmount);
    var perPerson = total / friendIds.length;

    for (var i = 0; i < friendIds.length; i++) {
      var fid = friendIds[i];
      var amount = (customAmounts && customAmounts[fid] !== undefined)
        ? parseFloat(customAmounts[fid])
        : perPerson;
      var id = Utilities.getUuid();
      sheet.appendRow([id, eventId, transactionId, payId, fid, amount, total, description, now]);
    }

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
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][2] === transactionId) {
        sheet.deleteRow(i + 1);
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ----------------------------------------------------------------
// Summary / Settlement Calculation
// ----------------------------------------------------------------

function getSummary(token, eventId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();

    // Get friends map for this user
    var friendsSheet = ss.getSheetByName('Friends');
    var friendsData = friendsSheet.getDataRange().getValues();
    var friendMap = {};
    for (var f = 1; f < friendsData.length; f++) {
      if (friendsData[f][1] === user.id) {
        friendMap[friendsData[f][0]] = friendsData[f][2];
      }
    }

    // Get all details for this event
    var detailsSheet = ss.getSheetByName('Details');
    var detailsData = detailsSheet.getDataRange().getValues();

    // debt[debtor][creditor] = amount debtor owes creditor
    var debt = {};

    for (var i = 1; i < detailsData.length; i++) {
      var row = detailsData[i];
      if (row[1] !== eventId) continue;

      var payerId = row[3];   // who paid
      var participantId = row[4]; // who shares this row
      var amount = parseFloat(row[5]); // per-person share

      // If the participant is NOT the payer, they owe the payer
      if (participantId !== payerId) {
        if (!debt[participantId]) debt[participantId] = {};
        if (!debt[participantId][payerId]) debt[participantId][payerId] = 0;
        debt[participantId][payerId] += amount;
      }
    }

    // Net pairwise debts
    var netDebt = {};
    var processed = {};
    Object.keys(debt).forEach(function(debtor) {
      Object.keys(debt[debtor]).forEach(function(creditor) {
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
    Object.keys(netDebt).forEach(function(from) {
      Object.keys(netDebt[from]).forEach(function(to) {
        settlements.push({
          from: from,
          fromName: friendMap[from] || from,
          to: to,
          toName: friendMap[to] || to,
          amount: Math.round(netDebt[from][to] * 100) / 100
        });
      });
    });

    return { success: true, settlements: settlements };
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
        role: data[i][6]
      });
    }
    return { success: true, accounts: accounts };
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
