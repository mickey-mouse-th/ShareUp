// ============================================================
// ShareUp - Expense Splitting App (Google Apps Script)
// ============================================================

var SPREADSHEET_NAME = 'ShareUp_Database';
var CACHE_EXPIRY = 21600; // 6 hours

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

// ----------------------------------------------------------------
// Entry Point
// ----------------------------------------------------------------

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('ShareUp - Expense Splitting');
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
      if (row[2] === username && row[3] === hashed) {
        // Update lastLogin
        sheet.getRange(i + 1, 6).setValue(new Date().toISOString());

        var token = generateToken();
        var userInfo = {
          id: row[0],
          displayName: row[1],
          username: row[2],
          role: row[6]
        };
        getCache().put('token_' + token, JSON.stringify(userInfo), CACHE_EXPIRY);
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
      if (data[i][2] === username) {
        return { success: false, error: 'Username already taken' };
      }
    }

    var now = new Date().toISOString();
    var id = Utilities.getUuid();
    var hashed = hashPassword(password);
    sheet.appendRow([id, displayName, username, hashed, now, now, 'user']);

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
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getSessionUser(token) {
  try {
    var cached = getCache().get('token_' + token);
    if (!cached) return { success: false, error: 'Session expired' };
    return { success: true, user: JSON.parse(cached) };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function requireAuth(token) {
  var cached = getCache().get('token_' + token);
  if (!cached) throw new Error('Unauthorized');
  return JSON.parse(cached);
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

function addFriend(token, name) {
  try {
    var user = requireAuth(token);
    if (!name || name.trim() === '') return { success: false, error: 'Name is required' };
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Friends');
    var id = Utilities.getUuid();
    sheet.appendRow([id, user.id, name.trim()]);
    return { success: true, friend: { id: id, accountId: user.id, name: name.trim() } };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteFriend(token, friendId) {
  try {
    var user = requireAuth(token);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Friends');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === friendId && data[i][1] === user.id) {
        if (data[i][2] === 'Me') return { success: false, error: 'Cannot delete "Me"' };
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Friend not found' };
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
    return { success: true, event: { id: id, name: name.trim(), accountId: user.id, createdAt: now } };
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
