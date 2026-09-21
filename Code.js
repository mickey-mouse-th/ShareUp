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

// Only used by the one-time migration script below (migrateRestToDb) to read
// the legacy Events/Details/EventFriends/... sheets - the live app no longer
// reads or writes this spreadsheet at all once that migration has run.
function getSpreadsheet() {
  var ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!ssId) throw new Error('SPREADSHEET_ID not set - nothing to migrate');
  return SpreadsheetApp.openById(ssId);
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

// ----------------------------------------------------------------
// Postgres (Neon) - the entire app lives here now (account, friend, event,
// transaction, split, participant, settlement, receipt, session). See
// migrateAccountsFriendsToDb() and migrateRestToDb() near the bottom of this
// file for the one-time cutovers that populated this from the old sheets.
//
// role/status stay numeric ONLY inside Postgres. Every other part of the
// app - every `user.role !== 'admin'` check in this file, plus the client
// JS - still expects the original 'admin'/'user' and 'active'/'disabled'
// strings, so that translation happens right here at the DB boundary and
// nowhere else in the app has to change.
// ----------------------------------------------------------------

function _dbConn() {
  var props = PropertiesService.getScriptProperties();
  return Jdbc.getConnection(props.getProperty('DB_URL'), props.getProperty('DB_USER'), props.getProperty('DB_PASS'));
}

function _roleToDb(role) { return role === 'admin' ? 2 : 0; } // staff(1) unused externally for now
function _roleFromDb(role) { return role >= 2 ? 'admin' : 'user'; }
function _statusToDb(status) { return status === 'disabled' ? 1 : 3; }
function _statusFromDb(status) { return status === 3 ? 'active' : 'disabled'; } // deleted(0)/pending(2) collapse to 'disabled' until those states have real behavior

function _accountFromRs(rs) {
  return {
    id: String(rs.getLong('id')), displayName: rs.getString('name'), username: rs.getString('username'),
    passwordHash: rs.getString('password_hash'), role: _roleFromDb(rs.getInt('role')), status: _statusFromDb(rs.getInt('status')),
    email: rs.getString('email') || '', photo: rs.getString('photo') || ''
  };
}

function _dbGetAccountByUsername(username, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT id, name, username, password_hash, role, status, email, photo FROM account WHERE lower(username) = lower(?)');
    stmt.setString(1, username);
    var rs = stmt.executeQuery();
    var out = rs.next() ? _accountFromRs(rs) : null;
    rs.close(); stmt.close();
    return out;
  });
}

function _dbGetAccountById(id, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT id, name, username, password_hash, role, status, email, photo FROM account WHERE id = ?');
    stmt.setLong(1, parseInt(id, 10));
    var rs = stmt.executeQuery();
    var out = rs.next() ? _accountFromRs(rs) : null;
    rs.close(); stmt.close();
    return out;
  });
}

// role/status default to 'user'/'active' (the only case registerUser needs) -
// pass them explicitly only from the migration script, which needs to
// preserve the existing admin account's role.
function _dbInsertAccount(name, username, passwordHash, role, status, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement(
      'INSERT INTO account (name, username, password_hash, role, status, first_login_at, last_login_at) ' +
      'VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP) RETURNING id'
    );
    stmt.setString(1, name);
    stmt.setString(2, username);
    stmt.setString(3, passwordHash);
    stmt.setInt(4, _roleToDb(role || 'user'));
    stmt.setInt(5, _statusToDb(status || 'active'));
    var rs = stmt.executeQuery();
    rs.next();
    var id = String(rs.getLong(1));
    rs.close(); stmt.close();
    return id;
  });
}

function _dbUpdateAccountLastLogin(id, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement(
      'UPDATE account SET last_login_at = CURRENT_TIMESTAMP, ' +
      'first_login_at = COALESCE(first_login_at, CURRENT_TIMESTAMP), mod_at = CURRENT_TIMESTAMP WHERE id = ?'
    );
    stmt.setLong(1, parseInt(id, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

// photo: pass a string to set it ('' clears it); omit/null to leave unchanged - same contract as updateProfile().
function _dbUpdateAccountProfile(id, displayName, photo, connOpt) {
  return _withConn(connOpt, function (conn) {
    var sql = (typeof photo === 'string')
      ? 'UPDATE account SET name = ?, photo = ?, mod_at = CURRENT_TIMESTAMP WHERE id = ?'
      : 'UPDATE account SET name = ?, mod_at = CURRENT_TIMESTAMP WHERE id = ?';
    var stmt = conn.prepareStatement(sql);
    stmt.setString(1, displayName);
    if (typeof photo === 'string') { stmt.setString(2, photo); stmt.setLong(3, parseInt(id, 10)); }
    else { stmt.setLong(2, parseInt(id, 10)); }
    stmt.executeUpdate();
    stmt.close();
  });
}

function _dbUpdateAccountPassword(id, passwordHash, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('UPDATE account SET password_hash = ?, mod_at = CURRENT_TIMESTAMP WHERE id = ?');
    stmt.setString(1, passwordHash);
    stmt.setLong(2, parseInt(id, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

function _dbIsAccountDisabled(id, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT status FROM account WHERE id = ?');
    stmt.setLong(1, parseInt(id, 10));
    var rs = stmt.executeQuery();
    var disabled = false;
    if (rs.next()) disabled = _statusFromDb(rs.getInt('status')) === 'disabled';
    rs.close(); stmt.close();
    return disabled;
  });
}

// Timestamps are cast to ISO-8601 text in SQL itself (to_char(...)) rather
// than read as JDBC Timestamp objects - Apps Script's JDBC wrapper doesn't
// reliably round-trip those to a JS Date the way the rest of this app
// expects (every other date already flows through as an ISO string).
function _dbGetAllAccounts(connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement(
      "SELECT id, name, username, role, status, photo, " +
      "to_char(first_login_at AT TIME ZONE 'UTC', 'YYYY-MM-DD\"T\"HH24:MI:SS\"Z\"') AS first_login_at, " +
      "to_char(last_login_at AT TIME ZONE 'UTC', 'YYYY-MM-DD\"T\"HH24:MI:SS\"Z\"') AS last_login_at " +
      "FROM account ORDER BY id"
    );
    var rs = stmt.executeQuery();
    var out = [];
    while (rs.next()) {
      out.push({
        id: String(rs.getLong('id')), displayName: rs.getString('name'), username: rs.getString('username'),
        role: _roleFromDb(rs.getInt('role')), status: _statusFromDb(rs.getInt('status')),
        photo: rs.getString('photo') || '',
        firstLogin: rs.getString('first_login_at') || '', lastLogin: rs.getString('last_login_at') || ''
      });
    }
    rs.close(); stmt.close();
    return out;
  });
}

function _dbUpdateAccountStatus(id, status, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('UPDATE account SET status = ?, mod_at = CURRENT_TIMESTAMP WHERE id = ?');
    stmt.setInt(1, _statusToDb(status));
    stmt.setLong(2, parseInt(id, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

function _dbUpdateAccountRole(id, role, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('UPDATE account SET role = ?, mod_at = CURRENT_TIMESTAMP WHERE id = ?');
    stmt.setInt(1, _roleToDb(role));
    stmt.setLong(2, parseInt(id, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

// friend rows cascade-delete with the account (ON DELETE CASCADE) - a real
// fix vs. the old Sheets version, which never cleaned up Friends at all.
function _dbDeleteAccount(id, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('DELETE FROM account WHERE id = ?');
    stmt.setLong(1, parseInt(id, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

function _dbGetFriendsByAccount(accountId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT id, name, is_self FROM friend WHERE account_id = ? ORDER BY id');
    stmt.setLong(1, parseInt(accountId, 10));
    var rs = stmt.executeQuery();
    var out = [];
    while (rs.next()) out.push({ id: String(rs.getLong('id')), name: rs.getString('name'), isSelf: rs.getBoolean('is_self') });
    rs.close(); stmt.close();
    return out;
  });
}

function _dbFindSelfFriend(accountId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT id, name FROM friend WHERE account_id = ? AND is_self = TRUE LIMIT 1');
    stmt.setLong(1, parseInt(accountId, 10));
    var rs = stmt.executeQuery();
    var out = rs.next() ? { id: String(rs.getLong('id')), name: rs.getString('name') } : null;
    rs.close(); stmt.close();
    return out;
  });
}

function _dbFindFriendByName(accountId, name, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT id, name FROM friend WHERE account_id = ? AND lower(name) = lower(?) LIMIT 1');
    stmt.setLong(1, parseInt(accountId, 10));
    stmt.setString(2, name);
    var rs = stmt.executeQuery();
    var out = rs.next() ? { id: String(rs.getLong('id')), name: rs.getString('name') } : null;
    rs.close(); stmt.close();
    return out;
  });
}

function _dbInsertFriend(accountId, name, isSelf, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('INSERT INTO friend (account_id, name, is_self) VALUES (?, ?, ?) RETURNING id');
    stmt.setLong(1, parseInt(accountId, 10));
    stmt.setString(2, name);
    stmt.setBoolean(3, !!isSelf);
    var rs = stmt.executeQuery();
    rs.next();
    var id = String(rs.getLong(1));
    rs.close(); stmt.close();
    return id;
  });
}

function _dbUpdateFriendName(id, name, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('UPDATE friend SET name = ?, mod_at = CURRENT_TIMESTAMP WHERE id = ?');
    stmt.setString(1, name);
    stmt.setLong(2, parseInt(id, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

// Binds `val` as a long, or a real SQL NULL when it's null/undefined - used
// for cre_by/mod_by everywhere below, since share-link (anonymous) writes
// and the migration script have no acting account to attribute a row to.
// -5 is java.sql.Types.BIGINT (a fixed JDBC spec constant, not an Apps
// Script API - Jdbc.getConnection()'s statements don't expose a Types helper).
function _setLongOrNull(stmt, idx, val) {
  if (val === null || val === undefined) stmt.setNull(idx, -5);
  else stmt.setLong(idx, parseInt(val, 10));
}

// Every _db* function below takes an optional trailing `connOpt` - pass an
// already-open connection to fold several calls into one round-trip's worth
// of connection setup (the dominant per-request cost against a remote DB,
// not the query itself); omit it and the function opens/closes its own,
// same as before. This is what lets RPCs like getHomeData/getDetailData
// collapse what used to be 3-7 separate connections into 1.
function _withConn(connOpt, fn) {
  var conn = connOpt || _dbConn();
  try {
    return fn(conn);
  } finally {
    if (!connOpt) conn.close();
  }
}

function _sharePermToDb(perm) { return perm === 'edit' ? 1 : 0; }
function _sharePermFromDb(perm) { return perm === 1 ? 'edit' : 'view'; }

// ----------------------------------------------------------------
// Postgres: session (replaces the Sessions sheet)
// ----------------------------------------------------------------

function _dbCreateSession(token, accountId, userInfo, expiresAtIso, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('INSERT INTO session (token, account_id, info, expires_at) VALUES (?, ?, ?::jsonb, ?::timestamptz)');
    stmt.setString(1, token);
    stmt.setLong(2, parseInt(accountId, 10));
    stmt.setString(3, JSON.stringify(userInfo));
    stmt.setString(4, expiresAtIso);
    stmt.executeUpdate();
    stmt.close();
  });
}

function _dbGetSession(token, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement(
      "SELECT info, to_char(expires_at AT TIME ZONE 'UTC', 'YYYY-MM-DD\"T\"HH24:MI:SS\"Z\"') AS expires_at " +
      'FROM session WHERE token = ?'
    );
    stmt.setString(1, token);
    var rs = stmt.executeQuery();
    var out = rs.next() ? { userInfo: JSON.parse(rs.getString('info')), expiresAt: rs.getString('expires_at') } : null;
    rs.close(); stmt.close();
    return out;
  });
}

function _dbDeleteSession(token, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('DELETE FROM session WHERE token = ?');
    stmt.setString(1, token);
    stmt.executeUpdate();
    stmt.close();
  });
}

function _dbUpdateSessionInfo(token, userInfo, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('UPDATE session SET info = ?::jsonb WHERE token = ?');
    stmt.setString(1, JSON.stringify(userInfo));
    stmt.setString(2, token);
    stmt.executeUpdate();
    stmt.close();
  });
}

function _dbCleanExpiredSessions(connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('DELETE FROM session WHERE expires_at < CURRENT_TIMESTAMP');
    stmt.executeUpdate();
    stmt.close();
  });
}

// ----------------------------------------------------------------
// Postgres: event (replaces Events + EventShares - share_token/share_perm/
// share_cre_at live as columns on event now, so there's no separate share table)
// ----------------------------------------------------------------

var EVENT_SELECT_COLS =
  "id, account_id, name, icon, is_active, share_token, share_perm, " +
  "to_char(cre_at AT TIME ZONE 'UTC', 'YYYY-MM-DD\"T\"HH24:MI:SS\"Z\"') AS cre_at";

function _dbEventFromRs(rs) {
  // share_perm reads as 0 (a JDBC-spec guarantee) when the column is NULL -
  // that happens to map to 'view' anyway, so no null-check is needed here.
  return {
    id: String(rs.getLong('id')), accountId: String(rs.getLong('account_id')), name: rs.getString('name'),
    icon: rs.getString('icon') || '', active: rs.getBoolean('is_active'), createdAt: rs.getString('cre_at'),
    shareToken: rs.getString('share_token') || '',
    sharePerm: _sharePermFromDb(rs.getInt('share_perm'))
  };
}

function _dbGetEventsByAccount(accountId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT ' + EVENT_SELECT_COLS + ' FROM event WHERE account_id = ? ORDER BY cre_at DESC');
    stmt.setLong(1, parseInt(accountId, 10));
    var rs = stmt.executeQuery();
    var out = [];
    while (rs.next()) out.push(_dbEventFromRs(rs));
    rs.close(); stmt.close();
    return out;
  });
}

function _dbGetEventById(eventId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT ' + EVENT_SELECT_COLS + ' FROM event WHERE id = ?');
    stmt.setLong(1, parseInt(eventId, 10));
    var rs = stmt.executeQuery();
    var out = rs.next() ? _dbEventFromRs(rs) : null;
    rs.close(); stmt.close();
    return out;
  });
}

function _dbGetEventByShareToken(shareToken, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT ' + EVENT_SELECT_COLS + ' FROM event WHERE share_token = ?');
    stmt.setString(1, shareToken);
    var rs = stmt.executeQuery();
    var out = rs.next() ? _dbEventFromRs(rs) : null;
    rs.close(); stmt.close();
    return out;
  });
}

function _dbInsertEvent(accountId, name, icon, createdBy, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement(
      'INSERT INTO event (account_id, name, icon, cre_by, mod_by) VALUES (?, ?, ?, ?, ?) ' +
      "RETURNING id, to_char(cre_at AT TIME ZONE 'UTC', 'YYYY-MM-DD\"T\"HH24:MI:SS\"Z\"') AS cre_at"
    );
    stmt.setLong(1, parseInt(accountId, 10));
    stmt.setString(2, name);
    stmt.setString(3, icon || '');
    _setLongOrNull(stmt, 4, createdBy);
    _setLongOrNull(stmt, 5, createdBy);
    var rs = stmt.executeQuery();
    rs.next();
    var out = { id: String(rs.getLong('id')), createdAt: rs.getString('cre_at') };
    rs.close(); stmt.close();
    return out;
  });
}

function _dbUpdateEventName(eventId, name, icon, modBy, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('UPDATE event SET name = ?, icon = ?, mod_at = CURRENT_TIMESTAMP, mod_by = ? WHERE id = ?');
    stmt.setString(1, name);
    stmt.setString(2, icon || '');
    _setLongOrNull(stmt, 3, modBy);
    stmt.setLong(4, parseInt(eventId, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

function _dbSetEventActive(eventId, active, modBy, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('UPDATE event SET is_active = ?, mod_at = CURRENT_TIMESTAMP, mod_by = ? WHERE id = ?');
    stmt.setBoolean(1, !!active);
    _setLongOrNull(stmt, 2, modBy);
    stmt.setLong(3, parseInt(eventId, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

function _dbDeleteEvent(eventId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('DELETE FROM event WHERE id = ?'); // cascades to transaction/split/participant/settlement/receipt
    stmt.setLong(1, parseInt(eventId, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

function _dbSetEventShare(eventId, shareToken, permission, modBy, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement(
      'UPDATE event SET share_token = ?, share_perm = ?, share_cre_at = COALESCE(share_cre_at, CURRENT_TIMESTAMP), ' +
      'mod_at = CURRENT_TIMESTAMP, mod_by = ? WHERE id = ?'
    );
    stmt.setString(1, shareToken);
    stmt.setInt(2, _sharePermToDb(permission));
    _setLongOrNull(stmt, 3, modBy);
    stmt.setLong(4, parseInt(eventId, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

function _dbClearEventShare(eventId, modBy, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement(
      'UPDATE event SET share_token = NULL, share_perm = NULL, share_cre_at = NULL, mod_at = CURRENT_TIMESTAMP, mod_by = ? WHERE id = ?'
    );
    _setLongOrNull(stmt, 1, modBy);
    stmt.setLong(2, parseInt(eventId, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

// ----------------------------------------------------------------
// Postgres: transaction + split (replaces Details, whose 'splits' column was
// a JSON blob {friendId: amount} - now real rows, one per participant)
// ----------------------------------------------------------------

// Returns this event's transactions, each with its splits already attached -
// the shape _buildDetailPayload/getHomeData/getSummary need, no JSON parsing.
function _dbGetTransactionsByEvent(eventId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var txStmt = conn.prepareStatement(
      "SELECT id, payer_id, amount, description, (excluded_at IS NOT NULL) AS excluded, " +
      "to_char(cre_at AT TIME ZONE 'UTC', 'YYYY-MM-DD\"T\"HH24:MI:SS\"Z\"') AS cre_at " +
      'FROM transaction WHERE event_id = ? ORDER BY cre_at'
    );
    txStmt.setLong(1, parseInt(eventId, 10));
    var txRs = txStmt.executeQuery();
    var out = [], byId = {};
    while (txRs.next()) {
      var tx = {
        id: String(txRs.getLong('id')), payId: String(txRs.getLong('payer_id')), amount: txRs.getDouble('amount'),
        description: txRs.getString('description') || '', excluded: txRs.getBoolean('excluded'), createdAt: txRs.getString('cre_at'),
        splits: []
      };
      out.push(tx);
      byId[tx.id] = tx;
    }
    txRs.close(); txStmt.close();

    if (out.length) {
      var spStmt = conn.prepareStatement(
        'SELECT s.transaction_id, s.friend_id, s.amount FROM split s JOIN transaction t ON t.id = s.transaction_id WHERE t.event_id = ?'
      );
      spStmt.setLong(1, parseInt(eventId, 10));
      var spRs = spStmt.executeQuery();
      while (spRs.next()) {
        var tid = String(spRs.getLong('transaction_id'));
        if (byId[tid]) byId[tid].splits.push({ friendId: String(spRs.getLong('friend_id')), amount: spRs.getDouble('amount') });
      }
      spRs.close(); spStmt.close();
    }
    return out;
  });
}

function _dbGetTransactionEventId(transactionId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT event_id FROM transaction WHERE id = ?');
    stmt.setLong(1, parseInt(transactionId, 10));
    var rs = stmt.executeQuery();
    var out = rs.next() ? String(rs.getLong('event_id')) : null;
    rs.close(); stmt.close();
    return out;
  });
}

// friendIds/customAmounts: same contract as the old _buildDetailRow - equal
// split unless customAmounts[fid] overrides it.
function _dbInsertTransactionWithSplits(eventId, payerId, friendIds, totalAmount, description, customAmounts, createdBy, connOpt) {
  return _withConn(connOpt, function (conn) {
    var total = parseFloat(totalAmount);
    var perPerson = total / friendIds.length;
    var txStmt = conn.prepareStatement(
      'INSERT INTO transaction (event_id, payer_id, amount, description, cre_by, mod_by) VALUES (?, ?, ?, ?, ?, ?) RETURNING id'
    );
    txStmt.setLong(1, parseInt(eventId, 10));
    txStmt.setLong(2, parseInt(payerId, 10));
    txStmt.setDouble(3, total);
    txStmt.setString(4, description || '');
    _setLongOrNull(txStmt, 5, createdBy);
    _setLongOrNull(txStmt, 6, createdBy);
    var rs = txStmt.executeQuery();
    rs.next();
    var transactionId = String(rs.getLong(1));
    rs.close(); txStmt.close();

    var spStmt = conn.prepareStatement('INSERT INTO split (transaction_id, friend_id, amount, cre_by, mod_by) VALUES (?, ?, ?, ?, ?)');
    friendIds.forEach(function (fid) {
      var amt = (customAmounts && customAmounts[fid] !== undefined) ? parseFloat(customAmounts[fid]) : perPerson;
      spStmt.setLong(1, parseInt(transactionId, 10));
      spStmt.setLong(2, parseInt(fid, 10));
      spStmt.setDouble(3, amt);
      _setLongOrNull(spStmt, 4, createdBy);
      _setLongOrNull(spStmt, 5, createdBy);
      spStmt.executeUpdate();
    });
    spStmt.close();
    return transactionId;
  });
}

function _dbUpdateTransactionWithSplits(transactionId, payerId, friendIds, totalAmount, description, customAmounts, modBy, connOpt) {
  return _withConn(connOpt, function (conn) {
    var total = parseFloat(totalAmount);
    var perPerson = total / friendIds.length;
    var txStmt = conn.prepareStatement('UPDATE transaction SET payer_id = ?, amount = ?, description = ?, mod_at = CURRENT_TIMESTAMP, mod_by = ? WHERE id = ?');
    txStmt.setLong(1, parseInt(payerId, 10));
    txStmt.setDouble(2, total);
    txStmt.setString(3, description || '');
    _setLongOrNull(txStmt, 4, modBy);
    txStmt.setLong(5, parseInt(transactionId, 10));
    txStmt.executeUpdate();
    txStmt.close();

    var delStmt = conn.prepareStatement('DELETE FROM split WHERE transaction_id = ?');
    delStmt.setLong(1, parseInt(transactionId, 10));
    delStmt.executeUpdate();
    delStmt.close();

    var spStmt = conn.prepareStatement('INSERT INTO split (transaction_id, friend_id, amount, cre_by, mod_by) VALUES (?, ?, ?, ?, ?)');
    friendIds.forEach(function (fid) {
      var amt = (customAmounts && customAmounts[fid] !== undefined) ? parseFloat(customAmounts[fid]) : perPerson;
      spStmt.setLong(1, parseInt(transactionId, 10));
      spStmt.setLong(2, parseInt(fid, 10));
      spStmt.setDouble(3, amt);
      _setLongOrNull(spStmt, 4, modBy);
      _setLongOrNull(spStmt, 5, modBy);
      spStmt.executeUpdate();
    });
    spStmt.close();
  });
}

function _dbDeleteTransaction(transactionId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('DELETE FROM transaction WHERE id = ?'); // cascades to split/receipt
    stmt.setLong(1, parseInt(transactionId, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

// Replaces TransactionPayments - a whole expense excluded from settlement
// math is now a column on transaction itself instead of a side table.
function _dbSetTransactionExcluded(transactionId, excluded, modBy, connOpt) {
  return _withConn(connOpt, function (conn) {
    var sql = excluded
      ? 'UPDATE transaction SET excluded_at = ?::timestamptz, mod_at = CURRENT_TIMESTAMP, mod_by = ? WHERE id = ?'
      : 'UPDATE transaction SET excluded_at = NULL, mod_at = CURRENT_TIMESTAMP, mod_by = ? WHERE id = ?';
    var stmt = conn.prepareStatement(sql);
    if (excluded) {
      stmt.setString(1, new Date().toISOString());
      _setLongOrNull(stmt, 2, modBy);
      stmt.setLong(3, parseInt(transactionId, 10));
    } else {
      _setLongOrNull(stmt, 1, modBy);
      stmt.setLong(2, parseInt(transactionId, 10));
    }
    stmt.executeUpdate();
    stmt.close();
  });
}

// ----------------------------------------------------------------
// Postgres: participant (replaces EventFriends)
// ----------------------------------------------------------------

function _dbGetParticipantsByEvent(eventId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT p.friend_id, f.name FROM participant p JOIN friend f ON f.id = p.friend_id WHERE p.event_id = ? ORDER BY p.cre_at');
    stmt.setLong(1, parseInt(eventId, 10));
    var rs = stmt.executeQuery();
    var out = [];
    while (rs.next()) out.push({ id: String(rs.getLong('friend_id')), name: rs.getString('name') });
    rs.close(); stmt.close();
    return out;
  });
}

function _dbAddParticipant(eventId, friendId, createdBy, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('INSERT INTO participant (event_id, friend_id, cre_by, mod_by) VALUES (?, ?, ?, ?) ON CONFLICT (event_id, friend_id) DO NOTHING');
    stmt.setLong(1, parseInt(eventId, 10));
    stmt.setLong(2, parseInt(friendId, 10));
    _setLongOrNull(stmt, 3, createdBy);
    _setLongOrNull(stmt, 4, createdBy);
    stmt.executeUpdate();
    stmt.close();
  });
}

// Bulk add/remove in one connection - setEventFriends changes several rows
// at once, same reasoning as the old single-read/single-write sheet version.
function _dbSyncParticipants(eventId, toAddIds, toRemoveIds, actorId, connOpt) {
  return _withConn(connOpt, function (conn) {
    if (toRemoveIds && toRemoveIds.length) {
      var delStmt = conn.prepareStatement('DELETE FROM participant WHERE event_id = ? AND friend_id = ?');
      toRemoveIds.forEach(function (fid) {
        delStmt.setLong(1, parseInt(eventId, 10));
        delStmt.setLong(2, parseInt(fid, 10));
        delStmt.executeUpdate();
      });
      delStmt.close();
    }
    if (toAddIds && toAddIds.length) {
      var insStmt = conn.prepareStatement('INSERT INTO participant (event_id, friend_id, cre_by, mod_by) VALUES (?, ?, ?, ?) ON CONFLICT (event_id, friend_id) DO NOTHING');
      toAddIds.forEach(function (fid) {
        insStmt.setLong(1, parseInt(eventId, 10));
        insStmt.setLong(2, parseInt(fid, 10));
        _setLongOrNull(insStmt, 3, actorId);
        _setLongOrNull(insStmt, 4, actorId);
        insStmt.executeUpdate();
      });
      insStmt.close();
    }
  });
}

// Friend ids that actually appear as a payer or split participant in this
// event's transactions - blocks removing them from participant below.
function _dbUsedFriendIdsInEvent(eventId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement(
      'SELECT DISTINCT payer_id AS friend_id FROM transaction WHERE event_id = ? ' +
      'UNION SELECT DISTINCT s.friend_id FROM split s JOIN transaction t ON t.id = s.transaction_id WHERE t.event_id = ?'
    );
    stmt.setLong(1, parseInt(eventId, 10));
    stmt.setLong(2, parseInt(eventId, 10));
    var rs = stmt.executeQuery();
    var set = {};
    while (rs.next()) set[String(rs.getLong('friend_id'))] = true;
    rs.close(); stmt.close();
    return set;
  });
}

// ----------------------------------------------------------------
// Postgres: settlement (replaces SettlementPayments)
// ----------------------------------------------------------------

// Keyed exactly like _settleKey below - a drop-in replacement for the old
// sheet-scanned paid set.
function _dbGetSettlementsByEvent(eventId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT from_id, to_id, amount FROM settlement WHERE event_id = ?');
    stmt.setLong(1, parseInt(eventId, 10));
    var rs = stmt.executeQuery();
    var set = {};
    while (rs.next()) set[_settleKey(String(rs.getLong('from_id')), String(rs.getLong('to_id')), rs.getDouble('amount'))] = true;
    rs.close(); stmt.close();
    return set;
  });
}

function _dbUpsertSettlement(eventId, fromId, toId, amount, actorId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement(
      'INSERT INTO settlement (event_id, from_id, to_id, amount, settled_at, cre_by, mod_by) VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP, ?, ?) ' +
      'ON CONFLICT (event_id, from_id, to_id) DO UPDATE SET amount = EXCLUDED.amount, settled_at = CURRENT_TIMESTAMP, mod_at = CURRENT_TIMESTAMP, mod_by = EXCLUDED.mod_by'
    );
    stmt.setLong(1, parseInt(eventId, 10));
    stmt.setLong(2, parseInt(fromId, 10));
    stmt.setLong(3, parseInt(toId, 10));
    stmt.setDouble(4, parseFloat(amount));
    _setLongOrNull(stmt, 5, actorId);
    _setLongOrNull(stmt, 6, actorId);
    stmt.executeUpdate();
    stmt.close();
  });
}

function _dbDeleteSettlement(eventId, fromId, toId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('DELETE FROM settlement WHERE event_id = ? AND from_id = ? AND to_id = ?');
    stmt.setLong(1, parseInt(eventId, 10));
    stmt.setLong(2, parseInt(fromId, 10));
    stmt.setLong(3, parseInt(toId, 10));
    stmt.executeUpdate();
    stmt.close();
  });
}

// ----------------------------------------------------------------
// Postgres: receipt (replaces TransactionSlips) - Drive upload/read/trash
// logic below is unchanged; only the row bookkeeping moves here.
// ----------------------------------------------------------------

function _dbGetReceiptsByEvent(eventId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement(
      'SELECT r.transaction_id, r.id, r.file_id FROM receipt r JOIN transaction t ON t.id = r.transaction_id WHERE t.event_id = ? ORDER BY r.cre_at'
    );
    stmt.setLong(1, parseInt(eventId, 10));
    var rs = stmt.executeQuery();
    var out = {};
    while (rs.next()) {
      var tid = String(rs.getLong('transaction_id'));
      if (!out[tid]) out[tid] = [];
      out[tid].push({ id: String(rs.getLong('id')), slip: rs.getString('file_id'), slipHi: rs.getString('file_id') });
    }
    rs.close(); stmt.close();
    return out;
  });
}

function _dbInsertReceipt(transactionId, fileId, createdBy, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('INSERT INTO receipt (transaction_id, file_id, cre_by, mod_by) VALUES (?, ?, ?, ?) RETURNING id');
    stmt.setLong(1, parseInt(transactionId, 10));
    stmt.setString(2, fileId);
    _setLongOrNull(stmt, 3, createdBy);
    _setLongOrNull(stmt, 4, createdBy);
    var rs = stmt.executeQuery();
    rs.next();
    var id = String(rs.getLong(1));
    rs.close(); stmt.close();
    return id;
  });
}

// Returns the deleted row's file_id (for Drive cleanup) or null if no match.
function _dbDeleteReceipt(transactionId, receiptId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('DELETE FROM receipt WHERE transaction_id = ? AND id = ? RETURNING file_id');
    stmt.setLong(1, parseInt(transactionId, 10));
    stmt.setLong(2, parseInt(receiptId, 10));
    var rs = stmt.executeQuery();
    var fileId = rs.next() ? rs.getString('file_id') : null;
    rs.close(); stmt.close();
    return fileId;
  });
}

function _dbIsKnownReceiptFile(fileId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT 1 FROM receipt WHERE file_id = ?');
    stmt.setString(1, fileId);
    var rs = stmt.executeQuery();
    var out = rs.next();
    rs.close(); stmt.close();
    return out;
  });
}

function _dbGetReceiptFileIdsByTransaction(transactionId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT file_id FROM receipt WHERE transaction_id = ?');
    stmt.setLong(1, parseInt(transactionId, 10));
    var rs = stmt.executeQuery();
    var out = [];
    while (rs.next()) out.push(rs.getString('file_id'));
    rs.close(); stmt.close();
    return out;
  });
}

function _dbGetReceiptFileIdsByEvent(eventId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var stmt = conn.prepareStatement('SELECT r.file_id FROM receipt r JOIN transaction t ON t.id = r.transaction_id WHERE t.event_id = ?');
    stmt.setLong(1, parseInt(eventId, 10));
    var rs = stmt.executeQuery();
    var out = [];
    while (rs.next()) out.push(rs.getString('file_id'));
    rs.close(); stmt.close();
    return out;
  });
}

function getCache() {
  return CacheService.getScriptCache();
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

function _readSlipImage(fileId, connOpt) {
  if (!_isKnownSlipFile(fileId, connOpt)) return _fail('Photo not found');
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
    var conn = _dbConn();
    try {
      if (!_shareEventId(shareToken, false, conn)) return _fail('Invalid link');
      return _readSlipImage(fileId, conn);
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function _isKnownSlipFile(fileId, connOpt) {
  return _dbIsKnownReceiptFile(fileId, connOpt);
}

function _deleteSlipFile(fileId) {
  if (!fileId) return; // legacy rows predating Drive storage have no fileId - nothing to trash
  try { DriveApp.getFileById(fileId).setTrashed(true) } catch (e) { /* already gone - ignore */ }
}

function _trashSlipFilesForTx(transactionId, connOpt) {
  _dbGetReceiptFileIdsByTransaction(transactionId, connOpt).forEach(_deleteSlipFile);
}

function _trashSlipFilesForEvent(eventId, connOpt) {
  _dbGetReceiptFileIdsByEvent(eventId, connOpt).forEach(_deleteSlipFile);
}

function _lookupSession(token) {
  var found = _dbGetSession(token);
  if (!found) return null;
  if (new Date() >= new Date(found.expiresAt)) { _dbDeleteSession(token); return null }
  if (_isAccountDisabled(found.userInfo.id)) { _dbDeleteSession(token); return null }
  return { userInfo: found.userInfo, expiresAt: found.expiresAt };
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
  return _dbIsAccountDisabled(accountId);
}

function _cleanExpiredSessions() {
  try { _dbCleanExpiredSessions(); } catch (e) {}
}

// ----------------------------------------------------------------
// Entry Point
// ----------------------------------------------------------------

// Looked up at render time (not via getSharedEventView) so doGet can skip
// sending the add/edit/delete transaction markup entirely for the common
// view-only case, instead of shipping it and hiding it with CSS.
function _sharePermissionByToken(shareToken) {
  if (!shareToken) return 'view';
  var event = _dbGetEventByShareToken(shareToken);
  return event ? event.sharePerm : 'view';
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
    var conn = _dbConn();
    var account, token, userInfo, expires;
    try {
      account = _dbGetAccountByUsername(username, conn);
      var hashed = hashPassword(password);
      if (!account || account.passwordHash !== hashed) return _fail('Invalid username or password');
      if (account.status === 'disabled') return _fail('This account has been disabled');
      _dbUpdateAccountLastLogin(account.id, conn);

      token = generateToken();
      userInfo = { id: account.id, displayName: account.displayName, username: account.username, role: account.role };
      var sessionMinutes = getAppSettings().sessionMinutes;
      expires = new Date(new Date().getTime() + sessionMinutes * 60000);
      var cacheTtl = Math.max(1, Math.min(CACHE_EXPIRY, sessionMinutes * 60));
      getCache().put('token_' + token, JSON.stringify(userInfo), cacheTtl);
      _dbCreateSession(token, account.id, userInfo, expires.toISOString(), conn);
      _dbCleanExpiredSessions(conn);
    } finally { conn.close(); }
    return { success: true, token: token, user: userInfo, url: ScriptApp.getService().getUrl() + '?tk=' + encodeURIComponent(token) };
  } catch (e) {
    return _fail(e);
  }
}

function registerUser(displayName, username, password) {
  try {
    var pwErr = _validatePassword(password);
    if (pwErr) return _fail(pwErr);

    var conn = _dbConn();
    try {
      if (_dbGetAccountByUsername(username, conn)) return _fail('Username already taken');
      var hashed = hashPassword(password);
      var id = _dbInsertAccount(displayName, username.toLowerCase(), hashed, 'user', 'active', conn);
      _dbInsertFriend(id, 'Me', true, conn);
      return { success: true };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function logoutUser(token) {
  try {
    getCache().remove('token_' + token);
    _dbDeleteSession(token);
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
  var found = _dbGetSession(token);
  var ttl = CACHE_EXPIRY;
  if (found) {
    _dbUpdateSessionInfo(token, userInfo);
    ttl = _cacheTtlFor(found.expiresAt);
  }
  getCache().put('token_' + token, JSON.stringify(userInfo), ttl);
}

// ----------------------------------------------------------------
// Profile (self-service account settings)
// ----------------------------------------------------------------

function getMyProfile(token) {
  try {
    var user = requireAuth(token);
    var account = _dbGetAccountById(user.id);
    if (!account) return _fail('Account not found');
    return { success: true, displayName: account.displayName, username: account.username, photo: account.photo };
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

    var conn = _dbConn();
    try {
      var account = _dbGetAccountById(user.id, conn);
      if (!account) return _fail('Account not found');
      _dbUpdateAccountProfile(user.id, trimmed, photo, conn);

      // Keep the self-friend (shown as payer/participant in every event) in sync.
      var selfFriend = _dbFindSelfFriend(user.id, conn);
      if (selfFriend) _dbUpdateFriendName(selfFriend.id, trimmed, conn);
    } finally { conn.close(); }

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
    var conn = _dbConn();
    try {
      var account = _dbGetAccountById(user.id, conn);
      if (!account) return _fail('Account not found');
      if (account.passwordHash !== hashPassword(currentPassword || '')) return _fail('Current password is incorrect');
      _dbUpdateAccountPassword(user.id, hashPassword(newPassword), conn);
      return { success: true };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

// ----------------------------------------------------------------
// Combined data fetchers (reduce round-trips)
// ----------------------------------------------------------------

// One connection covers every one of the account's events at once - avoids
// opening a fresh connection per event the way a naive per-event loop would.
function _dbGetHomeSettlementInput(accountId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var rowsByEvent = {};
    var txStmt = conn.prepareStatement(
      'SELECT t.event_id, t.payer_id, s.friend_id, s.amount FROM transaction t ' +
      'JOIN split s ON s.transaction_id = t.id JOIN event e ON e.id = t.event_id ' +
      'WHERE e.account_id = ? AND t.excluded_at IS NULL'
    );
    txStmt.setLong(1, parseInt(accountId, 10));
    var txRs = txStmt.executeQuery();
    while (txRs.next()) {
      var eid = String(txRs.getLong('event_id'));
      if (!rowsByEvent[eid]) rowsByEvent[eid] = [];
      rowsByEvent[eid].push({ payId: String(txRs.getLong('payer_id')), friendId: String(txRs.getLong('friend_id')), amount: txRs.getDouble('amount') });
    }
    txRs.close(); txStmt.close();

    var paidSetByEvent = {};
    var spStmt = conn.prepareStatement(
      'SELECT s.event_id, s.from_id, s.to_id, s.amount FROM settlement s JOIN event e ON e.id = s.event_id WHERE e.account_id = ?'
    );
    spStmt.setLong(1, parseInt(accountId, 10));
    var spRs = spStmt.executeQuery();
    while (spRs.next()) {
      var eid2 = String(spRs.getLong('event_id'));
      if (!paidSetByEvent[eid2]) paidSetByEvent[eid2] = {};
      paidSetByEvent[eid2][_settleKey(String(spRs.getLong('from_id')), String(spRs.getLong('to_id')), spRs.getDouble('amount'))] = true;
    }
    spRs.close(); spStmt.close();

    return { rowsByEvent: rowsByEvent, paidSetByEvent: paidSetByEvent };
  });
}

function getHomeData(token) {
  try {
    var user = requireAuth(token);
    // One shared connection for all three reads below, instead of three
    // separate round-trips to Neon - this RPC runs on every Home load.
    var conn = _dbConn();
    var events, friends, friendMap = {}, input;
    try {
      events = _dbGetEventsByAccount(user.id, conn);
      var dbFriends = _dbGetFriendsByAccount(user.id, conn);
      friends = [];
      dbFriends.forEach(function (f) {
        friends.push({ id: f.id, accountId: user.id, name: f.name });
        friendMap[f.id] = f.name;
      });
      // Settlement state per event for the Home filter tabs, via the same
      // engine as getSummary.
      input = _dbGetHomeSettlementInput(user.id, conn);
    } finally { conn.close(); }

    events.forEach(function (ev) {
      var rows = input.rowsByEvent[ev.id] || [];
      if (!rows.length) { ev.settled = true; return }
      var settlements = _computeSettlements(rows, friendMap);
      var paidSet = input.paidSetByEvent[ev.id] || {};
      ev.settled = settlements.every(function (s) { return !!paidSet[_settleKey(s.from, s.to, s.amount)] });
    });

    return { success: true, events: events, friends: friends };
  } catch (e) { return _fail(e) }
}

// Shared core behind getDetailData (authenticated) and getSharedEventView
// (public share link) so both paths compute details/friends/settlements/slips
// identically instead of maintaining two parallel implementations.
// connOpt: share one connection across every read below (transactions/
// splits, self-friend, participants, settlements, receipts) instead of
// opening 5-6 separate ones - this backs the most-loaded page in the app
// (Detail), both for the owner and for public share-link visitors.
function _buildDetailPayload(eventId, accountId, connOpt) {
  return _withConn(connOpt, function (conn) {
    var transactions = _dbGetTransactionsByEvent(eventId, conn);
    var details = [], rows = [];
    transactions.forEach(function (tx) {
      tx.splits.forEach(function (s) {
        details.push({
          id: tx.id + '_' + s.friendId, eventId: eventId, transactionId: tx.id,
          payId: tx.payId, friendId: s.friendId, amount: s.amount, totalAmount: tx.amount,
          description: tx.description, createdAt: tx.createdAt, paid: tx.excluded
        });
      });
      // Excluded transactions are left out of settlement math entirely - see markTransactionPaid.
      if (!tx.excluded) tx.splits.forEach(function (s) { rows.push({ payId: tx.payId, friendId: s.friendId, amount: s.amount }) });
    });

    var selfFriend = _dbFindSelfFriend(accountId, conn);
    var selfFriendId = selfFriend ? selfFriend.id : null;

    var friends = _dbGetParticipantsByEvent(eventId, conn);
    var friendMap = {};
    friends.forEach(function (f) { friendMap[f.id] = f.name });
    // _computeSettlementsWithPaid so share-link visitors see the same "paid"
    // checkmarks the owner does, bundled here to avoid a second round-trip.
    var settlements = _computeSettlementsWithPaid(rows, friendMap, eventId, conn);

    var slips = _dbGetReceiptsByEvent(eventId, conn);

    // selfFriendId: lets the client show the account's own profile photo.
    return { details: details, friends: friends, settlements: settlements, selfFriendId: selfFriendId, slips: slips };
  });
}

function getDetailData(token, eventId) {
  try {
    var user = requireAuth(token);
    var conn = _dbConn();
    try {
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Event not found');
      var payload = _buildDetailPayload(eventId, user.id, conn);
      payload.success = true;
      return payload;
    } finally { conn.close(); }
  } catch (e) { return _fail(e) }
}

// ----------------------------------------------------------------
// Event-scoped friend membership
// ----------------------------------------------------------------

function _eventOwnedBy(eventId, accountId, connOpt) {
  var event = _dbGetEventById(eventId, connOpt);
  return !!event && event.accountId === accountId;
}

// Combined fetch for the Add/Manage Friends sheet — one round trip instead of
// separate getFriends + getEventFriends calls.
function getEventFriendsData(token, eventId) {
  try {
    var user = requireAuth(token);
    var conn = _dbConn();
    try {
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Event not found');
      var allFriends = _dbGetFriendsByAccount(user.id, conn).map(function (f) { return { id: f.id, name: f.name }; });
      var linkedFriends = _dbGetParticipantsByEvent(eventId, conn);
      return { success: true, allFriends: allFriends, linkedFriends: linkedFriends };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function addFriendToEvent(token, eventId, name) {
  try {
    var user = requireAuth(token);
    if (!name || name.trim() === '') return _fail('Name is required');
    var trimmed = name.trim();
    var conn = _dbConn();
    try {
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Event not found');
      var existing = _dbFindFriendByName(user.id, trimmed, conn);
      var friendId = existing ? existing.id : _dbInsertFriend(user.id, trimmed, false, conn);
      _dbAddParticipant(eventId, friendId, user.id, conn);
      return { success: true, friend: { id: friendId, name: trimmed } };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function setEventFriends(token, eventId, friendIds) {
  try {
    var user = requireAuth(token);
    var conn = _dbConn();
    try {
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Event not found');

      var ownedFriendMap = {};
      _dbGetFriendsByAccount(user.id, conn).forEach(function (f) { ownedFriendMap[f.id] = f.name });
      var wantedIds = (friendIds || []).filter(function (fid) { return ownedFriendMap.hasOwnProperty(fid); });

      var usedInEvent = _dbUsedFriendIdsInEvent(eventId, conn);
      var currentIdSet = {};
      _dbGetParticipantsByEvent(eventId, conn).forEach(function (f) { currentIdSet[f.id] = true; });

      var blocked = [];
      var toAdd = wantedIds.filter(function (fid) { return !currentIdSet[fid]; });
      var toRemove = Object.keys(currentIdSet).filter(function (fid) {
        if (wantedIds.indexOf(fid) !== -1) return false;
        if (usedInEvent[fid]) { blocked.push({ id: fid, name: ownedFriendMap[fid] }); return false; }
        return true;
      });

      _dbSyncParticipants(eventId, toAdd, toRemove, user.id, conn);
      toRemove.forEach(function (fid) { delete currentIdSet[fid]; });
      toAdd.forEach(function (fid) { currentIdSet[fid] = true; });

      var finalFriends = Object.keys(currentIdSet).map(function (fid) {
        return { id: fid, name: ownedFriendMap[fid] };
      });
      return { success: true, friends: finalFriends, blocked: blocked };
    } finally { conn.close(); }
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
    var trimmed = name.trim();
    var conn = _dbConn();
    try {
      var created = _dbInsertEvent(user.id, trimmed, icon || '', user.id, conn);

      // Auto-link the account's own self-friend so every event starts with yourself in it
      var selfFriend = _dbFindSelfFriend(user.id, conn);
      if (selfFriend) _dbAddParticipant(created.id, selfFriend.id, user.id, conn);

      return { success: true, event: { id: created.id, name: trimmed, accountId: user.id, createdAt: created.createdAt, icon: icon || '' } };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function renameEvent(token, eventId, name, icon) {
  try {
    var user = requireAuth(token);
    if (!name || name.trim() === '') return _fail('Event name is required');
    var conn = _dbConn();
    try {
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Event not found');
      _dbUpdateEventName(eventId, name.trim(), icon || '', user.id, conn);
      return { success: true, name: name.trim(), icon: icon || '' };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function setEventActive(token, eventId, active) {
  try {
    var user = requireAuth(token);
    var conn = _dbConn();
    try {
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Event not found');
      _dbSetEventActive(eventId, active === true, user.id, conn);
      return { success: true, active: active === true };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function deleteEvent(token, eventId) {
  try {
    var user = requireAuth(token);
    var conn = _dbConn();
    try {
      // Check ownership before deleting anything.
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Event not found');

      _trashSlipFilesForEvent(eventId, conn);
      _dbDeleteEvent(eventId, conn); // cascades to transaction/split/participant/settlement/receipt

      return { success: true };
    } finally { conn.close(); }
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
    // One fetch covers both the ownership check and the share info - no
    // separate _eventOwnedBy round-trip needed.
    var event = _dbGetEventById(eventId);
    if (!event || event.accountId !== user.id) return _fail('Event not found');
    if (!event.shareToken) return { success: true, shareToken: null };
    return {
      success: true, shareToken: event.shareToken, permission: event.sharePerm,
      shareUrl: ScriptApp.getService().getUrl() + '?share=' + event.shareToken
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
    var conn = _dbConn();
    try {
      var event = _dbGetEventById(eventId, conn);
      if (!event || event.accountId !== user.id) return _fail('Event not found');
      var perm = permission === 'edit' ? 'edit' : 'view';

      var shareToken = event.shareToken || Utilities.getUuid();
      _dbSetEventShare(eventId, shareToken, perm, user.id, conn);
      return { success: true, shareToken: shareToken, permission: perm, shareUrl: ScriptApp.getService().getUrl() + '?share=' + shareToken };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function disableEventShare(token, eventId) {
  try {
    var user = requireAuth(token);
    var conn = _dbConn();
    try {
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Event not found');
      _dbClearEventShare(eventId, user.id, conn);
      return { success: true };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

// Public — intentionally takes no auth token. Only ever returns the one
// event a valid, unguessable share token points to; never account data.
function getSharedEventView(shareToken) {
  try {
    if (!shareToken) return _fail('Invalid link');
    var conn = _dbConn();
    try {
      var event = _dbGetEventByShareToken(shareToken, conn);
      if (!event) return _fail('This share link is no longer active');

      var ownerAccount = _dbGetAccountById(event.accountId, conn);
      var ownerPhoto = ownerAccount ? ownerAccount.photo : '';

      // Same core the authenticated getDetailData uses, so a share visitor sees
      // identical details/friends/settlements (including paid state)/slips
      // shapes — the client renders both through the exact same Detail code.
      var payload = _buildDetailPayload(event.id, event.accountId, conn);

      return {
        success: true,
        event: { name: event.name, createdAt: event.createdAt, icon: event.icon },
        details: payload.details,
        friends: payload.friends,
        settlements: payload.settlements,
        selfFriendId: payload.selfFriendId,
        ownerPhoto: ownerPhoto,
        slips: payload.slips,
        permission: event.sharePerm
      };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

// ----------------------------------------------------------------
// Details (Transactions)
// ----------------------------------------------------------------

function addDetail(token, eventId, payId, friendIds, totalAmount, description, customAmounts) {
  try {
    var user = requireAuth(token);
    if (!friendIds || !friendIds.length) return _fail('At least one person is required');
    var transactionId = _dbInsertTransactionWithSplits(eventId, payId, friendIds, totalAmount, description, customAmounts, user.id);
    return { success: true, transactionId: transactionId };
  } catch (e) {
    return _fail(e);
  }
}

function updateDetail(token, transactionId, payId, friendIds, totalAmount, description, customAmounts) {
  try {
    var user = requireAuth(token);
    if (!friendIds || !friendIds.length) return _fail('At least one person is required');
    var conn = _dbConn();
    try {
      var eventId = _dbGetTransactionEventId(transactionId, conn);
      if (!eventId) return _fail('Transaction not found');
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Transaction not found');

      _dbUpdateTransactionWithSplits(transactionId, payId, friendIds, totalAmount, description, customAmounts, user.id, conn);
      return { success: true, transactionId: transactionId };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function deleteDetail(token, transactionId) {
  try {
    var user = requireAuth(token);
    var conn = _dbConn();
    try {
      var eventId = _dbGetTransactionEventId(transactionId, conn);
      if (!eventId) return _fail('Transaction not found');
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Transaction not found');

      _trashSlipFilesForTx(transactionId, conn);
      _dbDeleteTransaction(transactionId, conn); // cascades to split/receipt
      return { success: true };
    } finally { conn.close(); }
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

function _shareEventId(shareToken, requireEdit, connOpt) {
  if (!shareToken) return null;
  var event = _dbGetEventByShareToken(shareToken, connOpt);
  if (!event) return null;
  if (requireEdit && event.sharePerm !== 'edit') return null;
  return event.id;
}

function addDetailViaShare(shareToken, payId, friendIds, totalAmount, description, customAmounts) {
  try {
    if (!friendIds || !friendIds.length) return _fail('At least one person is required');
    var conn = _dbConn();
    try {
      var eventId = _shareEventId(shareToken, true, conn);
      if (!eventId) return _fail('This share link cannot make changes');
      var transactionId = _dbInsertTransactionWithSplits(eventId, payId, friendIds, totalAmount, description, customAmounts, null, conn);
      return { success: true, transactionId: transactionId };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function updateDetailViaShare(shareToken, transactionId, payId, friendIds, totalAmount, description, customAmounts) {
  try {
    if (!friendIds || !friendIds.length) return _fail('At least one person is required');
    var conn = _dbConn();
    try {
      var eventId = _shareEventId(shareToken, true, conn);
      if (!eventId) return _fail('This share link cannot make changes');
      if (_dbGetTransactionEventId(transactionId, conn) !== eventId) return _fail('Transaction not found');

      _dbUpdateTransactionWithSplits(transactionId, payId, friendIds, totalAmount, description, customAmounts, null, conn);
      return { success: true, transactionId: transactionId };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

// deleteDetailViaShare / deleteTransactionSlipViaShare were removed on purpose
// (not merely hidden client-side): no share link, editable or not, may ever
// delete a transaction or a saved photo. Add/Edit stays available below.

function _saveUploadedSlip(transactionId, slip, createdBy, connOpt) {
  var uploaded = _uploadSlipToDrive(slip);
  var id = _dbInsertReceipt(transactionId, uploaded.fileId, createdBy, connOpt);
  return { success: true, id: id, slip: uploaded.fileId, slipHi: uploaded.fileId };
}

function uploadTransactionSlipViaShare(shareToken, transactionId, slip) {
  try {
    if (!slip) return _fail('No photo provided');
    var conn = _dbConn();
    try {
      var eventId = _shareEventId(shareToken, true, conn);
      if (!eventId) return _fail('This share link cannot make changes');
      if (_dbGetTransactionEventId(transactionId, conn) !== eventId) return _fail('Transaction not found');
      return _saveUploadedSlip(transactionId, slip, null, conn);
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

// Always adds a new photo (a transaction can have several) - returns its id
// so the client can target it with deleteTransactionSlip later.
function uploadTransactionSlip(token, transactionId, slip) {
  try {
    var user = requireAuth(token);
    if (!slip) return _fail('No photo provided');
    var conn = _dbConn();
    try {
      var eventId = _dbGetTransactionEventId(transactionId, conn);
      if (!eventId) return _fail('Transaction not found');
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Transaction not found');
      return _saveUploadedSlip(transactionId, slip, user.id, conn);
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function deleteTransactionSlip(token, transactionId, slipId) {
  try {
    var user = requireAuth(token);
    var conn = _dbConn();
    try {
      var eventId = _dbGetTransactionEventId(transactionId, conn);
      if (!eventId) return _fail('Transaction not found');
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Transaction not found');

      var fileId = _dbDeleteReceipt(transactionId, slipId, conn);
      if (fileId === null) return _fail('Photo not found');
      _deleteSlipFile(fileId);
      return { success: true };
    } finally { conn.close(); }
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

function _computeSettlementsWithPaid(detailRows, friendMap, eventId, connOpt) {
  var settlements = _computeSettlements(detailRows, friendMap);
  var paidSet = _dbGetSettlementsByEvent(eventId, connOpt);
  settlements.forEach(function (s) { s.paid = !!paidSet[_settleKey(s.from, s.to, s.amount)] });
  return settlements;
}

function _markSettlementPaid(eventId, fromId, toId, amount, paid, actorId, connOpt) {
  if (paid) _dbUpsertSettlement(eventId, fromId, toId, amount, actorId, connOpt);
  else _dbDeleteSettlement(eventId, fromId, toId, connOpt);
  return { success: true };
}

function markSettlementPaid(token, eventId, fromId, toId, amount, paid) {
  try {
    var user = requireAuth(token);
    var conn = _dbConn();
    try {
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Event not found');
      return _markSettlementPaid(eventId, fromId, toId, amount, paid, user.id, conn);
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

// Share-link 'edit' permission covers transaction CRUD already - marking a
// settlement/transaction paid is the same tier of access, unlike 'view'.
function markSettlementPaidViaShare(shareToken, fromId, toId, amount, paid) {
  try {
    var conn = _dbConn();
    try {
      var eventId = _shareEventId(shareToken, true, conn);
      if (!eventId) return _fail('This share link cannot make changes');
      return _markSettlementPaid(eventId, fromId, toId, amount, paid, null, conn);
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

// Marks a single transaction as already settled - it's then excluded from
// settlement math app-wide (see transaction.excluded_at, consumed by
// getHomeData/_buildDetailPayload/getSummary) instead of just noting a net
// debt as paid.
function markTransactionPaid(token, eventId, transactionId, paid) {
  try {
    var user = requireAuth(token);
    var conn = _dbConn();
    try {
      if (!_eventOwnedBy(eventId, user.id, conn)) return _fail('Event not found');
      _dbSetTransactionExcluded(transactionId, paid, user.id, conn);
      return { success: true };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function markTransactionPaidViaShare(shareToken, transactionId, paid) {
  try {
    var conn = _dbConn();
    try {
      var eventId = _shareEventId(shareToken, true, conn);
      if (!eventId) return _fail('This share link cannot make changes');
      _dbSetTransactionExcluded(transactionId, paid, null, conn);
      return { success: true };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function getSummary(token, eventId) {
  try {
    var user = requireAuth(token);
    var conn = _dbConn();
    try {
      var friendMap = {};
      _dbGetFriendsByAccount(user.id, conn).forEach(function (f) { friendMap[f.id] = f.name });

      var rows = [];
      _dbGetTransactionsByEvent(eventId, conn).forEach(function (tx) {
        if (!tx.excluded) tx.splits.forEach(function (s) { rows.push({ payId: tx.payId, friendId: s.friendId, amount: s.amount }) });
      });

      return { success: true, settlements: _computeSettlementsWithPaid(rows, friendMap, eventId, conn) };
    } finally { conn.close(); }
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
    return { success: true, accounts: _dbGetAllAccounts() };
  } catch (e) {
    return _fail(e);
  }
}

function updateAccountStatus(token, accountId, status) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return _fail('Forbidden');
    if (user.id === accountId) return _fail('Cannot disable your own account');
    var conn = _dbConn();
    try {
      if (!_dbGetAccountById(accountId, conn)) return _fail('Account not found');
      _dbUpdateAccountStatus(accountId, status, conn);
      return { success: true };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function updateAccountRole(token, accountId, role) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return _fail('Forbidden');
    if (user.id === accountId) return _fail('Cannot change your own role');
    var conn = _dbConn();
    try {
      if (!_dbGetAccountById(accountId, conn)) return _fail('Account not found');
      _dbUpdateAccountRole(accountId, role, conn);
      return { success: true };
    } finally { conn.close(); }
  } catch (e) {
    return _fail(e);
  }
}

function deleteAccount(token, accountId) {
  try {
    var user = requireAuth(token);
    if (user.role !== 'admin') return _fail('Forbidden');
    if (user.id === accountId) return _fail('Cannot delete your own account');
    var conn = _dbConn();
    try {
      if (!_dbGetAccountById(accountId, conn)) return _fail('Account not found');
      _dbDeleteAccount(accountId, conn);
      return { success: true };
    } finally { conn.close(); }
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
    var newSplits = _parseSplitsForMigration(row[6]);
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

// ----------------------------------------------------------------
// ONE-TIME MIGRATION: account/friend move from the Accounts/Friends sheets
// into Postgres (`account`/`friend`). Not client-callable - run manually
// from the Apps Script editor, once, then verify the log before deploying
// the rest of this file.
//
// account/friend get brand-new BIGINT ids in Postgres, so every other
// sheet that references the OLD uuid ids (Events.accountId, Details.payId
// + the splits JSON, EventFriends.friendId, SettlementPayments.fromId/toId)
// gets rewritten in place to point at the new ids instead. Sessions is
// cleared rather than remapped (its userInfo JSON embeds the old id too,
// and a session is cheap to just re-issue) - everyone gets logged out once.
//
// Safety: takes a full spreadsheet backup before touching anything, and
// never deletes the original Accounts/Friends sheets - just renames them
// aside. Refuses to run if `account` already has rows, so it can't be
// run twice by accident.
// ----------------------------------------------------------------

function migrateAccountsFriendsToDb() {
  var existing = _dbGetAllAccounts();
  if (existing.length) {
    throw new Error('Aborted: account table already has ' + existing.length + ' row(s). This migration only runs once, on an empty table.');
  }

  var ss = getSpreadsheet();
  var backupName = SPREADSHEET_NAME + '_backup_before_pg_migration_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  var backupSS = ss.copy(backupName);
  Logger.log('Backup created: ' + backupSS.getUrl());

  // --- accounts ---
  var acSheet = ss.getSheetByName('Accounts');
  var acData = acSheet.getDataRange().getValues();
  var accountIdMap = {}; // old uuid -> new bigint id (as string)
  for (var i = 1; i < acData.length; i++) {
    var r = acData[i];
    var oldId = r[0], displayName = r[1], username = r[2], passwordHash = r[3],
        firstLogin = r[4], lastLogin = r[5], role = r[6], status = r[7] || 'active',
        email = r[8] || '', photo = r[9] || '';
    var newId = _dbInsertAccountMigrated(displayName, username, passwordHash, role, status, firstLogin, lastLogin, email, photo);
    accountIdMap[oldId] = newId;
  }
  Logger.log('Migrated ' + Object.keys(accountIdMap).length + ' accounts.');

  // --- friends ---
  var frSheet = ss.getSheetByName('Friends');
  var frData = frSheet.getDataRange().getValues();
  var friendIdMap = {}; // old uuid -> new bigint id (as string)
  for (var i = 1; i < frData.length; i++) {
    var r = frData[i];
    var oldId = r[0], oldAccountId = r[1], name = r[2], isSelf = r[3] === 'true' || r[3] === true;
    var newAccountId = accountIdMap[oldAccountId];
    if (!newAccountId) { Logger.log('WARNING: friend ' + oldId + ' (' + name + ') references unknown account ' + oldAccountId + ' - skipped.'); continue; }
    var newId = _dbInsertFriend(newAccountId, name, isSelf);
    friendIdMap[oldId] = newId;
  }
  Logger.log('Migrated ' + Object.keys(friendIdMap).length + ' friends.');

  // --- rewrite Events.accountId ---
  var evSheet = ss.getSheetByName('Events');
  var evData = evSheet.getDataRange().getValues();
  var evUnmapped = 0;
  for (var i = 1; i < evData.length; i++) {
    var mapped = accountIdMap[evData[i][2]];
    if (mapped) evData[i][2] = mapped; else evUnmapped++;
  }
  if (evData.length > 1) evSheet.getRange(1, 1, evData.length, evData[0].length).setValues(evData);
  Logger.log('Rewrote Events.accountId for ' + (evData.length - 1) + ' rows (' + evUnmapped + ' unmapped).');

  // --- rewrite Details.payId + splits JSON keys ---
  var dtSheet = ss.getSheetByName('Details');
  var dtData = dtSheet.getDataRange().getValues();
  var dtUnmapped = 0;
  for (var i = 1; i < dtData.length; i++) {
    var payMapped = friendIdMap[dtData[i][2]];
    if (payMapped) dtData[i][2] = payMapped; else dtUnmapped++;
    var splits = _parseSplitsForMigration(dtData[i][6]);
    var newSplits = {};
    Object.keys(splits).forEach(function (fid) {
      var mapped = friendIdMap[fid];
      newSplits[mapped || fid] = splits[fid];
      if (!mapped) dtUnmapped++;
    });
    dtData[i][6] = JSON.stringify(newSplits);
  }
  if (dtData.length > 1) dtSheet.getRange(1, 1, dtData.length, dtData[0].length).setValues(dtData);
  Logger.log('Rewrote Details.payId/splits for ' + (dtData.length - 1) + ' rows (' + dtUnmapped + ' unmapped ids).');

  // --- rewrite EventFriends.friendId ---
  var efSheet = ss.getSheetByName('EventFriends');
  var efData = efSheet.getDataRange().getValues();
  var efUnmapped = 0;
  for (var i = 1; i < efData.length; i++) {
    var mapped = friendIdMap[efData[i][2]];
    if (mapped) efData[i][2] = mapped; else efUnmapped++;
  }
  if (efData.length > 1) efSheet.getRange(1, 1, efData.length, efData[0].length).setValues(efData);
  Logger.log('Rewrote EventFriends.friendId for ' + (efData.length - 1) + ' rows (' + efUnmapped + ' unmapped).');

  // --- rewrite SettlementPayments.fromId/toId ---
  var spSheet = ss.getSheetByName('SettlementPayments');
  var spData = spSheet.getDataRange().getValues();
  var spUnmapped = 0;
  for (var i = 1; i < spData.length; i++) {
    var fromMapped = friendIdMap[spData[i][2]], toMapped = friendIdMap[spData[i][3]];
    if (fromMapped) spData[i][2] = fromMapped; else spUnmapped++;
    if (toMapped) spData[i][3] = toMapped; else spUnmapped++;
  }
  if (spData.length > 1) spSheet.getRange(1, 1, spData.length, spData[0].length).setValues(spData);
  Logger.log('Rewrote SettlementPayments.fromId/toId for ' + (spData.length - 1) + ' rows (' + spUnmapped + ' unmapped).');

  // --- sessions: clear rather than remap (userInfo JSON embeds the old id too) ---
  var sessSheet = ss.getSheetByName('Sessions');
  var sessRows = sessSheet.getLastRow() - 1;
  if (sessRows > 0) sessSheet.deleteRows(2, sessRows);
  Logger.log('Cleared ' + Math.max(sessRows, 0) + ' session(s) - everyone will need to log in again.');

  // --- keep the old sheets, just rename them out of the way ---
  acSheet.setName('Accounts_v1_backup_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss'));
  frSheet.setName('Friends_v1_backup_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss'));

  var summary = {
    success: true,
    accountsMigrated: Object.keys(accountIdMap).length,
    friendsMigrated: Object.keys(friendIdMap).length,
    unmappedIdsFound: evUnmapped + dtUnmapped + efUnmapped + spUnmapped,
    sessionsCleared: Math.max(sessRows, 0),
    backupUrl: backupSS.getUrl()
  };
  Logger.log(JSON.stringify(summary, null, 2));
  return summary;
}

// Migration-only insert - preserves the real historical first_login/
// last_login timestamps and email/photo instead of defaulting them, unlike
// the normal _dbInsertAccount() new-registration path.
function _dbInsertAccountMigrated(name, username, passwordHash, role, status, firstLoginIso, lastLoginIso, email, photo) {
  var conn = _dbConn();
  try {
    var stmt = conn.prepareStatement(
      'INSERT INTO account (name, username, password_hash, role, status, email, photo, first_login_at, last_login_at) ' +
      'VALUES (?, ?, ?, ?, ?, ?, ?, ?::timestamptz, ?::timestamptz) RETURNING id'
    );
    stmt.setString(1, name);
    stmt.setString(2, username);
    stmt.setString(3, passwordHash);
    stmt.setInt(4, _roleToDb(role));
    stmt.setInt(5, _statusToDb(status));
    stmt.setString(6, email || '');
    stmt.setString(7, photo || '');
    stmt.setString(8, firstLoginIso || null);
    stmt.setString(9, lastLoginIso || null);
    var rs = stmt.executeQuery();
    rs.next();
    var id = String(rs.getLong(1));
    rs.close(); stmt.close();
    return id;
  } finally { conn.close(); }
}

// ----------------------------------------------------------------
// ONE-TIME MIGRATION: everything else (Events, Details+splits,
// EventFriends, EventShares, SettlementPayments, TransactionPayments,
// TransactionSlips, Sessions) moves from Sheets into Postgres
// (event/transaction/split/participant/settlement/receipt/session). Not
// client-callable - run manually from the Apps Script editor, once, after
// migrateAccountsFriendsToDb() has already run (this reuses the bigint
// account/friend ids that migration already wrote back into these sheets).
//
// Everything below shares ONE connection for the whole run (prepare each
// statement once, loop rows through it) - a first attempt that opened a
// fresh connection per row hit Apps Script's 6-minute execution cap partway
// through Details, since each Jdbc.getConnection() pays a real network
// handshake to Neon.
//
// Safety: takes a full spreadsheet backup before touching anything, and
// never deletes the original sheets - just renames them aside. Refuses to
// run if `event` already has rows, so it can't be run twice by accident.
// ----------------------------------------------------------------

// Migration-only: same JSON tolerance as the old _parseSplits (retired from
// the live app now that splits are real rows).
function _parseSplitsForMigration(json) {
  try {
    var o = JSON.parse(json || '{}');
    return (o && typeof o === 'object') ? o : {};
  } catch (e) {
    return {};
  }
}

function migrateRestToDb() {
  var conn = _dbConn();
  try {
    // Guard: abort if `event` already has rows (this only runs once).
    var cStmt = conn.prepareStatement('SELECT count(*) AS c FROM event');
    var cRs = cStmt.executeQuery();
    cRs.next();
    var eventCount = cRs.getInt('c');
    cRs.close(); cStmt.close();
    if (eventCount > 0) {
      throw new Error('Aborted: event table already has ' + eventCount + ' row(s). This migration only runs once, on an empty table.');
    }

    var ss = getSpreadsheet();
    var backupName = SPREADSHEET_NAME + '_backup_before_rest_migration_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    var backupSS = ss.copy(backupName);
    Logger.log('Backup created: ' + backupSS.getUrl());

    // Every friend_id reference below (payer_id, split, participant,
    // settlement from/to) gets checked against this set first - the prior
    // account/friend migration logged one EventFriends row with a friendId
    // that never mapped to a real friend, and every friend_id column here
    // has a hard FK, so a single bad id would otherwise abort the whole run.
    var validFriendIds = {};
    var vfStmt = conn.prepareStatement('SELECT id FROM friend');
    var vfRs = vfStmt.executeQuery();
    while (vfRs.next()) validFriendIds[String(vfRs.getLong('id'))] = true;
    vfRs.close(); vfStmt.close();

    // --- events ---
    var evData = ss.getSheetByName('Events').getDataRange().getValues();
    var eventIdMap = {}; // old uuid -> new bigint id (as string)
    var evStmt = conn.prepareStatement(
      'INSERT INTO event (account_id, name, icon, is_active, cre_at, mod_at) VALUES (?, ?, ?, ?, ?::timestamptz, ?::timestamptz) RETURNING id'
    );
    for (var i = 1; i < evData.length; i++) {
      var r = evData[i];
      evStmt.setLong(1, parseInt(r[2], 10));
      evStmt.setString(2, r[1]);
      evStmt.setString(3, r[5] || '');
      evStmt.setBoolean(4, r[4] !== false);
      evStmt.setString(5, r[3]);
      evStmt.setString(6, r[3]);
      var evRs = evStmt.executeQuery();
      evRs.next();
      eventIdMap[r[0]] = String(evRs.getLong(1));
      evRs.close();
    }
    evStmt.close();
    Logger.log('Migrated ' + Object.keys(eventIdMap).length + ' events.');

    // --- event shares -> event.share_token/share_perm/share_cre_at ---
    var shData = ss.getSheetByName('EventShares').getDataRange().getValues();
    var sharesApplied = 0, sharesUnmapped = 0;
    var shStmt = conn.prepareStatement('UPDATE event SET share_token = ?, share_perm = ?, share_cre_at = ?::timestamptz WHERE id = ?');
    for (var i = 1; i < shData.length; i++) {
      var r = shData[i];
      var newEventId = eventIdMap[r[0]];
      if (!newEventId) { sharesUnmapped++; continue; }
      shStmt.setString(1, r[1]);
      shStmt.setInt(2, _sharePermToDb(r[3] === 'edit' ? 'edit' : 'view'));
      shStmt.setString(3, r[2]);
      shStmt.setLong(4, parseInt(newEventId, 10));
      shStmt.executeUpdate();
      sharesApplied++;
    }
    shStmt.close();
    Logger.log('Applied ' + sharesApplied + ' event shares (' + sharesUnmapped + ' unmapped).');

    // --- transactions + splits ---
    var dtData = ss.getSheetByName('Details').getDataRange().getValues();
    var paidData = ss.getSheetByName('TransactionPayments').getDataRange().getValues();
    var excludedAtByTx = {};
    for (var i = 1; i < paidData.length; i++) excludedAtByTx[paidData[i][0]] = paidData[i][2];

    var transactionIdMap = {}; // old uuid -> new bigint id (as string)
    var splitsMigrated = 0, splitsUnmapped = 0;
    var txStmt = conn.prepareStatement(
      'INSERT INTO transaction (event_id, payer_id, amount, description, excluded_at, cre_at, mod_at) ' +
      'VALUES (?, ?, ?, ?, ?::timestamptz, ?::timestamptz, ?::timestamptz) RETURNING id'
    );
    var spStmt = conn.prepareStatement('INSERT INTO split (transaction_id, friend_id, amount) VALUES (?, ?, ?)');
    for (var i = 1; i < dtData.length; i++) {
      var r = dtData[i];
      var newEventId = eventIdMap[r[1]];
      if (!newEventId) { Logger.log('WARNING: transaction ' + r[0] + ' references unknown event ' + r[1] + ' - skipped.'); continue; }
      if (!validFriendIds[r[2]]) { Logger.log('WARNING: transaction ' + r[0] + ' has unknown payer friendId ' + r[2] + ' - skipped.'); continue; }
      var excludedAtIso = excludedAtByTx.hasOwnProperty(r[0]) ? excludedAtByTx[r[0]] : null;
      txStmt.setLong(1, parseInt(newEventId, 10));
      txStmt.setLong(2, parseInt(r[2], 10));
      txStmt.setDouble(3, parseFloat(r[3]));
      txStmt.setString(4, r[4] || '');
      txStmt.setString(5, excludedAtIso);
      txStmt.setString(6, r[5]);
      txStmt.setString(7, r[5]);
      var txRs = txStmt.executeQuery();
      txRs.next();
      var newTxId = String(txRs.getLong(1));
      txRs.close();
      transactionIdMap[r[0]] = newTxId;

      var splits = _parseSplitsForMigration(r[6]);
      Object.keys(splits).forEach(function (fid) {
        if (!validFriendIds[fid]) { Logger.log('WARNING: transaction ' + r[0] + ' has unknown split friendId ' + fid + ' - skipped.'); splitsUnmapped++; return; }
        spStmt.setLong(1, parseInt(newTxId, 10));
        spStmt.setLong(2, parseInt(fid, 10));
        spStmt.setDouble(3, parseFloat(splits[fid]));
        spStmt.executeUpdate();
        splitsMigrated++;
      });
    }
    txStmt.close(); spStmt.close();
    Logger.log('Migrated ' + Object.keys(transactionIdMap).length + ' transactions, ' + splitsMigrated + ' splits (' + splitsUnmapped + ' unmapped).');

    // --- participants ---
    var efData = ss.getSheetByName('EventFriends').getDataRange().getValues();
    var participantsMigrated = 0, participantsUnmapped = 0;
    var pStmt = conn.prepareStatement(
      'INSERT INTO participant (event_id, friend_id, cre_at, mod_at) VALUES (?, ?, ?::timestamptz, ?::timestamptz) ON CONFLICT (event_id, friend_id) DO NOTHING'
    );
    for (var i = 1; i < efData.length; i++) {
      var r = efData[i];
      var newEventId = eventIdMap[r[1]];
      if (!newEventId) { participantsUnmapped++; continue; }
      if (!validFriendIds[r[2]]) { Logger.log('WARNING: EventFriends row references unknown friendId ' + r[2] + ' - skipped.'); participantsUnmapped++; continue; }
      pStmt.setLong(1, parseInt(newEventId, 10));
      pStmt.setLong(2, parseInt(r[2], 10));
      pStmt.setString(3, r[3]);
      pStmt.setString(4, r[3]);
      pStmt.executeUpdate();
      participantsMigrated++;
    }
    pStmt.close();
    Logger.log('Migrated ' + participantsMigrated + ' participants (' + participantsUnmapped + ' unmapped).');

    // --- settlements ---
    var spData2 = ss.getSheetByName('SettlementPayments').getDataRange().getValues();
    var settlementsMigrated = 0, settlementsUnmapped = 0;
    var setStmt = conn.prepareStatement(
      'INSERT INTO settlement (event_id, from_id, to_id, amount, settled_at) VALUES (?, ?, ?, ?, ?::timestamptz) ' +
      'ON CONFLICT (event_id, from_id, to_id) DO UPDATE SET amount = EXCLUDED.amount, settled_at = EXCLUDED.settled_at'
    );
    for (var i = 1; i < spData2.length; i++) {
      var r = spData2[i];
      var newEventId = eventIdMap[r[1]];
      if (!newEventId) { settlementsUnmapped++; continue; }
      if (!validFriendIds[r[2]] || !validFriendIds[r[3]]) { Logger.log('WARNING: SettlementPayments row references unknown friendId (' + r[2] + '/' + r[3] + ') - skipped.'); settlementsUnmapped++; continue; }
      setStmt.setLong(1, parseInt(newEventId, 10));
      setStmt.setLong(2, parseInt(r[2], 10));
      setStmt.setLong(3, parseInt(r[3], 10));
      setStmt.setDouble(4, parseFloat(r[4]));
      setStmt.setString(5, r[5]);
      setStmt.executeUpdate();
      settlementsMigrated++;
    }
    setStmt.close();
    Logger.log('Migrated ' + settlementsMigrated + ' settlements (' + settlementsUnmapped + ' unmapped).');

    // --- receipts (TransactionSlips) - legacy base64 rows get uploaded to Drive first ---
    var slData = ss.getSheetByName('TransactionSlips').getDataRange().getValues();
    var receiptsMigrated = 0, receiptsUploadedFromLegacy = 0, receiptsUnmapped = 0;
    var rcStmt = conn.prepareStatement('INSERT INTO receipt (transaction_id, file_id, cre_at, mod_at) VALUES (?, ?, ?::timestamptz, ?::timestamptz)');
    for (var i = 1; i < slData.length; i++) {
      var r = slData[i];
      var newTxId = transactionIdMap[r[0]];
      if (!newTxId) { receiptsUnmapped++; continue; }
      var fileId = r[5];
      if (!fileId) {
        // Legacy row: 'slip' itself holds the original base64 data URI - upload it now to mint a real fileId.
        var uploaded = _uploadSlipToDrive(r[1]);
        fileId = uploaded.fileId;
        receiptsUploadedFromLegacy++;
      }
      rcStmt.setLong(1, parseInt(newTxId, 10));
      rcStmt.setString(2, fileId);
      rcStmt.setString(3, r[2]);
      rcStmt.setString(4, r[2]);
      rcStmt.executeUpdate();
      receiptsMigrated++;
    }
    rcStmt.close();
    Logger.log('Migrated ' + receiptsMigrated + ' receipts (' + receiptsUploadedFromLegacy + ' uploaded from legacy base64, ' + receiptsUnmapped + ' unmapped).');

    // --- sessions: clear rather than remap (userInfo JSON embeds account id, cheap to reissue) ---
    var sessSheet = ss.getSheetByName('Sessions');
    var sessRows = sessSheet.getLastRow() - 1;
    if (sessRows > 0) sessSheet.deleteRows(2, sessRows);
    Logger.log('Cleared ' + Math.max(sessRows, 0) + ' legacy session row(s) from the sheet.');

    // --- keep the old sheets, just rename them out of the way ---
    var suffix = '_v1_backup_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    ['Events', 'Details', 'EventFriends', 'EventShares', 'SettlementPayments', 'TransactionPayments', 'TransactionSlips', 'Sessions'].forEach(function (name) {
      var sheet = ss.getSheetByName(name);
      if (sheet) sheet.setName(name + suffix);
    });

    var summary = {
      success: true,
      eventsMigrated: Object.keys(eventIdMap).length,
      transactionsMigrated: Object.keys(transactionIdMap).length,
      splitsMigrated: splitsMigrated,
      participantsMigrated: participantsMigrated,
      settlementsMigrated: settlementsMigrated,
      receiptsMigrated: receiptsMigrated,
      receiptsUploadedFromLegacy: receiptsUploadedFromLegacy,
      unmappedIdsFound: sharesUnmapped + splitsUnmapped + participantsUnmapped + settlementsUnmapped + receiptsUnmapped,
      sessionsCleared: Math.max(sessRows, 0),
      backupUrl: backupSS.getUrl()
    };
    Logger.log(JSON.stringify(summary, null, 2));
    return summary;
  } finally {
    conn.close();
  }
}
