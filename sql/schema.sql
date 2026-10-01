-- ShareUp — PostgreSQL schema v3 (Neon)
-- Aligned to the audit pattern already applied to `account` in DBeaver:
-- singular table names, IDENTITY ids, timestamptz, and a standard
-- cre_at/mod_at/cre_by/mod_by audit quartet on every table.
--
-- Changed from v2:
-- - Tables renamed singular: account, friend, event, transaction, split,
--   settlement, participant, receipt, session.
-- - Every table gets cre_at/mod_at (timestamptz, auto-set) and cre_by/mod_by
--   (who did it, FK -> account.id, ON DELETE SET NULL so deleting a user
--   never blocks deleting their old edits). This replaces the old one-off
--   `created_at` column - cre_at covers the same fact.
-- - session is the one exception: it's a system-managed token, not a user-
--   edited row, so it keeps just cre_at/expires_at - no mod_at/cre_by/mod_by.
-- - account.role and account.status now carry more states than v2 had:
--     role:   0=user, 1=staff, 2=admin
--     status: 0=deleted (soft-delete), 1=disabled, 2=pending, 3=active (default)
--   status=0 means "deleted" accounts are a soft-delete (row stays, status
--   flips to 0) rather than an actual DELETE - Code.js's deleteAccount()
--   needs to become an UPDATE, not a DELETE, once this lands. staff (role=1)
--   and pending (status=2) aren't wired into the app yet - Code.js's
--   requireAuth/isAdmin-style checks only ever compared against user/admin
--   and active/disabled, so those two new states need actual behavior
--   defined before they mean anything at the app layer.
-- - account.name and event.name get a GIN trigram index (matching the
--   account_idx1 you already created) so an admin/owner search box can use
--   ILIKE '%text%' fast instead of a full scan - mirrors the search boxes
--   already in Admin_js.html and Home_js.html.
--
-- Trade-off worth flagging: split/participant were deliberately bare
-- (just their composite key) in v2 for minimalism. Adding the full audit
-- quartet to them is a deliberate choice to keep every table consistent,
-- at the cost of that minimalism - each split/participant row is now 5x
-- more columns for what's still just a fact.

CREATE EXTENSION IF NOT EXISTS pg_trgm;

CREATE TABLE account (
  id             BIGINT       GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  name           VARCHAR(255) NOT NULL,
  username       VARCHAR(255) NOT NULL UNIQUE,
  password_hash  TEXT         NOT NULL,
  first_login_at TIMESTAMPTZ,
  last_login_at  TIMESTAMPTZ,
  role           SMALLINT     NOT NULL DEFAULT 0 CHECK (role IN (0,1,2)),      -- 0=user, 1=staff, 2=admin
  status         SMALLINT     NOT NULL DEFAULT 3 CHECK (status IN (0,1,2,3)),  -- 0=deleted, 1=disabled, 2=pending, 3=active
  email          VARCHAR(255),
  photo          VARCHAR(255),
  cre_at         TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
  mod_at         TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
  cre_by         BIGINT       REFERENCES account(id) ON DELETE SET NULL,
  mod_by         BIGINT       REFERENCES account(id) ON DELETE SET NULL
);
CREATE INDEX account_idx1 ON account USING gin (name gin_trgm_ops);

CREATE TABLE friend (
  id         BIGINT       GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  account_id BIGINT       NOT NULL REFERENCES account(id) ON DELETE CASCADE,
  name       VARCHAR(255) NOT NULL,
  is_self    BOOLEAN      NOT NULL DEFAULT FALSE,
  cre_at     TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
  mod_at     TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
  cre_by     BIGINT       REFERENCES account(id) ON DELETE SET NULL,
  mod_by     BIGINT       REFERENCES account(id) ON DELETE SET NULL
);
CREATE INDEX friend_idx1 ON friend(account_id);
CREATE INDEX friend_idx2 ON friend USING gin (name gin_trgm_ops);

CREATE TABLE event (
  id            BIGINT       GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  account_id    BIGINT       NOT NULL REFERENCES account(id) ON DELETE CASCADE,  -- owner
  name          VARCHAR(255) NOT NULL,
  icon          VARCHAR(16),
  is_active     BOOLEAN      NOT NULL DEFAULT TRUE,
  share_token   VARCHAR(128) UNIQUE,      -- null = sharing off
  share_perm    SMALLINT     CHECK (share_perm IN (0,1)),  -- 0=view, 1=edit
  share_cre_at  TIMESTAMPTZ,
  cre_at        TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
  mod_at        TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
  cre_by        BIGINT       REFERENCES account(id) ON DELETE SET NULL,
  mod_by        BIGINT       REFERENCES account(id) ON DELETE SET NULL
);
CREATE INDEX event_idx1 ON event(account_id);
CREATE INDEX event_idx2 ON event USING gin (name gin_trgm_ops);

CREATE TABLE transaction (
  id          BIGINT        GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  event_id    BIGINT        NOT NULL REFERENCES event(id) ON DELETE CASCADE,
  payer_id    BIGINT        NOT NULL REFERENCES friend(id),
  amount      DECIMAL(12,2) NOT NULL CHECK (amount > 0),
  description VARCHAR(500)  NOT NULL DEFAULT '',
  excluded_at TIMESTAMPTZ,                 -- null = counted in settlement (replaces the old TransactionPayments sheet)
  cre_at      TIMESTAMPTZ   NOT NULL DEFAULT CURRENT_TIMESTAMP,
  mod_at      TIMESTAMPTZ   NOT NULL DEFAULT CURRENT_TIMESTAMP,
  cre_by      BIGINT        REFERENCES account(id) ON DELETE SET NULL,
  mod_by      BIGINT        REFERENCES account(id) ON DELETE SET NULL
);
CREATE INDEX transaction_idx1 ON transaction(event_id, cre_at DESC);
CREATE INDEX transaction_idx2 ON transaction(payer_id);

CREATE TABLE split (
  transaction_id BIGINT        NOT NULL REFERENCES transaction(id) ON DELETE CASCADE,
  friend_id      BIGINT        NOT NULL REFERENCES friend(id),
  amount         DECIMAL(12,2) NOT NULL CHECK (amount >= 0),
  cre_at         TIMESTAMPTZ   NOT NULL DEFAULT CURRENT_TIMESTAMP,
  mod_at         TIMESTAMPTZ   NOT NULL DEFAULT CURRENT_TIMESTAMP,
  cre_by         BIGINT        REFERENCES account(id) ON DELETE SET NULL,
  mod_by         BIGINT        REFERENCES account(id) ON DELETE SET NULL,
  PRIMARY KEY (transaction_id, friend_id)
);
CREATE INDEX split_idx1 ON split(friend_id);

CREATE TABLE participant (
  event_id  BIGINT      NOT NULL REFERENCES event(id) ON DELETE CASCADE,
  friend_id BIGINT      NOT NULL REFERENCES friend(id) ON DELETE CASCADE,
  cre_at    TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
  mod_at    TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
  cre_by    BIGINT      REFERENCES account(id) ON DELETE SET NULL,
  mod_by    BIGINT      REFERENCES account(id) ON DELETE SET NULL,
  PRIMARY KEY (event_id, friend_id)
);

CREATE TABLE settlement (
  id         BIGINT        GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  event_id   BIGINT        NOT NULL REFERENCES event(id) ON DELETE CASCADE,
  from_id    BIGINT        NOT NULL REFERENCES friend(id),
  to_id      BIGINT        NOT NULL REFERENCES friend(id),
  amount     DECIMAL(12,2) NOT NULL CHECK (amount > 0),
  settled_at TIMESTAMPTZ   NOT NULL,
  cre_at     TIMESTAMPTZ   NOT NULL DEFAULT CURRENT_TIMESTAMP,
  mod_at     TIMESTAMPTZ   NOT NULL DEFAULT CURRENT_TIMESTAMP,
  cre_by     BIGINT        REFERENCES account(id) ON DELETE SET NULL,
  mod_by     BIGINT        REFERENCES account(id) ON DELETE SET NULL,
  UNIQUE (event_id, from_id, to_id)
);

CREATE TABLE receipt (
  id             BIGINT       GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  transaction_id BIGINT       NOT NULL REFERENCES transaction(id) ON DELETE CASCADE,
  file_id        VARCHAR(128) NOT NULL,
  cre_at         TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
  mod_at         TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
  cre_by         BIGINT       REFERENCES account(id) ON DELETE SET NULL,
  mod_by         BIGINT       REFERENCES account(id) ON DELETE SET NULL
);
CREATE INDEX receipt_idx1 ON receipt(transaction_id);

-- Exception to the audit pattern: a session is issued and expires by
-- itself, no one "edits" one, so mod_at/cre_by/mod_by would always be
-- meaningless here.
CREATE TABLE session (
  token      VARCHAR(128) PRIMARY KEY,
  account_id BIGINT       NOT NULL REFERENCES account(id) ON DELETE CASCADE,
  info       JSONB        NOT NULL,
  cre_at     TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP,
  expires_at TIMESTAMPTZ  NOT NULL
);
CREATE INDEX session_idx1 ON session(account_id);
CREATE INDEX session_idx2 ON session(expires_at);

-- ------------------------------------------------------------------
-- Already applied directly in Neon during the account/friend migration
-- (account.photo -> TEXT, the two friend indexes above). Restated here,
-- guarded so this whole file stays safe to run end-to-end against a fresh
-- database - CREATE INDEX already has IF NOT EXISTS above; ALTER COLUMN
-- has no such guard, hence the DO block.
-- ------------------------------------------------------------------
DO $$
BEGIN
  IF EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_name = 'account' AND column_name = 'photo' AND data_type <> 'text'
  ) THEN
    ALTER TABLE account ALTER COLUMN photo TYPE TEXT;
  END IF;
END $$;
CREATE INDEX IF NOT EXISTS friend_idx1 ON friend(account_id);
CREATE INDEX IF NOT EXISTS friend_idx2 ON friend USING gin (name gin_trgm_ops);

-- receipt.file_id is NOT NULL - migrateRestToDb() satisfies this even for
-- legacy TransactionSlips rows that predate Drive storage (a blank fileId,
-- original photo inline as base64 in the 'slip' column) by uploading that
-- base64 blob to Drive during the migration to mint a real fileId first.
-- No schema change needed for that case.
