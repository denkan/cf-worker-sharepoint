-- Mirror of the SharePoint document library's file/folder metadata, kept
-- up to date by the delta sync. The whole library is mirrored (not only the
-- configured root folder) so moves into and out of the root are handled
-- correctly; every API query is restricted to the root's subtree.

CREATE TABLE items (
  id                TEXT PRIMARY KEY,
  parent_id         TEXT,
  name              TEXT NOT NULL,
  name_lower        TEXT NOT NULL,
  -- Drive-relative path: '' for the library root, otherwise '/Folder/file.pdf'.
  -- NULL while the parent is not yet known (fixed up at the end of each sync).
  path              TEXT,
  path_lower        TEXT,
  type              TEXT NOT NULL CHECK (type IN ('file', 'folder')),
  extension         TEXT,
  mime_type         TEXT,
  size              INTEGER,
  created_at        TEXT,
  modified_at       TEXT,
  modified_by       TEXT,
  etag              TEXT,
  ctag              TEXT,
  web_url           TEXT,
  -- Sync generation that last saw this item; a full (re)sync removes older ones.
  gen               INTEGER NOT NULL
);

CREATE INDEX idx_items_parent ON items (parent_id);
CREATE INDEX idx_items_path ON items (path_lower);
CREATE INDEX idx_items_gen ON items (gen);
CREATE INDEX idx_items_orphans ON items (parent_id) WHERE path IS NULL;

-- Full-text index over file and folder names.
CREATE VIRTUAL TABLE items_fts USING fts5 (
  name,
  content = 'items',
  content_rowid = 'rowid',
  tokenize = 'unicode61 remove_diacritics 2'
);

CREATE TRIGGER items_fts_insert AFTER INSERT ON items BEGIN
  INSERT INTO items_fts (rowid, name) VALUES (new.rowid, new.name);
END;

CREATE TRIGGER items_fts_delete AFTER DELETE ON items BEGIN
  INSERT INTO items_fts (items_fts, rowid, name) VALUES ('delete', old.rowid, old.name);
END;

CREATE TRIGGER items_fts_update AFTER UPDATE OF name ON items BEGIN
  INSERT INTO items_fts (items_fts, rowid, name) VALUES ('delete', old.rowid, old.name);
  INSERT INTO items_fts (rowid, name) VALUES (new.rowid, new.name);
END;

-- Sync state, resolved site/drive ids and the sync lock.
CREATE TABLE meta (
  key   TEXT PRIMARY KEY,
  value TEXT
);

-- Short-lived zip download manifests (POST /api/zip → GET /api/zip/:id).
CREATE TABLE zip_manifests (
  id          TEXT PRIMARY KEY,
  name        TEXT NOT NULL,
  entries     TEXT NOT NULL,
  file_count  INTEGER NOT NULL,
  total_bytes INTEGER NOT NULL,
  created_at  INTEGER NOT NULL,
  expires_at  INTEGER NOT NULL
);

CREATE INDEX idx_zip_manifests_expires ON zip_manifests (expires_at);
