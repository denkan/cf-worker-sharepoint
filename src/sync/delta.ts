/**
 * Delta sync: mirrors the document library's metadata into D1.
 *
 * SharePoint only supports delta on the drive root, so the whole library is
 * mirrored and paths are computed locally from parent ids (delta responses
 * carry no parentReference.path). The sync is resumable: it works within a
 * time budget and stores the continuation link between runs.
 */
import { configFingerprint, getConfig, type Env } from "../env";
import { getMeta, getMetaMany, setMeta, setMetaStmt } from "../db/meta";
import { GraphClient, GraphError } from "../graph/client";
import type { DeltaPage, DriveItem } from "../graph/types";
import { clearDriveContextCache, getStoredDriveContext, resolveDriveContext, type DriveContext } from "../sharepoint/drive";
import { codePointLength, lowerSegment, subtreeRange } from "../sharepoint/paths";

const DELTA_SELECT = [
  "id",
  "name",
  "size",
  "file",
  "folder",
  "package",
  "root",
  "deleted",
  "parentReference",
  "createdDateTime",
  "lastModifiedDateTime",
  "lastModifiedBy",
  "eTag",
  "cTag",
  "webUrl",
].join(",");

const LOCK_TTL_MS = 5 * 60 * 1000;
const MAX_PAGES_PER_RUN = 150;

export const SYNC_META_KEYS = [
  "sync_gen",
  "sync_mode",
  "delta_link",
  "next_link",
  "index_ready",
  "last_sync_started_at",
  "last_sync_completed_at",
  "last_full_sync_completed_at",
  "last_sync_error",
  "last_sync_result",
] as const;

export interface SyncResult {
  status: "completed" | "partial" | "skipped" | "failed";
  mode?: "full" | "incremental";
  pages: number;
  upserted: number;
  deleted: number;
  durationMs: number;
  error?: string;
}

interface Known {
  path: string | null;
  pathLower: string | null;
  type: "file" | "folder";
}

interface Row {
  id: string;
  parentId: string | null;
  name: string;
  nameLower: string;
  type: "file" | "folder";
  extension: string | null;
  mimeType: string | null;
  size: number | null;
  createdAt: string | null;
  modifiedAt: string | null;
  modifiedBy: string | null;
  etag: string | null;
  ctag: string | null;
  webUrl: string | null;
}

export async function runSync(env: Env, opts: { timeBudgetMs?: number; maxPages?: number } = {}): Promise<SyncResult> {
  const config = getConfig(env);
  const started = Date.now();
  const budget = opts.timeBudgetMs ?? config.syncTimeBudgetMs;
  const result: SyncResult = { status: "completed", pages: 0, upserted: 0, deleted: 0, durationMs: 0 };

  const lock = await acquireLock(env.DB);
  if (!lock) return { ...result, status: "skipped", durationMs: Date.now() - started };

  try {
    const graph = new GraphClient(env);
    const drive = await ensureDriveContext(env, graph);
    const meta = await getMetaMany(env.DB, SYNC_META_KEYS);

    let gen = Number(meta.sync_gen ?? 0);
    let mode = (meta.sync_mode as "full" | "incremental" | null) ?? "full";
    let url: string;

    if (meta.next_link) {
      url = meta.next_link;
    } else if (meta.delta_link) {
      url = meta.delta_link;
      mode = "incremental";
    } else {
      ({ gen, mode, url } = startFullSync(drive, gen));
    }
    result.mode = mode;
    await setMeta(env.DB, {
      sync_gen: String(gen),
      sync_mode: mode,
      last_sync_started_at: new Date(started).toISOString(),
    });

    let done = false;
    const maxPages = opts.maxPages ?? MAX_PAGES_PER_RUN;
    while (!done && result.pages < maxPages && Date.now() - started < budget) {
      let page: DeltaPage;
      try {
        page = await graph.json<DeltaPage>(url);
      } catch (err) {
        if (err instanceof GraphError && err.graphStatus === 410) {
          // Delta token expired: enumerate everything again; stale rows are removed at the end.
          ({ gen, mode, url } = startFullSync(drive, gen));
          result.mode = mode;
          await setMeta(env.DB, { sync_gen: String(gen), sync_mode: mode, next_link: null, delta_link: null });
          continue;
        }
        throw err;
      }

      const applied = await applyDeltaPage(env.DB, page.value, gen);
      result.pages++;
      result.upserted += applied.upserted;
      result.deleted += applied.deleted;

      if (page["@odata.nextLink"]) {
        url = page["@odata.nextLink"];
        await env.DB.batch([setMetaStmt(env.DB, "next_link", url), renewLockStmt(env.DB, lock)]);
      } else if (page["@odata.deltaLink"]) {
        const stmts = [
          setMetaStmt(env.DB, "delta_link", page["@odata.deltaLink"]),
          setMetaStmt(env.DB, "next_link", null),
          setMetaStmt(env.DB, "sync_mode", "incremental"),
          setMetaStmt(env.DB, "last_sync_completed_at", new Date().toISOString()),
          setMetaStmt(env.DB, "last_sync_error", null),
        ];
        if (mode === "full") {
          const removed = await env.DB.prepare("DELETE FROM items WHERE gen < ?").bind(gen).run();
          result.deleted += removed.meta.changes ?? 0;
          stmts.push(
            setMetaStmt(env.DB, "index_ready", "1"),
            setMetaStmt(env.DB, "last_full_sync_completed_at", new Date().toISOString()),
          );
        }
        await env.DB.batch(stmts);
        done = true;
      } else {
        throw new Error("Delta response had neither nextLink nor deltaLink");
      }
    }

    await fixOrphans(env.DB);
    result.status = done ? "completed" : "partial";
  } catch (err) {
    result.status = "failed";
    result.error = err instanceof Error ? err.message : String(err);
    await setMeta(env.DB, { last_sync_error: `${new Date().toISOString()} ${result.error}` }).catch(() => {});
  } finally {
    result.durationMs = Date.now() - started;
    await setMeta(env.DB, { last_sync_result: JSON.stringify(result) }).catch(() => {});
    await releaseLock(env.DB, lock).catch(() => {});
  }
  return result;
}

function startFullSync(drive: DriveContext, gen: number) {
  return {
    gen: gen + 1,
    mode: "full" as const,
    url: `/drives/${encodeURIComponent(drive.driveId)}/root/delta?$select=${DELTA_SELECT}`,
  };
}

/** Resolve the drive on first run; wipe the mirror if site or library configuration changed. */
async function ensureDriveContext(env: Env, graph: GraphClient): Promise<DriveContext> {
  const stored = await getStoredDriveContext(env);
  if (stored) return stored;

  const previous = await getMeta(env.DB, "drive_context");
  if (previous && (JSON.parse(previous) as DriveContext).fingerprint !== configFingerprint(getConfig(env))) {
    await env.DB.batch([
      env.DB.prepare("DELETE FROM items"),
      env.DB.prepare(
        `DELETE FROM meta WHERE key IN (${SYNC_META_KEYS.map(() => "?").join(",")}, 'drive_context')`,
      ).bind(...SYNC_META_KEYS),
    ]);
    clearDriveContextCache();
  }
  return resolveDriveContext(env, graph);
}

// --- Applying a delta page ------------------------------------------------------

function toRow(item: DriveItem): Row {
  const name = (item.name ?? "").normalize("NFC");
  const isFolder = Boolean(item.folder || item.package || item.root);
  const dot = name.lastIndexOf(".");
  return {
    id: item.id,
    parentId: item.root ? null : (item.parentReference?.id ?? null),
    name,
    nameLower: lowerSegment(name),
    type: isFolder ? "folder" : "file",
    extension: !isFolder && dot > 0 ? name.slice(dot + 1).toLowerCase() : null,
    mimeType: item.file?.mimeType ?? null,
    size: typeof item.size === "number" ? item.size : null,
    createdAt: item.createdDateTime ?? null,
    modifiedAt: item.lastModifiedDateTime ?? null,
    modifiedBy: item.lastModifiedBy?.user?.displayName ?? item.lastModifiedBy?.application?.displayName ?? null,
    etag: item.eTag ?? null,
    ctag: item.cTag ?? null,
    webUrl: item.webUrl ?? null,
  };
}

export async function applyDeltaPage(db: D1Database, items: DriveItem[], gen: number) {
  if (items.length === 0) return { upserted: 0, deleted: 0 };

  // Everything we know about the items in this page and their parents.
  const lookupIds = new Set<string>();
  for (const item of items) {
    lookupIds.add(item.id);
    if (item.parentReference?.id) lookupIds.add(item.parentReference.id);
  }
  const known = new Map<string, Known>();
  const { results } = await db
    .prepare("SELECT id, path, path_lower, type FROM items WHERE id IN (SELECT value FROM json_each(?))")
    .bind(JSON.stringify([...lookupIds]))
    .all<{ id: string; path: string | null; path_lower: string | null; type: "file" | "folder" }>();
  for (const r of results) known.set(r.id, { path: r.path, pathLower: r.path_lower, type: r.type });

  const pathStmts: D1PreparedStatement[] = [];
  const subtreeDeletes: string[] = [];
  const upserts = new Map<string, Row>();
  const deletes = new Set<string>();

  /** Apply a path change to everything below oldLower, both in D1 and in `known`. */
  const reparent = (oldPath: string, oldLower: string, newPath: string | null, newLower: string | null) => {
    const [lo, hi] = subtreeRange(oldLower);
    if (newPath === null || newLower === null) {
      pathStmts.push(
        db.prepare("UPDATE items SET path = NULL, path_lower = NULL WHERE path_lower >= ? AND path_lower < ?").bind(lo, hi),
      );
    } else {
      pathStmts.push(
        db
          .prepare(
            "UPDATE items SET path = ?1 || substr(path, ?2), path_lower = ?3 || substr(path_lower, ?4) WHERE path_lower >= ?5 AND path_lower < ?6",
          )
          .bind(newPath, codePointLength(oldPath) + 1, newLower, codePointLength(oldLower) + 1, lo, hi),
      );
    }
    for (const entry of known.values()) {
      if (entry.pathLower && entry.path && entry.pathLower >= lo && entry.pathLower < hi) {
        if (newPath === null || newLower === null) {
          entry.path = entry.pathLower = null;
        } else {
          entry.path = newPath + entry.path.slice(oldPath.length);
          entry.pathLower = newLower + entry.pathLower.slice(oldLower.length);
        }
      }
    }
  };

  for (const item of items) {
    const previous = known.get(item.id);

    if (item.deleted) {
      upserts.delete(item.id);
      deletes.add(item.id);
      if (previous?.type === "folder" && previous.pathLower) subtreeDeletes.push(previous.pathLower);
      known.delete(item.id);
      continue;
    }
    if (!item.name && !item.root) continue; // Nothing usable in this entry

    const row = toRow(item);
    let path: string | null = null;
    let pathLower: string | null = null;
    if (item.root) {
      path = pathLower = "";
    } else if (row.parentId) {
      const parent = known.get(row.parentId);
      if (parent?.path != null && parent.pathLower != null) {
        path = `${parent.path}/${row.name}`;
        pathLower = `${parent.pathLower}/${row.nameLower}`;
      }
    }

    if (row.type === "folder" && previous?.path != null && previous.pathLower != null && previous.pathLower !== pathLower) {
      reparent(previous.path, previous.pathLower, path, pathLower);
    }

    deletes.delete(item.id);
    upserts.set(item.id, row);
    known.set(item.id, { path, pathLower, type: row.type });
  }

  const stmts: D1PreparedStatement[] = [...pathStmts];

  if (subtreeDeletes.length > 0 && subtreeDeletes.length <= 20) {
    for (const prefix of subtreeDeletes) {
      const [lo, hi] = subtreeRange(prefix);
      stmts.push(db.prepare("DELETE FROM items WHERE path_lower >= ? AND path_lower < ?").bind(lo, hi));
    }
  } else if (subtreeDeletes.length > 20) {
    stmts.push(
      db
        .prepare(
          "DELETE FROM items WHERE EXISTS (SELECT 1 FROM json_each(?) p WHERE items.path_lower >= p.value || '/' AND items.path_lower < p.value || '0')",
        )
        .bind(JSON.stringify(subtreeDeletes)),
    );
  }
  if (deletes.size > 0) {
    stmts.push(db.prepare("DELETE FROM items WHERE id IN (SELECT value FROM json_each(?))").bind(JSON.stringify([...deletes])));
  }

  const rows = [...upserts.values()].map((row) => {
    const k = known.get(row.id);
    return { ...row, path: k?.path ?? null, pathLower: k?.pathLower ?? null };
  });
  for (let i = 0; i < rows.length; i += 500) {
    stmts.push(
      db
        .prepare(
          `INSERT INTO items (id, parent_id, name, name_lower, path, path_lower, type, extension, mime_type, size,
                              created_at, modified_at, modified_by, etag, ctag, web_url, gen)
           SELECT j.value->>'id', j.value->>'parentId', j.value->>'name', j.value->>'nameLower', j.value->>'path',
                  j.value->>'pathLower', j.value->>'type', j.value->>'extension', j.value->>'mimeType', j.value->>'size',
                  j.value->>'createdAt', j.value->>'modifiedAt', j.value->>'modifiedBy', j.value->>'etag', j.value->>'ctag',
                  j.value->>'webUrl', ?2
           FROM json_each(?1) j WHERE true
           ON CONFLICT (id) DO UPDATE SET
             parent_id = excluded.parent_id, name = excluded.name, name_lower = excluded.name_lower,
             path = excluded.path, path_lower = excluded.path_lower, type = excluded.type,
             extension = excluded.extension, mime_type = excluded.mime_type, size = excluded.size,
             created_at = excluded.created_at, modified_at = excluded.modified_at, modified_by = excluded.modified_by,
             etag = excluded.etag, ctag = excluded.ctag, web_url = excluded.web_url, gen = excluded.gen`,
        )
        .bind(JSON.stringify(rows.slice(i, i + 500)), gen),
    );
  }

  if (stmts.length) await db.batch(stmts);
  return { upserted: rows.length, deleted: deletes.size };
}

/** Give a path to items whose parent arrived later than they did. */
export async function fixOrphans(db: D1Database, maxRounds = 50): Promise<number> {
  let fixed = 0;
  for (let round = 0; round < maxRounds; round++) {
    const { results } = await db
      .prepare(
        `SELECT c.id, c.name, c.name_lower, p.path AS parent_path, p.path_lower AS parent_path_lower
         FROM items c JOIN items p ON p.id = c.parent_id
         WHERE c.path IS NULL AND p.path IS NOT NULL
         LIMIT 1000`,
      )
      .all<{ id: string; name: string; name_lower: string; parent_path: string; parent_path_lower: string }>();
    if (results.length === 0) break;
    const updates = results.map((r) => ({
      id: r.id,
      path: `${r.parent_path}/${r.name}`,
      pathLower: `${r.parent_path_lower}/${r.name_lower}`,
    }));
    await db
      .prepare(
        "UPDATE items SET path = j.value->>'path', path_lower = j.value->>'pathLower' FROM json_each(?) j WHERE items.id = j.value->>'id'",
      )
      .bind(JSON.stringify(updates))
      .run();
    fixed += updates.length;
  }
  return fixed;
}

// --- Lock -------------------------------------------------------------------------

async function acquireLock(db: D1Database): Promise<string | null> {
  const token = crypto.randomUUID();
  const now = Date.now();
  const res = await db
    .prepare(
      `INSERT INTO meta (key, value) VALUES ('sync_lock', ?1)
       ON CONFLICT (key) DO UPDATE SET value = excluded.value
       WHERE meta.value IS NULL OR CAST(substr(meta.value, 1, instr(meta.value, '|') - 1) AS INTEGER) < ?2`,
    )
    .bind(`${now + LOCK_TTL_MS}|${token}`, now)
    .run();
  return (res.meta.changes ?? 0) > 0 ? token : null;
}

function renewLockStmt(db: D1Database, token: string): D1PreparedStatement {
  return db
    .prepare("UPDATE meta SET value = ?1 WHERE key = 'sync_lock' AND substr(value, instr(value, '|') + 1) = ?2")
    .bind(`${Date.now() + LOCK_TTL_MS}|${token}`, token);
}

async function releaseLock(db: D1Database, token: string): Promise<void> {
  await db
    .prepare("UPDATE meta SET value = NULL WHERE key = 'sync_lock' AND substr(value, instr(value, '|') + 1) = ?")
    .bind(token)
    .run();
}

export async function isSyncRunning(db: D1Database): Promise<boolean> {
  const value = await getMeta(db, "sync_lock");
  if (!value) return false;
  return Number(value.split("|")[0]) > Date.now();
}
