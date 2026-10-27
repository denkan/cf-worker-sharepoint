/**
 * Zip downloads. A manifest (which files go where in the archive) is built
 * from the D1 index; the archive itself is streamed in "store" mode (no
 * compression, Zip64 when needed) while each file is fetched from SharePoint,
 * so memory use stays flat regardless of archive size.
 */
import { downloadZip, predictLength } from "client-zip";
import { getConfig, type Env } from "../env";
import { ApiError } from "../errors";
import { listDescendants, rootPathLower, type ItemRow } from "../db/items";
import { GraphClient } from "../graph/client";
import type { DriveItem } from "../graph/types";
import type { DriveContext } from "../sharepoint/drive";
import { resolvePath } from "../sharepoint/resolve";
import { contentDisposition, fetchFileBody } from "./files";

export interface ZipEntry {
  /** Item id; absent for folder entries. */
  id?: string;
  /** Path inside the archive; folders end with "/". */
  name: string;
  size?: number;
  modified?: string | null;
}

export interface ZipManifest {
  id?: string;
  name: string;
  entries: ZipEntry[];
  fileCount: number;
  folderCount: number;
  totalBytes: number;
}

const MANIFEST_TTL_MS = 60 * 60 * 1000;

function sanitizeName(name: string): string {
  return name.replace(/[\\/:*?"<>|\x00-\x1f]/g, "_").trim() || "files";
}

/** Resolve the requested paths and expand folders into the list of zip entries. */
export async function buildManifest(
  env: Env,
  drive: DriveContext,
  paths: string[],
  requestedName?: string,
): Promise<ZipManifest> {
  const config = getConfig(env);
  const entries: ZipEntry[] = [];
  const usedNames = new Set<string>();
  const seenIds = new Set<string>();
  let fileCount = 0;
  let folderCount = 0;
  let totalBytes = 0;

  const uniqueTopName = (name: string) => {
    let candidate = name;
    const dot = name.lastIndexOf(".");
    const [base, ext] = dot > 0 ? [name.slice(0, dot), name.slice(dot)] : [name, ""];
    for (let n = 2; usedNames.has(candidate.toLowerCase()); n++) candidate = `${base} (${n})${ext}`;
    usedNames.add(candidate.toLowerCase());
    return candidate;
  };

  const addFile = (item: ItemRow, name: string) => {
    if (seenIds.has(item.id)) return;
    seenIds.add(item.id);
    fileCount++;
    totalBytes += item.size ?? 0;
    if (fileCount > config.zipMaxFiles) {
      throw new ApiError(413, "LIMIT_EXCEEDED", `A zip can contain at most ${config.zipMaxFiles} files`, {
        maxFiles: config.zipMaxFiles,
      });
    }
    if (totalBytes > config.zipMaxBytes) {
      throw new ApiError(413, "LIMIT_EXCEEDED", `A zip can be at most ${config.zipMaxBytes} bytes`, {
        maxBytes: config.zipMaxBytes,
      });
    }
    entries.push({ id: item.id, name, size: item.size ?? 0, modified: item.modified_at });
  };

  const resolved: ItemRow[] = [];
  for (const p of paths) resolved.push((await resolvePath(env, drive, p)).item);

  for (const item of resolved) {
    if (item.type === "file") {
      addFile(item, uniqueTopName(item.name));
      continue;
    }
    if (seenIds.has(item.id)) continue;
    seenIds.add(item.id);
    const isRoot = item.path_lower === rootPathLower(config);
    const top = uniqueTopName(isRoot ? sanitizeName(item.name === "root" ? "files" : item.name) : item.name);
    entries.push({ name: `${top}/`, modified: item.modified_at });
    folderCount++;

    const descendants = await listDescendants(env.DB, item.path_lower, config.zipMaxFiles * 4 + 1);
    if (descendants.length > config.zipMaxFiles * 4) {
      throw new ApiError(413, "LIMIT_EXCEEDED", `"${item.name}" contains too many items to zip`);
    }
    const depth = item.path.split("/").filter(Boolean).length;
    for (const d of descendants) {
      const rel = d.path.split("/").filter(Boolean).slice(depth).join("/");
      if (d.type === "folder") {
        if (seenIds.has(d.id)) continue;
        seenIds.add(d.id);
        entries.push({ name: `${top}/${rel}/`, modified: d.modified_at });
        folderCount++;
      } else {
        addFile(d, `${top}/${rel}`);
      }
    }
  }

  let name: string;
  if (requestedName?.trim()) {
    name = sanitizeName(requestedName.replace(/\.zip$/i, ""));
  } else if (resolved.length === 1) {
    const only = resolved[0]!;
    name = sanitizeName(only.type === "folder" ? entries[0]!.name.replace(/\/$/, "") : only.name.replace(/\.[^.]+$/, ""));
  } else {
    name = `files-${new Date().toISOString().slice(0, 10)}`;
  }

  return { name, entries, fileCount, folderCount, totalBytes };
}

export async function storeManifest(db: D1Database, manifest: ZipManifest) {
  const id = crypto.randomUUID().replace(/-/g, "");
  const now = Date.now();
  const expiresAt = now + MANIFEST_TTL_MS;
  await db
    .prepare(
      "INSERT INTO zip_manifests (id, name, entries, file_count, total_bytes, created_at, expires_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
    )
    .bind(id, manifest.name, JSON.stringify(manifest.entries), manifest.fileCount, manifest.totalBytes, now, expiresAt)
    .run();
  return { id, expiresAt: new Date(expiresAt).toISOString() };
}

export async function loadManifest(db: D1Database, id: string): Promise<ZipManifest> {
  const row = await db
    .prepare("SELECT id, name, entries, file_count, total_bytes FROM zip_manifests WHERE id = ? AND expires_at > ?")
    .bind(id, Date.now())
    .first<{ id: string; name: string; entries: string; file_count: number; total_bytes: number }>();
  if (!row) throw new ApiError(404, "NOT_FOUND", "Zip manifest not found or expired");
  const entries = JSON.parse(row.entries) as ZipEntry[];
  return {
    id: row.id,
    name: row.name,
    entries,
    fileCount: row.file_count,
    folderCount: entries.filter((e) => !e.id).length,
    totalBytes: row.total_bytes,
  };
}

export async function deleteExpiredManifests(db: D1Database) {
  await db.prepare("DELETE FROM zip_manifests WHERE expires_at <= ?").bind(Date.now()).run();
}

/**
 * Stream the archive. File metadata is refreshed from Graph first (batched),
 * so Content-Length is computed from current sizes and download URLs are fresh.
 */
export async function streamZip(env: Env, drive: DriveContext, manifest: ZipManifest): Promise<Response> {
  const graph = new GraphClient(env);
  const fileEntries = manifest.entries.filter((e) => e.id);

  const fresh = new Map<string, { size: number; downloadUrl?: string; modified?: string }>();
  const batches: (typeof fileEntries)[] = [];
  for (let i = 0; i < fileEntries.length; i += 20) batches.push(fileEntries.slice(i, i + 20));
  const CONCURRENCY = 3;
  for (let i = 0; i < batches.length; i += CONCURRENCY) {
    await Promise.all(
      batches.slice(i, i + CONCURRENCY).map(async (batch) => {
        const responses = await graph.batch<DriveItem>(
          batch.map((e) => ({
            id: e.id!,
            method: "GET",
            url: `/drives/${encodeURIComponent(drive.driveId)}/items/${encodeURIComponent(e.id!)}`,
          })),
        );
        for (const r of responses) {
          if (r.status === 200 && r.body) {
            fresh.set(r.id, {
              size: r.body.size ?? 0,
              downloadUrl: r.body["@microsoft.graph.downloadUrl"],
              modified: r.body.lastModifiedDateTime,
            });
          } else if (r.status !== 404) {
            throw new ApiError(502, "UPSTREAM_ERROR", `Could not read file metadata (Graph status ${r.status})`);
          }
        }
      }),
    );
  }

  // Files deleted in SharePoint since the last sync are left out.
  const entries = manifest.entries
    .filter((e) => !e.id || fresh.has(e.id))
    .map((e) => (e.id ? { ...e, size: fresh.get(e.id)!.size, modified: fresh.get(e.id)!.modified ?? e.modified } : e));
  const skipped = fileEntries.length - fresh.size;

  const metadata = entries.map((e) => (e.id ? { name: e.name, size: e.size! } : { name: e.name }));
  const length = predictLength(metadata);

  async function* files() {
    for (const e of entries) {
      const lastModified = e.modified ? new Date(e.modified) : new Date();
      if (!e.id) {
        yield { name: e.name, lastModified };
        continue;
      }
      const info = fresh.get(e.id)!;
      const res = await fetchFileBody(graph, drive, e.id, {}, info.downloadUrl);
      if (!res.ok || !res.body) {
        throw new Error(`Download of "${e.name}" failed with ${res.status}`);
      }
      yield { name: e.name, lastModified, size: info.size, input: res.body };
    }
  }

  const zip = downloadZip(files(), { length });
  const headers = new Headers(zip.headers);
  headers.set("content-type", "application/zip");
  headers.set("content-disposition", contentDisposition("attachment", `${manifest.name}.zip`));
  headers.set("cache-control", "no-store");
  headers.set("x-zip-file-count", String(entries.length - manifest.folderCount));
  if (skipped > 0) headers.set("x-zip-skipped-files", String(skipped));
  // A FixedLengthStream makes the runtime send Content-Length instead of chunked
  // encoding, so browsers can show real download progress.
  const { readable, writable } = new FixedLengthStream(Number(length));
  zip.body!.pipeTo(writable).catch((err) => console.error("Zip stream failed", err));
  return new Response(readable, { status: 200, headers });
}
