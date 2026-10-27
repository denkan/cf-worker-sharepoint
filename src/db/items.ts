import type { Config } from "../env";
import { lowerPath, subtreeRange, toApiPath } from "../sharepoint/paths";

export interface ItemRow {
  id: string;
  parent_id: string | null;
  name: string;
  path: string;
  path_lower: string;
  type: "file" | "folder";
  extension: string | null;
  mime_type: string | null;
  size: number | null;
  created_at: string | null;
  modified_at: string | null;
  modified_by: string | null;
  etag: string | null;
  ctag: string | null;
  web_url: string | null;
  child_count?: number | null;
}

export interface ApiItem {
  id: string;
  name: string;
  path: string;
  type: "file" | "folder";
  size: number | null;
  mimeType: string | null;
  extension: string | null;
  childCount: number | null;
  created: string | null;
  modified: string | null;
  modifiedBy: string | null;
  etag: string | null;
  webUrl: string | null;
  links: {
    self: string;
    children?: string;
    content?: string;
    download?: string;
    thumbnail?: string;
  };
}

const COLUMNS =
  "i.id, i.parent_id, i.name, i.path, i.path_lower, i.type, i.extension, i.mime_type, i.size, i.created_at, i.modified_at, i.modified_by, i.etag, i.ctag, i.web_url";
const CHILD_COUNT = "(SELECT COUNT(*) FROM items c WHERE c.parent_id = i.id) AS child_count";

export function toApiItem(row: ItemRow, config: Config): ApiItem {
  const path = toApiPath(row.path, config.rootPath);
  const q = `?path=${encodeURIComponent(path)}`;
  const isFolder = row.type === "folder";
  return {
    id: row.id,
    name: path === "/" ? row.name || "/" : row.name,
    path,
    type: row.type,
    size: row.size,
    mimeType: row.mime_type,
    extension: row.extension,
    childCount: isFolder ? (row.child_count ?? null) : null,
    created: row.created_at,
    modified: row.modified_at,
    modifiedBy: row.modified_by,
    etag: row.etag,
    webUrl: row.web_url,
    links: isFolder
      ? { self: `/api/item${q}`, children: `/api/items${q}` }
      : {
          self: `/api/item${q}`,
          content: `/api/content${q}`,
          download: `/api/download${q}`,
          thumbnail: `/api/thumbnail${q}`,
        },
  };
}

/** Look up items by lower-cased drive paths; returns them keyed by path_lower. */
export async function getItemsByPathLower(db: D1Database, pathsLower: string[]): Promise<Map<string, ItemRow>> {
  const { results } = await db
    .prepare(`SELECT ${COLUMNS}, ${CHILD_COUNT} FROM items i WHERE i.path_lower IN (SELECT value FROM json_each(?))`)
    .bind(JSON.stringify(pathsLower))
    .all<ItemRow>();
  return new Map(results.map((r) => [r.path_lower, r]));
}

export async function getItemById(db: D1Database, id: string): Promise<ItemRow | null> {
  return db.prepare(`SELECT ${COLUMNS}, ${CHILD_COUNT} FROM items i WHERE i.id = ?`).bind(id).first<ItemRow>();
}

export type SortField = "name" | "modified" | "size" | "type";
export type SortOrder = "asc" | "desc";

function orderBy(sort: SortField, order: SortOrder, foldersFirst: boolean) {
  const dir = order === "desc" ? "DESC" : "ASC";
  const parts: string[] = [];
  if (foldersFirst) parts.push("CASE i.type WHEN 'folder' THEN 0 ELSE 1 END");
  switch (sort) {
    case "modified":
      parts.push(`i.modified_at ${dir}`);
      break;
    case "size":
      parts.push(`i.size ${dir}`);
      break;
    case "type":
      parts.push(`i.extension ${dir}`);
      break;
  }
  parts.push(`i.name_lower ${dir}`, "i.id");
  return parts.join(", ");
}

export async function listChildren(
  db: D1Database,
  parentId: string,
  opts: { limit: number; offset: number; sort: SortField; order: SortOrder; type?: "file" | "folder" },
): Promise<{ items: ItemRow[]; total: number }> {
  const typeFilter = opts.type ? "AND i.type = ?2" : "";
  const binds: unknown[] = [parentId];
  if (opts.type) binds.push(opts.type);
  const [rows, count] = await db.batch<ItemRow & { total?: number }>([
    db
      .prepare(
        `SELECT ${COLUMNS}, ${CHILD_COUNT} FROM items i
         WHERE i.parent_id = ?1 AND i.path IS NOT NULL ${typeFilter}
         ORDER BY ${orderBy(opts.sort, opts.order, true)}
         LIMIT ${opts.limit} OFFSET ${opts.offset}`,
      )
      .bind(...binds),
    db
      .prepare(`SELECT COUNT(*) AS total FROM items i WHERE i.parent_id = ?1 AND i.path IS NOT NULL ${typeFilter}`)
      .bind(...binds),
  ]);
  return { items: rows!.results, total: Number(count!.results[0]?.total ?? 0) };
}

export interface SearchOptions {
  query?: string;
  scopePathLower: string;
  type?: "file" | "folder";
  extensions?: string[];
  modifiedAfter?: string;
  modifiedBefore?: string;
  minSize?: number;
  maxSize?: number;
  sort: "relevance" | SortField;
  order: SortOrder;
  limit: number;
  offset: number;
}

/** Turn free text into an FTS5 query: every word must match as a prefix. */
export function toFtsQuery(text: string): string | null {
  const tokens = text
    .normalize("NFC")
    .split(/[^\p{L}\p{N}]+/u)
    .filter(Boolean)
    .slice(0, 12);
  if (tokens.length === 0) return null;
  return tokens.map((t) => `"${t.replace(/"/g, '""')}"*`).join(" ");
}

export async function searchItems(db: D1Database, opts: SearchOptions): Promise<{ items: ItemRow[]; total: number }> {
  const [lo, hi] = subtreeRange(opts.scopePathLower);
  const where = ["i.path_lower >= ?", "i.path_lower < ?"];
  const binds: unknown[] = [lo, hi];

  const fts = opts.query ? toFtsQuery(opts.query) : null;
  if (opts.query && !fts) return { items: [], total: 0 };
  let from = "items i";
  if (fts) {
    from = "items_fts f JOIN items i ON i.rowid = f.rowid";
    where.unshift("items_fts MATCH ?");
    binds.unshift(fts);
  }
  if (opts.type) {
    where.push("i.type = ?");
    binds.push(opts.type);
  }
  if (opts.extensions?.length) {
    where.push(`i.extension IN (${opts.extensions.map(() => "?").join(",")})`);
    binds.push(...opts.extensions.map((e) => e.toLowerCase().replace(/^\./, "")));
  }
  if (opts.modifiedAfter) {
    where.push("i.modified_at >= ?");
    binds.push(opts.modifiedAfter);
  }
  if (opts.modifiedBefore) {
    where.push("i.modified_at < ?");
    binds.push(opts.modifiedBefore);
  }
  if (opts.minSize !== undefined) {
    where.push("i.size >= ?");
    binds.push(opts.minSize);
  }
  if (opts.maxSize !== undefined) {
    where.push("i.size <= ?");
    binds.push(opts.maxSize);
  }

  const order =
    opts.sort === "relevance"
      ? fts
        ? "bm25(items_fts), i.name_lower, i.id"
        : "i.name_lower, i.id"
      : orderBy(opts.sort, opts.order, false);

  const whereSql = where.join(" AND ");
  const [rows, count] = await db.batch<ItemRow & { total?: number }>([
    db
      .prepare(
        `SELECT ${COLUMNS}, ${CHILD_COUNT} FROM ${from} WHERE ${whereSql} ORDER BY ${order} LIMIT ${opts.limit} OFFSET ${opts.offset}`,
      )
      .bind(...binds),
    db.prepare(`SELECT COUNT(*) AS total FROM ${from} WHERE ${whereSql}`).bind(...binds),
  ]);
  return { items: rows!.results, total: Number(count!.results[0]?.total ?? 0) };
}

/** All files below a folder (for zips), with their path relative to that folder. */
export async function listDescendants(
  db: D1Database,
  folderPathLower: string,
  limit: number,
): Promise<ItemRow[]> {
  const [lo, hi] = subtreeRange(folderPathLower);
  const { results } = await db
    .prepare(`SELECT ${COLUMNS} FROM items i WHERE i.path_lower >= ? AND i.path_lower < ? ORDER BY i.path_lower LIMIT ?`)
    .bind(lo, hi, limit)
    .all<ItemRow>();
  return results;
}

export const rootPathLower = (config: Config) => lowerPath(config.rootPath);
