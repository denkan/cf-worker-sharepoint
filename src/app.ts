import { createRoute, OpenAPIHono, z } from "@hono/zod-openapi";
import { cors } from "hono/cors";
import { getConfig, type Env } from "./env";
import { ApiError, errorBody } from "./errors";
import { getMetaMany } from "./db/meta";
import { listChildren, rootPathLower, searchItems, toApiItem } from "./db/items";
import { getStoredDriveContext, requireDriveContext } from "./sharepoint/drive";
import { resolvePath } from "./sharepoint/resolve";
import { isSyncRunning, runSync, SYNC_META_KEYS, type SyncResult } from "./sync/delta";
import { downloadRedirect, proxyContent, thumbnail } from "./api/files";
import { buildManifest, loadManifest, storeManifest, streamZip } from "./api/zip";
import {
  errorResponses,
  ItemSchema,
  ItemType,
  ListResponse,
  PaginationQuery,
  PathParam,
  ResolveResponse,
  SearchResponse,
  SortField,
  SortOrder,
  StatusResponse,
  SyncResultSchema,
  ZipManifestResponse,
  ZipRequest,
} from "./api/schemas";
import { uiHtml } from "./ui";

type AppEnv = { Bindings: Env };

export const app = new OpenAPIHono<AppEnv>({
  defaultHook: (result, c) => {
    if (!result.success) {
      return c.json(
        errorBody("BAD_REQUEST", "Invalid request parameters", {
          issues: result.error.issues.map((i) => ({ path: i.path.join("."), message: i.message })),
        }),
        400,
      );
    }
  },
});

app.onError((err, c) => {
  if (err instanceof ApiError) {
    return c.json(errorBody(err.code, err.message, err.details), err.status);
  }
  console.error(err);
  return c.json(errorBody("INTERNAL_ERROR", "Internal error"), 500);
});

app.notFound((c) => c.json(errorBody("NOT_FOUND", "No such endpoint"), 404));

app.use("/api/*", async (c, next) => {
  const origins = getConfig(c.env).corsOrigins;
  if (origins.length === 0) return next();
  return cors({
    origin: origins.includes("*") ? "*" : origins,
    allowMethods: ["GET", "HEAD", "POST", "OPTIONS"],
    allowHeaders: ["content-type", "range"],
    exposeHeaders: ["content-length", "content-range", "content-disposition", "accept-ranges", "x-zip-file-count", "x-zip-skipped-files"],
    maxAge: 600,
  })(c, next);
});

async function indexInfo(env: Env) {
  const meta = await getMetaMany(env.DB, ["index_ready", "last_sync_completed_at"]);
  return { ready: meta.index_ready === "1", syncedAt: meta.last_sync_completed_at };
}

// --- Browsing ---------------------------------------------------------------------

app.openapi(
  createRoute({
    method: "get",
    path: "/api/items",
    tags: ["Browse"],
    summary: "List a folder",
    description: "Folders come first, then files. Pagination with limit/offset; `total` is the full count.",
    request: {
      query: z.object({
        path: PathParam,
        ...PaginationQuery,
        sort: SortField.default("name"),
        order: SortOrder.default("asc"),
        type: ItemType.optional().openapi({ description: "Only files or only folders" }),
      }),
    },
    responses: { 200: { description: "Folder contents", content: { "application/json": { schema: ListResponse } } }, ...errorResponses },
  }),
  async (c) => {
    const q = c.req.valid("query");
    const config = getConfig(c.env);
    const drive = await requireDriveContext(c.env);
    const { item } = await resolvePath(c.env, drive, q.path);
    if (item.type !== "folder") throw new ApiError(400, "NOT_A_FOLDER", `"${item.name}" is a file`);
    const [{ items, total }, index] = await Promise.all([
      listChildren(c.env.DB, item.id, { limit: q.limit, offset: q.offset, sort: q.sort, order: q.order, type: q.type }),
      indexInfo(c.env),
    ]);
    return c.json(
      {
        folder: toApiItem(item, config),
        items: items.map((i) => toApiItem(i, config)),
        total,
        limit: q.limit,
        offset: q.offset,
        index,
      },
      200,
    );
  },
);

app.openapi(
  createRoute({
    method: "get",
    path: "/api/item",
    tags: ["Browse"],
    summary: "Metadata for one file or folder",
    request: { query: z.object({ path: PathParam }) },
    responses: { 200: { description: "The item", content: { "application/json": { schema: ItemSchema } } }, ...errorResponses },
  }),
  async (c) => {
    const drive = await requireDriveContext(c.env);
    const { item } = await resolvePath(c.env, drive, c.req.valid("query").path);
    return c.json(toApiItem(item, getConfig(c.env)), 200);
  },
);

app.openapi(
  createRoute({
    method: "get",
    path: "/api/resolve",
    tags: ["Browse"],
    summary: "Normalize any accepted path format to the canonical path",
    description: "Useful for turning a URL or full path pasted by a user into the canonical relative path.",
    request: { query: z.object({ path: PathParam }) },
    responses: { 200: { description: "Resolved item", content: { "application/json": { schema: ResolveResponse } } }, ...errorResponses },
  }),
  async (c) => {
    const input = c.req.valid("query").path ?? "";
    const drive = await requireDriveContext(c.env);
    const { item, matchedAs } = await resolvePath(c.env, drive, input);
    return c.json({ input, matchedAs, item: toApiItem(item, getConfig(c.env)) }, 200);
  },
);

app.openapi(
  createRoute({
    method: "get",
    path: "/api/search",
    tags: ["Browse"],
    summary: "Search file and folder names and metadata",
    description:
      "Every word in `q` must match the start of a word in the name (accent-insensitive). " +
      "Combine with filters, or omit `q` to filter only. Results are limited to the `path` subtree.",
    request: {
      query: z.object({
        q: z.string().max(200).optional().openapi({ example: "rapport q3" }),
        path: PathParam.openapi({ description: "Folder to search within (any accepted path format). Default: root." }),
        type: ItemType.optional(),
        ext: z.string().optional().openapi({ description: "Comma-separated extensions", example: "pdf,docx" }),
        modifiedAfter: z.iso.datetime({ offset: true }).or(z.iso.date()).optional(),
        modifiedBefore: z.iso.datetime({ offset: true }).or(z.iso.date()).optional(),
        minSize: z.coerce.number().int().min(0).optional(),
        maxSize: z.coerce.number().int().min(0).optional(),
        sort: z.enum(["relevance", "name", "modified", "size", "type"]).default("relevance"),
        order: SortOrder.default("asc"),
        ...PaginationQuery,
      }),
    },
    responses: { 200: { description: "Matches", content: { "application/json": { schema: SearchResponse } } }, ...errorResponses },
  }),
  async (c) => {
    const q = c.req.valid("query");
    const config = getConfig(c.env);
    const drive = await requireDriveContext(c.env);
    const { item: scope } = await resolvePath(c.env, drive, q.path);
    if (scope.type !== "folder") throw new ApiError(400, "NOT_A_FOLDER", `"${scope.name}" is a file`);
    const toIso = (v?: string) => (v ? new Date(v).toISOString().replace(/\.\d{3}Z$/, "Z") : undefined);
    const [{ items, total }, index] = await Promise.all([
      searchItems(c.env.DB, {
        query: q.q?.trim() || undefined,
        scopePathLower: scope.path_lower,
        type: q.type,
        extensions: q.ext?.split(",").map((e) => e.trim()).filter(Boolean),
        modifiedAfter: toIso(q.modifiedAfter),
        modifiedBefore: toIso(q.modifiedBefore),
        minSize: q.minSize,
        maxSize: q.maxSize,
        sort: q.sort,
        order: q.order,
        limit: q.limit,
        offset: q.offset,
      }),
      indexInfo(c.env),
    ]);
    return c.json(
      {
        scope: toApiItem(scope, config).path,
        query: q.q?.trim() || null,
        items: items.map((i) => toApiItem(i, config)),
        total,
        limit: q.limit,
        offset: q.offset,
        index,
      },
      200,
    );
  },
);

// --- Files ------------------------------------------------------------------------

const binary = (description: string) => ({ description, content: { "application/octet-stream": { schema: z.string().openapi({ format: "binary" }) } } });

app.openapi(
  createRoute({
    method: "get",
    path: "/api/download",
    tags: ["Files"],
    summary: "Download a file (redirect)",
    description: "Redirects to a short-lived, pre-authenticated SharePoint download URL.",
    request: { query: z.object({ path: PathParam }) },
    responses: { 302: { description: "Redirect to the file" }, ...errorResponses },
  }),
  async (c) => {
    const drive = await requireDriveContext(c.env);
    const { item } = await resolvePath(c.env, drive, c.req.valid("query").path);
    return downloadRedirect(c.env, drive, item);
  },
);

app.openapi(
  createRoute({
    method: "get",
    path: "/api/content",
    tags: ["Files"],
    summary: "Serve a file through the Worker",
    description:
      "Streams the file with its content type. Supports `Range` requests (206) for video seeking. " +
      "PDF, images, audio, video and plain text are served inline; anything else, or `disposition=attachment`, " +
      "is served as a download. SVG is sandboxed with a CSP.",
    request: {
      query: z.object({
        path: PathParam,
        disposition: z.enum(["inline", "attachment"]).default("inline"),
      }),
    },
    responses: { 200: binary("The file"), 206: binary("Partial content"), ...errorResponses },
  }),
  async (c) => {
    const q = c.req.valid("query");
    const drive = await requireDriveContext(c.env);
    const { item } = await resolvePath(c.env, drive, q.path);
    return proxyContent(c.env, drive, item, c.req.raw, q.disposition);
  },
);

app.openapi(
  createRoute({
    method: "get",
    path: "/api/thumbnail",
    tags: ["Files"],
    summary: "Thumbnail image for a file",
    request: {
      query: z.object({ path: PathParam, size: z.enum(["small", "medium", "large"]).default("medium") }),
    },
    responses: { 200: binary("Thumbnail image"), ...errorResponses },
  }),
  async (c) => {
    const q = c.req.valid("query");
    const drive = await requireDriveContext(c.env);
    const { item } = await resolvePath(c.env, drive, q.path);
    let ctx: ExecutionContext | undefined;
    try {
      ctx = c.executionCtx as unknown as ExecutionContext;
    } catch {
      ctx = undefined;
    }
    return thumbnail(c.env, drive, item, q.size, ctx);
  },
);

// --- Zip --------------------------------------------------------------------------

app.openapi(
  createRoute({
    method: "post",
    path: "/api/zip",
    tags: ["Zip"],
    summary: "Prepare a zip of files and/or folders",
    description: "Returns a short-lived URL (1 hour) that streams the zip. Folders are included recursively.",
    request: { body: { content: { "application/json": { schema: ZipRequest } }, required: true } },
    responses: {
      200: { description: "Manifest created", content: { "application/json": { schema: ZipManifestResponse } } },
      413: { description: "Too many files or bytes", content: { "application/json": { schema: errorResponses[400].content["application/json"].schema } } },
      ...errorResponses,
    },
  }),
  async (c) => {
    const body = c.req.valid("json");
    const drive = await requireDriveContext(c.env);
    const manifest = await buildManifest(c.env, drive, body.paths, body.name);
    const { id, expiresAt } = await storeManifest(c.env.DB, manifest);
    return c.json(
      {
        id,
        url: `/api/zip/${id}`,
        name: `${manifest.name}.zip`,
        fileCount: manifest.fileCount,
        folderCount: manifest.folderCount,
        totalBytes: manifest.totalBytes,
        expiresAt,
      },
      200,
    );
  },
);

app.openapi(
  createRoute({
    method: "get",
    path: "/api/zip/{id}",
    tags: ["Zip"],
    summary: "Download a prepared zip",
    request: { params: z.object({ id: z.string().regex(/^[0-9a-f]{32}$/) }) },
    responses: { 200: { description: "Zip archive", content: { "application/zip": { schema: z.string().openapi({ format: "binary" }) } } }, ...errorResponses },
  }),
  async (c) => {
    const drive = await requireDriveContext(c.env);
    const manifest = await loadManifest(c.env.DB, c.req.valid("param").id);
    return streamZip(c.env, drive, manifest);
  },
);

app.openapi(
  createRoute({
    method: "get",
    path: "/api/zip",
    tags: ["Zip"],
    summary: "Download a zip directly",
    description: "Convenience for plain links: repeat `path` for each file or folder, e.g. `?path=/A&path=/B/c.pdf`.",
    request: {
      query: z.object({
        path: z.union([z.string(), z.array(z.string())]).openapi({ description: "File or folder; repeat for several" }),
        name: z.string().max(200).optional(),
      }),
    },
    responses: { 200: { description: "Zip archive", content: { "application/zip": { schema: z.string().openapi({ format: "binary" }) } } }, ...errorResponses },
  }),
  async (c) => {
    const paths = c.req.queries("path") ?? [];
    if (paths.length === 0) throw new ApiError(400, "BAD_REQUEST", "At least one path is required");
    if (paths.length > 1000) throw new ApiError(400, "BAD_REQUEST", "At most 1000 paths");
    const drive = await requireDriveContext(c.env);
    const manifest = await buildManifest(c.env, drive, paths, c.req.valid("query").name);
    return streamZip(c.env, drive, manifest);
  },
);

// --- Operations -------------------------------------------------------------------

app.openapi(
  createRoute({
    method: "get",
    path: "/api/status",
    tags: ["Operations"],
    summary: "Configuration and sync status",
    responses: { 200: { description: "Status", content: { "application/json": { schema: StatusResponse } } } },
  }),
  async (c) => {
    const config = getConfig(c.env);
    const [drive, meta, running, counts] = await Promise.all([
      getStoredDriveContext(c.env),
      getMetaMany(c.env.DB, SYNC_META_KEYS),
      isSyncRunning(c.env.DB),
      countItems(c.env.DB, rootPathLower(config)),
    ]);
    return c.json(
      {
        site: config.siteUrl.toString(),
        library: drive?.driveName ?? (config.library || null),
        rootPath: config.rootPath || "/",
        mock: config.mock,
        drive: drive ? { siteId: drive.siteId, driveId: drive.driveId, name: drive.driveName, webUrl: drive.driveWebUrl } : null,
        index: {
          ready: meta.index_ready === "1",
          syncRunning: running,
          mode: meta.sync_mode,
          itemsInRoot: counts.inRoot,
          itemsInLibrary: counts.total,
          lastSyncStartedAt: meta.last_sync_started_at,
          lastSyncCompletedAt: meta.last_sync_completed_at,
          lastFullSyncCompletedAt: meta.last_full_sync_completed_at,
          lastError: meta.last_sync_error,
          lastResult: meta.last_sync_result ? (JSON.parse(meta.last_sync_result) as SyncResult) : null,
        },
      },
      200,
    );
  },
);

app.openapi(
  createRoute({
    method: "post",
    path: "/api/sync",
    tags: ["Operations"],
    summary: "Run a sync pass now",
    description: "Same as the scheduled sync. Returns `skipped` if a sync is already running.",
    responses: { 200: { description: "Sync result", content: { "application/json": { schema: SyncResultSchema } } } },
  }),
  async (c) => c.json(await runSync(c.env), 200),
);

async function countItems(db: D1Database, rootLower: string) {
  const row = await db
    .prepare(
      `SELECT COUNT(*) AS total,
              SUM(CASE WHEN path_lower = ?1 OR (path_lower >= ?1 || '/' AND path_lower < ?1 || '0') THEN 1 ELSE 0 END) AS in_root
       FROM items`,
    )
    .bind(rootLower)
    .first<{ total: number; in_root: number | null }>();
  return { total: row?.total ?? 0, inRoot: row?.in_root ?? 0 };
}

// --- Docs and PoC UI --------------------------------------------------------------

app.doc31("/openapi.json", {
  openapi: "3.1.0",
  info: {
    title: "SharePoint folder API",
    version: "1.0.0",
    description:
      "Read-only API over one SharePoint folder. Listing and search are served from a D1 mirror kept " +
      "up to date by a delta sync; file content is fetched from SharePoint on demand.",
  },
});

app.get("/", (c) => c.html(uiHtml));
