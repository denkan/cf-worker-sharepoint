import { z } from "@hono/zod-openapi";

export const PathParam = z
  .string()
  .optional()
  .openapi({
    description:
      "Item to act on. Accepts a path relative to the root folder (`/Reports/q3.pdf`), a full path " +
      "(`/sites/X/Shared Documents/Clients/Acme/Reports/q3.pdf` or `Shared Documents/...`), or a SharePoint URL " +
      "(browser URL, library view URL or sharing link). Omit for the root folder.",
    example: "/Rapporter 2026/Q3-rapport.pdf",
  });

const LinksSchema = z
  .object({
    self: z.string(),
    children: z.string().optional(),
    content: z.string().optional(),
    download: z.string().optional(),
    thumbnail: z.string().optional(),
  })
  .openapi("ItemLinks");

export const ItemSchema = z
  .object({
    id: z.string(),
    name: z.string(),
    path: z.string().openapi({ description: "Canonical path relative to the root folder", example: "/Rapporter 2026/Q3-rapport.pdf" }),
    type: z.enum(["file", "folder"]),
    size: z.number().nullable(),
    mimeType: z.string().nullable(),
    extension: z.string().nullable(),
    childCount: z.number().nullable(),
    created: z.string().nullable(),
    modified: z.string().nullable(),
    modifiedBy: z.string().nullable(),
    etag: z.string().nullable(),
    webUrl: z.string().nullable().openapi({ description: "Link to the item in SharePoint" }),
    links: LinksSchema,
  })
  .openapi("Item");

export const IndexInfoSchema = z
  .object({
    ready: z.boolean().openapi({ description: "False until the first full sync has completed" }),
    syncedAt: z.string().nullable().openapi({ description: "When the last sync pass completed" }),
  })
  .openapi("IndexInfo");

export const ErrorSchema = z
  .object({
    error: z.object({
      code: z.string(),
      message: z.string(),
      details: z.record(z.string(), z.unknown()).optional(),
    }),
  })
  .openapi("Error");

const intParam = (def: number, min: number, max: number) =>
  z.coerce.number().int().min(min).max(max).default(def);

export const PaginationQuery = {
  limit: intParam(100, 1, 1000).openapi({ description: "Page size (1–1000)" }),
  offset: intParam(0, 0, 10_000_000).openapi({ description: "Number of items to skip" }),
};

export const SortField = z.enum(["name", "modified", "size", "type"]);
export const SortOrder = z.enum(["asc", "desc"]);
export const ItemType = z.enum(["file", "folder"]);

export const ListResponse = z
  .object({
    folder: ItemSchema,
    items: z.array(ItemSchema),
    total: z.number(),
    limit: z.number(),
    offset: z.number(),
    index: IndexInfoSchema,
  })
  .openapi("ListResponse");

export const SearchResponse = z
  .object({
    scope: z.string(),
    query: z.string().nullable(),
    items: z.array(ItemSchema),
    total: z.number(),
    limit: z.number(),
    offset: z.number(),
    index: IndexInfoSchema,
  })
  .openapi("SearchResponse");

export const ResolveResponse = z
  .object({
    input: z.string(),
    matchedAs: z.enum(["relative", "fullPath", "url", "sharingLink"]),
    item: ItemSchema,
  })
  .openapi("ResolveResponse");

export const ZipRequest = z
  .object({
    paths: z.array(z.string()).min(1).max(1000).openapi({ description: "Files and/or folders (any accepted path format)" }),
    name: z.string().max(200).optional().openapi({ description: "Zip file name (without .zip)" }),
  })
  .openapi("ZipRequest");

export const ZipManifestResponse = z
  .object({
    id: z.string(),
    url: z.string().openapi({ description: "GET this URL to download the zip" }),
    name: z.string(),
    fileCount: z.number(),
    folderCount: z.number(),
    totalBytes: z.number(),
    expiresAt: z.string(),
  })
  .openapi("ZipManifest");

export const SyncResultSchema = z
  .object({
    status: z.enum(["completed", "partial", "skipped", "failed"]),
    mode: z.enum(["full", "incremental"]).optional(),
    pages: z.number(),
    upserted: z.number(),
    deleted: z.number(),
    durationMs: z.number(),
    error: z.string().optional(),
  })
  .openapi("SyncResult");

export const StatusResponse = z
  .object({
    site: z.string(),
    library: z.string().nullable(),
    rootPath: z.string(),
    mock: z.boolean(),
    drive: z
      .object({ siteId: z.string(), driveId: z.string(), name: z.string(), webUrl: z.string() })
      .nullable(),
    index: z.object({
      ready: z.boolean(),
      syncRunning: z.boolean(),
      mode: z.string().nullable(),
      itemsInRoot: z.number(),
      itemsInLibrary: z.number(),
      lastSyncStartedAt: z.string().nullable(),
      lastSyncCompletedAt: z.string().nullable(),
      lastFullSyncCompletedAt: z.string().nullable(),
      lastError: z.string().nullable(),
      lastResult: SyncResultSchema.nullable(),
    }),
  })
  .openapi("Status");

export const errorResponses = {
  400: { description: "Invalid request", content: { "application/json": { schema: ErrorSchema } } },
  404: { description: "Not found within the root folder", content: { "application/json": { schema: ErrorSchema } } },
  502: { description: "SharePoint / Graph error", content: { "application/json": { schema: ErrorSchema } } },
  503: { description: "Index not ready yet", content: { "application/json": { schema: ErrorSchema } } },
} as const;
