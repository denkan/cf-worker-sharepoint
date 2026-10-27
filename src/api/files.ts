import type { Env } from "../env";
import { ApiError } from "../errors";
import type { ItemRow } from "../db/items";
import { GraphClient, toGraphError } from "../graph/client";
import type { DriveContext } from "../sharepoint/drive";

/** Types that are safe to render inline on this Worker's origin. */
const INLINE_TYPES = [
  /^application\/pdf$/,
  /^image\/(png|jpeg|gif|webp|avif|bmp|svg\+xml)$/,
  /^video\//,
  /^audio\//,
  /^text\/plain$/,
];

export function contentDisposition(type: "inline" | "attachment", filename: string): string {
  const ascii = filename
    .normalize("NFKD")
    .replace(/\p{M}/gu, "")
    .replace(/[^\x20-\x7e]/g, "_")
    .replace(/["\\]/g, "_");
  const encoded = encodeURIComponent(filename).replace(/['()*]/g, (ch) => `%${ch.charCodeAt(0).toString(16).toUpperCase()}`);
  return `${type}; filename="${ascii}"; filename*=UTF-8''${encoded}`;
}

export function requireFile(item: ItemRow): ItemRow {
  if (item.type !== "file") throw new ApiError(400, "NOT_A_FILE", `"${item.name}" is a folder`);
  return item;
}

/** Fetch a download URL; if it has expired, get a fresh one and try once more. */
export async function fetchFileBody(
  graph: GraphClient,
  drive: DriveContext,
  itemId: string,
  init: RequestInit = {},
  downloadUrl?: string,
): Promise<Response> {
  let url = downloadUrl ?? (await graph.getDownloadUrl(drive.driveId, itemId));
  let res = await graph.fetchUnauthenticated(url, init);
  if (downloadUrl && (res.status === 401 || res.status === 403)) {
    await res.body?.cancel();
    url = await graph.getDownloadUrl(drive.driveId, itemId);
    res = await graph.fetchUnauthenticated(url, init);
  }
  return res;
}

/** GET /api/download: redirect to the pre-authenticated SharePoint download URL. */
export async function downloadRedirect(env: Env, drive: DriveContext, item: ItemRow): Promise<Response> {
  const url = await new GraphClient(env).getDownloadUrl(drive.driveId, requireFile(item).id);
  return new Response(null, { status: 302, headers: { location: url, "cache-control": "no-store" } });
}

/** GET /api/content: stream the file through the Worker, with Range support for media seeking. */
export async function proxyContent(
  env: Env,
  drive: DriveContext,
  item: ItemRow,
  request: Request,
  disposition: "inline" | "attachment",
): Promise<Response> {
  requireFile(item);
  const graph = new GraphClient(env);
  const upstreamHeaders = new Headers();
  const range = request.headers.get("range");
  if (range) upstreamHeaders.set("range", range);

  const upstream = await fetchFileBody(graph, drive, item.id, { headers: upstreamHeaders });
  if (!upstream.ok && upstream.status !== 416) {
    await upstream.body?.cancel();
    if (upstream.status === 404) throw new ApiError(404, "NOT_FOUND", "The file no longer exists in SharePoint");
    throw new ApiError(502, "UPSTREAM_ERROR", `SharePoint download failed with ${upstream.status}`);
  }

  const mime = (item.mime_type || upstream.headers.get("content-type") || "application/octet-stream").split(";")[0]!.trim().toLowerCase();
  const inlineAllowed = INLINE_TYPES.some((re) => re.test(mime));
  const effective = disposition === "inline" && inlineAllowed ? "inline" : "attachment";

  const headers = new Headers({
    "content-type": mime.startsWith("text/") ? `${mime}; charset=utf-8` : mime,
    "content-disposition": contentDisposition(effective, item.name),
    "accept-ranges": "bytes",
    "cache-control": "private, max-age=300",
    "x-content-type-options": "nosniff",
  });
  for (const h of ["content-length", "content-range"]) {
    const v = upstream.headers.get(h);
    if (v) headers.set(h, v);
  }
  if (item.etag) headers.set("etag", item.etag);
  if (item.modified_at) headers.set("last-modified", new Date(item.modified_at).toUTCString());
  if (mime === "image/svg+xml") {
    // SVG can carry scripts: render it, but never let it run on this origin.
    headers.set("content-security-policy", "sandbox; default-src 'none'; style-src 'unsafe-inline'; img-src data:");
  }
  return new Response(upstream.body, { status: upstream.status, headers });
}

const THUMB_SIZES = new Set(["small", "medium", "large"]);

/** GET /api/thumbnail: Graph-generated thumbnail, cached at the edge per item version. */
export async function thumbnail(
  env: Env,
  drive: DriveContext,
  item: ItemRow,
  size: string,
  ctx?: ExecutionContext,
): Promise<Response> {
  requireFile(item);
  if (!THUMB_SIZES.has(size)) throw new ApiError(400, "BAD_REQUEST", "size must be small, medium or large");

  const cacheKey = new Request(
    `https://thumbnail-cache.internal/${encodeURIComponent(drive.driveId)}/${encodeURIComponent(item.id)}/${encodeURIComponent(item.ctag ?? item.etag ?? "")}/${size}`,
  );
  const cache = (globalThis as unknown as { caches?: { default?: Cache } }).caches?.default;
  const hit = await cache?.match(cacheKey);
  if (hit) {
    const headers = new Headers(hit.headers);
    headers.set("cache-control", "private, max-age=86400");
    return new Response(hit.body, { headers });
  }

  const graph = new GraphClient(env);
  const res = await graph.request(
    `/drives/${encodeURIComponent(drive.driveId)}/items/${encodeURIComponent(item.id)}/thumbnails/0/${size}/content`,
    { redirect: "manual" },
  );
  let image: Response;
  const location = res.headers.get("location");
  if (res.status >= 300 && res.status < 400 && location) {
    await res.body?.cancel();
    image = await graph.fetchUnauthenticated(location);
  } else if (res.ok) {
    image = res;
  } else {
    if (res.status === 404) {
      await res.body?.cancel();
      throw new ApiError(404, "NOT_FOUND", "No thumbnail is available for this file");
    }
    throw await toGraphError(res);
  }
  if (!image.ok) {
    await image.body?.cancel();
    throw new ApiError(404, "NOT_FOUND", "No thumbnail is available for this file");
  }

  const bytes = await image.arrayBuffer();
  const headers = {
    "content-type": image.headers.get("content-type") ?? "image/jpeg",
    "x-content-type-options": "nosniff",
  };
  if (cache) {
    const put = cache.put(cacheKey, new Response(bytes, { headers: { ...headers, "cache-control": "public, max-age=86400" } }));
    if (ctx) ctx.waitUntil(put);
    else await put;
  }
  return new Response(bytes, { headers: { ...headers, "cache-control": "private, max-age=86400" } });
}
