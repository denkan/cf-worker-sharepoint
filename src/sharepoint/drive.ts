import { configFingerprint, getConfig, type Env } from "../env";
import { ApiError } from "../errors";
import { getMeta, setMeta } from "../db/meta";
import type { GraphClient } from "../graph/client";
import type { Drive, Site } from "../graph/types";

export interface DriveContext {
  fingerprint: string;
  siteId: string;
  driveId: string;
  /** Library display name, e.g. "Documents" or "Dokument". */
  driveName: string;
  driveWebUrl: string;
  /** Decoded server-relative library path, e.g. "/sites/Projects/Shared Documents". */
  libraryPath: string;
  /** Last segment of libraryPath, e.g. "Shared Documents". */
  librarySegment: string;
}

const META_KEY = "drive_context";
let cached: DriveContext | undefined;

export function clearDriveContextCache() {
  cached = undefined;
}

/** Stored drive context, or undefined if the first sync hasn't resolved it yet. */
export async function getStoredDriveContext(env: Env): Promise<DriveContext | undefined> {
  const fingerprint = configFingerprint(getConfig(env));
  if (cached?.fingerprint === fingerprint) return cached;
  const raw = await getMeta(env.DB, META_KEY);
  if (!raw) return undefined;
  const ctx = JSON.parse(raw) as DriveContext;
  if (ctx.fingerprint !== fingerprint) return undefined;
  cached = ctx;
  return ctx;
}

export async function requireDriveContext(env: Env): Promise<DriveContext> {
  const ctx = await getStoredDriveContext(env);
  if (!ctx) {
    throw new ApiError(
      503,
      "INDEX_NOT_READY",
      "The SharePoint library has not been synced yet. Wait for the scheduled sync or POST /api/sync.",
    );
  }
  return ctx;
}

/** Resolve site and document library through Graph and store the result. */
export async function resolveDriveContext(env: Env, graph: GraphClient): Promise<DriveContext> {
  const config = getConfig(env);
  const fingerprint = configFingerprint(config);

  const host = config.siteUrl.hostname;
  const sitePath = config.sitePath;
  const siteRef = sitePath ? `${host}:${encodePath(sitePath)}:` : host;
  const site = await graph.json<Site>(`/sites/${siteRef}?$select=id,webUrl,displayName`);

  let drive: Drive;
  if (!config.library) {
    drive = await graph.json<Drive>(`/sites/${site.id}/drive?$select=id,name,webUrl,driveType`);
  } else {
    const { value } = await graph.json<{ value: Drive[] }>(`/sites/${site.id}/drives?$select=id,name,webUrl,driveType`);
    const wanted = config.library.toLowerCase();
    const match = value.find(
      (d) => d.name.toLowerCase() === wanted || librarySegmentOf(d.webUrl).toLowerCase() === wanted,
    );
    if (!match) {
      throw new ApiError(
        500,
        "INTERNAL_ERROR",
        `Library "${config.library}" not found. Available: ${value.map((d) => `${d.name} (${librarySegmentOf(d.webUrl)})`).join(", ")}`,
      );
    }
    drive = match;
  }

  const libraryPath = decodeURIComponent(new URL(drive.webUrl).pathname).replace(/\/+$/, "");
  const ctx: DriveContext = {
    fingerprint,
    siteId: site.id,
    driveId: drive.id,
    driveName: drive.name,
    driveWebUrl: drive.webUrl,
    libraryPath,
    librarySegment: librarySegmentOf(drive.webUrl),
  };
  await setMeta(env.DB, { [META_KEY]: JSON.stringify(ctx) });
  cached = ctx;
  return ctx;
}

function librarySegmentOf(webUrl: string): string {
  const segments = new URL(webUrl).pathname.split("/").filter(Boolean);
  return decodeURIComponent(segments[segments.length - 1] ?? "");
}

/** Encode each segment of a path for use in a Graph URL. */
export function encodePath(path: string): string {
  return path
    .split("/")
    .map((s) => encodeURIComponent(s))
    .join("/");
}
