/**
 * Turns whatever a user supplies as `path` into an item inside the root folder:
 *
 *  - a path relative to the configured root:            "/Rapporter 2026/Q3.pdf"
 *  - a full path, server-relative or library-relative:  "/sites/Projects/Shared Documents/Clients/Acme/…"
 *                                                       "Shared Documents/Clients/Acme/…"
 *  - a SharePoint URL: a browser/webUrl URL, a library view URL (AllItems.aspx?id=…),
 *    a "people with existing access" link (/:f:/r/…), or a sharing link (/:b:/s/…),
 *    which is resolved through Graph.
 */
import { getConfig, type Env } from "../env";
import { notFound } from "../errors";
import { getItemById, getItemsByPathLower, rootPathLower, type ItemRow } from "../db/items";
import { GraphClient } from "../graph/client";
import type { DriveItem } from "../graph/types";
import type { DriveContext } from "./drive";
import { isInScope, joinPath, lowerPath, lowerSegment, splitPath } from "./paths";

export type MatchKind = "relative" | "fullPath" | "url" | "sharingLink";

export interface Resolved {
  item: ItemRow;
  matchedAs: MatchKind;
}

interface Candidate {
  drivePath: string;
  kind: MatchKind;
}

export async function resolvePath(env: Env, drive: DriveContext, input: string | undefined): Promise<Resolved> {
  const config = getConfig(env);
  const raw = (input ?? "").trim();
  const rootLower = rootPathLower(config);

  let candidates: Candidate[];
  if (/^https?:\/\//i.test(raw)) {
    const fromUrl = await candidatesFromUrl(env, drive, raw);
    if ("item" in fromUrl) {
      if (!isInScope(fromUrl.item.path_lower, rootLower)) throw notFound();
      return fromUrl;
    }
    candidates = fromUrl.candidates;
  } else {
    candidates = [...fullPathCandidates(drive, raw), { drivePath: joinPath([...splitPath(config.rootPath), ...splitPath(raw)]), kind: "relative" }];
  }

  const inScope = candidates.filter((c) => isInScope(lowerPath(c.drivePath), rootLower));
  if (inScope.length === 0) throw notFound();
  const found = await getItemsByPathLower(env.DB, inScope.map((c) => lowerPath(c.drivePath)));
  for (const c of inScope) {
    const item = found.get(lowerPath(c.drivePath));
    if (item) return { item, matchedAs: c.kind };
  }
  throw notFound();
}

/** Interpret input as a full path: server-relative ("/sites/X/Lib/...") or library-relative ("Lib/..."). */
function fullPathCandidates(drive: DriveContext, raw: string, kind: MatchKind = "fullPath"): Candidate[] {
  const segments = splitPath(raw);
  const out: Candidate[] = [];
  const lower = segments.map(lowerSegment);

  const libSegments = splitPath(drive.libraryPath).map(lowerSegment);
  if (raw.startsWith("/") && libSegments.every((s, i) => lower[i] === s)) {
    out.push({ drivePath: joinPath(segments.slice(libSegments.length)), kind });
  }
  const first = lower[0];
  if (first && (first === lowerSegment(drive.librarySegment) || first === lowerSegment(drive.driveName))) {
    out.push({ drivePath: joinPath(segments.slice(1)), kind });
  }
  return out;
}

async function candidatesFromUrl(
  env: Env,
  drive: DriveContext,
  raw: string,
): Promise<Resolved | { candidates: Candidate[] }> {
  const config = getConfig(env);
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    throw notFound("Not a valid URL");
  }
  if (url.hostname.toLowerCase() !== config.siteUrl.hostname.toLowerCase()) {
    throw notFound("The URL is not on the configured SharePoint host");
  }
  let pathname: string;
  try {
    pathname = decodeURIComponent(url.pathname);
  } catch {
    throw notFound("The URL path is not valid");
  }

  // "People with existing access" links carry the server-relative path: /:f:/r/sites/X/Lib/...
  const direct = pathname.match(/^\/:[a-z]:\/r(\/.+)$/i);
  if (direct) return { candidates: fullPathCandidates(drive, direct[1]!, "url") };

  // Other sharing links (/:b:/s/..., /:f:/g/...) and _layouts pages (Doc.aspx etc.) are opaque.
  const listPath = url.searchParams.get("id") ?? url.searchParams.get("RootFolder");
  if (/^\/:[a-z]:\//i.test(pathname) || (/\/_layouts\//i.test(pathname) && !listPath)) {
    return resolveSharingLink(env, drive, raw);
  }

  // Library views: /sites/X/Lib/Forms/AllItems.aspx?id=/sites/X/Lib/Folder
  if (listPath) return { candidates: fullPathCandidates(drive, listPath, "url") };
  const formsView = pathname.match(/^(.*)\/Forms\/[^/]+\.aspx$/i);
  if (formsView) return { candidates: fullPathCandidates(drive, formsView[1]!, "url") };

  return { candidates: fullPathCandidates(drive, pathname, "url") };
}

async function resolveSharingLink(env: Env, drive: DriveContext, link: string): Promise<Resolved> {
  const graph = new GraphClient(env);
  const encoded = btoa(String.fromCharCode(...new TextEncoder().encode(link)))
    .replace(/=+$/, "")
    .replace(/\//g, "_")
    .replace(/\+/g, "-");
  let shared: DriveItem;
  try {
    shared = await graph.json<DriveItem>(`/shares/u!${encoded}/driveItem?$select=id,parentReference`);
  } catch (err) {
    if (err instanceof Error && "status" in err && (err as { status: number }).status === 404) {
      throw notFound("The sharing link could not be resolved");
    }
    throw err;
  }
  if (shared.parentReference?.driveId && shared.parentReference.driveId !== drive.driveId) throw notFound();
  const item = await getItemById(env.DB, shared.id);
  if (!item || item.path === null) throw notFound("The shared item is not indexed (yet)");
  return { item, matchedAs: "sharingLink" };
}
