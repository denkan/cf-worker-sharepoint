import { badRequest } from "../errors";

/** Lower-casing used for every stored and queried path, applied per segment. */
export const lowerSegment = (segment: string) => segment.toLowerCase();

export const lowerPath = (path: string) => path.split("/").map(lowerSegment).join("/");

/** Length in Unicode code points, which is what SQLite's substr() counts. */
export const codePointLength = (s: string) => {
  let n = 0;
  for (const _ of s) n++;
  return n;
};

/**
 * Half-open range [lo, hi) of path_lower values strictly below `prefixLower`.
 * '/' is 0x2F and '0' is 0x30, so this is an index-friendly prefix match.
 */
export const subtreeRange = (prefixLower: string): [string, string] => [`${prefixLower}/`, `${prefixLower}0`];

/** Split a user-supplied path into clean segments. Rejects "." and "..". */
export function splitPath(input: string): string[] {
  const segments = input
    .normalize("NFC")
    .replace(/\\/g, "/")
    .split("/")
    .filter((s) => s.length > 0);
  for (const s of segments) {
    if (s === "." || s === "..") throw badRequest("Path segments '.' and '..' are not allowed");
  }
  return segments;
}

export const joinPath = (segments: string[]) => (segments.length ? `/${segments.join("/")}` : "");

/** Path relative to the configured root, as returned by the API ("/" for the root itself). */
export function toApiPath(drivePath: string, rootPath: string): string {
  const rootDepth = rootPath.split("/").filter(Boolean).length;
  const rest = drivePath.split("/").filter(Boolean).slice(rootDepth);
  return `/${rest.join("/")}`;
}

export function isInScope(pathLower: string | null, rootLower: string): boolean {
  if (pathLower === null) return false;
  return pathLower === rootLower || pathLower.startsWith(`${rootLower}/`);
}
