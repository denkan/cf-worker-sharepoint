import { createExecutionContext, waitOnExecutionContext } from "cloudflare:test";
import { env as testEnv } from "cloudflare:workers";
import worker from "../src/index";
import type { Env } from "../src/env";
import { mockSharePoint } from "../src/mock/graph-mock";
import { runSync } from "../src/sync/delta";

export const env = testEnv as unknown as Env;
export const mock = mockSharePoint;

export async function call(path: string, init?: RequestInit, overrides: Partial<Env> = {}): Promise<Response> {
  const ctx = createExecutionContext();
  const res = await worker.fetch(new Request(`https://worker.test${path}`, init), { ...env, ...overrides }, ctx);
  await waitOnExecutionContext(ctx);
  return res;
}

export async function json<T = any>(path: string, init?: RequestInit, overrides: Partial<Env> = {}): Promise<{ status: number; body: T; res: Response }> {
  const res = await call(path, init, overrides);
  return { status: res.status, body: (await res.json()) as T, res };
}

export async function syncAll() {
  for (let i = 0; i < 50; i++) {
    const r = await runSync(env);
    if (r.status === "failed") throw new Error(`Sync failed: ${r.error}`);
    if (r.status === "completed") return r;
  }
  throw new Error("Sync did not complete");
}

export const q = (path: string) => `?path=${encodeURIComponent(path)}`;

export interface UnzippedEntry {
  name: string;
  data: Uint8Array;
  method: number;
}

/** Minimal zip reader (stored entries, no Zip64) for checking archives in tests. */
export function unzip(buf: Uint8Array): UnzippedEntry[] {
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("No end of central directory record");
  const count = view.getUint16(eocd + 10, true);
  let p = view.getUint32(eocd + 16, true);
  const dec = new TextDecoder();
  const out: UnzippedEntry[] = [];
  for (let n = 0; n < count; n++) {
    if (view.getUint32(p, true) !== 0x02014b50) throw new Error("Bad central directory entry");
    const method = view.getUint16(p + 10, true);
    const compressedSize = view.getUint32(p + 20, true);
    const nameLen = view.getUint16(p + 28, true);
    const extraLen = view.getUint16(p + 30, true);
    const commentLen = view.getUint16(p + 32, true);
    const localOffset = view.getUint32(p + 42, true);
    const name = dec.decode(buf.subarray(p + 46, p + 46 + nameLen));
    const lNameLen = view.getUint16(localOffset + 26, true);
    const lExtraLen = view.getUint16(localOffset + 28, true);
    const start = localOffset + 30 + lNameLen + lExtraLen;
    out.push({ name, method, data: buf.subarray(start, start + compressedSize) });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}
