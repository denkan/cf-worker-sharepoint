import { getConfig, type Config, type Env } from "../env";
import { ApiError } from "../errors";
import { mockFetch } from "../mock/graph-mock";
import { getGraphToken } from "./auth";

export class GraphError extends ApiError {
  constructor(
    readonly graphStatus: number,
    readonly graphCode: string | undefined,
    message: string,
  ) {
    const mapped =
      graphStatus === 404
        ? ([404, "NOT_FOUND"] as const)
        : graphStatus === 401 || graphStatus === 403
          ? ([502, "UPSTREAM_AUTH_ERROR"] as const)
          : ([502, "UPSTREAM_ERROR"] as const);
    super(mapped[0], mapped[1], message, { graphStatus, graphCode });
  }
}

const RETRYABLE = new Set([429, 502, 503, 504]);
const MAX_ATTEMPTS = 4;
const MAX_RETRY_AFTER_MS = 30_000;

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

export interface BatchRequest {
  id: string;
  method: "GET";
  url: string;
}

export interface BatchResponse<T = unknown> {
  id: string;
  status: number;
  headers?: Record<string, string>;
  body?: T;
}

export class GraphClient {
  readonly config: Config;
  private readonly fetcher: typeof fetch;

  constructor(private readonly env: Env) {
    this.config = getConfig(env);
    this.fetcher = this.config.mock ? mockFetch : fetch.bind(globalThis);
  }

  url(pathOrUrl: string): string {
    return /^https?:\/\//.test(pathOrUrl) ? pathOrUrl : `${this.config.graphBaseUrl}${pathOrUrl}`;
  }

  /** Authenticated Graph request with retries on throttling and transient errors. */
  async request(pathOrUrl: string, init: RequestInit = {}): Promise<Response> {
    const url = this.url(pathOrUrl);
    for (let attempt = 1; ; attempt++) {
      const token = await getGraphToken(this.env, this.config, this.fetcher);
      const headers = new Headers(init.headers);
      headers.set("authorization", `Bearer ${token}`);
      const res = await this.fetcher(url, { ...init, headers });

      if (!RETRYABLE.has(res.status) || attempt >= MAX_ATTEMPTS) return res;

      const retryAfter = Number.parseFloat(res.headers.get("retry-after") ?? "");
      const waitMs = Number.isFinite(retryAfter)
        ? retryAfter * 1000
        : Math.min(2 ** (attempt - 1) * 1000, 8000) * (0.75 + Math.random() * 0.5);
      if (waitMs > MAX_RETRY_AFTER_MS) return res;
      await res.body?.cancel();
      await sleep(waitMs);
    }
  }

  async json<T>(pathOrUrl: string, init?: RequestInit): Promise<T> {
    const res = await this.request(pathOrUrl, init);
    if (!res.ok) throw await toGraphError(res);
    return (await res.json()) as T;
  }

  /** JSON $batch (max 20 requests per call). Individual failures are returned, not thrown. */
  async batch<T>(requests: BatchRequest[]): Promise<BatchResponse<T>[]> {
    const out: BatchResponse<T>[] = [];
    for (let i = 0; i < requests.length; i += 20) {
      const chunk = requests.slice(i, i + 20);
      const body = await this.json<{ responses: BatchResponse<T>[] }>("/$batch", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ requests: chunk }),
      });
      out.push(...body.responses);
    }
    return out;
  }

  /** Fetch a pre-authenticated URL (e.g. @microsoft.graph.downloadUrl) without a token. */
  fetchUnauthenticated(url: string, init?: RequestInit): Promise<Response> {
    return this.fetcher(url, init);
  }

  /** Short-lived, pre-authenticated download URL for a file. */
  async getDownloadUrl(driveId: string, itemId: string): Promise<string> {
    const item = await this.json<{ "@microsoft.graph.downloadUrl"?: string }>(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`,
    );
    const url = item["@microsoft.graph.downloadUrl"];
    if (url) return url;

    // Fallback: /content answers with a redirect to the same kind of URL.
    const res = await this.request(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`,
      { redirect: "manual" },
    );
    const location = res.headers.get("location");
    await res.body?.cancel();
    if (res.status >= 300 && res.status < 400 && location) return location;
    if (!res.ok) throw await toGraphError(res);
    throw new ApiError(502, "UPSTREAM_ERROR", "Graph returned no download URL for the item");
  }
}

export async function toGraphError(res: Response): Promise<GraphError> {
  let code: string | undefined;
  let message = `Graph request failed with ${res.status}`;
  try {
    const body = (await res.json()) as { error?: { code?: string; message?: string } };
    code = body.error?.code;
    if (body.error?.message) message = body.error.message;
  } catch {
    // Non-JSON error body
  }
  return new GraphError(res.status, code, message);
}
