/**
 * In-memory fake of the Microsoft Graph endpoints this Worker uses, backed by
 * a small fake SharePoint document library. Enabled with MOCK_GRAPH=true for
 * local development, and used by the test suite (which also mutates it to
 * exercise the delta sync).
 *
 * It imitates the SharePoint behaviours that matter here: delta only on the
 * drive root, no parentReference.path in delta responses, paging via
 * nextLink/deltaLink, 410 for expired delta tokens, pre-authenticated download
 * URLs with Range support, sharing links resolved through /shares, $batch and
 * 429 throttling with Retry-After.
 */
import { SAMPLE_JPG, SAMPLE_MP4, SAMPLE_PNG } from "./sample-media";

export const MOCK_HOST = "contoso.sharepoint.com";
export const MOCK_SITE_PATH = "/sites/Projects";
export const MOCK_SITE_URL = `https://${MOCK_HOST}${MOCK_SITE_PATH}`;
export const MOCK_LIBRARY_SEGMENT = "Shared Documents";
export const MOCK_LIBRARY_NAME = "Dokument";
export const MOCK_TOKEN = "mock-access-token";

const GRAPH_ORIGIN = "https://graph.microsoft.com";
const LOGIN_ORIGIN = "https://login.microsoftonline.com";
const SITE_ID = `${MOCK_HOST},6f1e0a52-0000-4000-8000-000000000001,0b3d2c1a-0000-4000-8000-000000000002`;
const DRIVE_ID = "b!mockDriveDokument000000000000000000000000000000000000000000";
const OTHER_DRIVE_ID = "b!mockDriveArkivbibliotek0000000000000000000000000000000000000";

interface MockItem {
  id: string;
  name: string;
  parentId: string | null;
  folder: boolean;
  mimeType?: string;
  content?: Uint8Array;
  sizeHint?: number;
  created: string;
  modified: string;
  modifiedBy: string;
  version: number;
  seq: number;
}

interface DeletedEntry {
  id: string;
  seq: number;
}

const enc = new TextEncoder();

function b64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

const MIME: Record<string, string> = {
  pdf: "application/pdf",
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  svg: "image/svg+xml",
  mp4: "video/mp4",
  txt: "text/plain",
  html: "text/html",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

const extOf = (name: string) => (name.includes(".") ? name.split(".").pop()!.toLowerCase() : "");

/** A minimal, valid single-page PDF showing `title`. */
export function makePdf(title: string): Uint8Array {
  const safe = title.replace(/[^\x20-\x7e]/g, "?").replace(/([()\\])/g, "\\$1");
  const stream = `BT /F1 28 Tf 72 720 Td (${safe}) Tj ET\nBT /F1 12 Tf 72 690 Td (Mock SharePoint document) Tj ET`;
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ];
  let pdf = "%PDF-1.4\n";
  const offsets: number[] = [];
  objects.forEach((obj, i) => {
    offsets.push(pdf.length);
    pdf += `${i + 1} 0 obj\n${obj}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, "0")} 00000 n \n`;
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return enc.encode(pdf);
}

function defaultContent(name: string): Uint8Array {
  switch (extOf(name)) {
    case "pdf":
      return makePdf(name);
    case "png":
      return b64ToBytes(SAMPLE_PNG);
    case "jpg":
    case "jpeg":
      return b64ToBytes(SAMPLE_JPG);
    case "mp4":
      return b64ToBytes(SAMPLE_MP4);
    case "svg":
      return enc.encode(
        `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="120"><rect width="200" height="120" fill="#0ea5e9"/><text x="20" y="66" fill="#fff" font-size="20">${name}</text><script>alert("svg script ran")</script></svg>`,
      );
    case "html":
      return enc.encode(`<!doctype html><title>${name}</title><script>alert("html script ran")</script><p>Uploaded HTML</p>`);
    case "txt":
      return enc.encode(`This is ${name} from the mock SharePoint library.\nRad två med åäö.\n`);
    default: {
      const bytes = new Uint8Array(2048 + (name.length % 7) * 333);
      for (let i = 0; i < bytes.length; i++) bytes[i] = (i * 31 + name.length) & 0xff;
      return bytes;
    }
  }
}

function jsonResponse(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json", ...headers },
  });
}

const graphError = (status: number, code: string, message: string, headers: Record<string, string> = {}) =>
  jsonResponse({ error: { code, message } }, status, headers);

export class MockSharePoint {
  items = new Map<string, MockItem>();
  deleted: DeletedEntry[] = [];
  seq = 0;
  private nextId = 1;
  private enumerations = new Map<string, unknown[]>();

  /** Delta page size. Small values exercise paging. */
  pageSize = 50;
  /** Delta tokens issued before this sequence number answer 410 Gone. */
  expireTokensBefore = 0;
  /** Number of upcoming Graph calls that answer 429. */
  throttleNext = 0;
  /** Return children before their parents in full delta enumerations. */
  childrenFirst = false;
  shareLinks = new Map<string, string>();
  requestLog: string[] = [];

  constructor() {
    this.reset();
  }

  reset() {
    this.items.clear();
    this.deleted = [];
    this.seq = 0;
    this.nextId = 1;
    this.enumerations.clear();
    this.pageSize = 50;
    this.expireTokensBefore = 0;
    this.throttleNext = 0;
    this.childrenFirst = false;
    this.shareLinks.clear();
    this.requestLog = [];
    this.seed();
  }

  // --- Library manipulation ---------------------------------------------------

  get rootId() {
    return "MOCKROOT";
  }

  private touch(item: MockItem) {
    item.seq = ++this.seq;
    item.version++;
  }

  private stamp(i: number) {
    return new Date(Date.UTC(2026, 0, 5) + i * 36e5 * 7).toISOString().replace(/\.\d{3}Z$/, "Z");
  }

  /** Create a file (or folder, if the path ends with "/") and any missing parent folders. */
  add(path: string, opts: { content?: Uint8Array | string; size?: number; modifiedBy?: string } = {}): MockItem {
    const isFolder = path.endsWith("/");
    const parts = path.split("/").filter(Boolean);
    let parentId = this.rootId;
    for (let i = 0; i < parts.length; i++) {
      const name = parts[i]!;
      const last = i === parts.length - 1;
      const existing = this.childByName(parentId, name);
      if (existing && (!last || isFolder)) {
        parentId = existing.id;
        continue;
      }
      const n = this.nextId++;
      const item: MockItem = {
        id: `MOCK${String(n).padStart(6, "0")}`,
        name,
        parentId,
        folder: !last || isFolder,
        created: this.stamp(n),
        modified: this.stamp(n + 3),
        modifiedBy: opts.modifiedBy ?? (n % 3 === 0 ? "Erik Lund" : "Anna Svensson"),
        version: 0,
        seq: 0,
      };
      if (!item.folder) {
        item.mimeType = MIME[extOf(name)] ?? "application/octet-stream";
        if (opts.content !== undefined) {
          item.content = typeof opts.content === "string" ? enc.encode(opts.content) : opts.content;
        } else if (opts.size !== undefined) {
          item.sizeHint = opts.size;
        }
      }
      this.touch(item);
      this.items.set(item.id, item);
      parentId = item.id;
      if (last) return item;
    }
    return this.items.get(parentId)!;
  }

  byPath(path: string): MockItem | undefined {
    let current: MockItem | undefined = this.items.get(this.rootId);
    for (const name of path.split("/").filter(Boolean)) {
      if (!current) return undefined;
      current = this.childByName(current.id, name);
    }
    return current;
  }

  mustPath(path: string): MockItem {
    const item = this.byPath(path);
    if (!item) throw new Error(`Mock item not found: ${path}`);
    return item;
  }

  rename(path: string, newName: string) {
    const item = this.mustPath(path);
    item.name = newName;
    item.modified = new Date().toISOString();
    this.touch(item);
  }

  move(path: string, newParentPath: string) {
    const item = this.mustPath(path);
    item.parentId = this.mustPath(newParentPath).id;
    this.touch(item);
  }

  /** Deletes an item. Like SharePoint, descendants of a deleted folder are reported deleted too. */
  remove(path: string) {
    const item = this.mustPath(path);
    for (const id of [...this.descendantIds(item.id), item.id]) {
      this.items.delete(id);
      this.deleted.push({ id, seq: ++this.seq });
    }
  }

  setContent(path: string, content: string | Uint8Array) {
    const item = this.mustPath(path);
    item.content = typeof content === "string" ? enc.encode(content) : content;
    item.sizeHint = undefined;
    item.modified = new Date().toISOString();
    this.touch(item);
  }

  private childByName(parentId: string, name: string) {
    const lower = name.toLowerCase();
    for (const item of this.items.values()) {
      if (item.parentId === parentId && item.name.toLowerCase() === lower) return item;
    }
    return undefined;
  }

  private descendantIds(id: string): string[] {
    const out: string[] = [];
    for (const item of this.items.values()) {
      if (item.parentId === id) out.push(...this.descendantIds(item.id), item.id);
    }
    return out;
  }

  private pathOf(item: MockItem): string {
    const names: string[] = [];
    let current: MockItem | undefined = item;
    while (current && current.parentId) {
      names.unshift(current.name);
      current = this.items.get(current.parentId);
    }
    return names.length ? `/${names.join("/")}` : "";
  }

  contentOf(item: MockItem): Uint8Array {
    if (!item.content) {
      const base = defaultContent(item.name);
      if (item.sizeHint !== undefined && item.sizeHint !== base.length) {
        const sized = new Uint8Array(item.sizeHint);
        sized.set(base.subarray(0, Math.min(base.length, sized.length)));
        item.content = sized;
      } else {
        item.content = base;
      }
    }
    return item.content;
  }

  private seed() {
    this.items.set(this.rootId, {
      id: this.rootId,
      name: "root",
      parentId: null,
      folder: true,
      created: this.stamp(0),
      modified: this.stamp(0),
      modifiedBy: "System",
      version: 0,
      seq: ++this.seq,
    });

    this.add("Clients/Acme/README.txt");
    for (const q of ["Q1", "Q2", "Q3"]) this.add(`Clients/Acme/Rapporter 2026/${q}-rapport.pdf`);
    this.add("Clients/Acme/Rapporter 2026/Årsredovisning 2025.pdf");
    this.add("Clients/Acme/Bilder/logo.png");
    this.add("Clients/Acme/Bilder/kontor.jpg");
    this.add("Clients/Acme/Bilder/diagram.svg");
    this.add("Clients/Acme/Video/intro.mp4");
    this.add("Clients/Acme/Avtal & offerter/Offert #12 (rev 2).pdf");
    this.add("Clients/Acme/Avtal & offerter/Ramavtal 100%.docx");
    this.add("Clients/Acme/Avtal & offerter/Tom mapp/");
    for (let i = 1; i <= 120; i++) {
      this.add(`Clients/Acme/Arkiv/Faktura ${String(i).padStart(3, "0")}.pdf`);
    }
    this.add("Clients/Acme/Webb/index.html");
    this.add("Clients/Other/Hemligt.pdf");
    this.add("Interna dokument/Budget.xlsx");

    this.shareLinks.set(
      `https://${MOCK_HOST}/:f:/s/Projects/EaRapporter2026SharedLinkAbc?e=x1Y2z3`,
      this.mustPath("Clients/Acme/Rapporter 2026").id,
    );
    this.shareLinks.set(
      `https://${MOCK_HOST}/:b:/s/Projects/EfQ3RapportSharedLinkXyz?e=a9B8c7`,
      this.mustPath("Clients/Acme/Rapporter 2026/Q3-rapport.pdf").id,
    );
    this.shareLinks.set(
      `https://${MOCK_HOST}/:b:/s/Projects/EfHemligtOutsideRootLink?e=q1w2e3`,
      this.mustPath("Clients/Other/Hemligt.pdf").id,
    );
  }

  // --- Graph representations --------------------------------------------------

  webUrlOf(item: MockItem): string {
    const segments = [MOCK_LIBRARY_SEGMENT, ...this.pathOf(item).split("/").filter(Boolean)];
    return `${MOCK_SITE_URL}/${segments.map(encodeURIComponent).join("/")}`;
  }

  downloadUrlOf(item: MockItem): string {
    return `https://${MOCK_HOST}${MOCK_SITE_PATH}/_layouts/15/download.aspx?UniqueId=${item.id}&tempauth=mock.${item.version}`;
  }

  toDriveItem(item: MockItem, opts: { withPath?: boolean; withDownloadUrl?: boolean } = {}) {
    const isRoot = item.parentId === null;
    const parent = item.parentId ? this.items.get(item.parentId) : undefined;
    const out: Record<string, unknown> = {
      id: item.id,
      name: item.name,
      eTag: `"{${item.id}},${item.version}"`,
      cTag: `"c:{${item.id}},${item.version}"`,
      createdDateTime: item.created,
      lastModifiedDateTime: item.modified,
      lastModifiedBy: { user: { displayName: item.modifiedBy } },
      webUrl: this.webUrlOf(item),
      size: item.folder ? this.folderSize(item.id) : this.contentOf(item).length,
    };
    if (isRoot) out.root = {};
    if (parent) {
      out.parentReference = {
        driveId: DRIVE_ID,
        driveType: "documentLibrary",
        id: parent.id,
        ...(opts.withPath ? { path: `/drives/${DRIVE_ID}/root:${this.pathOf(parent)}` } : {}),
      };
    }
    if (item.folder) {
      out.folder = { childCount: [...this.items.values()].filter((i) => i.parentId === item.id).length };
    } else {
      out.file = { mimeType: item.mimeType };
      if (opts.withDownloadUrl) out["@microsoft.graph.downloadUrl"] = this.downloadUrlOf(item);
    }
    return out;
  }

  private folderSize(id: string): number {
    let total = 0;
    for (const item of this.items.values()) {
      if (item.parentId === id) total += item.folder ? this.folderSize(item.id) : this.contentOf(item).length;
    }
    return total;
  }

  /** Items in parent-before-child order (or reversed, to test orphan handling). */
  private orderedItems(): MockItem[] {
    const out: MockItem[] = [];
    const visit = (id: string) => {
      const item = this.items.get(id);
      if (!item) return;
      out.push(item);
      const children = [...this.items.values()].filter((i) => i.parentId === id).sort((a, b) => a.id.localeCompare(b.id));
      for (const child of children) visit(child.id);
    };
    visit(this.rootId);
    return this.childrenFirst ? out.reverse() : out;
  }

  // --- HTTP handling -----------------------------------------------------------

  handle = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const request = new Request(input as RequestInfo, init);
    const url = new URL(request.url);
    this.requestLog.push(`${request.method} ${url.pathname}${url.search}`);

    if (url.origin === LOGIN_ORIGIN) return this.handleLogin(request);
    if (url.origin === `https://${MOCK_HOST}`) return this.handleDownload(request, url);
    if (url.origin === GRAPH_ORIGIN) {
      if (request.headers.get("authorization") !== `Bearer ${MOCK_TOKEN}`) {
        return graphError(401, "InvalidAuthenticationToken", "Access token is empty or invalid.");
      }
      if (this.throttleNext > 0) {
        this.throttleNext--;
        return graphError(429, "TooManyRequests", "Too many requests", { "retry-after": "0" });
      }
      if (url.pathname === "/v1.0/$batch" && request.method === "POST") return this.handleBatch(request);
      return this.route(request.method, url);
    }
    return new Response(`Mock has no handler for ${url.origin}`, { status: 502 });
  };

  private async handleLogin(request: Request): Promise<Response> {
    const form = await request.formData();
    if (form.get("grant_type") !== "client_credentials" || !form.get("client_id")) {
      return jsonResponse({ error: "invalid_request", error_description: "Bad token request" }, 400);
    }
    if (!form.get("client_secret") && !form.get("client_assertion")) {
      return jsonResponse({ error: "invalid_client", error_description: "No credentials" }, 401);
    }
    return jsonResponse({ token_type: "Bearer", expires_in: 3599, access_token: MOCK_TOKEN });
  }

  private async handleBatch(request: Request): Promise<Response> {
    const { requests } = (await request.json()) as { requests: { id: string; method: string; url: string }[] };
    if (requests.length > 20) return graphError(400, "BadRequest", "Too many requests in batch");
    const responses = [];
    for (const r of requests) {
      const res = await this.route(r.method, new URL(`${GRAPH_ORIGIN}/v1.0${r.url.startsWith("/") ? "" : "/"}${r.url}`));
      const headers: Record<string, string> = {};
      res.headers.forEach((v, k) => (headers[k] = v));
      const text = await res.text();
      responses.push({ id: r.id, status: res.status, headers, body: text ? JSON.parse(text) : undefined });
    }
    return jsonResponse({ responses });
  }

  private route(method: string, url: URL): Response {
    if (method !== "GET") return graphError(405, "MethodNotAllowed", "Read-only mock");
    const path = url.pathname.replace(/^\/v1\.0/, "");
    let m: RegExpMatchArray | null;

    if ((m = path.match(/^\/sites\/([^/]+?)(?::(.*?):?)?$/))) {
      const host = m[1]!;
      const sitePath = m[2] ? decodeURIComponent(m[2]) : "";
      if (host === SITE_ID || (host === MOCK_HOST && sitePath.toLowerCase() === MOCK_SITE_PATH.toLowerCase())) {
        return jsonResponse({ id: SITE_ID, webUrl: MOCK_SITE_URL, displayName: "Projects" });
      }
      return graphError(404, "itemNotFound", "Requested site could not be found");
    }
    if (path === `/sites/${SITE_ID}/drives`) {
      return jsonResponse({
        value: [
          { id: DRIVE_ID, name: MOCK_LIBRARY_NAME, webUrl: `${MOCK_SITE_URL}/Shared%20Documents`, driveType: "documentLibrary" },
          { id: OTHER_DRIVE_ID, name: "Arkivbibliotek", webUrl: `${MOCK_SITE_URL}/Arkivbibliotek`, driveType: "documentLibrary" },
        ],
      });
    }
    if (path === `/sites/${SITE_ID}/drive`) {
      return jsonResponse({ id: DRIVE_ID, name: MOCK_LIBRARY_NAME, webUrl: `${MOCK_SITE_URL}/Shared%20Documents`, driveType: "documentLibrary" });
    }
    if (path === `/drives/${DRIVE_ID}/root/delta`) return this.delta(url);

    if ((m = path.match(/^\/drives\/([^/]+)\/items\/([^/]+)(\/.*)?$/))) {
      if (m[1] !== DRIVE_ID) return graphError(404, "itemNotFound", "Drive not found");
      const item = this.items.get(decodeURIComponent(m[2]!));
      if (!item) return graphError(404, "itemNotFound", "The resource could not be found.");
      const rest = m[3] ?? "";
      if (rest === "") return jsonResponse(this.toDriveItem(item, { withPath: true, withDownloadUrl: true }));
      if (rest === "/content") {
        if (item.folder) return graphError(400, "invalidRequest", "Folders have no content");
        return new Response(null, { status: 302, headers: { location: this.downloadUrlOf(item) } });
      }
      const thumb = rest.match(/^\/thumbnails\/0\/(small|medium|large|c\d+x\d+)\/content$/);
      if (thumb) {
        if (item.folder || !/^(image|video)\/|pdf$/.test(item.mimeType ?? "")) {
          return graphError(404, "itemNotFound", "Thumbnail not available");
        }
        return new Response(null, {
          status: 302,
          headers: { location: `https://${MOCK_HOST}/_api/v2.1/thumbnail?UniqueId=${item.id}&size=${thumb[1]}` },
        });
      }
      return graphError(400, "invalidRequest", `Unsupported mock path ${rest}`);
    }

    if ((m = path.match(/^\/shares\/([^/]+)\/driveItem$/))) {
      const shareId = decodeURIComponent(m[1]!);
      if (!shareId.startsWith("u!")) return graphError(400, "invalidRequest", "Invalid share id");
      const b64 = shareId.slice(2).replace(/-/g, "+").replace(/_/g, "/");
      const sharedUrl = new TextDecoder().decode(b64ToBytes(b64 + "=".repeat((4 - (b64.length % 4)) % 4)));
      let id = this.shareLinks.get(sharedUrl);
      if (!id) {
        for (const item of this.items.values()) {
          if (this.webUrlOf(item).toLowerCase() === sharedUrl.split("?")[0]!.toLowerCase()) id = item.id;
        }
      }
      const item = id ? this.items.get(id) : undefined;
      if (!item) return graphError(404, "itemNotFound", "The sharing link could not be resolved");
      return jsonResponse(this.toDriveItem(item, { withPath: true }));
    }

    return graphError(400, "invalidRequest", `Unsupported mock path ${path}`);
  }

  private delta(url: URL): Response {
    const cursor = url.searchParams.get("cursor");
    const token = url.searchParams.get("token");
    let entries: unknown[];
    let cursorId: string;
    let skip = Number(url.searchParams.get("skip") ?? 0);

    if (cursor) {
      const cached = this.enumerations.get(cursor);
      if (!cached) return graphError(410, "resyncRequired", "Cursor expired");
      entries = cached;
      cursorId = cursor;
    } else {
      const since = token === null ? null : Number(token);
      if (since !== null && since < this.expireTokensBefore) {
        return graphError(410, "resyncRequired", "Resync required. Replace any local items with the server's version.", {
          location: `${GRAPH_ORIGIN}/v1.0/drives/${DRIVE_ID}/root/delta`,
        });
      }
      const snapshotSeq = this.seq;
      if (since === null) {
        entries = this.orderedItems().map((i) => this.toDriveItem(i));
      } else {
        const changed = [...this.items.values()]
          .filter((i) => i.seq > since)
          .map((i) => ({ seq: i.seq, value: this.toDriveItem(i) as unknown }));
        const removed = this.deleted
          .filter((d) => d.seq > since)
          .map((d) => ({ seq: d.seq, value: { id: d.id, deleted: { state: "deleted" } } as unknown }));
        entries = [...changed, ...removed].sort((a, b) => a.seq - b.seq).map((e) => e.value);
      }
      cursorId = `c${crypto.randomUUID().replace(/-/g, "")}`;
      this.enumerations.set(cursorId, entries);
      (entries as unknown as { snapshotSeq: number }).snapshotSeq = snapshotSeq;
      skip = 0;
    }

    const page = entries.slice(skip, skip + this.pageSize);
    const base = `${GRAPH_ORIGIN}/v1.0/drives/${DRIVE_ID}/root/delta`;
    const body: Record<string, unknown> = { value: page };
    if (skip + this.pageSize < entries.length) {
      body["@odata.nextLink"] = `${base}?cursor=${cursorId}&skip=${skip + this.pageSize}`;
    } else {
      const snapshotSeq = (entries as unknown as { snapshotSeq: number }).snapshotSeq;
      body["@odata.deltaLink"] = `${base}?token=${snapshotSeq}`;
      this.enumerations.delete(cursorId);
    }
    return jsonResponse(body);
  }

  private handleDownload(request: Request, url: URL): Response {
    const id = url.searchParams.get("UniqueId") ?? "";
    const item = this.items.get(id);
    if (!item || item.folder) return new Response("Not found", { status: 404 });

    if (url.pathname.endsWith("/thumbnail")) {
      return new Response(b64ToBytes(SAMPLE_PNG), { headers: { "content-type": "image/png" } });
    }
    if (url.searchParams.get("tempauth") !== `mock.${item.version}`) {
      return new Response("Access denied: download URL expired", { status: 401 });
    }

    const body = this.contentOf(item);
    const headers: Record<string, string> = {
      "content-type": item.mimeType ?? "application/octet-stream",
      "accept-ranges": "bytes",
      "content-disposition": `attachment; filename*=UTF-8''${encodeURIComponent(item.name)}`,
      etag: `"{${item.id}},${item.version}"`,
    };
    const range = request.headers.get("range")?.match(/^bytes=(\d*)-(\d*)$/);
    if (range && (range[1] || range[2])) {
      let start: number;
      let end: number;
      if (range[1] === "") {
        start = Math.max(0, body.length - Number(range[2]));
        end = body.length - 1;
      } else {
        start = Number(range[1]);
        end = range[2] ? Math.min(Number(range[2]), body.length - 1) : body.length - 1;
      }
      if (start >= body.length || start > end) {
        return new Response(null, { status: 416, headers: { "content-range": `bytes */${body.length}` } });
      }
      const slice = body.slice(start, end + 1);
      return new Response(slice, {
        status: 206,
        headers: { ...headers, "content-length": String(slice.length), "content-range": `bytes ${start}-${end}/${body.length}` },
      });
    }
    return new Response(body.slice(), { headers: { ...headers, "content-length": String(body.length) } });
  }
}

export const mockSharePoint = new MockSharePoint();
export const mockFetch: typeof fetch = (input, init) => mockSharePoint.handle(input, init);
