import { beforeEach, describe, expect, it } from "vitest";
import { MOCK_HOST } from "../src/mock/graph-mock";
import { json, q, syncAll } from "./helpers";

beforeEach(async () => {
  await syncAll();
});

describe("GET /api/items", () => {
  it("lists the root folder with folders first", async () => {
    const { status, body } = await json("/api/items");
    expect(status).toBe(200);
    expect(body.folder.path).toBe("/");
    expect(body.total).toBe(7);
    expect(body.items[0]).toMatchObject({ name: "Arkiv", path: "/Arkiv", type: "folder", childCount: 120 });
    expect(body.items.at(-1)).toMatchObject({ name: "README.txt", type: "file", mimeType: "text/plain", extension: "txt" });
    expect(body.items.at(-1).links.content).toBe("/api/content?path=%2FREADME.txt");
    expect(body.index.ready).toBe(true);
  });

  it("paginates with limit and offset", async () => {
    const page1 = await json(`/api/items${q("/Arkiv")}&limit=50`);
    const page3 = await json(`/api/items${q("/Arkiv")}&limit=50&offset=100`);
    expect(page1.body.total).toBe(120);
    expect(page1.body.items).toHaveLength(50);
    expect(page1.body.items[0].name).toBe("Faktura 001.pdf");
    expect(page3.body.items).toHaveLength(20);
    expect(page3.body.items[19].name).toBe("Faktura 120.pdf");
  });

  it("sorts and filters", async () => {
    const { body } = await json(`/api/items${q("/Arkiv")}&sort=name&order=desc&limit=1`);
    expect(body.items[0].name).toBe("Faktura 120.pdf");
    const folders = await json(`/api/items?type=folder`);
    expect(folders.body.items.every((i: { type: string }) => i.type === "folder")).toBe(true);
    expect(folders.body.total).toBe(6);
  });

  it("rejects bad parameters and files", async () => {
    expect((await json("/api/items?limit=0")).status).toBe(400);
    expect((await json("/api/items?sort=color")).status).toBe(400);
    const file = await json(`/api/items${q("/README.txt")}`);
    expect(file.status).toBe(400);
    expect(file.body.error.code).toBe("NOT_A_FOLDER");
  });
});

describe("path formats", () => {
  const target = "/Rapporter 2026/Q3-rapport.pdf";
  const cases: [string, string, string][] = [
    ["relative with leading slash", "/Rapporter 2026/Q3-rapport.pdf", "relative"],
    ["relative without leading slash", "Rapporter 2026/Q3-rapport.pdf", "relative"],
    ["relative, different case and backslashes", "rapporter 2026\\q3-RAPPORT.pdf", "relative"],
    ["server-relative full path", "/sites/Projects/Shared Documents/Clients/Acme/Rapporter 2026/Q3-rapport.pdf", "fullPath"],
    ["library-relative by URL segment", "Shared Documents/Clients/Acme/Rapporter 2026/Q3-rapport.pdf", "fullPath"],
    ["library-relative by display name", "Dokument/Clients/Acme/Rapporter 2026/Q3-rapport.pdf", "fullPath"],
    ["browser URL", `https://${MOCK_HOST}/sites/Projects/Shared%20Documents/Clients/Acme/Rapporter%202026/Q3-rapport.pdf?web=1`, "url"],
    ["direct link", `https://${MOCK_HOST}/:b:/r/sites/Projects/Shared%20Documents/Clients/Acme/Rapporter%202026/Q3-rapport.pdf?csf=1&web=1`, "url"],
    ["sharing link", `https://${MOCK_HOST}/:b:/s/Projects/EfQ3RapportSharedLinkXyz?e=a9B8c7`, "sharingLink"],
  ];
  for (const [label, input, kind] of cases) {
    it(`accepts a ${label}`, async () => {
      const { status, body } = await json(`/api/resolve${q(input)}`);
      expect(status, JSON.stringify(body)).toBe(200);
      expect(body.item.path).toBe(target);
      expect(body.matchedAs).toBe(kind);
    });
  }

  it("accepts a library view URL for a folder", async () => {
    const url = `https://${MOCK_HOST}/sites/Projects/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FProjects%2FShared%20Documents%2FClients%2FAcme%2FBilder&viewid=abc`;
    const { status, body } = await json(`/api/items${q(url)}`);
    expect(status).toBe(200);
    expect(body.folder.path).toBe("/Bilder");
    expect(body.total).toBe(3);
  });

  it("accepts a sharing link to a folder", async () => {
    const { body } = await json(`/api/items${q(`https://${MOCK_HOST}/:f:/s/Projects/EaRapporter2026SharedLinkAbc?e=x1Y2z3`)}`);
    expect(body.folder.path).toBe("/Rapporter 2026");
    expect(body.total).toBe(4);
  });

  it("handles special characters in names", async () => {
    const { status, body } = await json(`/api/item${q("/Avtal & offerter/Offert #12 (rev 2).pdf")}`);
    expect(status).toBe(200);
    expect(body.name).toBe("Offert #12 (rev 2).pdf");
    const pct = await json(`/api/item${q("/Avtal & offerter/Ramavtal 100%.docx")}`);
    expect(pct.status).toBe(200);
  });

  it("never exposes anything outside the root", async () => {
    for (const input of [
      "/sites/Projects/Shared Documents/Clients/Other/Hemligt.pdf",
      "Dokument/Interna dokument/Budget.xlsx",
      `https://${MOCK_HOST}/sites/Projects/Shared%20Documents/Clients/Other/Hemligt.pdf`,
      `https://${MOCK_HOST}/:b:/s/Projects/EfHemligtOutsideRootLink?e=q1w2e3`,
      "https://evil.example.com/sites/Projects/Shared%20Documents/Clients/Acme/README.txt",
      "/sites/Projects/Shared Documents",
    ]) {
      const res = await json(`/api/item${q(input)}`);
      expect(res.status, input).toBe(404);
    }
    const dotdot = await json(`/api/item${q("/Bilder/../../Other/Hemligt.pdf")}`);
    expect(dotdot.status).toBe(400);
  });
});

describe("GET /api/search", () => {
  it("matches word prefixes, accent-insensitively", async () => {
    const { body } = await json("/api/search?q=rapp");
    expect(body.items.map((i: { name: string }) => i.name).sort()).toEqual([
      "Q1-rapport.pdf",
      "Q2-rapport.pdf",
      "Q3-rapport.pdf",
      "Rapporter 2026",
    ]);
    const accents = await json("/api/search?q=arsredovisning");
    expect(accents.body.items[0].path).toBe("/Rapporter 2026/Årsredovisning 2025.pdf");
    const multi = await json("/api/search?q=q3%20rapport");
    expect(multi.body.total).toBe(1);
  });

  it("is scoped to the root and to the given folder", async () => {
    expect((await json("/api/search?q=hemligt")).body.total).toBe(0);
    expect((await json("/api/search?q=budget")).body.total).toBe(0);
    const scoped = await json(`/api/search?q=pdf${q("/Rapporter 2026").replace("?", "&")}`);
    expect(scoped.body.total).toBe(4);
    expect(scoped.body.scope).toBe("/Rapporter 2026");
  });

  it("paginates and filters by metadata", async () => {
    const page = await json("/api/search?q=faktura&limit=25&offset=100&sort=name");
    expect(page.body.total).toBe(120);
    expect(page.body.items).toHaveLength(20);
    const images = await json("/api/search?ext=png,jpg");
    expect(images.body.items.map((i: { name: string }) => i.name).sort()).toEqual(["kontor.jpg", "logo.png"]);
    const folders = await json("/api/search?q=mapp&type=folder");
    expect(folders.body.items[0].path).toBe("/Avtal & offerter/Tom mapp");
    const none = await json("/api/search?q=---");
    expect(none.body.total).toBe(0);
  });
});

describe("misc", () => {
  it("serves an OpenAPI document", async () => {
    const { status, body } = await json("/openapi.json");
    expect(status).toBe(200);
    expect(Object.keys(body.paths)).toEqual(
      expect.arrayContaining(["/api/items", "/api/search", "/api/content", "/api/zip", "/api/zip/{id}"]),
    );
  });

  it("returns JSON 404 for unknown endpoints", async () => {
    const res = await json("/api/nope");
    expect(res.status).toBe(404);
    expect(res.body.error.code).toBe("NOT_FOUND");
  });
});
