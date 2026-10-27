import { describe, expect, it } from "vitest";
import { runSync } from "../src/sync/delta";
import { env, json, mock, q, syncAll } from "./helpers";

async function paths(folder = "/") {
  const { body } = await json(`/api/items${q(folder)}&limit=1000`);
  return (body.items as { path: string }[]).map((i) => i.path);
}

async function pathOfId(id: string) {
  const row = await env.DB.prepare("SELECT path FROM items WHERE id = ?").bind(id).first<{ path: string | null }>();
  return row?.path ?? null;
}

describe("delta sync", () => {
  it("mirrors the library on the first run", async () => {
    const result = await syncAll();
    expect(result.mode).toBe("full");
    const total = await env.DB.prepare("SELECT COUNT(*) AS n FROM items").first<{ n: number }>();
    expect(total!.n).toBe(mock.items.size);
    expect(await paths()).toEqual([
      "/Arkiv",
      "/Avtal & offerter",
      "/Bilder",
      "/Rapporter 2026",
      "/Video",
      "/Webb",
      "/README.txt",
    ]);
    const status = await json("/api/status");
    expect(status.body.index.ready).toBe(true);
    expect(status.body.drive.name).toBe("Dokument");
  });

  it("resumes across runs when a run stops early", async () => {
    mock.pageSize = 10;
    const first = await runSync(env, { maxPages: 3 });
    expect(first.status).toBe("partial");
    expect((await json("/api/status")).body.index.ready).toBe(false);
    let runs = 1;
    while ((await runSync(env, { maxPages: 3 })).status === "partial") runs++;
    expect(runs).toBeGreaterThan(3);
    const total = await env.DB.prepare("SELECT COUNT(*) AS n FROM items").first<{ n: number }>();
    expect(total!.n).toBe(mock.items.size);
    expect((await json("/api/status")).body.index.ready).toBe(true);
  });

  it("handles children arriving before their parents", async () => {
    mock.childrenFirst = true;
    mock.pageSize = 7;
    await syncAll();
    const orphans = await env.DB.prepare("SELECT COUNT(*) AS n FROM items WHERE path IS NULL").first<{ n: number }>();
    expect(orphans!.n).toBe(0);
    expect(await pathOfId(mock.mustPath("Clients/Acme/Bilder/logo.png").id)).toBe("/Clients/Acme/Bilder/logo.png");
  });

  it("applies renames of folders to everything below them", async () => {
    await syncAll();
    const q3 = mock.mustPath("Clients/Acme/Rapporter 2026/Q3-rapport.pdf");
    mock.rename("Clients/Acme/Rapporter 2026", "Rapporter 2026 (arkiverat)");
    const r = await syncAll();
    expect(r.mode).toBe("incremental");
    expect(r.upserted).toBe(1); // only the folder is reported, like SharePoint
    expect(await pathOfId(q3.id)).toBe("/Clients/Acme/Rapporter 2026 (arkiverat)/Q3-rapport.pdf");
    expect((await json(`/api/item${q("/Rapporter 2026 (arkiverat)/Q3-rapport.pdf")}`)).status).toBe(200);
    expect((await json(`/api/item${q("/Rapporter 2026/Q3-rapport.pdf")}`)).status).toBe(404);
  });

  it("handles moves into and out of the root folder", async () => {
    await syncAll();
    mock.move("Clients/Acme/Bilder", "Clients/Other");
    mock.move("Clients/Other/Hemligt.pdf", "Clients/Acme");
    await syncAll();
    expect(await paths()).not.toContain("/Bilder");
    expect(await paths()).toContain("/Hemligt.pdf");
    expect((await json(`/api/item${q("/Bilder/logo.png")}`)).status).toBe(404);

    mock.move("Clients/Other/Bilder", "Clients/Acme/Video");
    await syncAll();
    expect(await paths("/Video/Bilder")).toEqual(["/Video/Bilder/diagram.svg", "/Video/Bilder/kontor.jpg", "/Video/Bilder/logo.png"]);
  });

  it("removes deleted folders and their contents", async () => {
    await syncAll();
    mock.remove("Clients/Acme/Arkiv");
    const r = await syncAll();
    expect(r.deleted).toBe(121);
    expect(await paths()).not.toContain("/Arkiv");
    const { body } = await json(`/api/search?q=faktura`);
    expect(body.total).toBe(0);
  });

  it("picks up new and changed files", async () => {
    await syncAll();
    mock.add("Clients/Acme/Nytt/Protokoll.pdf");
    mock.setContent("Clients/Acme/README.txt", "updated content, longer than before");
    await syncAll();
    expect(await paths("/Nytt")).toEqual(["/Nytt/Protokoll.pdf"]);
    const readme = await json(`/api/item${q("/README.txt")}`);
    expect(readme.body.size).toBe("updated content, longer than before".length);
  });

  it("starts over when the delta token has expired, removing stale rows", async () => {
    await syncAll();
    mock.remove("Clients/Acme/Webb");
    // Simulate the deletion being lost: the token predating it is no longer valid.
    mock.deleted = [];
    mock.expireTokensBefore = mock.seq + 1;
    const r = await syncAll();
    expect(r.mode).toBe("full");
    expect(await paths()).not.toContain("/Webb");
  });

  it("retries throttled Graph calls", async () => {
    mock.throttleNext = 2;
    const r = await syncAll();
    expect(r.status).toBe("completed");
    // The first Graph call (site lookup) was answered 429 twice, then succeeded.
    expect(mock.requestLog.filter((l) => l.includes(":/sites/Projects")).length).toBe(3);
  });

  it("runs only one sync at a time", async () => {
    const [a, b] = await Promise.all([runSync(env), runSync(env)]);
    expect([a.status, b.status].sort()).toEqual(["completed", "skipped"]);
  });

  it("reports a 503 before the first sync", async () => {
    const res = await json(`/api/items`);
    expect(res.status).toBe(503);
    expect(res.body.error.code).toBe("INDEX_NOT_READY");
  });
});
