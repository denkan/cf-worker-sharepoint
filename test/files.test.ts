import { beforeEach, describe, expect, it } from "vitest";
import { call, json, mock, q, syncAll, unzip } from "./helpers";

beforeEach(async () => {
  await syncAll();
});

const bytes = async (res: Response) => new Uint8Array(await res.arrayBuffer());

describe("GET /api/download", () => {
  it("redirects to a pre-authenticated SharePoint URL", async () => {
    const res = await call(`/api/download${q("/Rapporter 2026/Q3-rapport.pdf")}`);
    expect(res.status).toBe(302);
    const location = res.headers.get("location")!;
    expect(location).toContain("download.aspx");
    const file = await mock.handle(location);
    expect(file.status).toBe(200);
  });

  it("refuses folders", async () => {
    const { status, body } = await json(`/api/download${q("/Bilder")}`);
    expect(status).toBe(400);
    expect(body.error.code).toBe("NOT_A_FILE");
  });
});

describe("GET /api/content", () => {
  it("serves PDFs inline with the right headers", async () => {
    const res = await call(`/api/content${q("/Rapporter 2026/Årsredovisning 2025.pdf")}`);
    expect(res.status).toBe(200);
    expect(res.headers.get("content-type")).toBe("application/pdf");
    expect(res.headers.get("content-disposition")).toBe(
      `inline; filename="Arsredovisning 2025.pdf"; filename*=UTF-8''%C3%85rsredovisning%202025.pdf`,
    );
    expect(res.headers.get("accept-ranges")).toBe("bytes");
    const body = await bytes(res);
    expect(new TextDecoder().decode(body.subarray(0, 8))).toBe("%PDF-1.4");
    expect(Number(res.headers.get("content-length"))).toBe(body.length);
  });

  it("supports Range requests for video seeking", async () => {
    const full = await bytes(await call(`/api/content${q("/Video/intro.mp4")}`));
    const res = await call(`/api/content${q("/Video/intro.mp4")}`, { headers: { range: "bytes=100-199" } });
    expect(res.status).toBe(206);
    expect(res.headers.get("content-type")).toBe("video/mp4");
    expect(res.headers.get("content-range")).toBe(`bytes 100-199/${full.length}`);
    expect(await bytes(res)).toEqual(full.subarray(100, 200));

    const tail = await call(`/api/content${q("/Video/intro.mp4")}`, { headers: { range: "bytes=-10" } });
    expect(tail.status).toBe(206);
    expect(await bytes(tail)).toEqual(full.subarray(full.length - 10));
  });

  it("never renders active content inline", async () => {
    const html = await call(`/api/content${q("/Webb/index.html")}`);
    expect(html.headers.get("content-disposition")).toMatch(/^attachment;/);
    await html.body?.cancel();

    const svg = await call(`/api/content${q("/Bilder/diagram.svg")}`);
    expect(svg.headers.get("content-disposition")).toMatch(/^inline;/);
    expect(svg.headers.get("content-security-policy")).toContain("sandbox");
    expect(svg.headers.get("x-content-type-options")).toBe("nosniff");
    await svg.body?.cancel();
  });

  it("can force a download", async () => {
    const res = await call(`/api/content${q("/Bilder/logo.png")}&disposition=attachment`);
    expect(res.headers.get("content-disposition")).toMatch(/^attachment; filename="logo.png"/);
    await res.body?.cancel();
  });

  it("returns 404 when the file disappeared since the last sync", async () => {
    mock.remove("Clients/Acme/README.txt");
    const res = await json(`/api/content${q("/README.txt")}`);
    expect(res.status).toBe(404);
  });
});

describe("GET /api/thumbnail", () => {
  it("returns an image for images, PDFs and videos", async () => {
    for (const p of ["/Bilder/kontor.jpg", "/Rapporter 2026/Q1-rapport.pdf", "/Video/intro.mp4"]) {
      const res = await call(`/api/thumbnail${q(p)}&size=small`);
      expect(res.status, p).toBe(200);
      expect(res.headers.get("content-type")).toBe("image/png");
      await res.body?.cancel();
    }
  });

  it("returns 404 when SharePoint has no thumbnail", async () => {
    expect((await json(`/api/thumbnail${q("/README.txt")}`)).status).toBe(404);
  });
});

describe("zip", () => {
  it("zips a folder recursively, including empty folders", async () => {
    const prepared = await json("/api/zip", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ paths: ["/Avtal & offerter"] }),
    });
    expect(prepared.status).toBe(200);
    expect(prepared.body).toMatchObject({ name: "Avtal & offerter.zip", fileCount: 2, folderCount: 2 });

    const res = await call(prepared.body.url);
    expect(res.status).toBe(200);
    expect(res.headers.get("content-type")).toBe("application/zip");
    expect(res.headers.get("content-disposition")).toContain("Avtal%20%26%20offerter.zip");
    const zip = await bytes(res);
    expect(Number(res.headers.get("content-length"))).toBe(zip.length);

    const entries = unzip(zip);
    expect(entries.map((e) => e.name).sort()).toEqual([
      "Avtal & offerter/",
      "Avtal & offerter/Offert #12 (rev 2).pdf",
      "Avtal & offerter/Ramavtal 100%.docx",
      "Avtal & offerter/Tom mapp/",
    ]);
    const pdf = entries.find((e) => e.name.endsWith(".pdf"))!;
    expect(pdf.method).toBe(0);
    expect(pdf.data).toEqual(mock.contentOf(mock.mustPath("Clients/Acme/Avtal & offerter/Offert #12 (rev 2).pdf")));
  });

  it("zips a selection of files with unique names, via a plain GET link", async () => {
    mock.add("Clients/Acme/Bilder/README.txt");
    await syncAll();
    const res = await call(`/api/zip?path=${encodeURIComponent("/README.txt")}&path=${encodeURIComponent("/Bilder/README.txt")}&path=${encodeURIComponent("/Video/intro.mp4")}&name=urval`);
    expect(res.status).toBe(200);
    expect(res.headers.get("content-disposition")).toContain('filename="urval.zip"');
    const entries = unzip(await bytes(res));
    expect(entries.map((e) => e.name)).toEqual(["README.txt", "README (2).txt", "intro.mp4"]);
    expect(entries[2]!.data).toEqual(mock.contentOf(mock.mustPath("Clients/Acme/Video/intro.mp4")));
  });

  it("does not include the same file twice", async () => {
    const res = await call(`/api/zip?path=${encodeURIComponent("/Bilder")}&path=${encodeURIComponent("/Bilder/logo.png")}`);
    const names = unzip(await bytes(res)).map((e) => e.name);
    expect(names.filter((n) => n.endsWith("logo.png"))).toHaveLength(1);
  });

  it("uses current sizes and skips files deleted since the last sync", async () => {
    mock.setContent("Clients/Acme/README.txt", "this file grew after the sync ran");
    mock.remove("Clients/Acme/Bilder/kontor.jpg");
    const res = await call(`/api/zip?path=${encodeURIComponent("/README.txt")}&path=${encodeURIComponent("/Bilder")}`);
    expect(res.headers.get("x-zip-skipped-files")).toBe("1");
    const zip = await bytes(res);
    expect(Number(res.headers.get("content-length"))).toBe(zip.length);
    const entries = unzip(zip);
    expect(new TextDecoder().decode(entries.find((e) => e.name === "README.txt")!.data)).toBe("this file grew after the sync ran");
    expect(entries.map((e) => e.name)).not.toContain("Bilder/kontor.jpg");
  });

  it("enforces the file limit", async () => {
    const res = await json(
      "/api/zip",
      { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ paths: ["/Arkiv"] }) },
      { ZIP_MAX_FILES: "100" },
    );
    expect(res.status).toBe(413);
    expect(res.body.error.code).toBe("LIMIT_EXCEEDED");
  });

  it("rejects unknown or expired manifests", async () => {
    expect((await json(`/api/zip/${"0".repeat(32)}`)).status).toBe(404);
  });
});
