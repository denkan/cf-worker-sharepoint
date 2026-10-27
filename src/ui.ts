/**
 * Proof-of-concept browser for the API. Deliberately dependency-free: one HTML
 * page talking to /api/*. The real frontend is expected to replace it.
 */
export const uiHtml = String.raw`<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SharePoint folder</title>
<style>
  :root {
    --bg: #f6f7f9; --panel: #fff; --text: #16181d; --muted: #667085; --line: #e4e7ec;
    --accent: #0f6cbd; --accent-soft: #e8f1fb; --danger: #b42318; --radius: 10px;
    font-family: ui-sans-serif, system-ui, -apple-system, "Segoe UI", sans-serif;
  }
  @media (prefers-color-scheme: dark) {
    :root { --bg: #111317; --panel: #1a1d23; --text: #e8eaee; --muted: #98a2b3; --line: #2b2f37; --accent: #5aa9f0; --accent-soft: #1d2a3a; }
  }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--bg); color: var(--text); font-size: 14px; }
  header { display: flex; align-items: center; gap: 12px; padding: 12px 20px; border-bottom: 1px solid var(--line); background: var(--panel); flex-wrap: wrap; }
  header h1 { font-size: 15px; margin: 0; font-weight: 650; }
  .pill { font-size: 12px; padding: 3px 9px; border-radius: 999px; background: var(--accent-soft); color: var(--accent); }
  .pill.warn { background: #fef3c7; color: #92400e; }
  .spacer { flex: 1; }
  button, .btn { font: inherit; border: 1px solid var(--line); background: var(--panel); color: var(--text); padding: 6px 12px; border-radius: 8px; cursor: pointer; text-decoration: none; display: inline-flex; gap: 6px; align-items: center; }
  button:hover, .btn:hover { border-color: var(--accent); }
  button.primary { background: var(--accent); border-color: var(--accent); color: #fff; }
  button:disabled { opacity: .5; cursor: default; }
  input[type=text], input[type=search] { font: inherit; padding: 7px 10px; border: 1px solid var(--line); border-radius: 8px; background: var(--panel); color: var(--text); min-width: 0; }
  main { display: grid; grid-template-columns: minmax(0, 1fr) 380px; gap: 16px; padding: 16px 20px; }
  @media (max-width: 960px) { main { grid-template-columns: 1fr; } }
  .card { background: var(--panel); border: 1px solid var(--line); border-radius: var(--radius); }
  .toolbar { display: flex; gap: 8px; padding: 12px; border-bottom: 1px solid var(--line); flex-wrap: wrap; }
  .toolbar input { flex: 1 1 220px; }
  .crumbs { padding: 10px 14px; display: flex; flex-wrap: wrap; gap: 4px; align-items: center; color: var(--muted); border-bottom: 1px solid var(--line); }
  .crumbs a { color: var(--accent); text-decoration: none; cursor: pointer; }
  table { width: 100%; border-collapse: collapse; }
  th, td { text-align: left; padding: 8px 12px; border-bottom: 1px solid var(--line); white-space: nowrap; }
  th { font-weight: 600; color: var(--muted); font-size: 12px; cursor: pointer; user-select: none; }
  td.name { white-space: normal; word-break: break-word; width: 100%; }
  tr.row:hover { background: var(--accent-soft); }
  tr.row.active { background: var(--accent-soft); }
  .name-cell { display: flex; gap: 10px; align-items: center; cursor: pointer; }
  .thumb { width: 32px; height: 32px; border-radius: 6px; object-fit: cover; background: var(--bg); flex: none; display: grid; place-items: center; font-size: 11px; font-weight: 700; color: var(--muted); }
  .sub { color: var(--muted); font-size: 12px; }
  .pager { display: flex; justify-content: space-between; align-items: center; padding: 10px 12px; color: var(--muted); }
  .preview { position: sticky; top: 16px; align-self: start; padding: 14px; display: flex; flex-direction: column; gap: 12px; min-height: 200px; }
  .preview h2 { font-size: 15px; margin: 0; word-break: break-word; }
  .preview .media { border-radius: 8px; overflow: hidden; background: var(--bg); }
  .preview iframe { width: 100%; height: 480px; border: 0; display: block; }
  .preview img, .preview video { width: 100%; display: block; }
  .preview pre { margin: 0; padding: 10px; max-height: 360px; overflow: auto; font-size: 12px; white-space: pre-wrap; }
  dl { display: grid; grid-template-columns: auto 1fr; gap: 4px 12px; margin: 0; font-size: 12px; }
  dt { color: var(--muted); } dd { margin: 0; word-break: break-all; }
  .actions { display: flex; gap: 8px; flex-wrap: wrap; }
  .empty { padding: 40px; text-align: center; color: var(--muted); }
  .error { color: var(--danger); padding: 12px 14px; }
</style>
</head>
<body>
<header>
  <h1>SharePoint folder</h1>
  <span id="status" class="pill">…</span>
  <span class="spacer"></span>
  <button id="sync">Sync now</button>
  <a class="btn" href="/openapi.json" target="_blank">OpenAPI</a>
</header>
<main>
  <section class="card">
    <form class="toolbar" id="goForm">
      <input type="text" id="goInput" placeholder="Go to a path, full path or pasted SharePoint link…" aria-label="Path or link">
      <button type="submit">Go</button>
    </form>
    <form class="toolbar" id="searchForm">
      <input type="search" id="searchInput" placeholder="Search names in this folder…" aria-label="Search">
      <button type="submit">Search</button>
      <button type="button" id="zipBtn" class="primary" disabled>Download zip</button>
    </form>
    <nav class="crumbs" id="crumbs"></nav>
    <div id="error" class="error" hidden></div>
    <table>
      <thead><tr>
        <th style="width:28px"><input type="checkbox" id="selectAll" aria-label="Select all"></th>
        <th data-sort="name">Name</th>
        <th data-sort="modified">Modified</th>
        <th data-sort="size">Size</th>
      </tr></thead>
      <tbody id="rows"></tbody>
    </table>
    <div class="pager">
      <span id="range"></span>
      <span><button id="prev">Previous</button> <button id="next">Next</button></span>
    </div>
  </section>
  <aside class="card preview" id="preview">
    <div class="empty">Select a file to preview it.</div>
  </aside>
</main>
<script>
(function () {
  var PAGE = 50;
  var state = { path: "/", offset: 0, sort: "name", order: "asc", query: "", items: [], selected: new Set(), active: null };
  var $ = function (id) { return document.getElementById(id); };
  var enc = encodeURIComponent;

  function el(tag, attrs, children) {
    var e = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) {
      if (k === "text") e.textContent = attrs[k];
      else if (k.slice(0, 2) === "on") e.addEventListener(k.slice(2), attrs[k]);
      else e.setAttribute(k, attrs[k]);
    });
    (children || []).forEach(function (c) { if (c) e.appendChild(typeof c === "string" ? document.createTextNode(c) : c); });
    return e;
  }

  function api(url, init) {
    return fetch(url, init).then(function (res) {
      return res.json().then(function (body) {
        if (!res.ok) throw new Error((body.error && body.error.message) || res.statusText);
        return body;
      });
    });
  }

  function size(n) {
    if (n == null) return "";
    var u = ["B", "KB", "MB", "GB", "TB"], i = 0;
    while (n >= 1024 && i < u.length - 1) { n /= 1024; i++; }
    return (i ? n.toFixed(1) : n) + " " + u[i];
  }
  function date(s) { return s ? new Date(s).toLocaleString() : ""; }
  function showError(msg) { var e = $("error"); e.hidden = !msg; e.textContent = msg || ""; }

  function loadStatus() {
    api("/api/status").then(function (s) {
      var pill = $("status");
      pill.className = "pill" + (s.index.ready ? "" : " warn");
      pill.textContent = (s.mock ? "Mock data · " : "") + (s.index.ready
        ? s.index.itemsInRoot + " items · synced " + (s.index.lastSyncCompletedAt ? new Date(s.index.lastSyncCompletedAt).toLocaleTimeString() : "")
        : (s.index.syncRunning ? "Initial sync running…" : "Not synced yet"));
      pill.title = "Site: " + s.site + "\nLibrary: " + (s.library || "default") + "\nRoot: " + s.rootPath;
    });
  }

  function navigate(path) {
    location.hash = "#" + enc(path || "/");
  }

  function readHash() {
    var p = decodeURIComponent(location.hash.slice(1) || "") || "/";
    state.path = p; state.offset = 0; state.query = ""; $("searchInput").value = "";
    state.selected.clear(); load();
  }

  function load() {
    showError("");
    var url = state.query
      ? "/api/search?q=" + enc(state.query) + "&path=" + enc(state.path) + "&sort=" + (state.sort === "name" && state.order === "asc" ? "relevance" : state.sort)
      : "/api/items?path=" + enc(state.path) + "&sort=" + state.sort;
    url += "&order=" + state.order + "&limit=" + PAGE + "&offset=" + state.offset;
    api(url).then(function (res) {
      state.items = res.items;
      render(res);
    }).catch(function (err) {
      state.items = []; render({ items: [], total: 0 }); showError(err.message);
    });
  }

  function renderCrumbs() {
    var c = $("crumbs"); c.textContent = "";
    var parts = state.path.split("/").filter(Boolean);
    c.appendChild(el("a", { text: "Root", onclick: function () { navigate("/"); } }));
    parts.forEach(function (p, i) {
      c.appendChild(document.createTextNode(" / "));
      var target = "/" + parts.slice(0, i + 1).join("/");
      c.appendChild(i === parts.length - 1 && !state.query ? el("span", { text: p }) : el("a", { text: p, onclick: function () { navigate(target); } }));
    });
    if (state.query) c.appendChild(el("span", { text: "  ·  search results for “" + state.query + "”" }));
  }

  function icon(item) {
    if (item.type === "folder") return el("div", { class: "thumb", text: "▤" });
    if (/^(image|video)\//.test(item.mimeType || "") || item.extension === "pdf") {
      var img = el("img", { class: "thumb", loading: "lazy", alt: "", src: item.links.thumbnail + "&size=small" });
      img.onerror = function () { img.replaceWith(el("div", { class: "thumb", text: (item.extension || "").slice(0, 4).toUpperCase() })); };
      return img;
    }
    return el("div", { class: "thumb", text: (item.extension || "file").slice(0, 4).toUpperCase() });
  }

  function render(res) {
    renderCrumbs();
    var tbody = $("rows"); tbody.textContent = "";
    if (!res.items.length) tbody.appendChild(el("tr", {}, [el("td", { colspan: "4", class: "empty", text: state.query ? "No matches." : "This folder is empty." })]));
    res.items.forEach(function (item) {
      var cb = el("input", { type: "checkbox", "aria-label": "Select " + item.name });
      cb.checked = state.selected.has(item.path);
      cb.addEventListener("change", function () { cb.checked ? state.selected.add(item.path) : state.selected.delete(item.path); updateZip(); });
      var nameCell = el("div", { class: "name-cell" }, [icon(item), el("div", {}, [
        el("div", { text: item.name }),
        state.query ? el("div", { class: "sub", text: item.path }) : null,
      ])]);
      nameCell.addEventListener("click", function () { item.type === "folder" ? navigate(item.path) : preview(item); });
      var tr = el("tr", { class: "row" + (state.active === item.path ? " active" : "") }, [
        el("td", {}, [cb]),
        el("td", { class: "name" }, [nameCell]),
        el("td", { class: "sub", text: date(item.modified) }),
        el("td", { class: "sub", text: item.type === "folder" ? (item.childCount + " items") : size(item.size) }),
      ]);
      tbody.appendChild(tr);
    });
    var from = res.total ? state.offset + 1 : 0, to = state.offset + res.items.length;
    $("range").textContent = from + "–" + to + " of " + res.total;
    $("prev").disabled = state.offset === 0;
    $("next").disabled = to >= res.total;
    $("selectAll").checked = res.items.length > 0 && res.items.every(function (i) { return state.selected.has(i.path); });
    updateZip();
  }

  function updateZip() {
    var n = state.selected.size;
    $("zipBtn").disabled = n === 0;
    $("zipBtn").textContent = n ? "Download zip (" + n + ")" : "Download zip";
  }

  function preview(item) {
    state.active = item.path;
    Array.prototype.forEach.call(document.querySelectorAll("tr.row"), function (r) { r.classList.remove("active"); });
    var p = $("preview"); p.textContent = "";
    var url = item.links.content;
    var mime = item.mimeType || "";
    var media = null;
    if (mime === "application/pdf") media = el("iframe", { src: url, title: item.name });
    else if (/^image\//.test(mime)) media = el("img", { src: url, alt: item.name });
    else if (/^video\//.test(mime)) media = el("video", { src: url, controls: "", preload: "metadata" });
    else if (/^audio\//.test(mime)) media = el("audio", { src: url, controls: "" });
    else if (/^text\/plain/.test(mime)) {
      media = el("pre", { text: "Loading…" });
      fetch(url, { headers: { range: "bytes=0-65535" } }).then(function (r) { return r.text(); }).then(function (t) { media.textContent = t; });
    }
    p.appendChild(el("h2", { text: item.name }));
    if (media) p.appendChild(el("div", { class: "media" }, [media]));
    else p.appendChild(el("div", { class: "empty", text: "No inline preview for this file type." }));
    p.appendChild(el("div", { class: "actions" }, [
      el("a", { class: "btn", href: item.links.download, text: "Download" }),
      item.webUrl ? el("a", { class: "btn", href: item.webUrl, target: "_blank", rel: "noopener", text: "Open in SharePoint" }) : null,
      el("button", { text: "Copy path", onclick: function () { navigator.clipboard.writeText(item.path); } }),
    ]));
    p.appendChild(el("dl", {}, [
      el("dt", { text: "Path" }), el("dd", { text: item.path }),
      el("dt", { text: "Type" }), el("dd", { text: mime || "—" }),
      el("dt", { text: "Size" }), el("dd", { text: size(item.size) }),
      el("dt", { text: "Modified" }), el("dd", { text: date(item.modified) + (item.modifiedBy ? " by " + item.modifiedBy : "") }),
    ]));
  }

  $("goForm").addEventListener("submit", function (e) {
    e.preventDefault();
    var v = $("goInput").value.trim(); if (!v) return;
    api("/api/resolve?path=" + enc(v)).then(function (r) {
      $("goInput").value = "";
      if (r.item.type === "folder") navigate(r.item.path);
      else {
        var parent = r.item.path.split("/").slice(0, -1).join("/") || "/";
        if (parent !== state.path) navigate(parent);
        preview(r.item);
      }
    }).catch(function (err) { showError(err.message); });
  });

  $("searchForm").addEventListener("submit", function (e) {
    e.preventDefault();
    state.query = $("searchInput").value.trim(); state.offset = 0; load();
  });

  $("zipBtn").addEventListener("click", function () {
    var paths = Array.from(state.selected);
    $("zipBtn").disabled = true;
    api("/api/zip", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ paths: paths }) })
      .then(function (m) { location.href = m.url; })
      .catch(function (err) { showError(err.message); })
      .finally(updateZip);
  });

  $("selectAll").addEventListener("change", function (e) {
    state.items.forEach(function (i) { e.target.checked ? state.selected.add(i.path) : state.selected.delete(i.path); });
    render({ items: state.items, total: Number(($("range").textContent.split(" of ")[1]) || 0) });
  });

  Array.prototype.forEach.call(document.querySelectorAll("th[data-sort]"), function (th) {
    th.addEventListener("click", function () {
      var s = th.getAttribute("data-sort");
      state.order = state.sort === s && state.order === "asc" ? "desc" : "asc";
      state.sort = s; state.offset = 0; load();
    });
  });

  $("prev").addEventListener("click", function () { state.offset = Math.max(0, state.offset - PAGE); load(); });
  $("next").addEventListener("click", function () { state.offset += PAGE; load(); });
  $("sync").addEventListener("click", function () {
    $("sync").disabled = true;
    api("/api/sync", { method: "POST" }).then(function () { loadStatus(); load(); })
      .catch(function (err) { showError(err.message); })
      .finally(function () { $("sync").disabled = false; });
  });

  window.addEventListener("hashchange", readHash);
  loadStatus();
  setInterval(loadStatus, 30000);
  readHash();
})();
</script>
</body>
</html>`;
