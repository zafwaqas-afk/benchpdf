"use strict";

/* Verso in-app PDF editor. Renders each page as a text-free background with
   absolutely-positioned, contenteditable blocks laid over it (the same clustered
   blocks the converter produces). Pages render lazily as they scroll into view.
   Edits (text, move, add, delete) are collected per page and sent to
   /api/edit/<id>/export, which rebuilds only the touched pages. */

(function () {
  const $ = (id) => document.getElementById(id);
  const editor = $("editor"), edPages = $("ed-pages"), edScroll = $("ed-scroll");
  const edName = $("ed-name");
  const overlay = $("ed-overlay"), scrim = $("ed-scrim");

  let editId = null, pages = [], addMode = false, uid = 0;
  const nextId = () => "e" + (++uid);
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  // ---- open ----
  async function open(file) {
    if (!file) return;
    const fd = new FormData(); fd.append("file", file);
    editor.hidden = false; document.body.style.overflow = "hidden";
    edName.textContent = "Opening…"; edPages.innerHTML = ""; pages = [];
    let data;
    try {
      const res = await fetch("/api/edit/open", { method: "POST", body: fd });
      data = await res.json();
      if (!res.ok || data.error) throw new Error(data.error || "Couldn't open this PDF.");
    } catch (err) { return fail(err.message); }
    editId = data.edit_id;
    edName.textContent = data.name || "document.pdf";

    const avail = edScroll.clientWidth - 48;
    data.pages.forEach((p) => {
      const S = Math.min(1.7, Math.max(0.5, avail / p.width));
      const el = document.createElement("div");
      el.className = "ed-page is-loading";
      el.style.width = (p.width * S) + "px";
      el.style.height = (p.height * S) + "px";
      el.dataset.index = p.index;
      const num = document.createElement("div");
      num.className = "ed-pagenum"; num.textContent = "Page " + (p.index + 1);
      el.appendChild(num);
      edPages.appendChild(el);
      pages.push({ index: p.index, wpt: p.width, hpt: p.height, S, el, loaded: false, blocks: [], dirty: false });
      el.addEventListener("click", (ev) => onPageClick(ev, p.index));
    });
    edScroll.scrollTop = 0;
    observe();
  }

  // ---- lazy render ----
  let io = null;
  function observe() {
    if (io) io.disconnect();
    io = new IntersectionObserver((entries) => {
      entries.forEach((e) => { if (e.isIntersecting) renderPage(parseInt(e.target.dataset.index, 10)); });
    }, { root: edScroll, rootMargin: "500px 0px" });
    pages.forEach((p) => io.observe(p.el));
  }

  async function renderPage(index) {
    const st = pages[index];
    if (!st || st.loaded) return;
    st.loaded = true;
    const bg = document.createElement("img");
    bg.className = "ed-bg"; bg.alt = ""; bg.src = `/api/edit/${editId}/bg/${index}`;
    bg.addEventListener("load", () => st.el.classList.remove("is-loading"));
    st.el.appendChild(bg);
    let model;
    try {
      const res = await fetch(`/api/edit/${editId}/page/${index}`);
      model = await res.json();
      if (model.error) throw new Error(model.error);
    } catch (err) { st.loaded = false; st.el.classList.remove("is-loading"); return; }
    (model.blocks || []).forEach((b) => {
      if (b.type === "table") (b.cells || []).forEach((c) => addBlock(st, c, "cell", b.id));
      else addBlock(st, b, "text", null);
    });
  }

  // ---- blocks ----
  function addBlock(st, model, kind, tableId) {
    const S = st.S, [x0, y0, x1, y1] = model.bbox;
    const el = document.createElement("div");
    el.className = "ed-block" + (kind === "cell" ? " ed-cell" : "");
    el.contentEditable = "true"; el.spellcheck = false;
    el.innerHTML = model.html || "<div><br></div>";
    const oneLine = kind !== "cell" && (y1 - y0) <= (model.size || 10) * 1.7;
    if (oneLine) el.classList.add("oneline");
    Object.assign(el.style, {
      left: (x0 * S) + "px", top: (y0 * S) + "px",
      width: kind === "cell" ? ((x1 - x0) * S) + "px" : (oneLine ? "" : ((x1 - x0) * S) + "px"),
      maxWidth: ((st.wpt - x0) * S - 2) + "px",
      fontFamily: `"${model.font || "Arial"}"`, fontSize: ((model.size || 10) * S) + "px",
      fontWeight: model.bold ? "700" : "400", fontStyle: model.italic ? "italic" : "normal",
      color: model.color || "#000", textAlign: model.align || "left",
    });
    const rec = { id: nextId(), kind, tableId, model, el, origHtml: model.html || "",
      origBbox: model.bbox.slice(), bbox: model.bbox.slice(), deleted: false, isNew: false };
    st.blocks.push(rec);
    el.addEventListener("input", () => { if (el.innerHTML !== rec.origHtml) { el.classList.add("is-edited"); st.dirty = true; } });
    if (kind !== "cell") handles(st, rec);
    st.el.appendChild(el);
    return rec;
  }

  function handles(st, rec) {
    const grip = document.createElement("div");
    grip.className = "ed-grip"; grip.title = "Drag to move";
    grip.innerHTML = '<svg viewBox="0 0 16 16" width="12" height="12"><circle cx="4" cy="3" r="1.3"/><circle cx="4" cy="8" r="1.3"/><circle cx="4" cy="13" r="1.3"/><circle cx="10" cy="3" r="1.3"/><circle cx="10" cy="8" r="1.3"/><circle cx="10" cy="13" r="1.3"/></svg>';
    const del = document.createElement("button");
    del.className = "ed-del"; del.type = "button"; del.title = "Delete this block"; del.textContent = "×";
    del.addEventListener("mousedown", (e) => e.preventDefault());
    del.addEventListener("click", (e) => { e.stopPropagation(); rec.deleted = true; st.dirty = true; rec.el.remove(); });
    rec.el.appendChild(grip); rec.el.appendChild(del);
    drag(st, rec, grip);
  }

  function drag(st, rec, grip) {
    let sx, sy, ox, oy, on = false;
    grip.addEventListener("mousedown", (e) => {
      e.preventDefault(); e.stopPropagation(); on = true;
      sx = e.clientX; sy = e.clientY; ox = parseFloat(rec.el.style.left); oy = parseFloat(rec.el.style.top);
      rec.el.classList.add("is-dragging");
      document.addEventListener("mousemove", move); document.addEventListener("mouseup", up);
    });
    function move(e) {
      if (!on) return;
      rec.el.style.left = Math.max(0, ox + (e.clientX - sx)) + "px";
      rec.el.style.top = Math.max(0, oy + (e.clientY - sy)) + "px";
    }
    function up() {
      on = false; rec.el.classList.remove("is-dragging");
      document.removeEventListener("mousemove", move); document.removeEventListener("mouseup", up);
      const S = st.S, nx = parseFloat(rec.el.style.left) / S, ny = parseFloat(rec.el.style.top) / S;
      const w = rec.origBbox[2] - rec.origBbox[0], h = rec.origBbox[3] - rec.origBbox[1];
      rec.bbox = [nx, ny, nx + w, ny + h]; rec.el.classList.add("is-edited"); st.dirty = true;
    }
  }

  // ---- add text box ----
  function onPageClick(ev, index) {
    if (!addMode || ev.target.closest(".ed-block")) return;
    const st = pages[index], r = st.el.getBoundingClientRect(), S = st.S;
    const x0 = (ev.clientX - r.left) / S, y0 = (ev.clientY - r.top) / S;
    const model = { bbox: [x0, y0, x0 + 200, y0 + 16], font: "IBM Plex Sans", size: 11,
      bold: false, italic: false, color: getComputedStyle(document.documentElement).getPropertyValue("--ink").trim() || "#1A2238",
      align: "left", html: "New text" };
    // export needs a real installed font; use a doc-mapped default
    model.font = "Arial";
    const rec = addBlock(st, model, "text", null);
    rec.isNew = true; rec.origHtml = " "; st.dirty = true; rec.el.classList.add("is-edited");
    setAdd(false); rec.el.focus();
    try { document.getSelection().selectAllChildren(rec.el); } catch (e) {}
  }
  function setAdd(on) {
    addMode = on; editor.classList.toggle("is-adding", on); $("ed-add").classList.toggle("is-active", on);
    $("ed-hint").textContent = on ? "Click on a page to place a text box" : "Click any text to edit";
  }

  // ---- serialize to a clean HTML subset ----
  function serialize(node) {
    let out = "";
    node.childNodes.forEach((n) => {
      if (n.nodeType === 3) { out += esc(n.nodeValue); return; }
      if (n.nodeType !== 1) return;
      if (n.classList && (n.classList.contains("ed-grip") || n.classList.contains("ed-del"))) return;
      const tag = n.tagName, inner = serialize(n);
      if (tag === "BR") out += "<br>";
      else if (tag === "B" || tag === "STRONG") out += `<b>${inner}</b>`;
      else if (tag === "I" || tag === "EM") out += `<i>${inner}</i>`;
      else if (tag === "DIV" || tag === "P") out += `<div>${inner}</div>`;
      else {
        let piece = inner;
        const w = n.style && (n.style.fontWeight === "bold" || +n.style.fontWeight >= 600);
        const it = n.style && n.style.fontStyle === "italic";
        const col = n.style && n.style.color;
        if (w) piece = `<b>${piece}</b>`;
        if (it) piece = `<i>${piece}</i>`;
        if (col) piece = `<span style="color:${esc(col)}">${piece}</span>`;
        out += piece;
      }
    });
    return out;
  }

  function payload() {
    const p = { pages: {} };
    pages.forEach((st) => {
      if (!st.dirty) return;
      const tables = {}, blocks = [];
      st.blocks.forEach((rec) => {
        if (rec.deleted) return;
        const html = serialize(rec.el).trim() || "<div><br></div>";
        const m = rec.model;
        const base = { bbox: rec.bbox, font: m.font, size: m.size, bold: !!m.bold, italic: !!m.italic, color: m.color, align: m.align, html };
        if (rec.kind === "cell") (tables[rec.tableId] = tables[rec.tableId] || { type: "table", cells: [] }).cells.push(Object.assign({ fill: m.fill }, base));
        else blocks.push(Object.assign({ type: "text" }, base));
      });
      Object.values(tables).forEach((t) => blocks.push(t));
      p.pages[String(st.index)] = { blocks };
    });
    return p;
  }
  const anyEdits = () => pages.some((p) => p.dirty);

  // ---- export ----
  async function exportPdf() {
    ov("Exporting your PDF…", "Applying your edits and keeping every other page untouched.");
    let data;
    try {
      const res = await fetch(`/api/edit/${editId}/export`, { method: "POST",
        headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload()) });
      data = await res.json();
      if (!res.ok || data.error) throw new Error(data.error || "Export failed.");
    } catch (err) { return ovError(err.message); }
    const n = (data.edited_pages || []).length;
    ovDone("Your PDF is ready", n ? `${n} page${n === 1 ? "" : "s"} rebuilt with your edits; the rest are untouched.`
      : "No changes were needed; this is a copy of the original.", `/api/edit/${editId}/download`);
  }
  async function exportPptx() {
    ov("Building your PowerPoint…", "Converting the edited document to editable slides.");
    try {
      const res = await fetch(`/api/edit/${editId}/pptx`, { method: "POST",
        headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload()) });
      const data = await res.json();
      if (!res.ok || data.error) throw new Error(data.error || "Export failed.");
    } catch (err) { return ovError(err.message); }
    ovDone("Your PowerPoint is ready", "Editable slides built from your document.", `/api/edit/${editId}/download_pptx`);
  }

  // ---- overlay ----
  function ov(title, msg) {
    scrim.hidden = false; overlay.hidden = false;
    $("ed-spinner").hidden = false; $("ed-ov-check").hidden = true;
    $("ed-ov-title").textContent = title; $("ed-ov-msg").textContent = msg || "";
    $("ed-ov-actions").hidden = true; $("ed-dl").style.display = "";
  }
  function ovDone(title, msg, dl) {
    $("ed-spinner").hidden = true; $("ed-ov-check").hidden = false;
    $("ed-ov-title").textContent = title; $("ed-ov-msg").textContent = msg || "";
    $("ed-dl").href = dl;
    $("ed-open").onclick = () => fetch(`/api/edit/${editId}/open`, { method: "POST" });
    $("ed-ov-actions").hidden = false;
  }
  function ovError(msg) {
    $("ed-spinner").hidden = true; $("ed-ov-check").hidden = true;
    $("ed-ov-title").textContent = "That didn't work"; $("ed-ov-msg").textContent = msg;
    $("ed-dl").style.display = "none"; $("ed-ov-actions").hidden = false;
  }
  function shutOverlay() { overlay.hidden = true; scrim.hidden = true; $("ed-dl").style.display = ""; }
  function fail(msg) { ov("Couldn't open this PDF", ""); ovError(msg); }

  // ---- wiring ----
  $("ed-back").addEventListener("click", () => {
    if (anyEdits() && !confirm("Leave the editor? Unsaved edits will be lost.")) return;
    editor.hidden = true; document.body.style.overflow = ""; if (io) io.disconnect();
  });
  $("ed-add").addEventListener("click", () => setAdd(!addMode));
  $("ed-export").addEventListener("click", exportPdf);
  $("ed-pptx").addEventListener("click", exportPptx);
  $("ed-ov-close").addEventListener("click", shutOverlay);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") { if (addMode) setAdd(false); else if (!overlay.hidden) shutOverlay(); }
  });

  window.Editor = { open };
})();
