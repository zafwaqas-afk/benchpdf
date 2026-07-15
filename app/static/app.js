"use strict";

/* Verso converter surface. Talks to /api/matrix, /api/convert, /api/batch/<id>.
   Home states (on #home[data-state]): idle -> actions -> working -> done | error.
   Backend contract is unchanged from the hub build. */

const $ = (id) => document.getElementById(id);
const home = $("home");
const dropzone = $("dropzone");
const fileInput = $("file");
const gridTo = $("grid-to"), gridFrom = $("grid-from");
const actList = $("action-list"), actExtra = $("action-extra"), actSkipped = $("action-skipped");
const actType = $("act-type"), actName = $("act-name"), actSize = $("act-size");

const EXT_TO_TYPE = {
  ".pdf": "pdf", ".docx": "word", ".doc": "word", ".xlsx": "excel", ".xls": "excel",
  ".pptx": "powerpoint", ".ppt": "powerpoint", ".jpg": "image", ".jpeg": "image",
  ".png": "image", ".heic": "image", ".heif": "image",
};
const TYPE_LABEL = { pdf: "PDF", word: "Word", excel: "Excel", powerpoint: "PowerPoint", image: "Image" };

let MATRIX = [];
let currentBatch = null, pollTimer = null;
let scopedTarget = null;         // set when a grid cell / nav starts a tool-first flow
let pending = { files: [], type: null };
let inputSizes = {};             // filename -> bytes

const setState = (s) => { home.dataset.state = s; };
const extOf = (n) => { const i = n.lastIndexOf("."); return i >= 0 ? n.slice(i).toLowerCase() : ""; };
const esc = (s) => String(s).replace(/[&<>"']/g, (c) =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

function fmtSize(bytes) {
  if (!bytes && bytes !== 0) return "";
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(0) + " KB";
  return (bytes / 1048576).toFixed(1) + " MB";
}

/* ---------------- theme ---------------- */
(function themeInit() {
  const toggle = $("theme-toggle");
  toggle.addEventListener("click", () => {
    const root = document.documentElement;
    const isDark = matchMedia("(prefers-color-scheme: dark)").matches;
    const cur = root.getAttribute("data-theme") || (isDark ? "dark" : "light");
    const next = cur === "dark" ? "light" : "dark";
    root.setAttribute("data-theme", next);
    try { localStorage.setItem("benchpdf-theme", next); } catch (e) {}
  });
})();

/* ---------------- about dialog ---------------- */
(function aboutInit() {
  const nav = $("nav-about"), dlg = $("about-dialog"), scrim = $("about-scrim"), close = $("about-close");
  let lastFocus = null, loaded = false;
  async function load() {
    if (loaded) return;
    try {
      const a = await (await fetch("/api/about")).json();
      $("about-version").textContent = a.version;
      $("about-privacy").textContent = a.privacy;
      $("about-license").textContent = a.license;
      const src = $("about-source"); src.href = a.source_url; src.textContent = a.source_url;
      const off = a.office || {};
      const missing = ["word", "excel", "powerpoint"].filter((k) => !off[k]);
      $("about-office").textContent = missing.length
        ? "Microsoft Office was not detected, so Office conversions are unavailable on this PC."
        : "Microsoft Office is installed, so all conversions are available.";
      $("about-deps-list").innerHTML = (a.licenses || []).map((l) => `<li>${esc(l)}</li>`).join("");
      loaded = true;
    } catch (e) {
      $("about-privacy").textContent = "Couldn't load details.";
    }
  }
  function open() { lastFocus = document.activeElement; load(); dlg.hidden = false; scrim.hidden = false; close.focus(); }
  function shut() { dlg.hidden = true; scrim.hidden = true; if (lastFocus) lastFocus.focus(); }
  nav.addEventListener("click", open);
  close.addEventListener("click", shut);
  scrim.addEventListener("click", shut);
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && !dlg.hidden) shut(); });
})();

/* ---------------- boot: matrix + grid ---------------- */
fetch("/api/matrix").then((r) => r.json()).then((d) => {
  MATRIX = d.targets || [];
  renderGrid();
}).catch(() => { gridTo.innerHTML = '<p class="conv-item-desc">Couldn\'t load the conversion list.</p>'; });

function targetLabelVerb(t) {
  if (t.action === "edit") return "Edit in your browser";
  if (t.group === "to_pdf") return t.source_type === "image" ? "Combine into a PDF" : "Convert to PDF";
  return "Convert to " + t.label.split("→").pop().trim();
}

function renderGrid() {
  const toP = MATRIX.filter((t) => t.group === "to_pdf");
  const frP = MATRIX.filter((t) => t.group === "from_pdf");
  const item = (t) => `<button type="button" class="conv-item" data-target="${t.id}" ${t.enabled ? "" : "disabled"}>
      <span class="conv-item-top"><span class="conv-item-label">${esc(t.label)}</span>
      ${t.enabled ? "" : `<span class="conv-badge">${esc(t.note || "soon")}</span>`}</span>
      <span class="conv-item-desc">${esc(t.description)}</span></button>`;
  gridTo.innerHTML = toP.map(item).join("");
  gridFrom.innerHTML = frP.map(item).join("");
  [...gridTo.children, ...gridFrom.children].forEach((el) => {
    const t = MATRIX.find((m) => m.id === el.dataset.target);
    if (t && t.enabled) el.addEventListener("click", () => gridClick(t));
  });
}

function gridClick(t) {
  if (t.source_type === "url") { $("url-input").focus(); return; }
  scopedTarget = t;
  fileInput.multiple = !!t.multi;
  fileInput.accept = t.accepts.join(",");
  fileInput.click();
}

/* ---------------- file intake ---------------- */
dropzone.addEventListener("click", () => { scopedTarget = null; openPicker(true); });
dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); scopedTarget = null; openPicker(true); }
});
$("browse").addEventListener("click", (e) => { e.stopPropagation(); scopedTarget = null; openPicker(true); });
function openPicker(multi) { fileInput.multiple = multi; fileInput.accept = ""; fileInput.click(); }

fileInput.addEventListener("change", () => {
  const files = [...fileInput.files]; fileInput.value = "";
  if (!files.length) return;
  if (scopedTarget) {
    const t = scopedTarget; scopedTarget = null;
    if (t.action === "edit") return window.Editor.open(files[0]);
    recordSizes(files);
    startConversion(t, files, readExtraParams(t));
  } else intake(files);
});

["dragenter", "dragover"].forEach((ev) => document.addEventListener(ev, (e) => {
  if (home.dataset.state !== "idle") return;
  e.preventDefault(); dropzone.classList.add("is-over");
}));
["dragleave", "drop"].forEach((ev) => document.addEventListener(ev, (e) => {
  e.preventDefault();
  if (ev === "dragleave" && e.relatedTarget) return;
  dropzone.classList.remove("is-over");
}));
document.addEventListener("drop", (e) => {
  if (home.dataset.state !== "idle") return;
  const files = [...(e.dataTransfer ? e.dataTransfer.files : [])];
  if (files.length) intake(files);
});

function recordSizes(files) {
  inputSizes = {}; files.forEach((f) => { inputSizes[f.name] = f.size; });
}

function intake(files) {
  const groups = {};
  files.forEach((f) => { const ty = EXT_TO_TYPE[extOf(f.name)]; if (ty) (groups[ty] = groups[ty] || []).push(f); });
  const types = Object.keys(groups);
  if (!types.length) return showError("That file type isn't supported",
    "Verso converts PDF, Word, Excel, PowerPoint, and image files. Choose one of those, or paste a web address to convert a page to PDF.");
  types.sort((a, b) => groups[b].length - groups[a].length);
  const type = types[0], kept = groups[type];
  const skipped = types.slice(1).reduce((n, t) => n + groups[t].length, 0);
  recordSizes(kept);
  pending = { files: kept, type };
  showActions(type, kept, skipped);
}

/* ---------------- actions reveal ---------------- */
function showActions(type, files, skipped) {
  const n = files.length;
  actType.textContent = TYPE_LABEL[type] || type;
  actName.textContent = n === 1 ? files[0].name : `${n} ${TYPE_LABEL[type]} files`;
  actSize.textContent = n === 1 ? fmtSize(files[0].size)
    : fmtSize(files.reduce((s, f) => s + f.size, 0)) + " total";

  const targets = MATRIX.filter((t) => t.source_type === type);
  actList.innerHTML = targets.map((t) => {
    const label = (t.multi && n > 1) ? "Combine into one PDF" : targetLabelVerb(t);
    const primary = (t.action === "edit" || t.id === "pdf_to_pptx" || t.group === "to_pdf");
    return `<button type="button" class="action${primary ? " primary" : ""}${t.enabled ? "" : ""}"
        data-target="${t.id}" ${t.enabled ? "" : "disabled"}>
        <span class="action-label">${esc(label)}</span>
        <span class="action-desc">${esc(t.enabled ? t.description : (t.note || "Coming soon"))}</span>
      </button>`;
  }).join("");

  // extra params (PDF -> images dpi/format)
  actExtra.hidden = true; actExtra.innerHTML = "";
  const wp = targets.find((t) => t.params && t.params.length);
  if (wp) {
    let h = "";
    if (wp.params.includes("format")) h += `<label>Format <select id="opt-format"><option value="png">PNG</option><option value="jpg">JPG</option></select></label>`;
    if (wp.params.includes("dpi")) h += `<label>Resolution <select id="opt-dpi"><option value="100">100 DPI</option><option value="150" selected>150 DPI</option><option value="200">200 DPI</option><option value="300">300 DPI</option></select></label>`;
    actExtra.innerHTML = h; actExtra.hidden = false;
  }

  actSkipped.hidden = !skipped;
  if (skipped) actSkipped.textContent = `${skipped} file${skipped === 1 ? "" : "s"} of another type ${skipped === 1 ? "was" : "were"} set aside. Convert ${skipped === 1 ? "it" : "them"} separately.`;

  [...actList.children].forEach((el) => {
    const t = MATRIX.find((m) => m.id === el.dataset.target);
    if (!t || !t.enabled) return;
    el.addEventListener("click", () => {
      if (t.action === "edit") return window.Editor.open(files[0]);
      startConversion(t, files, readExtraParams(t));
    });
  });
  setState("actions");
}

function readExtraParams(t) {
  const p = {};
  const fmt = $("opt-format"), dpi = $("opt-dpi");
  if (fmt) p.format = fmt.value;
  if (dpi) p.dpi = dpi.value;
  return p;
}

$("act-cancel").addEventListener("click", reset);
$("done-more").addEventListener("click", reset);
$("err-retry").addEventListener("click", reset);
function reset() { clearTimeout(pollTimer); pending = { files: [], type: null }; setState("idle"); }

/* ---------------- URL flow ---------------- */
$("url-go").addEventListener("click", runUrl);
$("url-input").addEventListener("keydown", (e) => { if (e.key === "Enter") runUrl(); });
function runUrl() {
  const t = MATRIX.find((m) => m.id === "url_to_pdf");
  const v = $("url-input").value.trim();
  if (!v || !t) { $("url-input").focus(); return; }
  inputSizes = {};
  startConversion(t, [], { url: v });
}

/* ---------------- nav ---------------- */
$("nav-convert").addEventListener("click", () => {
  $("nav-convert").classList.add("is-active"); $("nav-edit").classList.remove("is-active");
  reset();
});
$("nav-edit").addEventListener("click", () => {
  const t = MATRIX.find((m) => m.id === "pdf_edit");
  scopedTarget = t; fileInput.multiple = false; fileInput.accept = ".pdf"; fileInput.click();
});

/* ---------------- conversion + polling ---------------- */
function startConversion(target, files, params) {
  setState("working");
  currentBatch = null;
  $("work-num").textContent = "0"; $("work-total").textContent = "0";
  $("work-phase").textContent = "Preparing…"; $("work-target").textContent = targetLabelVerb(target);
  $("work-fill").style.width = "3%"; $("work-list").innerHTML = "";

  const fd = new FormData();
  fd.append("target", target.id);
  files.forEach((f) => fd.append("files", f));
  if (params.url) fd.append("url", params.url);
  if (params.format) fd.append("format", params.format);
  if (params.dpi) fd.append("dpi", params.dpi);

  fetch("/api/convert", { method: "POST", body: fd })
    .then((r) => r.json().then((d) => ({ ok: r.ok, d })))
    .then(({ ok, d }) => {
      if (!ok || d.error) return showError("That conversion couldn't start", d.error || "Please try again.");
      currentBatch = d.job_id || d.batch_id;
      poll();
    })
    .catch(() => showError("Verso isn't responding", "The local converter window may have closed. Reopen the tool and try again."));
}

function poll() {
  fetch(`/api/batch/${currentBatch}`).then((r) => r.json()).then((b) => {
    if (b.error) return showError("Lost track of that conversion", b.error);
    renderWorking(b);
    if (b.status === "done") return finish(b);
    pollTimer = setTimeout(poll, 300);
  }).catch(() => showError("Lost contact with the converter", "Reopen the tool and try again."));
}

function renderWorking(b) {
  const items = b.items || [];
  const total = items.length;
  const settled = items.filter((i) => i.status === "done" || i.status === "error").length;
  const active = items.find((i) => i.status === "converting") || items[settled] || items[0];

  // progress bar: page-level if a single item exposes it, else file-level
  let pct;
  if (total === 1 && active && active.total) {
    pct = Math.min(100, Math.round((active.current / active.total) * 100));
    // show the page being read (matches the phase line), not the 0-based index
    const pm = (active.message || "").match(/page\s+(\d+)\s+of\s+(\d+)/i);
    $("work-num").textContent = pm ? pm[1] : Math.min((active.current || 0) + 1, active.total);
    $("work-total").textContent = active.total;
  } else {
    pct = total ? Math.round((settled / total) * 100) : 5;
    $("work-num").textContent = settled;
    $("work-total").textContent = total;
  }
  $("work-fill").style.width = Math.max(pct, 3) + "%";
  $("work-phase").textContent = active ? phaseText(active) : "Working…";

  $("work-list").innerHTML = items.map((it) => {
    const cls = it.status === "done" ? "is-done" : it.status === "error" ? "is-error"
      : it.status === "converting" ? "is-working" : "";
    const status = it.status === "error" ? (it.error || "Failed")
      : it.status === "done" ? "Done" : it.status === "converting" ? "Converting" : "Queued";
    return `<li class="filerow ${cls}"><span class="fr-dot"></span>
      <span class="fr-name">${esc(it.name)}</span><span class="fr-status">${esc(status)}</span></li>`;
  }).join("");
}

function phaseText(it) {
  const m = (it.message || "").match(/page\s+(\d+)\s+of\s+(\d+)/i);
  if (m) return `Reading page ${m[1]} of ${m[2]}`;
  if (/saving/i.test(it.message || "")) return "Saving the file";
  if (it.status === "converting") return "Converting " + it.name;
  return it.message || "Working…";
}

function finish(b) {
  const items = b.items || [];
  const done = items.filter((i) => i.status === "done");
  const errs = items.filter((i) => i.status === "error");
  $("work-fill").style.width = "100%";

  $("done-title").textContent = errs.length === 0
    ? (done.length === 1 ? "Your file is ready" : `${done.length} files ready`)
    : done.length === 0 ? "That didn't work" : `${done.length} of ${items.length} ready`;

  // scanned-PDF notice + report from a pptx item
  const rep = (done.find((i) => i.extra && i.extra.page_report) || {}).extra;
  renderScan(rep && rep.page_report);
  renderReport(rep && rep.page_report);

  $("result-list").innerHTML = items.map((it) => resultRow(b, it)).join("");
  [...$("result-list").querySelectorAll("[data-dl]")].forEach((a) => {
    a.href = `/api/batch/${currentBatch}/download/${a.dataset.dl}`;
  });
  [...$("result-list").querySelectorAll("[data-open]")].forEach((btn) => {
    btn.addEventListener("click", () => fetch(`/api/batch/${currentBatch}/open/${btn.dataset.open}`, { method: "POST" }));
  });
  [...$("result-list").querySelectorAll("[data-again]")].forEach((btn) => {
    btn.addEventListener("click", () => { if (pending.type) showActions(pending.type, pending.files, 0); });
  });

  if (errs.length && done.length === 0) {
    return showError("That file couldn't be converted", errs[0].error ||
      "The converter hit a problem. If it keeps happening, re-save the file from its original app and try again.");
  }
  setState("done");
}

function resultRow(b, it) {
  if (it.status === "error") {
    return `<li class="result"><div class="result-main"><div class="result-name">${esc(it.name)}</div>
      <div class="result-size" style="color:var(--error)">${esc(it.error || "Couldn't convert this file")}</div></div></li>`;
  }
  const before = inputSizes[it.name];
  const after = it.size;
  let sizeLine = fmtSize(after);
  if (before && after) {
    const pct = Math.round((1 - after / before) * 100);
    const delta = pct > 0 ? `<span class="delta">${pct}% smaller</span>` : "";
    sizeLine = `${fmtSize(before)}<span class="arw">→</span>${fmtSize(after)} ${delta}`;
  }
  const canAgain = pending.type && MATRIX.filter((t) => t.source_type === pending.type && t.enabled).length > 1;
  return `<li class="result">
    <div class="result-main">
      <div class="result-name">${esc(it.name)}</div>
      <div class="result-size">${sizeLine}</div>
    </div>
    <div class="result-actions">
      <a class="btn btn-primary btn-sm" data-dl="${it.id}" href="#" download>Download</a>
      <button class="btn btn-ghost btn-sm" data-open="${it.id}" type="button">Open</button>
      ${canAgain ? `<button class="btn btn-ghost btn-sm" data-again="1" type="button">Convert to something else</button>` : ""}
    </div>
  </li>`;
}

function renderScan(rep) {
  const el = $("scan-notice");
  if (rep && (rep.scanned_warning || (rep.warnings || []).some((w) => /scan|image-only|no selectable/i.test(w)))) {
    const w = (rep.warnings || []).find((x) => /scan|image-only|no selectable|no extractable/i.test(x));
    $("scan-text").textContent = w || "Some pages had no selectable text and were kept as images, so their text isn't editable. OCR isn't part of Verso.";
    el.hidden = false;
  } else el.hidden = true;
}

function renderReport(rep) {
  const box = $("report");
  if (!rep || !rep.pages) { box.hidden = true; return; }
  const modes = {};
  rep.pages.forEach((p) => { modes[p.mode] = (modes[p.mode] || 0) + 1; });
  $("report-sum").textContent = Object.entries(modes).map(([k, v]) => `${v} ${k}`).join(" · ");
  $("report-body").innerHTML = rep.pages.map((p) => {
    const el = [];
    if (p.tables) el.push(p.tables + " table" + (p.tables === 1 ? "" : "s"));
    if (p.text_boxes) el.push(p.text_boxes + " text");
    if (p.images) el.push(p.images + " image" + (p.images === 1 ? "" : "s"));
    return `<tr><td class="c-pg">${p.page}</td>
      <td><span class="mode ${esc(p.mode)}">${esc(p.mode)}</span></td>
      <td>${el.length ? el.join(", ") : "-"}</td></tr>`;
  }).join("");
  box.hidden = false;
}

function showError(title, body) {
  clearTimeout(pollTimer);
  $("err-title").textContent = title;
  $("err-body").textContent = body;
  setState("error");
}
