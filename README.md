# Convert — a local file-conversion hub

A local, private tool for converting between PDF, Word, Excel, PowerPoint,
images, and web pages. **No file ever leaves your machine.** Office documents
are converted by driving your own installed copy of Word, Excel, and
PowerPoint — never a web API. There is no account, no cloud, and the tool
makes no network calls of its own (the one exception is "web page → PDF",
which — like opening the page in a browser — fetches the address you give
it, and nothing else).

This replaces Smallpdf-style conversion tools for local, private use.

---

## How to use it

1. **Double-click the `PDF to PowerPoint` icon on your Desktop.**
   (The very first time, it spends a minute setting itself up. After that it
   starts in a couple of seconds.)
2. A black window opens and your web browser pops up automatically, showing
   the **Convert** hub.
3. **Drop a file** onto the drop zone — the tool detects what kind of file it
   is and offers the matching conversions. Or **click a cell in the grid**
   first (e.g. "PDF → Word") to pick the conversion, then choose your file.
4. Drop **several files at once** to convert them all, one after another,
   with live per-file status. Drop several **images** together and they
   merge into a single PDF, in order.
5. When it's done, **download** each result (or **Download all** for a
   batch), or click **Show files** to open the output folder.
6. Converting a PDF has one extra bonus: the **PDF → PowerPoint** cell keeps
   the detailed per-page report (which pages are natively editable vs. a
   locked background image, and any font substitutions).
7. When you're done, close the black window to stop the tool.

> If the Desktop icon is ever missing, double-click **`PDF to PowerPoint.bat`**
> inside the `PDF2PPTX` folder, or re-run **`create_desktop_shortcut.ps1`**.

---

## What it can convert (v1)

| To PDF | From PDF |
|---|---|
| Word (.docx/.doc) → PDF | **PDF → Edit & export** (edit text in the browser) |
| Excel (.xlsx) → PDF | PDF → PowerPoint (editable slides) |
| PowerPoint (.pptx/.ppt) → PDF | PDF → Word (.docx) |
| Images (.jpg/.png/.heic, multi-select) → one merged PDF | PDF → Images (PNG/JPG, choice of resolution) |
| Web page (paste a URL) → PDF | PDF → Text (.txt) |
| | PDF → Excel — **coming soon** |

### Editing a PDF in the browser

Choose **PDF → Edit & export** (or drop a PDF and pick it) to open the editor.
Each page is shown as your document with its text laid over it as editable
blocks:

- **Click any text** to edit it in place; it reflows within its box.
- **Drag** a block by the grip handle on its left to move it; hover a block and
  click **×** to delete it; use **+ Text box** to add a new one.
- Tables are editable cell-by-cell, over their original fills.
- **Export PDF** rebuilds only the pages you changed and leaves every other page
  **byte-identical** to the original. Edited text stays real, searchable,
  selectable text (never flattened to an image) in the closest matching
  installed font. **Download PPTX** is also available from the same screen.

Large documents stay responsive — pages render only as you scroll to them.

Word, Excel, and PowerPoint conversions are done by automating your real,
installed Office apps in the background (invisibly) — the same engine your
own copy of Office uses, just driven for you. If Office ever hits a problem
partway through (a password-protected file, a corrupted document, or Office
simply not responding), the tool reports a plain-language error instead of
freezing, and stays ready for the next file.

### PDF → PowerPoint, in detail

The tool decides automatically, page by page, and tells you which mode it
used:

| Mode | When it's used | What you get |
|------|----------------|--------------|
| **Native** | Mostly text on a plain background | Real, editable PowerPoint **text boxes** and **tables**, plus embedded photos placed as pictures. |
| **Hybrid** | Pages with charts, vector drawings, or complex backgrounds | The artwork is rendered as a **locked background image** and the **editable text is laid on top**. |
| **Image-only** | Scanned pages with no selectable text | The page is placed as an image and you get a clear warning. (OCR is **not** part of this version.) |

---

## Requirements

- **Windows 11** with **Microsoft Word, Excel, and PowerPoint** installed
  (desktop versions) for the Office-format conversions. PDF-only conversions
  (images, text extraction, PDF → PowerPoint) work even without Office.
- **Microsoft Edge or Google Chrome** installed, for "web page → PDF".
- **Python 3** installed and on your PATH. If it isn't, get it from
  <https://www.python.org/downloads/> and tick **"Add Python to PATH"**
  during install. (The launcher handles everything else itself.)

---

## Privacy

- The tool runs a small web page on **your PC only** (`http://127.0.0.1:5000`).
  The address `127.0.0.1` means *this computer* — it is not reachable from
  the internet, and no other device or user can see your files or activity.
- Office documents are converted by scripting your own local copy of Word /
  Excel / PowerPoint — the files are never uploaded anywhere.
- All page styling and scripts are stored **locally** in the `app/static`
  folder; nothing is fetched from a CDN.

---

## A note on Office automation reliability

Occasionally, Word/Excel/PowerPoint can show a dialog during automation (a
password prompt, a "this file previously caused a problem" recovery notice,
etc.) that isn't visible on screen but can block a single conversion. The
tool has a built-in timeout: if Office doesn't respond within 30–90 seconds,
it force-closes that one stuck attempt, reports a clear error for that file,
and is immediately ready for your next conversion — it never leaves the app
frozen. If a specific file keeps failing this way, opening it once directly
in Office (and dismissing whatever it asks) usually clears the underlying
Office-level flag for good.

---

## Licensing note (important if you ever share this tool)

This tool uses **PyMuPDF**, which is licensed under the **GNU AGPL v3**. That is
perfectly fine for your own **personal use**. However, if you ever
**distribute** this tool to others or run it as a service for other people, the
AGPL has obligations (broadly: you must make the source code available under the
AGPL). If you plan to do that, review the PyMuPDF licence first, or contact
Artifex about a commercial licence.

Other components used: `python-pptx` (MIT), `python-docx` (MIT), `Flask` (BSD),
`pywin32` (PSF), `img2pdf` (LGPL), `pillow-heif` (BSD), `psutil` (BSD),
`reportlab` (BSD — only used to generate the bundled test files). The interface
is set in **IBM Plex** (Serif, Sans, Mono), subset and vendored locally under
the SIL Open Font License — no fonts are loaded from a network at runtime.

---

## Conversion fidelity — the regression guard

Conversion quality (text placement, tables, fonts) is protected by an
automated **fidelity suite**. **No engine change ships without a green suite.**

- **Run it:** double-click **`run-tests.bat`**, or
  `venv\Scripts\python tests\fidelity_suite.py`.
- **What it asserts** on the committed fixtures in `tests/fixtures/`
  (a text-heavy doc, a table-heavy doc, and a slide-style doc), for both the
  **PDF → PPTX** and **editor export** paths:
  paragraph-level blocks (no line fragments), no fabricated `" / "` joins,
  every detected table rebuilt as a native table, **zero text-box insets**,
  consistent font mapping, **no two elements overlapping by >10%**, nothing past
  the page edge, and **text-box positions within 2% of a committed golden
  layout**.
- **The guard:** a git `pre-commit` hook (in `.githooks/`) runs the suite
  automatically whenever engine code (`extraction.py`, `converter.py`,
  `edit_model.py`, `pdf_edit_export.py`, `fonts_local.py`) or the fixtures
  change, and **blocks the commit if it goes red**. If a layout change is
  *intentional*, re-bless the golden with
  `venv\Scripts\python tests\fidelity_suite.py --update-golden` and review the
  diff before committing.

**Architecture note.** Extraction and placement are deliberately separated so a
change made for one output can't silently shift another's:
`app/extraction.py` is the single shared layer (clustering, font mapping, fill
sampling); each target keeps its own placement policy — `converter.py` (PPTX),
`edit_model.py` + `pdf_edit_export.py` (editor). Don't put placement logic in
`extraction.py`.

> This is the public source copy. In private development, the suite is also
> run against an additional confidential real-world document; that fixture
> (and anything that names it) is deliberately excluded here.

---

## Packaged desktop app (BenchPDF)

The app ships as **BenchPDF**, a self-contained Windows desktop application built
with PyInstaller (embedded Python runtime, all dependencies, vendored fonts, and
an app icon — no Python install required on the target PC).

- **Installer:** `dist_installer\BenchPDF-Setup-1.0.0.exe` — a **per-user**
  installer that needs **no administrator rights** (installs to
  `%LOCALAPPDATA%\Programs\BenchPDF`). It adds Start Menu shortcuts and a clean
  uninstaller.
- **Runtime:** no console window; on launch it starts the local server, opens
  your default browser to it, and shows a **system-tray icon** (Open / About /
  Quit) so there's a clear way to close it. First run has no setup and no
  account — it opens straight to the drop zone.
- **Office detection:** Word/Excel/PowerPoint conversions are greyed out with an
  honest note on machines without Microsoft Office; the app never tries to use
  Office that isn't there.
- **About** (header → About): version, the privacy statement, the AGPL license,
  a source-availability link, and the bundled open-source components.
- **Hidden diagnostics:** `BenchPDF.exe --diagnostics` runs the fidelity
  regression suite against the bundled synthetic fixtures and reports pass/fail
  (also on the "BenchPDF diagnostics" Start Menu shortcut). Verified green in the
  packaged build.
- **Rebuild:** `venv\Scripts\python -m PyInstaller packaging\benchpdf.spec`
  then compile `packaging\benchpdf.iss` with Inno Setup's `ISCC.exe`.

### Licensing (important)

BenchPDF bundles **PyMuPDF, which is AGPL-3.0**, so **the whole app is
distributed under the GNU AGPL-3.0** (`LICENSE`) and its source must be made
available — set `SOURCE_URL` in `app/version.py` to your real source location
(it ships as a placeholder). Every bundled component and its license is
enumerated in `THIRD_PARTY_LICENSES.md`; copyleft/attention items are flagged
there (img2pdf/pystray/libheif are LGPL; IBM Plex is OFL-1.1, verified; HEIC
decoding relies on patent-encumbered HEVC). The installer and exe are **not
code-signed**, so Windows SmartScreen may warn on first run until you sign them.

---

## Folder layout

```
PDF2PPTX/
  PDF to PowerPoint.bat        <- double-click this (or the Desktop shortcut)
  create_desktop_shortcut.ps1  <- re-creates the Desktop icon
  requirements.txt
  app/
    server.py                  <- local web server + batch/job model
    registry.py                <- the conversion matrix (what converts to what)
    office_com.py               <- the single Office COM session manager
    converter.py                <- PDF -> PPTX engine
    convert_images.py           <- images -> PDF
    convert_web.py               <- web page -> PDF (headless browser)
    convert_pdf_misc.py          <- PDF -> images / PDF -> text
    templates/index.html
    static/style.css, app.js    <- local UI assets (no internet)
  tests/
    make_test_pdfs.py, make_extra_pdfs.py, make_office_test_files.py
                                 <- build sample PDFs, Office files, images
    verify_engine.py, verify_hub.py, verify_outputs.py, test_com_crash.py
                                 <- automated conversion + browser + crash-path checks
    input/  output/             <- sample files and their converted outputs
  venv/                         <- the private Python environment (auto-created)
  _work/                        <- temporary files from each conversion
```

## Re-running the built-in tests (optional)

```
venv\Scripts\python tests\make_test_pdfs.py
venv\Scripts\python tests\make_extra_pdfs.py
venv\Scripts\python tests\make_office_test_files.py
venv\Scripts\python tests\verify_engine.py
venv\Scripts\python tests\verify_outputs.py
venv\Scripts\python tests\test_com_crash.py
```

The hub's browser-driven suite (`tests\verify_hub.py`) needs the server
running first (`venv\Scripts\python app\server.py`) and Playwright's Chromium
installed (`venv\Scripts\python -m playwright install chromium`).
