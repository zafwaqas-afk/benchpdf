# Third-party licenses bundled in BenchPDF

BenchPDF is distributed under the **GNU AGPL-3.0** because it bundles PyMuPDF (AGPL). The full BenchPDF source is available (see About → Source). Components shipped in the installed app:

| Component | Version | License | Notes |
|---|---|---|---|
| PyMuPDF (fitz) | 1.28 | AGPL-3.0 | GOVERNING copyleft: the whole app is AGPL-3.0; source must be available. Commercial license available from Artifex. |
| python-pptx | 1.0 | MIT | permissive |
| python-docx | 1.2 | MIT | permissive |
| Flask | 3.x | BSD-3-Clause | permissive (with Werkzeug, Jinja2, click, itsdangerous, markupsafe, blinker — all BSD/MIT) |
| pywin32 | 3xx | PSF-2.0 | permissive; drives installed Microsoft Office via COM |
| Pillow | 12.x | MIT-CMU (HPND) | permissive |
| pillow-heif | 1.x | BSD-3-Clause | wraps libheif/libde265 (LGPL-3.0). HEIC decode uses HEVC — HEVC is patent-encumbered; distributing HEIC support may need a codec/patent license in some jurisdictions. |
| img2pdf | 0.6 | LGPL-3.0 | WEAK COPYLEFT: allow relinking or provide the library's source; AGPL-compatible. |
| psutil | 7.x | BSD-3-Clause | permissive |
| pystray | 0.19 | LGPL-3.0 | WEAK COPYLEFT (system-tray icon); AGPL-compatible. |
| IBM Plex (Serif/Sans/Mono) | subset | SIL OFL 1.1 | OFL: bundle the license text; do not sell the fonts on their own; Reserved Font Name applies. Verified OFL. |
| Embedded CPython runtime | 3.14 | PSF-2.0 | permissive |
| PyInstaller bootloader | 6.x | GPL-2.0 with bootloader exception | the exception permits shipping the frozen app under any license (here AGPL-3.0). |

## Build/test-only (NOT shipped in the installer)

| Component | License | Purpose |
|---|---|---|
| reportlab | BSD | generates the bundled synthetic test PDFs |
| pikepdf | MPL-2.0 | strict PDF validation in tests |
| playwright | Apache-2.0 | browser-driven UI tests |
| fonttools | MIT | font subsetting at build time |
| PyInstaller | GPL-2.0 + exception | the packager itself |

## Copyleft / attention items
- **PyMuPDF (AGPL-3.0)** — governs the whole app; source must be offered.
- **img2pdf, pystray, libheif/libde265 (LGPL-3.0)** — weak copyleft; allow relinking or provide their source. All AGPL-compatible.
- **IBM Plex (OFL-1.1)** — ship the OFL text; Reserved Font Name; don't sell the fonts standalone.
- **HEIC/HEVC** — decoding HEIC relies on HEVC, which is patent-encumbered; shipping HEIC support may require a patent/codec license in some regions.
