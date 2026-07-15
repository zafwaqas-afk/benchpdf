"""
Focused, single-cell verification for PDF -> Word: drives the REAL browser
against the running hub server (grid cell -> file chooser -> real Word COM
conversion -> download), then opens the result with python-docx to confirm
it's a valid document with real text.

Run standalone (server must already be running on :5000) so it doesn't
re-touch Word/Excel/PowerPoint/Chromium for the other matrix cells.
"""
import os
import sys
import time

from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
IN = os.path.join(HERE, "input")
SHOTS = os.path.join(HERE, "shots_hub")
os.makedirs(SHOTS, exist_ok=True)
BASE = "http://127.0.0.1:5000/"


def wait_state(page, state, timeout=15000):
    page.wait_for_function("s => document.getElementById('stage').dataset.state === s",
                           arg=state, timeout=timeout)


def wait_batch_done(page, timeout=120000):
    page.wait_for_function(
        """() => {
            const rows = document.querySelectorAll('#batch-list .brow');
            if (rows.length === 0) return false;
            return [...rows].every(r => r.classList.contains('done') || r.classList.contains('error'));
        }""",
        timeout=timeout)


def main():
    src_pdf = os.path.join(IN, "text_report.pdf")
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1180, "height": 900})
        console_errors = []
        page.on("console", lambda m: console_errors.append(m.text) if m.type == "error" else None)

        page.goto(BASE)
        wait_state(page, "empty")

        t0 = time.time()
        with page.expect_file_chooser() as fc_info:
            page.click('.cell[data-target="pdf_to_docx"]')
        fc_info.value.set_files([src_pdf])
        wait_batch_done(page)
        elapsed = time.time() - t0

        line = page.evaluate("document.getElementById('batch-line').textContent")
        row_status = page.eval_on_selector("#batch-list .brow", "el => el.className")
        page.screenshot(path=os.path.join(SHOTS, "07_pdf_to_docx_retry.png"), full_page=True)

        print(f"batch line: {line!r}")
        print(f"row status: {row_status!r}")
        print(f"elapsed: {elapsed:.1f}s")
        print(f"console errors: {console_errors}")

        ok_ui = "done" in row_status and "Done" in line
        downloaded_path = None
        if ok_ui:
            with page.expect_download() as dl_info:
                page.click("#batch-list .brow-dl")
            dl = dl_info.value
            downloaded_path = os.path.join(SHOTS, "pdf_to_docx_downloaded.docx")
            dl.save_as(downloaded_path)

        browser.close()

    if not ok_ui:
        print("\nFAILED: batch did not reach a successful 'done' state in the browser.")
        sys.exit(1)

    from docx import Document
    doc = Document(downloaded_path)
    paras = [x.text for x in doc.paragraphs if x.text.strip()]
    print(f"\ndownloaded .docx opened with python-docx: {len(paras)} non-empty paragraph(s)")
    print("sample:", paras[:2])

    ok = len(paras) > 0
    print("\n" + ("PDF -> WORD LIVE VERIFICATION: PASSED" if ok else "PDF -> WORD LIVE VERIFICATION: FAILED"))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
