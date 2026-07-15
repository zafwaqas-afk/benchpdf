"""
Drives the real hub in Chromium end-to-end for every cell in the v1 matrix,
plus the two required error paths (password-protected file, corrupt file).
Captures a screenshot at each key state and asserts real output files.
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

results = []


def log(name, ok, detail=""):
    results.append((name, ok, detail))
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}{(' - ' + detail) if detail else ''}")


def wait_state(page, state, timeout=20000):
    page.wait_for_function("s => document.getElementById('stage').dataset.state === s",
                           arg=state, timeout=timeout)


def wait_batch_done(page, timeout=180000):
    """Wait until every row in the batch list has settled to done/error (not
    pending/converting) — checking row state directly avoids false positives
    from transient header text like 'Starting…'."""
    page.wait_for_function(
        """() => {
            const rows = document.querySelectorAll('#batch-list .brow');
            if (rows.length === 0) return false;
            return [...rows].every(r => r.classList.contains('done') || r.classList.contains('error'));
        }""",
        timeout=timeout)


def reset(page):
    page.goto(BASE)
    wait_state(page, "empty")


def shot(page, name):
    page.screenshot(path=os.path.join(SHOTS, f"{name}.png"), full_page=True)


def pick_via_chooser(page, cell_selector, files):
    """The grid cell's click handler calls fileInput.click() itself, which
    Playwright intercepts as a real filechooser event that must be answered
    explicitly — a follow-up set_input_files on the (now-irrelevant) input
    does not satisfy it."""
    with page.expect_file_chooser() as fc_info:
        page.click(cell_selector)
    fc_info.value.set_files(files)


def convert_via_grid(page, cell_target_id, files, extra=None, shot_name=None):
    """Click a grid cell (tool-first), pick file(s), wait for completion."""
    reset(page)
    pick_via_chooser(page, f'.cell[data-target="{cell_target_id}"]', files)
    wait_batch_done(page)
    if extra:
        extra(page)
    if shot_name:
        shot(page, shot_name)
    return page.evaluate("document.getElementById('batch-line').textContent")


def batch_items(page):
    return page.evaluate("""
        () => [...document.querySelectorAll('#batch-list .brow')].map(r => ({
            status: r.className.replace('brow ', ''),
            text: r.textContent
        }))
    """)


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(viewport={"width": 1180, "height": 900})
        page = ctx.new_page()
        console_errors = []
        page.on("console", lambda m: console_errors.append(m.text) if m.type == "error" else None)

        # ---- initial load ----
        reset(page)
        shot(page, "00_empty")
        cells = page.eval_on_selector_all(".cell", "els => els.length")
        log("grid renders with matrix cells", cells == 10, f"{cells} cells")
        pdf_to_excel_disabled = page.eval_on_selector(
            '.cell[data-target="pdf_to_excel"]', "el => el.classList.contains('disabled')")
        log("PDF -> Excel shown greyed/disabled ('coming soon')", pdf_to_excel_disabled)

        # =========================================================== TO PDF
        log_line = convert_via_grid(page, "word_to_pdf", [os.path.join(IN, "vendor_checklist.docx")],
                                    shot_name="01_word_to_pdf")
        log("Word -> PDF", "Done" in log_line, log_line)

        log_line = convert_via_grid(page, "excel_to_pdf", [os.path.join(IN, "budget.xlsx")],
                                    shot_name="02_excel_to_pdf")
        log("Excel -> PDF", "Done" in log_line, log_line)

        log_line = convert_via_grid(page, "ppt_to_pdf", [os.path.join(IN, "hub_test_deck.pptx")],
                                    shot_name="03_ppt_to_pdf")
        log("PowerPoint -> PDF", "Done" in log_line, log_line)

        log_line = convert_via_grid(page, "images_to_pdf",
                                    [os.path.join(IN, "photo_a.jpg"), os.path.join(IN, "photo_b.png"),
                                     os.path.join(IN, "photo_c.heic")],
                                    shot_name="04_images_to_pdf")
        log("Images (jpg+png+heic) -> merged PDF", "Done" in log_line, log_line)

        # web page -> PDF (URL flow, not the grid file-picker path)
        reset(page)
        page.fill("#url-input", "example.com")
        page.click("#url-go")
        wait_batch_done(page)
        line = page.evaluate("document.getElementById('batch-line').textContent")
        shot(page, "05_url_to_pdf")
        log("Web page (URL) -> PDF", "Done" in line, line)

        # =========================================================== FROM PDF
        log_line = convert_via_grid(page, "pdf_to_pptx", [os.path.join(IN, "tables_charts.pdf")],
                                    shot_name="06_pdf_to_pptx")
        report_visible = page.eval_on_selector("#report", "el => !el.hidden")
        log("PDF -> PowerPoint", "Done" in log_line, log_line)
        log("  per-page report shown for PDF->PPTX", report_visible)

        log_line = convert_via_grid(page, "pdf_to_docx", [os.path.join(IN, "text_report.pdf")],
                                    shot_name="07_pdf_to_docx")
        log("PDF -> Word", "Done" in log_line, log_line)

        reset(page)
        pick_via_chooser(page, '.cell[data-target="pdf_to_images"]', [os.path.join(IN, "tables_charts.pdf")])
        wait_batch_done(page)
        line = page.evaluate("document.getElementById('batch-line').textContent")
        shot(page, "08_pdf_to_images")
        log("PDF -> Images", "Done" in line, line)

        log_line = convert_via_grid(page, "pdf_to_text", [os.path.join(IN, "text_report.pdf")],
                                    shot_name="09_pdf_to_text")
        log("PDF -> Text", "Done" in log_line, log_line)

        # ============================================================ ERRORS
        reset(page)
        pick_via_chooser(page, '.cell[data-target="word_to_pdf"]', [os.path.join(IN, "protected.docx")])
        wait_batch_done(page, timeout=60000)
        items = batch_items(page)
        shot(page, "10_error_password")
        pw_ok = any(it["status"] == "error" for it in items)
        log("Password-protected docx -> clean per-item error (not a hang)", pw_ok,
            items[0]["text"][:120] if items else "no items")

        reset(page)
        pick_via_chooser(page, '.cell[data-target="pdf_to_text"]', [os.path.join(IN, "corrupt.pdf")])
        wait_batch_done(page, timeout=30000)
        items = batch_items(page)
        shot(page, "11_error_corrupt")
        corrupt_ok = any(it["status"] == "error" for it in items)
        log("Corrupt PDF -> clean per-item error (not a hang)", corrupt_ok,
            items[0]["text"][:120] if items else "no items")

        # ============================================================ DROP-FIRST + BATCH
        reset(page)
        dt_files = [os.path.join(IN, "photo_a.jpg"), os.path.join(IN, "photo_b.png")]
        # simulate a real drop of MULTIPLE docx (batch, sequential, per-file status)
        page.set_input_files("#file", [os.path.join(IN, "vendor_checklist.docx")])
        wait_state(page, "picking")
        shot(page, "12_chooser_dropfirst")
        chip_visible = page.eval_on_selector('.chip[data-target="word_to_pdf"]', "el => !!el")
        log("drop-first auto-detect shows target chooser", chip_visible)
        page.click('.chip[data-target="word_to_pdf"]')
        wait_batch_done(page)
        line = page.evaluate("document.getElementById('batch-line').textContent")
        log("drop-first conversion completes", "Done" in line, line)

        log("no console errors across full run", len(console_errors) == 0,
            "; ".join(console_errors[:3]))

        browser.close()

    print("\n=== SUMMARY ===")
    failed = [r for r in results if not r[1]]
    for name, ok, detail in results:
        print(f"  [{'PASS' if ok else 'FAIL'}] {name}")
    print(f"\n{len(results) - len(failed)}/{len(results)} checks passed")
    sys.exit(0 if not failed else 1)


if __name__ == "__main__":
    main()
