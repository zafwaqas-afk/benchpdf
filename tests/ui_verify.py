"""
Drives the real app in Chromium (Playwright), captures all six UI states in
light and dark, verifies the end-to-end journey, and records every network
request to prove nothing loads from outside localhost.

Run with the server already listening on http://127.0.0.1:5000
"""
import os
import sys
import time
import zipfile

from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
INP = os.path.join(HERE, "input")
SHOTS = os.path.join(HERE, "shots")
BASE = "http://127.0.0.1:5000/"

os.makedirs(SHOTS, exist_ok=True)
# a non-PDF file to exercise the error state
NOTPDF = os.path.join(SHOTS, "notes.txt")
open(NOTPDF, "w").write("i am not a pdf")

external_requests = []
all_requests = []


def is_local(url):
    return (url.startswith("http://127.0.0.1:5000")
            or url.startswith("http://localhost:5000")
            or url.startswith("data:") or url.startswith("blob:"))


def wait_state(page, state, timeout=60):
    page.wait_for_function(
        "s => document.getElementById('stage').dataset.state === s",
        arg=state, timeout=timeout * 1000)


def shot(page, theme, name):
    page.evaluate("document.fonts.ready")
    time.sleep(0.25)
    path = os.path.join(SHOTS, theme, f"{name}.png")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    page.screenshot(path=path, full_page=True)
    print(f"  shot: {theme}/{name}.png")


def run_theme(browser, theme):
    ctx = browser.new_context(color_scheme=theme, viewport={"width": 1180, "height": 900},
                              device_scale_factor=2)
    page = ctx.new_page()
    page.on("request", lambda r: (all_requests.append(r.url),
            None if is_local(r.url) else external_requests.append(r.url)))

    # 1 · EMPTY
    page.goto(BASE)
    wait_state(page, "empty")
    shot(page, theme, "1_empty")

    # 2 · DRAG-OVER (apply the real class the drag handler uses)
    page.eval_on_selector("#dropzone", "el => el.classList.add('over')")
    page.eval_on_selector("#dz-head", "el => el.textContent = 'Release to convert'")
    shot(page, theme, "2_dragover")
    page.eval_on_selector("#dropzone", "el => el.classList.remove('over')")

    # 3 · PROCESSING — capture a real mid-conversion frame from the 52-page file
    page.goto(BASE); wait_state(page, "empty")
    page.set_input_files("#file", os.path.join(INP, "big50.pdf"))
    captured = False
    for _ in range(400):
        st = page.evaluate("document.getElementById('stage').dataset.state")
        if st == "processing":
            pct = page.evaluate("parseInt(document.getElementById('proc-pct').textContent)||0")
            if 12 <= pct <= 88:
                shot(page, theme, "3_processing"); captured = True; break
        elif st in ("success", "error"):
            break
        time.sleep(0.05)
    if not captured:
        # fallback: representative processing frame (real conversion was too fast)
        page.goto(BASE); wait_state(page, "empty")
        page.evaluate("""() => {
            const st=document.getElementById('stage'); st.dataset.state='processing';
            document.getElementById('proc-page').textContent='18';
            document.getElementById('proc-total').textContent='52';
            document.getElementById('proc-phase').textContent='Reading page 18';
            document.getElementById('proc-file').textContent='big50.pdf';
            document.getElementById('proc-pct').textContent='35';
            document.getElementById('fill').style.width='35%';
        }""")
        shot(page, theme, "3_processing")

    # 4 · SUCCESS (real conversion of the text report)
    page.goto(BASE); wait_state(page, "empty")
    page.set_input_files("#file", os.path.join(INP, "slide_deck.pdf"))
    wait_state(page, "success")
    shot(page, theme, "4_success")

    # 5 · ERROR (real path: a non-PDF file)
    page.goto(BASE); wait_state(page, "empty")
    page.set_input_files("#file", NOTPDF)
    wait_state(page, "error")
    shot(page, theme, "5_error")

    # 6 · SCANNED warning (real conversion of the image-only PDF)
    page.goto(BASE); wait_state(page, "empty")
    page.set_input_files("#file", os.path.join(INP, "scanned.pdf"))
    wait_state(page, "success")
    page.wait_for_selector("#scan-notice:not([hidden])", timeout=30000)
    shot(page, theme, "6_scanned")

    ctx.close()


def verify_journey(browser):
    """Full end-to-end: upload, progress, success, download a valid .pptx."""
    ctx = browser.new_context(viewport={"width": 1180, "height": 900})
    page = ctx.new_page()
    page.on("request", lambda r: (all_requests.append(r.url),
            None if is_local(r.url) else external_requests.append(r.url)))
    page.goto(BASE); wait_state(page, "empty")
    page.set_input_files("#file", os.path.join(INP, "tables_charts.pdf"))
    wait_state(page, "success")
    # per-page report present?
    rows = page.eval_on_selector_all("#report-body tr", "els => els.length")
    resline = page.inner_text("#res-line")
    # download and validate the file
    with page.expect_download() as dl_info:
        page.click("#download")
    dl = dl_info.value
    out = os.path.join(SHOTS, "downloaded.pptx")
    dl.save_as(out)
    ok_zip = zipfile.is_zipfile(out)
    has_pres = False
    if ok_zip:
        with zipfile.ZipFile(out) as z:
            has_pres = "ppt/presentation.xml" in z.namelist()
    ctx.close()
    print(f"\nJOURNEY: result='{resline}', report_rows={rows}, "
          f"valid_pptx={ok_zip and has_pres}, bytes={os.path.getsize(out)}")
    return (rows > 0 and ok_zip and has_pres)


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        print("== capturing states (light) ==");  run_theme(browser, "light")
        print("== capturing states (dark) ==");   run_theme(browser, "dark")
        print("== full journey verification ==")
        journey_ok = verify_journey(browser)
        browser.close()

    print("\n== network audit ==")
    print(f"total requests observed: {len(all_requests)}")
    if external_requests:
        print("EXTERNAL REQUESTS FOUND:")
        for u in sorted(set(external_requests)):
            print("  ", u)
    else:
        print("OK: every request stayed on localhost (no CDN / external calls)")
    # show the unique local asset URLs for transparency
    fonts = sorted({u for u in all_requests if "/fonts/" in u})
    print("font requests (local):", [u.split("/static/")[-1] for u in fonts])

    ok = journey_ok and not external_requests
    print("\n" + ("UI VERIFY PASSED" if ok else "UI VERIFY FAILED"))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
