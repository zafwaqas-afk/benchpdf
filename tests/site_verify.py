"""Site journey verification: the converter's visible states are honest.

Runs the home-page drop zone through a real conversion in headless Chromium
and asserts the states a visitor actually sees. Exists because the done-state
warning banner rendered as an empty shell on every successful conversion for
weeks: display:flex on .notice defeated the hidden attribute, and no check
ever asserted the ABSENCE of an element.

Usage:
  python tests/site_verify.py             # serves the local benchpdf-site dir
  python tests/site_verify.py --live      # runs against https://benchpdf.pages.dev
"""

import http.server
import os
import socketserver
import sys
import threading

HERE = os.path.dirname(os.path.abspath(__file__))
SITE = os.environ.get("BENCHPDF_SITE_DIR") or os.path.join(
    os.path.dirname(os.path.dirname(HERE)), "benchpdf-site")
FIXTURE = os.path.join(HERE, "fixtures", "tables_charts.pdf")


class H(http.server.SimpleHTTPRequestHandler):
    def translate_path(self, path):
        p = path.split("?", 1)[0].lstrip("/")
        if p == "":
            p = "index.html"
        elif "." not in os.path.basename(p):
            p += ".html"
        return os.path.join(SITE, p.replace("/", os.sep))

    def log_message(self, *a):
        pass


def check(ok, label, results):
    results.append((bool(ok), label))
    print(f"  [{'PASS' if ok else 'FAIL'}] {label}")
    return ok


def main():
    from playwright.sync_api import sync_playwright

    live = "--live" in sys.argv
    if live:
        base = "https://benchpdf.pages.dev"
        srv = None
    else:
        srv = socketserver.ThreadingTCPServer(("127.0.0.1", 0), H)
        srv.daemon_threads = True
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        base = f"http://127.0.0.1:{srv.server_address[1]}"

    results = []
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(viewport={"width": 1440, "height": 950})
        console_errors = []
        pg.on("console", lambda m: console_errors.append(m.text) if m.type == "error" else None)

        print(f"== converter journey against {base} ==")
        pg.goto(base + "/", wait_until="load")
        pg.set_input_files("#file", FIXTURE)
        pg.wait_for_selector("#action-list button", timeout=25000)
        pg.click("#action-list button:has-text('PowerPoint')")
        pg.wait_for_selector("#conv-done", state="visible", timeout=300000)

        check(pg.locator("#conv-done h2").is_visible(), "done state shows its heading", results)
        check(pg.locator("#result-download").is_visible(), "download button is visible", results)

        # The regression this file exists for: a successful conversion must
        # show no warning shell. Assert the notice both by visibility and by
        # the general rule that nothing visible in the done panel is an empty
        # styled box.
        check(not pg.locator("#done-notice").is_visible(),
              "no warning banner after a successful conversion", results)
        empties = pg.evaluate("""() => {
          const out = [];
          for (const el of document.querySelectorAll('#conv-done [class]')) {
            const cs = getComputedStyle(el);
            if (cs.display === 'none' || cs.visibility === 'hidden') continue;
            const styled = cs.borderStyle !== 'none' || cs.backgroundColor !== 'rgba(0, 0, 0, 0)';
            // an icon does not make a box non-empty: only text or an
            // interactive element does
            if (styled && el.innerText.trim() === ''
                && !el.querySelector('a,button,input')) {
              out.push(el.className);
            }
          }
          return out;
        }""")
        check(not empties, f"no visible empty styled boxes in the done state {empties or ''}", results)

        # a real warning must still render when one is set
        pg.evaluate("""() => {
          document.getElementById('done-notice-text').textContent = 'probe';
          document.getElementById('done-notice').hidden = false;
        }""")
        check(pg.locator("#done-notice").is_visible()
              and "probe" in pg.inner_text("#done-notice"),
              "a genuine warning still renders when set", results)

        check(not console_errors, f"no console errors {console_errors[:2] or ''}", results)
        b.close()

    if srv:
        srv.shutdown()
    ok = all(r for r, _ in results)
    print("SITE VERIFY:", "PASS" if ok else "FAIL")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
