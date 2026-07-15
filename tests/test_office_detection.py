"""Verify Office-dependent conversions grey out honestly when Office is absent,
and work when present — both cases exercised by mocking the availability check.
Also confirms the real registry probe runs without launching Office."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app import registry
from app import office_com

OFFICE_TARGETS = {"word_to_pdf", "excel_to_pdf", "ppt_to_pdf", "pdf_to_docx"}
results = []
def check(c, label):
    results.append(bool(c)); print(f"  [{'PASS' if c else 'FAIL'}] {label}")


def main():
    # 1) real probe never raises and returns a full map (doesn't launch Office)
    real = office_com.office_availability(force=True)
    check(set(real) == {"word", "excel", "powerpoint"} and all(isinstance(v, bool) for v in real.values()),
          f"real registry probe returns a clean map: {real}")

    # 2) Office PRESENT -> all Office targets enabled
    present = {"word": True, "excel": True, "powerpoint": True}
    m = {t["id"]: t for t in registry.matrix_json(present)}
    check(all(m[i]["enabled"] for i in OFFICE_TARGETS), "Office present: all Office targets enabled")

    # 3) Office ABSENT -> Office targets greyed with the honest note; others fine
    absent = {"word": False, "excel": False, "powerpoint": False}
    m = {t["id"]: t for t in registry.matrix_json(absent)}
    greyed = all(not m[i]["enabled"] for i in OFFICE_TARGETS)
    noted = all("Microsoft Office" in m[i]["note"] for i in OFFICE_TARGETS)
    check(greyed, "Office absent: all Office targets disabled")
    check(noted, f"Office absent: honest note shown ({m['word_to_pdf']['note']!r})")
    check(m["pdf_to_pptx"]["enabled"] and m["images_to_pdf"]["enabled"] and m["pdf_to_text"]["enabled"],
          "Office absent: non-Office conversions stay enabled")

    # 4) convert guard: is_available False for Office target when absent, True when present
    tgt = registry.TARGETS["word_to_pdf"]
    check(not registry.is_available(tgt, absent), "guard blocks Office target when absent (no crash path)")
    check(registry.is_available(tgt, present), "guard allows Office target when present")
    check(registry.is_available(registry.TARGETS["pdf_to_text"], absent),
          "guard allows non-Office target regardless of Office")

    print("\n" + ("OFFICE DETECTION: PASS" if all(results) else "OFFICE DETECTION: FAIL"))
    sys.exit(0 if all(results) else 1)


if __name__ == "__main__":
    main()
