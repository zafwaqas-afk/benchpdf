"""Assemble the REAL-WORLD fidelity corpus: public PDFs from diverse generators.

Unlike the synthetic corpus (build_corpus.py), every document here was produced
by a real-world PDF generator: LaTeX (arXiv), government form pipelines (IRS,
gov.uk), InDesign/print-stream reports (ECB, BoE, SEC), troff (RFC), Word
(NIST, W3C), EU Publications Office (EUR-Lex). This is the corpus that catches
the failure classes synthetic fixtures cannot imagine.

Rules:
  * PUBLIC documents only. Every URL is recorded in manifest.json.
  * Download failures are skipped and recorded; the corpus is whatever succeeds.
  * Documents are trimmed to the first TRIM_PAGES pages to keep runs fast; the
    manifest records the original page count and the PDF producer string.
  * docs/ is gitignored (no downloaded binaries in the repo); manifest.json is
    committed so anyone can rebuild the same corpus.
"""

import json
import os
import urllib.request

import fitz

HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(HERE, "docs")
MANIFEST = os.path.join(HERE, "manifest.json")
TRIM_PAGES = 3
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) BenchPDF-corpus/1.0 (test corpus builder; contact zafwaqas@gmail.com)"

# (name, url, generator-class guess)
SOURCES = [
    # ---- arXiv: LaTeX ----
    ("arxiv_1706_03762", "https://arxiv.org/pdf/1706.03762", "latex"),
    ("arxiv_1810_04805", "https://arxiv.org/pdf/1810.04805", "latex"),
    ("arxiv_2005_14165", "https://arxiv.org/pdf/2005.14165", "latex"),
    ("arxiv_1512_03385", "https://arxiv.org/pdf/1512.03385", "latex"),
    ("arxiv_1409_0473", "https://arxiv.org/pdf/1409.0473", "latex"),
    ("arxiv_2010_11929", "https://arxiv.org/pdf/2010.11929", "latex"),
    ("arxiv_1412_6980", "https://arxiv.org/pdf/1412.6980", "latex"),
    ("arxiv_1502_03167", "https://arxiv.org/pdf/1502.03167", "latex"),
    ("arxiv_1606_06565", "https://arxiv.org/pdf/1606.06565", "latex"),
    ("arxiv_1707_06347", "https://arxiv.org/pdf/1707.06347", "latex"),
    ("arxiv_2203_02155", "https://arxiv.org/pdf/2203.02155", "latex"),
    ("arxiv_2302_13971", "https://arxiv.org/pdf/2302.13971", "latex"),
    ("arxiv_1301_3781", "https://arxiv.org/pdf/1301.3781", "latex"),
    ("arxiv_1611_05431", "https://arxiv.org/pdf/1611.05431", "latex"),
    # ---- IRS: government form pipeline ----
    ("irs_f1040", "https://www.irs.gov/pub/irs-pdf/f1040.pdf", "gov-form"),
    ("irs_fw9", "https://www.irs.gov/pub/irs-pdf/fw9.pdf", "gov-form"),
    ("irs_fw4", "https://www.irs.gov/pub/irs-pdf/fw4.pdf", "gov-form"),
    ("irs_f4506t", "https://www.irs.gov/pub/irs-pdf/f4506t.pdf", "gov-form"),
    ("irs_f941", "https://www.irs.gov/pub/irs-pdf/f941.pdf", "gov-form"),
    ("irs_f8949", "https://www.irs.gov/pub/irs-pdf/f8949.pdf", "gov-form"),
    ("irs_fss4", "https://www.irs.gov/pub/irs-pdf/fss4.pdf", "gov-form"),
    ("irs_f1065", "https://www.irs.gov/pub/irs-pdf/f1065.pdf", "gov-form"),
    ("irs_f1120", "https://www.irs.gov/pub/irs-pdf/f1120.pdf", "gov-form"),
    ("irs_p15", "https://www.irs.gov/pub/irs-pdf/p15.pdf", "gov-publication"),
    # ---- RFC: troff/xml2rfc print stream ----
    ("rfc_793", "https://www.rfc-editor.org/rfc/pdfrfc/rfc793.txt.pdf", "troff"),
    ("rfc_2616", "https://www.rfc-editor.org/rfc/pdfrfc/rfc2616.txt.pdf", "troff"),
    ("rfc_5321", "https://www.rfc-editor.org/rfc/pdfrfc/rfc5321.txt.pdf", "troff"),
    ("rfc_7231", "https://www.rfc-editor.org/rfc/pdfrfc/rfc7231.txt.pdf", "troff"),
    ("rfc_8446", "https://www.rfc-editor.org/rfc/pdfrfc/rfc8446.txt.pdf", "troff"),
    # ---- gov.uk: HMG publishing (Word/InDesign exports) ----
    ("govuk_r43_notes",
     "https://assets.publishing.service.gov.uk/media/69bc066f8006048065f73d03/R43_Notes.pdf",
     "govuk"),
    ("govuk_frem_2024_25",
     "https://assets.publishing.service.gov.uk/media/67ed5550632d0f88e8248c16/MASTER_FINAL_DRAFT_2024-25_FReM_APRIL_2025_RELEASE.pdf",
     "govuk"),
    ("govuk_spend_guidance",
     "https://assets.publishing.service.gov.uk/media/5a821d93ed915d74e6235dd7/guidance_for_publishing_spend.pdf",
     "govuk"),
    # ---- ECB: publication pipeline ----
    ("ecb_eb202404", "https://www.ecb.europa.eu/pub/pdf/ecbu/eb202404.en.pdf", "ecb-report"),
    ("ecb_eb202501", "https://www.ecb.europa.eu/pub/pdf/ecbu/eb202501.en.pdf", "ecb-report"),
    ("ecb_eb202503", "https://www.ecb.europa.eu/pub/pdf/ecbu/eb202503.en.pdf", "ecb-report"),
    ("ecb_eb202508", "https://www.ecb.europa.eu/pub/pdf/ecbu/eb202508.en.pdf", "ecb-report"),
    # ---- Bank of England: InDesign reports ----
    ("boe_mpr_2024_11",
     "https://www.bankofengland.co.uk/-/media/boe/files/monetary-policy-report/2024/november/monetary-policy-report-november-2024.pdf",
     "boe-report"),
    ("boe_mpr_2025_08",
     "https://www.bankofengland.co.uk/-/media/boe/files/monetary-policy-report/2025/august/monetary-policy-report-august-2025.pdf",
     "boe-report"),
    ("boe_mpr_2026_02",
     "https://www.bankofengland.co.uk/-/media/boe/files/monetary-policy-report/2026/february/monetary-policy-report-february-2026.pdf",
     "boe-report"),
    # ---- W3C: spec PDFs ----
    ("w3c_svg11_2011", "https://www.w3.org/TR/SVG11/REC-SVG11-20110816.pdf", "w3c-spec"),
    ("w3c_svg10_2001", "https://www.w3.org/TR/2001/REC-SVG-20010904/REC-SVG-20010904.pdf", "w3c-spec"),
    ("w3c_wcag20_wd", "https://www.w3.org/WAI/GL/WCAG20/WD-WCAG20-20010824.pdf", "w3c-spec"),
    ("w3c_svg_tutorial", "https://www.w3.org/2002/Talks/www2002-svgtut-ih/hwtut.pdf", "w3c-slides"),
    # ---- SEC EDGAR: annual report print streams ----
    ("sec_pnrg_ars_2026",
     "https://www.sec.gov/Archives/edgar/data/56868/000143774926013009/pnrg20260422_ars.pdf",
     "sec-filing"),
    ("sec_addus_ar_2023",
     "https://www.sec.gov/Archives/edgar/data/1468328/000095017024053594/2023_annual_report.pdf",
     "sec-filing"),
    ("sec_tm262321_ars",
     "https://www.sec.gov/Archives/edgar/data/908255/000110465926031786/tm262321d2_ars.pdf",
     "sec-filing"),
    # ---- NIST: Word-derived standards ----
    ("nist_sp800_53r5",
     "https://nvlpubs.nist.gov/nistpubs/SpecialPublications/NIST.SP.800-53r5.pdf", "nist"),
    ("nist_sp800_63b",
     "https://nvlpubs.nist.gov/nistpubs/SpecialPublications/NIST.SP.800-63b.pdf", "nist"),
    ("nist_fips_203", "https://nvlpubs.nist.gov/nistpubs/FIPS/NIST.FIPS.203.pdf", "nist"),
    # ---- Federal Reserve ----
    ("fed_fomc_2024_06",
     "https://www.federalreserve.gov/monetarypolicy/files/fomcminutes20240612.pdf", "fed"),
    ("fed_fomc_2025_01",
     "https://www.federalreserve.gov/monetarypolicy/files/fomcminutes20250129.pdf", "fed"),
    # ---- EUR-Lex: EU Publications Office ----
    ("eurlex_gdpr",
     "https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32016R0679", "eurlex"),
    ("eurlex_dsa",
     "https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32022R2065", "eurlex"),
    ("eurlex_dsm",
     "https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32019L0790", "eurlex"),
]


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/pdf,*/*"})
    return urllib.request.urlopen(req, timeout=90).read()


def main():
    os.makedirs(DOCS, exist_ok=True)
    manifest = {"trim_pages": TRIM_PAGES, "documents": []}
    ok = 0
    for name, url, gen in SOURCES:
        path = os.path.join(DOCS, name + ".pdf")
        entry = {"name": name, "url": url, "generator_class": gen}
        try:
            if not os.path.exists(path):
                data = fetch(url)
                if not data.startswith(b"%PDF"):
                    raise ValueError("not a PDF (%d bytes)" % len(data))
                src = fitz.open(stream=data, filetype="pdf")
                out = fitz.open()
                out.insert_pdf(src, from_page=0, to_page=min(TRIM_PAGES - 1, src.page_count - 1))
                entry["source_pages"] = src.page_count
                entry["producer"] = src.metadata.get("producer", "")
                entry["creator"] = src.metadata.get("creator", "")
                out.save(path)
                out.close()
                src.close()
            else:
                d = fitz.open(path)
                entry["producer"] = d.metadata.get("producer", "")
                entry["creator"] = d.metadata.get("creator", "")
                d.close()
            entry["status"] = "ok"
            ok += 1
        except Exception as e:
            entry["status"] = "skipped: %s: %s" % (type(e).__name__, str(e)[:120])
            print("skip", name, entry["status"])
            if os.path.exists(path):
                os.remove(path)
        manifest["documents"].append(entry)
    json.dump(manifest, open(MANIFEST, "w"), indent=1)
    print(f"real corpus: {ok}/{len(SOURCES)} documents in {DOCS}")
    print("manifest written:", MANIFEST)


if __name__ == "__main__":
    main()
