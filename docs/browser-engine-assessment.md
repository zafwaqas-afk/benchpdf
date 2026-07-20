# Porting the PDF→PPTX quality pipeline to the browser

**Status: decision pending. Do not start the port on the strength of this document alone.**
Written 2026-07-19, after the browser converter was withdrawn from the site.

## What we shipped, and what it actually did

The JS converter was not a weaker version of the engine. It was a different,
much simpler program that happened to emit a `.pptx`. Measured by
`tests/fidelity_suite.py --engines=browser` on the committed fixtures:

| Invariant | python | browser |
|---|---|---|
| Paragraph-level blocks | 0 fragments | **242 fragments** on the 9-page report |
| Native tables | 17 of 17 | **0 of 17** |
| Font families preserved | 3 of 3 | **1** (hardcoded Arial) |
| Bold weight | 21 bold runs | **0** |
| Graphic layer | present | **absent on every page** |
| Element overlaps >10% | 0 | **240** |
| Boxes on the worst page | 6 | **116** |

Two causes, both structural rather than a bug to fix:

1. `groupLines()` groups text items onto shared baselines and stops. There is no
   second pass clustering lines into paragraphs, and no column model, so items
   on the same baseline in different table cells are concatenated into one string.
2. The slide is built from text alone. `page.render()` is only called when the
   page has under four characters, as a scanned-page fallback. Every fill, rule,
   chart and table shading is simply never drawn.

## Is parity feasible in the browser?

Yes for four of the five pieces. pdf.js exposes enough:

| Piece | Feasibility | Notes |
|---|---|---|
| **Background layer** | **Straightforward.** | `page.getOperatorList()`, filter out the text-showing operators (`OPS.showText`, `showSpacedText`), render the remainder to canvas, place as a full-slide image behind the text. This alone fixes the single most visible failure. |
| **Block clustering** | **Moderate.** | `getTextContent()` gives per-item transform, width, height and `fontName`. Paragraph clustering is geometry over that: line grouping (exists), then vertical-gap and left-edge agreement to merge lines into blocks. This is a direct port of the Python logic, not new research. |
| **Font mapping** | **Moderate, and permanently approximate.** | `getTextContent({includeMarkedContent:false})` plus `page.commonObjs` gives the source font name and flags, so weight and italic are recoverable and the current total loss of both is inexcusable. Mapping to a *target* face is the problem: see below. |
| **Table detection** | **Hard, the real cost.** | The Python engine leans on PyMuPDF `find_tables(strategy="lines")`, which has no JS equivalent. Ruling lines are recoverable from the operator list as stroked paths, so a lines-strategy detector is buildable: collect stroked segments, cluster into a grid, assign text to cells. This is the bulk of the work and the bulk of the risk. |

## What stays impossible client-side

- **Knowing which fonts the reader has.** The desktop engine resolves a source
  face against the fonts actually installed on the machine. A browser cannot
  enumerate local fonts without the Local Font Access API, which is Chromium-only
  and permission-gated. The realistic browser ceiling is a fixed mapping table to
  the common web-safe families. That is a genuine quality gap, not a bug.
- **Anything requiring Office.** Unchanged: PDF→Word, Word/Excel/PowerPoint→PDF
  stay desktop-only regardless of what happens here.
- **OCR for scanned pages.** Same as desktop; neither has it.
- **Large documents at comfort.** The whole pipeline runs on the main thread of
  the visitor's machine. The 9-page fixture is fine; a 200-page report with a
  rendered background layer per page is a memory problem a server or a desktop
  process does not have.

## Estimate

Assuming one engineer, and counting only work to reach suite parity on the
committed fixtures:

| Workstream | Estimate |
|---|---|
| Background layer via operator-list filtering | 2 to 3 days |
| Paragraph clustering ported from the Python engine | 4 to 6 days |
| Font family/weight/style mapping (fixed table) | 2 to 3 days |
| Table detection from stroked paths | 8 to 12 days |
| Harness, suite parity, cross-browser, memory work | 4 to 5 days |
| **Total** | **4 to 6 weeks** |

The table detector is over a third of that and carries most of the schedule risk.
It is also the piece most likely to land at "usually right" rather than "right",
which is the quality posture we just withdrew a product for.

## Recommendation: keep PDF→PPTX desktop-only for now

Three reasons, in order of weight:

1. **The ceiling is below the promise.** Even a finished port cannot resolve
   local fonts, so a document in a non-web-safe face converts visibly worse in
   the browser than on the desktop. Our positioning is output fidelity. Shipping
   a knowingly second-best path under the same name reintroduces exactly the
   problem we are fixing, just with better numbers.
2. **The cost buys a duplicate, not a new capability.** Four to six weeks
   produces a second implementation of something that already works, and then two
   engines to keep at parity forever. The same weeks spent on OCR, or on a Mac
   build, would add something no BenchPDF user can do today.
3. **The funnel argument is weaker than it looks.** The drop-off concern that
   motivated the browser converter is real, but PDF→Images and PDF→Text still
   run in the browser and still convert visitors. The interstitial now sets an
   honest expectation before the click rather than after the download.

**Revisit if** one of these changes: PDF→PPTX becomes the dominant requested
conversion from browser traffic; or the background layer alone (2 to 3 days,
the cheapest item here) turns out to close enough of the perceived gap to ship a
deliberately-labelled "quick preview" conversion that is *named* as lower
fidelity rather than presented as the product.

**If we do port it**, the order is fixed: background layer, then clustering, then
fonts, then tables. Each lands as a suite improvement measurable by
`--engines=browser`, and the site relinks the conversion only when that column
reads GREEN.
