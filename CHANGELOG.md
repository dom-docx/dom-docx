# Changelog

All notable changes to this project will be documented in this file.

Format based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## 0.1.14

### Fixed

- **Table rows whose cells hold only invisible-but-real content are no longer squashed.** The spacer-row collapse (which keeps truly empty divider rows from inflating to a full line box) also fired on cells containing `&nbsp;`, zero-width space, soft hyphen, `<br>`, or `<wbr>` — browsers keep a full-height line box for all of these. Such cells now count as intentional content and the row gets its natural height; genuinely whitespace-only rows still collapse. Also covers typographic spaces (`&ensp;`, `&emsp;`, `&thinsp;`), which JS `\s` matches but browsers render at full width. Suite case: `table-empty-cell-row-height`. (Thanks to Alexander Wilms for a community patch.)

## 0.1.13

### Changed

- **Default text color now uses Word's `auto` instead of a forced near-black.** Text with no explicit CSS color previously had `#111111` stamped on every run; it now omits the run color so Word/LibreOffice apply their default (which renders black, and adapts to the background). This is more native and fixes a latent bug: a dark-background block whose text had no explicit color rendered dark-on-dark (invisible) — `auto` renders it light-on-dark. Hyperlinks keep their link color. (Thanks to Alexander Wilms for a community patch.)

### Fixed

- **CSS `padding` on a table cell is now applied.** Previously only the table's `cellpadding` attribute set cell margins, so `<td style="padding:16px">` produced no `w:tcMar` and rendered flush. Per-side CSS padding on a `<td>`/`<th>` now becomes cell margins, with the table `cellpadding` attribute filling any side the cell doesn't set. Suite case: `table-cell-padding`. (From a community patch.)
- **Physical CSS length units (`mm`, `cm`, `in`, `pc`) are now parsed correctly.** They previously fell through to the bare-number branch and were treated as pixels, so `10mm` came out as `10px` (~2.6mm — about 3.85× too small); `in`/`pc` had the same silent mis-parse. All four now convert to the correct twip lengths (exact factors: 1440 twips/inch, 25.4mm/inch), in both the inline and computed length parsers. Suite case: `css-length-units`. (From a community patch.)

## 0.1.12

Add Related Projects section to README.

## 0.1.11

### Fixed

- **`<th>` header cells are centered by default, matching browsers.** Chromium/Firefox apply the UA default `th { text-align: center }`; dom-docx left header text left-aligned unless an explicit `text-align`/`align` was present, so plain `<th>` headers sat visibly off from the browser render (the `adjacent-tables` suite case scored H 74 from this alone — 85.4% → 96.1% after the fix). An explicit alignment still wins.
- **Suite oracle: hidden/overlay content is excluded from expected text regardless of how it hides.** The text-fidelity oracle only stripped hidden elements carrying a `style` attribute, so content hidden via overlay semantics (`role="tooltip"`, closed `<dialog>`, `-dialog`/`-tooltip` components) stayed in the EXPECTED text and the converter was penalized for correctly skipping it (`tooltip-skipped` 86.6% → 96.3%). All 42 suite cases now score ≥ 90% (avg 96.3%).

## 0.1.10

### Fixed

- **Adjacent tables no longer merge (data tables collapsed to sliver width).** Word and LibreOffice fuse consecutive sibling `w:tbl` elements into ONE table, so any table emitted directly after another inherited a merged grid — on a real docs page an empty 0.4"-wide icon-chrome table followed every data table, and the merge collapsed each data table to a ~1-character-wide column. Two fixes: an invisible 1-twip separator paragraph is inserted between consecutive tables so each keeps its own grid, and flex containers whose items have no visible content (no text, media, fill, or border — e.g. icon-only web-component chrome whose glyph lives in shadow DOM) no longer emit an empty table at all. Suite case: `adjacent-tables`.
- **Images no longer smash against the next heading or paragraph.** Web layouts separate figures with margins the computed path faithfully zeroes (flex/grid `gap`, container padding), so a docs figure rendered flush against the following heading or body text in the flat docx. An image paragraph's before/after spacing is now floored to ~0.5em (a larger real margin still wins); skipped for flex-card content, which manages its own tight rhythm. Guard: `npm run guard:image-spacing` (structural — the floor deliberately diverges from a bare `margin:0` browser render, so it can't be scored by the visual suite).
- **Transient overlay content (dialogs, tooltips, popovers) is no longer rendered.** Overlay content shown only on a user action is not part of the linear document, yet two kinds leaked in: a figure's click-to-expand modal held a full-size **duplicate** of the same image (two `<img>`, both emitted), and a heading's copy-link **tooltip** bled its label into the text (headings ended in "Copy link"). Native `<dialog>` without `open`, overlay web components matched by the `-dialog`/`-modal`/`-tooltip`/`-popover` custom-element suffix (`<rh-dialog>`, `<rh-tooltip>`, `<sl-popover>`), and `role="dialog"`/`role="alertdialog"`/`role="tooltip"`/`aria-modal="true"` are now skipped. Suite cases: `modal-dialog-skipped`, `tooltip-skipped`.

## 0.1.8

### Fixed

- **Admonition / alert callouts render as a box instead of plain text.** Docs sites mark notes up as web-component alerts (`<rh-alert>`) or `class="admonition note"` blocks whose box (fill, accent bar, icon) is drawn in shadow DOM or an external stylesheet — the light-DOM host reads as transparent, so a "Note" flattened to plain text. Callout containers (recognized by alert tag or `note`/`tip`/`important`/`warning`/`caution`/`danger`/`admonition`/`callout`/`alert` class token) now get a shaded box with a severity-colored left accent bar (amber for warning/caution, red for danger/important, green for tip, blue for note), and the title label ("Note") is bolded. A callout that already carries a real background keeps it. Suite case: `admonition-note`.
- **Table column widths from `<colgroup>` are now honored.** Tables that size their columns entirely through `<colgroup><col style="width:N%">` (common on docs sites and DocBook/AsciiDoc output) carried no width signal dom-docx recognized, so every column collapsed to its min-content width — a real docs table rendered as a single ~1-character-wide column. `<col>` widths (percent or px, with `span`) now seed the per-column widths and the table's natural width, so a colgroup-sized table fills the page with the declared proportions. Suite case: `table-colgroup-widths`.
- **Computed styles no longer follow dark-mode browser themes into Word.** A DOCX is a light canvas (white page, dark text). Snapshotting `getComputedStyle` from a tab with `prefers-color-scheme: dark` stamped near-white text colors with no fill — invisible in Word/LibreOffice (seen converting Red Hat docs from a dark macOS browser). The Node/Playwright computed path now forces `colorScheme: "light"` for the snapshot; both paths drop near-white text colors when there is no dark block background (light text on an intentionally dark shaded callout is kept). Guard: `npm run guard:document-canvas`.
- **Empty flex containers no longer render as solid colored boxes.** On the computed path, docs sites wrap figures in a `display:flex` "open in modal" control with a background color and hidden/lazy inner content; when every flex item was hidden, the container still emitted an empty flex table painted with its background — a bare colored box (the "blue box" seen converting a Red Hat docs page, where all five figure wrappers rendered this way). A flex container with no visible items now renders nothing. Flex-item visibility is also resolved with the active style source, so computed `display:none` items are filtered correctly. Separately, a `background-color` on a media-only container (an `<img>`/`<svg>`/`<canvas>`/`<picture>` wrapper with no text) is dropped so the image stands on its own; backgrounds on containers with real text (callouts, blockquotes) are unaffected.
- **Images with `display: inline-block` (or `block`) are no longer silently dropped.** On the computed path, docs sites style images (e.g. a `modal-img` thumbnail) as `inline-block`, which made `isBlockElement` route the `<img>` to the block-container dispatch — which has no image handling — so the element vanished with no drawing and no alt text (every figure on a Red Hat docs page: 10 images resolved, 0 rendered). `<img>`/`<picture>` are now always treated as replaced inline media and rendered via the inline-run path regardless of computed `display`.
- **Oversized images are clamped to the page width.** An `<img>` with no `width`/`height` attributes was emitted at its full natural pixel size, so a high-resolution image (e.g. a 1520px docs diagram) rendered ~16" wide — overflowing the margins, clipping its right edge, and spilling across pages. Display size is now capped at the printable content width (aspect-preserved), like a browser's `max-width: 100%`. Images already sized at or below the content width are unaffected.
- **Flex cards no longer clip wrapping content.** A `display:flex` container sizes each card row from an estimate of one line per block child; that estimate was applied as an EXACT row height, so any content taller than the estimate was cropped — a long heading inside a flex wrapper (e.g. a docs page's `.titlepage`) wrapped to a second line and had that line sliced off in LibreOffice. Flex card rows now use AT_LEAST height: they still floor at the estimate (so tight single-line cards don't inflate) but grow to fit wrapping/nested content. The `flex-row-horizontal` visual suite case now uses multi-line wrapping content as a regression guard.

### Added

- **`npm run guard:document-canvas`** — regression for light-canvas color handling: near-white computed text with no dark fill must not be stamped into OOXML; light text on a dark shaded block is kept; optional Playwright check that the Node computed path forces `prefers-color-scheme: light`.
- **`onWarning`** — diagnostic callback for conditions that don't fail conversion but silently degrade the output: images with no `imageResolver` (or that it couldn't resolve) fall back to alt text, and class/stylesheet-based CSS is ignored on the default `styleSource: "inline"`. Previously both failed silently — a real-world page (external stylesheet, hosted images) would convert "successfully" while quietly losing every image and all table/callout styling. Defaults to `console.warn`; pass `null` to suppress. Available on both the Node and browser entries.

## [0.1.6] - 2026-07-12

### Added

- **`tocHtml`** — caller-provided table-of-contents "slot": an HTML fragment placed after the cover page (if any) and before the body. You control the markup and styling (numbered, boxed, columns…); in-page links (`<a href="#id">`) jump to the matching `id` in the body. Add a trailing `<div style="break-after:page"></div>` to put it on its own page. Guard: `npm run guard:toc-slot`.
- **Internal links (`href="#id"`)** — same-document fragments become Word internal hyperlinks (`w:hyperlink w:anchor`) targeting bookmarks emitted for matching `id` attributes (and legacy `<a name>`). External URLs unchanged. Guard: `npm run guard:internal-href`.

### Removed

- **`tableOfContents`** (breaking) — the auto-generated, page-number-less TOC field added in 0.1.5 is removed in favor of `tocHtml` + internal links. Rather than dom-docx generating and styling the TOC, you provide the exact TOC markup and it wires up the `#id` navigation — full styling control, no field, no "update fields" prompt, nothing to keep in sync. (`TableOfContentsConfig` type and `guard:toc` removed.)

### Fixed

- **Headings no longer smash into the block above them.** Web layouts routinely zero heading margins and rely on flex/grid `gap` or container padding for spacing — which has no equivalent in a flat docx, so on the computed path a `margin-top: 0` heading rendered flush against the previous paragraph or table. Heading top spacing is now floored to ~0.5em of the heading font (a larger real margin still wins; the floor is skipped inside flex cards).

## [0.1.5] - 2026-07-10

### Added

- **`coverHtml`** — HTML fragment rendered as a cover page: the first content in the document, before the table of contents, followed by an automatic page break. Inline styles + `data:` images (e.g. a logo); header/footer/page number are suppressed on the cover page, and headings inside the cover are not treated as TOC entries.

### Changed

- **`tableOfContents` is now a clickable, page-number-less table of contents** (reworked from the page-numbered field shipped in 0.1.4). Each entry hyperlinks to a bookmark on its heading; page numbers are omitted — they depend on layout the library does not compute — so the entry list is **complete at creation**: correct in every viewer (Word, LibreOffice, Google Docs, PDF/preview) with **no field update and no "update fields" prompt**. Entries are styled as a document outline for readability — bold near-black top levels, lighter/greyer deeper levels, indent + spacing per level — rather than a wall of blue underlined links (they stay clickable). The field carries the `\n` switch so a reader who refreshes it still gets a number-less TOC, and heading styles keep their `w:outlineLvl` so that refresh rebuilds correctly in LibreOffice.

### Fixed

- **Flex row of bordered cards wrapping an `<img>`** rendered wrong in LibreOffice: a doubled border (the item was emitted both as the flex cell _and_ as a nested bordered table), the top of the image clipped (an EXACT line box cropped the inline image above the baseline), and `min-height` ignored (the card hugged the image). Flex items now render their children into the single bordered cell; image lines use AT_LEAST height so they aren't cropped; `min-height` is parsed and applied as an AT_LEAST row height; and card content is top-aligned like the browser. (`flex-row-images` visual 51.6% → 93.7%; full-suite score unchanged.)

## [0.1.4] - 2026-07-10

### Added

- **`tableOfContents`** — insert a native Word Table of Contents field built from the document's `h1`–`h6` (mapped to Word Heading 1–6). `true` for defaults, or `{ title?, headingRange?, hyperlink?, pageBreakAfter? }`. Heading titles are cached into the field so it is visible immediately; page numbers are filled by the word processor on field update (Word updates on open via the dirty/`updateFields` flags). Heading styles get explicit `w:outlineLvl` so LibreOffice's "Update Table of Contents" repopulates instead of emptying the table. _(Reworked to a clickable, page-number-less TOC in Unreleased — see above.)_
- **`npm run guard:toc`** — structural regression for the TOC field, cached entries, heading outline levels, and OOXML schema validity (CI, no Playwright/LibreOffice)

### Changed

- The Node and browser entries now forward the whole options object to the builder (which reads only the `DocumentConfig` fields), so a new option reaches both entries automatically — replacing the two hand-maintained per-entry option lists that could silently drift
- **`guard:config`** runs its option → OOXML assertions through **both** the Node (`convertHtmlToDocx`) and browser (`convertHtmlToDocxUint8Array`) entries, so per-entry forwarding drift fails in CI
- Browser entry types (`dist/browser.d.ts`) are generated by tsc from `src/browser.ts` instead of a hand-maintained declaration file, so the published `dom-docx/browser` types can't fall out of sync with the code

## [0.1.3] - 2026-07-09

### Added

- **`rasterizeInPlace`** — rasterize live `<canvas>` and complex `<svg>` (e.g. chart libraries) to PNG `<img>` before conversion, on both the browser bundle and the Node Playwright path
- **`rasterizeInPlace.scale`** — optional supersampling factor (default `1`, max `4`); `scale: 2` recommended for sharper chart PNGs in Word
- Page breaks are **CSS-only** — use `break-before` / `break-after` (inline or via computed stylesheets); no special class names
- **`options.root`** on the browser API — pass the live export element when using `styleSource: "computed"` from an SPA so computed styles resolve against the correct subtree
- **`npm run guard:page-break`** — structural regression (OOXML + multi-page PDF via LibreOffice; not part of the single-page visual suite)

### Fixed

- **SPA computed-style export** — style snapshots are scoped to the export root with inline-style fallback when computed values are unavailable
- **Flex layout + rasterized charts** — flex row/column sizing and paragraph line spacing now account for rasterized `<img>` dimensions so chart images are not clipped
- **Visual scoring** — rasterized text inside SVGs (e.g. chart labels) is no longer incorrectly penalized by the fidelity metric

### Changed

- Test/score npm scripts reorganized into **`score:*`**, **`guard:*`**, and **`research:*`** tiers; hardcoded suite case counts removed from docs
- **`npm run docs:sync`** — regenerate scoring docs from harness JSON
- README, API.md, AGENTS.md, CONTRIBUTING.md, and BENCHMARK.md refreshed (browser quick start, `rasterizeInPlace`, page breaks, dom-docx.com links)

## [0.1.1] - 2026-07-05

### Fixed

- Shaded callouts and blockquotes no longer lose their accent bar or background fill on the computed-style path
- Computed-style row-height leak that could inflate paragraph spacing in table rows

### Changed

- API.md and TEST-SCORES documentation updates

## [0.1.2] - 2026-07-05

### Fixed

- **`dom-docx/browser` npm exports** — bundlers now resolve the ESM entry correctly; the IIFE bundle is included in the published tarball

## [0.1.0] - 2026-07-02

### Added

- `convertHtmlToDocx()` — HTML body fragment → native OOXML (paragraphs, runs, lists, tables, images)
- Default **`inline`** style path (pure JS: `docx` + `cheerio` + `fflate`, no browser)
- Optional **`computed`** style path via Playwright `getComputedStyle` (optional dependency)
- Browser bundle entry (`dom-docx/browser`) for in-page conversion
- Page options: size, orientation, margins, default font, metadata, headers/footers, page numbers, lang/direction
- `list-style-type` support for ordered and unordered lists
- `data:` URL image embedding with optional `imageResolver` hook
- Low-complexity inline SVG → native DOCX bands
- Committed [examples/](./examples/) with sample HTML, DOCX, and side-by-side previews
- Visual regression harness in `tools/` (33-case loop, OSS benchmark, scoring methodology)

[0.1.6]: https://github.com/dom-docx/dom-docx/releases/tag/v0.1.6
[0.1.5]: https://github.com/dom-docx/dom-docx/releases/tag/v0.1.5
[0.1.4]: https://github.com/dom-docx/dom-docx/releases/tag/v0.1.4
[0.1.3]: https://github.com/dom-docx/dom-docx/releases/tag/v0.1.3
[0.1.2]: https://github.com/dom-docx/dom-docx/releases/tag/v0.1.2
[0.1.1]: https://github.com/dom-docx/dom-docx/releases/tag/v0.1.1
[0.1.0]: https://github.com/dom-docx/dom-docx/releases/tag/v0.1.0
