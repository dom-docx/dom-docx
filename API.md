# dom-docx API reference

dom-docx converts **semantic HTML fragments** into **native Word OOXML** (paragraphs, runs, lists, tables, images)—not raster snapshots or layout hacks.

For HTML authoring guidance (what converts well), see [AGENTS.md](./AGENTS.md). For validation scoring and test commands, see [SCORING.md](./docs/SCORING.md) and [README.md](./README.md).

---

## Quick start

```bash
npm install dom-docx
```

### Browser

```typescript
import { convertHtmlToDocx } from "dom-docx/browser";

const html = `
<h1 style="color:#1a1a2e">Quarterly Report</h1>
<p>Revenue grew <strong>12%</strong> year over year.</p>
<ul>
  <li>North America</li>
  <li>EMEA</li>
</ul>
`;

const blob = await convertHtmlToDocx(html);
// hand the Blob to a download (e.g. saveAs(blob, "output.docx"))
```

Runs entirely in the user's tab — **no Playwright, no Node `Buffer`**. See the browser entry point section below for the full API.

### Node

```typescript
import { writeFile } from "node:fs/promises";
import { convertHtmlToDocx } from "dom-docx";

const html = `
<h1 style="color:#1a1a2e">Quarterly Report</h1>
<p>Revenue grew <strong>12%</strong> year over year.</p>
<ul>
  <li>North America</li>
  <li>EMEA</li>
</ul>
`;

const docx = await convertHtmlToDocx(html);
await writeFile("output.docx", docx);
```

Pass a **body fragment only**—no `<!DOCTYPE>`, `<html>`, or `<body>` wrapper. Defaults: US Letter, 1″ margins, Arial 10.5 pt (14 px) body text — all configurable via [options](#options-convertoptions).

The default install is pure JS (`docx`, `cheerio`, `fflate`) — **no browser, no Playwright, no LibreOffice**.

## Two entry points

| | `dom-docx` (Node) | `dom-docx/browser` |
|--|-------------------|--------------------|
| Returns | `Promise<Buffer>` | `Promise<Blob>` (or `Uint8Array`) |
| `styleSource: "inline"` (default) | Pure JS, no browser | Pure JS, no live DOM required |
| `styleSource: "computed"` | Headless Chromium via **Playwright** (optional peer dep) | Native `getComputedStyle` on the **live page** — Playwright never involved |
| Typical use | Server-side batch conversion, agents with inline HTML | In-app "Export to Word" from rendered React/Vue/etc. |

---

## `convertHtmlToDocx(html, options?)` — Node

Primary entry point. Resolves styles, converts HTML, returns a **`Promise<Buffer>`** containing a valid `.docx` file.

```typescript
function convertHtmlToDocx(
  html: string,
  options?: ConvertOptions,
): Promise<Buffer>;
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `html` | `string` | Body fragment (trimmed and wrapped in `<body>…</body>` internally). |
| `options` | `ConvertOptions` | Optional. See [Options](#options-convertoptions). Defaults to inline style resolution. |

**Behavior by `styleSource`:**

| `styleSource` | Playwright needed? | What happens |
|---------------|--------------------|--------------|
| `"inline"` (default) | No | Parses `style=""` attributes only. Fast (~15–30 ms typical). |
| `"computed"` | **Yes** | Renders the fragment in headless Chromium, snapshots `getComputedStyle` for every element, then converts. Install once: `npm install playwright && npx playwright install chromium`. Slower (~100–500 ms+ depending on launch/reuse). |

When `styleSource: "computed"` and neither `page` nor `browser` is provided, the function launches Chromium, snapshots styles, converts, and **closes the browser** in a `finally` block. Pass `browser` or `page` to avoid the per-call launch cost (see [Usage patterns](#usage-patterns)).

---

## `convertHtmlToDocx(html, options?)` — browser (`dom-docx/browser`)

Client-side entry — **no Playwright, no Node `Buffer`**. Use this when the HTML is already rendered in the user's browser. Computed styles come from the **live DOM** via native `getComputedStyle`.

```typescript
import { convertHtmlToDocx } from "dom-docx/browser";

const blob = await convertHtmlToDocx(htmlFragment, { styleSource: "computed" });
// hand the Blob to a download (e.g. saveAs(blob, "output.docx"))
```

**Script tag** — the prebuilt IIFE (`dist/browser/dom-docx.browser.js`, built with `npm run build:browser`) exposes `window.domDocx`:

```html
<script src="dom-docx.browser.js"></script>
<script>
  const blob = await domDocx.convertHtmlToDocx(htmlFragment);
</script>
```

```typescript
interface BrowserConvertOptions extends DocumentConfig {
  styleSource?: "inline" | "computed";  // default "inline"
  document?: Document;                  // computed only; defaults to the host page's document
  imageResolver?: ImageResolver;
}
```

| `styleSource` | Live DOM required? | Behavior |
|---------------|--------------------|----------|
| `"inline"` (default) | No | Parses `style=""` only — works on a string fragment with no rendered page. |
| `"computed"` | **Yes** | Batch-reads native `getComputedStyle` from `document.body` (or `options.document`). The page must already render the same fragment — the converter does not inject HTML for you. |

All [document options](#options-convertoptions) (`pageSize`, `margins`, `metadata`, …) work here too. `browser` / `page` are Node-only.

| Export | Returns | Notes |
|--------|---------|-------|
| `convertHtmlToDocx(html, options?)` | `Promise<Blob>` | Primary browser API |
| `convertHtmlToDocxUint8Array(html, options?)` | `Promise<Uint8Array>` | Same bytes, no Blob wrapper |
| `buildDocxBlob` / `buildDocxUint8Array` | | Lower-level, bring your own `StyleResolver` |
| `snapshotComputedStylesFromDocument(doc?)` | `ComputedStyleSnapshot[]` | Style snapshots from a live `document.body` |

---

## Options (`ConvertOptions`)

```typescript
interface ConvertOptions extends DocumentConfig {
  styleSource?: "inline" | "computed";  // default "inline"
  browser?: Browser;                    // Node computed only (Playwright)
  page?: Page;                          // Node computed only (Playwright)
  imageResolver?: ImageResolver;
}

interface DocumentConfig {
  pageSize?: "letter" | "a4" | { width: number; height: number }; // custom in inches
  orientation?: "portrait" | "landscape";
  margins?: { top?: number; right?: number; bottom?: number; left?: number }; // inches
  defaultFont?: { family?: string; sizePt?: number };
  metadata?: {
    title?: string;
    subject?: string;
    creator?: string;
    keywords?: string[];
    description?: string;
  };
  headerHtml?: string;   // HTML fragment rendered as the page header
  footerHtml?: string;   // HTML fragment rendered as the page footer
  pageNumber?: boolean;  // centered "Page N" field appended to the footer
  lang?: string;         // spell-check locale, e.g. "en-US", "ar-SA"
  direction?: "ltr" | "rtl";
}
```

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `styleSource` | `"inline" \| "computed"` | `"inline"` | Which style resolution path to use. |
| `pageSize` | `"letter" \| "a4" \| {width,height}` | `"letter"` | Page size. Custom `{width, height}` in **inches**. |
| `orientation` | `"portrait" \| "landscape"` | `"portrait"` | Landscape swaps the page dimensions. |
| `margins` | `{top,right,bottom,left}` | `1` each | Page margins in **inches**; each side defaults to 1″. |
| `defaultFont` | `{family?, sizePt?}` | Arial, 10.5 pt | Default body font family and size (points). Applies to text with no explicit CSS font. |
| `metadata` | `{title,subject,creator,keywords[],description}` | — | Core document properties → `docProps/core.xml`. `keywords` is joined with `, `. |
| `headerHtml` / `footerHtml` | `string` | — | HTML fragment rendered as the page header / footer (its own inline-styled fragment). |
| `pageNumber` | `boolean` | `false` | Appends a centered `Page N` field to the footer (creates one if `footerHtml` is absent). |
| `lang` | `string` | — | Document spell-check locale (`w:lang`), e.g. `"en-US"`. |
| `direction` | `"ltr" \| "rtl"` | `"ltr"` | `"rtl"` sets right-to-left runs (`w:rtl`) — e.g. Arabic/Hebrew. |
| `imageResolver` | `ImageResolver` | — | Resolve non-`data:` `<img src>`. See [Images](#images--the-resolver-hook). |
| `browser` | Playwright `Browser` | — | **Node computed only.** Reuse an already-launched browser across many conversions. |
| `page` | Playwright `Page` | — | **Node computed only.** Snapshot styles from a page you already rendered. For in-browser apps use `dom-docx/browser` instead. |

```typescript
const docx = await convertHtmlToDocx(html, {
  pageSize: "a4",
  orientation: "landscape",
  margins: { top: 0.75, bottom: 0.75 },   // inches; left/right default to 1
  defaultFont: { family: "Georgia", sizePt: 11 },
  metadata: { title: "Q3 Report", creator: "Finance", keywords: ["revenue", "q3"] },
  headerHtml: "<p style='font-size:12px;color:#666'>Confidential</p>",
  pageNumber: true,
});
```

**Resolution order for `styleSource: "computed"` (Node):**

1. If `options.page` is set → snapshot that page (no new page)
2. Else if `options.browser` is set → new page per call, browser kept open
3. Else → launch Chromium, convert, close browser

---

## Images & the resolver hook

`<img>` embeds as a native DOCX `ImageRun` (not a link). Display size comes from the
`width`/`height` attributes, falling back to the image's intrinsic size (png/jpg/gif/bmp
headers are decoded), aspect-preserved if only one dimension is given.

**By default the library never makes a network or filesystem request.** Only inline
`data:` URLs (base64 png/jpg/gif/bmp) embed automatically. Any other `src` — `http(s):`,
`file:`, relative — is **not fetched**; it falls back to the `alt` text. This keeps
conversion deterministic and preserves a zero-egress guarantee (nothing leaves your
process based on input HTML — important for SSRF-safety and PHI/on-prem use).

To enable remote or local images, pass an **`imageResolver`**. You own the fetch and its
security policy; the library only orchestrates placement.

```typescript
type ResolvedImage = {
  data: Uint8Array | ArrayBuffer;      // raw image bytes (not base64)
  type: "png" | "jpg" | "gif" | "bmp";
  width?: number;                       // used only if the <img> omits width/height
  height?: number;
};

type ImageResolver = (
  src: string,
) => Promise<ResolvedImage | null> | ResolvedImage | null;
```

Behavior:

- Called once per `<img>` whose `src` is **not** already a `data:` URL, before conversion.
- Return `ResolvedImage` to embed, or `null` to skip (image falls back to alt text).
- A resolver that **throws** for one image is caught per-image (that image falls back);
  the conversion never aborts.
- Multiple images resolve concurrently (`Promise.all`).

```typescript
const docx = await convertHtmlToDocx(html, {
  imageResolver: async (src) => {
    // YOUR policy: allowlist hosts, block private IPs/SSRF, add auth, cap size…
    const url = new URL(src);
    if (url.hostname !== "cdn.example.com") return null;
    const res = await fetch(src);
    if (!res.ok) return null;
    return { data: new Uint8Array(await res.arrayBuffer()), type: "png" };
  },
});
```

> The library ships no default network resolver on purpose — making a request must be an
> explicit, visible line of caller code, never an implicit side effect of conversion.

---

## Usage patterns

### Default — inline styles

```typescript
const docx = await convertHtmlToDocx(html);
```

Best for agent-generated HTML with explicit inline styles. No Playwright required on Node or in the browser.

### Stylesheet / class-based HTML (Node — Playwright required)

```typescript
const html = `
<style>
  .hero { background: #eaeaea; padding: 10px 16px; }
  .hero h1 { color: #1a1a2e; margin: 0; }
</style>
<div class="hero">
  <h1>Title</h1>
</div>
`;

const docx = await convertHtmlToDocx(html, { styleSource: "computed" });
```

On **Node**, the computed path renders the fragment in headless Chromium. External `<link rel="stylesheet">` works if the URL loads during `setContent` (`waitUntil: "networkidle"`).

In a **browser app**, use `dom-docx/browser` with `styleSource: "computed"` instead — render the HTML in the page, then convert; no Playwright.

### Reuse a browser in a loop

```typescript
import { chromium } from "playwright";
import { convertHtmlToDocx } from "dom-docx";

const browser = await chromium.launch();
try {
  for (const html of fragments) {
    const docx = await convertHtmlToDocx(html, {
      styleSource: "computed",
      browser,
    });
    // ...
  }
} finally {
  await browser.close();
}
```

### Snapshot from an existing page

```typescript
// page already has your HTML rendered (Playwright)
const docx = await convertHtmlToDocx(html, {
  styleSource: "computed",
  page,
});
```

Styles come from the **same DOM** as a reference screenshot—no second render.

---

## Supported HTML & CSS

### Elements

`h1`–`h6`, `p`, `div`, `section`, `ul`, `ol`, `li`, `table`, `thead`, `tbody`, `tfoot`, `tr`, `td`, `th`, `blockquote`, `hr`, `figure`, `figcaption`, `img`, `svg` (low-complexity), `strong`, `b`, `em`, `i`, `u`, `a`, `span`, `code`, `br`, `pre` (limited).

Element attributes: table `border` / `cellpadding` / `cellspacing` / `colspan`; `href` on links; `src` / `width` / `height` / `alt` on images; list `type`. `<img>` embeds `data:` URLs by default; other `src` schemes require an [`imageResolver`](#images--the-resolver-hook). Other attributes are mostly ignored.

Unsupported tags are treated as generic block containers or skipped.

### Inline CSS properties

Parsed from `style=""` (and from computed snapshots on the computed path):

| Property | Notes |
|----------|-------|
| `color` | Hex, `rgb()`, `rgba()` (alpha 0 → ignored) |
| `background`, `background-color` | Hex / rgb; `transparent` ignored |
| `text-align` | `left`, `center`, `right`, `justify` |
| `font-size` | `px`, `pt`, `em` |
| `font-weight`, `font-style` | Including `bold`, `600`, `italic` |
| `list-style-type` | decimal, lower/upper-alpha, lower/upper-roman, disc, circle, square |
| `margin`, `margin-*` | `px`, `pt`, `em` |
| `padding`, `padding-*` | `px`, `pt`, `em` |
| `border`, `border-*` | Width + color shorthand |
| `display` | `block`, `inline-block`, `flex` |
| `flex-direction` | `row`, `column` |
| `gap`, `row-gap`, `column-gap` | px |

All other CSS properties are silently ignored.

### What converts well / poorly

See [AGENTS.md](./AGENTS.md) for the full tier list. In short:

- **Excellent:** headings, paragraphs, lists, simple tables, inline formatting, short span highlights, blockquotes, `<hr>`.
- **Good:** shaded div banners, flex rows (≤4 items), table row/cell backgrounds, bordered boxes, `data:` images.
- **Avoid:** complex SVG, CSS grid/float/absolute layout, external stylesheets (inline path), forms, deep layout div nesting.

### Input contract

- **Fragment only** — content that would go inside `<body>`, not a full document.
- **Inline CSS preferred** for the default path — `style="..."` on elements.
- **Stylesheets** (`<style>` blocks, classes) — require `styleSource: "computed"`.
- **No JavaScript** — conversion is static HTML → OOXML.

---

## Lower-level API

### `buildDocxBuffer(html, styleResolver, imageResolver?, documentConfig?)`

Use when you already have a **`StyleResolver`** (or want to convert many fragments with one resolver / one browser session).

```typescript
import { buildDocxBuffer, INLINE_STYLE_RESOLVER } from "dom-docx";

const docx = await buildDocxBuffer(html, INLINE_STYLE_RESOLVER);
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `html` | `string` | Same body fragment as `convertHtmlToDocx`. |
| `styleResolver` | `StyleResolver` | Supplies `getCss()` for every element during the visit. |
| `imageResolver` | `ImageResolver` | Optional. Same hook as [`ConvertOptions`](#images--the-resolver-hook). |
| `documentConfig` | `DocumentConfig` | Optional. Page/font/metadata options. |

Platform-neutral variants: **`buildDocxUint8Array`** (same signature, returns `Uint8Array`) and **`buildDocxBlob`** (returns `Blob`) — exported from both entry points.

Use these for benchmark loops that reuse one resolver, tests that inject a mock resolver, or pipelines that snapshot styles once and convert multiple times.

### `StyleResolver`

Every element asks a `StyleResolver` for a normalized `ParsedCss` object (color, margins, borders, flex, …). The visitor, table builder, and inline collector all call `styleResolver.getCss(element)` — they never read `style=""` directly. Implement the interface to inject test doubles or alternate style sources.

```typescript
interface StyleResolver {
  readonly source: "inline" | "computed";
  getCss(element: Element): ParsedCss; // cheerio/domhandler Element; {} if no styles apply
}
```

Built-in resolvers and helpers (exported from `dom-docx`):

| Export | Description |
|--------|-------------|
| `INLINE_STYLE_RESOLVER` | Singleton inline resolver — parses each element's `style=""` attribute. No stylesheets, no class selectors. |
| `ComputedStyleResolver` | Computed resolver built from a snapshot array; looks up elements by stable DOM path. Construct via `ComputedStyleResolver.fromSnapshots(snapshots)`. |
| `createComputedStyleResolver(html, browser)` | Node: renders the fragment in a new Playwright page, snapshots styles, closes the page. |
| `computedStyleResolverFromPage(page)` | Node: snapshot an existing Playwright page (same DOM as a reference screenshot). |
| `snapshotComputedStyles(page)` | Node: raw `{ path, styles }[]` snapshots from a Playwright page. |
| `snapshotComputedStylesFromDocument(doc?)` | Browser (`dom-docx/browser`): same snapshots from a live `document`. |

**Computed style properties captured** (via `getComputedStyle`): `color`, `backgroundColor`, `display`, `flexDirection`, `gap` / `columnGap` / `rowGap`, `textAlign`, `fontSize`, `fontWeight`, `fontStyle`, margin and padding sides, per-side border width/color. UA defaults on headings are partially stripped when the element has no inline override, so the computed path stays aligned with the inline path for bare `<h1>`–`<h6>`.

---

## How the engine works

Conversion is a three-stage pipeline: **style resolution → HTML visitor → OOXML pack + patch**.

```
HTML fragment
     │
     ▼
┌────────────────────────────────────────────────────────────────┐
│ 1. Style resolution (StyleResolver)                            │
│    inline: parse style="" on each element                      │
│    computed (Node): getComputedStyle via Playwright/Chromium   │
│    computed (browser): getComputedStyle on live document.body  │
└────────────────────────────────────────────────────────────────┘
     │
     ▼
┌────────────────────────────────────────────────────────────────┐
│ 2. Visitor (cheerio → docx objects)                            │
│    Walk body children; map blocks to Paragraph / Table         │
│    Inline nodes → TextRun, Hyperlink, list numbering           │
│    Flex divs → borderless tables; shaded divs → table wrap     │
└────────────────────────────────────────────────────────────────┘
     │
     ▼
┌────────────────────────────────────────────────────────────────┐
│ 3. Pack + post-process                                         │
│    docx Packer → unzip → patch document.xml / numbering.xml    │
│    → re-zip → Buffer (Node) or Blob (browser)                  │
└────────────────────────────────────────────────────────────────┘
```

### Visitor mapping

| HTML | OOXML |
|------|-------|
| `h1`–`h6` | Heading paragraphs (Word heading levels) |
| `p`, text flow | `Paragraph` + `TextRun` |
| `ul` / `ol` / `li` | Native numbering; honors `list-style-type` and the `type` attr |
| `table` / `tr` / `td` / `th` | `Table` with borders, shading, colspan |
| `div` with `display:flex` | Borderless flex table (row or column) |
| `div` with background/border | Shaded or bordered 1×1 table wrapper when needed |
| `blockquote` | Indented paragraph with left border |
| `a` | Hyperlink runs |
| `img` | Embedded `ImageRun` (`data:` URLs, or via `imageResolver`) |
| `svg` (low-complexity) | Native DOCX blocks — `<rect>` bands + `<text>` (bar/funnel charts) |
| `strong`, `em`, `span`, `code`, `br` | Inline runs / breaks |

### OOXML post-processing

After `docx` packs the document, the engine unzips the buffer and patches XML the library cannot express cleanly:

- **`numbering.xml`** — LibreOffice needs list tab stops as `w:val="num"` (not `"left"`) and drops tentative numbering flags.
- **`document.xml`** — Shaded paragraphs with exact line spacing get vertical text alignment so PDF export centers padding correctly.

These patches are applied automatically; callers receive a finished `.docx`.

---

## Document defaults

Omitting all options produces:

| Setting | Value |
|---------|-------|
| Page size | US Letter (8.5″ × 11″), portrait |
| Margins | 1″ all sides |
| Body font | Arial 10.5 pt (= 14 px at 96 dpi) |
| Line height | 1.4 |
| Text color | `#111` on white |

These match the harness wrapper (`wrapHtml()` in `src/html-wrap.ts`) that the computed path and the visual validator render against, so default output aligns byte-for-byte with the validated baseline. Override any of them via [`ConvertOptions`](#options-convertoptions).

---

## Dependencies & environment

| Requirement | Node (`dom-docx`) | Browser (`dom-docx/browser`) |
|-------------|-------------------|------------------------------|
| **Node.js ≥ 20** | Required | N/A (runs in the user's browser) |
| **cheerio, docx, fflate** | Installed as dependencies | Bundled in the IIFE |
| **playwright** | **Only** for `styleSource: "computed"` (optional peer dependency, installed separately) | **Not used** |
| **Live DOM** | Not required (computed uses Playwright) | Required only for `styleSource: "computed"` |

LibreOffice (`soffice`) is **not** required for conversion — only for the visual validation loop in `npm run score:suite`.

For the Node computed path, install Playwright plus Chromium yourself, once:

```bash
npm install playwright
npx playwright install chromium
```

Contributors cloning this repo: `npm run setup` (same command, via package script).

**Playwright is not a dependency of the browser bundle.** It appears only in the Node computed path (and in this repo's test harness, which uses Playwright to *test* the browser bundle — dev tooling, not a runtime requirement for client apps).

---

## Limitations

- **Computed path cost (Node)** — launching Chromium per call is expensive; pass `browser` or `page` in hot loops. In the browser bundle, computed styles are free aside from normal layout.
- **Stylesheet fidelity** — the computed path improves cascade support, but some patterns (e.g. complex themed sections) still score lower than inline-authored equivalents.
- **Layout CSS** — CSS grid, floats, and absolute positioning have no OOXML equivalent and are ignored; flex support covers simple row/column cases.
- **Images** — png/jpg/gif/bmp only; svg `<img>` sources are not rasterized (inline low-complexity `<svg>` elements convert natively).
- **No JavaScript execution** — the fragment is converted statically; anything rendered client-side must be in the HTML string (or live DOM for the browser computed path).

---

## Validation commands (repo development)

These exercise the API and write artifacts under `output/`:

| Command | What it runs |
|---------|----------------|
| `npm run score:suite` | Full visual + XML regression suite (cases: `tools/generator.ts`; needs Chromium + LibreOffice) |
| `npm run score:suite:priority` | Fast subset of the same cases |
| `npm run guard:inline` | Asserts inline path OOXML equivalence (normalized XML) |
| `npm run guard:config` | `ConvertOptions` OOXML checks |
| `npm run score:benchmark` | OSS html-to-docx / TurboDocx comparison |
| `npm run build:browser` | esbuild → `dist/browser/dom-docx.browser.js` |
| `npm run typecheck` | TypeScript compile check |

Browser bundle parity, style-source, and CSS-cascade guards: see [CONTRIBUTING.md](./CONTRIBUTING.md#maintainer-only-harness-commands).

---

## Related documentation

| Doc | Contents |
|-----|----------|
| [AGENTS.md](./AGENTS.md) | How to write HTML that converts well |
| [README.md](./README.md) | Install, quick start, API overview |
| [SCORING.md](./docs/SCORING.md) | Validation methodology and engine score |
| [TEST-SCORES.md](./docs/TEST-SCORES.md) | Latest suite metrics |
| [BENCHMARK.md](./docs/BENCHMARK.md) | Inline vs computed and CSS cascade benchmarks |
