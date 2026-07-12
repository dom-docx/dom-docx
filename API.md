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
| `rasterizeInPlace` | Same Playwright page as computed (or spawn one) | Live DOM — requires `root`; clones off-screen by default |
| Fragment export (`root` / `rootSelector`) | `rootSelector` + optional live `page` | `root` (live `Element`) |
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

**Behavior by `styleSource` and `rasterizeInPlace`:**

| Path | Playwright needed? | What happens |
|------|--------------------|--------------|
| `"inline"` (default), no rasterize | No | Parses `style=""` attributes only. Fast (~15–30 ms typical). |
| `"computed"` | **Yes** | Renders the fragment in headless Chromium, snapshots `getComputedStyle` for every element, then converts. |
| `rasterizeInPlace` (any `styleSource`) | **Yes** | Loads or reuses a Chromium page, rasterizes `<canvas>` / complex `<svg>` to PNG `<img>`, then converts. Can be combined with `"computed"`. |

Install Playwright once for computed or rasterize paths: `npm install playwright && npx playwright install chromium`.

When Playwright is needed and neither `page` nor `browser` is provided, the function launches Chromium, converts, and **closes the browser** in a `finally` block. Pass `browser` or `page` to avoid the per-call launch cost (see [Usage patterns](#usage-patterns)).

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
  root?: Element;                       // export root for computed + rasterize (SPA pattern)
  rasterizeInPlace?: boolean | RasterizeInPlaceOptions;
  imageResolver?: ImageResolver;
}

interface RasterizeInPlaceOptions {
  /** Mutate the live `root` instead of cloning. Default false (browser / live page). */
  mutate?: boolean;
  /** Extra selectors to rasterize, e.g. [".highcharts-container"]. */
  selectors?: string[];
  /** Supersampling factor (default `1`, max `4`). Renders PNG at scale× density; display size unchanged. Recommended: `2` for charts. */
  scale?: number;
}
```

| `styleSource` | Live DOM required? | Behavior |
|---------------|--------------------|----------|
| `"inline"` (default) | No* | Parses `style=""` only — works on a string fragment with no rendered page. |
| `"computed"` | **Yes** | Batch-reads native `getComputedStyle` from `options.root` or `document.body`. The page must already render the same fragment. |

\*Unless `rasterizeInPlace` is set — that always requires a live DOM with rendered charts.

| Export | Returns | Notes |
|--------|---------|-------|
| `convertHtmlToDocx(html, options?)` | `Promise<Blob>` | Primary browser API |
| `convertHtmlToDocxUint8Array(html, options?)` | `Promise<Uint8Array>` | Same bytes, no Blob wrapper |
| `buildDocxBlob` / `buildDocxUint8Array` | | Lower-level, bring your own `StyleResolver` |
| `snapshotComputedStylesFromDocument(doc?, root?)` | `ComputedStyleSnapshot[]` | Style snapshots scoped to `root` or `document.body` |
| `rasterizeInPlace(root, options?)` | `Promise<void>` | Mutates live DOM; prefer `convertHtmlToDocx({ rasterizeInPlace: { scale: 2 } })` for chart exports |
| `isSimpleSvgElement(svg)` | `boolean` | True when SVG is simple enough for native rect/text conversion |

All [document options](#options-convertoptions) (`pageSize`, `margins`, `metadata`, …) work here too. `browser`, `page`, and `rootSelector` are Node-only.

---

## Options (`ConvertOptions`)

```typescript
interface ConvertOptions extends DocumentConfig {
  styleSource?: "inline" | "computed";  // default "inline"
  browser?: Browser;                    // Node — reuse Playwright browser
  page?: Page;                          // Node — reuse Playwright page (computed / rasterize)
  rootSelector?: string;                // Node — export root when converting element.innerHTML from a live page
  rasterizeInPlace?: boolean | RasterizeInPlaceOptions;
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
  coverHtml?: string;    // cover page: page 1, before the TOC, then a page break
  tableOfContents?: boolean | {
    title?: string;         // bold heading above the TOC, e.g. "Contents"
    headingRange?: string;  // Word heading levels to include, default "1-3"
    hyperlink?: boolean;    // clickable entries, default true
    pageBreakAfter?: boolean; // start the body on a fresh page after the TOC
  };
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
| `coverHtml` | `string` | — | HTML fragment rendered as a **cover page**: the first content in the document, before the `tableOfContents` (if any), followed by an automatic page break so the TOC/body start on the next page. Uses the inline style path like `headerHtml`/`footerHtml` (inline `style=""` + `data:` images such as a logo). When a header/footer/`pageNumber` is set, it is suppressed on the cover page (Word "different first page"). Headings inside the cover are **not** included as TOC entries. |
| `tableOfContents` | `boolean \| { title?, headingRange?, hyperlink?, pageBreakAfter? }` | — | Inserts a **clickable, page-number-less** Table of Contents at the top, built from the document's `h1`–`h6` (mapped to Word Heading 1–6). Each entry is a hyperlink to a bookmark on its heading. Because there are no page numbers (which would require layout), the entry list is **complete at creation** — it renders correctly in every viewer (Word, LibreOffice, Google Docs, PDF/preview) with no field update and no "update fields" prompt. `headingRange` defaults to `"1-3"`; `hyperlink` (the clickable links) to `true` — set it `false` for a plain-text list. |
| `imageResolver` | `ImageResolver` | — | Resolve non-`data:` `<img src>`. See [Images](#images--the-resolver-hook). |
| `browser` | Playwright `Browser` | — | **Node only.** Reuse an already-launched browser across many conversions. |
| `page` | Playwright `Page` | — | **Node only.** Snapshot styles and/or rasterize from a page you already rendered. |
| `rootSelector` | `string` | — | **Node only.** CSS selector for the export root when passing `element.innerHTML` from a live `page`. Must match the node whose HTML you convert. |
| `rasterizeInPlace` | `boolean \| RasterizeInPlaceOptions` | — | Rasterize `<canvas>` and complex `<svg>` (e.g. Highcharts) to PNG `<img>` before conversion. Uses Playwright on Node. See [Charts & rasterizeInPlace](#charts--rasterizeinplace). **Recommended for charts:** `{ scale: 2 }`. |

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

**Resolution order for Playwright paths (Node):**

1. If `options.page` is set → use that page (no `setContent`)
2. Else if `options.browser` is set → new page per call, `setContent(html)`, browser kept open
3. Else → launch Chromium, `setContent(html)`, convert, close browser

When `rasterizeInPlace` is set, rasterization runs on the Playwright page **before** style resolution and conversion. With a live `page` + `rootSelector`, the export root is cloned off-screen by default (same as browser `root`). Ephemeral spawn pages (no `page`, no `rootSelector`) mutate in place by default.

When `styleSource: "computed"` and you pass `rootSelector`, computed-style snapshots are scoped to that element so paths match `element.innerHTML` (SPA fragment export).

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

### SPA fragment export (browser)

When exporting a subtree from a live Vue/React app, pass the **same element** as both the HTML source and the computed-style root:

```typescript
import { convertHtmlToDocx } from "dom-docx/browser";

async function exportSection(root: HTMLElement) {
  return convertHtmlToDocx(root.innerHTML, {
    styleSource: "computed",
    root, // paths in the style snapshot match the fragment cheerio tree
  });
}
```

Without `root`, computed styles are snapshotted from `document.body` while cheerio parses only the fragment — paths miss and styles fall back to inline attributes only.

### SPA fragment export (Node + live Playwright page)

Same pattern with a CSS selector on an already-rendered page:

```typescript
const rootSelector = ".page-body";
const html = await page.$eval(rootSelector, (el) => el.innerHTML);

const docx = await convertHtmlToDocx(html, {
  styleSource: "computed",
  page,
  rootSelector,
});
```

### Charts & `rasterizeInPlace`

dom-docx converts simple inline SVG (rect + text bars) natively. Chart libraries (Highcharts, Chart.js canvas, complex SVG paths/gradients) need rasterization to PNG `<img>` first.

#### `RasterizeInPlaceOptions`

| Option | Default | Description |
|--------|---------|-------------|
| `mutate` | `false` on live `root` / `page`; `true` on ephemeral spawn pages | Replace charts in the caller's DOM instead of cloning off-screen. |
| `selectors` | — | Extra CSS selectors to rasterize (e.g. `[".highcharts-container"]`). |
| `scale` | `1` (max `4`) | **Supersampling factor.** Renders the PNG at `scale`× pixel density but keeps `<img width>` / `<img height>` at layout size. |

**Recommended for chart exports:** `scale: 2`. At `scale: 1`, chart PNGs match layout pixels 1:1 and often look soft or jagged in Word. `scale: 2` produces sharper images at the cost of larger embedded PNGs (typically ~2–4× file size per chart). Use `scale: 1` when file size matters more than chart crispness.

**Browser** — requires `root`; clones off-screen by default:

```typescript
const blob = await convertHtmlToDocx(root.innerHTML, {
  styleSource: "computed",
  root,
  rasterizeInPlace: { scale: 2 },
  // or: { selectors: [".highcharts-container"], scale: 2 },
});
```

**Node** — uses the same Playwright page as computed styles:

```typescript
// Ephemeral page (HTML string with rendered chart markup)
const docx = await convertHtmlToDocx(html, {
  styleSource: "computed",
  rasterizeInPlace: { scale: 2 }, // mutates the throwaway page in place
});

// Live Playwright page (e.g. E2E export from a running app)
const docx2 = await convertHtmlToDocx(fragmentHtml, {
  styleSource: "computed",
  page,
  rootSelector: "#dashboard",
  rasterizeInPlace: { scale: 2 }, // clones off-screen unless { mutate: true }
});
```

`rasterizeInPlace: true` is shorthand for default options (`scale: 1`).

| Context | Default `mutate` | Why |
|---------|------------------|-----|
| Browser / live `page` + `rootSelector` | `false` (clone) | Avoid disturbing the user's SPA |
| Node spawn (no `page`, no `rootSelector`) | `true` | Ephemeral Chromium tab — safe to mutate |

Pass `{ mutate: true }` to replace charts in the caller's live DOM (browser or Playwright page).

Lower-level Node helpers: `preparePlaywrightRasterizedExport(page, rootSelector?, options?, mutate?)`, `openPlaywrightPage(html, browser)`.

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
| `break-before`, `page-break-before` | `page`, `always`, `left`, `right` → page break before block |
| `break-after`, `page-break-after` | same values → page break before **next** block sibling |

**Page break examples** (inline styles or computed stylesheet):

Use inline `style=""` or a stylesheet with `styleSource: "computed"`. Example:

```html
<p style="break-after: page">End of section</p>
<h2 style="break-before: page">Next section</h2>
<div style="break-after: page"></div> <!-- break before next block -->
```

Verified by `npm run guard:page-break` (OOXML + multi-page PDF; not part of the single-page visual suite).

All other CSS properties are silently ignored.

### What converts well / poorly

See [AGENTS.md](./AGENTS.md) for the full tier list. In short:

- **Excellent:** headings, paragraphs, lists, simple tables, inline formatting, short span highlights, blockquotes, `<hr>`.
- **Good:** shaded div banners, flex rows (≤4 items), table row/cell backgrounds, bordered boxes, `data:` images.
- **Avoid:** complex SVG, CSS grid/float/absolute layout, external stylesheets (inline path), forms, deep layout div nesting. Chart libraries may work with [`rasterizeInPlace`](#charts--rasterizeinplace) on a live page.

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
| `computedStyleResolverFromPage(page, rootSelector?)` | Node: snapshot an existing Playwright page, optionally scoped to `rootSelector`. |
| `snapshotComputedStyles(page, rootSelector?)` | Node: raw `{ path, styles }[]` snapshots from a Playwright page. |
| `preparePlaywrightRasterizedExport(page, rootSelector?, options?, mutate?)` | Node: rasterize charts on a Playwright page; returns `{ html, snapshotRootSelector, cleanupSelector }`. |
| `openPlaywrightPage(html, browser)` | Node: `setContent` wrapper with harness viewport defaults. |
| `snapshotComputedStylesFromDocument(doc?, root?)` | Browser (`dom-docx/browser`): same snapshots from a live `document`, scoped to `root` or `document.body`. |
| `rasterizeInPlace(root, options?)` | Browser: mutate live DOM to replace charts with PNG `<img>`. |
| `isSimpleSvgElement(svg)` | Browser: true when SVG converts natively without rasterization. |

**Computed style properties captured** (via `getComputedStyle`): `color`, `backgroundColor`, `display`, `flexDirection`, `gap` / `columnGap` / `rowGap`, `textAlign`, `fontSize`, `fontWeight`, `fontStyle`, margin and padding sides, per-side border width/color. UA defaults on headings are partially stripped when the element has no inline override, so the computed path stays aligned with the inline path for bare `<h1>`–`<h6>`.

---

## How the engine works

Conversion is a three-stage pipeline: **style resolution → HTML visitor → OOXML pack + patch**.

```
HTML fragment
     │
     ▼
┌────────────────────────────────────────────────────────────────┐
│ 0. Optional rasterize (rasterizeInPlace)                       │
│    canvas / complex SVG → PNG <img> in browser or Playwright   │
└────────────────────────────────────────────────────────────────┘
     │
     ▼
┌────────────────────────────────────────────────────────────────┐
│ 1. Style resolution (StyleResolver)                            │
│    inline: parse style="" on each element                      │
│    computed (Node): getComputedStyle via Playwright/Chromium   │
│    computed (browser): getComputedStyle on live root / body    │
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
| `svg` (low-complexity) | Native DOCX blocks — `<rect>` bands + `<text>` (bar/funnel charts); complex SVG rasterized when `rasterizeInPlace` is set |
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
| **playwright** | For `styleSource: "computed"` and/or `rasterizeInPlace` (optional peer dependency, installed separately) | **Not used** |
| **Live DOM** | Not required for inline path; computed/rasterize use Playwright | Required for computed and `rasterizeInPlace` |

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
- **Images** — png/jpg/gif/bmp only; svg `<img>` sources are not rasterized. Inline low-complexity `<svg>` elements convert natively; complex chart SVG and `<canvas>` can be rasterized with `rasterizeInPlace`.
- **No JavaScript execution on the inline/spawn path** — static HTML is converted as-is. Client-rendered charts must already be in the DOM (browser computed path, live Playwright `page`, or pre-rendered markup). The spawn path does not run your app's JS bundle unless you load a full page URL in Playwright yourself.

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
