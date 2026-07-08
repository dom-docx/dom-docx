# dom-docx

Convert semantic **HTML fragments** to native, editable **Word documents** (OOXML): paragraphs, runs, lists, tables, images. Not screenshots or layout hacks.

**Live demo:** [dom-docx.com](https://dom-docx.com/). Try the converter, browse showcases, read the learn guide.

Built with a visual regression loop: render HTML in Chromium, convert to docx, rasterize via LibreOffice, score layout + structural fidelity against a human-validated metric, iterate. Latest scores: [TEST-SCORES.md](./docs/TEST-SCORES.md) · methodology: [SCORING.md](./docs/SCORING.md).

## Install

```bash
npm install dom-docx
```

Requires **Node.js ≥ 20**. No browser or Playwright is needed for the default **`inline`** path.

### When is Playwright needed?

| Entry                            | `styleSource: "inline"` | `styleSource: "computed"` | `rasterizeInPlace` |
| -------------------------------- | ----------------------- | ------------------------- | ------------------ |
| **Node** (`dom-docx`)            | Pure JS, no browser     | **Playwright + Chromium** | **Playwright + Chromium** (same headless page) |
| **Browser** (`dom-docx/browser`) | Pure JS, no live DOM    | **Live page**: native `getComputedStyle` | **Live page**: canvas/SVG → PNG `<img>` in the tab |

On Node, `playwright` is an **optional peer dependency**. `npm install dom-docx` pulls only `docx`, `cheerio` and `fflate`, nothing heavy. It is loaded lazily when you pass `styleSource: "computed"` or `rasterizeInPlace`. To use those paths, install Playwright and Chromium yourself, once:

```bash
npm install playwright
npx playwright install chromium
```

Playwright is also used by the **dev test harness** (not required to use the library). Contributors: `npm run setup` after clone.

LibreOffice is **not** needed to convert. It is only used for the visual test harness.

## CLI

Convert a file without writing any code:

```bash
npx dom-docx input.html -o output.docx
npx dom-docx input.html                      # writes input.docx next to it
cat fragment.html | npx dom-docx - -o -      # stdin → binary stdout (pipelines)
npx dom-docx input.html -s computed          # stylesheet/class HTML (needs playwright installed)
```

Input is a **body HTML fragment**, same as the API. `--help` for all options.

## Quick start

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
// e.g. trigger a download in the browser
const a = document.createElement("a");
a.href = URL.createObjectURL(blob);
a.download = "output.docx";
a.click();
```

No Playwright, no Node. This runs entirely in the user's tab. See [Browser bundle](#browser-bundle) below.

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

Pass a **body fragment only** (no `<!DOCTYPE>` / `<html>` / `<body>` required). Defaults: US Letter, 1″ margins, Arial 10.5pt body text.

## v0.1.x capability

**Supported (default `styleSource: "inline"`):**

- Headings, paragraphs, lists (`<ul>`/`<ol>` including `list-style-type`), tables, links, inline formatting
- Block backgrounds, blockquotes, `<hr>`, simple flex rows (≤4 items)
- `data:` images; remote images via your `imageResolver`
- Page size/orientation/margins, default font, metadata, header/footer HTML, page numbers, lang/direction
- Low-complexity inline SVG (bars + text)
- CSS bar divs in table cells (background + height/width → native shaded bands)

**Advanced (optional `styleSource: "computed"`):**

- Resolves `<style>` blocks and class/`#id` selectors via `getComputedStyle`
- **Node:** requires **`playwright`** (optional peer dependency, installed separately) + Chromium. The library launches headless Chromium to render the fragment
- **Browser bundle:** uses the **live DOM** in the user's tab. No Playwright, no extra install
- **SPA fragment export:** pass `root` (browser) or `rootSelector` (Node + live `page`) when converting `element.innerHTML` so computed-style paths match the fragment tree
- **Inline is the supported default** for npm installs; computed is for stylesheets/classes or when you already have a rendered page

**Charts & complex SVG (optional `rasterizeInPlace`):**

- Rasterizes `<canvas>` and complex `<svg>` (e.g. Highcharts) to PNG `<img>` before conversion
- **Browser:** requires `root`. Clones off-screen by default so the live page is not mutated
- **Node:** uses the same Playwright/Chromium context as computed styles; ephemeral spawn pages mutate in place by default
- Simple inline SVG (rect + text bar charts) still converts natively without rasterization

**Not supported in v0.1.x:**

- External stylesheets on the inline path (use computed or inline all styles)
- Web fonts, CSS grid/float layout, forms, `<pre>` polish, `<dl>`, table `rowspan`
- Header/footer first/even page variants; guaranteed multi-page layout fidelity
- Complex SVG (paths, gradients, `<use>`) unless `rasterizeInPlace` is used on a live rendered page

See [AGENTS.md](./AGENTS.md) for HTML authoring tiers and [API.md](./API.md) for full options.

## API

### `convertHtmlToDocx(html, options?)`

Returns `Promise<Buffer>` (Node) with a valid `.docx` file.

```typescript
import { convertHtmlToDocx, type ConvertOptions } from "dom-docx";

const docx = await convertHtmlToDocx(html, {
  pageSize: "a4",
  orientation: "landscape",
  margins: { top: 0.75, bottom: 0.75 }, // inches; omitted sides default to 1
  defaultFont: { family: "Georgia", sizePt: 11 },
  metadata: { title: "Q3 Report", creator: "Finance" },
  headerHtml: "<p style='font-size:12px;color:#666'>Confidential</p>",
  footerHtml: "<p style='font-size:12px'>© 2026 ACME</p>",
  pageNumber: true,
  lang: "en-US",
  direction: "ltr",
});
```

### Options

| Option                      | Default       | Description                                                                                                                                                                                                      |
| --------------------------- | ------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `styleSource`               | `"inline"`    | `"inline"` parses `style=""` only (pure JS, fast). `"computed"` uses `getComputedStyle`. On **Node** this requires Playwright + Chromium; in the **browser bundle** it reads from the live DOM (no Playwright). |
| `browser` / `page`          | —             | **Node only.** Reuse an open Playwright browser or page (computed styles and/or `rasterizeInPlace`). Not used by `dom-docx/browser`. |
| `rootSelector`              | —             | **Node only.** CSS selector for the export root when converting `element.innerHTML` from a live Playwright `page`. Must match the node whose HTML you pass. |
| `rasterizeInPlace`          | —             | Rasterize `<canvas>` / chart `<svg>` to PNG `<img>` before conversion. **Node:** Playwright page; **browser:** requires `root`. See [API.md](./API.md#charts--rasterizeinplace). |
| `imageResolver`             | —             | Hook to fetch non-`data:` `<img src>` (library never fetches on its own).                                                                                                                                        |
| `pageSize`                  | `"letter"`    | `"letter"`, `"a4"` or `{ width, height }` in inches.                                                                                                                                                            |
| `orientation`               | `"portrait"`  | `"landscape"` swaps dimensions.                                                                                                                                                                                  |
| `margins`                   | `1` inch each | Per-side overrides in inches.                                                                                                                                                                                    |
| `defaultFont`               | Arial 10.5pt  | `{ family, sizePt }` for body text without explicit CSS.                                                                                                                                                         |
| `metadata`                  | —             | `title`, `subject`, `creator`, `keywords[]`, `description` → `docProps/core.xml`.                                                                                                                                |
| `headerHtml` / `footerHtml` | —             | HTML fragments for page header/footer.                                                                                                                                                                           |
| `pageNumber`                | `false`       | Appends centered `Page N` field to footer.                                                                                                                                                                       |
| `lang` / `direction`        | —             | Spell-check locale; `"rtl"` for right-to-left.                                                                                                                                                                   |

### Images

Only **`data:`** URLs embed automatically. For `http(s):` or file paths, supply a resolver. You control fetch policy and security:

```typescript
const docx = await convertHtmlToDocx(html, {
  imageResolver: async (src) => {
    const res = await fetch(src); // your allowlist / SSRF checks
    if (!res.ok) return null;
    return { data: new Uint8Array(await res.arrayBuffer()), type: "png" };
  },
});
```

### Browser bundle

For client-side conversion (returns a `Blob`). **No Playwright.** The bundle runs entirely in the user's browser.

```typescript
import { convertHtmlToDocx } from "dom-docx/browser";

// Inline styles: pass HTML string directly
const blob = await convertHtmlToDocx(htmlFragment, { styleSource: "inline" });

// Computed styles: render the fragment in the live DOM first, then convert
document.body.innerHTML = htmlFragment;
const blob2 = await convertHtmlToDocx(htmlFragment, { styleSource: "computed" });

// SPA export: pass the live root so computed paths match element.innerHTML
const root = document.querySelector(".page-body")!;
const blob3 = await convertHtmlToDocx(root.innerHTML, {
  styleSource: "computed",
  root,
});

// Charts (Highcharts, canvas): rasterize to PNG <img> on a clone, then convert
const blob4 = await convertHtmlToDocx(root.innerHTML, {
  styleSource: "computed",
  root,
  rasterizeInPlace: true, // or { selectors: [".highcharts-container"] }
});
```

Browser-only options: `root`, `document`, `rasterizeInPlace`. Build: `npm run build:browser` → `dist/browser/dom-docx.browser.js`.

For advanced usage (`buildDocxBuffer`, custom `StyleResolver`, engine architecture), see [API.md](./API.md).

---

## What converts well

Optimized for **Word-friendly semantic HTML**: headings, paragraphs, lists, data tables, inline formatting, shaded callouts, simple flex rows.

| Excellent                      | Good                            | Avoid (or use `rasterizeInPlace`)  |
| ------------------------------ | ------------------------------- | ---------------------------------- |
| Headings, lists, simple tables | Shaded banners, flex (≤4 items) | Live chart libraries (Highcharts, canvas) unless rasterized |
| Inline `strong` / `em` / links | Table row/cell backgrounds      | Complex SVG paths/gradients unless rasterized |
| Short span highlights          | Blockquotes, `<hr>`             | External stylesheets (inline path) |
| Simple SVG bars (rect + text)  | `data:` images                  | Forms, web fonts, CSS grid/float layout |

Full authoring guide for agents: [AGENTS.md](./AGENTS.md).

---

## How it was built

dom-docx maps a practical HTML subset to native OOXML through a three-stage pipeline:

1. **Style resolution**: inline `style=""` (default) or browser computed styles
2. **Visitor**: Cheerio walk → `docx` paragraphs, tables, numbering, hyperlinks
3. **Pack + patch**: generate OOXML, patch list numbering and shaded-block alignment for LibreOffice/Word PDF export

Quality is driven by an autonomous loop rather than one-off visual checks:

- **30+ regression cases** (defined in [`tools/generator.ts`](./tools/generator.ts), run via `npm run score:suite`): human-validated **layout fidelity** (ink-projection structure comparison, 85.6% concordance with blind human quality ratings), plus guards for bad contrast, missing list markers, wrong text and imbalanced shaded blocks; raw pixel match is recorded as a regression tripwire
- **Engine score**: 50% visual (layout-based) + 35% editability (native structure, not 1×1 layout tables) + 15% compile speed
- **OSS benchmark**: same harness scores [html-to-docx](https://www.npmjs.com/package/html-to-docx) and [@turbodocx/html-to-docx](https://www.npmjs.com/package/@turbodocx/html-to-docx) for ongoing comparison ([BENCHMARK.md](./docs/BENCHMARK.md))

The default **`inline`** path is pure JavaScript (`docx` + `cheerio` + `fflate`) with no browser required. **Playwright is Node-only**: for `styleSource: "computed"` on the server and for the dev harness. The **`dom-docx/browser`** bundle never uses Playwright; computed styles come from the live page's `getComputedStyle`.

Full scoring formulas, subscores, calibration and the agent iteration workflow: [SCORING.md](./docs/SCORING.md).

---

## Development

For contributors and harness runs (not required to use the library):

```bash
git clone … && npm install && npm run setup
npm run build              # dist/ for npm pack
npm run typecheck
npm run score:suite          # full visual + XML regression suite (cases: tools/generator.ts)
npm run score:suite:priority # fast subset of the same cases
# Run one case: SUITE_ONLY=rasterize-in-place-chart npm run score:suite
npm run score:benchmark     # vs html-to-docx + TurboDocx
npm run guard:config        # ConvertOptions OOXML checks
```

Prerequisites for the harness: **LibreOffice** (`soffice`) for PDF rasterization, **Playwright Chromium** (`npm run setup`).

---

## Documentation

| Doc                                     | Contents                                                |
| --------------------------------------- | ------------------------------------------------------- |
| [API.md](./API.md)                      | Full API reference, engine architecture, usage patterns |
| [AGENTS.md](./AGENTS.md)                | HTML authoring guide for AI agents                      |
| [SCORING.md](./docs/SCORING.md)         | Validation methodology and engine score                 |
| [TEST-SCORES.md](./docs/TEST-SCORES.md) | Latest suite metrics and per-case scores                |
| [BENCHMARK.md](./docs/BENCHMARK.md)     | Comparison vs OSS html-to-docx libraries                |
| [examples/](./examples/)                | Sample HTML, DOCX output and side-by-side previews     |
| [SHOWCASE.md](./docs/SHOWCASE.md)       | How to run and extend showcase examples                 |
| [CONTRIBUTING.md](./CONTRIBUTING.md)    | Library vs harness layout, dev setup                    |

## License

[MIT](./LICENSE)
