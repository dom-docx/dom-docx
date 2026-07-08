import { TEST_IMAGE_260x140, TEST_IMAGE_H, TEST_IMAGE_W } from "./test-image.js";
import type { ConvertOptions } from "../src/converter.js";

export interface TestCase {
  name: string;
  html: string;
  /** One-line human description used to render docs tables (see tools/docs-sync.ts). */
  description?: string;
  /** Custom conversion path (e.g. `rasterizeInPlace`). Omitted cases use default inline. */
  convertOptions?: ConvertOptions;
}

/** Cases that skip default inline / computed parity guards. */
export function isCustomConvertCase(testCase: TestCase): boolean {
  return testCase.convertOptions != null;
}

export function resolveHarnessConvertOptions(
  testCase: TestCase,
  browser: import("playwright").Browser,
): ConvertOptions | undefined {
  const opts = testCase.convertOptions;
  if (!opts) return undefined;
  const needsBrowser = opts.rasterizeInPlace || opts.styleSource === "computed";
  return needsBrowser ? { ...opts, browser: opts.browser ?? browser } : opts;
}

/** Plain, everyday HTML — the validation baseline. */
const STANDARD_TEST_CASES: TestCase[] = [
  {
    name: "plain-paragraph",
    description: "Single unstyled `<p>`",
    html: `<p>This is a single plain paragraph with no formatting.</p>`,
  },
  {
    name: "multiple-paragraphs",
    description: "Three sequential paragraphs",
    html: `
      <p>First paragraph of a short document.</p>
      <p>Second paragraph follows normally.</p>
      <p>Third paragraph closes the section.</p>
    `,
  },
  {
    name: "heading-hierarchy",
    description: "h1 / h2 / h3 with body text",
    html: `
      <h1>Document Title</h1>
      <h2>Section One</h2>
      <p>Introductory text under section one.</p>
      <h2>Section Two</h2>
      <p>Introductory text under section two.</p>
      <h3>Subsection</h3>
      <p>Detail text under the subsection.</p>
    `,
  },
  {
    name: "simple-unordered-list",
    description: "Basic `<ul>` with 3 items",
    html: `
      <ul>
        <li>Apples</li>
        <li>Bananas</li>
        <li>Cherries</li>
      </ul>
    `,
  },
  {
    name: "simple-ordered-list",
    description: "Basic `<ol>` with 3 items",
    html: `
      <ol>
        <li>Preheat the oven</li>
        <li>Mix the ingredients</li>
        <li>Bake for 30 minutes</li>
      </ol>
    `,
  },
  {
    name: "ordered-list-rich-inline",
    description: "`<ol>` with `<strong>` + highlighted `<span>` per item",
    html: `
      <h2 style="font-size:15px">Rep leaderboard (top 3)</h2>
      <ol>
        <li><strong>A. Chen</strong> — $840k (<span style="background:#cfc">128% quota</span>)</li>
        <li><strong>M. Ortiz</strong> — $790k (<span style="background:#cfc">119% quota</span>)</li>
        <li><strong>S. Patel</strong> — $755k (<span style="background:#ff0">104% quota</span>)</li>
      </ol>
    `,
  },
  {
    name: "paragraph-and-list",
    description: "Intro paragraph + `<ul>`",
    html: `
      <p>Shopping list for the week:</p>
      <ul>
        <li>Milk</li>
        <li>Bread</li>
        <li>Eggs</li>
      </ul>
    `,
  },
  {
    name: "simple-link",
    description: "One hyperlinked anchor",
    html: `<p>Visit <a href="https://example.com">Example Domain</a> for more info.</p>`,
  },
  {
    name: "multiple-links",
    description: "Two links in one sentence",
    html: `
      <p>
        See <a href="https://example.com/a">link A</a> and
        <a href="https://example.com/b">link B</a> in one sentence.
      </p>
    `,
  },
  {
    name: "basic-inline-formatting",
    description: "`<strong>`, `<em>`, nested bold-italic",
    html: `
      <p>
        This sentence has <strong>bold</strong>, <em>italic</em>, and
        <strong><em>bold italic</em></strong> text.
      </p>
    `,
  },
  {
    name: "pre-code-block",
    description: "Fenced `<pre><code>` + inline `<code>`",
    html: `
      <p>Install and convert:</p>
      <pre style="background:#f5f5f5;padding:12px 14px;border:1px solid #ddd;font-size:13px;line-height:1.45;white-space:pre"><code>npm install dom-docx
const { convertHtmlToDocx } = await import("dom-docx");
const docx = await convertHtmlToDocx(html);</code></pre>
      <p>Save the buffer with <code>writeFile()</code>.</p>
    `,
  },
  {
    name: "simple-table-2x2",
    description: "2-column table, header + one row",
    html: `
      <table border="1" cellpadding="4" style="border-collapse:collapse;width:100%">
        <tr><td>Name</td><td>Value</td></tr>
        <tr><td>Alpha</td><td>100</td></tr>
      </table>
    `,
  },
  {
    name: "simple-table-3col",
    description: "3-column table, 3 rows",
    html: `
      <table border="1" cellpadding="4" style="border-collapse:collapse;width:100%">
        <tr><td>Item</td><td>Qty</td><td>Price</td></tr>
        <tr><td>Widget</td><td>2</td><td>$9.99</td></tr>
        <tr><td>Gadget</td><td>1</td><td>$14.50</td></tr>
      </table>
    `,
  },
  {
    name: "paragraph-with-line-break",
    description: "Address block with `<br>` tags",
    html: `
      <p>
        Line one of the address.<br>
        Line two of the address.<br>
        Line three of the address.
      </p>
    `,
  },
  {
    name: "simple-blockquote",
    description: "Plain blockquote + paragraph",
    html: `
      <blockquote>
        <p>Simplicity is the ultimate sophistication.</p>
      </blockquote>
    `,
  },
  {
    name: "centered-paragraph",
    description: "`text-align: center`",
    html: `<p style="text-align:center">This paragraph is centered.</p>`,
  },
  {
    name: "horizontal-rule",
    description: "Content separated by `<hr>`",
    html: `
      <p>Content above the rule.</p>
      <hr>
      <p>Content below the rule.</p>
    `,
  },
];

/** Stressful HTML — colspan drift, nested structures, asymmetric spacing. */
const EDGE_TEST_CASES: TestCase[] = [
  {
    name: "typography-colors",
    description: "Foreground/background colors, mixed inline & block",
    html: `
      <h1 style="color:#1a1a2e;background:#eaeaea;padding:8px">Heading Alpha</h1>
      <p style="color:#e63946;font-size:18px">Red foreground text with <strong>bold</strong> and <em>italic</em>.</p>
      <p style="background:#457b9d;color:#f1faee;padding:12px">Light text on blue background block.</p>
      <span style="color:#2a9d8f">Inline green</span>
      <span style="color:#e76f51"> and orange siblings.</span>
    `,
  },
  {
    name: "table-mismatched-cells",
    description: "Colspan, short rows, extra cells",
    html: `
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%">
        <tr><td>A1</td><td>A2</td><td>A3</td></tr>
        <tr><td>B1</td><td colspan="2">B2 spans two</td></tr>
        <tr><td>C1</td></tr>
        <tr><td>D1</td><td>D2</td><td>D3</td><td>D4 extra cell</td></tr>
      </table>
    `,
  },
  {
    name: "borderless-table",
    description: "Label/value table with `border:none`",
    html: `
      <h2 style="font-size:16px">Connection details</h2>
      <table border="0" cellpadding="8" style="border-collapse:collapse;width:100%;border:none">
        <tr>
          <td style="width:130px;color:#666"><strong>Host</strong></td>
          <td>api.example.com</td>
        </tr>
        <tr>
          <td style="color:#666"><strong>Port</strong></td>
          <td>443</td>
        </tr>
        <tr>
          <td style="color:#666"><strong>Protocol</strong></td>
          <td>HTTPS / REST</td>
        </tr>
        <tr>
          <td style="color:#666"><strong>Auth</strong></td>
          <td>Bearer token</td>
        </tr>
      </table>
    `,
  },
  {
    name: "table-row-backgrounds",
    description: "Shaded `<tr>` bands",
    html: `
      <table border="1" cellpadding="8" style="border-collapse:collapse;width:100%">
        <tr style="background:#1a1a2e;color:#f1faee">
          <td><strong>Item</strong></td>
          <td style="text-align:right"><strong>Value</strong></td>
        </tr>
        <tr><td>Alpha</td><td style="text-align:right">100</td></tr>
        <tr style="background:#f5f5f5"><td><strong>Subtotal</strong></td><td style="text-align:right"><strong>100</strong></td></tr>
        <tr><td>Beta</td><td style="text-align:right;color:#2a9d8f">+12%</td></tr>
      </table>
    `,
  },
  {
    name: "nested-blockquotes-lists",
    description: "Nested quotes, `<ol>` inside `<ul>`",
    html: `
      <blockquote style="border-left:4px solid #333;padding-left:12px;margin:8px 0">
        Outer quote
        <blockquote style="border-left:4px solid #666;padding-left:12px">
          Nested quote level 2
          <ol>
            <li>Ordered one
              <ul>
                <li>Unordered inside ordered A</li>
                <li>Unordered inside ordered B</li>
              </ul>
            </li>
            <li>Ordered two</li>
          </ol>
        </blockquote>
      </blockquote>
    `,
  },
  {
    name: "inline-vs-block",
    description: "Spans, links, code, styled divs",
    html: `
      <p>
        <span style="background:#ff0">Inline span</span>
        <a href="#">link</a>
        <code>code()</code>
        stretching across the line with more inline content here.
      </p>
      <div style="background:#ddd;padding:8px;margin:8px 0">Block div one</div>
      <p>Paragraph between blocks.</p>
      <div style="background:#ccc;padding:8px">Block div two</div>
    `,
  },
  {
    name: "inline-backgrounds",
    description: "Multi-color inline highlights, bold in shaded span",
    html: `
      <p>
        <span style="background:#cfc">Green highlight</span>
        and
        <span style="background:#ccf">blue highlight</span>
        on one line.
      </p>
      <p>
        Mixed with <span style="background:#fcc"><strong>bold pink</strong></span>
        and plain text after.
      </p>
    `,
  },
  {
    name: "mixed-margins-paddings",
    description: "Asymmetric margin/padding, bordered box",
    html: `
      <div style="margin:40px 20px 10px 60px;padding:16px 32px 8px 12px;background:#f5f5f5">
        Box with asymmetric margin and padding.
      </div>
      <p style="margin-left:80px;padding-top:24px">Indented paragraph with top padding.</p>
      <div style="margin:0;padding:0">
        <div style="margin:12px;padding:4px;border:1px solid #999">Nested margin box</div>
      </div>
    `,
  },
  {
    name: "flex-row-horizontal",
    description: "`display:flex; flex-direction:row` — three columns with gap",
    html: `
      <div style="display:flex;flex-direction:row;gap:12px;padding:8px;background:#f0f0f0">
        <div style="background:#ccffcc;padding:8px;text-align:center">Alpha</div>
        <div style="background:#ccccff;padding:8px;text-align:center">Beta</div>
        <div style="background:#ffcccc;padding:8px;text-align:center">Gamma</div>
      </div>
    `,
  },
  {
    name: "flex-column-vertical",
    description: "`display:flex; flex-direction:column` — stacked rows with gap",
    html: `
      <div style="display:flex;flex-direction:column;gap:10px;padding:12px;background:#f5f5f5">
        <div style="background:#dddddd;padding:8px">First row</div>
        <div style="background:#bbbbbb;padding:8px">Second row</div>
        <div style="background:#888888;padding:8px;color:#ffffff">Third row</div>
      </div>
    `,
  },
  {
    name: "inline-svg-chart",
    description: "Inline SVG bar chart → native DOCX bands",
    html: `
      <p><strong>Activation funnel</strong></p>
      <figure style="margin:8px 0;text-align:center">
        <svg width="420" height="100" viewBox="0 0 420 100" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Activation funnel chart">
          <rect x="0" y="60" width="380" height="28" fill="#457b9d" rx="2"/>
          <rect x="40" y="40" width="300" height="28" fill="#2a9d8f" rx="2"/>
          <rect x="80" y="20" width="220" height="28" fill="#e9c46a" rx="2"/>
          <rect x="120" y="0" width="140" height="28" fill="#e76f51" rx="2"/>
          <text x="0" y="98" fill="#666" font-size="11" font-family="Arial">Signup → Trial → Qualified → Closed</text>
        </svg>
        <figcaption style="color:#666;font-size:12px">Fig 1. Illustrative activation funnel</figcaption>
      </figure>
    `,
  },
  {
    name: "rasterize-in-place-chart",
    description: "Complex SVG + canvas rasterized via `rasterizeInPlace` before conversion",
    html: `
      <p><strong>Monthly active users</strong></p>
      <figure style="margin:8px 0;text-align:center">
        <svg width="420" height="160" viewBox="0 0 420 160" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="MAU trend line chart">
          <defs>
            <linearGradient id="mauAreaFill" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stop-color="#457b9d" stop-opacity="0.35"/>
              <stop offset="100%" stop-color="#457b9d" stop-opacity="0.02"/>
            </linearGradient>
          </defs>
          <line x1="48" y1="16" x2="48" y2="128" stroke="#ccc" stroke-width="1"/>
          <line x1="48" y1="128" x2="404" y2="128" stroke="#ccc" stroke-width="1"/>
          <line x1="48" y1="56" x2="404" y2="56" stroke="#eee" stroke-width="1"/>
          <line x1="48" y1="92" x2="404" y2="92" stroke="#eee" stroke-width="1"/>
          <path d="M48,128 L48,104 L108,88 L168,96 L228,72 L288,80 L348,52 L404,44 L404,128 Z" fill="url(#mauAreaFill)"/>
          <polyline points="48,104 108,88 168,96 228,72 288,80 348,52 404,44" fill="none" stroke="#457b9d" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
          <circle cx="108" cy="88" r="4" fill="#457b9d"/>
          <circle cx="228" cy="72" r="4" fill="#457b9d"/>
          <circle cx="348" cy="52" r="4" fill="#457b9d"/>
          <circle cx="404" cy="44" r="4" fill="#457b9d"/>
          <text x="48" y="148" fill="#666" font-size="11" font-family="Arial">Jan</text>
          <text x="168" y="148" fill="#666" font-size="11" font-family="Arial">Mar</text>
          <text x="288" y="148" fill="#666" font-size="11" font-family="Arial">May</text>
          <text x="384" y="148" fill="#666" font-size="11" font-family="Arial">Jun</text>
        </svg>
        <figcaption style="color:#666;font-size:12px">Fig 1. Six-month MAU trend</figcaption>
      </figure>
      <p style="margin-top:12px"><strong>Channel mix</strong></p>
      <canvas id="channelMix" width="420" height="100" aria-label="Channel mix bar chart"></canvas>
      <script>
        (function () {
          var canvas = document.getElementById("channelMix");
          if (!canvas) return;
          var ctx = canvas.getContext("2d");
          if (!ctx) return;
          var bars = [
            { label: "Organic", value: 0.42, color: "#457b9d" },
            { label: "Paid", value: 0.28, color: "#2a9d8f" },
            { label: "Partner", value: 0.18, color: "#e9c46a" },
            { label: "Direct", value: 0.12, color: "#e76f51" },
          ];
          var x = 0;
          var w = canvas.width;
          var h = canvas.height - 18;
          bars.forEach(function (bar) {
            var bw = Math.round(bar.value * w);
            ctx.fillStyle = bar.color;
            ctx.fillRect(x, h * (1 - bar.value), bw, h * bar.value);
            ctx.fillStyle = "#666";
            ctx.font = "11px Arial";
            ctx.fillText(bar.label, x + 4, canvas.height - 4);
            x += bw;
          });
        })();
      </script>
    `,
    convertOptions: {
      rasterizeInPlace: true,
    },
  },
  {
    name: "table-cell-bar-divs",
    description: "CSS bar divs inside table cells",
    html: `
      <table border="1" cellpadding="8" style="border-collapse:collapse;width:100%">
        <tr style="background:#457b9d;color:#f1faee">
          <td><strong>Region</strong></td>
          <td style="text-align:right"><strong>Actual</strong></td>
          <td><strong>Trend</strong></td>
        </tr>
        <tr>
          <td>West</td>
          <td style="text-align:right">$2.41M</td>
          <td><div style="background:#457b9d;height:14px;width:92%;max-width:180px"></div></td>
        </tr>
        <tr>
          <td>East</td>
          <td style="text-align:right">$1.98M</td>
          <td><div style="background:#457b9d;height:14px;width:76%;max-width:180px"></div></td>
        </tr>
        <tr>
          <td>EMEA</td>
          <td style="text-align:right">$1.62M</td>
          <td><div style="background:#457b9d;height:14px;width:68%;max-width:180px"></div></td>
        </tr>
      </table>
    `,
  },
  {
    name: "unicode-emoji-content",
    description: "Emoji in body text",
    html: `
      <h2 style="font-size:15px">✅ What went well</h2>
      <table border="1" cellpadding="10" style="border-collapse:collapse;width:100%">
        <tr>
          <td style="text-align:center;background:#eaeaea"><span style="font-size:22px">📈</span><br><strong>12k</strong><br><span style="color:#666;font-size:12px">Activations</span></td>
          <td style="text-align:center;background:#eaeaea"><span style="font-size:22px">⚡</span><br><strong>68%</strong><br><span style="color:#666;font-size:12px">Retention</span></td>
        </tr>
      </table>
      <ul>
        <li>🥇 <strong>A. Chen</strong> — $840k</li>
        <li>🥈 <strong>M. Ortiz</strong> — $790k</li>
      </ul>
    `,
  },
  {
    name: "image-block",
    description: "`data:` URL `<img>` in a centered paragraph",
    html: `
      <p><strong>Quarterly revenue</strong></p>
      <p style="text-align:center"><img src="${TEST_IMAGE_260x140}" width="${TEST_IMAGE_W}" height="${TEST_IMAGE_H}" alt="Quarterly revenue bar chart"></p>
      <p style="color:#666;font-size:12px">Revenue grew across all three product lines.</p>
    `,
  },
  {
    name: "image-figure",
    description: "`<figure>` → `<img>` + `<figcaption>`",
    html: `
      <figure style="margin:8px 0;text-align:center">
        <img src="${TEST_IMAGE_260x140}" width="${TEST_IMAGE_W}" height="${TEST_IMAGE_H}" alt="Quarterly revenue bar chart">
        <figcaption style="color:#666;font-size:12px">Fig 1. Quarterly revenue by product line</figcaption>
      </figure>
    `,
  },
  {
    name: "ordered-list-lower-alpha",
    description: "`<ol list-style-type:lower-alpha>`",
    html: `
      <p>Steps:</p>
      <ol style="list-style-type:lower-alpha">
        <li>Assess the current state</li>
        <li>Design the target</li>
        <li>Execute the migration</li>
      </ol>
    `,
  },
  {
    name: "ordered-list-upper-roman",
    description: "`<ol list-style-type:upper-roman>`",
    html: `
      <p>Phases:</p>
      <ol style="list-style-type:upper-roman">
        <li>Discovery</li>
        <li>Delivery</li>
        <li>Review</li>
      </ol>
    `,
  },
  {
    name: "unordered-list-square",
    description: "`<ul list-style-type:square>`",
    html: `
      <p>Checklist:</p>
      <ul style="list-style-type:square">
        <li>Backups verified</li>
        <li>Access reviewed</li>
        <li>Runbook updated</li>
      </ul>
    `,
  },
];

export function generateTestCases(): TestCase[] {
  return [...STANDARD_TEST_CASES, ...EDGE_TEST_CASES];
}

export function generateStandardTestCases(): TestCase[] {
  return [...STANDARD_TEST_CASES];
}

export function generateEdgeTestCases(): TestCase[] {
  return [...EDGE_TEST_CASES];
}

/**
 * Fast regression subset (~10 cases) — one anchor per major pattern.
 * Full suite: `npm run score:suite` · subset: `npm run score:suite:priority`
 */
export const PRIORITY_LOOP_CASE_NAMES = [
  "plain-paragraph",
  "simple-unordered-list",
  "simple-ordered-list",
  "simple-table-2x2",
  "simple-link",
  "basic-inline-formatting",
  "paragraph-and-list",
  "table-row-backgrounds",
  "flex-row-horizontal",
  "nested-blockquotes-lists",
] as const;

export type LoopCaseMode = "full" | "priority";

export function generatePriorityTestCases(): TestCase[] {
  const byName = new Map(generateTestCases().map((c) => [c.name, c]));
  const missing = PRIORITY_LOOP_CASE_NAMES.filter((name) => !byName.has(name));
  if (missing.length > 0) {
    throw new Error(`Priority loop cases missing from generator: ${missing.join(", ")}`);
  }
  return PRIORITY_LOOP_CASE_NAMES.map((name) => byName.get(name)!);
}

export function resolveLoopTestCases(mode: LoopCaseMode): TestCase[] {
  return mode === "priority" ? generatePriorityTestCases() : generateTestCases();
}
