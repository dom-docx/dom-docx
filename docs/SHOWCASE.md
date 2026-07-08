# Showcase Examples

Nine real-world HTML documents designed to **convert cleanly to DOCX**: financial statements, invoices, balance sheets, contracts, product briefs, sales dashboards, sprint retrospectives, and a React-rendered dashboard.

These are **not** part of the regression loop (`npm run score:suite`). Run them separately to preview conversion quality on richer, portable content without slowing iteration.

**Committed artifacts** live in [`examples/`](../examples/) (`input.html`, `output.docx`, `compare_side_by_side.png`). The harness writes additional scratch files to gitignored `output/showcase/`.

Stress patterns removed from earlier showcase drafts (inline SVG, CSS bar divs in table cells, emoji-heavy layouts) now live in edge test cases — see `tools/generator.ts`.

## Run

```bash
npm run showcase      # 9 showcases → examples/ + output/showcase/
```

Requires LibreOffice (`soffice`) and Playwright Chromium (same as the main loop).

## Output

| Location | Contents |
|----------|----------|
| **`examples/{name}/`** | `input.html`, `output.docx`, `compare_side_by_side.png` (committed) |
| **`output/showcase/{name}/`** | Full harness: wrapped `source.html`, PDFs, PNGs, `meta.json`, diffs |
| **`output/showcase/results.json`** | Scored summary (comparison only, not CI-gated) |

## Examples

| Name | Category | Content |
|------|----------|---------|
| `quarterly-financials` | financial | Condensed income statement with YoY columns and subtotals |
| `invoice` | financial | Professional services invoice with logo via imageResolver, line items, tax, and payment terms |
| `balance-sheet` | financial | Condensed balance sheet — assets, liabilities, and equity |
| `sales-contract` | legal | Software subscription agreement with numbered terms and signatures |
| `product-launch-brief` | work-product | Hero banner, KPI table, funnel table, roadmap list (embedded CSS → computed styles) |
| `javascript-essay` | essay | Three reasons JavaScript is great — headings, blockquote, editorial layout |
| `regional-sales-dashboard` | charts | Regional metrics table with conditional coloring and leaderboard (embedded CSS → computed styles) |
| `sprint-retrospective` | work-product | Retro lists, action-item table with status badges |
| `react-dashboard` | work-product | KPI cards rendered by React (CDN); converts `root.innerHTML` snapshot |

See [examples/README.md](../examples/README.md) for side-by-side previews.

## HTML conventions

Showcase HTML sticks to patterns the converter handles well:

- Semantic tables (`border`, `cellpadding`, row/cell shading, aligned columns)
- Headings, paragraphs, ordered/unordered lists, blockquotes, horizontal rules
- Block backgrounds via shaded `<div>` elements
- Inline highlights via `<span style="background:#…">`
- Links, bold, italic, and `<code>`

Avoid in showcases (use edge test cases instead): inline SVG, nested layout divs inside table cells, emoji-as-icons.

## Scoring

Uses the same content-region pixelmatch and engine score weights as the regression suite, for comparison only.

Add new showcases in `tools/showcase.ts` (not `tools/generator.ts`). Static HTML uses the `html` field; React previews use `reactPreviewPath` (see `react-dashboard`).
