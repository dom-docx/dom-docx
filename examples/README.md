# Examples

Curated HTML → DOCX demos you can open, convert, and compare without running the test harness.

Each folder contains:

| File | Description |
|------|-------------|
| `input.html` | Body fragment passed to `convertHtmlToDocx()` |
| `output.docx` | Generated Word document |
| `compare_side_by_side.png` | Chromium reference (left) vs LibreOffice-rendered DOCX (right) |

Regenerate with `npm run showcase`. The runner refreshes these files automatically; full harness artifacts (PDFs, diffs, scores) stay in gitignored `output/showcase/`.

## Showcases

| Folder | Description |
|--------|-------------|
| [quarterly-financials](./quarterly-financials/) | Condensed income statement with YoY columns and subtotals |
| [invoice](./invoice/) | Professional services invoice with logo (`logo.png` + imageResolver), line items, tax, and payment terms |
| [balance-sheet](./balance-sheet/) | Condensed balance sheet — assets, liabilities, and equity |
| [sales-contract](./sales-contract/) | Software subscription agreement with numbered terms and signatures |
| [product-launch-brief](./product-launch-brief/) | Hero banner, KPI table, funnel table, roadmap list (embedded CSS) |
| [javascript-essay](./javascript-essay/) | Essay — three reasons JavaScript is great |
| [regional-sales-dashboard](./regional-sales-dashboard/) | Regional metrics table with conditional coloring and leaderboard (embedded CSS) |
| [sprint-retrospective](./sprint-retrospective/) | Retro lists, action-item table with status badges |
| [react-dashboard](./react-dashboard/) | KPI cards rendered by React (CDN, no build); `preview.html` is the source app |

## Try one locally

```typescript
import { readFile, writeFile } from "node:fs/promises";
import { convertHtmlToDocx } from "dom-docx";

const html = await readFile("examples/javascript-essay/input.html", "utf-8");
await writeFile("javascript-essay.docx", await convertHtmlToDocx(html));
```

HTML authoring tips: [AGENTS.md](../AGENTS.md).
