# Agent guide: HTML → DOCX (dom-docx)

Use this when an AI agent needs to produce **editable Word documents** from HTML. dom-docx converts semantic, inline-styled HTML into native OOXML (paragraphs, runs, lists, tables)—not screenshots or layout hacks.

**Workflow:** agent writes HTML fragment → `convertHtmlToDocx(html)` → `.docx` buffer → save or attach.

```typescript
import { convertHtmlToDocx } from "./converter.js";

const docx = await convertHtmlToDocx(htmlFragment);
await writeFile("output.docx", docx);
```

Pass a **body fragment only** (no `<html>`, `<head>`, or `<body>` required). The converter wraps content in a letter-size document with 1″ margins and Arial 14pt body text.

---

## Design goal for agents

Optimize for three properties (in order):

1. **Valid, native structure** — paragraphs, list numbering, table rows/cells, hyperlinks
2. **Visual fidelity** — layout that survives Word/LibreOffice rendering
3. **Simplicity** — fewer nested wrappers, inline styles over external CSS, no JS

Think “Word-friendly semantic HTML,” not “web app layout.”

---

## What converts best

### Tier 1 — Excellent (~90–96% visual fidelity)

| Pattern | Notes |
|---------|--------|
| **Headings + paragraphs** | `<h1>`–`<h3>`, plain `<p>`, optional `color`, `font-size`, `text-align` |
| **Lists** | `<ul>`, `<ol>`, nested lists; one idea per `<li>` |
| **Simple tables** | `border="1"`, `cellpadding`, `width:100%`, header row, `text-align:right` on numeric columns |
| **Inline formatting** | `<strong>`, `<em>`, `<a href="…">`, `<code>`, `<br>` |
| **Inline highlights** | `<span style="background:#cfc">term</span>` on short phrases |
| **Blockquote** | `border-left:4px solid #333; padding-left:12px` |
| **Horizontal rule** | `<hr>` between sections |
| **Thematic breaks** | Section headings, not layout divs |

### Tier 2 — Good (~80–92%)

| Pattern | Notes |
|---------|--------|
| **Shaded banner blocks** | `<div style="background:#eaeaea;padding:10px 16px">` with headings inside |
| **Table row/cell styling** | `background`, `color` on `<tr>` or `<td>`; subtotal rows with `<strong>` |
| **Financial / KPI tables** | Label column left-aligned, numbers right-aligned; pastel row bands |
| **Flex rows/columns** | `display:flex; gap:8px` with 2–4 child divs (converted to borderless tables) |
| **Multi-side bordered boxes** | `border:1px solid #ccc; padding:12px` on a div (uses a 1×1 table wrapper when needed) |
| **Typography colors** | `color:#666` on subtitles, accent colors on deltas (`#2a9d8f`, `#e76f51`) |

### Tier 3 — Weak or unsupported — avoid

| Pattern | Why |
|---------|-----|
| **Inline SVG / Canvas / `<img>` charts** | Not rendered as vectors; use tables, describe data in text, or enable **`rasterizeInPlace`** when exporting from a live rendered page |
| **CSS bar charts inside cells** | `<div style="height:14px;width:80%;background:…">` in `<td>` does not port |
| **Emoji as UI icons** | Font/glyph mismatch in Word |
| **External stylesheets** | Only inline `style=""` and a few attributes (`border`, `cellpadding`, `colspan`) |
| **Grid / absolute / float layout** | No `position`, `float`, `grid-template` |
| **Multi-column CSS** | No `column-count` |
| **Web fonts** | Defaults to Arial; custom `@font-face` ignored |
| **Forms, inputs, buttons** | Not supported |
| **Deeply nested layout divs** | Prefer flat block flow: heading → p → table → p |
| **`<colgroup>` / `<col width>`** | Column widths inferred from content, not col tags |

---

## HTML style guide for agents

### Page and typography defaults

The engine assumes **US Letter**, **1″ margins**, **Arial 14px**, **line-height ~1.4**. You do not need a wrapper document; if you preview in a browser, match:

```html
<!-- Optional preview wrapper only — omit when calling convertHtmlToDocx -->
<style>
  body {
    margin: 0;
    padding: 96px;
    width: 816px;
    font-family: Arial, Helvetica, sans-serif;
    font-size: 14px;
    line-height: 1.4;
    color: #111;
  }
</style>
```

Use **hex colors** (`#1a1a2e`, `#666`, `#f5f5f5`). Named CSS colors work when parsed, but hex is predictable.

### Block structure

Prefer a **linear document flow**:

```
h1 title
p subtitle (muted color, smaller font-size)
table | ul | blockquote
h2 section
p …
```

Each visual “section” should be a heading or table—not a tower of anonymous divs.

**Good — shaded hero:**

```html
<div style="background:#eaeaea;padding:10px 16px;margin-bottom:12px">
  <h1 style="margin:0;font-size:20px;color:#1a1a2e">Sprint 24 Retrospective</h1>
  <p style="margin:4px 0 0;color:#666;font-size:13px">Platform Team · Mar 3 – Mar 14, 2026</p>
</div>
```

**Bad — layout soup:**

```html
<div><div><div style="display:grid;grid-template-columns:1fr 1fr">…</div></div></div>
```

### Tables (data, not layout)

Use tables for **tabular data only**. Always include:

- `border="1"` (or explicit border in style) when grid lines are wanted
- `cellpadding="6"` or `8`
- `style="border-collapse:collapse;width:100%"`
- Header row with `<strong>` labels
- `text-align:right` on numeric `<td>` cells

**Good — financial row bands:**

```html
<table border="1" cellpadding="8" style="border-collapse:collapse;width:100%">
  <tr style="background:#1a1a2e;color:#f1faee">
    <td><strong>Line item</strong></td>
    <td style="text-align:right"><strong>Q1 2026</strong></td>
  </tr>
  <tr style="background:#f5f5f5">
    <td><strong>Subtotal</strong></td>
    <td style="text-align:right"><strong>$33,424</strong></td>
  </tr>
</table>
```

Put **text inside cells** (`<td>`, `<th>`)—not nested layout divs, charts, or icons.

### Inline vs block backgrounds

| Intent | HTML | DOCX mapping |
|--------|------|----------------|
| Highlight a phrase | `<span style="background:#cfc">text</span>` | Shading on `TextRun` |
| Banner / callout / row band | `background` on `<div>`, `<tr>`, or `<td>` | Paragraph or cell shading |
| Whole paragraph tint | `background` on `<p>` | Paragraph shading |

Keep inline highlights **short** (a few words). Long shaded spans across line wraps are OK but may differ slightly from browser wrapping.

### Lists

- Use `<ul>` / `<ol>` for enumerations, action items, roadmaps
- Keep list item content in inline flow; avoid block elements inside `<li>` when possible
- Nested lists are supported (ordered inside unordered, etc.)

### Links

```html
<p>See <a href="https://example.com/report">the full report</a> for details.</p>
```

Links render as Word hyperlinks (underlined, Hyperlink style). Use real `href` values.

### Flex (simple only)

Supported for **small toolbars, KPI chips, column stacks**:

```html
<div style="display:flex;flex-direction:row;gap:10px;padding:12px;background:#f5f5f5">
  <div style="background:#ddd;padding:8px">Left</div>
  <div style="background:#bbb;padding:8px">Right</div>
</div>
```

Do **not** use flex for page layout, sticky headers, or precise pixel alignment. Max ~4 items per row.

### Line breaks and addresses

```html
<p>
  742 Evergreen Terrace<br>
  Springfield, OR 97403
</p>
```

`<br>` is supported inside paragraphs. Do not simulate line breaks with empty `<p>` tags.

---

## Supported inline CSS properties

Parsed from `style=""` on any element:

| Property | Example |
|----------|---------|
| `color` | `color:#666` |
| `background` / `background-color` | `background:#f5f5f5` |
| `text-align` | `left`, `center`, `right`, `justify` |
| `font-size` | `13px`, `12px` (px preferred) |
| `font-weight` | `bold`, `600` |
| `font-style` | `italic` |
| `margin`, `margin-*` | `margin-top:0`, `margin:8px 0` |
| `padding`, `padding-*` | `padding:8px`, `padding-left:12px` |
| `border`, `border-*` | Blockquote left rule, boxed sections |
| `display` | `flex`, `block`, `inline-block` |
| `flex-direction` | `row`, `column` |
| `gap` | `gap:10px` (px; converted to twips internally) |
| `width` | `width:100%` on tables |

Everything else is ignored silently.

---

## Supported elements

| Element | Role |
|---------|------|
| `h1`–`h6` | Headings (mapped to Word heading levels) |
| `p` | Body paragraphs |
| `div`, `section` | Block containers; shading and padding |
| `ul`, `ol`, `li` | Bulleted / numbered lists |
| `table`, `tr`, `td`, `th` | Data tables (`colspan` supported) |
| `thead`, `tbody`, `tfoot` | Table structure (rows collected) |
| `blockquote` | Indented quotations |
| `hr` | Thematic break |
| `strong`, `b`, `em`, `i`, `u` | Inline emphasis |
| `a` | Hyperlinks |
| `span` | Inline color / background |
| `code` | Monospace inline text |
| `br` | Line break within paragraph |
| `pre` | Monospace block (limited) |

Unsupported tags are treated as generic block containers or skipped.

---

## Document archetypes (copy these shapes)

### Memo / policy essay (~92% fidelity)

```
h1 title
p routing line (To / From / Date — muted)
table routing metadata (2 columns)
h2 section
p body
blockquote excerpt
ol requirements
p closing
```

### Financial statement (~83% fidelity)

```
h1 company
p subtitle (period, units — small gray)
table 4-column grid, shaded header + subtotal rows
p footnote with inline highlight + link
```

### Product one-pager (~78% fidelity)

```
div hero banner (shaded)
table KPI metrics (2–4 columns)
table funnel stages (label + value columns)
ol roadmap milestones
p contact line
```

### Retro / action tracker (~90% fidelity)

```
div banner
h2 + ul (went well / to improve)
table Owner | Action | Status with row backgrounds
```

---

## Agent checklist before convert

- [ ] Fragment is **body content only** (no `<!DOCTYPE>`, scripts, or external CSS)
- [ ] **Tables** used for data, not page layout; numeric columns right-aligned
- [ ] **Backgrounds** on `span` (inline) vs `div`/`tr`/`td` (block) used intentionally
- [ ] **No SVG, images, emoji icons, or CSS drawing** inside tables
- [ ] **Headings** encode document structure (`h1` once, then `h2` sections)
- [ ] **Colors** are hex; contrast readable (light text on dark row bands)
- [ ] **Links** have `href`; important labels use `<strong>` not custom font hacks
- [ ] **Flex** limited to simple rows/columns with ≤4 items
- [ ] Run `npm run typecheck` after pipeline changes; smoke-test with `convertHtmlToDocx`

---

## Validation commands (human or CI agent)

```bash
npm run typecheck          # TypeScript
npm run test:suite          # full suite: visual + XML + editability (cases: tools/generator.ts)
npm run test:suite:priority # fast subset of the same cases
npm run test:showcase      # 6 rich real-world examples
```

After conversion, OOXML should pass schema validation. For critical documents, open `output.docx` in Word and verify tables are editable cell-by-cell and lists use native numbering.

---

## Minimal starter template

```html
<h1 style="color:#1a1a2e;margin-bottom:4px">Document Title</h1>
<p style="color:#666;margin-top:0;font-size:13px">Subtitle or date line</p>

<h2 style="font-size:15px">Section</h2>
<p>Opening paragraph with <strong>emphasis</strong> and a <a href="https://example.com">link</a>.</p>

<ul>
  <li>First point</li>
  <li>Second point</li>
</ul>

<table border="1" cellpadding="8" style="border-collapse:collapse;width:100%">
  <tr style="background:#1a1a2e;color:#f1faee">
    <td><strong>Column A</strong></td>
    <td style="text-align:right"><strong>Column B</strong></td>
  </tr>
  <tr>
    <td>Row label</td>
    <td style="text-align:right">123</td>
  </tr>
</table>

<p style="font-size:12px;color:#666;margin-top:12px">
  Footnote with <span style="background:#cfc">highlighted term</span>.
</p>
```

---

## Mental model

```
Browser HTML (layout engine)  ≠  Word (flow + styles + tables)
dom-docx maps a practical subset:  semantic blocks → native OOXML
```

Agents should generate HTML that would still read well as a **printed memo**—not HTML that only makes sense with CSS layout engines, JavaScript, or embedded graphics.

When in doubt: **heading, paragraph, table, list, span highlight.** That path produces the highest-fidelity, most editable DOCX.
