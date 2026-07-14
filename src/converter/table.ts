import {
  AlignmentType,
  BorderStyle,
  HeightRule,
  LineRuleType,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import type { CheerioAPI } from "cheerio";
import type { AnyNode, Element } from "domhandler";
import { BODY_LINE_EXACT_TWIPS, HEADING_FONT_HALF_POINTS } from "./constants.js";
import {
  cssToBlockTypography,
  inheritedFontFamily,
  isHiddenElement,
  mapTextAlign,
  pxToTwips,
  type ParsedBorder,
  type ParsedCss,
} from "./css.js";
import type { StyleResolver } from "./style-resolver.js";
import { INLINE_STYLE_RESOLVER } from "./style-resolver.js";
import { collectInlineRunsFromNodes } from "./inline.js";
import {
  estimateTextWidthTwips,
  minContentWidthTwips,
} from "./text-metrics.js";
import type { DocxBlock } from "./types.js";
import type { RunTypography } from "./types.js";
import { DEFAULT_VISITOR_CONTEXT } from "./types.js";
// Runtime-only cycle with visitor.ts (function declarations hoist): cells reuse
// the body list-numbering path for `<ul>`/`<ol>` children.
import { processList } from "./visitor.js";

/** Usable content width: 8.5" page − 2 × 1" margins = 6.5" → 9360 twips. */
const CONTENT_WIDTH_TWIPS = 9360;

// Chromium renders `border="1"` collapsed cell borders as a 1px #eee hairline;
// 0.5pt strokes rasterize 2px wide in LibreOffice and double the border ink.
const BORDER_COLOR = "eeeeee";
const BORDER_SIZE = 2;

interface ParsedCell {
  element: Element;
  row: Element;
  table: Element;
  colspan: number;
  rowspan: number;
}

/** A cell with its resolved grid column (accounts for row-spanning cells above). */
interface PlacedCell {
  cell: ParsedCell;
  columnIndex: number;
}

interface GridAnalysis {
  maxColumns: number;
  placedRows: PlacedCell[][];
  columnWidths: number[];
  /** Resolved table width (twips) — declared width or shrink-wrapped content width. */
  totalWidth: number;
}

function parseColspan(cell: Element): number {
  return Math.max(1, parseInt(cell.attribs?.colspan ?? "1", 10) || 1);
}

function parseRowspan(cell: Element): number {
  return Math.max(1, parseInt(cell.attribs?.rowspan ?? "1", 10) || 1);
}

/**
 * Resolve each cell's grid column, skipping columns occupied by rowspanning
 * cells from earlier rows (HTML grid placement).
 */
function placeCells(rows: ParsedCell[][]): { placedRows: PlacedCell[][]; maxColumns: number } {
  const placedRows: PlacedCell[][] = [];
  const carry: Array<{ col: number; span: number; rowsLeft: number }> = [];
  let maxColumns = 1;

  for (const row of rows) {
    const occupied = new Set<number>();
    for (const c of carry) {
      if (c.rowsLeft > 0) for (let i = 0; i < c.span; i++) occupied.add(c.col + i);
    }

    const placed: PlacedCell[] = [];
    const newCarries: Array<{ col: number; span: number; rowsLeft: number }> = [];
    let col = 0;
    for (const cell of row) {
      while (occupied.has(col)) col += 1;
      placed.push({ cell, columnIndex: col });
      if (cell.rowspan > 1) {
        newCarries.push({ col, span: cell.colspan, rowsLeft: cell.rowspan - 1 });
      }
      col += cell.colspan;
    }

    maxColumns = Math.max(maxColumns, col, ...[...occupied].map((c) => c + 1));
    placedRows.push(placed);
    // Decrement only pre-existing spans — a span opened THIS row still covers
    // all of its rowspan-1 following rows.
    for (const c of carry) c.rowsLeft -= 1;
    carry.push(...newCarries);
  }

  return { placedRows, maxColumns };
}

function collectRowElements($: CheerioAPI, table: Element): Element[] {
  const trElements: Element[] = [];
  for (const child of $(table).children("thead, tbody, tfoot, tr").toArray()) {
    if (child.name.toLowerCase() === "tr") {
      trElements.push(child);
    } else {
      trElements.push(...$(child).children("tr").toArray());
    }
  }
  return trElements;
}

function parseRows(
  $: CheerioAPI,
  trElements: Element[],
  table: Element,
  styleResolver: StyleResolver,
): ParsedCell[][] {
  return trElements
    .filter((tr) => !isHiddenElement(tr, styleResolver))
    .map((tr) =>
      $(tr)
        .children("td, th")
        .toArray()
        .filter((cell) => !isHiddenElement(cell, styleResolver))
        .map((cell) => ({
          element: cell,
          row: tr,
          table,
          colspan: parseColspan(cell),
          rowspan: parseRowspan(cell),
        })),
    );
}

/** Cell styles override row, row overrides inherited table styles (font, color, bg). */
function resolveCellCss(cell: ParsedCell, styleResolver: StyleResolver): ParsedCss {
  const tableCss = styleResolver.getCss(cell.table);
  const rowCss = styleResolver.getCss(cell.row);
  const cellCss = styleResolver.getCss(cell.element);
  return {
    ...rowCss,
    ...cellCss,
    color: cellCss.color ?? rowCss.color ?? tableCss.color,
    fontSize: cellCss.fontSize ?? rowCss.fontSize ?? tableCss.fontSize,
    fontWeight: cellCss.fontWeight ?? rowCss.fontWeight ?? tableCss.fontWeight,
    fontStyle: cellCss.fontStyle ?? rowCss.fontStyle ?? tableCss.fontStyle,
    textAlign: cellCss.textAlign ?? rowCss.textAlign ?? tableCss.textAlign,
    // Table background shows through cells with no fill of their own.
    backgroundColor: cellCss.backgroundColor ?? rowCss.backgroundColor ?? tableCss.backgroundColor,
  };
}

function cellTypography(cell: ParsedCell, styleResolver: StyleResolver): RunTypography {
  const typography = cssToBlockTypography(resolveCellCss(cell, styleResolver));
  if (!typography.font) {
    // Walks cell → row → table → wrapper (CSS font-family inheritance).
    const font = inheritedFontFamily(cell.element, styleResolver);
    if (font) typography.font = font;
  }
  if (cell.element.name.toLowerCase() === "th") {
    typography.bold = true;
  }
  return typography;
}

function cellShading(cell: ParsedCell, styleResolver: StyleResolver) {
  const { backgroundColor } = resolveCellCss(cell, styleResolver);
  if (!backgroundColor) return undefined;
  return {
    type: ShadingType.CLEAR,
    fill: backgroundColor,
    color: "auto" as const,
  };
}

function elementPlainText(element: Element): string {
  const parts: string[] = [];
  function walk(nodes: AnyNode[]): void {
    for (const node of nodes) {
      if (node.type === "text") parts.push(node.data ?? "");
      else if (node.type === "tag") walk(node.children ?? []);
    }
  }
  walk(element.children ?? []);
  return parts.join("").replace(/\s+/g, " ").trim();
}

function cellContentIsBold(cell: ParsedCell, styleResolver: StyleResolver): boolean {
  if (cell.element.name.toLowerCase() === "th") return true;
  const css = resolveCellCss(cell, styleResolver);
  if (css.fontWeight === "bold" || Number(css.fontWeight) >= 600) return true;
  return Boolean(
    cell.element.children?.some(
      (node) => node.type === "tag" && node.name.toLowerCase() === "strong",
    ),
  );
}

function cellTextMeasureOptions(
  cell: ParsedCell,
  styleResolver: StyleResolver,
): { bold?: boolean; fontSizeHalfPoints?: number } {
  const css = resolveCellCss(cell, styleResolver);
  const bold =
    cellContentIsBold(cell, styleResolver) ||
    css.fontWeight === "bold" ||
    Number(css.fontWeight) >= 600;
  return {
    bold,
    fontSizeHalfPoints: css.fontSize,
  };
}

/** Minimal width for layout-only columns (Word/LibreOffice reject true zero-width grid cols). */
const LAYOUT_GUTTER_TWIPS = pxToTwips(2);

function estimateCellTextWidthTwips(cell: ParsedCell, styleResolver: StyleResolver): number {
  const text = elementPlainText(cell.element);
  return Math.max(
    minContentWidthTwips(),
    estimateTextWidthTwips(text, cellTextMeasureOptions(cell, styleResolver)),
  );
}

function estimateCellContentWidthTwips(
  cell: ParsedCell,
  styleResolver: StyleResolver,
  cellPadding?: number,
): number {
  const padX = cellPadding ? cellPadding * 2 : pxToTwips(16);
  return padX + estimateCellTextWidthTwips(cell, styleResolver);
}

/**
 * Empty cells used only for table-layout alignment (vote gutters, rank columns,
 * indent spacers). Collapse unless the author declared a width via CSS or attrs.
 */
function isDecorativeSpacerCell(cell: ParsedCell): boolean {
  if (elementPlainText(cell.element).length > 0) return false;
  return !cell.element.children?.some(
    (node) => node.type === "tag" && (node as Element).name.toLowerCase() === "table",
  );
}

function columnsWithContent(placedRows: GridAnalysis["placedRows"]): boolean[] {
  let maxCol = 0;
  for (const row of placedRows) {
    for (const { cell, columnIndex } of row) {
      maxCol = Math.max(maxCol, columnIndex + cell.colspan);
    }
  }
  const has = Array.from({ length: maxCol }, () => false);
  for (const row of placedRows) {
    for (const { cell, columnIndex } of row) {
      if (isDecorativeSpacerCell(cell)) continue;
      for (let i = 0; i < cell.colspan; i++) has[columnIndex + i] = true;
    }
  }
  return has;
}

/** Explicit `width` on a single-column cell (style px/% or legacy attr) → pinned column twips. */
function explicitCellWidthTwips(
  cell: ParsedCell,
  styleResolver: StyleResolver,
  contentWidthTwips: number,
): number | undefined {
  if (cell.colspan !== 1) return undefined;
  const css = styleResolver.getCss(cell.element);
  if (css.widthTwips !== undefined) return css.widthTwips;
  if (css.widthPercent !== undefined) {
    return Math.round((css.widthPercent / 100) * contentWidthTwips);
  }
  const attr = cell.element.attribs?.width?.trim();
  if (attr) {
    const percent = attr.match(/^(\d+(?:\.\d+)?)%$/);
    if (percent) return Math.round((parseFloat(percent[1]!) / 100) * contentWidthTwips);
    const px = parseFloat(attr);
    if (Number.isFinite(px)) return pxToTwips(px);
  }
  return undefined;
}

/**
 * Pinned columns keep their explicit width; the rest share the remaining width
 * by content weight. Falls back to pure content weighting when the pins leave
 * no room for the unpinned columns.
 */
function distributeWithPinnedColumns(
  weights: number[],
  pinned: Array<number | undefined>,
  total: number,
): number[] {
  const unpinnedIdx = pinned
    .map((p, i) => (p === undefined ? i : -1))
    .filter((i) => i >= 0);
  if (unpinnedIdx.length === weights.length) return distributeColumnWidths(weights, total);
  if (unpinnedIdx.length === 0) {
    return distributeColumnWidths(pinned.map((p) => p ?? 0), total);
  }

  const pinnedSum = pinned.reduce<number>((acc, p) => acc + (p ?? 0), 0);
  const free = total - pinnedSum;
  if (free < unpinnedIdx.length * pxToTwips(32)) {
    return distributeColumnWidths(weights, total);
  }

  const freeWidths = distributeColumnWidths(
    unpinnedIdx.map((i) => weights[i]!),
    free,
  );
  const out = pinned.map((p) => p ?? 0);
  unpinnedIdx.forEach((i, k) => {
    out[i] = freeWidths[k]!;
  });
  return out;
}

/** Scale column weights to exactly `total` twips (table is width:100%). */
function distributeColumnWidths(weights: number[], total: number): number[] {
  if (weights.length === 0) return [];

  const sum = weights.reduce((acc, w) => acc + w, 0);
  if (sum <= 0) {
    const base = Math.floor(total / weights.length);
    const equal = Array.from({ length: weights.length }, () => base);
    for (let i = 0; i < total - base * weights.length; i++) {
      equal[i]! += 1;
    }
    return equal;
  }

  const scaled = weights.map((w) => (w / sum) * total);
  const columnWidths = scaled.map((w) => Math.floor(w));
  let remainder = total - columnWidths.reduce((acc, w) => acc + w, 0);
  const fractional = scaled
    .map((w, index) => ({ index, fraction: w - Math.floor(w) }))
    .sort((a, b) => b.fraction - a.fraction);
  for (let i = 0; remainder > 0; i++, remainder--) {
    columnWidths[fractional[i % fractional.length]!.index]! += 1;
  }
  return columnWidths;
}

/** Pass 1 — derive max grid columns and content-weighted column widths. */
function analyzeGrid(
  rows: ParsedCell[][],
  styleResolver: StyleResolver,
  cellPadding: number | undefined,
  contentWidthTwips: number,
  declaredWidthTwips?: number,
  fillParent = false,
  colWidths: Array<number | undefined> = [],
): GridAnalysis {
  const { placedRows, maxColumns } = placeCells(rows);

  const columnHasContent = columnsWithContent(placedRows);

  const columnMinWidths = Array.from({ length: maxColumns }, (_, i) =>
    columnHasContent[i] ? pxToTwips(32) : LAYOUT_GUTTER_TWIPS,
  );
  // Seed pinned widths from `<colgroup>`; an explicit per-cell width can still raise
  // a column below (max wins). Colgroup-pinned columns also feed the natural table
  // width, so a colgroup-sized table stretches to its declared column total.
  const pinnedWidths: Array<number | undefined> = Array.from(
    { length: maxColumns },
    (_, i) => colWidths[i],
  );

  for (const row of placedRows) {
    for (const { cell, columnIndex } of row) {
      const span = cell.colspan;
      const explicit = explicitCellWidthTwips(cell, styleResolver, contentWidthTwips);
      if (explicit !== undefined) {
        pinnedWidths[columnIndex] = Math.max(pinnedWidths[columnIndex] ?? 0, explicit);
      }
      if (isDecorativeSpacerCell(cell)) continue;
      const need = estimateCellContentWidthTwips(cell, styleResolver, cellPadding);
      const currentSpanWidth = sumColumnWidths(columnMinWidths, columnIndex, span);
      if (need > currentSpanWidth) {
        const perCol = (need - currentSpanWidth) / span;
        for (let i = 0; i < span; i++) {
          columnMinWidths[columnIndex + i]! += perCol;
        }
      }
    }
  }

  // HTML tables shrink-to-fit by default — only a declared width stretches
  // them (agent HTML habitually writes width:100%; wild HTML often doesn't).
  const naturalWidth = Math.round(
    columnMinWidths.reduce((sum, w, i) => sum + (pinnedWidths[i] ?? w), 0),
  );
  const totalWidth = Math.min(
    contentWidthTwips,
    declaredWidthTwips ?? (fillParent ? contentWidthTwips : naturalWidth),
  );
  const columnWidths = distributeWithPinnedColumns(columnMinWidths, pinnedWidths, totalWidth);

  return { maxColumns, placedRows, columnWidths, totalWidth };
}

function sumColumnWidths(columnWidths: number[], start: number, span: number): number {
  return columnWidths.slice(start, start + span).reduce((total, w) => total + w, 0);
}

function parseCellPadding(table: Element): number | undefined {
  const cellpadding = table.attribs?.cellpadding;
  if (!cellpadding) return undefined;
  const px = parseFloat(cellpadding);
  return Number.isFinite(px) ? pxToTwips(px) : undefined;
}

/** One `<col>`'s width in twips: `style="width:N%|Npx"` or `width="N%|N"`, else undefined. */
function colWidthTwips(col: Element, contentWidthTwips: number): number | undefined {
  const style = col.attribs?.style ?? "";
  const styleMatch = style.match(/width\s*:\s*([\d.]+)\s*(%|px)?/i);
  if (styleMatch) {
    const value = parseFloat(styleMatch[1]!);
    if (Number.isFinite(value)) {
      return styleMatch[2] === "%"
        ? Math.round((value / 100) * contentWidthTwips)
        : pxToTwips(value);
    }
  }
  const attr = col.attribs?.width?.trim();
  if (attr) {
    const pct = attr.match(/^([\d.]+)%$/);
    if (pct) return Math.round((parseFloat(pct[1]!) / 100) * contentWidthTwips);
    const px = parseFloat(attr);
    if (Number.isFinite(px)) return pxToTwips(px);
  }
  return undefined;
}

/**
 * Per-column widths (twips) declared by `<colgroup><col>` — expanded for `span`.
 * Real-world tables (docs sites, DocBook output) size columns entirely via colgroup
 * `width:33%` cols with no `width` on the table itself; without this, dom-docx has no
 * width signal and collapses every column to its min-content width (~1 char wide).
 */
function colgroupColumnWidths(
  $: CheerioAPI,
  table: Element,
  contentWidthTwips: number,
): Array<number | undefined> {
  const cols = $(table).children("colgroup").children("col").toArray();
  if (cols.length === 0) return [];
  const widths: Array<number | undefined> = [];
  for (const col of cols) {
    const span = Math.max(1, parseInt(col.attribs?.span ?? "1", 10) || 1);
    const width = colWidthTwips(col, contentWidthTwips);
    for (let i = 0; i < span; i++) widths.push(width);
  }
  return widths;
}

/** Declared table width from CSS or the legacy `width` attribute. */
function tableDeclaredWidth(
  table: Element,
  styleResolver: StyleResolver,
): { twips?: number; percent?: number } {
  const css = styleResolver.getCss(table);
  if (css.widthPercent !== undefined) return { percent: css.widthPercent };
  if (css.widthTwips !== undefined) return { twips: css.widthTwips };
  const maxWidth = css.maxWidthTwips;
  const attr = table.attribs?.width?.trim();
  if (attr) {
    const pct = attr.match(/^(\d+(?:\.\d+)?)%$/);
    if (pct) return { percent: parseFloat(pct[1]!) };
    const px = parseFloat(attr);
    if (Number.isFinite(px)) return { twips: pxToTwips(px) };
  }
  // Email idiom: fluid width capped by max-width (`width:100%;max-width:600px`
  // parses as percent above; bare max-width acts as the effective width).
  if (maxWidth !== undefined) return { twips: maxWidth };
  return {};
}

/** Table positioning: `align` attribute or CSS auto margins. */
function tableAlignment(
  table: Element,
): (typeof AlignmentType)[keyof typeof AlignmentType] | undefined {
  const align = table.attribs?.align?.trim().toLowerCase();
  if (align === "center") return AlignmentType.CENTER;
  if (align === "right") return AlignmentType.RIGHT;
  const style = table.attribs?.style ?? "";
  const leftAuto = /margin-left\s*:\s*auto/i.test(style);
  const rightAuto = /margin-right\s*:\s*auto/i.test(style);
  const shorthandAuto = /margin\s*:\s*[^;]*\bauto\b/i.test(style);
  if (shorthandAuto || (leftAuto && rightAuto)) return AlignmentType.CENTER;
  if (leftAuto) return AlignmentType.RIGHT;
  return undefined;
}

/** `border` attribute ≥ 1 → Chromium paints every cell border (full grid). */
function hasAttrGrid(table: Element): boolean {
  const borderAttr = table.attribs?.border;
  return borderAttr !== undefined && borderAttr !== "0";
}

/**
 * HTML separate-borders model: without `border-collapse:collapse`, cells sit
 * 2px apart (UA border-spacing) unless `cellspacing` says otherwise, and the
 * page background shows through — visible as white lines across shaded rows.
 * (Word's native w:tblCellSpacing is ignored by LibreOffice, so the gaps are
 * emulated with white borders instead.)
 */
function separateBorderSpacingPx(table: Element, styleResolver: StyleResolver): number {
  if (styleResolver.getCss(table).borderCollapse === "collapse") return 0;
  // Layout tables (`border="0"`) have no visible cell grid in browsers even when
  // cellspacing defaults to 2 — emulating gaps as white borders shows as gridlines
  // on shaded backgrounds (HN, email templates).
  if (table.attribs?.border === "0") return 0;
  const spacingAttr = table.attribs?.cellspacing;
  if (spacingAttr !== undefined) return parseFloat(spacingAttr) || 0;
  return 2;
}

const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "auto" };

interface TableBorderPlan {
  /** Full cell grid (border attr) — padding cells carve holes out of this. */
  grid: boolean;
  gridColor: string;
  borders: Record<string, { style: (typeof BorderStyle)[keyof typeof BorderStyle]; size: number; color: string }>;
}

/**
 * Border plan: the `border` attr means a full hairline cell grid; CSS borders on
 * the table element style only the OUTER frame (per-side width/color), never the
 * inside gridlines.
 */
function tableBorderPlan(table: Element, styleResolver: StyleResolver): TableBorderPlan {
  const css = styleResolver.getCss(table);
  const gridColor = css.borderColor ?? css.border?.color ?? BORDER_COLOR;

  if (hasAttrGrid(table)) {
    const side = { style: BorderStyle.SINGLE, size: BORDER_SIZE, color: gridColor };
    return {
      grid: true,
      gridColor,
      borders: {
        top: side,
        bottom: side,
        left: side,
        right: side,
        insideHorizontal: side,
        insideVertical: side,
      },
    };
  }

  const toSide = (b: ParsedBorder | undefined) =>
    b
      ? {
          style: BorderStyle.SINGLE,
          size: Math.max(2, Math.round(b.widthPx * 6)),
          color: b.color ?? css.borderColor ?? "000000",
        }
      : undefined;
  const frame = {
    top: toSide(css.borderTop ?? css.border),
    right: toSide(css.borderRight ?? css.border),
    bottom: toSide(css.borderBottom ?? css.border),
    left: toSide(css.borderLeft ?? css.border),
  };

  const spacingPx = separateBorderSpacingPx(table, styleResolver);
  const hasVisibleFrame = Boolean(
    frame.top || frame.right || frame.bottom || frame.left,
  );
  const gap =
    spacingPx > 0 && (hasAttrGrid(table) || hasVisibleFrame)
      ? {
          style: BorderStyle.SINGLE,
          size: Math.max(2, Math.round(spacingPx * 6)),
          color: "FFFFFF",
        }
      : NO_BORDER;

  return {
    grid: false,
    gridColor,
    borders: {
      top: frame.top ?? gap,
      right: frame.right ?? gap,
      bottom: frame.bottom ?? gap,
      left: frame.left ?? gap,
      insideHorizontal: gap,
      insideVertical: gap,
    },
  };
}

/** EXACT line twips for a font size (half-points): px × 1.4 line-height × 15. */
function exactLineForFontSize(fontSizeHalfPoints: number | undefined): number {
  if (!fontSizeHalfPoints) return BODY_LINE_EXACT_TWIPS;
  return Math.round(fontSizeHalfPoints * 14);
}

/** EXACT line boxes clip inline images — image-bearing paragraphs need AUTO. */
function nodesContainImage(nodes: AnyNode[]): boolean {
  for (const node of nodes) {
    if (node.type !== "tag") continue;
    const el = node as Element;
    if (el.name.toLowerCase() === "img") return true;
    if (nodesContainImage(el.children ?? [])) return true;
  }
  return false;
}

function cellParagraph(
  cell: ParsedCell,
  styleResolver: StyleResolver,
  nodes?: AnyNode[],
): Paragraph {
  const css = resolveCellCss(cell, styleResolver);
  const typography = cellTypography(cell, styleResolver);
  const content = nodes ?? cell.element.children ?? [];
  const children = content.length
    ? collectInlineRunsFromNodes(content, typography, undefined, styleResolver)
    : [new TextRun("")];
  const hasImage = nodesContainImage(content);
  return new Paragraph({
    ...(css.textAlign ? { alignment: mapTextAlign(css.textAlign) } : {}),
    spacing: hasImage
      ? { before: 0, after: 0 }
      : {
          before: 0,
          after: 0,
          line: exactLineForFontSize(typography.fontSize),
          lineRule: LineRuleType.EXACT,
        },
    children,
  });
}

/**
 * Block child of a cell (`<p>`, `<h1>`–`<h6>`) → its own paragraph with the
 * block's typography, margins, and a line box sized to ITS font (a 28px heading
 * squeezed into the body EXACT line overlaps the paragraph above).
 */
function cellBlockParagraph(
  cell: ParsedCell,
  styleResolver: StyleResolver,
  element: Element,
): Paragraph {
  const cellCss = resolveCellCss(cell, styleResolver);
  const blockCss = styleResolver.getCss(element);
  const tag = element.name.toLowerCase();
  const isHeading = /^h[1-6]$/.test(tag);

  const fontSize =
    blockCss.fontSize ??
    (isHeading
      ? HEADING_FONT_HALF_POINTS[tag as keyof typeof HEADING_FONT_HALF_POINTS]
      : undefined) ??
    cellCss.fontSize;
  const typography: RunTypography = {
    ...cellTypography(cell, styleResolver),
    ...cssToBlockTypography(blockCss),
    ...(isHeading ? { bold: true } : {}),
    fontSize,
  };
  const align = blockCss.textAlign ?? cellCss.textAlign;
  const hasImage = nodesContainImage(element.children ?? []);

  // UA default `<p>` margins (1em of the paragraph's font), same as body flow:
  // margin-bottom always; margin-top only when the previous sibling isn't a
  // `<p>` (adjacent-paragraph margins collapse in HTML, but stack in Word).
  const isParagraphTag = tag === "p";
  const uaMarginTwips = Math.round(((fontSize ?? 21) / 1.5) * 15);
  let prev = element.prev;
  while (prev && prev.type !== "tag") prev = prev.prev;
  const prevIsParagraph =
    prev?.type === "tag" && (prev as Element).name.toLowerCase() === "p";
  const before =
    blockCss.marginTop ?? (isParagraphTag && !prevIsParagraph ? uaMarginTwips : 0);
  const after = blockCss.marginBottom ?? (isParagraphTag ? uaMarginTwips : 0);

  return new Paragraph({
    ...(align ? { alignment: mapTextAlign(align) } : {}),
    spacing: hasImage
      ? { before, after }
      : {
          before,
          after,
          line: exactLineForFontSize(fontSize),
          lineRule: LineRuleType.EXACT,
        },
    children: collectInlineRunsFromNodes(
      element.children ?? [],
      typography,
      undefined,
      styleResolver,
    ),
  });
}

const DEFAULT_BAR_HEIGHT_TWIPS = pxToTwips(12);

/**
 * Empty block with a background fill and an explicit height/width — a CSS "trend bar"
 * (e.g. `<div style="background:#457b9d;height:14px;width:76%">`). These carry no text,
 * so the inline-run path drops them; render as a shaded band instead.
 */
function colorBarCss(node: AnyNode, styleResolver: StyleResolver): ParsedCss | null {
  if (node.type !== "tag") return null;
  const element = node as Element;
  if (elementPlainText(element).length > 0) return null;
  if (element.children?.some((child) => child.type === "tag")) return null;
  const css = styleResolver.getCss(element);
  if (!css.backgroundColor) return null;
  if (
    css.heightTwips === undefined &&
    css.widthPercent === undefined &&
    css.widthTwips === undefined
  ) {
    return null;
  }
  return css;
}

/**
 * Shaded band sized like the source bar: EXACT line height from CSS `height`, and the
 * paragraph's right indent trims the shading to `width` (percent of the cell's content
 * width, capped by `max-width`). Same pattern as the SVG bar renderer — native
 * paragraph shading, no nested layout tables.
 */
function barParagraph(css: ParsedCss, contentWidthTwips: number): Paragraph {
  const height = css.heightTwips ?? DEFAULT_BAR_HEIGHT_TWIPS;
  let barWidth =
    css.widthPercent !== undefined
      ? Math.round((css.widthPercent / 100) * contentWidthTwips)
      : (css.widthTwips ?? contentWidthTwips);
  if (css.maxWidthTwips !== undefined) barWidth = Math.min(barWidth, css.maxWidthTwips);
  barWidth = Math.max(0, Math.min(barWidth, contentWidthTwips));
  return new Paragraph({
    shading: { type: ShadingType.CLEAR, fill: css.backgroundColor!, color: "auto" },
    indent: { right: contentWidthTwips - barWidth },
    spacing: { before: 0, after: 0, line: height, lineRule: LineRuleType.EXACT },
    children: [new TextRun("")],
  });
}

function nodeHasInlineContent(node: AnyNode): boolean {
  if (node.type === "text") return (node.data ?? "").trim().length > 0;
  return node.type === "tag";
}

/** Cell content: inline runs, shaded bars, nested tables, and explicit `<p>` blocks. */
/** Structural wrappers inside cells (`td > div.content > …`) — recursed, not flattened. */
const CELL_CONTAINER_TAG = /^(?:div|section|center)$/;

/** Block content anywhere under the cell, seen through container wrappers. */
function cellHasBlockContent(nodes: AnyNode[], styleResolver: StyleResolver): boolean {
  for (const node of nodes) {
    if (node.type !== "tag") continue;
    const el = node as Element;
    const tag = el.name.toLowerCase();
    if (/^(?:p|h[1-6]|ul|ol|table)$/.test(tag)) return true;
    if (colorBarCss(node, styleResolver) !== null) return true;
    if (CELL_CONTAINER_TAG.test(tag) && cellHasBlockContent(el.children ?? [], styleResolver)) {
      return true;
    }
  }
  return false;
}

function cellBlocks(
  $: CheerioAPI,
  cell: ParsedCell,
  columnIndex: number,
  columnWidths: number[],
  styleResolver: StyleResolver,
  cellPadding?: number,
): (Paragraph | Table)[] {
  const nodes = cell.element.children ?? [];
  if (!cellHasBlockContent(nodes, styleResolver)) {
    return [cellParagraph(cell, styleResolver)];
  }

  const cellWidth = sumColumnWidths(columnWidths, columnIndex, cell.colspan);
  const contentWidth = Math.max(0, cellWidth - (cellPadding ?? 0) * 2);

  const blocks: (Paragraph | Table)[] = [];
  let pending: AnyNode[] = [];
  const flushInline = (): void => {
    if (pending.some(nodeHasInlineContent)) {
      blocks.push(cellParagraph(cell, styleResolver, pending));
    }
    pending = [];
  };

  const walk = (children: AnyNode[]): void => {
    for (const node of children) {
      if (node.type === "tag") {
        const el = node as Element;
        if (isHiddenElement(el, styleResolver)) continue;
        const tag = el.name.toLowerCase();
        if (tag === "table") {
          flushInline();
          blocks.push(convertTable($, el, styleResolver, contentWidth, true));
          continue;
        }
        if (/^(?:p|h[1-6])$/.test(tag)) {
          flushInline();
          blocks.push(cellBlockParagraph(cell, styleResolver, el));
          continue;
        }
        if (tag === "ul" || tag === "ol") {
          flushInline();
          blocks.push(
            ...processList($, el, { ...DEFAULT_VISITOR_CONTEXT, styleResolver }),
          );
          continue;
        }
        const barCss = colorBarCss(node, styleResolver);
        if (barCss) {
          flushInline();
          blocks.push(barParagraph(barCss, contentWidth));
          continue;
        }
        if (CELL_CONTAINER_TAG.test(tag)) {
          // A wrapper div is a block boundary: end any open inline run, then
          // process its children at this level so nested paragraphs/tables stay
          // real blocks instead of flattening into one run-on paragraph.
          flushInline();
          walk(el.children ?? []);
          flushInline();
          continue;
        }
      }
      pending.push(node);
    }
  };

  walk(nodes);
  flushInline();

  return blocks.length > 0 ? blocks : [cellParagraph(cell, styleResolver)];
}

/** Explicit style borders on the cell itself (e.g. `border-right:1px solid #e2e8f0`). */
function cellStyleBorders(cell: ParsedCell, styleResolver: StyleResolver) {
  // Computed styles report UA-default cell borders (`border="1"` tables) as
  // concrete near-black values on EVERY cell — the table-level grid already
  // paints those edges. Only honor borders the author declared on the cell.
  if (styleResolver.source === "computed") {
    const inline = cell.element.attribs?.style ?? "";
    if (!/\bborder(?:-top|-right|-bottom|-left)?\s*:/i.test(inline)) return undefined;
  }
  const css = styleResolver.getCss(cell.element);
  const toSide = (b: ParsedBorder | undefined) =>
    b
      ? {
          style: BorderStyle.SINGLE,
          size: Math.max(2, Math.round(b.widthPx * 6)),
          color: b.color ?? css.borderColor ?? "000000",
        }
      : undefined;
  const sides = {
    top: toSide(css.borderTop ?? css.border),
    right: toSide(css.borderRight ?? css.border),
    bottom: toSide(css.borderBottom ?? css.border),
    left: toSide(css.borderLeft ?? css.border),
  };
  const declared = Object.fromEntries(
    Object.entries(sides).filter(([, side]) => side !== undefined),
  );
  return Object.keys(declared).length > 0 ? declared : undefined;
}

function buildTableCell(
  $: CheerioAPI,
  cell: ParsedCell,
  columnIndex: number,
  columnWidths: number[],
  styleResolver: StyleResolver,
  cellPadding?: number,
): TableCell {
  const span = cell.colspan;
  const borders = cellStyleBorders(cell, styleResolver);
  return new TableCell({
    columnSpan: span > 1 ? span : undefined,
    rowSpan: cell.rowspan > 1 ? cell.rowspan : undefined,
    width: {
      size: sumColumnWidths(columnWidths, columnIndex, span),
      type: WidthType.DXA,
    },
    margins: cellPadding
      ? { top: cellPadding, bottom: cellPadding, left: cellPadding, right: cellPadding }
      : undefined,
    shading: cellShading(cell, styleResolver),
    ...(borders ? { borders } : {}),
    children: cellBlocks($, cell, columnIndex, columnWidths, styleResolver, cellPadding),
  });
}

interface PaddingCellEdges {
  bordered: boolean;
  gridColor: string;
  isFirstRow: boolean;
  isLastRow: boolean;
  isLastColumn: boolean;
}

function buildPaddingCell(
  columnIndex: number,
  columnWidths: number[],
  cellPadding: number | undefined,
  edges: PaddingCellEdges,
): TableCell {
  // Browsers leave missing-cell regions as borderless "holes" while still
  // painting the table's outer frame. Suppress the grid inside the hole and
  // keep only the frame edges this cell sits on.
  const none = { style: BorderStyle.NONE, size: 0, color: "auto" };
  const frame = { style: BorderStyle.SINGLE, size: BORDER_SIZE, color: edges.gridColor };
  return new TableCell({
    width: { size: columnWidths[columnIndex]!, type: WidthType.DXA },
    margins: cellPadding
      ? { top: cellPadding, bottom: cellPadding, left: cellPadding, right: cellPadding }
      : undefined,
    ...(edges.bordered
      ? {
          borders: {
            top: edges.isFirstRow ? frame : none,
            bottom: edges.isLastRow ? frame : none,
            left: none,
            right: edges.isLastColumn ? frame : none,
          },
        }
      : {}),
    children: [new Paragraph({ children: [new TextRun("")] })],
  });
}

/** Pass 2 — emit docx rows aligned to the grid matrix with explicit spans and widths. */
/** Cell with nothing renderable — no text, no table/list/bar/image anywhere under it. */
function isRenderableEmptyCell(cell: ParsedCell, styleResolver: StyleResolver): boolean {
  if (elementPlainText(cell.element).length > 0) return false;
  if (cellHasBlockContent(cell.element.children ?? [], styleResolver)) return false;
  const hasImage = (nodes: AnyNode[]): boolean =>
    nodes.some(
      (node) =>
        node.type === "tag" &&
        ((node as Element).name.toLowerCase() === "img" ||
          hasImage((node as Element).children ?? [])),
    );
  return !hasImage(cell.element.children ?? []);
}

/** Explicit row height: max of the tr's and its cells' `height` (CSS or attr), in twips. */
function explicitRowHeightTwips(
  row: PlacedCell[],
  styleResolver: StyleResolver,
): number | undefined {
  const candidates: number[] = [];
  const collect = (element: Element): void => {
    // Computed styles report a concrete USED height for every row — only
    // author-declared heights count (same rule as cellStyleBorders).
    const declaredInline = /\bheight\s*:/i.test(element.attribs?.style ?? "");
    const css = styleResolver.getCss(element);
    if (css.heightTwips && (styleResolver.source !== "computed" || declaredInline)) {
      candidates.push(css.heightTwips);
    }
    const attr = element.attribs?.height?.trim();
    if (attr && !attr.endsWith("%")) {
      const px = parseFloat(attr);
      if (Number.isFinite(px) && px > 0) candidates.push(pxToTwips(px));
    }
  };
  const tr = row[0]?.cell.row;
  if (tr) collect(tr);
  for (const { cell } of row) collect(cell.element);
  return candidates.length > 0 ? Math.max(...candidates) : undefined;
}

/**
 * Row height plan (legacy/email HTML controls vertical rhythm with spacer rows):
 * - Spacer rows (nothing renderable in any cell) collapse to their declared
 *   height — or to just the cell padding, matching browsers, instead of the
 *   full default line box that inflated every divider/gap row (~20px each).
 * - Content rows with a declared height get AT_LEAST — HTML treats tr/td
 *   height as a minimum; content may grow past it.
 */
function rowHeightOptions(
  row: PlacedCell[],
  styleResolver: StyleResolver,
  cellPadding: number | undefined,
): { value: number; rule: (typeof HeightRule)[keyof typeof HeightRule] } | undefined {
  const explicit = explicitRowHeightTwips(row, styleResolver);
  const isSpacerRow =
    row.length === 0 || row.every(({ cell }) => isRenderableEmptyCell(cell, styleResolver));
  if (isSpacerRow) {
    const padding = (cellPadding ?? 0) * 2;
    return { value: explicit ?? Math.max(padding, 20), rule: HeightRule.EXACT };
  }
  if (explicit) return { value: explicit, rule: HeightRule.ATLEAST };
  return undefined;
}

function buildTableRows(
  $: CheerioAPI,
  analysis: GridAnalysis,
  styleResolver: StyleResolver,
  cellPadding: number | undefined,
  bordered: boolean,
  gridColor: string,
): TableRow[] {
  const { maxColumns, placedRows, columnWidths } = analysis;

  // Columns covered by a rowspan from an earlier row: the docx library emits the
  // vertical-merge continuation cells itself, so those columns must stay empty.
  const occupied: Array<Set<number>> = placedRows.map(() => new Set<number>());
  placedRows.forEach((row, rowIndex) => {
    for (const { cell, columnIndex } of row) {
      const lastRow = Math.min(rowIndex + cell.rowspan, placedRows.length);
      for (let r = rowIndex + 1; r < lastRow; r++) {
        for (let i = 0; i < cell.colspan; i++) occupied[r]!.add(columnIndex + i);
      }
    }
  });

  return placedRows.map((row, rowIndex) => {
    const byColumn = new Map(row.map((p) => [p.columnIndex, p.cell]));
    const docxCells: TableCell[] = [];
    let columnIndex = 0;

    while (columnIndex < maxColumns) {
      const cell = byColumn.get(columnIndex);
      if (cell) {
        docxCells.push(buildTableCell($, cell, columnIndex, columnWidths, styleResolver, cellPadding));
        columnIndex += cell.colspan;
        continue;
      }
      if (occupied[rowIndex]!.has(columnIndex)) {
        columnIndex += 1;
        continue;
      }
      docxCells.push(
        buildPaddingCell(columnIndex, columnWidths, cellPadding, {
          bordered,
          gridColor,
          isFirstRow: rowIndex === 0,
          isLastRow: rowIndex === placedRows.length - 1,
          isLastColumn: columnIndex === maxColumns - 1,
        }),
      );
      columnIndex += 1;
    }

    const height = rowHeightOptions(row, styleResolver, cellPadding);
    return new TableRow({ children: docxCells, ...(height ? { height } : {}) });
  });
}

export function convertTable(
  $: CheerioAPI,
  table: Element,
  styleResolver: StyleResolver = INLINE_STYLE_RESOLVER,
  contentWidthTwips: number = CONTENT_WIDTH_TWIPS,
  fillParent = false,
): Table {
  const trElements = collectRowElements($, table);
  const parsedRows = parseRows($, trElements, table, styleResolver);
  const cellPadding = parseCellPadding(table);
  const declared = tableDeclaredWidth(table, styleResolver);
  const declaredTwips =
    declared.percent !== undefined
      ? Math.round((declared.percent / 100) * contentWidthTwips)
      : declared.twips;
  const colWidths = colgroupColumnWidths($, table, contentWidthTwips);
  const analysis = analyzeGrid(
    parsedRows,
    styleResolver,
    cellPadding,
    contentWidthTwips,
    declaredTwips,
    fillParent,
    colWidths,
  );
  const plan = tableBorderPlan(table, styleResolver);
  const alignment = tableAlignment(table);

  return new Table({
    width:
      declared.percent !== undefined
        ? { size: declared.percent, type: WidthType.PERCENTAGE }
        : { size: analysis.totalWidth, type: WidthType.DXA },
    // Positioning only matters when the table is narrower than its container.
    ...(alignment && analysis.totalWidth < contentWidthTwips ? { alignment } : {}),
    columnWidths: analysis.columnWidths,
    layout: "fixed",
    // `borders: undefined` means docx-library defaults (a black grid) — always
    // pass an explicit plan (grid, frame-only, or all-NONE).
    borders: plan.borders,
    rows: buildTableRows($, analysis, styleResolver, cellPadding, plan.grid, plan.gridColor),
  });
}

export function convertTableBlock(
  $: CheerioAPI,
  table: Element,
  styleResolver: StyleResolver = INLINE_STYLE_RESOLVER,
): DocxBlock {
  return convertTable($, table, styleResolver);
}
