import {
  BorderStyle,
  HeightRule,
  LineRuleType,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  VerticalAlign,
  WidthType,
} from "docx";
import type { AnyNode, Element } from "domhandler";
import { PRINTABLE_CONTENT_WIDTH_TWIPS } from "./constants.js";
import { visibleCellBorders } from "./bordered-block.js";
import { cssToBlockLayout, isHiddenElement, layoutForNativeShadedBlock, pxToTwips, type ParsedCss } from "./css.js";
import type { StyleResolver } from "./style-resolver.js";
import { INLINE_STYLE_RESOLVER } from "./style-resolver.js";
import { estimateTextWidthTwips, minContentWidthTwips } from "./text-metrics.js";
import type { BlockLayout, DocxBlock, FlexLayout } from "./types.js";

const BORDERLESS = { style: BorderStyle.NONE, size: 0, color: "auto" };

export function borderlessTableBorders() {
  return {
    top: BORDERLESS,
    bottom: BORDERLESS,
    left: BORDERLESS,
    right: BORDERLESS,
    insideHorizontal: BORDERLESS,
    insideVertical: BORDERLESS,
  };
}

export interface FlexItemContent {
  layout: BlockLayout;
  blocks: DocxBlock[];
  /** Shrink-wrapped column width when flex items have no flex-grow. */
  intrinsicWidthTwips?: number;
  /**
   * Estimated content height (twips) for text card-style items with block children —
   * used as an AT_LEAST floor on the row so LibreOffice doesn't collapse/inflate tight
   * cards (it sizes rows by natural font metrics). The row still grows past this to fit
   * wrapping/nested content. Omitted for items that wrap raster `<img>` / `<canvas>` so
   * the drawing sizes the row instead.
   */
  contentHeightTwips?: number;
  /** CSS `min-height` (twips) — a floor applied to the flex row as AT_LEAST height. */
  minHeightTwips?: number;
  /** Item has stacked block children (a "card"): content flows from the top, not centered. */
  stackedContent?: boolean;
}

export function elementPlainText(element: Element): string {
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

/** Word/LibreOffice Arial runs slightly wider than our glyph estimate — pad shrink-wrapped flex columns. */
const FLEX_ITEM_WIDTH_SAFETY_TWIPS = pxToTwips(8);

function mediaWidthPx(el: Element): number | undefined {
  const w = el.attribs?.width;
  if (!w) return undefined;
  const n = parseFloat(w);
  return Number.isFinite(n) && n > 0 ? n : undefined;
}

function estimateSubtreeMediaWidthTwips(element: Element, styleResolver: StyleResolver): number {
  let maxTwips = 0;
  const walk = (el: Element): void => {
    const tag = el.name.toLowerCase();
    if (tag === "img" || tag === "canvas") {
      const w = mediaWidthPx(el);
      if (w) maxTwips = Math.max(maxTwips, pxToTwips(w));
    }
    const css = styleResolver.getCss(el);
    if (css.widthTwips) maxTwips = Math.max(maxTwips, css.widthTwips);
    for (const child of el.children ?? []) {
      if (child.type === "tag") walk(child);
    }
  };
  walk(element);
  return maxTwips;
}

export function estimateFlexItemWidthTwips(
  element: Element,
  layout: BlockLayout,
  styleResolver: StyleResolver = INLINE_STYLE_RESOLVER,
): number {
  const text = elementPlainText(element);
  const css = styleResolver.getCss(element);
  const bold = css.fontWeight === "bold" || Number(css.fontWeight) >= 600;
  const padX = (layout.paddingLeft ?? 0) + (layout.paddingRight ?? 0);
  const mediaWidth = estimateSubtreeMediaWidthTwips(element, styleResolver);
  const content = Math.max(
    minContentWidthTwips(32),
    estimateTextWidthTwips(text, { bold, fontSizeHalfPoints: css.fontSize }),
    mediaWidth,
  );
  return padX + content + FLEX_ITEM_WIDTH_SAFETY_TWIPS;
}

function distributeFlexItemWidths(weights: number[], total: number): number[] {
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

function rowColumnPlan(
  items: FlexItemContent[],
  gapTwips: number,
): { columnWidths: number[]; totalWidth: number } {
  const itemCount = items.length;
  const gapTotal = gapTwips * Math.max(0, itemCount - 1);
  const available = Math.max(0, PRINTABLE_CONTENT_WIDTH_TWIPS - gapTotal);
  const weights = items.map((item) =>
    Math.max(minContentWidthTwips(32), item.intrinsicWidthTwips ?? 0),
  );
  // Browsers shrink-wrap row items (no flex-grow parsing) — keep intrinsic
  // widths and only scale down when they would overflow the printable width.
  const wanted = weights.reduce((acc, w) => acc + w, 0);
  const itemWidths =
    wanted > available ? distributeFlexItemWidths(weights, available) : weights;
  const columnWidths: number[] = [];
  let totalWidth = 0;

  itemWidths.forEach((width, index) => {
    columnWidths.push(width);
    totalWidth += width;
    if (index < itemCount - 1 && gapTwips > 0) {
      columnWidths.push(gapTwips);
      totalWidth += gapTwips;
    }
  });

  return { columnWidths, totalWidth };
}

export function parseFlexLayoutFromCss(css: ParsedCss): FlexLayout | null {
  if (css.display !== "flex" && css.display !== "inline-flex") return null;
  return {
    direction: css.flexDirection === "column" ? "column" : "row",
    gap: css.gap,
  };
}

export function parseFlexLayout(
  css: ParsedCss,
): FlexLayout | null {
  return parseFlexLayoutFromCss(css);
}

export function isFlexContainer(
  element: Element,
  styleResolver: StyleResolver = INLINE_STYLE_RESOLVER,
): boolean {
  return parseFlexLayoutFromCss(styleResolver.getCss(element)) !== null;
}

export function flexItemElements(
  element: Element,
  styleResolver: StyleResolver = INLINE_STYLE_RESOLVER,
): Element[] {
  return (element.children ?? []).filter(
    (node): node is Element => node.type === "tag" && !isHiddenElement(node, styleResolver),
  );
}

/** Horizontal flex → single-row borderless table with explicit gap columns. */
export function makeFlexRowTable(
  items: FlexItemContent[],
  gap: number | undefined,
  containerLayout: BlockLayout,
): Table {
  const gapTwips = gap ?? 0;
  const { columnWidths, totalWidth } = rowColumnPlan(items, gapTwips);
  const cells: TableCell[] = [];

  items.forEach((item, index) => {
    const width = columnWidths[cells.length]!;
    cells.push(
      new TableCell({
        width: { size: width, type: WidthType.DXA },
        // Card items (stacked block children) flow from the top like the browser;
        // simple inline items stay vertically centered.
        verticalAlign:
          item.contentHeightTwips || item.stackedContent ? VerticalAlign.TOP : VerticalAlign.CENTER,
        shading: item.layout.shading
          ? { type: ShadingType.CLEAR, ...item.layout.shading }
          : undefined,
        ...(item.layout.borders ? { borders: visibleCellBorders(item.layout.borders) } : {}),
        margins: {
          top: item.layout.paddingTop ?? 0,
          bottom: item.layout.paddingBottom ?? 0,
          left: item.layout.paddingLeft ?? 0,
          right: item.layout.paddingRight ?? 0,
        },
        children: item.blocks.length > 0 ? item.blocks : [new Paragraph({ children: [] })],
      }),
    );

    if (index < items.length - 1 && gapTwips > 0) {
      cells.push(
        new TableCell({
          width: { size: gapTwips, type: WidthType.DXA },
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [new Paragraph({ children: [] })],
        }),
      );
    }
  });

  // Card rows: floor the row at the tallest item's estimated content height so
  // LibreOffice (which sizes rows by natural font metrics) doesn't inflate tight
  // cards — but use AT_LEAST, never EXACT. EXACT clips any content taller than the
  // estimate, and the estimate (one line per direct block child) badly under-counts
  // wrapped headings and nested content: a long heading in a flex wrapper wrapped to
  // two lines and had its second line sliced off. AT_LEAST keeps single-line cards
  // tight (natural height ≈ the estimate) while letting wrapping content grow.
  // Items with raster media omit contentHeightTwips so the drawing sizes the row.
  const maxContentHeight = Math.max(0, ...items.map((it) => it.contentHeightTwips ?? 0));
  const maxMinHeight = Math.max(0, ...items.map((it) => it.minHeightTwips ?? 0));
  const floor = Math.max(maxContentHeight, maxMinHeight);
  const rowOptions: { children: TableCell[]; height?: { value: number; rule: (typeof HeightRule)[keyof typeof HeightRule] } } =
    floor > 0
      ? { children: cells, height: { value: floor, rule: HeightRule.ATLEAST } }
      : { children: cells };

  const inner = new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths,
    layout: "fixed",
    borders: borderlessTableBorders(),
    rows: [new TableRow(rowOptions)],
  });

  return wrapInnerTable(inner, containerLayout);
}

function containerPaddingMargins(layout: BlockLayout) {
  return {
    top: layout.paddingTop ?? 0,
    bottom: layout.paddingBottom ?? 0,
    left: layout.paddingLeft ?? 0,
    right: layout.paddingRight ?? 0,
  };
}

function wrapInnerTable(inner: Table, containerLayout: BlockLayout): Table {
  const needsWrapper =
    Boolean(containerLayout.shading?.fill) ||
    Boolean(
      containerLayout.paddingTop ||
        containerLayout.paddingBottom ||
        containerLayout.paddingLeft ||
        containerLayout.paddingRight,
    );

  if (!needsWrapper) return inner;

  const padLayout = containerLayout.shading
    ? layoutForNativeShadedBlock(containerLayout)
    : containerLayout;

  return new Table({
    width: { size: PRINTABLE_CONTENT_WIDTH_TWIPS, type: WidthType.DXA },
    columnWidths: [PRINTABLE_CONTENT_WIDTH_TWIPS],
    layout: "fixed",
    borders: borderlessTableBorders(),
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: PRINTABLE_CONTENT_WIDTH_TWIPS, type: WidthType.DXA },
            shading: containerLayout.shading
              ? { type: ShadingType.CLEAR, ...containerLayout.shading }
              : undefined,
            margins: containerPaddingMargins(padLayout),
            children: [inner],
          }),
        ],
      }),
    ],
  });
}

/**
 * Unshaded spacer row of exact `gapTwips` height — renders the flex `gap` as a band of
 * the container's own fill (nested cells without shading are transparent, so the wrapper
 * cell's background shows through). Gap-as-cell-margin does NOT work: cell margins are
 * painted with the cell's shading, which visually fuses adjacent shaded items.
 */
function gapSpacerRow(widthTwips: number, gapTwips: number): TableRow {
  return new TableRow({
    height: { value: gapTwips, rule: HeightRule.EXACT },
    children: [
      new TableCell({
        width: { size: widthTwips, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        children: [
          new Paragraph({
            spacing: { before: 0, after: 0, line: gapTwips, lineRule: LineRuleType.EXACT },
            children: [],
          }),
        ],
      }),
    ],
  });
}

/** Vertical flex → multi-row borderless table; gap via unshaded exact-height spacer rows. */
export function makeFlexColumnTable(
  items: FlexItemContent[],
  gap: number | undefined,
  containerLayout: BlockLayout,
): Table {
  const gapTwips = gap ?? 0;
  // The inner table sits inside the wrapper cell, inset by the container's padding —
  // sizing it to the full printable width would overflow the wrapper to the page edge.
  const innerWidth = Math.max(
    0,
    PRINTABLE_CONTENT_WIDTH_TWIPS -
      (containerLayout.paddingLeft ?? 0) -
      (containerLayout.paddingRight ?? 0),
  );

  const rows: TableRow[] = [];
  items.forEach((item, index) => {
    if (index > 0 && gapTwips > 0) {
      rows.push(gapSpacerRow(innerWidth, gapTwips));
    }
    rows.push(
      new TableRow({
        children: [
          new TableCell({
            width: { size: innerWidth, type: WidthType.DXA },
            shading: item.layout.shading
              ? { type: ShadingType.CLEAR, ...item.layout.shading }
              : undefined,
            margins: {
              top: item.layout.paddingTop ?? 0,
              bottom: item.layout.paddingBottom ?? 0,
              left: item.layout.paddingLeft ?? 0,
              right: item.layout.paddingRight ?? 0,
            },
            children: item.blocks.length > 0 ? item.blocks : [new Paragraph({ children: [] })],
          }),
        ],
      }),
    );
  });

  const inner = new Table({
    width: { size: innerWidth, type: WidthType.DXA },
    columnWidths: [innerWidth],
    layout: "fixed",
    borders: borderlessTableBorders(),
    rows,
  });

  return wrapInnerTable(inner, containerLayout);
}
