import {
  AlignmentType,
  BorderStyle,
  LineRuleType,
  Paragraph,
  TabStopType,
  TextRun,
  type FileChild,
} from "docx";
import type { CheerioAPI } from "cheerio";
import type { AnyNode, Element } from "domhandler";
import {
  BODY_FONT_HALF_POINTS,
  BODY_LINE_EXACT_TWIPS,
  BODY_LINE_HEIGHT,
  BLOCKQUOTE_MARGIN_PX,
  BLOCKQUOTE_INDENT_PX,
  BLOCKQUOTE_UA_SIDE_MARGIN_PX,
  LIST_HANGING_TWIPS,
  LIST_LEVEL_LEFT_TWIPS,
  DEFAULT_PARAGRAPH_MARGIN_PX,
  HEADING_FONT_HALF_POINTS,
  HEADING_LEVELS,
  HEADING_MARGIN_EM,
  LIST_STYLE_REFERENCES,
} from "./constants.js";
import {
  blockLayoutToParagraphProps,
  cssToBlockLayout,
  isBlockElement,
  isHiddenElement,
  layoutForNativeShadedBlock,
  layoutFromElement,
  pxPaddingToBorderSpace,
  pxToHalfPoints,
  pxToTwips,
  shadedBlockParagraphSpacing,
  typographyFromBlockElement,
} from "./css.js";
import { makeBorderedTableBlock, makeShadedBlockTable, makeShadedContainerTable, needsShadedTableWrapper, shouldUseBorderedTable } from "./bordered-block.js";
import {
  flexItemElements,
  estimateFlexItemWidthTwips,
  isFlexContainer,
  makeFlexColumnTable,
  makeFlexRowTable,
  parseFlexLayoutFromCss,
} from "./flex.js";
import {
  collectInlineRunsFromNodes,
  getDirectBlockChildren,
  getInlineOrTextNodes,
  hasDirectBlockChild,
} from "./inline.js";
import { convertTable } from "./table.js";
import { convertSvg } from "./svg.js";
import type { BlockLayout, DocxBlock, RunTypography, VisitorContext } from "./types.js";
import { DEFAULT_VISITOR_CONTEXT } from "./types.js";
import type { StyleResolver } from "./style-resolver.js";
import { INLINE_STYLE_RESOLVER } from "./style-resolver.js";

function isElement(node: AnyNode): node is Element {
  return node.type === "tag";
}

function mergeBlockLayouts(base: BlockLayout, overlay: BlockLayout): BlockLayout {
  return {
    alignment: overlay.alignment ?? base.alignment,
    shading: overlay.shading ?? base.shading,
    spacingBefore: overlay.spacingBefore ?? base.spacingBefore,
    spacingAfter: overlay.spacingAfter ?? base.spacingAfter,
    indentLeft: sumIndent(base.indentLeft, overlay.indentLeft),
    indentRight: overlay.indentRight ?? base.indentRight,
    borders: overlay.borders ?? base.borders,
  };
}

function sumIndent(base: number | undefined, overlay: number | undefined): number | undefined {
  const total = (base ?? 0) + (overlay ?? 0);
  return total > 0 ? total : undefined;
}

function sumTwips(...values: Array<number | undefined>): number | undefined {
  const total = values.reduce<number>((acc, v) => acc + (v ?? 0), 0);
  return total > 0 ? total : undefined;
}

function nextElementSibling(element: Element): Element | undefined {
  let next = element.next;
  while (next) {
    if (isElement(next)) return next;
    next = next.next;
  }
  return undefined;
}

function prevElementSibling(element: Element): Element | undefined {
  let prev = element.prev;
  while (prev) {
    if (isElement(prev)) return prev;
    prev = prev.prev;
  }
  return undefined;
}

/** Block containers never inherit shading — each block gets its own paragraph shading. */
function applyDefaultParagraphMargins(
  element: Element,
  layout: BlockLayout,
  ctx: VisitorContext,
): BlockLayout {
  if (element.name.toLowerCase() !== "p") return layout;
  // `<p>` inside blockquote — vertical margin lives on the blockquote container.
  if (ctx.blockquoteDepth > 0) return layout;
  if (ctx.styleResolver.source === "computed") return layout;

  const css = ctx.styleResolver.getCss(element);
  let defaultMargin = pxToTwips(DEFAULT_PARAGRAPH_MARGIN_PX);
  if (css.fontSize) {
    defaultMargin = pxToTwips(css.fontSize / 1.5);
  }
  const result = { ...layout };

  if (css.marginBottom === undefined) {
    result.marginBottom = (result.marginBottom ?? 0) + defaultMargin;
  }
  if (css.marginTop === undefined) {
    // Collapse with previous `<p>` margin-bottom (HTML max-margin behavior).
    const prev = prevElementSibling(element);
    if (!(prev && isParagraphElement(prev))) {
      result.marginTop = (result.marginTop ?? 0) + defaultMargin;
    }
  }

  return {
    ...result,
    spacingBefore: sumTwips(result.paddingTop, result.marginTop),
    spacingAfter: sumTwips(result.paddingBottom, result.marginBottom),
  };
}

function collapseMarginTopAfterBlock(
  element: Element,
  layout: BlockLayout,
  ctx: VisitorContext,
): BlockLayout {
  if (ctx.styleResolver.source === "computed") return layout;
  if (!layout.marginTop) return layout;

  const tag = element.name.toLowerCase();
  if (tag !== "ul" && tag !== "ol" && tag !== "div" && tag !== "section" && tag !== "blockquote") {
    return layout;
  }

  const prev = prevElementSibling(element);
  if (!prev || !isElement(prev)) return layout;
  const prevTag = prev.name.toLowerCase();
  if (prevTag === "p" || /^h[1-6]$/.test(prevTag) || prevTag === "blockquote") {
    return {
      ...layout,
      marginTop: 0,
      spacingBefore: layout.paddingTop,
    };
  }
  return layout;
}

function collapseMarginTopAfterParagraph(
  element: Element,
  layout: BlockLayout,
  ctx: VisitorContext,
): BlockLayout {
  if (ctx.styleResolver.source === "computed") return layout;
  if (element.name.toLowerCase() !== "p") return layout;
  if (!layout.marginTop) return layout;

  const prev = prevElementSibling(element);
  if (!prev || !isElement(prev)) return layout;
  const prevTag = prev.name.toLowerCase();
  if (prevTag === "p" || /^h[1-6]$/.test(prevTag) || prevTag === "blockquote") {
    return {
      ...layout,
      marginTop: 0,
      spacingBefore: layout.paddingTop,
    };
  }
  return layout;
}

function resolveBlockLayout(ctx: VisitorContext, element: Element): BlockLayout {
  const local = collapseMarginTopAfterBlock(
    element,
    collapseMarginTopAfterParagraph(
      element,
      applyDefaultParagraphMargins(
        element,
        layoutFromElement(element, ctx.styleResolver),
        ctx,
      ),
      ctx,
    ),
    ctx,
  );
  const inherited = ctx.inheritedLayout ?? {};
  return {
    ...local,
    alignment: local.alignment ?? inherited.alignment,
    indentLeft: sumIndent(inherited.indentLeft, local.indentLeft),
    indentRight: local.indentRight ?? inherited.indentRight,
    borders: local.borders ?? inherited.borders,
  };
}

function emptyRun(size = BODY_FONT_HALF_POINTS): TextRun {
  return new TextRun({ text: "", size });
}

/** Pin spacer paragraphs to ~1pt so LO does not add a full default line box before `w:after`. */
const MARGIN_SPACER_LINE_TWIPS = 20;

function marginSpacer(
  afterTwips: number | undefined,
  carry?: BlockLayout,
): Paragraph | undefined {
  if (!afterTwips) return undefined;
  // Inside a blockquote the spacer carries the quote's left border + indent so
  // the vertical rule stays continuous across margin gaps (matches the browser).
  const carryProps =
    carry && (carry.borders?.left || carry.indentLeft)
      ? blockLayoutToParagraphProps({
          indentLeft: carry.indentLeft,
          borders: carry.borders?.left ? { left: carry.borders.left } : undefined,
        })
      : {};
  // Borders only paint along the line box, so a carried bar needs the gap's
  // full height in the EXACT line (not in w:after) to span it.
  const hasBar = Boolean(carry?.borders?.left);
  return new Paragraph({
    ...carryProps,
    spacing: hasBar
      ? {
          after: 0,
          before: 0,
          line: afterTwips,
          lineRule: LineRuleType.EXACT,
        }
      : {
          after: afterTwips,
          before: 0,
          line: MARGIN_SPACER_LINE_TWIPS,
          lineRule: LineRuleType.EXACT,
        },
    children: [emptyRun()],
  });
}

function isParagraphElement(node: AnyNode): boolean {
  return isElement(node) && node.name.toLowerCase() === "p";
}

function nodesHaveLineBreaks(nodes: AnyNode[]): boolean {
  for (const node of nodes) {
    if (!isElement(node)) continue;
    if (node.name.toLowerCase() === "br") return true;
    if (nodesHaveLineBreaks(node.children ?? [])) return true;
  }
  return false;
}

/** HTML margin collapse with adjacent `<p>` — avoid double-counting in Word spacers. */
function shadedMarginTopSpacer(element: Element | undefined, layout: BlockLayout): Paragraph | undefined {
  if (!layout.marginTop) return undefined;
  const defaultPMargin = pxToTwips(DEFAULT_PARAGRAPH_MARGIN_PX);
  const prev = element ? prevElementSibling(element) : undefined;
  if (prev && isParagraphElement(prev) && layout.marginTop <= defaultPMargin) {
    return undefined;
  }
  return marginSpacer(layout.marginTop);
}

function emitBlockContent(
  runs: ReturnType<typeof collectInlineRunsFromNodes>,
  layout: BlockLayout,
  extra: Record<string, unknown> = {},
  element?: Element,
): DocxBlock[] {
  if (shouldUseBorderedTable(layout)) {
    const blocks: DocxBlock[] = [];
    const topSpacer = marginSpacer(layout.marginTop);
    if (topSpacer) blocks.push(topSpacer);
    blocks.push(makeBorderedTableBlock(runs, layout, extra));
    const bottomSpacer = marginSpacer(layout.marginBottom);
    if (bottomSpacer) blocks.push(bottomSpacer);
    return blocks;
  }

  if (layout.shading?.fill) {
    const blocks: DocxBlock[] = [];
    const topSpacer = shadedMarginTopSpacer(element, layout);
    if (topSpacer) blocks.push(topSpacer);
    blocks.push(
      needsShadedTableWrapper(layout)
        ? makeShadedBlockTable(runs, layoutForNativeShadedBlock(layout))
        : makeParagraph(runs, layoutForNativeShadedBlock(layout), extra),
    );
    return blocks;
  }

  return [makeParagraph(runs, layout, extra)];
}

function makeParagraph(
  children: ReturnType<typeof collectInlineRunsFromNodes>,
  layout: BlockLayout,
  extra: Record<string, unknown> = {},
): Paragraph {
  const isShaded = Boolean(layout.shading?.fill);
  const hasVerticalPad = (layout.paddingTop ?? 0) > 0 || (layout.paddingBottom ?? 0) > 0;
  const padLeft = layout.paddingLeft ?? 0;
  const padRight = layout.paddingRight ?? 0;
  const isListItem = Boolean(extra.numbering);

  const layoutProps = blockLayoutToParagraphProps(
    isShaded
      ? {
          ...layout,
          spacingBefore: undefined,
          spacingAfter: undefined,
          indentLeft: undefined,
          indentRight: padRight > 0 ? padRight : undefined,
        }
      : layout,
  );

  let paragraphChildren =
    children.length > 0 ? children : [emptyRun()];

  const paragraphExtra: Record<string, unknown> = { ...extra };

  if (isShaded && padLeft > 0) {
    paragraphExtra.tabStops = [{ type: TabStopType.LEFT, position: padLeft }];
    paragraphChildren = [
      new TextRun({
        text: "\t",
        size: layout.shadedTabHalfPoints ?? BODY_FONT_HALF_POINTS,
      }),
      ...paragraphChildren,
    ];
  }

  const shadedSpacing =
    isShaded && hasVerticalPad
      ? { spacing: shadedBlockParagraphSpacing(layout) }
      : {};

  const hasLineBreaks = Boolean(extra.hasLineBreaks);
  const isFlexCompact = Boolean(extra.flexCompact);

  const metricSpacing =
    hasLineBreaks || isFlexCompact
      ? {
          spacing: {
            ...(layoutProps.spacing ?? {}),
            line: BODY_LINE_HEIGHT,
            lineRule: LineRuleType.EXACT,
            before: layoutProps.spacing?.before ?? 0,
            after: layoutProps.spacing?.after ?? 0,
          },
        }
      : {};

  const listItemSpacing = isListItem
    ? {
        spacing: {
          before: 0,
          after: 0,
          line: BODY_LINE_EXACT_TWIPS,
          lineRule: LineRuleType.EXACT,
        },
        contextualSpacing: true,
      }
    : {};

  // EXACT per-font line box for flex-item block children (card lines).
  const exactFontLine = typeof extra.exactLineTwips === "number" ? (extra.exactLineTwips as number) : undefined;
  const exactFontSpacing = exactFontLine
    ? {
        spacing: {
          ...(layoutProps.spacing ?? {}),
          line: exactFontLine,
          lineRule: LineRuleType.EXACT,
          before: layoutProps.spacing?.before ?? 0,
          after: layoutProps.spacing?.after ?? 0,
        },
      }
    : {};

  // CSS line-height 1.4 → AT_LEAST font×1.4 (AUTO multiplies Word's natural
  // ~1.15em line — too tall; EXACT clips taller inline content like images).
  // Merged WITH the margin spacing — a separate spread would be overwritten.
  const atLeastLine =
    typeof extra.atLeastLineTwips === "number"
      ? (extra.atLeastLineTwips as number)
      : BODY_LINE_EXACT_TWIPS;
  const defaultLineSpacing = {
    spacing: {
      ...(layoutProps.spacing ?? {}),
      line: atLeastLine,
      lineRule: LineRuleType.AT_LEAST,
    },
  };

  return new Paragraph({
    ...layoutProps,
    ...(isShaded && hasVerticalPad
      ? {}
      : hasLineBreaks || isFlexCompact || isListItem || exactFontLine
        ? {}
        : defaultLineSpacing),
    ...(isShaded && hasVerticalPad
      ? shadedSpacing
      : isListItem
        ? listItemSpacing
        : exactFontLine
          ? exactFontSpacing
          : metricSpacing),
    ...(isShaded ? { contextualSpacing: false } : {}),
    ...paragraphExtra,
    children: paragraphChildren,
  });
}

function blockquoteLayout(
  depth: number,
  element: Element,
  styleResolver: StyleResolver,
  parentIndentTwips = 0,
): BlockLayout {
  const css = styleResolver.getCss(element);
  let local = cssToBlockLayout(css);
  const hasExplicitLeftBorder = Boolean(css.borderLeft ?? css.border);

  let layout: BlockLayout;
  if (styleResolver.source === "computed") {
    const inlineStyle = element.attribs?.style ?? "";
    const hasInlineMargin = /\bmargin(?:\s|:|-)/i.test(inlineStyle);
    if (!hasInlineMargin) {
      local = {
        ...local,
        marginTop: undefined,
        marginBottom: undefined,
        marginLeft: undefined,
        marginRight: undefined,
        spacingBefore: local.paddingTop,
        spacingAfter: local.paddingBottom,
        indentLeft: local.paddingLeft,
        indentRight: local.paddingRight,
      };
    } else {
      local = {
        ...local,
        marginLeft: undefined,
        marginRight: undefined,
        indentLeft: local.paddingLeft,
        indentRight: local.paddingRight,
      };
    }
    layout = {
      ...local,
      indentLeft: sumIndent(local.indentLeft, pxToTwips(BLOCKQUOTE_INDENT_PX * depth)),
      marginTop: sumTwips(local.marginTop, pxToTwips(BLOCKQUOTE_MARGIN_PX)),
      marginBottom: sumTwips(local.marginBottom, pxToTwips(BLOCKQUOTE_MARGIN_PX)),
    };
  } else {
    // UA box model: margin 1em 40px unless the inline style says otherwise.
    // Content indent accumulates parent quote indent + margin + border + padding.
    const borderPx = (css.borderLeft ?? css.border)?.widthPx ?? 0;
    const ownOffset =
      (css.marginLeft ?? pxToTwips(BLOCKQUOTE_UA_SIDE_MARGIN_PX)) +
      pxToTwips(borderPx) +
      (css.paddingLeft ?? 0);
    layout = {
      ...local,
      marginLeft: undefined,
      marginRight: undefined,
      marginTop: css.marginTop ?? pxToTwips(BLOCKQUOTE_MARGIN_PX),
      marginBottom: css.marginBottom ?? pxToTwips(BLOCKQUOTE_MARGIN_PX),
      spacingBefore: local.paddingTop,
      spacingAfter: local.paddingBottom,
      indentLeft: sumIndent(parentIndentTwips, ownOffset),
      indentRight: local.paddingRight,
    };
  }

  if (hasExplicitLeftBorder) {
    const border = local.borders?.left ?? {
      color: depth > 1 ? "666666" : "333333",
      size: 24,
      space: pxPaddingToBorderSpace(css.paddingLeft),
    };
    layout.borders = { ...local.borders, left: border };
  }

  return layout;
}

/** Flush a fresh paragraph on block boundaries; inline backgrounds stay on TextRuns. */
function emitFlowBlocks(
  $: CheerioAPI,
  nodes: AnyNode[],
  blockLayout: BlockLayout,
  inheritedTypography: RunTypography,
  ctx: VisitorContext,
  extra: Record<string, unknown> = {},
  containerElement?: Element,
): DocxBlock[] {
  const blocks: DocxBlock[] = [];
  let pendingInline: AnyNode[] = [];

  const flush = (): void => {
    if (pendingInline.length === 0 && !blockLayout.shading) return;
    const hasLineBreaks = nodesHaveLineBreaks(pendingInline);
    // In a flex-item's block children, size each line box to its own font (CSS
    // line-height 1.4), not AUTO — halfPt/2 → pt × 1.4 × 20 twips = halfPt × 14.
    const exactLineTwips = ctx.flexBlockContent
      ? (inheritedTypography.fontSize ?? ctx.defaultSizeHalfPoints) * 14
      : undefined;
    blocks.push(
      ...emitBlockContent(
        collectInlineRunsFromNodes(pendingInline, inheritedTypography, undefined, ctx.styleResolver, ctx.defaultSizeHalfPoints),
        blockLayout,
        {
          ...extra,
          hasLineBreaks,
          exactLineTwips,
          ...(inheritedTypography.fontSize
            ? { atLeastLineTwips: inheritedTypography.fontSize * 14 }
            : {}),
        },
        containerElement,
      ),
    );
    pendingInline = [];
  };

  for (const node of nodes) {
    if (isElement(node) && isBlockElement(node, ctx.styleResolver)) {
      flush();
      blocks.push(...visitElement($, node, { ...ctx, inheritedLayout: undefined }));
      continue;
    }
    pendingInline.push(node);
  }

  flush();
  return blocks.length > 0 ? blocks : emitBlockContent([], blockLayout, extra, containerElement);
}

/** Tight content height (twips) of a card-style flex item: one EXACT line box per
 *  block child (fontSize × 14 = size × 1.4) plus the item's top/bottom padding. */
function estimateFlexItemContentHeight(
  item: Element,
  itemLayout: BlockLayout,
  ctx: VisitorContext,
): number {
  let lines = 0;
  for (const child of getDirectBlockChildren(item, ctx.styleResolver)) {
    const fontSize = typographyFromBlockElement(child, ctx.styleResolver).fontSize ?? ctx.defaultSizeHalfPoints;
    lines += fontSize * 14;
  }
  return lines + (itemLayout.paddingTop ?? 0) + (itemLayout.paddingBottom ?? 0);
}

function processFlexContainer(
  $: CheerioAPI,
  element: Element,
  ctx: VisitorContext,
): DocxBlock[] {
  const flex = parseFlexLayoutFromCss(ctx.styleResolver.getCss(element));
  if (!flex) return [];

  const containerLayout = resolveBlockLayout(ctx, element);
  const childCtx: VisitorContext = { ...ctx, inheritedLayout: undefined };
  const items = flexItemElements(element).map((item) => {
    const itemCss = ctx.styleResolver.getCss(item);
    const itemLayout = cssToBlockLayout(itemCss);
    const typography = typographyFromBlockElement(item, ctx.styleResolver);
    const intrinsicWidthTwips =
      flex.direction === "row"
        ? estimateFlexItemWidthTwips(item, itemLayout, ctx.styleResolver)
        : undefined;

    if (hasDirectBlockChild(item, ctx.styleResolver)) {
      // Card-style flex items: tighten stacked lines to CSS line boxes (EXACT per font)
      // and force an exact row height, since LibreOffice sizes rows by natural metrics.
      const itemCtx: VisitorContext = { ...childCtx, flexBlockContent: true };
      const contentHeightTwips = estimateFlexItemContentHeight(item, itemLayout, ctx);
      return {
        layout: itemLayout,
        blocks: visitElement($, item, itemCtx),
        intrinsicWidthTwips,
        contentHeightTwips,
      };
    }

    return {
      layout: itemLayout,
      intrinsicWidthTwips,
      blocks: [
        makeParagraph(
          collectInlineRunsFromNodes(item.children ?? [], typography, undefined, ctx.styleResolver, ctx.defaultSizeHalfPoints),
          { alignment: itemLayout.alignment },
          { flexCompact: true },
        ),
      ],
    };
  });

  const table =
    flex.direction === "column"
      ? makeFlexColumnTable(items, flex.gap, containerLayout)
      : makeFlexRowTable(items, flex.gap, containerLayout);

  const wrapped: DocxBlock[] = [];
  const topSpacer = marginSpacer(containerLayout.marginTop);
  if (topSpacer) wrapped.push(topSpacer);
  wrapped.push(table);
  const bottomSpacer = marginSpacer(containerLayout.marginBottom);
  if (bottomSpacer) wrapped.push(bottomSpacer);
  return wrapped;
}

function processBlockContainer(
  $: CheerioAPI,
  element: Element,
  ctx: VisitorContext,
  extra: Record<string, unknown> = {},
): DocxBlock[] {
  if (isFlexContainer(element, ctx.styleResolver)) {
    return processFlexContainer($, element, ctx);
  }

  const layout = resolveBlockLayout(ctx, element);
  const typography = typographyFromBlockElement(element, ctx.styleResolver);
  // text-align inherits (CSS): a container's alignment flows to block children
  // (`<figure>` captions, `<center>`, `<div align=center>` wrappers).
  const inheritAlignment = layout.alignment ? { alignment: layout.alignment } : undefined;
  const childCtx: VisitorContext = { ...ctx, inheritedLayout: inheritAlignment };

  if (hasDirectBlockChild(element, ctx.styleResolver)) {
    const children = getDirectBlockChildren(element, ctx.styleResolver);

    if (layout.shading?.fill) {
      const blocks: DocxBlock[] = [];
      const topSpacer = marginSpacer(layout.marginTop);
      if (topSpacer) blocks.push(topSpacer);

      const childBlocks: DocxBlock[] = [];
      for (const child of children) {
        childBlocks.push(...visitElement($, child, childCtx));
      }
      // Pass the container's own borders (e.g. a callout's border-left) — the
      // cell paints fill AND bar together.
      blocks.push(makeShadedContainerTable(childBlocks, layout, layout.borders));

      const bottomSpacer = marginSpacer(layout.marginBottom);
      if (bottomSpacer) blocks.push(bottomSpacer);
      return blocks;
    }

    // Walk children in order so inline content between block children (e.g. an
    // `<img>` inside a `<figure>` alongside its `<figcaption>`) isn't dropped.
    const blocks: DocxBlock[] = [];
    let pendingInline: AnyNode[] = [];
    const flushInline = (): void => {
      if (pendingInline.length === 0) return;
      const runs = collectInlineRunsFromNodes(pendingInline, typography, undefined, ctx.styleResolver, ctx.defaultSizeHalfPoints);
      if (runs.length > 0) {
        // First emitted block carries the container's top margin (e.g. `<figure margin:8px>`).
        const isFirst = blocks.length === 0;
        const flushLayout: BlockLayout = {
          alignment: layout.alignment,
          spacingBefore: isFirst ? layout.marginTop : undefined,
        };
        blocks.push(...emitBlockContent(runs, flushLayout, {}, element));
      }
      pendingInline = [];
    };
    for (const child of element.children ?? []) {
      if (isElement(child) && isBlockElement(child, ctx.styleResolver)) {
        flushInline();
        blocks.push(...visitElement($, child, childCtx));
      } else {
        pendingInline.push(child);
      }
    }
    flushInline();
    return blocks;
  }

  return emitFlowBlocks($, element.children ?? [], layout, typography, ctx, extra, element);
}

function listContainerMargins(
  listElement: Element,
  styleResolver: StyleResolver,
): { marginTop?: number; marginBottom?: number } {
  const css = styleResolver.getCss(listElement);
  if (styleResolver.source === "computed") {
    return { marginTop: css.marginTop, marginBottom: css.marginBottom };
  }
  // UA: nested lists (`ul ul`, `ol ul`, …) have zero vertical margins.
  const parent = listElement.parent;
  if (parent && isElement(parent)) {
    const parentTag = parent.name.toLowerCase();
    if (parentTag === "li" || parentTag === "ul" || parentTag === "ol") {
      return { marginTop: css.marginTop, marginBottom: css.marginBottom };
    }
  }
  const defaultMargin = pxToTwips(DEFAULT_PARAGRAPH_MARGIN_PX);
  let marginTop: number | undefined = css.marginTop ?? defaultMargin;
  const prev = prevElementSibling(listElement);
  if (prev && isElement(prev) && marginTop) {
    const prevTag = prev.name.toLowerCase();
    if (prevTag === "p" || /^h[1-6]$/.test(prevTag) || prevTag === "blockquote") {
      marginTop = undefined;
    }
  }
  return {
    marginTop,
    marginBottom: css.marginBottom ?? defaultMargin,
  };
}

/** `<ol type="a">` / `<ul type="square">` → CSS list-style-type keyword. */
const LIST_TYPE_ATTR: Record<string, string> = {
  "1": "decimal",
  a: "lower-alpha",
  A: "upper-alpha",
  i: "lower-roman",
  I: "upper-roman",
  disc: "disc",
  circle: "circle",
  square: "square",
};

function resolveListReference(
  listElement: Element,
  tag: string,
  styleResolver: StyleResolver,
): string {
  const styleType =
    styleResolver.getCss(listElement).listStyleType ??
    (listElement.attribs?.type ? LIST_TYPE_ATTR[listElement.attribs.type] : undefined);
  if (styleType && LIST_STYLE_REFERENCES[styleType]) return LIST_STYLE_REFERENCES[styleType];
  return tag === "ol" ? "numbers" : "bullets";
}

/** Exported for table cells — lists inside `<td>` use the same numbering path. */
export function processList(
  $: CheerioAPI,
  listElement: Element,
  ctx: VisitorContext,
): DocxBlock[] {
  const tag = listElement.name.toLowerCase();
  const listRef = resolveListReference(listElement, tag, ctx.styleResolver);
  const blocks: DocxBlock[] = [];
  const { marginTop, marginBottom } = listContainerMargins(listElement, ctx.styleResolver);
  const topSpacer = marginSpacer(
    marginTop,
    ctx.blockquoteDepth > 0 ? ctx.inheritedLayout : undefined,
  );
  if (topSpacer) blocks.push(topSpacer);

  $(listElement)
    .children("li")
    .each((_, li) => {
      blocks.push(...processListItem($, li, { ...ctx, listRef, listLevel: ctx.listLevel }));
    });

  const bottomSpacer = marginSpacer(marginBottom);
  if (bottomSpacer) blocks.push(bottomSpacer);
  return blocks;
}

function listItemParagraphExtra(
  ctx: VisitorContext,
): Record<string, unknown> {
  return {
    style: ctx.listRef?.startsWith("numbers") ? "ListNumber" : "ListBullet",
    numbering: { reference: ctx.listRef ?? "bullets", level: ctx.listLevel },
  };
}

/** OOXML caps paragraph border `w:space` at 31pt. */
const MAX_BORDER_SPACE_PT = 31;

/**
 * A direct `w:ind` on a numbered paragraph overrides the numbering style's indent.
 * When a list item inherits an indent (blockquote), stack the list level indent on
 * top of it, restore the hanging marker column, and pin the tab stop at the text
 * start. The inherited quote bar is re-aimed at the quote content edge (clamped).
 */
function listItemLayout(layout: BlockLayout, ctx: VisitorContext): {
  layout: BlockLayout;
  extra: Record<string, unknown>;
} {
  const extra = listItemParagraphExtra(ctx);
  if (!layout.indentLeft) return { layout, extra };

  const left = layout.indentLeft + LIST_LEVEL_LEFT_TWIPS * (ctx.listLevel + 1);
  const itemLayout: BlockLayout = {
    ...layout,
    indentLeft: left,
    hangingIndent: LIST_HANGING_TWIPS,
  };
  const quoteBar = layout.borders?.left;
  if (quoteBar) {
    // The border anchors at the paragraph's leftmost extent (left − hanging).
    // Aim the bar back at the quote content edge; if the OOXML space cap (31pt)
    // would strand it far from the browser's bar, omit it instead.
    const space =
      Math.round((left - LIST_HANGING_TWIPS - layout.indentLeft) / 20) +
      quoteBar.space;
    if (space <= MAX_BORDER_SPACE_PT) {
      itemLayout.borders = { ...layout.borders, left: { ...quoteBar, space } };
    } else {
      const { left: _left, ...rest } = layout.borders ?? {};
      itemLayout.borders = Object.keys(rest).length > 0 ? rest : undefined;
    }
  }
  extra.tabStops = [{ type: TabStopType.LEFT, position: left }];
  return { layout: itemLayout, extra };
}

function processListItem(
  $: CheerioAPI,
  li: Element,
  ctx: VisitorContext,
): DocxBlock[] {
  const blocks: DocxBlock[] = [];
  const inlineNodes = getInlineOrTextNodes(li, ctx.styleResolver);
  const resolved = resolveBlockLayout(ctx, li);
  const { layout, extra } = listItemLayout(resolved, ctx);
  const typography = typographyFromBlockElement(li, ctx.styleResolver);

  if (inlineNodes.length > 0) {
    blocks.push(
      ...emitBlockContent(
        collectInlineRunsFromNodes(inlineNodes, typography, undefined, ctx.styleResolver, ctx.defaultSizeHalfPoints),
        layout,
        extra,
      ),
    );
  }

  for (const child of li.children ?? []) {
    if (!isElement(child)) continue;
    const tag = child.name.toLowerCase();
    if (tag === "ul" || tag === "ol") {
      blocks.push(...processList($, child, { ...ctx, listLevel: ctx.listLevel + 1 }));
    } else if (isBlockElement(child, ctx.styleResolver)) {
      blocks.push(...visitElement($, child, ctx));
    }
  }

  if (blocks.length === 0) {
    blocks.push(...emitBlockContent([], layout, extra));
  }

  return blocks;
}

function applyDefaultHeadingMargins(
  element: Element,
  layout: BlockLayout,
  fontSizeHalfPoints: number,
  styleResolver: StyleResolver,
): BlockLayout {
  if (styleResolver.source === "computed") return layout;
  const inlineStyle = element.attribs?.style ?? "";
  if (/\bmargin(?:\s|:|-)/i.test(inlineStyle)) return layout;

  const tag = element.name.toLowerCase() as keyof typeof HEADING_MARGIN_EM;
  const factor = HEADING_MARGIN_EM[tag] ?? 1;
  const fontPx = fontSizeHalfPoints / 1.5;
  const margin = pxToTwips(fontPx * factor);

  return {
    ...layout,
    marginTop: layout.marginTop ?? margin,
    marginBottom: layout.marginBottom ?? margin,
  };
}

function processHeading($: CheerioAPI, element: Element, ctx: VisitorContext): DocxBlock[] {
  let layout = resolveBlockLayout(ctx, element);
  const tag = element.name.toLowerCase() as keyof typeof HEADING_LEVELS;
  const fromElement = typographyFromBlockElement(element, ctx.styleResolver);
  const fontSize = fromElement.fontSize ?? HEADING_FONT_HALF_POINTS[tag];
  layout = applyDefaultHeadingMargins(element, layout, fontSize, ctx.styleResolver);
  layout = {
    ...layout,
    spacingBefore: sumTwips(layout.paddingTop, layout.marginTop),
    spacingAfter: sumTwips(layout.paddingBottom, layout.marginBottom),
  };
  const typography: RunTypography = {
    ...fromElement,
    bold: true,
    fontSize,
  };
  const hasBlockShading = Boolean(layout.shading?.fill);
  const fontSizePx = fontSize / 1.5;
  const headingLayout: BlockLayout = hasBlockShading
    ? {
        ...layout,
        shadedContentLinePx: fontSizePx * 1.4,
        shadedTabHalfPoints: fontSize,
      }
    : layout;

  return emitBlockContent(
    collectInlineRunsFromNodes(element.children ?? [], typography, undefined, ctx.styleResolver, ctx.defaultSizeHalfPoints),
    headingLayout,
    hasBlockShading ? {} : { heading: HEADING_LEVELS[tag], atLeastLineTwips: fontSize * 14 },
    element,
  );
}

function processBlockquote($: CheerioAPI, element: Element, ctx: VisitorContext): DocxBlock[] {
  const depth = ctx.blockquoteDepth + 1;
  const parentIndent =
    ctx.blockquoteDepth > 0 ? (ctx.inheritedLayout?.indentLeft ?? 0) : 0;
  const quoteLayout = blockquoteLayout(depth, element, ctx.styleResolver, parentIndent);
  const childCtx: VisitorContext = {
    ...ctx,
    blockquoteDepth: depth,
    inheritedLayout: quoteLayout,
  };

  let blocks: DocxBlock[];
  if (hasDirectBlockChild(element) && quoteLayout.shading?.fill) {
    // Shaded blockquote with block children: one container cell paints fill +
    // bar + padding together (shading is not inheritable, so per-child
    // inheritance would render the bar but lose the background). Children must
    // NOT also inherit the bar/indent — the cell handles both.
    const shadedCtx: VisitorContext = { ...childCtx, inheritedLayout: undefined };
    const childBlocks = visitNodes($, element.children ?? [], shadedCtx);
    blocks = [makeShadedContainerTable(childBlocks, quoteLayout, quoteLayout.borders)];
  } else if (hasDirectBlockChild(element)) {
    blocks = visitNodes($, element.children ?? [], childCtx);
  } else {
    const typography = typographyFromBlockElement(element, ctx.styleResolver);
    blocks = emitBlockContent(
      collectInlineRunsFromNodes(element.children ?? [], typography, undefined, ctx.styleResolver, ctx.defaultSizeHalfPoints),
      quoteLayout,
      {},
      element,
    );
  }

  const wrapped: DocxBlock[] = [];
  const topSpacer = marginSpacer(
    quoteLayout.marginTop,
    ctx.blockquoteDepth > 0 ? ctx.inheritedLayout : undefined,
  );
  if (topSpacer) wrapped.push(topSpacer);
  wrapped.push(...blocks);
  const bottomSpacer = marginSpacer(quoteLayout.marginBottom);
  if (bottomSpacer) wrapped.push(bottomSpacer);
  return wrapped;
}

/** Chromium default monospace font size (`<pre>`/`<code>` with no explicit size). */
const MONOSPACE_DEFAULT_PX = 13;

/** Text content with literal newlines preserved (`white-space: pre`); `<br>` → newline. */
function rawPreText(element: Element): string {
  const parts: string[] = [];
  const walk = (nodes: AnyNode[]): void => {
    for (const node of nodes) {
      if (node.type === "text") parts.push(node.data ?? "");
      else if (isElement(node)) {
        if (node.name.toLowerCase() === "br") parts.push("\n");
        else walk(node.children ?? []);
      }
    }
  };
  walk(element.children ?? []);
  return parts.join("").replace(/^\n+/, "").replace(/\s+$/, "");
}

/**
 * `<pre>` → one run per line with explicit breaks (inline collection would
 * collapse the newlines), sized per the element's font (UA monospace 13px).
 */
function processPreBlock(element: Element, ctx: VisitorContext): DocxBlock[] {
  const css = ctx.styleResolver.getCss(element);
  const layout = resolveBlockLayout(ctx, element);
  const fontSize = css.fontSize ?? pxToHalfPoints(MONOSPACE_DEFAULT_PX);
  // Own font-family only — UA `pre { font-family: monospace }` beats inheritance.
  const font = css.fontFamily ?? "Consolas";
  const runs = rawPreText(element)
    .split("\n")
    .map(
      (line, index) =>
        new TextRun({
          text: line,
          font,
          size: fontSize,
          ...(index > 0 ? { break: 1 } : {}),
        }),
    );
  const preLayout: BlockLayout = {
    ...layout,
    // EXACT per-line box matching the CSS line height of the pre's own font.
    shadedContentLinePx: (fontSize / 1.5) * 1.4,
  };
  return emitBlockContent(runs, preLayout, {}, element);
}

/** Chromium `<hr>`: 1px inset border paints ~2px gray; margin 0.5em collapses with neighbors. */
const HR_DEFAULT_COLOR = "999999";
const HR_DEFAULT_SIZE_EIGHTHS = 8;
const HR_MARGIN_PX = 7;
/** EXACT line height (twips) placing the bottom border at the browser's rule y. */
const HR_LINE_TWIPS = 65;

/**
 * `<hr>` → minimal-height paragraph with a bottom border (the Word idiom for a
 * rule). The EXACT spacer line keeps the border at the browser's y-position
 * instead of hanging below an empty full-height line box.
 */
function processHorizontalRule(element: Element, ctx: VisitorContext): Paragraph {
  const css = ctx.styleResolver.getCss(element);
  const side = css.borderTop ?? css.borderBottom ?? css.border;
  const color = side?.color ?? css.backgroundColor ?? HR_DEFAULT_COLOR;
  const size = side
    ? Math.max(4, Math.round(side.widthPx * 6))
    : HR_DEFAULT_SIZE_EIGHTHS;

  // Adjacent paragraphs/headings/blockquotes already emit collapsed margins
  // (which exceed the hr's own 0.5em); only provide a gap when they don't.
  const prev = prevElementSibling(element);
  const prevTag = prev?.name.toLowerCase() ?? "";
  const prevProvidesGap =
    prevTag === "p" || /^h[1-6]$/.test(prevTag) || prevTag === "blockquote";
  const before = css.marginTop ?? (prevProvidesGap ? 0 : pxToTwips(HR_MARGIN_PX));
  const after = css.marginBottom ?? 0;

  return new Paragraph({
    spacing: {
      before,
      after,
      line: HR_LINE_TWIPS,
      lineRule: LineRuleType.EXACT,
    },
    border: {
      bottom: { style: BorderStyle.SINGLE, size, space: 0, color },
    },
    children: [emptyRun()],
  });
}

export function visitElement(
  $: CheerioAPI,
  element: Element,
  ctx: VisitorContext = DEFAULT_VISITOR_CONTEXT,
): DocxBlock[] {
  if (isHiddenElement(element, ctx.styleResolver)) return [];
  const tag = element.name.toLowerCase();

  switch (tag) {
    case "h1":
    case "h2":
    case "h3":
    case "h4":
    case "h5":
    case "h6":
      return processHeading($, element, ctx);
    case "p":
    case "div":
    case "section":
      return processBlockContainer($, element, ctx);
    case "blockquote":
      return processBlockquote($, element, ctx);
    case "ul":
    case "ol":
      return processList($, element, ctx);
    case "table": {
      // CSS margins on the table element → spacer paragraphs (also keeps
      // adjacent tables from merging into one in Word).
      const tableCss = ctx.styleResolver.getCss(element);
      const blocks: DocxBlock[] = [];
      const topSpacer = marginSpacer(tableCss.marginTop);
      if (topSpacer) blocks.push(topSpacer);
      // `<caption>` renders above the table, centered (UA default).
      const caption = $(element).children("caption").first().toArray()[0];
      if (caption) {
        blocks.push(
          makeParagraph(
            collectInlineRunsFromNodes(
              caption.children ?? [],
              typographyFromBlockElement(caption, ctx.styleResolver),
              undefined,
              ctx.styleResolver,
              ctx.defaultSizeHalfPoints,
            ),
            { alignment: AlignmentType.CENTER },
          ),
        );
      }
      blocks.push(convertTable($, element, ctx.styleResolver));
      const bottomSpacer = marginSpacer(tableCss.marginBottom);
      if (bottomSpacer) blocks.push(bottomSpacer);
      return blocks;
    }
    case "svg":
      return convertSvg(element);
    case "pre":
      return processPreBlock(element, ctx);
    case "hr":
      return [processHorizontalRule(element, ctx)];
    default:
      return processBlockContainer($, element, ctx);
  }
}

export function visitNodes(
  $: CheerioAPI,
  nodes: AnyNode[],
  ctx: VisitorContext = DEFAULT_VISITOR_CONTEXT,
): DocxBlock[] {
  const blocks: DocxBlock[] = [];
  let pendingInline: AnyNode[] = [];

  const flushInline = (): void => {
    if (pendingInline.length === 0) return;
    blocks.push(
      makeParagraph(
        collectInlineRunsFromNodes(pendingInline, {}, undefined, ctx.styleResolver, ctx.defaultSizeHalfPoints),
        ctx.inheritedLayout ?? {},
      ),
    );
    pendingInline = [];
  };

  for (const node of nodes) {
    if (node.type === "text") {
      if ((node.data ?? "").trim()) pendingInline.push(node);
      continue;
    }

    if (!isElement(node)) continue;

    if (isBlockElement(node, ctx.styleResolver)) {
      flushInline();
      const childCtx =
        ctx.blockquoteDepth > 0 || ctx.inheritedLayout
          ? ctx
          : { ...ctx, inheritedLayout: undefined };
      blocks.push(...visitElement($, node, childCtx));
    } else {
      pendingInline.push(node);
    }
  }

  flushInline();
  return blocks;
}

export function htmlToDocxBlocks(
  $: CheerioAPI,
  styleResolver: StyleResolver = INLINE_STYLE_RESOLVER,
  defaultSizeHalfPoints: number = DEFAULT_VISITOR_CONTEXT.defaultSizeHalfPoints,
): FileChild[] {
  const ctx: VisitorContext = { ...DEFAULT_VISITOR_CONTEXT, styleResolver, defaultSizeHalfPoints };
  const nodes = $("body").contents().toArray();
  const blocks = visitNodes($, nodes, ctx);

  if (blocks.length === 0) {
    return [new Paragraph({ children: [new TextRun("")] })];
  }
  return blocks;
}
