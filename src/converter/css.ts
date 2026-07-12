import { AlignmentType, BorderStyle, LineRuleType, ShadingType, type IParagraphStylePropertiesOptions } from "docx";
import type { Element } from "domhandler";
import { BLOCK_TAGS, BODY_LINE_BOX_PX } from "./constants.js";
import type { StyleResolver } from "./style-resolver.js";
import { INLINE_STYLE_RESOLVER } from "./style-resolver.js";
import type { BlockBorders, BlockBorderSide, BlockLayout, RunTypography } from "./types.js";

export interface ParsedBorder {
  widthPx: number;
  color?: string;
}

export interface ParsedCss {
  color?: string;
  backgroundColor?: string;
  display?: string;
  flexDirection?: string;
  gap?: number;
  textAlign?: string;
  /** `font-family` mapped to a Word-safe installed font (first recognized family wins). */
  fontFamily?: string;
  /** `border-collapse` keyword — separate (default) tables get UA border-spacing gaps. */
  borderCollapse?: string;
  /** `text-transform` keyword (uppercase → Word all-caps display property). */
  textTransform?: string;
  /** `letter-spacing` in twips (absolute units). */
  letterSpacingTwips?: number;
  /** `letter-spacing` in em (resolved against font size at typography build). */
  letterSpacingEm?: number;
  fontSize?: number;
  fontWeight?: string;
  fontStyle?: string;
  listStyleType?: string;
  marginTop?: number;
  marginRight?: number;
  marginBottom?: number;
  marginLeft?: number;
  paddingTop?: number;
  paddingRight?: number;
  paddingBottom?: number;
  paddingLeft?: number;
  /** Explicit CSS height in twips (px/pt/em) — used for color-bar sizing. */
  heightTwips?: number;
  /** CSS `min-height` in twips — a floor for flex-item box height. */
  minHeightTwips?: number;
  /** Explicit CSS width: absolute in twips, or percentage of the containing block. */
  widthTwips?: number;
  widthPercent?: number;
  maxWidthTwips?: number;
  maxHeightTwips?: number;
  visibility?: string;
  opacity?: number;
  overflow?: string;
  border?: ParsedBorder;
  /** `border-color` — color override for borders declared elsewhere (attr or shorthand). */
  borderColor?: string;
  borderTop?: ParsedBorder;
  borderRight?: ParsedBorder;
  borderBottom?: ParsedBorder;
  borderLeft?: ParsedBorder;
  /** CSS break-before / page-break-before → Word pageBreakBefore. */
  pageBreakBefore?: boolean;
  /** CSS break-after / page-break-after — applied to the next block sibling. */
  pageBreakAfter?: boolean;
}

const PX_TO_TWIPS = 15;
const PX_TO_HALF_POINTS = 1.5;
const PX_TO_BORDER_SIZE = 6;

export function pxToTwips(px: number): number {
  return Math.round(px * PX_TO_TWIPS);
}

export function pxToHalfPoints(px: number): number {
  return Math.round(px * PX_TO_HALF_POINTS);
}

/** Flex gap in twips from px or bare number (treated as px). */
export function parseGap(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const trimmed = value.trim();
  if (trimmed.endsWith("px")) return pxToTwips(parseFloat(trimmed));
  const num = parseFloat(trimmed);
  return Number.isFinite(num) ? pxToTwips(num) : undefined;
}

export function parseColor(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const trimmed = value.trim().toLowerCase();
  if (trimmed === "transparent") return undefined;
  if (trimmed.startsWith("#")) {
    let hex = trimmed.slice(1);
    if (hex.length === 3) {
      hex = hex
        .split("")
        .map((c) => c + c)
        .join("");
    }
    return hex.length === 6 ? hex : undefined;
  }
  if (trimmed.startsWith("rgb")) {
    const match = trimmed.match(
      /rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)(?:\s*,\s*([\d.]+))?\s*\)/,
    );
    if (!match) return undefined;
    const alpha = match[4] !== undefined ? parseFloat(match[4]) : 1;
    if (!Number.isFinite(alpha) || alpha <= 0) return undefined;
    const [, r, g, b] = match;
    return [r, g, b]
      .map((n) => Math.round(Number(n)).toString(16).padStart(2, "0"))
      .join("");
  }
  return undefined;
}

/**
 * CSS font stacks → Word/LibreOffice-safe installed fonts. First recognized
 * family in the stack wins (browser semantics); generic keywords are the
 * fallback. Unrecognized stacks stay on the document default (Arial).
 */
const FONT_FAMILY_MAP: Record<string, string> = {
  georgia: "Georgia",
  "times new roman": "Times New Roman",
  times: "Times New Roman",
  cambria: "Cambria",
  garamond: "Garamond",
  "palatino linotype": "Palatino Linotype",
  palatino: "Palatino Linotype",
  arial: "Arial",
  helvetica: "Arial",
  verdana: "Verdana",
  tahoma: "Tahoma",
  "trebuchet ms": "Trebuchet MS",
  "segoe ui": "Arial",
  "courier new": "Courier New",
  courier: "Courier New",
  // Consolas: native in Word (ships with Office); LibreOffice substitutes its
  // bundled DejaVu Sans Mono. Both are full-weight — Courier New's thin strokes
  // read faint next to a browser's Menlo/Consolas.
  consolas: "Consolas",
  menlo: "Consolas",
  monaco: "Consolas",
  serif: "Times New Roman",
  "sans-serif": "Arial",
  monospace: "Consolas",
};

function mapFontFamily(value: string): string | undefined {
  for (const family of value.split(",")) {
    const key = family.trim().replace(/^['"]|['"]$/g, "").toLowerCase();
    const mapped = FONT_FAMILY_MAP[key];
    if (mapped) return mapped;
  }
  return undefined;
}

function parseLength(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const trimmed = value.trim();
  if (trimmed.endsWith("px")) return pxToTwips(parseFloat(trimmed));
  if (trimmed.endsWith("pt")) return Math.round(parseFloat(trimmed) * 20);
  if (trimmed.endsWith("em")) return pxToTwips(parseFloat(trimmed) * 16);
  const num = parseFloat(trimmed);
  return Number.isFinite(num) ? pxToTwips(num) : undefined;
}

export function parseFontSize(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const trimmed = value.trim();
  if (trimmed.endsWith("px")) return pxToHalfPoints(parseFloat(trimmed));
  if (trimmed.endsWith("pt")) return Math.round(parseFloat(trimmed) * 2);
  if (trimmed.endsWith("em")) return pxToHalfPoints(parseFloat(trimmed) * 16);
  const num = parseFloat(trimmed);
  return Number.isFinite(num) ? pxToHalfPoints(num) : undefined;
}

function parseBorderShorthand(value: string): ParsedBorder | undefined {
  if (/\b(?:none|hidden)\b/i.test(value)) return undefined;
  const widthMatch = value.match(/(\d+(?:\.\d+)?)\s*px/i);
  const widthPx = widthMatch ? parseFloat(widthMatch[1]!) : 1;
  if (widthPx <= 0) return undefined;
  const hexMatch = value.match(/#([0-9a-f]{3,8})/i);
  const rgbMatch = value.match(/rgba?\([^)]+\)/i);
  const colorRaw = hexMatch?.[0] ?? rgbMatch?.[0];
  return {
    widthPx,
    color: parseColor(colorRaw),
  };
}

function applyBoxShorthand(
  value: string,
  result: ParsedCss,
  prefix: "margin" | "padding",
): void {
  const parts = value.split(/\s+/).map(parseLength);
  if (parts.some((p) => p === undefined)) return;

  const [top, right = top, bottom = top, left = right] = parts;
  if (prefix === "margin") {
    result.marginTop = top;
    result.marginRight = right;
    result.marginBottom = bottom;
    result.marginLeft = left;
  } else {
    result.paddingTop = top;
    result.paddingRight = right;
    result.paddingBottom = bottom;
    result.paddingLeft = left;
  }
}

export function parseInlineStyle(style: string | undefined): ParsedCss {
  if (!style) return {};

  const result: ParsedCss = {};
  for (const declaration of style.split(";")) {
    const colon = declaration.indexOf(":");
    if (colon === -1) continue;
    const property = declaration.slice(0, colon).trim().toLowerCase();
    const value = declaration.slice(colon + 1).trim();

    switch (property) {
      case "color":
        result.color = parseColor(value);
        break;
      case "background-color":
      case "background":
        result.backgroundColor = parseColor(value);
        break;
      case "display":
        result.display = value.trim().toLowerCase();
        break;
      case "flex-direction":
        result.flexDirection = value.trim().toLowerCase();
        break;
      case "gap":
      case "row-gap":
      case "column-gap":
        result.gap = parseGap(value);
        break;
      case "text-align":
        result.textAlign = value.toLowerCase();
        break;
      case "text-transform":
        result.textTransform = value.trim().toLowerCase();
        break;
      case "border-collapse":
        result.borderCollapse = value.trim().toLowerCase();
        break;
      case "letter-spacing": {
        const trimmed = value.trim();
        if (trimmed.endsWith("em")) {
          const em = parseFloat(trimmed);
          if (Number.isFinite(em)) result.letterSpacingEm = em;
        } else {
          result.letterSpacingTwips = parseLength(trimmed);
        }
        break;
      }
      case "list-style-type":
      case "list-style":
        // For the shorthand, the type is the last keyword (e.g. "none inside square").
        result.listStyleType = value.trim().toLowerCase().split(/\s+/).pop();
        break;
      case "font-size":
        result.fontSize = parseFontSize(value);
        break;
      case "font-weight":
        result.fontWeight = value;
        break;
      case "font-style":
        result.fontStyle = value;
        break;
      case "font-family":
        result.fontFamily = mapFontFamily(value);
        break;
      case "margin":
        applyBoxShorthand(value, result, "margin");
        break;
      case "margin-top":
        result.marginTop = parseLength(value);
        break;
      case "margin-right":
        result.marginRight = parseLength(value);
        break;
      case "margin-bottom":
        result.marginBottom = parseLength(value);
        break;
      case "margin-left":
        result.marginLeft = parseLength(value);
        break;
      case "padding":
        applyBoxShorthand(value, result, "padding");
        break;
      case "padding-top":
        result.paddingTop = parseLength(value);
        break;
      case "padding-right":
        result.paddingRight = parseLength(value);
        break;
      case "padding-bottom":
        result.paddingBottom = parseLength(value);
        break;
      case "padding-left":
        result.paddingLeft = parseLength(value);
        break;
      case "height":
        result.heightTwips = parseLength(value);
        break;
      case "min-height":
        result.minHeightTwips = parseLength(value);
        break;
      case "width": {
        const percent = value.trim().match(/^(\d+(?:\.\d+)?)\s*%$/);
        if (percent) result.widthPercent = parseFloat(percent[1]!);
        else result.widthTwips = parseLength(value);
        break;
      }
      case "max-width":
        result.maxWidthTwips = parseLength(value);
        break;
      case "max-height":
        result.maxHeightTwips = parseLength(value);
        break;
      case "visibility":
        result.visibility = value.trim().toLowerCase();
        break;
      case "opacity": {
        const opacity = parseFloat(value);
        if (Number.isFinite(opacity)) result.opacity = opacity;
        break;
      }
      case "overflow":
        result.overflow = value.trim().toLowerCase();
        break;
      case "border":
        result.border = parseBorderShorthand(value);
        break;
      case "border-color":
        result.borderColor = parseColor(value.trim().split(/\s+/)[0]);
        break;
      case "border-top":
        result.borderTop = parseBorderShorthand(value);
        break;
      case "border-right":
        result.borderRight = parseBorderShorthand(value);
        break;
      case "border-bottom":
        result.borderBottom = parseBorderShorthand(value);
        break;
      case "border-left":
        result.borderLeft = parseBorderShorthand(value);
        break;
      case "break-before":
      case "page-break-before":
        if (isPageBreakCssValue(value)) result.pageBreakBefore = true;
        break;
      case "break-after":
      case "page-break-after":
        if (isPageBreakCssValue(value)) result.pageBreakAfter = true;
        break;
      default:
        break;
    }
  }
  return result;
}

/** True for CSS break values that start a new page in print layout. */
export function isPageBreakCssValue(value: string): boolean {
  const v = value.trim().toLowerCase();
  return v === "page" || v === "always" || v === "left" || v === "right";
}

export function elementRequestsPageBreakAfter(
  element: Element,
  resolver: StyleResolver = INLINE_STYLE_RESOLVER,
): boolean {
  return Boolean(resolver.getCss(element).pageBreakAfter);
}

/**
 * Element is invisible in a browser — skip it entirely. Covers `display:none`,
 * `visibility:hidden`, `opacity:0`, and the email-preheader idiom
 * (`max-height:0` + `overflow:hidden`).
 */
export function isHiddenCss(css: ParsedCss): boolean {
  if (css.display === "none") return true;
  if (css.visibility === "hidden" || css.visibility === "collapse") return true;
  if (css.opacity !== undefined && css.opacity <= 0) return true;
  if (css.maxHeightTwips !== undefined && css.maxHeightTwips <= 0 && css.overflow === "hidden") {
    return true;
  }
  return false;
}

export function isHiddenElement(
  element: Element,
  resolver: StyleResolver = INLINE_STYLE_RESOLVER,
): boolean {
  return isHiddenCss(resolver.getCss(element));
}

/** HTML `<font size=1..7>` → CSS absolute-size px (Chromium keyword scale). */
const FONT_SIZE_ATTR_PX: Record<number, number> = {
  1: 10,
  2: 13,
  3: 16,
  4: 18,
  5: 24,
  6: 32,
  7: 48,
};

/**
 * Legacy presentational attributes (`bgcolor`, `<font color/size/face>`,
 * `align`, `<center>`) → their CSS equivalents. Inline `style=""` wins over
 * attributes, matching browser precedence.
 */
export function presentationalAttributesCss(element: Element): ParsedCss {
  const attribs = element.attribs ?? {};
  const tag = element.name.toLowerCase();
  const css: ParsedCss = {};

  if (attribs.bgcolor) {
    const color = parseColor(attribs.bgcolor);
    if (color) css.backgroundColor = color;
  }

  if (tag === "font") {
    if (attribs.color) {
      const color = parseColor(attribs.color);
      if (color) css.color = color;
    }
    const sizeAttr = attribs.size?.trim();
    if (sizeAttr) {
      const relative = /^[+-]/.test(sizeAttr);
      const parsed = parseInt(sizeAttr, 10);
      if (Number.isFinite(parsed)) {
        const level = Math.min(7, Math.max(1, relative ? 3 + parsed : parsed));
        css.fontSize = pxToHalfPoints(FONT_SIZE_ATTR_PX[level]!);
      }
    }
    if (attribs.face) css.fontFamily = mapFontFamily(attribs.face);
  }

  // `align` on table means table positioning (handled by the table converter),
  // not text alignment; on images it means float.
  const align = attribs.align?.trim().toLowerCase();
  if (align && tag !== "table" && tag !== "img") {
    if (["left", "center", "right", "justify"].includes(align)) css.textAlign = align;
  }
  if (tag === "center") css.textAlign = "center";

  return css;
}

/** Block if tag is structural block OR display is block / inline-block. */
export function isBlockElement(element: Element, resolver: StyleResolver = INLINE_STYLE_RESOLVER): boolean {
  // `<svg>` is always handled as a block via convertSvg. Computed `getComputedStyle`
  // reports display:inline for inline SVG, which would otherwise skip the svg dispatch.
  if (element.name.toLowerCase() === "svg") return true;
  const css = resolver.getCss(element);
  const display = css.display;

  if (display === "block" || display === "inline-block") return true;
  if (display === "flex") return true;
  if (display === "inline" || display === "inline-flex") return false;

  return BLOCK_TAGS.has(element.name.toLowerCase());
}

function sumTwips(...values: Array<number | undefined>): number | undefined {
  const total = values.reduce<number>((acc, v) => acc + (v ?? 0), 0);
  return total > 0 ? total : undefined;
}

export function mapTextAlign(value: string | undefined): BlockLayout["alignment"] {
  switch (value) {
    case "center":
      return AlignmentType.CENTER;
    case "right":
      return AlignmentType.RIGHT;
    case "justify":
      return AlignmentType.JUSTIFIED;
    default:
      return AlignmentType.LEFT;
  }
}

export function pxPaddingToBorderSpace(paddingTwips: number | undefined): number {
  if (!paddingTwips) return 4;
  return Math.max(1, Math.round(paddingTwips / 20));
}

function borderSideToDocx(
  side: ParsedBorder | undefined,
  paddingTwips: number | undefined,
): BlockBorderSide | undefined {
  if (!side) return undefined;
  return {
    size: Math.max(1, Math.round(side.widthPx * PX_TO_BORDER_SIZE)),
    space: pxPaddingToBorderSpace(paddingTwips),
    color: side.color ?? "000000",
  };
}

function buildBlockBorders(css: ParsedCss): BlockBorders | undefined {
  const fallback = css.border;
  const top = css.borderTop ?? fallback;
  const right = css.borderRight ?? fallback;
  const bottom = css.borderBottom ?? fallback;
  const left = css.borderLeft ?? fallback;

  if (!top && !right && !bottom && !left) return undefined;

  const borders: BlockBorders = {
    top: borderSideToDocx(top, css.paddingTop),
    right: borderSideToDocx(right, css.paddingRight),
    bottom: borderSideToDocx(bottom, css.paddingBottom),
    left: borderSideToDocx(left, css.paddingLeft),
  };

  return Object.values(borders).some(Boolean) ? borders : undefined;
}

/** Block containers: background-color → paragraph shading only. */
export function cssToBlockLayout(css: ParsedCss): BlockLayout {
  return {
    // Leave undefined when text-align is absent so block children can inherit a
    // container's alignment (Word still defaults to left when nothing is set).
    alignment: css.textAlign ? mapTextAlign(css.textAlign) : undefined,
    shading: css.backgroundColor
      ? { fill: css.backgroundColor, color: "auto" }
      : undefined,
    paddingTop: css.paddingTop,
    paddingBottom: css.paddingBottom,
    paddingLeft: css.paddingLeft,
    paddingRight: css.paddingRight,
    marginTop: css.marginTop,
    marginBottom: css.marginBottom,
    marginLeft: css.marginLeft,
    marginRight: css.marginRight,
    spacingBefore: sumTwips(css.paddingTop, css.marginTop),
    spacingAfter: sumTwips(css.paddingBottom, css.marginBottom),
    indentLeft: sumTwips(css.marginLeft, css.paddingLeft),
    indentRight: sumTwips(css.marginRight, css.paddingRight),
    borders: buildBlockBorders(css),
    pageBreakBefore: css.pageBreakBefore || undefined,
    pageBreakAfter: css.pageBreakAfter || undefined,
  };
}

/** Block foreground typography — never carries background-color. */
export function cssToBlockTypography(css: ParsedCss): RunTypography {
  const typography: RunTypography = {
    color: css.color,
    fontSize: css.fontSize,
  };
  if (css.fontWeight === "bold" || Number(css.fontWeight) >= 600) {
    typography.bold = true;
  }
  if (css.fontStyle === "italic") {
    typography.italics = true;
  }
  if (css.fontFamily) {
    typography.font = css.fontFamily;
  }
  if (css.textTransform === "uppercase") {
    typography.allCaps = true;
  }
  const letterSpacing =
    css.letterSpacingTwips ??
    (css.letterSpacingEm !== undefined
      ? Math.round(css.letterSpacingEm * (((css.fontSize ?? 21) / 1.5) * 15))
      : undefined);
  if (letterSpacing) {
    typography.characterSpacing = letterSpacing;
  }
  return typography;
}

/** Inline elements: background-color → TextRun shading only. */
export function cssToInlineRunTypography(css: ParsedCss): RunTypography {
  const typography = cssToBlockTypography(css);
  if (css.backgroundColor) {
    typography.shading = { fill: css.backgroundColor, color: "auto" };
    const hex = css.backgroundColor.replace(/^#/, "").toUpperCase();
    const expanded =
      hex.length === 3 ? hex.split("").map((c) => c + c).join("") : hex;
    if (expanded === "FFFF00") {
      typography.color = typography.color ?? "111111";
    }
  }
  return typography;
}

export function layoutFromElement(
  element: Element,
  resolver: StyleResolver = INLINE_STYLE_RESOLVER,
): BlockLayout {
  return cssToBlockLayout(resolver.getCss(element));
}

/** Nearest `font-family` up the ancestor chain (CSS inheritance), if any. */
export function inheritedFontFamily(
  element: Element,
  resolver: StyleResolver = INLINE_STYLE_RESOLVER,
): string | undefined {
  let node: Element | undefined = element;
  while (node) {
    const font = resolver.getCss(node).fontFamily;
    if (font) return font;
    node =
      node.parent && node.parent.type === "tag" ? (node.parent as Element) : undefined;
  }
  return undefined;
}

export function typographyFromBlockElement(
  element: Element,
  resolver: StyleResolver = INLINE_STYLE_RESOLVER,
): RunTypography {
  const typography = cssToBlockTypography(resolver.getCss(element));
  if (!typography.font) {
    const font = inheritedFontFamily(element, resolver);
    if (font) typography.font = font;
  }
  return typography;
}

/** Native shaded block: keep padding fields for makeParagraph; no indent/spacing bleed. */
export function layoutForNativeShadedBlock(layout: BlockLayout): BlockLayout {
  if (!layout.shading?.fill) return layout;

  return {
    ...layout,
    spacingBefore: undefined,
    spacingAfter: undefined,
    indentLeft: undefined,
    indentRight: undefined,
  };
}

/** Normalize fill hex for consistent LO PDF rendering. */
export function runShadingForFill(fill: string): {
  type: typeof ShadingType.CLEAR;
  fill: string;
  color: string;
} {
  const hex = fill.replace(/^#/, "").toUpperCase();
  const expanded =
    hex.length === 3 ? hex.split("").map((c) => c + c).join("") : hex;
  return {
    type: ShadingType.CLEAR,
    fill: `#${expanded}`,
    color: "auto",
  };
}

/** Extra EXACT line twips — LO PDF band height vs HTML harness (0 = omit for tighter band). */
const SHADED_LINE_FUDGE_TWIPS = 0;

/**
 * Vertical padding inside a shaded paragraph band.
 * Embed all padding in one EXACT line — LO paints spacing.before outside w:shd.
 */
export function shadedBlockParagraphSpacing(layout: BlockLayout): IParagraphStylePropertiesOptions["spacing"] {
  const padTop = layout.paddingTop ?? 0;
  const padBottom = layout.paddingBottom ?? 0;
  const contentLine = pxToTwips(layout.shadedContentLinePx ?? BODY_LINE_BOX_PX);
  return {
    before: 0,
    after: layout.marginBottom ?? 0,
    line: padTop + contentLine + padBottom + SHADED_LINE_FUDGE_TWIPS,
    lineRule: LineRuleType.EXACT,
  };
}

export function blockLayoutToParagraphProps(
  layout: BlockLayout,
): IParagraphStylePropertiesOptions {
  const borderEntries = layout.borders
    ? Object.entries(layout.borders).filter(([, side]) => side !== undefined)
    : [];

  let indent: IParagraphStylePropertiesOptions["indent"];
  if (layout.indentLeft || layout.indentRight) {
    indent = {
      left: layout.indentLeft,
      right: layout.indentRight,
      ...(layout.hangingIndent ? { hanging: layout.hangingIndent } : {}),
    };
  }

  return {
    ...(layout.alignment ? { alignment: layout.alignment } : {}),
    ...(layout.shading
      ? { shading: { type: ShadingType.CLEAR, ...layout.shading } }
      : {}),
    ...(layout.spacingBefore || layout.spacingAfter
      ? {
          spacing: {
            before: layout.spacingBefore,
            after: layout.spacingAfter,
          },
        }
      : {}),
    ...(layout.pageBreakBefore ? { pageBreakBefore: true } : {}),
    ...(indent ? { indent } : {}),
    ...(borderEntries.length > 0
      ? {
          border: Object.fromEntries(
            borderEntries.map(([side, spec]) => [
              side,
              {
                style: BorderStyle.SINGLE,
                size: spec!.size,
                space: spec!.space,
                color: spec!.color,
              },
            ]),
          ),
        }
      : {}),
  };
}

/** @deprecated */
export function cssToBlockStyle(css: ParsedCss): BlockLayout {
  return cssToBlockLayout(css);
}

/** @deprecated */
export function cssToRunTypography(_css: ParsedCss, _tag?: string): RunTypography {
  return cssToBlockTypography(_css);
}

/** @deprecated */
export function cssToRunStyle(css: ParsedCss, tag?: string): RunTypography {
  return cssToRunTypography(css, tag);
}

/** @deprecated */
export function blockStyleToParagraphProps(layout: BlockLayout): IParagraphStylePropertiesOptions {
  return blockLayoutToParagraphProps(layout);
}
