import type { Element } from "domhandler";
import {
  parseColor,
  parseFontSize,
  parseGap,
  parseInlineStyle,
  presentationalAttributesCss,
  pxToTwips,
  type ParsedBorder,
  type ParsedCss,
} from "./css.js";
import { elementStylePath } from "./style-path.js";
import { HEADING_FONT_HALF_POINTS, HEADING_MARGIN_EM } from "./constants.js";
import type { ComputedStyleSnapshot } from "./computed-style-snapshot.js";

export type { ComputedStyleSnapshot } from "./computed-style-snapshot.js";
export { snapshotComputedStylesFromDocument } from "./computed-style-snapshot.js";

export type StyleSource = "inline" | "computed";

export interface StyleResolver {
  readonly source: StyleSource;
  getCss(element: Element): ParsedCss;
}

function parseLengthPx(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const trimmed = value.trim();
  if (!trimmed || trimmed === "auto" || trimmed === "normal") return undefined;
  if (trimmed.endsWith("px")) return pxToTwips(parseFloat(trimmed));
  if (trimmed.endsWith("pt")) return Math.round(parseFloat(trimmed) * 20);
  if (trimmed.endsWith("em")) return pxToTwips(parseFloat(trimmed) * 16);
  const num = parseFloat(trimmed);
  return Number.isFinite(num) ? pxToTwips(num) : undefined;
}

function computedBorderSide(width: string, color: string): ParsedBorder | undefined {
  const w = parseFloat(width);
  if (!Number.isFinite(w) || w <= 0) return undefined;
  return { widthPx: w, color: parseColor(color) };
}

export function parsedCssFromComputedRecord(raw: Record<string, string>): ParsedCss {
  const css: ParsedCss = {};

  css.color = parseColor(raw.color);
  css.backgroundColor = parseColor(raw.backgroundColor);
  css.display = raw.display?.trim().toLowerCase() || undefined;
  css.flexDirection = raw.flexDirection?.trim().toLowerCase() || undefined;
  css.gap = parseGap(raw.gap) ?? parseGap(raw.columnGap) ?? parseGap(raw.rowGap);
  css.textAlign = raw.textAlign?.trim().toLowerCase() || undefined;
  css.fontSize = parseFontSize(raw.fontSize);
  css.fontWeight = raw.fontWeight?.trim() || undefined;
  css.fontStyle = raw.fontStyle?.trim().toLowerCase() || undefined;

  css.marginTop = parseLengthPx(raw.marginTop);
  css.marginRight = parseLengthPx(raw.marginRight);
  css.marginBottom = parseLengthPx(raw.marginBottom);
  css.marginLeft = parseLengthPx(raw.marginLeft);
  css.paddingTop = parseLengthPx(raw.paddingTop);
  css.paddingRight = parseLengthPx(raw.paddingRight);
  css.paddingBottom = parseLengthPx(raw.paddingBottom);
  css.paddingLeft = parseLengthPx(raw.paddingLeft);

  // getComputedStyle resolves percentages to concrete px, so width lands in widthTwips.
  css.heightTwips = parseLengthPx(raw.height);
  const rawWidth = raw.width?.trim();
  if (rawWidth?.endsWith("%")) css.widthPercent = parseFloat(rawWidth);
  else css.widthTwips = parseLengthPx(raw.width);
  css.maxWidthTwips = parseLengthPx(raw.maxWidth);

  css.borderTop = computedBorderSide(raw.borderTopWidth ?? "", raw.borderTopColor ?? "");
  css.borderRight = computedBorderSide(raw.borderRightWidth ?? "", raw.borderRightColor ?? "");
  css.borderBottom = computedBorderSide(raw.borderBottomWidth ?? "", raw.borderBottomColor ?? "");
  css.borderLeft = computedBorderSide(raw.borderLeftWidth ?? "", raw.borderLeftColor ?? "");

  return css;
}

export class InlineStyleResolver implements StyleResolver {
  readonly source = "inline" as const;

  getCss(element: Element): ParsedCss {
    const attrCss = presentationalAttributesCss(element);
    const styleCss = parseInlineStyle(element.attribs?.style);
    // Inline style wins over presentational attributes (browser precedence);
    // parsed style objects only carry keys that were actually declared.
    return { ...attrCss, ...styleCss };
  }
}

/** Strip UA defaults that inline style="" resolution never sees (headings, blockquotes). */
function approxEquals(value: number | undefined, expected: number, tolerance: number): boolean {
  return value !== undefined && Math.abs(value - expected) <= tolerance;
}

function normalizeComputedUACss(element: Element, css: ParsedCss): ParsedCss {
  const tag = element.name.toLowerCase();
  const inlineStyle = element.attribs?.style ?? "";

  if (/^h[1-6]$/.test(tag)) {
    // Strip heading values ONLY when they match the UA default — a stylesheet
    // `.banner h1 { font-size: 20px }` must survive; blanket-stripping (the old
    // behavior) threw author CSS away along with the UA noise.
    const result = { ...css };
    const headingTag = tag as keyof typeof HEADING_FONT_HALF_POINTS;
    const uaFontHalfPoints = HEADING_FONT_HALF_POINTS[headingTag];
    if (
      !/\bfont-size\s*:/i.test(inlineStyle) &&
      approxEquals(css.fontSize, uaFontHalfPoints, 1)
    ) {
      result.fontSize = undefined;
    }
    if (!/\bmargin(?:\s|:|-)/i.test(inlineStyle)) {
      const headingFontPx = (css.fontSize ?? uaFontHalfPoints) / 1.5;
      const uaMarginTwips = Math.round(HEADING_MARGIN_EM[headingTag] * headingFontPx * 15);
      if (approxEquals(css.marginTop, uaMarginTwips, 15)) result.marginTop = undefined;
      if (approxEquals(css.marginBottom, uaMarginTwips, 15)) result.marginBottom = undefined;
      if (approxEquals(css.marginLeft, 0, 1)) result.marginLeft = undefined;
      if (approxEquals(css.marginRight, 0, 1)) result.marginRight = undefined;
    }
    if (
      !/\bfont-weight\s*:/i.test(inlineStyle) &&
      (css.fontWeight === "700" || css.fontWeight === "bold")
    ) {
      result.fontWeight = undefined;
    }
    return result;
  }

  return css;
}

export class ComputedStyleResolver implements StyleResolver {
  readonly source = "computed" as const;
  private readonly byPath: Map<string, ParsedCss>;

  constructor(snapshots: ComputedStyleSnapshot[]) {
    this.byPath = new Map(
      snapshots.map((s) => [s.path, parsedCssFromComputedRecord(s.styles)]),
    );
  }

  static fromSnapshots(snapshots: ComputedStyleSnapshot[]): ComputedStyleResolver {
    return new ComputedStyleResolver(snapshots);
  }

  getCss(element: Element): ParsedCss {
    const path = elementStylePath(element);
    const fromComputed = this.byPath.get(path);
    if (fromComputed !== undefined) {
      return normalizeComputedUACss(element, fromComputed);
    }
    // Path miss (e.g. fragment export without root-scoped snapshot) — preserve inline style="".
    return normalizeComputedUACss(element, INLINE_STYLE_RESOLVER.getCss(element));
  }
}

export const INLINE_STYLE_RESOLVER: StyleResolver = new InlineStyleResolver();
