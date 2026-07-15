import * as cheerio from "cheerio";
import type { CheerioAPI } from "cheerio";
import type { Element } from "domhandler";
import { isHiddenElement } from "../src/converter/css.js";

export interface TextContentFidelityResult {
  /** 0–100; 100 = HTML and DOCX display text match (multiset). */
  score: number;
  /**
   * 0–100 order-aware similarity (2·LCS / (lenA+lenB) on condensed text).
   * Report-only — not folded into `score`. The char-bag `score` is order-blind,
   * so reordered blocks/cells score 100 there but show up here.
   */
  orderedSimilarity: number;
  htmlTokenCount: number;
  docxTokenCount: number;
  missingTokenUnits: number;
  extraTokenUnits: number;
  sampleMissing: string[];
  sampleExtra: string[];
}

function isElement(node: unknown): node is Element {
  return typeof node === "object" && node !== null && "name" in node && (node as Element).type === "tag";
}

function cellText($: CheerioAPI, cell: Element): string {
  return $(cell).text().replace(/\s+/g, " ").trim();
}

function tableDisplayText($: CheerioAPI, table: Element): string {
  const rows: string[] = [];
  $(table)
    .find("tr")
    .each((_, tr) => {
      const cells: string[] = [];
      $(tr)
        .children("td, th")
        .each((_, cell) => {
          const text = cellText($, cell);
          if (text) cells.push(text);
        });
      if (cells.length > 0) rows.push(cells.join(" "));
    });
  return rows.join(" ");
}

function blockDisplayText($: CheerioAPI, element: Element): string {
  const tag = element.name.toLowerCase();
  if (tag === "table") return tableDisplayText($, element);

  if (tag === "ol" || tag === "ul") {
    const lines: string[] = [];
    walkList($, element, tag === "ol", lines);
    return lines.join(" ");
  }

  const css = element.attribs?.style ?? "";
  if (/display\s*:\s*flex/i.test(css)) {
    const parts: string[] = [];
    $(element)
      .children()
      .each((_, child) => {
        if (!isElement(child)) return;
        const text = blockDisplayText($, child);
        if (text) parts.push(text);
      });
    return parts.join(" ");
  }

  // Containers with ordered lists anywhere inside (blockquote, div, …) need the
  // synthesized `1.` markers too — DOCX native numbering renders them in the PDF.
  if ($(element).find("ol").length > 0) {
    const parts: string[] = [];
    for (const child of element.children ?? []) {
      if (isElement(child)) {
        const text = blockDisplayText($, child);
        if (text) parts.push(text);
      } else if (child.type === "text") {
        const text = (child.data ?? "").replace(/\s+/g, " ").trim();
        if (text) parts.push(text);
      }
    }
    return parts.join(" ");
  }

  return $(element).text().replace(/\s+/g, " ").trim();
}

function liInlineText($: CheerioAPI, li: Element): string {
  const clone = $(li).clone();
  clone.children("ol, ul").remove();
  return clone.text().replace(/\s+/g, " ").trim();
}

const TYPE_ATTR_TO_STYLE: Record<string, string> = {
  "1": "decimal",
  a: "lower-alpha",
  A: "upper-alpha",
  i: "lower-roman",
  I: "upper-roman",
};

/** CSS `list-style-type` (from the inline style or `type` attr) for a list element. */
function listStyleTypeOf(list: Element): string {
  const style = list.attribs?.style ?? "";
  const m = /list-style(?:-type)?\s*:\s*([a-z-]+)/i.exec(style);
  if (m) return m[1].toLowerCase();
  const typeAttr = list.attribs?.type;
  return (typeAttr && TYPE_ATTR_TO_STYLE[typeAttr]) || "decimal";
}

function toAlpha(n: number, upper: boolean): string {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(97 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return upper ? s.toUpperCase() : s;
}

function toRoman(n: number): string {
  const map: [number, string][] = [
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
    [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
  ];
  let out = "";
  for (const [v, sym] of map) while (n >= v) { out += sym; n -= v; }
  return out;
}

/** Ordered-list marker (`1.`, `a.`, `IV.`, …) matching the list's list-style-type. */
function orderedMarker(styleType: string, index: number): string {
  const n = index + 1;
  switch (styleType) {
    case "lower-alpha":
    case "lower-latin":
      return `${toAlpha(n, false)}.`;
    case "upper-alpha":
    case "upper-latin":
      return `${toAlpha(n, true)}.`;
    case "lower-roman":
      return `${toRoman(n).toLowerCase()}.`;
    case "upper-roman":
      return `${toRoman(n)}.`;
    default:
      return `${n}.`;
  }
}

function walkList($: CheerioAPI, list: Element, ordered: boolean, lines: string[]): void {
  const styleType = ordered ? listStyleTypeOf(list) : "";
  $(list)
    .children("li")
    .each((index, li) => {
      const text = liInlineText($, li);
      if (text) lines.push(ordered ? `${orderedMarker(styleType, index)} ${text}` : text);
      $(li)
        .children("ol, ul")
        .each((_, nested) => {
          walkList($, nested, nested.name.toLowerCase() === "ol", lines);
        });
    });
}

export interface HtmlFragmentDisplayTextOptions {
  /**
   * Drop `<svg>` / `<canvas>` subtrees before extracting expected text.
   * Use when the export path rasterizes chart media to PNG — axis labels and
   * canvas-drawn text remain visible in the image but are not in the PDF text layer.
   */
  excludeRasterizedMediaText?: boolean;
}

/** HTML oracle options for suite cases with custom convert paths (e.g. `rasterizeInPlace`). */
export function htmlDisplayTextOptionsForConvertOptions(
  convertOptions?: { rasterizeInPlace?: boolean | object } | null,
): HtmlFragmentDisplayTextOptions | undefined {
  if (convertOptions?.rasterizeInPlace) {
    return { excludeRasterizedMediaText: true };
  }
  return undefined;
}

/** Build display text with explicit `<ol>` markers — Playwright innerText omits list numbers. */
export function htmlFragmentDisplayText(
  html: string,
  options?: HtmlFragmentDisplayTextOptions,
): string {
  const $ = cheerio.load(`<body>${html.trim()}</body>`, { xml: false });

  // Browsers don't display hidden subtrees (display:none, preheader idioms) or
  // transient overlays (closed dialogs, tooltips) — exclude them from the EXPECTED
  // text, matching the converter's skip. Check every element, not just `[style]`:
  // overlay semantics (`role="tooltip"`, `<dialog>`, `-dialog`/`-tooltip` components)
  // hide without an inline style attribute.
  $("*").each((_, el) => {
    if (el.type === "tag" && isHiddenElement(el)) $(el).remove();
  });

  if (options?.excludeRasterizedMediaText) {
    $("svg, canvas").remove();
  }

  const lines: string[] = [];

  $("body")
    .children()
    .each((_, child) => {
      if (!isElement(child)) return;
      const tag = child.name.toLowerCase();
      if (tag === "ol" || tag === "ul") {
        walkList($, child, tag === "ol", lines);
        return;
      }
      const text = blockDisplayText($, child);
      if (text) lines.push(text);
    });

  return lines.join("\n");
}

/** Count `1.` `2.` … markers in PDF/display text (handles single-line PDF collapse). */
export function countOrderedListLinesInDisplayText(text: string): number {
  const matches = text.match(/\d+\.\s+\S/g);
  return matches ? matches.length : 0;
}

/** Count bullet glyphs in PDF/display text (handles single-line PDF collapse). */
export function countBulletMarkersInDisplayText(text: string): number {
  const matches = text.match(/[•◦‣⁃·▪▫●○◘◙]/g);
  return matches ? matches.length : 0;
}

function normalizeDisplayText(raw: string): string {
  return raw
    .normalize("NFKC")
    .replace(/\u00a0/g, " ")
    .replace(/[·•▪◦‣⁃]/g, " ")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

/** Remove whitespace so PDF cell gaps and mid-word splits compare equal. */
function condenseText(raw: string): string {
  return normalizeDisplayText(raw).replace(/\s+/g, "");
}

function charBag(text: string): Map<string, number> {
  const bag = new Map<string, number>();
  for (const ch of text) {
    bag.set(ch, (bag.get(ch) ?? 0) + 1);
  }
  return bag;
}

/** Cap DP size — page-1 display text is ~1–3k chars; beyond this, truncation noise is negligible. */
const LCS_MAX_CHARS = 6000;

/** Length of the longest common subsequence (two-row DP). */
function lcsLength(a: string, b: string): number {
  if (a.length === 0 || b.length === 0) return 0;
  let prev = new Uint16Array(b.length + 1);
  let curr = new Uint16Array(b.length + 1);

  for (let i = 1; i <= a.length; i++) {
    const ca = a.charCodeAt(i - 1);
    for (let j = 1; j <= b.length; j++) {
      curr[j] = ca === b.charCodeAt(j - 1) ? prev[j - 1] + 1 : Math.max(prev[j], curr[j - 1]);
    }
    [prev, curr] = [curr, prev];
  }

  return prev[b.length];
}

/** 0–100 order-aware similarity on condensed strings: 2·LCS / (lenA+lenB). */
export function orderedSimilarityScore(a: string, b: string): number {
  const ta = a.slice(0, LCS_MAX_CHARS);
  const tb = b.slice(0, LCS_MAX_CHARS);
  if (ta.length === 0 && tb.length === 0) return 100;
  if (ta.length === 0 || tb.length === 0) return 0;
  const lcs = lcsLength(ta, tb);
  return Math.round(((2 * lcs) / (ta.length + tb.length)) * 10000) / 100;
}

function charBagDefectUnits(html: string, docx: string): {
  missing: number;
  extra: number;
  sampleMissing: string[];
  sampleExtra: string[];
} {
  const htmlBag = charBag(html);
  const docxBag = charBag(docx);
  let missing = 0;
  let extra = 0;
  const sampleMissing: string[] = [];
  const sampleExtra: string[] = [];

  for (const [ch, htmlCount] of htmlBag) {
    const docxCount = docxBag.get(ch) ?? 0;
    if (docxCount < htmlCount) {
      missing += htmlCount - docxCount;
      if (sampleMissing.length < 5) sampleMissing.push(ch === " " ? "<space>" : ch);
    }
  }

  for (const [ch, docxCount] of docxBag) {
    const htmlCount = htmlBag.get(ch) ?? 0;
    if (docxCount > htmlCount) {
      extra += docxCount - htmlCount;
      if (sampleExtra.length < 5) sampleExtra.push(ch === " " ? "<space>" : ch);
    }
  }

  return { missing, extra, sampleMissing, sampleExtra };
}

/**
 * Penalize when rendered DOCX text is missing content vs HTML or has duplicated/extra
 * characters (e.g. status labels emitted twice as block shading).
 */
export function textContentFidelityScore(
  htmlDisplayText: string,
  docxDisplayText: string,
): TextContentFidelityResult {
  const htmlCondensed = condenseText(htmlDisplayText);
  const docxCondensed = condenseText(docxDisplayText);

  if (htmlCondensed.length === 0) {
    return {
      score: 100,
      orderedSimilarity: 100,
      htmlTokenCount: 0,
      docxTokenCount: docxCondensed.length,
      missingTokenUnits: 0,
      extraTokenUnits: 0,
      sampleMissing: [],
      sampleExtra: [],
    };
  }

  if (htmlCondensed === docxCondensed) {
    return {
      score: 100,
      orderedSimilarity: 100,
      htmlTokenCount: htmlCondensed.length,
      docxTokenCount: docxCondensed.length,
      missingTokenUnits: 0,
      extraTokenUnits: 0,
      sampleMissing: [],
      sampleExtra: [],
    };
  }

  const { missing, extra, sampleMissing, sampleExtra } = charBagDefectUnits(
    htmlCondensed,
    docxCondensed,
  );
  const defectRatio = (missing + extra) / htmlCondensed.length;
  const score = Math.max(0, 100 * (1 - defectRatio));

  return {
    score: Math.round(score * 100) / 100,
    orderedSimilarity: orderedSimilarityScore(htmlCondensed, docxCondensed),
    htmlTokenCount: htmlCondensed.length,
    docxTokenCount: docxCondensed.length,
    missingTokenUnits: missing,
    extraTokenUnits: extra,
    sampleMissing,
    sampleExtra,
  };
}
