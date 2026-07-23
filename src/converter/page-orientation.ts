import * as cheerio from "cheerio";
import type { Element } from "domhandler";
import type { CheerioAPI } from "cheerio";
import { PageOrientation, type FileChild } from "docx";
import { htmlToDocxBlocks } from "./visitor.js";
import type { InlineFieldOptions } from "./fields.js";
import type { StyleResolver } from "./style-resolver.js";

export type PageOrientationName = "portrait" | "landscape";

export interface PageRules {
  defaultOrientation?: PageOrientationName;
  namedOrientations: Record<string, PageOrientationName>;
  classToPage: Record<string, string>;
}

export interface BodySection {
  orientation: PageOrientationName | undefined;
  children: FileChild[];
}

export interface ResolvedPageSize {
  width: number;
  height: number;
}

function toInches(value: number, unit: string): number | null {
  switch (unit.toLowerCase()) {
    case "in":
      return value;
    case "cm":
      return value / 2.54;
    case "mm":
      return value / 25.4;
    case "q":
      return value / 101.6;
    case "pt":
      return value / 72;
    case "pc":
      return value / 6;
    case "px":
      return value / 96;
    default:
      return null;
  }
}

function parseLengthToken(token: string): number | null {
  const m = token.trim().toLowerCase().match(/^([0-9]*\.?[0-9]+)(in|cm|mm|q|pt|pc|px)$/);
  if (!m) return null;
  const numeric = Number(m[1]);
  if (!Number.isFinite(numeric)) return null;
  return toInches(numeric, m[2]!);
}

/** Infer portrait vs landscape from an `@page { size: … }` value. Dimensions only — page size stays on DocumentConfig. */
export function inferOrientationFromPageSizeValue(value: string): PageOrientationName | null {
  const lower = value.trim().toLowerCase();
  if (lower.includes("landscape")) return "landscape";
  if (lower.includes("portrait")) return "portrait";
  const tokens = lower.split(/\s+/).filter(Boolean);
  const lengthTokens = tokens.filter((token) => parseLengthToken(token) !== null);
  if (lengthTokens.length < 2) return null;
  const widthIn = parseLengthToken(lengthTokens[0]!);
  const heightIn = parseLengthToken(lengthTokens[1]!);
  if (widthIn === null || heightIn === null || widthIn === heightIn) return null;
  return widthIn > heightIn ? "landscape" : "portrait";
}

/**
 * Parse `@page` size rules from embedded `<style>` blocks. Class→page selector
 * mappings are parsed only when `includeClassMapping` is true (computed path).
 */
export function parseCssPageRules(html: string, includeClassMapping: boolean): PageRules {
  const namedOrientations: Record<string, PageOrientationName> = {};
  const classToPage: Record<string, string> = {};
  let defaultOrientation: PageOrientationName | undefined;
  const styleBlocks = [...html.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/gi)].map((match) => match[1]!);
  for (const cssText of styleBlocks) {
    for (const match of cssText.matchAll(/@page\s*([^{]*)\{([\s\S]*?)\}/gi)) {
      const selector = (match[1] ?? "").trim();
      const body = match[2] ?? "";
      const sizeDecl = body.match(/(?:^|[;\s])size\s*:\s*([^;}]*)/i);
      if (!sizeDecl?.[1]) continue;
      const inferred = inferOrientationFromPageSizeValue(sizeDecl[1]);
      if (!inferred) continue;
      if (!selector || selector.startsWith(":")) {
        if (!defaultOrientation) defaultOrientation = inferred;
        continue;
      }
      const pageName = selector.split(":")[0]?.trim().toLowerCase();
      if (pageName) namedOrientations[pageName] = inferred;
    }
    if (!includeClassMapping) continue;
    for (const ruleMatch of cssText.matchAll(/([^{}]+)\{([\s\S]*?)\}/g)) {
      const selectorList = (ruleMatch[1] ?? "").trim();
      const body = ruleMatch[2] ?? "";
      if (!selectorList || selectorList.startsWith("@")) continue;
      const pageDecl = body.match(/(?:^|[;\s])page\s*:\s*([^;}]*)/i);
      if (!pageDecl?.[1]) continue;
      const pageTarget = pageDecl[1].trim().replace(/^['"]|['"]$/g, "").toLowerCase();
      if (!pageTarget || pageTarget === "auto") continue;
      for (const selector of selectorList.split(",").map((s) => s.trim().toLowerCase()).filter(Boolean)) {
        const dotIdx = selector.indexOf(".");
        if (dotIdx === -1) continue;
        const classMatch = selector.match(/(?:^|\s|>)(?:[a-z0-9_-]+\.)?([a-z0-9_-]+)/i);
        if (!classMatch?.[1]) continue;
        classToPage[classMatch[1].toLowerCase()] = pageTarget;
      }
    }
  }
  return { defaultOrientation, namedOrientations, classToPage };
}

export function inlinePageName(styleValue: string | undefined): string | undefined {
  if (!styleValue) return undefined;
  const pageDecl = styleValue
    .split(";")
    .map((part) => part.trim())
    .find((part) => /^page\s*:/i.test(part));
  if (!pageDecl) return undefined;
  const [, value = ""] = pageDecl.split(":", 2);
  const page = value.trim().toLowerCase();
  if (!page || page === "auto") return undefined;
  return page;
}

export function pageNameToOrientation(
  pageName: string | undefined,
  rules: PageRules,
): PageOrientationName | undefined {
  if (!pageName) return undefined;
  if (pageName === "portrait" || pageName === "landscape") return pageName;
  return rules.namedOrientations[pageName];
}

function classPageName(classAttr: string | undefined, rules: PageRules): string | undefined {
  if (!classAttr) return undefined;
  for (const className of classAttr.split(/\s+/).map((c) => c.trim().toLowerCase()).filter(Boolean)) {
    const page = rules.classToPage[className];
    if (page) return page;
  }
  return undefined;
}

function orientationForTopLevelNode(
  $: CheerioAPI,
  node: Element,
  baseOrientation: PageOrientationName | undefined,
  pageRules: PageRules,
  allowMixedOrientation: boolean,
  allowClassPageMapping: boolean,
): PageOrientationName | undefined {
  if (!allowMixedOrientation || node.type !== "tag") return baseOrientation;
  const pageName =
    inlinePageName($(node).attr("style")) ??
    (allowClassPageMapping ? classPageName($(node).attr("class"), pageRules) : undefined);
  return pageNameToOrientation(pageName, pageRules) ?? baseOrientation;
}

/**
 * Split top-level body nodes into contiguous same-orientation chunks, convert
 * each chunk, and return one BodySection per chunk (each → one Word section).
 */
export function buildBodySections(
  $: CheerioAPI,
  styleResolver: StyleResolver,
  fontHalfPoints: number,
  baseOrientation: PageOrientationName | undefined,
  pageRules: PageRules,
  allowMixedOrientation: boolean,
  allowClassPageMapping: boolean,
  fieldOptions?: InlineFieldOptions,
): BodySection[] {
  const chunks: Array<{ orientation: PageOrientationName | undefined; htmlParts: string[] }> = [];
  for (const node of $("body").contents().toArray()) {
    if (node.type === "tag" && node.name.toLowerCase() === "style") continue;
    if (node.type === "text" && !(node.data ?? "").trim()) continue;

    if (node.type !== "tag") continue;

    const nodeOrientation = orientationForTopLevelNode(
      $,
      node,
      baseOrientation,
      pageRules,
      allowMixedOrientation,
      allowClassPageMapping,
    );
    const rendered = $.html(node);
    if (!rendered) continue;

    const last = chunks.at(-1);
    if (!last || last.orientation !== nodeOrientation) {
      chunks.push({ orientation: nodeOrientation, htmlParts: [rendered] });
    } else {
      last.htmlParts.push(rendered);
    }
  }

  const sections: BodySection[] = [];
  for (const chunk of chunks) {
    const chunkHtml = chunk.htmlParts.join("").trim();
    if (!chunkHtml) continue;
    const chunk$ = cheerio.load(`<body>${chunkHtml}</body>`, { xml: false });
    const children = htmlToDocxBlocks(chunk$, styleResolver, fontHalfPoints, fieldOptions);
    if (children.length === 0) continue;
    sections.push({ orientation: chunk.orientation, children });
  }
  return sections;
}

/** docx swaps width/height for landscape — pass portrait dims + orient flag. */
export function resolveSectionPageSize(
  pageSize: ResolvedPageSize,
  orientation: PageOrientationName | undefined,
): { width: number; height: number; orientation?: (typeof PageOrientation)[keyof typeof PageOrientation] } {
  return orientation === "landscape"
    ? { width: pageSize.width, height: pageSize.height, orientation: PageOrientation.LANDSCAPE }
    : { width: pageSize.width, height: pageSize.height };
}
