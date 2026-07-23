import {
  Document,
  Footer,
  Header,
  PageBreak,
  Packer,
  Paragraph,
  convertInchesToTwip,
  type FileChild,
} from "docx";
import * as cheerio from "cheerio";
import type { CheerioAPI } from "cheerio";
import { unzipSync, zipSync } from "fflate";
import { BODY_FONT, BODY_FONT_HALF_POINTS, NUMBERING_CONFIG, PAGE_MARGIN_TWIPS } from "./constants.js";
import { patchChromeFieldFiles, patchDocumentXml, patchNumberingXml } from "./ooxml-patch.js";
import { injectFieldTokens, type InlineFieldOptions } from "./fields.js";
import { applyImageResolver, hasUnresolvedSrc, resetImageDocPrIds, type ImageResolver } from "./image.js";
import { INLINE_STYLE_RESOLVER, type StyleResolver } from "./style-resolver.js";
import {
  buildBodySections,
  parseCssPageRules,
  resolveSectionPageSize,
  type BodySection,
  type PageOrientationName,
} from "./page-orientation.js";
import { htmlToDocxBlocks } from "./visitor.js";

/**
 * Diagnostic callback for conditions that don't fail the conversion but silently
 * degrade the output (an image that never embeds, styling that never applies).
 * Pass `null` to suppress the default `console.warn`.
 */
export type WarningHandler = (message: string) => void;

function resolveWarningHandler(onWarning: WarningHandler | null | undefined): WarningHandler {
  if (onWarning === null) return () => {};
  return onWarning ?? ((message) => console.warn(message));
}

/**
 * Flag conditions that produce a structurally valid but visually degraded docx —
 * dropped images, ignored class/stylesheet CSS — so they surface instead of being
 * discovered only by opening the file. Best-effort heuristics, not exhaustive.
 */
function emitDegradationWarnings(
  $: CheerioAPI,
  styleResolver: StyleResolver,
  hadImageResolver: boolean,
  warn: WarningHandler,
): void {
  const unresolvedImages = $("img").toArray().filter(hasUnresolvedSrc).length;
  if (unresolvedImages > 0) {
    warn(
      hadImageResolver
        ? `dom-docx: ${unresolvedImages} image(s) could not be resolved by imageResolver (returned null or threw) — they will render as alt text only.`
        : `dom-docx: ${unresolvedImages} image(s) have a non-data: src (remote/relative URL) and no imageResolver was provided — they will render as alt text only. Pass options.imageResolver to embed them.`,
    );
  }

  if (styleResolver === INLINE_STYLE_RESOLVER) {
    const hasStylesheetSignal = $("style").length > 0 || $('link[rel="stylesheet"]').length > 0;
    if (hasStylesheetSignal) {
      warn(
        'dom-docx: detected a <style> block or external stylesheet, but styleSource is "inline" (the default) — ' +
          "only inline style=\"\" attributes are read, so class-based CSS is ignored. Pass styleSource: \"computed\" " +
          "to resolve it (Node: requires Playwright; browser: uses the live DOM).",
      );
    }
  }
}

/** Page/font/metadata options (Tier 1 `ConvertOptions`). All lengths in inches / points. */
export interface DocumentConfig {
  /** `"letter"` (default), `"a4"`, or a custom size in inches. */
  pageSize?: "letter" | "a4" | { width: number; height: number };
  orientation?: "portrait" | "landscape";
  /** Page margins in inches (each side defaults to 1). */
  margins?: { top?: number; right?: number; bottom?: number; left?: number };
  /** Default body font family and size (points). */
  defaultFont?: { family?: string; sizePt?: number };
  /** Core document properties written to `docProps/core.xml`. */
  metadata?: {
    title?: string;
    subject?: string;
    creator?: string;
    keywords?: string[];
    description?: string;
  };
  /** HTML fragment rendered as the page header (its own inline-styled fragment). */
  headerHtml?: string;
  /** HTML fragment rendered as the page footer. */
  footerHtml?: string;
  /**
   * Append a page-number paragraph to the footer (created if `footerHtml` is absent).
   * - `true` — shorthand for `"Page {page}"`
   * - `string` — template with `{page}` and/or `{pages}` sugar, e.g. `"{page} / {pages}"`
   *
   * Prefer `<span data-docx-field="page">` markers in chrome HTML for full control;
   * `{page}` / `{pages}` lower to the same markers. See allowlist in API.md.
   */
  pageNumber?: boolean | string;
  /** Document language (spell-check locale), e.g. `"en-US"`, `"ar-SA"`. */
  lang?: string;
  /** Text direction; `"rtl"` sets right-to-left paragraphs. */
  direction?: "ltr" | "rtl";
  /**
   * HTML fragment rendered as a cover page: the first content in the document,
   * before the table of contents (if any), followed by an automatic page break so
   * the TOC/body start on the next page. Uses the inline style path like
   * `headerHtml`/`footerHtml` — inline `style="…"` and `data:` images (e.g. a logo)
   * work. When a header/footer/page number is configured, it is suppressed on the
   * cover page (Word "different first page").
   */
  coverHtml?: string;
  /**
   * HTML fragment rendered as a table-of-contents "slot": placed after the cover
   * page (if any) and before the body. You control the markup and styling — a
   * numbered or boxed list, columns, whatever — and in-page links (`<a href="#id">`)
   * jump to the matching `id` in the body (dom-docx bookmarks `id` attributes). Add
   * a trailing `<div style="break-after:page"></div>` if you want it on its own page.
   */
  tocHtml?: string;
}

// Portrait dimensions in twips. Letter matches convertInchesToTwip(8.5)×(11).
const PAGE_PRESETS_TWIPS = {
  letter: { width: 12240, height: 15840 },
  a4: { width: 11906, height: 16838 },
} as const;

interface ResolvedConfig {
  pageSize: { width: number; height: number };
  defaultOrientation: PageOrientationName | undefined;
  margin: { top: number; right: number; bottom: number; left: number };
  font: string;
  fontHalfPoints: number;
  metadata: {
    title?: string;
    subject?: string;
    creator?: string;
    keywords?: string;
    description?: string;
  };
  lang?: string;
  rtl: boolean;
}

function resolveDocumentConfig(
  config: DocumentConfig | undefined,
  inferredOrientation: PageOrientationName | undefined,
): ResolvedConfig {
  const ps = config?.pageSize;
  const base =
    !ps || ps === "letter"
      ? PAGE_PRESETS_TWIPS.letter
      : ps === "a4"
        ? PAGE_PRESETS_TWIPS.a4
        : { width: convertInchesToTwip(ps.width), height: convertInchesToTwip(ps.height) };

  const effectiveOrientation = config?.orientation ?? inferredOrientation;

  const m = config?.margins;
  const marginIn = (v: number | undefined): number =>
    v !== undefined ? convertInchesToTwip(v) : PAGE_MARGIN_TWIPS;

  const meta = config?.metadata;
  const metadata: ResolvedConfig["metadata"] = {};
  if (meta?.title) metadata.title = meta.title;
  if (meta?.subject) metadata.subject = meta.subject;
  if (meta?.creator) metadata.creator = meta.creator;
  if (meta?.keywords?.length) metadata.keywords = meta.keywords.join(", ");
  if (meta?.description) metadata.description = meta.description;

  return {
    pageSize: { width: base.width, height: base.height },
    defaultOrientation: effectiveOrientation,
    margin: { top: marginIn(m?.top), right: marginIn(m?.right), bottom: marginIn(m?.bottom), left: marginIn(m?.left) },
    font: config?.defaultFont?.family ?? BODY_FONT,
    fontHalfPoints:
      config?.defaultFont?.sizePt !== undefined
        ? Math.round(config.defaultFont.sizePt * 2)
        : BODY_FONT_HALF_POINTS,
    metadata,
    lang: config?.lang,
    rtl: config?.direction === "rtl",
  };
}

/** Convert a standalone HTML fragment (header/footer/cover/toc) to DOCX blocks via the inline resolver. */
function fragmentToBlocks(
  html: string,
  sizeHalfPoints: number,
  fieldOptions: InlineFieldOptions,
): FileChild[] {
  const $ = cheerio.load(`<body>${injectFieldTokens(html.trim())}</body>`, { xml: false });
  return htmlToDocxBlocks($, INLINE_STYLE_RESOLVER, sizeHalfPoints, fieldOptions);
}

function chromeFieldOptions(warn: WarningHandler): InlineFieldOptions {
  return { enabled: true, onWarning: warn };
}

/**
 * Caller-provided table-of-contents fragment, placed after the cover (if any) and
 * before the body. Rendered as ordinary body content via the inline style path;
 * in-page links (`<a href="#id">`) resolve to `id` bookmarks in the body. The
 * caller adds a trailing page break in the fragment if they want it on its own
 * page. Returns `[]` when no `tocHtml` is configured.
 */
function buildTocSlot(
  config: DocumentConfig | undefined,
  resolved: ResolvedConfig,
  warn: WarningHandler,
): FileChild[] {
  if (!config?.tocHtml) return [];
  return fragmentToBlocks(config.tocHtml, resolved.fontHalfPoints, chromeFieldOptions(warn));
}

/**
 * Cover fragment as the document's first content, followed by a page break so the
 * TOC/body start on the next page. Converted via the inline style path (like
 * header/footer). Returns `[]` when no `coverHtml` is configured.
 */
function buildCover(
  config: DocumentConfig | undefined,
  resolved: ResolvedConfig,
  warn: WarningHandler,
): FileChild[] {
  if (!config?.coverHtml) return [];
  return [
    ...fragmentToBlocks(config.coverHtml, resolved.fontHalfPoints, chromeFieldOptions(warn)),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

function buildFooter(
  config: DocumentConfig | undefined,
  resolved: ResolvedConfig,
  warn: WarningHandler,
): Footer | undefined {
  const hasFooterHtml = Boolean(config?.footerHtml);
  const pnTemplate =
    config?.pageNumber === true
      ? "Page {page}"
      : typeof config?.pageNumber === "string"
        ? config.pageNumber
        : null;
  if (!hasFooterHtml && pnTemplate === null) return undefined;
  const fieldOpts = chromeFieldOptions(warn);
  const children: FileChild[] = hasFooterHtml
    ? fragmentToBlocks(config!.footerHtml!, resolved.fontHalfPoints, fieldOpts)
    : [];
  if (pnTemplate !== null) {
    const isPlainTemplate = !pnTemplate.trimStart().startsWith("<");
    const html = isPlainTemplate ? `<p style="text-align:center">${pnTemplate}</p>` : pnTemplate;
    children.push(...fragmentToBlocks(html, resolved.fontHalfPoints, fieldOpts));
  }
  return new Footer({ children });
}

function buildHeader(
  config: DocumentConfig | undefined,
  resolved: ResolvedConfig,
  warn: WarningHandler,
): Header | undefined {
  if (!config?.headerHtml) return undefined;
  return new Header({
    children: fragmentToBlocks(config.headerHtml, resolved.fontHalfPoints, chromeFieldOptions(warn)),
  });
}

async function packDocxToUint8Array(
  bodySections: BodySection[],
  resolved: ResolvedConfig,
  chrome: { header?: Header; footer?: Footer },
  coverBlocks: FileChild[],
  tocSlotBlocks: FileChild[],
): Promise<Uint8Array> {
  const listStyleRun = { font: resolved.font, size: resolved.fontHalfPoints };
  const suppressFirstChrome = coverBlocks.length > 0 && Boolean(chrome.header || chrome.footer);
  const nonEmptySections = bodySections.filter((section) => section.children.length > 0);
  const normalizedSections: BodySection[] =
    nonEmptySections.length > 0
      ? nonEmptySections
      : [{ orientation: resolved.defaultOrientation, children: [] }];
  const doc = new Document({
    ...resolved.metadata,
    numbering: NUMBERING_CONFIG,
    styles: {
      default: {
        document: {
          run: {
            font: resolved.font,
            size: resolved.fontHalfPoints,
            ...(resolved.lang ? { language: { value: resolved.lang } } : {}),
            ...(resolved.rtl ? { rightToLeft: true } : {}),
          },
        },
      },
      paragraphStyles: [
        {
          id: "ListNumber",
          name: "List Number",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: listStyleRun,
        },
        {
          id: "ListBullet",
          name: "List Bullet",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: listStyleRun,
        },
      ],
    },
    sections: normalizedSections.map((section, idx) => ({
      properties: {
        page: {
          size: resolveSectionPageSize(resolved.pageSize, section.orientation),
          margin: resolved.margin,
        },
        ...(idx === 0 && suppressFirstChrome ? { titlePage: true } : {}),
      },
      ...(chrome.header
        ? {
            headers:
              idx === 0 && suppressFirstChrome
                ? { default: chrome.header, first: new Header({ children: [] }) }
                : { default: chrome.header },
          }
        : {}),
      ...(chrome.footer
        ? {
            footers:
              idx === 0 && suppressFirstChrome
                ? { default: chrome.footer, first: new Footer({ children: [] }) }
                : { default: chrome.footer },
          }
        : {}),
      children: idx === 0 ? [...coverBlocks, ...tocSlotBlocks, ...section.children] : section.children,
    })),
  });

  const blob = await Packer.toBlob(doc);
  return new Uint8Array(await blob.arrayBuffer());
}

function patchPackedDocx(packed: Uint8Array): Uint8Array {
  const files = unzipSync(packed);
  const documentXml = new TextDecoder().decode(files["word/document.xml"]!);
  files["word/document.xml"] = new TextEncoder().encode(patchDocumentXml(documentXml));
  if (files["word/numbering.xml"]) {
    const numberingXml = new TextDecoder().decode(files["word/numbering.xml"]);
    files["word/numbering.xml"] = new TextEncoder().encode(patchNumberingXml(numberingXml));
  }
  patchChromeFieldFiles(files);
  return zipSync(files);
}

/** Platform-neutral DOCX bytes from an HTML body fragment and style resolver. */
export async function buildDocxUint8Array(
  html: string,
  styleResolver: StyleResolver,
  imageResolver?: ImageResolver,
  documentConfig?: DocumentConfig,
  onWarning?: WarningHandler | null,
): Promise<Uint8Array> {
  resetImageDocPrIds();
  const allowClassPageMapping = styleResolver.source === "computed";
  const pageRules = parseCssPageRules(html, allowClassPageMapping);
  const inferredOrientation = documentConfig?.orientation ? undefined : pageRules.defaultOrientation;
  const resolved = resolveDocumentConfig(documentConfig, inferredOrientation);
  const warn = resolveWarningHandler(onWarning);
  const $ = cheerio.load(`<body>${html.trim()}</body>`, { xml: false });
  if (imageResolver) await applyImageResolver($, imageResolver);
  emitDegradationWarnings($, styleResolver, Boolean(imageResolver), warn);
  const bodyFieldOptions: InlineFieldOptions = {
    enabled: false,
    onWarning: warn,
  };
  const bodySections = buildBodySections(
    $,
    styleResolver,
    resolved.fontHalfPoints,
    resolved.defaultOrientation,
    pageRules,
    !documentConfig?.orientation,
    allowClassPageMapping,
    bodyFieldOptions,
  );
  const chrome = {
    header: buildHeader(documentConfig, resolved, warn),
    footer: buildFooter(documentConfig, resolved, warn),
  };
  const coverBlocks = buildCover(documentConfig, resolved, warn);
  const tocSlotBlocks = buildTocSlot(documentConfig, resolved, warn);
  const packed = await packDocxToUint8Array(bodySections, resolved, chrome, coverBlocks, tocSlotBlocks);
  return patchPackedDocx(packed);
}

/** Browser entry — returns a `.docx` Blob. */
export async function buildDocxBlob(
  html: string,
  styleResolver: StyleResolver,
  imageResolver?: ImageResolver,
  documentConfig?: DocumentConfig,
  onWarning?: WarningHandler | null,
): Promise<Blob> {
  const bytes = await buildDocxUint8Array(html, styleResolver, imageResolver, documentConfig, onWarning);
  return new Blob([bytes.slice()], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}
