import {
  AlignmentType,
  Document,
  Footer,
  Header,
  PageBreak,
  PageNumber,
  Packer,
  PageOrientation,
  Paragraph,
  TableOfContents,
  TextRun,
  convertInchesToTwip,
  type FileChild,
} from "docx";
import * as cheerio from "cheerio";
import { unzipSync, zipSync } from "fflate";
import { BODY_FONT, BODY_FONT_HALF_POINTS, NUMBERING_CONFIG, PAGE_MARGIN_TWIPS } from "./constants.js";
import {
  patchDocumentXml,
  patchHeadingOutlineLevels,
  patchNumberingXml,
  patchTableOfContents,
} from "./ooxml-patch.js";
import { applyImageResolver, resetImageDocPrIds, type ImageResolver } from "./image.js";
import { INLINE_STYLE_RESOLVER, type StyleResolver } from "./style-resolver.js";
import { htmlToDocxBlocks } from "./visitor.js";

/** Table-of-contents field options (see `DocumentConfig.tableOfContents`). */
export interface TableOfContentsConfig {
  /**
   * Heading text rendered above the TOC (e.g. `"Contents"`). Rendered as a plain
   * bold paragraph — not a Word heading — so it never appears as its own entry.
   * Omit for no title.
   */
  title?: string;
  /**
   * Heading levels to include, as a Word range, e.g. `"1-3"` (default) or `"1-2"`.
   * `h1`–`h6` map to Word Heading 1–6.
   */
  headingRange?: string;
  /** Render entries as clickable hyperlinks to their headings (default `true`). */
  hyperlink?: boolean;
  /** Insert a page break after the TOC so body content starts on a new page. */
  pageBreakAfter?: boolean;
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
  /** Append a centered `Page N` field to the footer (created if `footerHtml` is absent). */
  pageNumber?: boolean;
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
   * Insert a clickable, page-number-less Table of Contents at the top of the
   * document, built from the headings present (`h1`–`h6` become Word Heading 1–6).
   * Each entry is a hyperlink to its heading. Page numbers depend on layout, which
   * this library does not do — so instead of a live field that must be refreshed
   * (and prompts to "update fields"), the TOC omits page numbers and is complete at
   * creation: correct in every viewer with no update. `true` uses defaults (levels
   * 1–3, clickable).
   */
  tableOfContents?: boolean | TableOfContentsConfig;
}

// Portrait dimensions in twips. Letter matches convertInchesToTwip(8.5)×(11).
const PAGE_PRESETS_TWIPS = {
  letter: { width: 12240, height: 15840 },
  a4: { width: 11906, height: 16838 },
} as const;

interface ResolvedConfig {
  size: { width: number; height: number; orientation?: (typeof PageOrientation)[keyof typeof PageOrientation] };
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
  toc?: TableOfContentsConfig;
}

function resolveTableOfContents(
  toc: DocumentConfig["tableOfContents"],
): TableOfContentsConfig | undefined {
  if (!toc) return undefined;
  return toc === true ? {} : toc;
}

function resolveDocumentConfig(config?: DocumentConfig): ResolvedConfig {
  const ps = config?.pageSize;
  const base =
    !ps || ps === "letter"
      ? PAGE_PRESETS_TWIPS.letter
      : ps === "a4"
        ? PAGE_PRESETS_TWIPS.a4
        : { width: convertInchesToTwip(ps.width), height: convertInchesToTwip(ps.height) };

  // docx swaps width/height itself for landscape, so pass portrait dims + the flag.
  const size =
    config?.orientation === "landscape"
      ? { width: base.width, height: base.height, orientation: PageOrientation.LANDSCAPE }
      : { width: base.width, height: base.height };

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
    size,
    margin: { top: marginIn(m?.top), right: marginIn(m?.right), bottom: marginIn(m?.bottom), left: marginIn(m?.left) },
    font: config?.defaultFont?.family ?? BODY_FONT,
    fontHalfPoints:
      config?.defaultFont?.sizePt !== undefined
        ? Math.round(config.defaultFont.sizePt * 2)
        : BODY_FONT_HALF_POINTS,
    metadata,
    lang: config?.lang,
    rtl: config?.direction === "rtl",
    toc: resolveTableOfContents(config?.tableOfContents),
  };
}

/** Convert a standalone HTML fragment (header/footer) to DOCX blocks via the inline resolver. */
function fragmentToBlocks(html: string, sizeHalfPoints: number): FileChild[] {
  const $ = cheerio.load(`<body>${html.trim()}</body>`, { xml: false });
  return htmlToDocxBlocks($, INLINE_STYLE_RESOLVER, sizeHalfPoints);
}

function pageNumberParagraph(): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun("Page "), new TextRun({ children: [PageNumber.CURRENT] })],
  });
}

const DEFAULT_HEADING_RANGE = "1-3";

/** The resolved heading range + hyperlink flag the TOC patch needs, or `undefined`. */
function tocPatchOptions(
  resolved: ResolvedConfig,
): { headingRange: string; hyperlink: boolean } | undefined {
  if (!resolved.toc) return undefined;
  return {
    headingRange: resolved.toc.headingRange ?? DEFAULT_HEADING_RANGE,
    hyperlink: resolved.toc.hyperlink ?? true,
  };
}

/**
 * TOC field + optional title, prepended to the body. docx emits the bare field
 * (its `\o`/`\h` instruction and an empty result); `patchTableOfContents` then
 * fills the result with clickable, page-number-less entries and bookmarks the
 * headings. The field is NOT dirty and we set no `updateFields`: a number-less TOC
 * is complete at creation, so there is nothing to refresh and no "update fields"
 * prompt. Returns `[]` when no TOC is configured.
 */
function buildTableOfContents(resolved: ResolvedConfig): FileChild[] {
  const toc = resolved.toc;
  if (!toc) return [];

  const blocks: FileChild[] = [];
  if (toc.title) {
    blocks.push(
      new Paragraph({
        children: [
          new TextRun({
            text: toc.title,
            bold: true,
            font: resolved.font,
            size: Math.round(resolved.fontHalfPoints * 1.5),
          }),
        ],
      }),
    );
  }

  blocks.push(
    new TableOfContents("Table of Contents", {
      hyperlink: toc.hyperlink ?? true,
      headingStyleRange: toc.headingRange ?? DEFAULT_HEADING_RANGE,
      beginDirty: false,
    }),
  );

  if (toc.pageBreakAfter) {
    blocks.push(new Paragraph({ children: [new PageBreak()] }));
  }
  return blocks;
}

/**
 * Cover fragment as the document's first content, followed by a page break so the
 * TOC/body start on the next page. Converted via the inline style path (like
 * header/footer). Returns `[]` when no `coverHtml` is configured.
 */
function buildCover(config: DocumentConfig | undefined, resolved: ResolvedConfig): FileChild[] {
  if (!config?.coverHtml) return [];
  return [
    ...fragmentToBlocks(config.coverHtml, resolved.fontHalfPoints),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

function buildFooter(config: DocumentConfig | undefined, resolved: ResolvedConfig): Footer | undefined {
  const hasFooterHtml = Boolean(config?.footerHtml);
  if (!hasFooterHtml && !config?.pageNumber) return undefined;
  const children: FileChild[] = hasFooterHtml
    ? fragmentToBlocks(config!.footerHtml!, resolved.fontHalfPoints)
    : [];
  if (config?.pageNumber) children.push(pageNumberParagraph());
  return new Footer({ children });
}

function buildHeader(config: DocumentConfig | undefined, resolved: ResolvedConfig): Header | undefined {
  if (!config?.headerHtml) return undefined;
  return new Header({ children: fragmentToBlocks(config.headerHtml, resolved.fontHalfPoints) });
}

async function packDocxToUint8Array(
  children: FileChild[],
  resolved: ResolvedConfig,
  chrome: { header?: Header; footer?: Footer },
  coverBlocks: FileChild[],
): Promise<Uint8Array> {
  const listStyleRun = { font: resolved.font, size: resolved.fontHalfPoints };
  // A cover page suppresses the header/footer/page-number on page 1 via Word's
  // "different first page" (titlePg + empty first-page header/footer).
  const suppressFirstChrome = coverBlocks.length > 0 && Boolean(chrome.header || chrome.footer);
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
    sections: [
      {
        properties: {
          page: {
            size: resolved.size,
            margin: resolved.margin,
          },
          ...(suppressFirstChrome ? { titlePage: true } : {}),
        },
        ...(chrome.header
          ? { headers: { default: chrome.header, ...(suppressFirstChrome ? { first: new Header({ children: [] }) } : {}) } }
          : {}),
        ...(chrome.footer
          ? { footers: { default: chrome.footer, ...(suppressFirstChrome ? { first: new Footer({ children: [] }) } : {}) } }
          : {}),
        children: [...coverBlocks, ...buildTableOfContents(resolved), ...children],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  return new Uint8Array(await blob.arrayBuffer());
}

function patchPackedDocx(
  packed: Uint8Array,
  toc?: { headingRange: string; hyperlink: boolean },
): Uint8Array {
  const files = unzipSync(packed);
  let documentXml = patchDocumentXml(new TextDecoder().decode(files["word/document.xml"]));
  if (toc) documentXml = patchTableOfContents(documentXml, toc.headingRange, toc.hyperlink);
  files["word/document.xml"] = new TextEncoder().encode(documentXml);
  if (files["word/numbering.xml"]) {
    const numberingXml = new TextDecoder().decode(files["word/numbering.xml"]);
    files["word/numbering.xml"] = new TextEncoder().encode(patchNumberingXml(numberingXml));
  }
  // A TOC collects by outline level; give the heading styles explicit levels so a
  // manual field refresh rebuilds correctly in LibreOffice (Word infers them).
  if (toc && files["word/styles.xml"]) {
    const stylesXml = new TextDecoder().decode(files["word/styles.xml"]);
    files["word/styles.xml"] = new TextEncoder().encode(patchHeadingOutlineLevels(stylesXml));
  }
  return zipSync(files);
}

/** Platform-neutral DOCX bytes from an HTML body fragment and style resolver. */
export async function buildDocxUint8Array(
  html: string,
  styleResolver: StyleResolver,
  imageResolver?: ImageResolver,
  documentConfig?: DocumentConfig,
): Promise<Uint8Array> {
  resetImageDocPrIds();
  const resolved = resolveDocumentConfig(documentConfig);
  const $ = cheerio.load(`<body>${html.trim()}</body>`, { xml: false });
  if (imageResolver) await applyImageResolver($, imageResolver);
  const children = htmlToDocxBlocks($, styleResolver, resolved.fontHalfPoints);
  const chrome = {
    header: buildHeader(documentConfig, resolved),
    footer: buildFooter(documentConfig, resolved),
  };
  const coverBlocks = buildCover(documentConfig, resolved);
  const packed = await packDocxToUint8Array(children, resolved, chrome, coverBlocks);
  return patchPackedDocx(packed, tocPatchOptions(resolved));
}

/** Browser entry — returns a `.docx` Blob. */
export async function buildDocxBlob(
  html: string,
  styleResolver: StyleResolver,
  imageResolver?: ImageResolver,
  documentConfig?: DocumentConfig,
): Promise<Blob> {
  const bytes = await buildDocxUint8Array(html, styleResolver, imageResolver, documentConfig);
  return new Blob([bytes.slice()], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}
