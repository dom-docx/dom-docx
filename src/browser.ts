import { buildDocxUint8Array, type DocumentConfig } from "./converter/build-docx.js";
import { snapshotComputedStylesFromDocument } from "./converter/computed-style-snapshot.js";
import type { ImageResolver } from "./converter/image.js";
import {
  ComputedStyleResolver,
  INLINE_STYLE_RESOLVER,
  type StyleSource,
} from "./converter/style-resolver.js";

export type { ImageResolver, ResolvedImage } from "./converter/image.js";
export type { StyleResolver, StyleSource } from "./converter/style-resolver.js";
export type { DocumentConfig } from "./converter/build-docx.js";
export { buildDocxBlob, buildDocxUint8Array } from "./converter/build-docx.js";
export { snapshotComputedStylesFromDocument } from "./converter/computed-style-snapshot.js";

export interface BrowserConvertOptions extends DocumentConfig {
  styleSource?: StyleSource;
  /** Document to snapshot for `styleSource: "computed"`. Defaults to the host page. */
  document?: Document;
  /**
   * Export root for `styleSource: "computed"`. Pass the live element whose `innerHTML`
   * you convert so computed-style paths match the fragment tree (SPA export pattern).
   */
  root?: Element;
  /** Resolve non-`data:` `<img src>` before conversion (caller owns fetch policy). */
  imageResolver?: ImageResolver;
}

function documentConfig(options?: BrowserConvertOptions): DocumentConfig | undefined {
  if (!options) return undefined;
  const { pageSize, orientation, margins, defaultFont, metadata, headerHtml, footerHtml, pageNumber, lang, direction } = options;
  return { pageSize, orientation, margins, defaultFont, metadata, headerHtml, footerHtml, pageNumber, lang, direction };
}

export async function convertHtmlToDocxUint8Array(
  html: string,
  options?: BrowserConvertOptions,
): Promise<Uint8Array> {
  const styleSource = options?.styleSource ?? "inline";
  const resolver =
    styleSource === "inline"
      ? INLINE_STYLE_RESOLVER
      : ComputedStyleResolver.fromSnapshots(
          snapshotComputedStylesFromDocument(options?.document ?? document, options?.root),
        );
  return buildDocxUint8Array(html, resolver, options?.imageResolver, documentConfig(options));
}

/**
 * Convert an HTML body fragment to a `.docx` Blob in the browser.
 *
 * - `inline` (default): parses `style=""` only — no live DOM required for styles.
 * - `computed`: batch-reads `getComputedStyle` from the export root (or full `document.body`
 *   when `root` is omitted). Pass `root` when converting `element.innerHTML` from a live SPA.
 */
export async function convertHtmlToDocx(
  html: string,
  options?: BrowserConvertOptions,
): Promise<Blob> {
  const bytes = await convertHtmlToDocxUint8Array(html, options);
  return new Blob([bytes.slice()], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

export interface DomDocxGlobal {
  convertHtmlToDocx: typeof convertHtmlToDocx;
  convertHtmlToDocxUint8Array: typeof convertHtmlToDocxUint8Array;
}

declare global {
  interface Window {
    domDocx?: DomDocxGlobal;
  }
}

if (typeof window !== "undefined") {
  window.domDocx = { convertHtmlToDocx, convertHtmlToDocxUint8Array };
}
