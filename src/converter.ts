import type { Browser, Page } from "playwright";
import { buildDocxUint8Array } from "./converter/build-docx.js";
import {
  INLINE_STYLE_RESOLVER,
  type StyleResolver,
  type StyleSource,
} from "./converter/style-resolver.js";
import {
  computedStyleResolverFromPage,
  createComputedStyleResolver,
} from "./converter/style-resolver-node.js";
import type { ImageResolver, ResolvedImage } from "./converter/image.js";
import type { DocumentConfig } from "./converter/build-docx.js";

export type { StyleResolver, StyleSource } from "./converter/style-resolver.js";
export type { ImageResolver, ResolvedImage } from "./converter/image.js";
export type { DocumentConfig } from "./converter/build-docx.js";
export { buildDocxUint8Array, buildDocxBlob } from "./converter/build-docx.js";

export interface ConvertOptions extends DocumentConfig {
  styleSource?: StyleSource;
  /** Reuse an open Playwright browser for computed style resolution (benchmark loops). */
  browser?: Browser;
  /**
   * Browser-native computed styles: read from this already-rendered page's ambient DOM
   * instead of spawning a fresh page. Only used when `styleSource: "computed"`.
   */
  page?: Page;
  /**
   * CSS selector for the export root when converting `element.innerHTML` from a live page
   * (must match the node whose innerHTML was passed as `html`). Playwright `page` path only.
   */
  rootSelector?: string;
  /**
   * Resolve non-`data:` `<img src>` (e.g. `http(s):` / `file:`). The library never fetches
   * on its own — supply this hook to enable remote/local images, and own the fetch plus its
   * security policy (host allowlist, SSRF/private-IP blocking, auth, size caps). Without it,
   * only inline `data:` URLs embed; other images fall back to alt text.
   */
  imageResolver?: ImageResolver;
  // Page layout, default font, and metadata come from DocumentConfig (pageSize,
  // orientation, margins, defaultFont, metadata).
}

function documentConfig(options?: ConvertOptions): DocumentConfig | undefined {
  if (!options) return undefined;
  const { pageSize, orientation, margins, defaultFont, metadata, headerHtml, footerHtml, pageNumber, lang, direction } = options;
  return { pageSize, orientation, margins, defaultFont, metadata, headerHtml, footerHtml, pageNumber, lang, direction };
}

async function resolveStyleResolver(
  html: string,
  styleSource: StyleSource,
  options?: ConvertOptions,
): Promise<{ resolver: StyleResolver; ownsBrowser: Browser | null }> {
  if (styleSource === "inline") {
    return { resolver: INLINE_STYLE_RESOLVER, ownsBrowser: null };
  }

  if (options?.page) {
    return {
      resolver: await computedStyleResolverFromPage(options.page, options.rootSelector),
      ownsBrowser: null,
    };
  }

  // Lazy-load Playwright only on the spawn path — the default inline path never reaches
  // here, so `playwright` stays an optional dependency (no Chromium on `npm i`).
  let ownsBrowser: Browser | null = null;
  let browser = options?.browser;
  if (!browser) {
    const { chromium } = await import("playwright");
    ownsBrowser = await chromium.launch();
    browser = ownsBrowser;
  }
  const resolver = await createComputedStyleResolver(html, browser);
  return { resolver, ownsBrowser };
}

export async function buildDocxBuffer(
  html: string,
  styleResolver: StyleResolver,
  imageResolver?: ImageResolver,
  docConfig?: DocumentConfig,
): Promise<Buffer> {
  return Buffer.from(await buildDocxUint8Array(html, styleResolver, imageResolver, docConfig));
}

export async function convertHtmlToDocx(
  html: string,
  options?: ConvertOptions,
): Promise<Buffer> {
  const styleSource = options?.styleSource ?? "inline";
  const { resolver, ownsBrowser } = await resolveStyleResolver(html, styleSource, options);
  try {
    return await buildDocxBuffer(html, resolver, options?.imageResolver, documentConfig(options));
  } finally {
    if (ownsBrowser) await ownsBrowser.close();
  }
}
