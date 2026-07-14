import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import type { Browser, Page } from "playwright";
import { wrapHtml } from "../html-wrap.js";
import { VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX } from "./constants.js";
import type { ComputedStyleSnapshot } from "./computed-style-snapshot.js";
import type { RasterizeInPlaceOptions } from "./rasterize-subtree.browser.js";
import { ComputedStyleResolver } from "./style-resolver.js";

/** Playwright evaluate payloads — kept as raw JS to avoid tsx `__name` injection. */
const COMPUTED_STYLE_SNAPSHOT_FN = readFileSync(
  fileURLToPath(new URL("./computed-style-snapshot.browser.js", import.meta.url)),
  "utf-8",
);

const RASTERIZE_EXPORT_PREP_FN = readFileSync(
  fileURLToPath(new URL("./rasterize-export-prep.browser.js", import.meta.url)),
  "utf-8",
);

export interface PlaywrightPreparedExport {
  html: string;
  snapshotRootSelector: string | null;
  cleanupSelector: string | null;
}

export async function preparePlaywrightRasterizedExport(
  page: Page,
  rootSelector: string | null | undefined,
  options?: RasterizeInPlaceOptions,
  mutate?: boolean,
): Promise<PlaywrightPreparedExport> {
  const rootArg = rootSelector ?? null;
  const optsArg = options ?? null;
  const mutateArg = mutate ?? false;
  return (await page.evaluate(
    `(${RASTERIZE_EXPORT_PREP_FN.trim()})(${JSON.stringify(rootArg)}, ${JSON.stringify(optsArg)}, ${mutateArg})`,
  )) as PlaywrightPreparedExport;
}

export async function openPlaywrightPage(
  html: string,
  browser: Browser,
): Promise<Page> {
  const page = await browser.newPage({
    viewport: { width: VIEWPORT_WIDTH_PX, height: VIEWPORT_HEIGHT_PX },
  });
  // DOCXs are a light canvas — never snapshot prefers-color-scheme: dark colors.
  await page.emulateMedia({ colorScheme: "light" });
  await page.setContent(wrapHtml(html), { waitUntil: "networkidle" });
  return page;
}

export async function snapshotComputedStyles(
  page: Page,
  rootSelector?: string,
): Promise<ComputedStyleSnapshot[]> {
  const rootArg = rootSelector
    ? `document.querySelector(${JSON.stringify(rootSelector)})`
    : "null";
  return (await page.evaluate(
    `(${COMPUTED_STYLE_SNAPSHOT_FN.trim()})(${rootArg})`,
  )) as ComputedStyleSnapshot[];
}

export async function createComputedStyleResolver(
  html: string,
  browser: Browser,
): Promise<ComputedStyleResolver> {
  const page = await openPlaywrightPage(html, browser);
  try {
    return ComputedStyleResolver.fromSnapshots(await snapshotComputedStyles(page));
  } finally {
    await page.close();
  }
}

/**
 * Snapshot computed styles from a live Playwright page for DOCX export.
 * Forces `prefers-color-scheme: light` for the duration of the snapshot (Word is a
 * white-page medium); restores the previous media emulation afterward.
 */
export async function computedStyleResolverFromPage(
  page: Page,
  rootSelector?: string,
): Promise<ComputedStyleResolver> {
  await page.emulateMedia({ colorScheme: "light" });
  try {
    return ComputedStyleResolver.fromSnapshots(await snapshotComputedStyles(page, rootSelector));
  } finally {
    await page.emulateMedia({ colorScheme: null });
  }
}
