import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import type { Browser, Page } from "playwright";
import { wrapHtml } from "../html-wrap.js";
import { VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX } from "./constants.js";
import type { ComputedStyleSnapshot } from "./computed-style-snapshot.js";
import { ComputedStyleResolver } from "./style-resolver.js";

/** Playwright evaluate payload — kept as raw JS to avoid tsx `__name` injection. */
const COMPUTED_STYLE_SNAPSHOT_FN = readFileSync(
  fileURLToPath(new URL("./computed-style-snapshot.browser.js", import.meta.url)),
  "utf-8",
);

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
  const page = await browser.newPage({
    viewport: { width: VIEWPORT_WIDTH_PX, height: VIEWPORT_HEIGHT_PX },
  });
  try {
    await page.setContent(wrapHtml(html), { waitUntil: "networkidle" });
    return ComputedStyleResolver.fromSnapshots(await snapshotComputedStyles(page));
  } finally {
    await page.close();
  }
}

export async function computedStyleResolverFromPage(
  page: Page,
  rootSelector?: string,
): Promise<ComputedStyleResolver> {
  return ComputedStyleResolver.fromSnapshots(await snapshotComputedStyles(page, rootSelector));
}
