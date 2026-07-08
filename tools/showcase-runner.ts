import { access, copyFile, mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { validateFile } from "@xarsh/ooxml-validator";
import { docxToPdf } from "./docx2pdf.js";
import { pdfToContentFlowPng } from "./pdf-raster.js";
import { chromium, type Browser } from "playwright";
import { convertHtmlToDocx } from "../src/converter.js";
import type { ImageResolver } from "../src/converter/image.js";
import { PAGE_MARGIN_PX, VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX } from "../src/converter/constants.js";
import { generateShowcaseCases, type ShowcaseCase } from "./showcase.js";
import {
  createShowcaseImageResolver,
  resolveFragmentImages,
} from "./showcase-image-resolver.js";
import {
  compositeEngineScore,
  measureEditability,
  performanceScore,
  rollupSuite,
  SCORE_WEIGHTS,
  type EditabilityBreakdown,
} from "./scoring.js";
import type { CaseResult } from "./validator.js";
import { comparePngs, composeSideBySidePng } from "./visual-compare.js";
import { extractPdfDisplayText } from "./pdf-text.js";
import { htmlFragmentDisplayText } from "./text-content-fidelity.js";
import { SHOWCASE_OUTPUT } from "./output-paths.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(__dirname, "..");
const OUTPUT_DIR = SHOWCASE_OUTPUT;
const EXAMPLES_DIR = path.resolve(__dirname, "../examples");
const RESULTS_JSON = path.join(OUTPUT_DIR, "results.json");
const VIEWPORT = { width: VIEWPORT_WIDTH_PX, height: VIEWPORT_HEIGHT_PX };

function wrapHtml(fragment: string): string {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <style>
    * { box-sizing: border-box; }
    body {
      margin: 0;
      padding: ${PAGE_MARGIN_PX}px;
      width: ${VIEWPORT.width}px;
      min-height: ${VIEWPORT.height}px;
      font-family: Arial, Helvetica, sans-serif;
      font-size: 14px;
      line-height: 1.4;
      color: #111;
      background: #fff;
    }
  </style>
</head>
<body>${fragment}</body>
</html>`;
}

async function captureHtmlPng(browser: Browser, html: string, outPath: string): Promise<string> {
  const page = await browser.newPage({ viewport: VIEWPORT });
  try {
    await page.setContent(wrapHtml(html), { waitUntil: "networkidle" });
    await page.screenshot({ path: outPath, fullPage: true });
    return htmlFragmentDisplayText(html);
  } finally {
    await page.close();
  }
}

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}

/** Run a CDN React preview and return the rendered root innerHTML. */
async function renderReactFragment(
  browser: Browser,
  previewPath: string,
  rootId = "root",
): Promise<string> {
  const abs = path.resolve(REPO_ROOT, previewPath);
  const page = await browser.newPage({ viewport: VIEWPORT });
  try {
    await page.goto(`file://${abs}`, { waitUntil: "networkidle" });
    await page.waitForFunction(
      (id) => {
        const root = document.getElementById(id);
        return !!root && root.children.length > 0;
      },
      rootId,
      { timeout: 30_000 },
    );
    return await page.evaluate(
      (id) => document.getElementById(id)!.innerHTML,
      rootId,
    );
  } finally {
    await page.close();
  }
}

async function resolveShowcaseHtml(showcase: ShowcaseCase, browser: Browser): Promise<string> {
  if (showcase.html) return showcase.html;
  if (showcase.reactPreviewPath) {
    return renderReactFragment(browser, showcase.reactPreviewPath, showcase.reactRootId);
  }
  throw new Error(`Showcase "${showcase.name}" needs html or reactPreviewPath`);
}

/** Commit-able artifacts for docs and npm package browsing. */
async function publishExample(
  name: string,
  fragment: string,
  docxPath: string,
  docxPngPath: string,
  sideBySidePath: string,
): Promise<void> {
  const dir = path.join(EXAMPLES_DIR, name);
  await mkdir(dir, { recursive: true });
  await writeFile(path.join(dir, "input.html"), `${fragment.trim()}\n`, "utf-8");
  await copyFile(docxPath, path.join(dir, "output.docx"));
  await copyFile(docxPngPath, path.join(dir, "preview.png"));
  await copyFile(sideBySidePath, path.join(dir, "compare_side_by_side.png"));
}

async function runShowcaseCase(
  showcase: ShowcaseCase,
  html: string,
  browser: Browser,
): Promise<CaseResult> {
  const caseDir = path.join(OUTPUT_DIR, showcase.name);
  await mkdir(caseDir, { recursive: true });

  const imageResolver: ImageResolver | undefined = showcase.usesImageResolver
    ? createShowcaseImageResolver(showcase.name, EXAMPLES_DIR)
    : undefined;
  const previewHtml = imageResolver ? await resolveFragmentImages(html, imageResolver) : html;

  const htmlPath = path.join(caseDir, "source.html");
  const targetPng = path.join(caseDir, "target_html.png");
  const docxPath = path.join(caseDir, "output.docx");
  const pdfPath = path.join(caseDir, "output.pdf");
  const docxPng = path.join(caseDir, "output_docx.png");
  const sideBySidePng = path.join(caseDir, "compare_side_by_side.png");
  const diffPath = path.join(OUTPUT_DIR, `diff_${showcase.name}.png`);

  const result: CaseResult = {
    name: showcase.name,
    xmlPassed: false,
    xmlErrors: [],
    visualCompared: false,
    visualScore: null,
    editability: null,
    compileMs: null,
    performanceScore: null,
    engineScore: null,
  };

  try {
    await writeFile(htmlPath, wrapHtml(previewHtml), "utf-8");
    await writeFile(
      path.join(caseDir, "meta.json"),
      `${JSON.stringify({ title: showcase.title, description: showcase.description, category: showcase.category, styleSource: showcase.styleSource ?? "inline", usesImageResolver: showcase.usesImageResolver ?? false }, null, 2)}\n`,
    );

    const page = await browser.newPage({ viewport: VIEWPORT });
    let htmlDisplayText: string;
    try {
      await page.setContent(wrapHtml(previewHtml), { waitUntil: "networkidle" });
      await page.screenshot({ path: targetPng, fullPage: true });
      htmlDisplayText = htmlFragmentDisplayText(previewHtml);

      const compileStart = performance.now();
      const styleSource = showcase.styleSource ?? "inline";
      const convertOptions = imageResolver ? { imageResolver } : undefined;
      const docxBuffer =
        styleSource === "computed"
          ? await convertHtmlToDocx(html, { styleSource: "computed", page, ...convertOptions })
          : await convertHtmlToDocx(html, convertOptions);
      result.compileMs = performance.now() - compileStart;
      result.performanceScore = performanceScore(result.compileMs);

      await writeFile(docxPath, docxBuffer);
    } finally {
      await page.close();
    }

    const validation = await validateFile(docxPath, { officeVersion: "Office2019" });
    result.xmlPassed = validation.ok;
    if (!validation.ok) {
      result.xmlErrors = validation.errors.map(
        (e) => `[${e.errorType}] ${e.path}: ${e.description}`,
      );
    }

    const editability = await measureEditability(docxPath);
    result.editability = editability.score;
    result.editabilityBreakdown = editability;

    if (!validation.ok) {
      return result;
    }

    await docxToPdf(docxPath, pdfPath);
    if (!(await fileExists(pdfPath))) {
      throw new Error(`PDF not created at ${pdfPath}`);
    }
    await pdfToContentFlowPng(pdfPath, docxPng);
    const docxDisplayText = await extractPdfDisplayText(pdfPath);

    const comparison = await comparePngs(targetPng, docxPng, diffPath, {
      htmlDisplayText,
      docxDisplayText,
      htmlFragment: previewHtml,
    });
    await composeSideBySidePng(targetPng, docxPng, sideBySidePng);
    await publishExample(showcase.name, html, docxPath, docxPng, sideBySidePng);
    result.visualCompared = true;
    result.mismatchedPixels = comparison.mismatchedPixels;
    result.matchPercent = comparison.matchPercent;
    result.legibilityScore = comparison.legibility.score;
    result.illegibleBandCount = comparison.legibility.illegibleBandCount;
    result.backgroundBalanceScore = comparison.backgroundBalance.score;
    result.imbalancedFillBlockCount = comparison.backgroundBalance.imbalancedBlockCount;
    result.listMarkerFidelityScore = comparison.listMarkerFidelity.score;
    result.htmlMarkerLineCount = comparison.listMarkerFidelity.htmlMarkerLineCount;
    result.docxMarkerLineCount = comparison.listMarkerFidelity.docxMarkerLineCount;
    if (comparison.textContentFidelity) {
      result.textContentFidelityScore = comparison.textContentFidelity.score;
      result.missingTokenUnits = comparison.textContentFidelity.missingTokenUnits;
      result.extraTokenUnits = comparison.textContentFidelity.extraTokenUnits;
      result.sampleMissingTokens = comparison.textContentFidelity.sampleMissing;
      result.sampleExtraTokens = comparison.textContentFidelity.sampleExtra;
    }
    result.visualScore = comparison.visualScore;
    result.diffPath = diffPath;
    result.sideBySidePath = sideBySidePng;

    result.engineScore = compositeEngineScore(
      result.visualScore,
      result.editability,
      result.performanceScore,
    );
  } catch (err) {
    result.error = err instanceof Error ? err.message : String(err);
  }

  return result;
}

function printShowcaseReport(cases: ShowcaseCase[], results: CaseResult[]): void {
  const line = "─".repeat(60);
  console.log(`\n${line}`);
  console.log("  HTML → DOCX Showcase — Conversion Report");
  console.log(line);
  console.log(`  Output: ${OUTPUT_DIR}`);
  console.log(`  Results JSON: ${RESULTS_JSON}`);
  console.log(`  (Not included in npm run score:suite regression suite)\n`);

  for (const showcase of cases) {
    const r = results.find((x) => x.name === showcase.name);
    if (!r) continue;

    console.log(`  [${showcase.name}] ${showcase.title}`);
    console.log(`    ${showcase.category} · ${showcase.description}`);

    if (r.error) {
      console.log(`    [!] Error: ${r.error}`);
      continue;
    }

    const xmlIcon = r.xmlPassed ? "✓" : "✗";
    console.log(`    [${xmlIcon}] XML Schema`);

    if (r.compileMs !== null) {
      console.log(`    [•] Compile: ${r.compileMs.toFixed(1)} ms`);
    }

    if (r.editabilityBreakdown) {
      const b: EditabilityBreakdown = r.editabilityBreakdown;
      console.log(
        `    [•] Editability: ${b.score.toFixed(0)} (1×1 tables: ${b.oneByOneTables})`,
      );
    }

    if (r.visualCompared && r.matchPercent !== undefined) {
      console.log(`    [•] Visual match: ${r.matchPercent.toFixed(2)}% (content region)`);
      if (
        r.legibilityScore !== undefined &&
        r.legibilityScore < 99.5
      ) {
        console.log(
          `        Legibility: ${r.legibilityScore.toFixed(2)}% (${r.illegibleBandCount ?? 0} illegible bands)`,
        );
      }
      if (
        r.backgroundBalanceScore !== undefined &&
        r.backgroundBalanceScore < 99.5
      ) {
        console.log(
          `        Background balance: ${r.backgroundBalanceScore.toFixed(2)}% (${r.imbalancedFillBlockCount ?? 0} imbalanced fills)`,
        );
      }
      if (
        r.textContentFidelityScore !== undefined &&
        r.visualScore !== null &&
        r.textContentFidelityScore < 99.5
      ) {
        const missing = r.sampleMissingTokens?.length
          ? ` missing: ${r.sampleMissingTokens.join(", ")}`
          : "";
        const extra = r.sampleExtraTokens?.length ? ` extra: ${r.sampleExtraTokens.join(", ")}` : "";
        console.log(
          `        Text content: ${r.textContentFidelityScore.toFixed(2)}% (${r.missingTokenUnits ?? 0} missing · ${r.extraTokenUnits ?? 0} extra units${missing}${extra})`,
        );
      }
      if (
        r.visualScore !== null &&
        r.matchPercent !== undefined &&
        r.visualScore < r.matchPercent - 0.5
      ) {
        console.log(`        Adjusted visual: ${r.visualScore.toFixed(2)}%`);
      }
      if (r.diffPath) console.log(`        Diff: ${r.diffPath}`);
      if (r.sideBySidePath) console.log(`        Side-by-side: ${r.sideBySidePath}`);
    }

    if (r.engineScore !== null) {
      console.log(`    [★] Engine score: ${r.engineScore.toFixed(2)}`);
    }
    console.log("");
  }

  const suite = rollupSuite(
    results.map((r) => ({
      visualMatch: r.visualScore,
      editability: r.editability,
      performance: r.performanceScore,
      engineScore: r.engineScore,
      compileMs: r.compileMs,
      xmlPassed: r.xmlPassed,
      error: r.error,
    })),
  );

  console.log(line);
  console.log(
    `  Showcase avg: visual ${suite.visual?.toFixed(2) ?? "—"} · editability ${suite.editability?.toFixed(2) ?? "—"} · engine ${suite.engine?.toFixed(2) ?? "—"}`,
  );
  console.log(`${line}\n`);
}

async function main(): Promise<void> {
  await mkdir(OUTPUT_DIR, { recursive: true });
  const only = process.env.SHOWCASE_ONLY?.split(",").map((s) => s.trim()).filter(Boolean);
  const cases = generateShowcaseCases().filter((c) => !only?.length || only.includes(c.name));
  console.log(`Running ${cases.length} showcase examples → ${OUTPUT_DIR}`);

  const browser = await chromium.launch();
  const results: CaseResult[] = [];

  try {
    for (const showcase of cases) {
      console.log(`  • ${showcase.name}`);
      const html = await resolveShowcaseHtml(showcase, browser);
      results.push(await runShowcaseCase(showcase, html, browser));
    }
  } finally {
    await browser.close();
  }

  const payload = {
    version: 1,
    kind: "showcase",
    runAt: new Date().toISOString(),
    weights: SCORE_WEIGHTS,
    suite: rollupSuite(
      results.map((r) => ({
        visualMatch: r.visualScore,
        editability: r.editability,
        performance: r.performanceScore,
        engineScore: r.engineScore,
        compileMs: r.compileMs,
        xmlPassed: r.xmlPassed,
        error: r.error,
      })),
    ),
    cases: results.map((r) => {
      const meta = cases.find((c) => c.name === r.name);
      return {
        ...r,
        title: meta?.title,
        category: meta?.category,
        description: meta?.description,
      };
    }),
  };

  await writeFile(RESULTS_JSON, `${JSON.stringify(payload, null, 2)}\n`, "utf-8");
  printShowcaseReport(cases, results);

  if (results.some((r) => !r.xmlPassed || r.error)) {
    process.exitCode = 1;
  }
}

main();
