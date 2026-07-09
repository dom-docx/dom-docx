import { access, mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { validateFile } from "@xarsh/ooxml-validator";
import { docxToPdf } from "./docx2pdf.js";
import { chromium, type Browser } from "playwright";
import { convertHtmlToDocx } from "../src/converter.js";
import { VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX, PAGE_MARGIN_PX } from "../src/converter/constants.js";
import {
  generateTestCases,
  resolveHarnessConvertOptions,
  resolveLoopTestCases,
  type LoopCaseMode,
  type TestCase,
} from "./generator.js";
import {
  SCORE_WEIGHTS,
  compositeEngineScore,
  measureEditability,
  performanceScore,
  rollupSuite,
  type EditabilityBreakdown,
  type SuiteRollup,
} from "./scoring.js";
import { comparePngs, composeSideBySidePng } from "./visual-compare.js";
import { pdfToContentFlowPng } from "./pdf-raster.js";
import { extractPdfDisplayText } from "./pdf-text.js";
import {
  htmlDisplayTextOptionsForConvertOptions,
  htmlFragmentDisplayText,
  type HtmlFragmentDisplayTextOptions,
} from "./text-content-fidelity.js";
import type { ListMarkerFidelityDetail } from "./list-marker-fidelity.js";
import { captureEnvironment, type HarnessEnvironment } from "./environment.js";
import { SUITE_OUTPUT } from "./output-paths.js";

const OUTPUT_DIR = SUITE_OUTPUT;
const STRICT_VISUAL = process.argv.includes("--strict-visual");
const LOOP_MODE: LoopCaseMode = process.argv.includes("--priority") ? "priority" : "full";
const SUITE_ONLY = process.env.SUITE_ONLY?.split(",").map((s) => s.trim()).filter(Boolean);
const RESULTS_JSON = path.join(
  OUTPUT_DIR,
  LOOP_MODE === "priority" ? "results-priority.json" : "results.json",
);
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

async function captureHtmlPng(
  browser: Browser,
  html: string,
  outPath: string,
  displayTextOptions?: HtmlFragmentDisplayTextOptions,
): Promise<string> {
  const page = await browser.newPage({ viewport: VIEWPORT });
  try {
    await page.setContent(wrapHtml(html), { waitUntil: "networkidle" });
    await page.screenshot({ path: outPath, fullPage: true });
    return htmlFragmentDisplayText(html, displayTextOptions);
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

export interface CaseResult {
  name: string;
  xmlPassed: boolean;
  xmlErrors: string[];
  visualCompared: boolean;
  mismatchedPixels?: number;
  matchPercent?: number;
  legibilityScore?: number;
  illegibleBandCount?: number;
  backgroundBalanceScore?: number;
  imbalancedFillBlockCount?: number;
  listMarkerFidelityScore?: number;
  listMarkerEffectiveScore?: number;
  listMarkerDetail?: ListMarkerFidelityDetail;
  htmlMarkerLineCount?: number;
  docxMarkerLineCount?: number;
  textContentFidelityScore?: number;
  /** Report-only order-aware text similarity (LCS); not folded into visualScore. */
  textOrderedSimilarity?: number;
  /** Raw ink-projection layout similarity (before content-quality guards → visualScore). */
  layoutFidelityScore?: number;
  layoutVerticalScore?: number;
  layoutHorizontalScore?: number;
  layoutInkRatio?: number;
  layoutBandCountRef?: number;
  layoutBandCountDocx?: number;
  missingTokenUnits?: number;
  extraTokenUnits?: number;
  sampleMissingTokens?: string[];
  sampleExtraTokens?: string[];
  visualScore: number | null;
  editability: number | null;
  editabilityBreakdown?: EditabilityBreakdown;
  compileMs: number | null;
  performanceScore: number | null;
  engineScore: number | null;
  diffPath?: string;
  sideBySidePath?: string;
  error?: string;
}

export interface LoopResults {
  version: 1;
  runAt: string;
  mode: LoopCaseMode;
  environment: HarnessEnvironment;
  caseCount: number;
  totalCaseCount: number;
  weights: typeof SCORE_WEIGHTS;
  strictVisual: boolean;
  objective: number | null;
  suite: SuiteRollup;
  cases: CaseResult[];
}

async function runCase(testCase: TestCase, browser: Browser): Promise<CaseResult> {
  const caseDir = path.join(OUTPUT_DIR, testCase.name);
  await mkdir(caseDir, { recursive: true });

  const htmlPath = path.join(caseDir, "source.html");
  const targetPng = path.join(caseDir, "target_html.png");
  const docxPath = path.join(caseDir, "output.docx");
  const pdfPath = path.join(caseDir, "output.pdf");
  const docxPng = path.join(caseDir, "output_docx.png");
  const sideBySidePng = path.join(caseDir, "compare_side_by_side.png");
  const diffPath = path.join(OUTPUT_DIR, `diff_${testCase.name}.png`);

  const result: CaseResult = {
    name: testCase.name,
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
    await writeFile(htmlPath, wrapHtml(testCase.html), "utf-8");
    const htmlDisplayText = await captureHtmlPng(
      browser,
      testCase.html,
      targetPng,
      htmlDisplayTextOptionsForConvertOptions(testCase.convertOptions),
    );

    const compileStart = performance.now();
    const harnessOpts = resolveHarnessConvertOptions(testCase, browser);
    const docxBuffer = harnessOpts
      ? await convertHtmlToDocx(testCase.html, harnessOpts)
      : await convertHtmlToDocx(testCase.html);
    result.compileMs = performance.now() - compileStart;
    result.performanceScore = performanceScore(result.compileMs);

    await writeFile(docxPath, docxBuffer);

    const validation = await validateFile(docxPath, { officeVersion: "Office2019" });
    result.xmlPassed = validation.ok;

    if (!validation.ok) {
      result.xmlErrors = validation.errors.map(
        (e) => `[${e.errorType}] ${e.path}: ${e.description}`,
      );
    }

    try {
      const editability = await measureEditability(docxPath);
      result.editability = editability.score;
      result.editabilityBreakdown = editability;
    } catch (editErr) {
      result.error = editErr instanceof Error ? editErr.message : String(editErr);
      return result;
    }

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
      htmlFragment: testCase.html,
    });
    await composeSideBySidePng(targetPng, docxPng, sideBySidePng);
    result.visualCompared = true;
    result.mismatchedPixels = comparison.mismatchedPixels;
    result.matchPercent = comparison.matchPercent;
    result.legibilityScore = comparison.legibility.score;
    result.illegibleBandCount = comparison.legibility.illegibleBandCount;
    result.backgroundBalanceScore = comparison.backgroundBalance.score;
    result.imbalancedFillBlockCount = comparison.backgroundBalance.imbalancedBlockCount;
    result.listMarkerFidelityScore = comparison.listMarkerFidelity.score;
    result.listMarkerEffectiveScore = comparison.listMarkerEffectiveScore;
    result.listMarkerDetail = comparison.listMarkerFidelity.detail;
    result.htmlMarkerLineCount = comparison.listMarkerFidelity.htmlMarkerLineCount;
    result.docxMarkerLineCount = comparison.listMarkerFidelity.docxMarkerLineCount;
    if (comparison.textContentFidelity) {
      result.textContentFidelityScore = comparison.textContentFidelity.score;
      result.textOrderedSimilarity = comparison.textContentFidelity.orderedSimilarity;
      result.missingTokenUnits = comparison.textContentFidelity.missingTokenUnits;
      result.extraTokenUnits = comparison.textContentFidelity.extraTokenUnits;
      result.sampleMissingTokens = comparison.textContentFidelity.sampleMissing;
      result.sampleExtraTokens = comparison.textContentFidelity.sampleExtra;
    }
    result.layoutFidelityScore = comparison.layoutFidelity.score;
    result.layoutVerticalScore = comparison.layoutFidelity.verticalScore;
    result.layoutHorizontalScore = comparison.layoutFidelity.horizontalScore;
    result.layoutInkRatio = comparison.layoutFidelity.inkRatio;
    result.layoutBandCountRef = comparison.layoutFidelity.bandCountRef;
    result.layoutBandCountDocx = comparison.layoutFidelity.bandCountDocx;
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

function printScorecard(results: CaseResult[], suite: SuiteRollup, mode: LoopCaseMode): void {
  const line = "─".repeat(60);
  const total = generateTestCases().length;
  console.log(`\n${line}`);
  console.log("  HTML → DOCX Validation Loop — Engine Scorecard");
  if (mode === "priority") {
    console.log(`  Mode: priority (${results.length}/${total} cases) — run npm run score:suite for full suite`);
  }
  console.log(line);
  console.log(
    `\n  Objective (Engine Score): ${suite.engine !== null ? suite.engine.toFixed(2) : "—"}`,
  );
  console.log(
    `  Weights: visual ${SCORE_WEIGHTS.visual} · editability ${SCORE_WEIGHTS.editability} · performance ${SCORE_WEIGHTS.performance}`,
  );
  console.log(
    `  Suite avg: visual (layout-based) ${fmt(suite.visual)} · editability ${fmt(suite.editability)} · performance ${fmt(suite.performance)} · pixel tripwire ${fmt(suite.pixelMatch)}`,
  );
  if (suite.avgPerformanceMs !== null) {
    console.log(`  Avg compile: ${suite.avgPerformanceMs.toFixed(1)} ms (excl. LibreOffice)`);
  }
  console.log(`  Results JSON: ${RESULTS_JSON}`);

  for (const r of results) {
    const xmlIcon = r.xmlPassed ? "✓" : "✗";
    console.log(`\n  [${r.name}]`);
    console.log(`    [${xmlIcon}] XML Schema Passed`);

    if (r.xmlErrors.length > 0) {
      for (const err of r.xmlErrors) console.log(`        ↳ ${err}`);
    }

    if (r.error) {
      console.log(`    [!] Error: ${r.error}`);
      continue;
    }

    if (r.compileMs !== null) {
      console.log(
        `    [•] Compile: ${r.compileMs.toFixed(1)} ms → performance ${fmt(r.performanceScore)}`,
      );
    }

    if (r.editabilityBreakdown) {
      const b = r.editabilityBreakdown;
      console.log(
        `    [•] Editability: ${b.score.toFixed(0)} (1×1 tables: ${b.oneByOneTables}, cantSplit: ${b.cantSplitLocks})`,
      );
    }

    if (r.visualCompared && r.matchPercent !== undefined && r.mismatchedPixels !== undefined) {
      const visualIcon = (r.visualScore ?? 0) >= 99.5 ? "✓" : "✗";
      console.log(
        `    [${visualIcon}] Visual (layout-based): ${r.visualScore !== null ? r.visualScore.toFixed(2) : "—"}%` +
          (r.layoutVerticalScore !== undefined && r.layoutHorizontalScore !== undefined
            ? ` (V ${r.layoutVerticalScore.toFixed(0)} · H ${r.layoutHorizontalScore.toFixed(0)} · bands ${r.layoutBandCountRef ?? "?"}/${r.layoutBandCountDocx ?? "?"})`
            : ""),
      );
      if (
        r.legibilityScore !== undefined &&
        r.visualScore !== null &&
        r.legibilityScore < 99.5
      ) {
        console.log(
          `        Legibility: ${r.legibilityScore.toFixed(2)}% (${r.illegibleBandCount ?? 0} illegible bands)`,
        );
      }
      if (
        r.backgroundBalanceScore !== undefined &&
        r.visualScore !== null &&
        r.backgroundBalanceScore < 99.5
      ) {
        console.log(
          `        Background balance: ${r.backgroundBalanceScore.toFixed(2)}% (${r.imbalancedFillBlockCount ?? 0} imbalanced fills)`,
        );
      }
      if (
        r.listMarkerFidelityScore !== undefined &&
        r.visualScore !== null &&
        r.listMarkerFidelityScore < 99.5
      ) {
        console.log(
          `        List markers: ${r.listMarkerFidelityScore.toFixed(2)}% (HTML ${r.htmlMarkerLineCount ?? 0} · DOCX ${r.docxMarkerLineCount ?? 0} marker lines)`,
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
        r.textOrderedSimilarity !== undefined &&
        r.visualScore !== null &&
        r.textOrderedSimilarity < 99.5
      ) {
        console.log(
          `        Text order (report-only): ${r.textOrderedSimilarity.toFixed(2)}%`,
        );
      }
      if (
        r.visualScore !== null &&
        r.layoutFidelityScore !== undefined &&
        r.visualScore < r.layoutFidelityScore - 0.5
      ) {
        console.log(
          `        Raw layout before guards: ${r.layoutFidelityScore.toFixed(2)}% (guards cost ${(r.layoutFidelityScore - r.visualScore).toFixed(2)} pp)`,
        );
      }
      console.log(
        `        Pixel tripwire: ${r.matchPercent.toFixed(2)}% · ${r.mismatchedPixels.toLocaleString()} misaligned px`,
      );
      if (r.mismatchedPixels > 0 && r.diffPath) {
        console.log(`        Diff: ${r.diffPath}`);
      }
      if (r.sideBySidePath) {
        console.log(`        Side-by-side: ${r.sideBySidePath}`);
      }
    } else if (r.xmlPassed) {
      console.log("    [—] Visual comparison skipped");
    }

    if (r.engineScore !== null) {
      console.log(`    [★] Engine Score: ${r.engineScore.toFixed(2)}`);
    }
  }

  console.log(`\n${line}\n`);
}

function fmt(value: number | null): string {
  return value !== null ? value.toFixed(2) : "—";
}

function buildLoopResults(
  results: CaseResult[],
  mode: LoopCaseMode,
  totalCaseCount: number,
  environment: HarnessEnvironment,
): LoopResults {
  const suite = rollupSuite(
    results.map((r) => ({
      visualMatch: r.visualScore,
      editability: r.editability,
      performance: r.performanceScore,
      engineScore: r.engineScore,
      compileMs: r.compileMs,
      xmlPassed: r.xmlPassed,
      layout: r.layoutFidelityScore ?? null,
      pixelMatch: r.matchPercent ?? null,
      error: r.error,
    })),
  );
  return {
    version: 1,
    runAt: new Date().toISOString(),
    mode,
    environment,
    caseCount: results.length,
    totalCaseCount,
    weights: SCORE_WEIGHTS,
    strictVisual: STRICT_VISUAL,
    objective: suite.engine,
    suite,
    cases: results,
  };
}

async function writeResultsJson(
  results: CaseResult[],
  mode: LoopCaseMode,
  totalCaseCount: number,
  environment: HarnessEnvironment,
): Promise<LoopResults> {
  const payload = buildLoopResults(results, mode, totalCaseCount, environment);
  await writeFile(RESULTS_JSON, `${JSON.stringify(payload, null, 2)}\n`, "utf-8");
  return payload;
}

async function main(): Promise<void> {
  await mkdir(OUTPUT_DIR, { recursive: true });
  const totalCaseCount = generateTestCases().length;
  let cases = resolveLoopTestCases(LOOP_MODE);
  if (SUITE_ONLY?.length) {
    cases = cases.filter((c) => SUITE_ONLY.includes(c.name));
    if (cases.length === 0) {
      throw new Error(`SUITE_ONLY matched no cases: ${SUITE_ONLY.join(", ")}`);
    }
  }
  console.log(
    `Running ${cases.length}/${totalCaseCount} test cases (${LOOP_MODE}) → ${OUTPUT_DIR}`,
  );

  const browser = await chromium.launch();
  try {
    const environment = await captureEnvironment(browser);
    console.log(
      `Environment: chromium ${environment.chromium ?? "?"} · ${environment.libreoffice ?? "LibreOffice ?"} · node ${environment.node}`,
    );
    const results: CaseResult[] = [];
    for (const testCase of cases) {
      console.log(`  • ${testCase.name}`);
      results.push(await runCase(testCase, browser));
    }

    const payload = await writeResultsJson(results, LOOP_MODE, totalCaseCount, environment);
    printScorecard(results, payload.suite, LOOP_MODE);

    const hardFailures = results.filter((r) => !r.xmlPassed || r.error);
    const visualFailures = results.filter((r) => (r.mismatchedPixels ?? 0) > 0);

    if (hardFailures.length > 0) {
      process.exitCode = 1;
    } else if (STRICT_VISUAL && visualFailures.length > 0) {
      process.exitCode = 1;
    } else {
      process.exitCode = 0;
    }
  } finally {
    await browser.close();
  }
}

main();
