import { access, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { validateFile } from "@xarsh/ooxml-validator";
import { docxToPdf } from "./docx2pdf.js";
import { chromium, type Browser } from "playwright";
import { convertHtmlToDocx } from "../src/converter.js";
import { wrapHtml } from "../src/html-wrap.js";
import { VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX } from "../src/converter/constants.js";
import { generateTestCases, type TestCase } from "./generator.js";
import {
  SCORE_WEIGHTS,
  compositeEngineScore,
  measureEditability,
  performanceScore,
  rollupSuite,
  type SuiteRollup,
} from "./scoring.js";
import type { CaseResult } from "./validator.js";
import { comparePngs, composeSideBySidePng } from "./visual-compare.js";
import { extractPdfDisplayText } from "./pdf-text.js";
import { htmlFragmentDisplayText } from "./text-content-fidelity.js";
import { BENCHMARK_OUTPUT, SUITE_OUTPUT } from "./output-paths.js";

const INLINE_RESULTS = path.join(SUITE_OUTPUT, "results.json");
const INLINE_OUTPUT_DIR = SUITE_OUTPUT;
const OUTPUT_ROOT = path.join(BENCHMARK_OUTPUT, "style-source");
const COMPUTED_DIR = path.join(OUTPUT_ROOT, "dom-docx-computed");
const NATIVE_DIR = path.join(OUTPUT_ROOT, "dom-docx-computed-native");
const RESULTS_JSON = path.join(OUTPUT_ROOT, "results-dom-docx-computed.json");
const VIEWPORT = { width: VIEWPORT_WIDTH_PX, height: VIEWPORT_HEIGHT_PX };

/**
 * "oracle"  — server-side: `convertHtmlToDocx` spawns its own page (setContent + render) per call.
 * "native"  — browser-side: snapshot computed styles from the ambient already-rendered page,
 *             the same way the in-browser deployment lane works (no second render).
 */
type ComputedMode = "oracle" | "native";

interface InlineCaseSnapshot {
  name: string;
  visualScore: number | null;
  editability: number | null;
  engineScore: number | null;
  xmlPassed?: boolean;
}

interface LaneCaseResult extends CaseResult {
  inline?: InlineCaseSnapshot;
  deltaVisual?: number | null;
  deltaEngine?: number | null;
  regression?: boolean;
}

export interface StyleSourceComparisonResults {
  version: 2;
  kind: "style-source-comparison";
  runAt: string;
  weights: typeof SCORE_WEIGHTS;
  inline: {
    runAt: string | null;
    suite: SuiteRollup | null;
  };
  computed: {
    suite: SuiteRollup;
  };
  computedNative: {
    suite: SuiteRollup;
  };
  comparison: {
    deltaVisual: number | null;
    deltaEditability: number | null;
    deltaEngine: number | null;
  };
  comparisonNative: {
    deltaVisual: number | null;
    deltaEditability: number | null;
    deltaEngine: number | null;
    /** native − oracle: should be ~0 visual (same snapshots) and a large positive perf gain. */
    deltaVisualVsOracle: number | null;
    avgCompileMsOracle: number | null;
    avgCompileMsNative: number | null;
  };
  cases: LaneCaseResult[];
  casesNative: LaneCaseResult[];
}

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function loadInlineBaseline(): Promise<{
  runAt: string | null;
  suite: SuiteRollup | null;
  byName: Map<string, InlineCaseSnapshot>;
}> {
  if (!(await fileExists(INLINE_RESULTS))) {
    return { runAt: null, suite: null, byName: new Map() };
  }

  const raw = JSON.parse(await readFile(INLINE_RESULTS, "utf-8")) as {
    runAt?: string;
    suite?: SuiteRollup;
    cases?: InlineCaseSnapshot[];
  };

  const byName = new Map<string, InlineCaseSnapshot>();
  for (const c of raw.cases ?? []) {
    byName.set(c.name, c);
  }

  return { runAt: raw.runAt ?? null, suite: raw.suite ?? null, byName };
}

async function captureHtmlPng(browser: Browser, html: string, outPath: string): Promise<string> {
  const page = await browser.newPage({ viewport: VIEWPORT });
  try {
    await page.setContent(wrapHtml(html), { waitUntil: "networkidle" });
    await page.screenshot({ path: outPath, fullPage: false });
    return htmlFragmentDisplayText(html);
  } finally {
    await page.close();
  }
}

async function pdfFirstPageToPng(pdfPath: string, outPath: string): Promise<void> {
  const { pdf } = await import("pdf-to-img");
  const document = await pdf(pdfPath, { scale: VIEWPORT.width / 612 });
  for await (const page of document) {
    await writeFile(outPath, page);
    return;
  }
  throw new Error(`No pages rendered from ${pdfPath}`);
}

async function runComputedCase(
  testCase: TestCase,
  browser: Browser,
  mode: ComputedMode = "oracle",
): Promise<CaseResult> {
  const isNative = mode === "native";
  const caseDir = path.join(isNative ? NATIVE_DIR : COMPUTED_DIR, testCase.name);
  await mkdir(caseDir, { recursive: true });

  const htmlPath = path.join(caseDir, "source.html");
  const inlineTargetPng = path.join(INLINE_OUTPUT_DIR, testCase.name, "target_html.png");
  const targetPng = path.join(caseDir, "target_html.png");
  const docxPath = path.join(caseDir, "output.docx");
  const pdfPath = path.join(caseDir, "output.pdf");
  const docxPng = path.join(caseDir, "output_docx.png");
  const sideBySidePng = path.join(caseDir, "compare_side_by_side.png");
  const diffPath = path.join(OUTPUT_ROOT, `diff_${testCase.name}_${mode}.png`);

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

    let htmlDisplayText: string;
    if (await fileExists(inlineTargetPng)) {
      htmlDisplayText = htmlFragmentDisplayText(testCase.html);
    } else {
      htmlDisplayText = await captureHtmlPng(browser, testCase.html, targetPng);
    }

    const referencePng = (await fileExists(inlineTargetPng)) ? inlineTargetPng : targetPng;

    let docxBuffer: Buffer;
    if (isNative) {
      // Browser-native: render once, then read computed styles from that ambient page.
      // The render is what a real browser already did to show the document, so the
      // measured conversion cost excludes it — it covers only the snapshot + build.
      const page = await browser.newPage({ viewport: VIEWPORT });
      try {
        await page.setContent(wrapHtml(testCase.html), { waitUntil: "networkidle" });
        const compileStart = performance.now();
        docxBuffer = await convertHtmlToDocx(testCase.html, { styleSource: "computed", page });
        result.compileMs = performance.now() - compileStart;
      } finally {
        await page.close();
      }
    } else {
      // Server-side oracle: convertHtmlToDocx spawns and renders its own page.
      const compileStart = performance.now();
      docxBuffer = await convertHtmlToDocx(testCase.html, { styleSource: "computed", browser });
      result.compileMs = performance.now() - compileStart;
    }
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
    await pdfFirstPageToPng(pdfPath, docxPng);
    const docxDisplayText = await extractPdfDisplayText(pdfPath);

    const comparison = await comparePngs(referencePng, docxPng, diffPath, {
      htmlDisplayText,
      docxDisplayText,
      htmlFragment: testCase.html,
    });
    await composeSideBySidePng(referencePng, docxPng, sideBySidePng);
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

function fmt(value: number | null | undefined): string {
  return value !== null && value !== undefined ? value.toFixed(2) : "—";
}

function delta(a: number | null | undefined, b: number | null | undefined): number | null {
  if (a === null || a === undefined || b === null || b === undefined) return null;
  return a - b;
}

function signed(value: number | null): string {
  return value !== null ? `${value >= 0 ? "+" : ""}${value.toFixed(2)}` : "—";
}

function printComparisonReport(payload: StyleSourceComparisonResults): void {
  const enrichedCases = payload.cases;
  const computedSuite = payload.computed.suite;
  const nativeSuite = payload.computedNative.suite;
  const inlineSuite = payload.inline.suite;
  const inlineRunAt = payload.inline.runAt;

  const line = "─".repeat(72);
  console.log(`\n${line}`);
  console.log("  Style source benchmark: inline vs computed-oracle vs computed-native");
  console.log(line);
  console.log(`  Output: ${OUTPUT_ROOT}`);
  console.log(
    `  Weights: visual ${SCORE_WEIGHTS.visual} · editability ${SCORE_WEIGHTS.editability} · performance ${SCORE_WEIGHTS.performance}\n`,
  );

  console.log("  Suite averages (adjusted visual)");
  console.log(
    `    inline            visual ${fmt(inlineSuite?.visual)} · engine ${fmt(inlineSuite?.engine)}`,
  );
  console.log(
    `    computed-oracle   visual ${fmt(computedSuite.visual)} · engine ${fmt(computedSuite.engine)} · ~${fmt(payload.comparisonNative.avgCompileMsOracle)} ms/case`,
  );
  console.log(
    `    computed-native   visual ${fmt(nativeSuite.visual)} · engine ${fmt(nativeSuite.engine)} · ~${fmt(payload.comparisonNative.avgCompileMsNative)} ms/case`,
  );
  console.log(
    `    Δ visual: oracle−inline ${signed(payload.comparison.deltaVisual)} pp · native−inline ${signed(payload.comparisonNative.deltaVisual)} pp · native−oracle ${signed(payload.comparisonNative.deltaVisualVsOracle)} pp`,
  );
  console.log(
    `    Δ engine: oracle−inline ${signed(payload.comparison.deltaEngine)} · native−inline ${signed(payload.comparisonNative.deltaEngine)} (native reclaims the Playwright perf penalty)`,
  );
  if (inlineRunAt) {
    console.log(`    inline baseline from: ${inlineRunAt}`);
  } else {
    console.log("    inline baseline: not found — run npm run test:suite first");
  }

  const regressions = enrichedCases.filter((c) => c.regression);
  const improvements = enrichedCases.filter(
    (c) => c.deltaVisual !== null && c.deltaVisual !== undefined && c.deltaVisual > 0.05,
  );
  const unchanged = enrichedCases.filter(
    (c) =>
      c.deltaVisual !== null &&
      c.deltaVisual !== undefined &&
      Math.abs(c.deltaVisual) <= 0.05,
  );

  console.log(`\n  Summary: ${improvements.length} improved · ${unchanged.length} ~unchanged · ${regressions.length} regressed`);

  if (regressions.length > 0) {
    console.log("\n  ⚠ Regressions (computed < inline — likely resolver mapping bug):");
    for (const r of regressions.sort(
      (a, b) => (a.deltaVisual ?? 0) - (b.deltaVisual ?? 0),
    )) {
      console.log(
        `    ${r.name.padEnd(28)} inline ${fmt(r.inline?.visualScore)}% → computed ${fmt(r.visualScore)}% (Δ ${fmt(r.deltaVisual)} pp)`,
      );
    }
  }

  console.log("\n  Per-case Δ adjusted visual (computed − inline)");
  const sorted = [...enrichedCases].sort(
    (a, b) => (b.deltaVisual ?? 0) - (a.deltaVisual ?? 0),
  );
  for (const r of sorted) {
    const dVis =
      r.deltaVisual !== null && r.deltaVisual !== undefined
        ? `${r.deltaVisual >= 0 ? "+" : ""}${r.deltaVisual.toFixed(2)}`
        : "—";
    const flag = r.regression ? " ⚠" : "";
    const xml = r.xmlPassed ? "✓" : "✗";
    const err = r.error ? ` · ERROR: ${r.error}` : "";
    console.log(
      `    [${xml}] ${r.name.padEnd(28)} inline ${fmt(r.inline?.visualScore)}% · computed ${fmt(r.visualScore)}% · Δ ${dVis}${flag}${err}`,
    );
  }

  console.log(`\n${line}\n`);
}

function rollup(results: CaseResult[]): SuiteRollup {
  return rollupSuite(
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
}

function enrich(
  results: CaseResult[],
  byName: Map<string, InlineCaseSnapshot>,
): LaneCaseResult[] {
  return results.map((r) => {
    const inline = byName.get(r.name);
    const deltaVisual = delta(r.visualScore, inline?.visualScore);
    return {
      ...r,
      inline,
      deltaVisual,
      deltaEngine: delta(r.engineScore, inline?.engineScore),
      regression: deltaVisual !== null && deltaVisual !== undefined && deltaVisual < -0.05,
    };
  });
}

function avgCompileMs(results: CaseResult[]): number | null {
  const vals = results.map((r) => r.compileMs).filter((v): v is number => v !== null);
  return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
}

async function main(): Promise<void> {
  await mkdir(OUTPUT_ROOT, { recursive: true });
  const inlineBaseline = await loadInlineBaseline();
  const cases = generateTestCases();

  console.log(
    `Running ${cases.length} cases · oracle → ${COMPUTED_DIR} · native → ${NATIVE_DIR}`,
  );

  const browser = await chromium.launch();
  const results: CaseResult[] = [];
  const nativeResults: CaseResult[] = [];

  try {
    for (const testCase of cases) {
      console.log(`  • ${testCase.name} [oracle]`);
      results.push(await runComputedCase(testCase, browser, "oracle"));
      console.log(`  • ${testCase.name} [native]`);
      nativeResults.push(await runComputedCase(testCase, browser, "native"));
    }
  } finally {
    await browser.close();
  }

  const computedSuite = rollup(results);
  const nativeSuite = rollup(nativeResults);
  const enrichedCases = enrich(results, inlineBaseline.byName);
  const enrichedNative = enrich(nativeResults, inlineBaseline.byName);

  const payload: StyleSourceComparisonResults = {
    version: 2,
    kind: "style-source-comparison",
    runAt: new Date().toISOString(),
    weights: SCORE_WEIGHTS,
    inline: {
      runAt: inlineBaseline.runAt,
      suite: inlineBaseline.suite,
    },
    computed: {
      suite: computedSuite,
    },
    computedNative: {
      suite: nativeSuite,
    },
    comparison: {
      deltaVisual: delta(computedSuite.visual, inlineBaseline.suite?.visual),
      deltaEditability: delta(computedSuite.editability, inlineBaseline.suite?.editability),
      deltaEngine: delta(computedSuite.engine, inlineBaseline.suite?.engine),
    },
    comparisonNative: {
      deltaVisual: delta(nativeSuite.visual, inlineBaseline.suite?.visual),
      deltaEditability: delta(nativeSuite.editability, inlineBaseline.suite?.editability),
      deltaEngine: delta(nativeSuite.engine, inlineBaseline.suite?.engine),
      deltaVisualVsOracle: delta(nativeSuite.visual, computedSuite.visual),
      avgCompileMsOracle: avgCompileMs(results),
      avgCompileMsNative: avgCompileMs(nativeResults),
    },
    cases: enrichedCases,
    casesNative: enrichedNative,
  };

  await writeFile(RESULTS_JSON, `${JSON.stringify(payload, null, 2)}\n`, "utf-8");
  printComparisonReport(payload);
  console.log(`Results JSON: ${RESULTS_JSON}`);

  if ([...results, ...nativeResults].some((r) => !r.xmlPassed || r.error)) {
    process.exitCode = 1;
  }
}

main();
