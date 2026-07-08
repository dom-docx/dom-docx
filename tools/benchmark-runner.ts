import { access, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { validateFile } from "@xarsh/ooxml-validator";
import { docxToPdf } from "./docx2pdf.js";
import { chromium, type Browser } from "playwright";
import { resolveBenchmarkLibraries } from "./benchmark/libraries.js";
import type { BenchmarkLibrary } from "./benchmark/types.js";
import { PAGE_MARGIN_PX, VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX } from "../src/converter/constants.js";
import { generateTestCases, type TestCase } from "./generator.js";
import {
  SCORE_WEIGHTS,
  compositeEngineScore,
  measureEditability,
  performanceScore,
  rollupSuite,
  type EditabilityBreakdown,
  type SuiteRollup,
} from "./scoring.js";
import type { CaseResult } from "./validator.js";
import { comparePngs } from "./visual-compare.js";
import { extractPdfDisplayText } from "./pdf-text.js";
import { htmlFragmentDisplayText } from "./text-content-fidelity.js";
import { BENCHMARK_OUTPUT, SUITE_OUTPUT } from "./output-paths.js";

const OUTPUT_ROOT = BENCHMARK_OUTPUT;
const DOM_DOCX_RESULTS = path.join(SUITE_OUTPUT, "results.json");
const VIEWPORT = { width: VIEWPORT_WIDTH_PX, height: VIEWPORT_HEIGHT_PX };

interface DomDocxCaseSnapshot {
  name: string;
  visualScore: number | null;
  editability: number | null;
  engineScore: number | null;
  xmlPassed?: boolean;
}

export interface BenchmarkResults {
  version: 1;
  kind: "benchmark";
  library: {
    id: string;
    npm: string;
    version: string;
    description: string;
  };
  runAt: string;
  weights: typeof SCORE_WEIGHTS;
  suite: SuiteRollup;
  comparison: {
    domDocxRunAt: string | null;
    domDocxSuite: SuiteRollup | null;
    delta: {
      visual: number | null;
      editability: number | null;
      engine: number | null;
    };
  };
  cases: Array<
    CaseResult & {
      domDocx?: DomDocxCaseSnapshot;
      deltaVisual?: number | null;
      deltaEngine?: number | null;
    }
  >;
}

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

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function loadDomDocxBaseline(): Promise<{
  runAt: string | null;
  suite: SuiteRollup | null;
  byName: Map<string, DomDocxCaseSnapshot>;
}> {
  if (!(await fileExists(DOM_DOCX_RESULTS))) {
    return { runAt: null, suite: null, byName: new Map() };
  }

  const raw = JSON.parse(await readFile(DOM_DOCX_RESULTS, "utf-8")) as {
    runAt?: string;
    suite?: SuiteRollup;
    cases?: Array<{
      name: string;
      visualScore: number | null;
      editability: number | null;
      engineScore: number | null;
      xmlPassed?: boolean;
    }>;
  };

  const byName = new Map<string, DomDocxCaseSnapshot>();
  for (const c of raw.cases ?? []) {
    byName.set(c.name, {
      name: c.name,
      visualScore: c.visualScore,
      editability: c.editability,
      engineScore: c.engineScore,
      xmlPassed: c.xmlPassed,
    });
  }

  return { runAt: raw.runAt ?? null, suite: raw.suite ?? null, byName };
}

async function runBenchmarkCase(
  testCase: TestCase,
  browser: Browser,
  libraryDir: string,
  library: BenchmarkLibrary,
): Promise<CaseResult> {
  const caseDir = path.join(libraryDir, testCase.name);
  await mkdir(caseDir, { recursive: true });

  const htmlPath = path.join(caseDir, "source.html");
  const targetPng = path.join(caseDir, "target_html.png");
  const docxPath = path.join(caseDir, "output.docx");
  const pdfPath = path.join(caseDir, "output.pdf");
  const docxPng = path.join(caseDir, "output_docx.png");
  const diffPath = path.join(libraryDir, `diff_${testCase.name}.png`);

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
    const wrapped = wrapHtml(testCase.html);
    await writeFile(htmlPath, wrapped, "utf-8");
    const htmlDisplayText = await captureHtmlPng(browser, testCase.html, targetPng);

    const compileStart = performance.now();
    const docxBuffer = await library.convertHtmlFragment(testCase.html);
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

    // html-to-docx emits sectPr before body content (schema-invalid but LO-readable).
    // Continue visual scoring so we can compare fidelity fairly.
    await docxToPdf(docxPath, pdfPath);
    if (!(await fileExists(pdfPath))) {
      throw new Error(`PDF not created at ${pdfPath}`);
    }
    await pdfFirstPageToPng(pdfPath, docxPng);
    const docxDisplayText = await extractPdfDisplayText(pdfPath);

    const comparison = await comparePngs(targetPng, docxPng, diffPath, {
      htmlDisplayText,
      docxDisplayText,
      htmlFragment: testCase.html,
    });
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

function printBenchmarkReport(
  library: BenchmarkLibrary,
  results: BenchmarkResults["cases"],
  suite: SuiteRollup,
  domDocxSuite: SuiteRollup | null,
  domDocxRunAt: string | null,
  libraryDir: string,
): void {
  const line = "─".repeat(72);
  console.log(`\n${line}`);
  console.log(`  Benchmark: ${library.npm}@${library.version} vs dom-docx`);
  console.log(line);
  console.log(`  ${library.description}`);
  console.log(`  Output: ${libraryDir}`);
  console.log(`  Weights: visual ${SCORE_WEIGHTS.visual} · editability ${SCORE_WEIGHTS.editability} · performance ${SCORE_WEIGHTS.performance}\n`);

  console.log("  Suite averages");
  console.log(
    `    ${library.id.padEnd(14)} visual ${fmt(suite.visual)} · editability ${fmt(suite.editability)} · engine ${fmt(suite.engine)}`,
  );
  if (domDocxSuite) {
    console.log(
      `    dom-docx       visual ${fmt(domDocxSuite.visual)} · editability ${fmt(domDocxSuite.editability)} · engine ${fmt(domDocxSuite.engine)}`,
    );
    const dVisual = delta(suite.visual, domDocxSuite.visual);
    const dEngine = delta(suite.engine, domDocxSuite.engine);
    console.log(
      `    Δ (${library.id} − dom-docx)  visual ${dVisual !== null ? (dVisual >= 0 ? "+" : "") + dVisual.toFixed(2) : "—"} · engine ${dEngine !== null ? (dEngine >= 0 ? "+" : "") + dEngine.toFixed(2) : "—"}`,
    );
    if (domDocxRunAt) {
      console.log(`    dom-docx baseline from: ${domDocxRunAt}`);
    }
  } else {
    console.log("    dom-docx baseline: not found — run npm run test:suite first for comparison");
  }

  console.log(`\n  Per-case (visual % · engine score · Δ visual vs dom-docx)`);
  const sorted = [...results].sort(
    (a, b) => (a.visualScore ?? 0) - (b.visualScore ?? 0),
  );
  for (const r of sorted) {
    const dom = r.domDocx;
    const dVis =
      r.deltaVisual !== null && r.deltaVisual !== undefined
        ? `${r.deltaVisual >= 0 ? "+" : ""}${r.deltaVisual.toFixed(2)}`
        : "—";
    const xml = r.xmlPassed ? "✓" : "✗";
    const err = r.error ? ` · ERROR: ${r.error}` : "";
    console.log(
      `    [${xml}] ${r.name.padEnd(28)} ${fmt(r.visualScore)}% · engine ${fmt(r.engineScore)} · Δ ${dVis}${err}`,
    );
    if (dom?.visualScore !== null && dom?.visualScore !== undefined) {
      console.log(
      `         dom-docx ${fmt(dom.visualScore)}% · engine ${fmt(dom.engineScore)}`,
    );
    }
  }

  console.log(`\n${line}\n`);
}

async function runBenchmarkForLibrary(
  library: BenchmarkLibrary,
  browser: Browser,
  domDocxBaseline: Awaited<ReturnType<typeof loadDomDocxBaseline>>,
): Promise<{ payload: BenchmarkResults; hadErrors: boolean }> {
  const libraryDir = path.join(OUTPUT_ROOT, library.id);
  await mkdir(libraryDir, { recursive: true });

  const cases = generateTestCases();
  console.log(`\nBenchmarking ${library.npm}@${library.version} on ${cases.length} regression cases`);
  console.log(`→ ${libraryDir}`);

  const results: CaseResult[] = [];
  for (const testCase of cases) {
    console.log(`  • ${testCase.name}`);
    results.push(await runBenchmarkCase(testCase, browser, libraryDir, library));
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

  const enrichedCases = results.map((r) => {
    const dom = domDocxBaseline.byName.get(r.name);
    return {
      ...r,
      domDocx: dom,
      deltaVisual: delta(r.visualScore, dom?.visualScore),
      deltaEngine: delta(r.engineScore, dom?.engineScore),
    };
  });

  const payload: BenchmarkResults = {
    version: 1,
    kind: "benchmark",
    library: {
      id: library.id,
      npm: library.npm,
      version: library.version,
      description: library.description,
    },
    runAt: new Date().toISOString(),
    weights: SCORE_WEIGHTS,
    suite,
    comparison: {
      domDocxRunAt: domDocxBaseline.runAt,
      domDocxSuite: domDocxBaseline.suite,
      delta: {
        visual: delta(suite.visual, domDocxBaseline.suite?.visual),
        editability: delta(suite.editability, domDocxBaseline.suite?.editability),
        engine: delta(suite.engine, domDocxBaseline.suite?.engine),
      },
    },
    cases: enrichedCases,
  };

  const resultsPath = path.join(OUTPUT_ROOT, `results-${library.id}.json`);
  await writeFile(resultsPath, `${JSON.stringify(payload, null, 2)}\n`, "utf-8");

  printBenchmarkReport(
    library,
    enrichedCases,
    suite,
    domDocxBaseline.suite,
    domDocxBaseline.runAt,
    libraryDir,
  );
  console.log(`Results JSON: ${resultsPath}`);

  return { payload, hadErrors: results.some((r) => r.error) };
}

async function main(): Promise<void> {
  const libraryArg = process.argv[2] ?? "all";
  const libraries = resolveBenchmarkLibraries(libraryArg);
  const domDocxBaseline = await loadDomDocxBaseline();

  const browser = await chromium.launch();
  const allPayloads: BenchmarkResults[] = [];
  let hadErrors = false;

  try {
    for (const library of libraries) {
      const { payload, hadErrors: libErrors } = await runBenchmarkForLibrary(
        library,
        browser,
        domDocxBaseline,
      );
      allPayloads.push(payload);
      hadErrors ||= libErrors;
    }
  } finally {
    await browser.close();
  }

  if (allPayloads.length === 1) {
    await writeFile(
      path.join(OUTPUT_ROOT, "results.json"),
      `${JSON.stringify(allPayloads[0], null, 2)}\n`,
      "utf-8",
    );
  } else if (allPayloads.length > 1) {
    await writeFile(
      path.join(OUTPUT_ROOT, "results.json"),
      `${JSON.stringify({ version: 1, kind: "benchmark-multi", runAt: new Date().toISOString(), libraries: allPayloads }, null, 2)}\n`,
      "utf-8",
    );
  }

  if (hadErrors) {
    process.exitCode = 1;
  }
}

main();
