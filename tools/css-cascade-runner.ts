import { access, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, unzipSync } from "fflate";
import { validateFile } from "@xarsh/ooxml-validator";
import { docxToPdf } from "./docx2pdf.js";
import { chromium, type Browser } from "playwright";
import { convertHtmlToDocx } from "../src/converter.js";
import type { StyleSource } from "../src/converter/style-resolver.js";
import { wrapHtml } from "../src/html-wrap.js";
import { VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX } from "../src/converter/constants.js";
import {
  DEFAULT_COMPUTED_MIN_VISUAL,
  DEFAULT_MIN_COMPUTED_ADVANTAGE,
  generateCssCascadeCases,
  type CssCascadeCase,
} from "./css-cascade-cases.js";
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
import { CSS_CASCADE_OUTPUT } from "./output-paths.js";

const OUTPUT_ROOT = CSS_CASCADE_OUTPUT;
const RESULTS_JSON = path.join(OUTPUT_ROOT, "results.json");
const VIEWPORT = { width: VIEWPORT_WIDTH_PX, height: VIEWPORT_HEIGHT_PX };

interface CssCascadeCaseResult {
  case: CssCascadeCase;
  inline: CaseResult;
  computed: CaseResult;
  deltaVisual: number | null;
  passed: boolean;
  failures: string[];
}

export interface CssCascadeResults {
  version: 1;
  kind: "css-cascade";
  runAt: string;
  weights: typeof SCORE_WEIGHTS;
  thresholds: {
    defaultComputedMinVisual: number;
    defaultMinComputedAdvantage: number;
  };
  suite: {
    inline: SuiteRollup;
    computed: SuiteRollup;
    deltaVisual: number | null;
  };
  summary: {
    caseCount: number;
    passedCount: number;
    failedCount: number;
  };
  cases: CssCascadeCaseResult[];
}

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function captureHtmlPng(browser: Browser, html: string, outPath: string): Promise<string> {
  const page = await browser.newPage({ viewport: VIEWPORT });
  try {
    await page.setContent(wrapHtml(html), { waitUntil: "domcontentloaded" });
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

async function runStyleCase(
  testCase: CssCascadeCase,
  styleSource: StyleSource,
  browser: Browser,
  targetPng: string,
  htmlDisplayText: string,
): Promise<CaseResult> {
  const caseDir = path.join(OUTPUT_ROOT, testCase.name, styleSource);
  await mkdir(caseDir, { recursive: true });

  const docxPath = path.join(caseDir, "output.docx");
  const pdfPath = path.join(caseDir, "output.pdf");
  const docxPng = path.join(caseDir, "output_docx.png");
  const sideBySidePng = path.join(caseDir, "compare_side_by_side.png");
  const diffPath = path.join(OUTPUT_ROOT, `diff_${testCase.name}_${styleSource}.png`);

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
    const compileStart = performance.now();
    const docxBuffer = await convertHtmlToDocx(testCase.html, {
      styleSource,
      browser: styleSource === "computed" ? browser : undefined,
    });
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
    await pdfFirstPageToPng(pdfPath, docxPng);
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

/**
 * Structural stylesheet-leak check: colors that appear ONLY inside `<style>`
 * blocks must never reach the inline-path OOXML (the inline resolver reads
 * `style=""` attributes only). Deterministic and metric-independent — unlike
 * the old score-ceiling heuristic, it never needs recalibration when the
 * converter or the visual metric improves.
 */
async function findStylesheetLeaks(
  testCase: CssCascadeCase,
  inlineDocxPath: string,
): Promise<string[]> {
  const styleBlocks = [...testCase.html.matchAll(/<style>([\s\S]*?)<\/style>/g)]
    .map((m) => m[1]!)
    .join("\n");
  const body = testCase.html.replace(/<style>[\s\S]*?<\/style>/g, "");
  const colorsIn = (s: string) =>
    new Set([...s.matchAll(/#([0-9a-fA-F]{3,6})\b/g)].map((m) => m[1]!.toLowerCase()));
  const styleOnly = [...colorsIn(styleBlocks)].filter((c) => !colorsIn(body).has(c));
  if (styleOnly.length === 0) return [];

  const archive = unzipSync(new Uint8Array(await readFile(inlineDocxPath)));
  const entry = archive["word/document.xml"];
  if (!entry) return [];
  const xml = strFromU8(entry).toLowerCase();
  return styleOnly.filter((color) => {
    const expanded = color.length === 3 ? color.split("").map((ch) => ch + ch).join("") : color;
    return xml.includes(expanded) || xml.includes(color);
  });
}

function evaluateCase(
  testCase: CssCascadeCase,
  inline: CaseResult,
  computed: CaseResult,
  leakedColors: string[],
): { passed: boolean; failures: string[]; deltaVisual: number | null } {
  const failures: string[] = [];
  const computedMin = testCase.computedMinVisual ?? DEFAULT_COMPUTED_MIN_VISUAL;
  const minAdvantage = testCase.minComputedAdvantage ?? DEFAULT_MIN_COMPUTED_ADVANTAGE;

  const deltaVisual =
    computed.visualScore !== null &&
    computed.visualScore !== undefined &&
    inline.visualScore !== null &&
    inline.visualScore !== undefined
      ? computed.visualScore - inline.visualScore
      : null;

  if (computed.error) {
    failures.push(`computed error: ${computed.error}`);
  }
  if (!computed.xmlPassed) {
    failures.push("computed XML schema failed");
  }
  if (computed.visualScore === null || computed.visualScore === undefined) {
    failures.push("computed visual score missing");
  } else if (computed.visualScore < computedMin) {
    failures.push(
      `computed visual ${computed.visualScore.toFixed(2)}% below minimum ${computedMin}%`,
    );
  }
  if (minAdvantage > 0) {
    if (deltaVisual === null) {
      failures.push("could not compute inline vs computed delta");
    } else if (deltaVisual < minAdvantage) {
      failures.push(
        `computed advantage ${deltaVisual.toFixed(2)} pp below minimum ${minAdvantage} pp (inline ${inline.visualScore?.toFixed(2) ?? "—"}% · computed ${computed.visualScore?.toFixed(2) ?? "—"}%)`,
      );
    }
  }

  if (leakedColors.length > 0) {
    failures.push(
      `stylesheet-only colors leaked into inline OOXML: ${leakedColors.map((c) => `#${c}`).join(", ")}`,
    );
  }

  return { passed: failures.length === 0, failures, deltaVisual };
}

function fmt(value: number | null | undefined): string {
  return value !== null && value !== undefined ? value.toFixed(2) : "—";
}

function printReport(payload: CssCascadeResults): void {
  const line = "─".repeat(72);
  console.log(`\n${line}`);
  console.log("  CSS cascade suite — inline vs computed (stylesheet / class selectors)");
  console.log(line);
  console.log(`  Output: ${OUTPUT_ROOT}`);
  console.log(
    `  Thresholds: computed min ${payload.thresholds.defaultComputedMinVisual}% · advantage min ${payload.thresholds.defaultMinComputedAdvantage} pp (per-case overrides apply)\n`,
  );

  console.log("  Suite averages (adjusted visual)");
  console.log(
    `    inline    ${fmt(payload.suite.inline.visual)}% · engine ${fmt(payload.suite.inline.engine)}`,
  );
  console.log(
    `    computed  ${fmt(payload.suite.computed.visual)}% · engine ${fmt(payload.suite.computed.engine)}`,
  );
  const dVis = payload.suite.deltaVisual;
  console.log(
    `    Δ (computed − inline) ${dVis !== null ? (dVis >= 0 ? "+" : "") + dVis.toFixed(2) : "—"} pp`,
  );

  console.log(
    `\n  ${payload.summary.passedCount}/${payload.summary.caseCount} cases passed`,
  );

  for (const row of payload.cases) {
    const icon = row.passed ? "✓" : "✗";
    const d =
      row.deltaVisual !== null
        ? `${row.deltaVisual >= 0 ? "+" : ""}${row.deltaVisual.toFixed(2)}`
        : "—";
    console.log(`\n  [${icon}] ${row.case.name}`);
    console.log(`      ${row.case.notes}`);
    console.log(
      `      inline ${fmt(row.inline.visualScore)}% · computed ${fmt(row.computed.visualScore)}% · Δ ${d} pp`,
    );
    if (!row.passed) {
      for (const f of row.failures) {
        console.log(`      ↳ ${f}`);
      }
    }
    if (row.computed.sideBySidePath) {
      console.log(`      side-by-side (computed): ${row.computed.sideBySidePath}`);
    }
  }

  console.log(`\n${line}\n`);
}

async function main(): Promise<void> {
  await mkdir(OUTPUT_ROOT, { recursive: true });
  const cases = generateCssCascadeCases();

  console.log(`Running ${cases.length} CSS cascade cases → ${OUTPUT_ROOT}`);

  const browser = await chromium.launch();
  const caseResults: CssCascadeCaseResult[] = [];

  try {
    for (const testCase of cases) {
      console.log(`  • ${testCase.name}`);
      const caseDir = path.join(OUTPUT_ROOT, testCase.name);
      await mkdir(caseDir, { recursive: true });

      const htmlPath = path.join(caseDir, "source.html");
      const targetPng = path.join(caseDir, "target_html.png");
      await writeFile(htmlPath, wrapHtml(testCase.html), "utf-8");
      const htmlDisplayText = await captureHtmlPng(browser, testCase.html, targetPng);

      const inline = await runStyleCase(testCase, "inline", browser, targetPng, htmlDisplayText);
      const computed = await runStyleCase(
        testCase,
        "computed",
        browser,
        targetPng,
        htmlDisplayText,
      );

      const inlineDocxPath = path.join(OUTPUT_ROOT, testCase.name, "inline", "output.docx");
      const leakedColors = await findStylesheetLeaks(testCase, inlineDocxPath).catch(() => []);
      const evaluation = evaluateCase(testCase, inline, computed, leakedColors);
      caseResults.push({
        case: testCase,
        inline,
        computed,
        deltaVisual: evaluation.deltaVisual,
        passed: evaluation.passed,
        failures: evaluation.failures,
      });
    }
  } finally {
    await browser.close();
  }

  const inlineSuite = rollupSuite(
    caseResults.map((r) => ({
      visualMatch: r.inline.visualScore,
      editability: r.inline.editability,
      performance: r.inline.performanceScore,
      engineScore: r.inline.engineScore,
      compileMs: r.inline.compileMs,
      xmlPassed: r.inline.xmlPassed,
      error: r.inline.error,
    })),
  );
  const computedSuite = rollupSuite(
    caseResults.map((r) => ({
      visualMatch: r.computed.visualScore,
      editability: r.computed.editability,
      performance: r.computed.performanceScore,
      engineScore: r.computed.engineScore,
      compileMs: r.computed.compileMs,
      xmlPassed: r.computed.xmlPassed,
      error: r.computed.error,
    })),
  );

  const passedCount = caseResults.filter((r) => r.passed).length;
  const payload: CssCascadeResults = {
    version: 1,
    kind: "css-cascade",
    runAt: new Date().toISOString(),
    weights: SCORE_WEIGHTS,
    thresholds: {
      defaultComputedMinVisual: DEFAULT_COMPUTED_MIN_VISUAL,
      defaultMinComputedAdvantage: DEFAULT_MIN_COMPUTED_ADVANTAGE,
    },
    suite: {
      inline: inlineSuite,
      computed: computedSuite,
      deltaVisual:
        computedSuite.visual !== null && inlineSuite.visual !== null
          ? computedSuite.visual - inlineSuite.visual
          : null,
    },
    summary: {
      caseCount: cases.length,
      passedCount,
      failedCount: cases.length - passedCount,
    },
    cases: caseResults,
  };

  await writeFile(RESULTS_JSON, `${JSON.stringify(payload, null, 2)}\n`, "utf-8");
  printReport(payload);
  console.log(`Results JSON: ${RESULTS_JSON}`);

  if (passedCount < cases.length) {
    process.exitCode = 1;
  }
}

main();
