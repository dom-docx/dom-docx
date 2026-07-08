import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chromium, type Browser } from "playwright";
import { VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX } from "../src/converter/constants.js";
import { captureEnvironment, type HarnessEnvironment } from "./environment.js";
import { resolveLoopTestCases, type LoopCaseMode, type TestCase } from "./generator.js";
import { wrapHtml } from "../src/html-wrap.js";
import { pdfFirstPageToPng } from "./pdf-raster.js";
import { extractPdfDisplayText } from "./pdf-text.js";
import { average } from "./scoring.js";
import { htmlFragmentDisplayText } from "./text-content-fidelity.js";
import { comparePngs } from "./visual-compare.js";
import { SUITE_OUTPUT } from "./output-paths.js";

/**
 * Score-ceiling calibration: render the SAME HTML through both sides of the harness —
 * Chromium screenshot (reference) vs Chromium-printed PDF rasterized by pdf-to-img
 * (the DOCX-side pipeline, minus the converter and LibreOffice). No conversion happens,
 * so any mismatch is pure pipeline noise: font rasterization deltas between Chromium
 * and pdfjs, PDF print pagination, and subscore heuristics.
 *
 * The per-case `matchPercent`/`visualScore` here is the CEILING a perfect conversion
 * could reach. Interpret loop scores as a fraction of this, not of 100.
 */

const OUTPUT_DIR = SUITE_OUTPUT;
const CALIBRATION_DIR = path.join(OUTPUT_DIR, "calibration");
const RESULTS_JSON = path.join(OUTPUT_DIR, "calibration.json");
const LOOP_MODE: LoopCaseMode = process.argv.includes("--full") ? "full" : "priority";
const VIEWPORT = { width: VIEWPORT_WIDTH_PX, height: VIEWPORT_HEIGHT_PX };

interface CalibrationCaseResult {
  name: string;
  /** Raw pixelmatch ceiling (%) — max achievable matchPercent for this case. */
  rawCeiling: number;
  /** Adjusted-visual ceiling (%) — max achievable visualScore for this case. */
  adjustedCeiling: number;
  legibilityScore: number;
  backgroundBalanceScore: number;
  listMarkerFidelityScore: number;
  textContentFidelityScore: number | null;
  error?: string;
}

export interface CalibrationResults {
  version: 1;
  runAt: string;
  mode: LoopCaseMode;
  environment: HarnessEnvironment;
  meanRawCeiling: number | null;
  meanAdjustedCeiling: number | null;
  minRawCeiling: number | null;
  minAdjustedCeiling: number | null;
  cases: CalibrationCaseResult[];
}

async function runCase(testCase: TestCase, browser: Browser): Promise<CalibrationCaseResult> {
  const caseDir = path.join(CALIBRATION_DIR, testCase.name);
  await mkdir(caseDir, { recursive: true });

  const targetPng = path.join(caseDir, "target_html.png");
  const pdfPath = path.join(caseDir, "reference_print.pdf");
  const pdfPng = path.join(caseDir, "reference_print.png");
  const diffPath = path.join(caseDir, "diff.png");

  const page = await browser.newPage({ viewport: VIEWPORT });
  try {
    await page.setContent(wrapHtml(testCase.html), { waitUntil: "networkidle" });
    await page.screenshot({ path: targetPng, fullPage: false });
    await page.pdf({
      path: pdfPath,
      width: "8.5in",
      height: "11in",
      margin: { top: "0", right: "0", bottom: "0", left: "0" },
      printBackground: true,
    });
  } finally {
    await page.close();
  }

  await pdfFirstPageToPng(pdfPath, pdfPng);
  const pdfDisplayText = await extractPdfDisplayText(pdfPath);

  const comparison = await comparePngs(targetPng, pdfPng, diffPath, {
    htmlDisplayText: htmlFragmentDisplayText(testCase.html),
    docxDisplayText: pdfDisplayText,
    htmlFragment: testCase.html,
  });

  return {
    name: testCase.name,
    rawCeiling: comparison.matchPercent,
    adjustedCeiling: comparison.visualScore,
    legibilityScore: comparison.legibility.score,
    backgroundBalanceScore: comparison.backgroundBalance.score,
    listMarkerFidelityScore: comparison.listMarkerFidelity.score,
    textContentFidelityScore: comparison.textContentFidelity?.score ?? null,
  };
}

async function main(): Promise<void> {
  await mkdir(CALIBRATION_DIR, { recursive: true });
  const cases = resolveLoopTestCases(LOOP_MODE);
  console.log(
    `Calibration (score ceiling, no conversion): ${cases.length} cases (${LOOP_MODE}) → ${CALIBRATION_DIR}`,
  );

  const browser = await chromium.launch();
  const results: CalibrationCaseResult[] = [];
  try {
    const environment = await captureEnvironment(browser);
    for (const testCase of cases) {
      try {
        const result = await runCase(testCase, browser);
        results.push(result);
        console.log(
          `  • ${testCase.name}: raw ceiling ${result.rawCeiling.toFixed(2)}% · adjusted ${result.adjustedCeiling.toFixed(2)}%`,
        );
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        results.push({
          name: testCase.name,
          rawCeiling: 0,
          adjustedCeiling: 0,
          legibilityScore: 0,
          backgroundBalanceScore: 0,
          listMarkerFidelityScore: 0,
          textContentFidelityScore: null,
          error: message,
        });
        console.log(`  • ${testCase.name}: ERROR ${message}`);
      }
    }

    const ok = results.filter((r) => !r.error);
    const payload: CalibrationResults = {
      version: 1,
      runAt: new Date().toISOString(),
      mode: LOOP_MODE,
      environment,
      meanRawCeiling: average(ok.map((r) => r.rawCeiling)),
      meanAdjustedCeiling: average(ok.map((r) => r.adjustedCeiling)),
      minRawCeiling: ok.length > 0 ? Math.min(...ok.map((r) => r.rawCeiling)) : null,
      minAdjustedCeiling: ok.length > 0 ? Math.min(...ok.map((r) => r.adjustedCeiling)) : null,
      cases: results,
    };
    await writeFile(RESULTS_JSON, `${JSON.stringify(payload, null, 2)}\n`, "utf-8");

    console.log("\n─── Score ceiling (perfect conversion would score at most) ───");
    console.log(
      `  Mean: raw ${fmt(payload.meanRawCeiling)}% · adjusted ${fmt(payload.meanAdjustedCeiling)}%`,
    );
    console.log(
      `  Min:  raw ${fmt(payload.minRawCeiling)}% · adjusted ${fmt(payload.minAdjustedCeiling)}%`,
    );
    console.log(`  Results JSON: ${RESULTS_JSON}\n`);

    process.exitCode = results.some((r) => r.error) ? 1 : 0;
  } finally {
    await browser.close();
  }
}

function fmt(value: number | null): string {
  return value !== null ? value.toFixed(2) : "—";
}

main();
