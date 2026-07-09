import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chromium, type Browser } from "playwright";
import { convertHtmlToDocx } from "../src/converter.js";
import { VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX } from "../src/converter/constants.js";
import { convertWithLibreOffice, convertWithWord } from "./docx2pdf.js";
import { captureEnvironment, type HarnessEnvironment } from "./environment.js";
import { generateTestCases, type TestCase } from "./generator.js";
import { wrapHtml } from "../src/html-wrap.js";
import { pdfFirstPageToPng } from "./pdf-raster.js";
import { extractPdfDisplayText } from "./pdf-text.js";
import { average } from "./scoring.js";
import {
  htmlDisplayTextOptionsForConvertOptions,
  htmlFragmentDisplayText,
} from "./text-content-fidelity.js";
import { comparePngs } from "./visual-compare.js";
import { SUITE_OUTPUT } from "./output-paths.js";

/**
 * Word-render spot check: the loop scores against LibreOffice, but the real consumer is
 * Microsoft Word — and the two disagree (LO ignores EXACT `w:line` in table rows, treats
 * exact `trHeight` as a minimum). This lane renders a handful of cases through BOTH and
 * reports the score delta, quantifying how much of the metric is LibreOffice-specific.
 *
 * Requires Microsoft Word (macOS, via AppleScript). Skips cleanly when unavailable.
 */

const OUTPUT_DIR = SUITE_OUTPUT;
const SPOTCHECK_DIR = path.join(OUTPUT_DIR, "word-spotcheck");
const RESULTS_JSON = path.join(OUTPUT_DIR, "word-spotcheck.json");
const VIEWPORT = { width: VIEWPORT_WIDTH_PX, height: VIEWPORT_HEIGHT_PX };

/** One anchor per major pattern: text, ordered list, table, shading, flex cards. */
const SPOTCHECK_CASE_NAMES = [
  "plain-paragraph",
  "simple-ordered-list",
  "simple-table-2x2",
  "table-row-backgrounds",
  "flex-row-horizontal",
] as const;

interface RendererScore {
  matchPercent: number;
  visualScore: number;
}

interface SpotcheckCaseResult {
  name: string;
  libreoffice: RendererScore | null;
  word: RendererScore | null;
  /** word.visualScore − libreoffice.visualScore; positive = Word renders closer to HTML. */
  visualDelta: number | null;
  error?: string;
}

interface SpotcheckResults {
  version: 1;
  runAt: string;
  environment: HarnessEnvironment;
  meanVisualDelta: number | null;
  cases: SpotcheckCaseResult[];
}

function resolveSpotcheckCases(): TestCase[] {
  const byName = new Map(generateTestCases().map((c) => [c.name, c]));
  const missing = SPOTCHECK_CASE_NAMES.filter((name) => !byName.has(name));
  if (missing.length > 0) {
    throw new Error(`Spot-check cases missing from generator: ${missing.join(", ")}`);
  }
  return SPOTCHECK_CASE_NAMES.map((name) => byName.get(name)!);
}

async function scoreRender(
  targetPng: string,
  pdfPath: string,
  pngPath: string,
  diffPath: string,
  testCase: TestCase,
  htmlDisplayText: string,
): Promise<RendererScore> {
  await pdfFirstPageToPng(pdfPath, pngPath);
  const docxDisplayText = await extractPdfDisplayText(pdfPath);
  const comparison = await comparePngs(targetPng, pngPath, diffPath, {
    htmlDisplayText,
    docxDisplayText,
    htmlFragment: testCase.html,
  });
  return { matchPercent: comparison.matchPercent, visualScore: comparison.visualScore };
}

async function runCase(testCase: TestCase, browser: Browser): Promise<SpotcheckCaseResult> {
  const caseDir = path.join(SPOTCHECK_DIR, testCase.name);
  await mkdir(caseDir, { recursive: true });

  const targetPng = path.join(caseDir, "target_html.png");
  const docxPath = path.join(caseDir, "output.docx");
  const loPdf = path.join(caseDir, "libreoffice.pdf");
  const wordPdf = path.join(caseDir, "word.pdf");

  const page = await browser.newPage({ viewport: VIEWPORT });
  try {
    await page.setContent(wrapHtml(testCase.html), { waitUntil: "networkidle" });
    await page.screenshot({ path: targetPng, fullPage: false });
  } finally {
    await page.close();
  }
  const htmlDisplayText = htmlFragmentDisplayText(
    testCase.html,
    htmlDisplayTextOptionsForConvertOptions(testCase.convertOptions),
  );

  const docxBuffer = await convertHtmlToDocx(testCase.html);
  await writeFile(docxPath, docxBuffer);

  const result: SpotcheckCaseResult = {
    name: testCase.name,
    libreoffice: null,
    word: null,
    visualDelta: null,
  };

  if (await convertWithLibreOffice(docxPath, loPdf)) {
    result.libreoffice = await scoreRender(
      targetPng,
      loPdf,
      path.join(caseDir, "libreoffice.png"),
      path.join(caseDir, "diff_libreoffice.png"),
      testCase,
      htmlDisplayText,
    );
  }

  if (await convertWithWord(docxPath, wordPdf)) {
    result.word = await scoreRender(
      targetPng,
      wordPdf,
      path.join(caseDir, "word.png"),
      path.join(caseDir, "diff_word.png"),
      testCase,
      htmlDisplayText,
    );
  }

  if (result.libreoffice && result.word) {
    result.visualDelta = result.word.visualScore - result.libreoffice.visualScore;
  }

  return result;
}

async function main(): Promise<void> {
  await mkdir(SPOTCHECK_DIR, { recursive: true });
  const cases = resolveSpotcheckCases();
  console.log(`Word spot check: ${cases.length} cases → ${SPOTCHECK_DIR}`);

  const browser = await chromium.launch();
  const results: SpotcheckCaseResult[] = [];
  try {
    const environment = await captureEnvironment(browser);

    for (const testCase of cases) {
      try {
        const result = await runCase(testCase, browser);
        results.push(result);
        const lo = result.libreoffice ? result.libreoffice.visualScore.toFixed(2) : "—";
        const word = result.word ? result.word.visualScore.toFixed(2) : "—";
        const delta =
          result.visualDelta !== null
            ? `${result.visualDelta >= 0 ? "+" : ""}${result.visualDelta.toFixed(2)}`
            : "—";
        console.log(`  • ${testCase.name}: LO ${lo}% · Word ${word}% · Δ ${delta} pp`);
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        results.push({
          name: testCase.name,
          libreoffice: null,
          word: null,
          visualDelta: null,
          error: message,
        });
        console.log(`  • ${testCase.name}: ERROR ${message}`);
      }
    }

    const wordAvailable = results.some((r) => r.word !== null);
    if (!wordAvailable) {
      console.log(
        "\nMicrosoft Word not available (macOS + Word required) — no deltas measured. Skipping.\n",
      );
      process.exitCode = 0;
      return;
    }

    const payload: SpotcheckResults = {
      version: 1,
      runAt: new Date().toISOString(),
      environment,
      meanVisualDelta: average(results.map((r) => r.visualDelta)),
      cases: results,
    };
    await writeFile(RESULTS_JSON, `${JSON.stringify(payload, null, 2)}\n`, "utf-8");

    console.log("\n─── Word vs LibreOffice render delta (adjusted visual, pp) ───");
    console.log(
      `  Mean Δ: ${payload.meanVisualDelta !== null ? payload.meanVisualDelta.toFixed(2) : "—"} (positive = Word closer to HTML reference)`,
    );
    console.log(`  Results JSON: ${RESULTS_JSON}\n`);

    process.exitCode = results.some((r) => r.error) ? 1 : 0;
  } finally {
    await browser.close();
  }
}

main();
