/**
 * Wild-HTML corpus runner — launch-readiness scoring on real-world pages the
 * converter was NOT tuned on (email templates, legacy table layouts, wiki
 * tables, book prose, spec text, rendered markdown).
 *
 * Reads `internal/wild-corpus/<case>/fragment.html` (build with
 * `tsx tools/wild-corpus-build.ts`), converts each fragment, renders both
 * sides, scores with the same metrics as the suite, and writes:
 *   - output/wild/results.json
 *   - output/wild/<case>/compare_side_by_side.png (+ diff, docx, pdf)
 *   - output/wild/labeling.html — hand-label with keys 1/2/3 (NO scores shown,
 *     so the labeler isn't anchored); Export downloads wild-labels.json →
 *     save to internal/research/wild-labels.json for metric-concordance work.
 *
 * Run: npm run research:wild
 */
import { access, mkdir, readdir, readFile, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { validateFile } from "@xarsh/ooxml-validator";
import { chromium, type Browser } from "playwright";
import { convertHtmlToDocx } from "../src/converter.js";
import { PAGE_MARGIN_PX, VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX } from "../src/converter/constants.js";
import { docxToPdf } from "./docx2pdf.js";
import { pdfToContentFlowPng } from "./pdf-raster.js";
import { extractPdfDisplayText } from "./pdf-text.js";
import { htmlFragmentDisplayText } from "./text-content-fidelity.js";
import {
  compositeEngineScore,
  measureEditability,
  performanceScore,
} from "./scoring.js";
import type { CaseResult } from "./validator.js";
import { comparePngs, composeSideBySidePng } from "./visual-compare.js";
import { WILD_OUTPUT } from "./output-paths.js";
import { WILD_CORPUS_DIR } from "./wild-corpus-build.js";

const VIEWPORT = { width: VIEWPORT_WIDTH_PX, height: VIEWPORT_HEIGHT_PX };
const RESULTS_JSON = path.join(WILD_OUTPUT, "results.json");

interface WildCaseMeta {
  name: string;
  title: string;
  url: string;
  description: string;
  fetchedAt?: string;
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

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function loadCorpus(): Promise<Array<{ meta: WildCaseMeta; html: string }>> {
  if (!existsSync(WILD_CORPUS_DIR)) {
    throw new Error(
      `No wild corpus at ${WILD_CORPUS_DIR} — build it first: tsx tools/wild-corpus-build.ts`,
    );
  }
  const entries = await readdir(WILD_CORPUS_DIR, { withFileTypes: true });
  const cases: Array<{ meta: WildCaseMeta; html: string }> = [];
  for (const entry of entries) {
    if (!entry.isDirectory() || entry.name.startsWith(".")) continue;
    const dir = path.join(WILD_CORPUS_DIR, entry.name);
    const fragmentPath = path.join(dir, "fragment.html");
    if (!(await fileExists(fragmentPath))) continue;
    const html = await readFile(fragmentPath, "utf-8");
    let meta: WildCaseMeta = {
      name: entry.name,
      title: entry.name,
      url: "",
      description: "",
    };
    try {
      meta = { ...meta, ...JSON.parse(await readFile(path.join(dir, "meta.json"), "utf-8")) };
    } catch {
      // meta optional
    }
    cases.push({ meta, html });
  }
  return cases.sort((a, b) => a.meta.name.localeCompare(b.meta.name));
}

async function runWildCase(
  meta: WildCaseMeta,
  html: string,
  browser: Browser,
): Promise<CaseResult> {
  const caseDir = path.join(WILD_OUTPUT, meta.name);
  await mkdir(caseDir, { recursive: true });

  const targetPng = path.join(caseDir, "target_html.png");
  const docxPath = path.join(caseDir, "output.docx");
  const pdfPath = path.join(caseDir, "output.pdf");
  const docxPng = path.join(caseDir, "output_docx.png");
  const sideBySidePng = path.join(caseDir, "compare_side_by_side.png");
  const diffPath = path.join(WILD_OUTPUT, `diff_${meta.name}.png`);

  const result: CaseResult = {
    name: meta.name,
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
    await writeFile(path.join(caseDir, "source.html"), wrapHtml(html), "utf-8");

    const page = await browser.newPage({ viewport: VIEWPORT });
    try {
      await page.setContent(wrapHtml(html), { waitUntil: "networkidle" });
      await page.screenshot({ path: targetPng, fullPage: true });
    } finally {
      await page.close();
    }
    const htmlDisplayText = htmlFragmentDisplayText(html);

    const compileStart = performance.now();
    const docxBuffer = await convertHtmlToDocx(html);
    result.compileMs = performance.now() - compileStart;
    result.performanceScore = performanceScore(result.compileMs);
    await writeFile(docxPath, docxBuffer);

    const validation = await validateFile(docxPath, { officeVersion: "Office2019" });
    result.xmlPassed = validation.ok;
    if (!validation.ok) {
      result.xmlErrors = validation.errors
        .map((e) => `[${e.errorType}] ${e.path}: ${e.description}`)
        .slice(0, 5);
    }

    const editability = await measureEditability(docxPath);
    result.editability = editability.score;
    result.editabilityBreakdown = editability;

    await docxToPdf(docxPath, pdfPath);
    if (!(await fileExists(pdfPath))) {
      throw new Error(`PDF not created at ${pdfPath}`);
    }
    await pdfToContentFlowPng(pdfPath, docxPng);
    const docxDisplayText = await extractPdfDisplayText(pdfPath);

    const comparison = await comparePngs(targetPng, docxPng, diffPath, {
      htmlDisplayText,
      docxDisplayText,
      htmlFragment: html,
    });
    await composeSideBySidePng(targetPng, docxPng, sideBySidePng);

    result.visualCompared = true;
    result.mismatchedPixels = comparison.mismatchedPixels;
    result.matchPercent = comparison.matchPercent;
    result.legibilityScore = comparison.legibility.score;
    result.backgroundBalanceScore = comparison.backgroundBalance.score;
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

/** Deterministic shuffle (mulberry32) — stable labeling order across re-runs. */
function seededShuffle<T>(items: T[], seed: number): T[] {
  let state = seed >>> 0;
  const rand = () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  const out = [...items];
  for (let i = out.length - 1; i > 0; i--) {
    const j = Math.floor(rand() * (i + 1));
    [out[i], out[j]] = [out[j], out[i]];
  }
  return out;
}

/** Self-contained labeling page — shows NO scores so labels stay unanchored. */
async function writeLabelingPage(caseNames: string[]): Promise<string> {
  const cases = seededShuffle(
    caseNames.filter((name) =>
      existsSync(path.join(WILD_OUTPUT, name, "compare_side_by_side.png")),
    ),
    20260703,
  );

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>dom-docx wild-corpus labeling</title>
<style>
  body { margin: 0; font-family: system-ui, sans-serif; background: #1e1e1e; color: #ddd; }
  header { position: sticky; top: 0; background: #2a2a2a; padding: 10px 16px; display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
  header b { color: #fff; }
  button { font-size: 14px; padding: 6px 14px; border-radius: 6px; border: 1px solid #555; background: #333; color: #ddd; cursor: pointer; }
  .active-1 { background: #14532d; border-color: #22c55e; }
  .active-2 { background: #713f12; border-color: #eab308; }
  .active-3 { background: #7f1d1d; border-color: #ef4444; }
  input { flex: 1; min-width: 220px; font-size: 13px; padding: 6px 8px; border-radius: 6px; border: 1px solid #555; background: #262626; color: #ddd; }
  img { display: block; width: 100%; max-width: 1700px; margin: 0 auto; background: #fff; }
  .hint { color: #888; font-size: 12px; }
</style>
</head>
<body>
<header>
  <b id="casename"></b>
  <span id="progress" class="hint"></span>
  <button id="b1" onclick="rate(1)">1 · looks right</button>
  <button id="b2" onclick="rate(2)">2 · minor issues</button>
  <button id="b3" onclick="rate(3)">3 · broken</button>
  <input id="note" placeholder="note (optional)" onchange="saveNote()">
  <button onclick="prev()">←</button>
  <button onclick="next()">→</button>
  <button onclick="exportLabels()">Export</button>
  <span class="hint">keys: 1/2/3 rate · arrows navigate</span>
</header>
<main><img id="img" src="" alt="side by side render"></main>
<script>
const CASES = ${JSON.stringify(cases)};
const KEY = "dom-docx-wild-labels-v1";
let idx = 0;
let labels = JSON.parse(localStorage.getItem(KEY) || "{}");
function firstUnlabeled() {
  const i = CASES.findIndex((c) => !labels[c] || !labels[c].rating);
  return i < 0 ? 0 : i;
}
function render() {
  const name = CASES[idx];
  document.getElementById("img").src = name + "/compare_side_by_side.png";
  document.getElementById("casename").textContent = name;
  const done = CASES.filter((c) => labels[c] && labels[c].rating).length;
  document.getElementById("progress").textContent = (idx + 1) + "/" + CASES.length + " · " + done + " labeled";
  document.getElementById("note").value = (labels[name] && labels[name].note) || "";
  for (const r of [1, 2, 3]) {
    const b = document.getElementById("b" + r);
    b.className = labels[name] && labels[name].rating === r ? "active-" + r : "";
  }
  window.scrollTo(0, 0);
}
function save() { localStorage.setItem(KEY, JSON.stringify(labels)); }
function rate(r) {
  labels[CASES[idx]] = Object.assign({}, labels[CASES[idx]], { rating: r });
  save();
  next();
}
function saveNote() {
  labels[CASES[idx]] = Object.assign({}, labels[CASES[idx]], { note: document.getElementById("note").value });
  save();
  render();
}
function next() { if (idx < CASES.length - 1) idx++; render(); }
function prev() { if (idx > 0) idx--; render(); }
function exportLabels() {
  const payload = {
    version: 1,
    corpus: "wild",
    labeledAt: new Date().toISOString(),
    scale: { 1: "looks right", 2: "minor issues", 3: "broken" },
    labels: labels,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "wild-labels.json";
  a.click();
}
document.addEventListener("keydown", (e) => {
  if (e.target.tagName === "INPUT") return;
  if (e.key === "1") rate(1);
  else if (e.key === "2") rate(2);
  else if (e.key === "3") rate(3);
  else if (e.key === "ArrowRight") next();
  else if (e.key === "ArrowLeft") prev();
});
idx = firstUnlabeled();
render();
</script>
</body>
</html>
`;

  const outPath = path.join(WILD_OUTPUT, "labeling.html");
  await writeFile(outPath, html, "utf-8");
  return outPath;
}

async function main(): Promise<void> {
  const corpus = await loadCorpus();
  if (corpus.length === 0) {
    throw new Error(`Wild corpus is empty — run: tsx tools/wild-corpus-build.ts`);
  }
  await mkdir(WILD_OUTPUT, { recursive: true });

  const browser = await chromium.launch();
  const results: CaseResult[] = [];
  try {
    for (const { meta, html } of corpus) {
      results.push(await runWildCase(meta, html, browser));
    }
  } finally {
    await browser.close();
  }

  const line = "─".repeat(60);
  console.log(`\n${line}`);
  console.log("  Wild-HTML Corpus — Conversion Report");
  console.log(line);
  console.log(`  Output: ${WILD_OUTPUT}`);
  console.log("  (Real-world pages — NOT part of the regression suite)\n");

  for (const { meta } of corpus) {
    const r = results.find((x) => x.name === meta.name);
    if (!r) continue;
    console.log(`  [${meta.name}] ${meta.title}`);
    if (r.error) {
      console.log(`    [!] Error: ${r.error}\n`);
      continue;
    }
    console.log(`    [${r.xmlPassed ? "✓" : "✗"}] XML Schema`);
    if (r.editabilityBreakdown) {
      console.log(
        `    [•] Editability: ${r.editabilityBreakdown.score.toFixed(0)} (1×1 tables: ${r.editabilityBreakdown.oneByOneTables})`,
      );
    }
    if (r.visualScore !== null) {
      console.log(`    [•] Visual (layout-based): ${r.visualScore.toFixed(2)}%`);
    }
    if (r.textContentFidelityScore !== undefined && r.textContentFidelityScore < 99.5) {
      console.log(
        `        Text content: ${r.textContentFidelityScore.toFixed(2)}% (${r.missingTokenUnits ?? 0} missing · ${r.extraTokenUnits ?? 0} extra)`,
      );
    }
    if (r.engineScore !== null) {
      console.log(`    [★] Engine score: ${r.engineScore.toFixed(2)}`);
    }
    console.log("");
  }

  const scored = results.filter((r) => r.visualScore !== null);
  const avg = (xs: number[]) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0);
  console.log(line);
  console.log(
    `  Wild avg: visual ${avg(scored.map((r) => r.visualScore!)).toFixed(2)} · engine ${avg(
      results.filter((r) => r.engineScore !== null).map((r) => r.engineScore!),
    ).toFixed(2)} · ${scored.length}/${results.length} compared`,
  );
  console.log(line);

  await writeFile(
    RESULTS_JSON,
    `${JSON.stringify(
      { version: 1, runAt: new Date().toISOString(), corpus: "wild", cases: results },
      null,
      2,
    )}\n`,
    "utf-8",
  );

  const labelingPath = await writeLabelingPage(results.map((r) => r.name));
  console.log(`  Results JSON: ${RESULTS_JSON}`);
  console.log(`  Labeling page: ${labelingPath}`);
  console.log("  Open it, rate 1/2/3, Export → save as internal/research/wild-labels.json\n");
}

main();
