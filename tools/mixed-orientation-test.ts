/**
 * Mixed orientation guard — per-section portrait/landscape from CSS page rules.
 *
 * Inline `style="page:landscape|portrait|Name"` and `@page { size: … }` work on
 * the default inline path. Class→page mappings (div.WordSection2 { page: … })
 * are honored only under styleSource: "computed".
 */
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { unzipSync } from "fflate";
import { convertHtmlToDocx, type ConvertOptions } from "../src/converter.js";
import { wrapHtml } from "../src/html-wrap.js";
import { docxToPdf } from "./docx2pdf.js";
import { writeGuardResult } from "./guard-result.js";
import { GUARDS_OUTPUT } from "./output-paths.js";

const OUT_DIR = path.join(GUARDS_OUTPUT, "mixed-orientation");

let failures = 0;
let checksRun = 0;

function check(name: string, cond: boolean, detail?: string): void {
  checksRun += 1;
  if (cond) console.log(`  ✓ ${name}`);
  else {
    console.error(`  ✗ ${name}${detail ? ` — ${detail}` : ""}`);
    failures += 1;
  }
}

async function documentXml(html: string, options?: ConvertOptions): Promise<string> {
  const buffer = await convertHtmlToDocx(html, { onWarning: null, ...options });
  const files = unzipSync(new Uint8Array(buffer));
  return new TextDecoder().decode(files["word/document.xml"]!);
}

function pageSizes(xml: string): Array<{ w: number; h: number; orient?: string }> {
  return [...xml.matchAll(/<w:pgSz w:w="(\d+)" w:h="(\d+)"(?: w:orient="(\w+)")?\s*\/>/g)].map(
    (m) => ({ w: Number(m[1]), h: Number(m[2]), orient: m[3] }),
  );
}

function sectPrCount(xml: string): number {
  return (xml.match(/<w:sectPr/g) ?? []).length;
}

const CLASS_MAPPING_HTML = `
<style>
  @page { size: 8.5in 11in; }
  @page WordSection2 { size: 11in 8.5in; }
  div.WordSection2 { page: WordSection2; }
</style>
<div class="WordSection1"><p>portrait-marker-alpha</p></div>
<div class="WordSection2"><p>landscape-marker-beta</p></div>
`;

const INLINE_THREE_SECTION_HTML = `
<p>portrait-before</p>
<div style="page:landscape"><p>landscape-middle</p></div>
<p>portrait-after</p>
`;

/** Three sections for optional PDF render: portrait → landscape → portrait (inline path). */
const RENDER_HTML = `
<p>render-marker-one</p>
<div style="page:landscape"><p>render-marker-two</p></div>
<p>render-marker-three</p>
`;

interface RenderedPage {
  widthIn: number;
  heightIn: number;
  text: string;
}

async function renderedPages(pdfPath: string): Promise<RenderedPage[]> {
  const { getDocument } = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const data = new Uint8Array(await readFile(pdfPath));
  const doc = await getDocument({ data, useSystemFonts: true }).promise;
  const pages: RenderedPage[] = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const { width, height } = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();
    pages.push({
      widthIn: width / 72,
      heightIn: height / 72,
      text: content.items.map((it) => ("str" in it ? it.str : "")).join(" "),
    });
  }
  return pages;
}

function isPortraitLetter(page: RenderedPage): boolean {
  return Math.abs(page.widthIn - 8.5) < 0.05 && Math.abs(page.heightIn - 11) < 0.05;
}

function isLandscapeLetter(page: RenderedPage): boolean {
  return Math.abs(page.widthIn - 11) < 0.05 && Math.abs(page.heightIn - 8.5) < 0.05;
}

async function main(): Promise<void> {
  console.log("inline path — class→page mapping ignored:");
  const inlineClassXml = await documentXml(CLASS_MAPPING_HTML);
  const inlineClassSizes = pageSizes(inlineClassXml);
  check(
    "single section (class mapping ignored on inline path)",
    sectPrCount(inlineClassXml) === 1 && inlineClassSizes.length === 1,
    `${sectPrCount(inlineClassXml)} sectPr · ${JSON.stringify(inlineClassSizes)}`,
  );
  check(
    "both markers in one section",
    inlineClassXml.includes("portrait-marker-alpha") && inlineClassXml.includes("landscape-marker-beta"),
  );

  console.log("\ninline path — style=\"page:landscape\" mid-body:");
  const threeXml = await documentXml(INLINE_THREE_SECTION_HTML);
  const threeSizes = pageSizes(threeXml);
  check("three sections emitted", sectPrCount(threeXml) === 3, String(sectPrCount(threeXml)));
  check(
    "portrait → landscape → portrait",
    threeSizes.length === 3 &&
      threeSizes[0]?.orient === "portrait" &&
      threeSizes[1]?.orient === "landscape" &&
      threeSizes[2]?.orient === "portrait",
    JSON.stringify(threeSizes),
  );
  check(
    "no blank paragraph + sectPr prefix at body start",
    !/<w:body[^>]*>\s*<w:p>\s*<w:r>\s*(?:<w:t\b[^>]*\/>|<w:t\b[^>]*>\s*<\/w:t>)\s*<\/w:r>\s*<\/w:p>\s*<w:p>\s*<w:pPr>\s*<w:sectPr>/.test(
      threeXml,
    ),
  );

  console.log("\ninline named-page reference:");
  const namedHtml = `
<style>@page Wide { size: 11in 8.5in; }</style>
<p>upright</p><div style="page:Wide"><p>sideways</p></div><p>upright-again</p>
`;
  const namedSizes = pageSizes(await documentXml(namedHtml));
  check(
    "style=\"page:Wide\" → portrait, landscape, portrait",
    namedSizes.length === 3 &&
      namedSizes[0]?.orient === "portrait" &&
      namedSizes[1]?.orient === "landscape" &&
      namedSizes[2]?.orient === "portrait",
    JSON.stringify(namedSizes),
  );

  console.log("\nexplicit config still wins:");
  const forced = pageSizes(await documentXml(CLASS_MAPPING_HTML, { orientation: "landscape" }));
  check(
    "orientation:'landscape' → one landscape section, CSS ignored",
    forced.length === 1 && forced[0]?.orient === "landscape",
    JSON.stringify(forced),
  );
  const plain = pageSizes(await documentXml(`<p>plain</p>`));
  check(
    "document without page CSS → one portrait section",
    plain.length === 1 && plain[0]?.orient === "portrait",
    JSON.stringify(plain),
  );

  console.log("\ncomputed path — class→page mapping (optional, needs Playwright + Chromium):");
  try {
    const { chromium } = await import("playwright");
    const { VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX } = await import("../src/converter/constants.js");
    const browser = await chromium.launch();
    try {
      const page = await browser.newPage({
        viewport: { width: VIEWPORT_WIDTH_PX, height: VIEWPORT_HEIGHT_PX },
      });
      await page.setContent(wrapHtml(CLASS_MAPPING_HTML), { waitUntil: "networkidle" });
      const computedXml = await documentXml(CLASS_MAPPING_HTML, { styleSource: "computed", page });
      const computedSizes = pageSizes(computedXml);
      check(
        "computed: two sections from class mapping",
        sectPrCount(computedXml) === 2 && computedSizes.length === 2,
        `${sectPrCount(computedXml)} sectPr`,
      );
      check(
        "computed: portrait then landscape",
        computedSizes[0]?.orient === "portrait" &&
          computedSizes[1]?.orient === "landscape" &&
          computedSizes[1]?.w === 15840 &&
          computedSizes[1]?.h === 12240,
        JSON.stringify(computedSizes),
      );
    } finally {
      await browser.close();
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.log(`  · skipped computed section (${msg.split("\n")[0]})`);
  }

  console.log("\nrendered PDF pages (optional, needs LibreOffice):");
  try {
    await mkdir(OUT_DIR, { recursive: true });
    const docxPath = path.join(OUT_DIR, "output.docx");
    const pdfPath = path.join(OUT_DIR, "output.pdf");
    await writeFile(docxPath, await convertHtmlToDocx(RENDER_HTML, { onWarning: null }));
    await docxToPdf(docxPath, pdfPath);
    const rendered = await renderedPages(pdfPath);
    check("three pages rendered", rendered.length === 3, `${rendered.length} pages`);
    check(
      "page orientations portrait → landscape → portrait",
      rendered.length === 3 &&
        isPortraitLetter(rendered[0]!) &&
        isLandscapeLetter(rendered[1]!) &&
        isPortraitLetter(rendered[2]!),
      rendered.map((p) => `${p.widthIn.toFixed(1)}x${p.heightIn.toFixed(1)}in`).join(" · "),
    );
    check(
      "page 1 carries first section content (no blank first page)",
      rendered[0]?.text.includes("render-marker-one") ?? false,
      rendered[0]?.text.slice(0, 80),
    );
    check(
      "each section marker on its own page",
      (rendered[1]?.text.includes("render-marker-two") ?? false) &&
        (rendered[2]?.text.includes("render-marker-three") ?? false),
    );
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.log(`  · skipped render section (${msg.split("\n")[0]})`);
  }

  const ok = failures === 0;
  await writeGuardResult({
    id: "mixed-orientation",
    label: "Mixed orientation",
    passed: ok ? checksRun : checksRun - failures,
    total: checksRun,
    ok,
    unit: "per-section w:pgSz + optional PDF",
    command: "npm run guard:mixed-orientation",
  });
  console.log(
    ok ? `\nMixed orientation guard passed (${checksRun} checks).` : `\n${failures} check(s) failed.`,
  );
  if (!ok) process.exitCode = 1;
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
