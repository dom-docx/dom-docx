/**
 * Structural image-spacing guard — CI-safe (inline path, no Playwright/LibreOffice).
 *
 * An image in the document flow must keep breathing room from adjacent blocks. Web
 * layouts separate figures with margins that the computed path faithfully zeroes
 * (flex/grid `gap`, container padding), which left the flat docx with the image
 * smashed against the next heading/paragraph. An image paragraph's before/after
 * spacing is floored to ~0.5em so it never renders flush against neighbors.
 *
 * This can't live in the visual suite: with `margin:0` the *browser* itself renders
 * the image smashed, so a browser-comparison case would penalize the floor (removing
 * it would score higher). Hence a structural assertion on the OOXML spacing.
 */
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { unzipSync } from "fflate";
import { validateFile } from "@xarsh/ooxml-validator";
import { convertHtmlToDocx } from "../src/converter.js";
import { writeGuardResult } from "./guard-result.js";
import { GUARDS_OUTPUT } from "./output-paths.js";
import { TEST_IMAGE_260x140, TEST_IMAGE_W, TEST_IMAGE_H } from "./test-image.js";

const OUT_DIR = path.join(GUARDS_OUTPUT, "image-spacing");
const FLOOR_TWIPS = 120; // pxToTwips(8)
const IMG = `<img src="${TEST_IMAGE_260x140}" width="${TEST_IMAGE_W}" height="${TEST_IMAGE_H}" alt="diagram">`;

// Zeroed margins everywhere — mimics the computed path (docs sites zero figure margins
// and space with flex/grid gap), the exact case that smashed the image against neighbors.
const ZEROED_HTML = `
  <p style="margin:0"><strong>Text before the image.</strong></p>
  <p style="margin:0">${IMG}</p>
  <h2 style="margin:0">Next section heading</h2>
  <p style="margin:0">${IMG}</p>
  <p style="margin:0">Body immediately after the image.</p>
`;
// A flex card holding an image must NOT be floored (cards manage their own tight rhythm).
const FLEX_HTML = `
  <div style="display:flex;gap:12px">
    <div style="flex:1;border:1px solid #ccc;padding:6px"><div style="height:140px">${IMG}</div></div>
    <div style="flex:1;border:1px solid #ccc;padding:6px"><div style="height:140px">${IMG}</div></div>
  </div>
`;

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

function documentXml(buffer: Buffer): string {
  const files = unzipSync(new Uint8Array(buffer));
  const data = files["word/document.xml"];
  return data ? new TextDecoder().decode(data) : "";
}

/** before/after twips of every paragraph containing an image drawing. */
function imageParagraphSpacings(xml: string): Array<{ before: number; after: number }> {
  const body = xml.slice(xml.indexOf("<w:body>"), xml.indexOf("</w:body>"));
  const out: Array<{ before: number; after: number }> = [];
  for (const p of body.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? []) {
    if (!/<w:drawing>/.test(p)) continue;
    out.push({
      before: Number(p.match(/w:before="(\d+)"/)?.[1] ?? 0),
      after: Number(p.match(/w:after="(\d+)"/)?.[1] ?? 0),
    });
  }
  return out;
}

async function main(): Promise<void> {
  await mkdir(OUT_DIR, { recursive: true });

  console.log("Image-spacing guard — structural checks (inline path):");

  const zeroedDocx = await convertHtmlToDocx(ZEROED_HTML);
  const zeroedPath = path.join(OUT_DIR, "zeroed.docx");
  await writeFile(zeroedPath, zeroedDocx);
  const zeroedImgs = imageParagraphSpacings(documentXml(zeroedDocx));
  check("both zeroed-margin images found", zeroedImgs.length === 2, `found ${zeroedImgs.length}`);
  check(
    "zeroed-margin image paragraphs are floored before+after",
    zeroedImgs.length > 0 && zeroedImgs.every((s) => s.before >= FLOOR_TWIPS && s.after >= FLOOR_TWIPS),
    JSON.stringify(zeroedImgs),
  );

  const validation = await validateFile(zeroedPath, { officeVersion: "Office2019" });
  check("OOXML schema valid", validation.ok, validation.errors.slice(0, 2).map((e) => e.description).join("; "));

  const flexDocx = await convertHtmlToDocx(FLEX_HTML);
  await writeFile(path.join(OUT_DIR, "flex.docx"), flexDocx);
  const flexImgs = imageParagraphSpacings(documentXml(flexDocx));
  check("flex-card images found", flexImgs.length >= 1, `found ${flexImgs.length}`);
  check(
    "flex-card image paragraphs are NOT floored (cards stay tight)",
    flexImgs.every((s) => s.before < FLOOR_TWIPS && s.after < FLOOR_TWIPS),
    JSON.stringify(flexImgs),
  );

  const ok = failures === 0;
  await writeGuardResult({
    id: "image-spacing",
    label: "Image spacing",
    passed: ok ? checksRun : checksRun - failures,
    total: checksRun,
    ok,
    unit: "flow images floored, flex images tight",
    command: "npm run guard:image-spacing",
  });

  console.log(ok ? `\nImage-spacing guard passed.` : `\n${failures} check(s) failed.`);
  if (!ok) process.exitCode = 1;
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
