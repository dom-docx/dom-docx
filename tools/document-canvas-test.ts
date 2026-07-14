/**
 * Document-canvas guard — DOCX exports must not follow dark-theme colors.
 *
 * Word is a white-page medium. Computed styles from a dark-mode browser tab
 * often yield near-white `color` with a transparent fill; we drop those so text
 * stays readable. Light text on an intentionally dark shaded block is kept.
 *
 * The Node/Playwright computed path also forces `prefers-color-scheme: light`
 * before snapshotting — covered by an optional Playwright section (skipped when
 * Chromium is unavailable; CI runs the snapshot/remap checks only).
 */
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { unzipSync } from "fflate";
import { validateFile } from "@xarsh/ooxml-validator";
import { buildDocxUint8Array } from "../src/converter/build-docx.js";
import {
  relativeLuminance,
  remapComputedColorsForDocumentCanvas,
} from "../src/converter/document-canvas.js";
import {
  ComputedStyleResolver,
  parsedCssFromComputedRecord,
} from "../src/converter/style-resolver.js";
import { writeGuardResult } from "./guard-result.js";
import { GUARDS_OUTPUT } from "./output-paths.js";

const OUT_DIR = path.join(GUARDS_OUTPUT, "document-canvas");

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

function documentXml(buffer: Uint8Array): string {
  const files = unzipSync(buffer);
  const data = files["word/document.xml"];
  return data ? new TextDecoder().decode(data) : "";
}

function emptySnapshotStyles(overrides: Record<string, string>): Record<string, string> {
  return {
    color: "",
    backgroundColor: "rgba(0, 0, 0, 0)",
    display: "block",
    flexDirection: "row",
    gap: "0px",
    columnGap: "0px",
    rowGap: "0px",
    textAlign: "start",
    fontSize: "16px",
    fontWeight: "400",
    fontStyle: "normal",
    marginTop: "0px",
    marginRight: "0px",
    marginBottom: "0px",
    marginLeft: "0px",
    paddingTop: "0px",
    paddingRight: "0px",
    paddingBottom: "0px",
    paddingLeft: "0px",
    height: "auto",
    width: "auto",
    maxWidth: "none",
    borderTopWidth: "0px",
    borderTopColor: "rgb(0, 0, 0)",
    borderRightWidth: "0px",
    borderRightColor: "rgb(0, 0, 0)",
    borderBottomWidth: "0px",
    borderBottomColor: "rgb(0, 0, 0)",
    borderLeftWidth: "0px",
    borderLeftColor: "rgb(0, 0, 0)",
    breakBefore: "auto",
    breakAfter: "auto",
    ...overrides,
  };
}

async function main(): Promise<void> {
  await mkdir(OUT_DIR, { recursive: true });

  console.log("Document-canvas guard — luminance / remap helpers:");
  check("white luminance ≈ 1", Math.abs(relativeLuminance("ffffff") - 1) < 0.001);
  check("black luminance ≈ 0", relativeLuminance("000000") < 0.001);
  check("dark body text is not light", relativeLuminance("151515") < 0.55);
  check("dark-mode grey (#c7c7c7) is light", relativeLuminance("c7c7c7") >= 0.55);
  check("pure red keeps moderate luma", relativeLuminance("ee0000") < 0.55);

  const dropped = remapComputedColorsForDocumentCanvas({
    color: "ffffff",
    backgroundColor: undefined,
  });
  check("remap drops white text on transparent bg", dropped.color === undefined);

  const droppedRgb = parsedCssFromComputedRecord(
    emptySnapshotStyles({ color: "rgb(255, 255, 255)", backgroundColor: "rgba(0, 0, 0, 0)" }),
  );
  check("parsedCss drops white text on transparent bg", droppedRgb.color === undefined);

  const kept = remapComputedColorsForDocumentCanvas({
    color: "ffffff",
    backgroundColor: "1a1a2e",
  });
  check("remap keeps white text on dark fill", kept.color === "ffffff");

  const keptDark = remapComputedColorsForDocumentCanvas({
    color: "151515",
    backgroundColor: undefined,
  });
  check("remap keeps dark text on transparent bg", keptDark.color === "151515");

  console.log("\nDocument-canvas guard — OOXML via computed snapshots:");

  const whiteOnClear = await buildDocxUint8Array(
    "<p>Hello light theme</p>",
    ComputedStyleResolver.fromSnapshots([
      {
        path: "p[0]",
        styles: emptySnapshotStyles({
          color: "rgb(255, 255, 255)",
          backgroundColor: "rgba(0, 0, 0, 0)",
        }),
      },
    ]),
    undefined,
    undefined,
    null,
  );
  const whiteXml = documentXml(whiteOnClear);
  await writeFile(path.join(OUT_DIR, "white-on-clear.docx"), whiteOnClear);
  check("white-on-clear emits body text", /Hello light theme/.test(whiteXml));
  check(
    "white-on-clear does not stamp w:color ffffff",
    !/w:color[^>]*w:val="FFFFFF"/i.test(whiteXml),
  );

  const whiteOnDark = await buildDocxUint8Array(
    '<p style="background:#1a1a2e">Hello callout</p>',
    ComputedStyleResolver.fromSnapshots([
      {
        path: "p[0]",
        styles: emptySnapshotStyles({
          color: "rgb(255, 255, 255)",
          backgroundColor: "rgb(26, 26, 46)",
        }),
      },
    ]),
    undefined,
    undefined,
    null,
  );
  const darkXml = documentXml(whiteOnDark);
  await writeFile(path.join(OUT_DIR, "white-on-dark.docx"), whiteOnDark);
  check("white-on-dark keeps light foreground", /w:color[^>]*w:val="FFFFFF"/i.test(darkXml));
  check("white-on-dark keeps dark shading", /w:shd[^>]*w:fill="1A1A2E"/i.test(darkXml));

  const validation = await validateFile(path.join(OUT_DIR, "white-on-clear.docx"), {
    officeVersion: "Office2019",
  });
  check("OOXML schema valid", validation.ok, validation.errors?.[0]?.description);

  // Optional: Playwright forces light scheme so dark-media CSS never reaches snapshots.
  console.log("\nDocument-canvas guard — Playwright colorScheme light (optional):");
  try {
    const { chromium } = await import("playwright");
    const browser = await chromium.launch({ headless: true });
    try {
      const page = await browser.newPage({
        viewport: { width: 800, height: 600 },
        colorScheme: "dark",
      });
      await page.setContent(
        `<!DOCTYPE html><html><head><style>
          p { color: #111; }
          @media (prefers-color-scheme: dark) {
            p { color: #ffffff; }
          }
        </style></head><body><p>Scheme probe</p></body></html>`,
        { waitUntil: "networkidle" },
      );

      const darkLuma = await page.evaluate(() => getComputedStyle(document.querySelector("p")!).color);
      check(
        "prefers-dark tab would paint white text",
        darkLuma === "rgb(255, 255, 255)",
        darkLuma,
      );

      const { convertHtmlToDocx } = await import("../src/converter.js");
      const buf = await convertHtmlToDocx(await page.locator("body").innerHTML(), {
        styleSource: "computed",
        page,
        rootSelector: "body",
        onWarning: null,
      });
      await writeFile(path.join(OUT_DIR, "playwright-light.docx"), buf);
      const xml = documentXml(new Uint8Array(buf));
      check("Node computed path still emits text", /Scheme probe/.test(xml));
      check(
        "Node computed path does not stamp dark-scheme white",
        !/w:color[^>]*w:val="FFFFFF"/i.test(xml),
      );
    } finally {
      await browser.close();
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.log(`  · skipped Playwright section (${msg.split("\n")[0]})`);
  }

  await writeGuardResult({
    id: "document-canvas",
    label: "Document canvas colors",
    passed: checksRun - failures,
    total: checksRun,
    ok: failures === 0,
    unit: "structural checks",
    command: "npm run guard:document-canvas",
  });

  if (failures > 0) {
    console.error(`\n${failures} check(s) failed`);
    process.exit(1);
  }
  console.log(`\nAll ${checksRun} checks passed`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
