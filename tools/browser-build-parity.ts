import { access } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync } from "fflate";
import { chromium } from "playwright";
import { convertHtmlToDocx } from "../src/converter.js";
import { wrapHtml } from "../src/html-wrap.js";
import { VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX } from "../src/converter/constants.js";
import { generateTestCases } from "./generator.js";
import { writeGuardResult } from "./guard-result.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const BUNDLE_PATH = path.join(ROOT, "dist/browser/dom-docx.browser.js");
const VIEWPORT = { width: VIEWPORT_WIDTH_PX, height: VIEWPORT_HEIGHT_PX };

function normalizeWordXml(xml: string): string {
  return xml
    .replace(/\bw14:paraId="[^"]+"/g, 'w14:paraId="0"')
    .replace(/\bw14:textId="[^"]+"/g, 'w14:textId="0"')
    .replace(/\bw:rsidR="[^"]+"/g, 'w:rsidR="0"')
    .replace(/\bw:rsidRDefault="[^"]+"/g, 'w:rsidRDefault="0"')
    .replace(/\bw:rsidP="[^"]+"/g, 'w:rsidP="0"')
    .replace(/\bw:rsidRPr="[^"]+"/g, 'w:rsidRPr="0"')
    .replace(/\br:id="rId[^"]+"/g, 'r:id="rId0"')
    .replace(/\bId="[^"]+"/g, 'Id="0"');
}

function normalizedWordXmlParts(bytes: Uint8Array): Map<string, string> {
  const files = unzipSync(bytes);
  const parts = new Map<string, string>();
  for (const [name, data] of Object.entries(files)) {
    if (name.startsWith("word/") && name.endsWith(".xml")) {
      parts.set(name, normalizeWordXml(new TextDecoder().decode(data)));
    }
  }
  return parts;
}

function firstDivergingPart(a: Uint8Array, b: Uint8Array): string | null {
  const partsA = normalizedWordXmlParts(a);
  const partsB = normalizedWordXmlParts(b);
  if (partsA.size !== partsB.size) return `part count ${partsA.size} vs ${partsB.size}`;
  for (const [name, xmlA] of partsA) {
    if (partsB.get(name) !== xmlA) return name;
  }
  return null;
}

async function main(): Promise<void> {
  await access(BUNDLE_PATH);

  const cases = generateTestCases();
  const browser = await chromium.launch();

  try {
    let passed = 0;
    for (const testCase of cases) {
      const page = await browser.newPage({ viewport: VIEWPORT });
      try {
        await page.setContent(wrapHtml(testCase.html), { waitUntil: "domcontentloaded" });
        await page.addScriptTag({ path: BUNDLE_PATH });

        const nodeBuf = await convertHtmlToDocx(testCase.html, { styleSource: "computed", page });
        const byteList = await page.evaluate(
          async (html) =>
            Array.from(
              await window.domDocx!.convertHtmlToDocxUint8Array(html, { styleSource: "computed" }),
            ),
          testCase.html,
        );
        const browserBytes = new Uint8Array(byteList);

        const diverging = firstDivergingPart(nodeBuf, browserBytes);
        if (diverging) {
          console.error(`browser-build parity mismatch: ${testCase.name} (${diverging})`);
          await writeGuardResult({
            id: "browser-build-parity",
            label: "Browser bundle parity",
            passed,
            total: cases.length,
            ok: false,
            unit: "equivalent to Node computed-native",
            command: "npm run guard:browser-parity",
          });
          process.exitCode = 1;
          return;
        }
        passed += 1;
      } finally {
        await page.close();
      }
    }

    console.log(
      `browser-build parity: ${passed}/${cases.length} cases equivalent to Node computed-native`,
    );
    await writeGuardResult({
      id: "browser-build-parity",
      label: "Browser bundle parity",
      passed,
      total: cases.length,
      ok: true,
      unit: "equivalent to Node computed-native",
      command: "npm run guard:browser-parity",
    });
  } finally {
    await browser.close();
  }
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
