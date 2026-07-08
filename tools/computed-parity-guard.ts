import { unzipSync } from "fflate";
import { chromium } from "playwright";
import { convertHtmlToDocx } from "../src/converter.js";
import { wrapHtml } from "../src/html-wrap.js";
import { VIEWPORT_HEIGHT_PX, VIEWPORT_WIDTH_PX } from "../src/converter/constants.js";
import { generateTestCases } from "./generator.js";
import { writeGuardResult } from "./guard-result.js";

const VIEWPORT = { width: VIEWPORT_WIDTH_PX, height: VIEWPORT_HEIGHT_PX };

/** Strip docx-library random ids so two conversions of the same input compare equal. */
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

function normalizedWordXmlParts(buffer: Buffer): Map<string, string> {
  const files = unzipSync(new Uint8Array(buffer));
  const parts = new Map<string, string>();
  for (const [name, data] of Object.entries(files)) {
    if (name.startsWith("word/") && name.endsWith(".xml")) {
      parts.set(name, normalizeWordXml(new TextDecoder().decode(data)));
    }
  }
  return parts;
}

function firstDivergingPart(a: Buffer, b: Buffer): string | null {
  const partsA = normalizedWordXmlParts(a);
  const partsB = normalizedWordXmlParts(b);
  if (partsA.size !== partsB.size) return `part count ${partsA.size} vs ${partsB.size}`;
  for (const [name, xmlA] of partsA) {
    if (partsB.get(name) !== xmlA) return name;
  }
  return null;
}

/**
 * The browser-native path (computed styles read from an ambient already-rendered page)
 * and the server-side oracle (a freshly spawned page) feed the same snapshots into the
 * same pure OOXML builder. For identical input HTML they MUST emit byte-identical
 * normalized word/*.xml. A divergence isolates a bug to snapshot extraction — the only
 * thing that differs between the two lanes.
 */
async function main(): Promise<void> {
  const cases = generateTestCases();
  const browser = await chromium.launch();
  let passed = 0;

  try {
    for (const testCase of cases) {
      const oracleBuf = await convertHtmlToDocx(testCase.html, {
        styleSource: "computed",
        browser,
      });

      const page = await browser.newPage({ viewport: VIEWPORT });
      let nativeBuf: Buffer;
      try {
        await page.setContent(wrapHtml(testCase.html), { waitUntil: "networkidle" });
        nativeBuf = await convertHtmlToDocx(testCase.html, { styleSource: "computed", page });
      } finally {
        await page.close();
      }

      const diverging = firstDivergingPart(oracleBuf, nativeBuf);
      if (diverging) {
        console.error(`computed parity mismatch: ${testCase.name} (diverges in ${diverging})`);
        await writeGuardResult({
          id: "computed-parity",
          label: "Computed parity (oracle vs native)",
          passed,
          total: cases.length,
          ok: false,
          unit: "byte-identical",
          command: "npm run guard:computed-parity",
        });
        process.exitCode = 1;
        return;
      }
      passed += 1;
    }
  } finally {
    await browser.close();
  }

  console.log(
    `computed parity guard: ${passed}/${cases.length} cases byte-identical (oracle vs browser-native)`,
  );
  await writeGuardResult({
    id: "computed-parity",
    label: "Computed parity (oracle vs native)",
    passed,
    total: cases.length,
    ok: true,
    unit: "byte-identical",
    command: "npm run guard:computed-parity",
  });
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
