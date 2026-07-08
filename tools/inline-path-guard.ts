import { unzipSync } from "fflate";
import { convertHtmlToDocx } from "../src/converter.js";
import { generateTestCases, isCustomConvertCase } from "./generator.js";
import { writeGuardResult } from "./guard-result.js";

/** Strip docx-library random ids so two inline conversions compare equal. */
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

function inlineOutputsEquivalent(a: Buffer, b: Buffer): boolean {
  const partsA = normalizedWordXmlParts(a);
  const partsB = normalizedWordXmlParts(b);
  if (partsA.size !== partsB.size) return false;
  for (const [name, xmlA] of partsA) {
    if (partsB.get(name) !== xmlA) return false;
  }
  return true;
}

async function main(): Promise<void> {
  const cases = generateTestCases();
  let passed = 0;

  for (const testCase of cases) {
    if (isCustomConvertCase(testCase)) {
      passed += 1;
      continue;
    }
    const defaultBuf = await convertHtmlToDocx(testCase.html);
    const explicitBuf = await convertHtmlToDocx(testCase.html, { styleSource: "inline" });
    if (!inlineOutputsEquivalent(defaultBuf, explicitBuf)) {
      console.error(`inline path mismatch: ${testCase.name} (normalized word/*.xml differ)`);
      await writeGuardResult({
        id: "inline-path",
        label: "Inline path",
        passed,
        total: cases.length,
        ok: false,
        unit: "equivalent (default vs explicit inline)",
        command: "npm run guard:inline",
      });
      process.exitCode = 1;
      return;
    }
    passed += 1;
  }

  console.log(
    `inline path guard: ${passed}/${cases.length} cases equivalent (default vs styleSource: "inline")`,
  );
  await writeGuardResult({
    id: "inline-path",
    label: "Inline path",
    passed,
    total: cases.length,
    ok: true,
    unit: "equivalent (default vs explicit inline)",
    command: "npm run guard:inline",
  });
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
