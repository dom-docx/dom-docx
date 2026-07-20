/**
 * Page-number field guard — verifies that {page} / {pages} tokens in footerHtml /
 * headerHtml and the pageNumber option produce correct OOXML PAGE / NUMPAGES fields.
 */
import { unzipSync } from "fflate";
import { convertHtmlToDocx } from "../src/converter.js";
import { writeGuardResult } from "./guard-result.js";

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

async function footerXml(options: Parameters<typeof convertHtmlToDocx>[1]): Promise<string> {
  const buf = await convertHtmlToDocx("<p>body</p>", { onWarning: null, ...options });
  const files = unzipSync(new Uint8Array(buf));
  const footer = files["word/footer1.xml"];
  return footer ? new TextDecoder().decode(footer) : "";
}

const PAGE_FIELD = /<w:instrText[^>]*>\s*PAGE\s*<\/w:instrText>/;
const NUMPAGES_FIELD = /<w:instrText[^>]*>\s*NUMPAGES\s*<\/w:instrText>/;

async function main(): Promise<void> {
  console.log("pageNumber: true (backward-compat shorthand):");
  {
    const xml = await footerXml({ pageNumber: true });
    check("produces a footer", xml.length > 0);
    check("contains PAGE field", PAGE_FIELD.test(xml));
    check("no NUMPAGES field", !NUMPAGES_FIELD.test(xml));
  }

  console.log("\npageNumber: plain template string:");
  {
    const xml = await footerXml({ pageNumber: "{page} / {pages}" });
    check("PAGE field present", PAGE_FIELD.test(xml));
    check("NUMPAGES field present", NUMPAGES_FIELD.test(xml));
    check("literal ' / ' text run present", xml.includes("/ ") || xml.includes(" /"));
  }

  console.log("\npageNumber: localised template string:");
  {
    const xml = await footerXml({ pageNumber: "Seite {page} von {pages}" });
    check("PAGE field present", PAGE_FIELD.test(xml));
    check("NUMPAGES field present", NUMPAGES_FIELD.test(xml));
    check("literal 'von' text present", xml.includes("von"));
  }

  console.log("\npageNumber: HTML template with formatting:");
  {
    const xml = await footerXml({
      pageNumber: '<p style="text-align:right;font-weight:bold">{page} of {pages}</p>',
    });
    check("PAGE field present", PAGE_FIELD.test(xml));
    check("NUMPAGES field present", NUMPAGES_FIELD.test(xml));
    check("bold run property present", xml.includes("<w:b/>") || xml.includes("<w:b />"));
  }

  console.log("\n{page} token inside footerHtml:");
  {
    const xml = await footerXml({
      footerHtml: '<p style="text-align:center"><i>Page {page} of {pages}</i></p>',
    });
    check("PAGE field present", PAGE_FIELD.test(xml));
    check("NUMPAGES field present", NUMPAGES_FIELD.test(xml));
    check("italic run property present", xml.includes("<w:i/>") || xml.includes("<w:i />"));
  }

  console.log("\n{page} token inside headerHtml:");
  {
    const buf = await convertHtmlToDocx("<p>body</p>", {
      onWarning: null,
      headerHtml: '<p style="text-align:right">{page}</p>',
    });
    const files = unzipSync(new Uint8Array(buf));
    const headerXml = files["word/header1.xml"]
      ? new TextDecoder().decode(files["word/header1.xml"])
      : "";
    check("header file present", headerXml.length > 0);
    check("PAGE field in header", PAGE_FIELD.test(headerXml));
  }

  const ok = failures === 0;
  await writeGuardResult({
    id: "page-number",
    label: "Page-number fields",
    passed: ok ? checksRun : checksRun - failures,
    total: checksRun,
    ok,
    unit: "PAGE/NUMPAGES fields",
    command: "npm run guard:page-number",
  });
  console.log(ok ? `\nPage-number guard passed (${checksRun} checks).` : `\n${failures} check(s) failed.`);
  if (!ok) process.exitCode = 1;
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
