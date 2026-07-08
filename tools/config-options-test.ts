import { unzipSync } from "fflate";
import { convertHtmlToDocx, type ConvertOptions } from "../src/converter.js";
import { writeGuardResult } from "./guard-result.js";

const HTML = `<p>Config option test document.</p>`;

let failures = 0;
let checksRun = 0;
function check(name: string, cond: boolean, detail?: string): void {
  checksRun += 1;
  if (cond) {
    console.log(`  ✓ ${name}`);
  } else {
    console.error(`  ✗ ${name}${detail ? ` — ${detail}` : ""}`);
    failures += 1;
  }
}

async function part(options: ConvertOptions | undefined, path: string): Promise<string> {
  const buf = await convertHtmlToDocx(HTML, options);
  const files = unzipSync(new Uint8Array(buf));
  const data = files[path];
  return data ? new TextDecoder().decode(data) : "";
}

/** The document's default run props (font/size) live in <w:docDefaults>. */
async function docDefaults(options: ConvertOptions | undefined): Promise<string> {
  const styles = await part(options, "word/styles.xml");
  return styles.match(/<w:docDefaults>.*?<\/w:docDefaults>/s)?.[0] ?? "";
}

/** The size on the body text run for the single `<p>` (runs hardcode w:sz). */
async function bodyRunSize(options: ConvertOptions | undefined): Promise<string | null> {
  const doc = await part(options, "word/document.xml");
  return doc.match(/<w:sz w:val="(\d+)"\s*\/>/)?.[1] ?? null;
}

async function main(): Promise<void> {
  // ---- Defaults (no options) must stay Letter / 1" / Arial 14pt ----
  console.log("Defaults (unchanged when options omitted):");
  const defDoc = await part(undefined, "word/document.xml");
  check("page size Letter (w=12240 h=15840)", /w:w="12240"/.test(defDoc) && /w:h="15840"/.test(defDoc));
  check("portrait orientation (docx always emits orient)", /w:orient="portrait"/.test(defDoc));
  check("margins 1440 twips", /w:top="1440"/.test(defDoc) && /w:left="1440"/.test(defDoc));
  const defDD = await docDefaults(undefined);
  check("default font Arial (docDefaults)", /w:ascii="Arial"/.test(defDD));
  check("default size 21 half-pt (14px) in docDefaults", /<w:sz w:val="21"\s*\/>/.test(defDD));
  check("default body run size 21", (await bodyRunSize(undefined)) === "21");

  // ---- pageSize ----
  console.log("\npageSize:");
  const a4 = await part({ pageSize: "a4" }, "word/document.xml");
  check("a4 (w=11906 h=16838)", /w:w="11906"/.test(a4) && /w:h="16838"/.test(a4));
  const custom = await part({ pageSize: { width: 5, height: 7 } }, "word/document.xml");
  check("custom 5x7in (w=7200 h=10080)", /w:w="7200"/.test(custom) && /w:h="10080"/.test(custom));

  // ---- orientation ----
  console.log("\norientation:");
  const land = await part({ orientation: "landscape" }, "word/document.xml");
  check("landscape swaps dims (w=15840 h=12240)", /w:w="15840"/.test(land) && /w:h="12240"/.test(land));
  check("landscape sets orient attr", /w:orient="landscape"/.test(land));
  const landA4 = await part({ pageSize: "a4", orientation: "landscape" }, "word/document.xml");
  check("a4 landscape (w=16838 h=11906)", /w:w="16838"/.test(landA4) && /w:h="11906"/.test(landA4));

  // ---- margins (inches) ----
  console.log("\nmargins:");
  const marg = await part({ margins: { top: 0.5, bottom: 2 } }, "word/document.xml");
  check("top 0.5in = 720 twips", /w:top="720"/.test(marg));
  check("bottom 2in = 2880 twips", /w:bottom="2880"/.test(marg));
  check("unspecified sides default 1440", /w:left="1440"/.test(marg) && /w:right="1440"/.test(marg));

  // ---- defaultFont ----
  console.log("\ndefaultFont:");
  const fontDD = await docDefaults({ defaultFont: { family: "Georgia", sizePt: 12 } });
  check("family Georgia (docDefaults)", /w:ascii="Georgia"/.test(fontDD));
  check("size 24 half-pt (12pt) in docDefaults", /<w:sz w:val="24"\s*\/>/.test(fontDD));
  check(
    "body run size follows sizePt (12pt → 24)",
    (await bodyRunSize({ defaultFont: { sizePt: 12 } })) === "24",
  );

  // ---- metadata ----
  console.log("\nmetadata (docProps/core.xml):");
  const core = await part(
    {
      metadata: {
        title: "Q3 Report",
        subject: "Financials",
        creator: "Blair",
        keywords: ["revenue", "growth"],
        description: "Quarterly summary",
      },
    },
    "docProps/core.xml",
  );
  check("title", /<dc:title>Q3 Report<\/dc:title>/.test(core), core.slice(0, 200));
  check("subject", /<dc:subject>Financials<\/dc:subject>/.test(core));
  check("creator", /<dc:creator>Blair<\/dc:creator>/.test(core));
  check("keywords joined", /<cp:keywords>revenue, growth<\/cp:keywords>/.test(core));
  check("description", /<dc:description>Quarterly summary<\/dc:description>/.test(core));

  // ---- Tier 2: headers / footers / page number ----
  console.log("\nheaders / footers / pageNumber:");
  const buf = await convertHtmlToDocx(HTML, {
    headerHtml: "<p>ACME Confidential</p>",
    footerHtml: "<p>© 2026 ACME</p>",
    pageNumber: true,
  });
  const files = unzipSync(new Uint8Array(buf));
  const dec = (p: string): string => (files[p] ? new TextDecoder().decode(files[p]) : "");
  check("header1.xml exists", Boolean(files["word/header1.xml"]));
  check("header content", dec("word/header1.xml").includes("ACME Confidential"));
  check("footer1.xml exists", Boolean(files["word/footer1.xml"]));
  check("footer content", dec("word/footer1.xml").includes("2026"));
  check("footer PAGE field", /PAGE/.test(dec("word/footer1.xml")));
  const doc = dec("word/document.xml");
  check("section references header", /<w:headerReference/.test(doc));
  check("section references footer", /<w:footerReference/.test(doc));

  // pageNumber without footerHtml still creates a footer with the field
  const pnOnly = unzipSync(new Uint8Array(await convertHtmlToDocx(HTML, { pageNumber: true })));
  check("pageNumber alone creates footer", Boolean(pnOnly["word/footer1.xml"]));

  // ---- Tier 2: lang / direction ----
  console.log("\nlang / direction:");
  const langDD = await docDefaults({ lang: "ar-SA", direction: "rtl" });
  check("lang ar-SA in docDefaults", /w:lang w:val="ar-SA"/.test(langDD));
  check("rtl flag in docDefaults", /<w:rtl\s*\/>/.test(langDD));
  check("no rtl when ltr (default)", !/<w:rtl\s*\/>/.test(await docDefaults(undefined)));

  console.log(failures === 0 ? "\nAll config-option tests passed." : `\n${failures} check(s) failed.`);
  await writeGuardResult({
    id: "config-options",
    label: "Config options",
    passed: checksRun - failures,
    total: checksRun,
    ok: failures === 0,
    unit: "checks passed",
    command: "npm run guard:config",
  });
  if (failures > 0) process.exitCode = 1;
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
