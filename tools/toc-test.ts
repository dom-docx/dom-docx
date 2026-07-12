/**
 * Table-of-contents guard — structural OOXML checks for the `tableOfContents`
 * config option.
 *
 * The TOC is a clickable, page-number-less list: each entry is a hyperlink to a
 * bookmark on its heading, with no dot leader and no page number. Because there are
 * no page numbers (which would need layout), the entry list is complete the moment
 * we write it — so the field is NOT dirty, there is no `w:updateFields` (no "update
 * fields" prompt), and it renders correctly in every viewer with no update. The
 * `\n` switch keeps it number-less even if a user does refresh the field, and the
 * heading styles carry explicit `w:outlineLvl` so that refresh rebuilds correctly
 * in LibreOffice (which, unlike Word, collects strictly by outline level).
 *
 * Entries and heading bookmarks are produced by one pass over `document.xml`
 * (`patchTableOfContents`), so they cannot drift out of sync.
 */
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { unzipSync } from "fflate";
import { validateFile } from "@xarsh/ooxml-validator";
import { convertHtmlToDocx, type ConvertOptions } from "../src/converter.js";
import { writeGuardResult } from "./guard-result.js";
import { GUARDS_OUTPUT } from "./output-paths.js";

const OUT_DIR = path.join(GUARDS_OUTPUT, "toc");

const HTML = `
  <h1>Introduction</h1>
  <p>Opening paragraph.</p>
  <h2>Background</h2>
  <p>Some background.</p>
  <h3>Details</h3>
  <p>Fine print.</p>
  <h4>Appendix</h4>
  <p>Aside.</p>
  <h1>Conclusion</h1>
  <p>Closing paragraph.</p>
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

interface Parts {
  doc: string;
  settings: string;
  styles: string;
  buf: Buffer;
}

async function build(options: ConvertOptions | undefined): Promise<Parts> {
  const buf = await convertHtmlToDocx(HTML, options);
  const files = unzipSync(new Uint8Array(buf));
  const dec = (p: string): string =>
    files[p] ? new TextDecoder().decode(files[p]) : "";
  return {
    doc: dec("word/document.xml"),
    settings: dec("word/settings.xml"),
    styles: dec("word/styles.xml"),
    buf,
  };
}

/** The `<w:style>` definition for a given style id, or "" if absent. */
function styleDef(styles: string, id: string): string {
  return styles.match(new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[\\s\\S]*?</w:style>`))?.[0] ?? "";
}

/** The TOC field's own markup (`<w:sdt>…</w:sdt>`) — its entries live here. */
function tocField(doc: string): string {
  const start = doc.indexOf("<w:sdt>");
  const end = doc.indexOf("</w:sdt>");
  return start >= 0 && end > start ? doc.slice(start, end) : "";
}

async function main(): Promise<void> {
  await mkdir(OUT_DIR, { recursive: true });
  console.log("Table-of-contents guard — structural checks:");

  // ---- Omitted: no TOC field, no outline levels, headings intact ----
  console.log("\ntableOfContents omitted (baseline):");
  const off = await build(undefined);
  check("no w:sdt / TOC field", !/<w:sdt>/.test(off.doc) && !/\bTOC\b/.test(off.doc));
  check("no updateFields in settings", !/<w:updateFields/.test(off.settings));
  check("no heading bookmarks", !/<w:bookmarkStart[^>]*w:name="_Toc/.test(off.doc));
  check("headings still map to Heading styles", /w:pStyle w:val="Heading1"/.test(off.doc));
  check("no outline levels added without a TOC", !/<w:outlineLvl/.test(styleDef(off.styles, "Heading1")));

  // ---- tableOfContents: true → defaults (levels 1-3, clickable) ----
  console.log("\ntableOfContents: true (defaults):");
  const on = await build({ tableOfContents: true });
  const onField = tocField(on.doc);
  check("emits a w:sdt block", /<w:sdt>/.test(on.doc));
  check("emits a TOC field instruction", /<w:instrText[^>]*>\s*TOC\b/.test(on.doc));
  check('default heading range \\o "1-3"', /TOC\b[^<]*\\o\s+&quot;1-3&quot;/.test(on.doc));
  check("clickable by default (\\h switch)", /TOC\b[^<]*\\h\b/.test(on.doc));

  // The whole point: number-less + final at creation, so no prompt / no update.
  check("number-less (\\n switch on the field)", /TOC \\n\b/.test(on.doc));
  check("no updateFields setting (no 'update fields' prompt)", !/<w:updateFields/.test(on.settings));
  check("field is NOT dirty (treated as final)", !/w:dirty/.test(onField));
  check("no page-number fields in entries", !/PAGEREF|<w:fldSimple/.test(onField));
  check("no dot leader in entries", !/w:leader="dot"/.test(onField));

  check(
    "TOC precedes first body heading",
    on.doc.indexOf("TOC") < on.doc.indexOf('w:pStyle w:val="Heading1"'),
  );
  check("no title paragraph by default", !on.doc.includes("Table of Contents</w:t>"));
  check("no trailing page break by default", !/<w:br w:type="page"\/>/.test(on.doc));

  // Clickable entries: each is a hyperlink to a heading bookmark, produced from the
  // same scan so anchors and bookmarks line up. Range 1-3 includes h1/h2/h3, not h4.
  check("entry: Introduction (h1) hyperlinked in field", /<w:hyperlink w:anchor="_Toc\d+"[^>]*>[\s\S]*?>Introduction</.test(onField));
  check("entry: Background (h2) in field", />Background</.test(onField));
  check("entry: Details (h3) in field", />Details</.test(onField));
  check("h4 'Appendix' excluded from default 1-3 range", !/>Appendix</.test(onField));
  check(
    "every entry anchor has a matching heading bookmark",
    [...onField.matchAll(/w:anchor="(_Toc\d+)"/g)].every((m) =>
      on.doc.includes(`<w:bookmarkStart w:id="`) && on.doc.includes(`w:name="${m[1]}"`),
    ),
  );

  // ---- Full config: title + custom range + hyperlink:false + page break ----
  console.log("\ntableOfContents: { title, headingRange: '1-2', hyperlink: false, pageBreakAfter }:");
  const full = await build({
    tableOfContents: {
      title: "Contents",
      headingRange: "1-2",
      hyperlink: false,
      pageBreakAfter: true,
    },
  });
  const fullField = tocField(full.doc);
  check("title rendered as bold paragraph", /<w:b\/>[\s\S]*?<w:t[^>]*>Contents<\/w:t>/.test(full.doc));
  check(
    "title is not itself a heading (no self-referential entry)",
    full.doc.indexOf("Contents</w:t>") < full.doc.indexOf("<w:sdt>") &&
      !/Heading\d"[^>]*\/>[\s\S]{0,80}Contents<\/w:t>/.test(full.doc),
  );
  check('custom heading range \\o "1-2"', /TOC\b[^<]*\\o\s+&quot;1-2&quot;/.test(full.doc));
  check("hyperlink: false omits \\h switch", !/TOC\b[^<]*\\h\b/.test(full.doc));
  check("hyperlink: false emits plain-text entries (no hyperlink)", !/<w:hyperlink/.test(fullField));
  check("hyperlink: false emits no heading bookmarks", !/w:name="_Toc/.test(full.doc));
  check("entry: Introduction (h1) in field", />Introduction</.test(fullField));
  check("entry: Background (h2) in field", />Background</.test(fullField));
  check("entry: Conclusion (h1) in field", />Conclusion</.test(fullField));
  check("h3 'Details' excluded from 1-2 range", !/>Details</.test(fullField));
  check("h4 'Appendix' excluded from 1-2 range", !/>Appendix</.test(fullField));
  check("pageBreakAfter emits a page break", /<w:br w:type="page"\/>/.test(full.doc));
  check(
    "page break sits after the TOC field, before the body",
    full.doc.indexOf('<w:br w:type="page"/>') > full.doc.indexOf("</w:sdt>") &&
      full.doc.indexOf('<w:br w:type="page"/>') < full.doc.indexOf('w:pStyle w:val="Heading1"'),
  );

  // Outline levels: needed only so a *manual* field refresh rebuilds correctly in
  // LibreOffice (it collects strictly by outline level; Word infers them).
  console.log("\noutline levels (for graceful manual refresh):");
  check('Heading1 style has outlineLvl "0"', /<w:outlineLvl w:val="0"\/>/.test(styleDef(full.styles, "Heading1")));
  check('Heading2 style has outlineLvl "1"', /<w:outlineLvl w:val="1"\/>/.test(styleDef(full.styles, "Heading2")));

  // ---- Cover page > TOC > content ordering ----
  console.log("\ncoverHtml + tableOfContents (Cover > TOC > Content):");
  const cover = await build({
    coverHtml: "<h1>Report Cover Title</h1><p>Subtitle</p>",
    tableOfContents: true,
    headerHtml: "<p>CONFIDENTIAL</p>",
    pageNumber: true,
  });
  const cDoc = cover.doc;
  const iCoverText = cDoc.indexOf("Report Cover Title");
  const iCoverBreak = cDoc.indexOf('<w:br w:type="page"/>');
  const iTocField = cDoc.indexOf("<w:sdt>");
  const iTocEnd = cDoc.indexOf("</w:sdt>") + "</w:sdt>".length;
  const iBodyHeading = cDoc.indexOf("Introduction", iTocEnd); // real body heading, after the field
  check(
    "order: cover < page break < TOC field < body",
    iCoverText >= 0 && iCoverText < iCoverBreak && iCoverBreak < iTocField && iTocEnd < iBodyHeading,
  );
  check(
    "cover heading is NOT a TOC entry",
    !cDoc.slice(iTocField, iTocEnd).includes("Report Cover Title"),
  );
  check(
    "body headings ARE TOC entries",
    cDoc.slice(iTocField, iTocEnd).includes("Introduction"),
  );
  check("header/footer suppressed on cover page (titlePg)", /<w:titlePg\s*\/>/.test(cDoc));
  check("empty first-page header referenced", /w:type="first"/.test(cDoc));

  // ---- OOXML schema validity ----
  console.log("\nOOXML schema:");
  const docxPath = path.join(OUT_DIR, "output.docx");
  await writeFile(docxPath, on.buf);
  await writeFile(path.join(OUT_DIR, "source.html"), HTML.trim(), "utf-8");
  const validation = await validateFile(docxPath, { officeVersion: "Office2019" });
  check(
    "TOC document is schema-valid",
    validation.ok,
    validation.errors.slice(0, 2).map((e) => e.description).join("; "),
  );

  const coverPath = path.join(OUT_DIR, "cover-toc.docx");
  await writeFile(coverPath, cover.buf);
  const coverValidation = await validateFile(coverPath, { officeVersion: "Office2019" });
  check(
    "cover + TOC document is schema-valid",
    coverValidation.ok,
    coverValidation.errors.slice(0, 2).map((e) => e.description).join("; "),
  );

  const ok = failures === 0;
  await writeGuardResult({
    id: "toc",
    label: "Table of contents",
    passed: checksRun - failures,
    total: checksRun,
    ok,
    unit: "OOXML field structure + schema",
    command: "npm run guard:toc",
  });

  console.log(ok ? "\nAll table-of-contents checks passed." : `\n${failures} check(s) failed.`);
  if (!ok) process.exitCode = 1;
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
