/**
 * Leading blank section prefix guard — post-pack document.xml patching.
 *
 * Multi-section docs from the docx library can open <w:body> with an empty
 * paragraph followed by a paragraph whose only content is <w:sectPr>, which
 * Word renders as a blank first page. patchDocumentXml must strip exactly
 * that prefix and nothing else.
 */
import { patchDocumentXml } from "../src/converter/ooxml-patch.js";
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

const SECT_PR = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
const EMPTY_P = `<w:p><w:r><w:t/></w:r></w:p>`;
const SECT_ONLY_P = `<w:p><w:pPr>${SECT_PR}</w:pPr></w:p>`;
const CONTENT_P = `<w:p><w:r><w:t>Real content</w:t></w:r></w:p>`;

function main(): void {
  console.log("leading blank prefix (stripped):");
  const prefixed = `<w:body>${EMPTY_P}${SECT_ONLY_P}${CONTENT_P}</w:body>`;
  const patched = patchDocumentXml(prefixed);
  check("empty paragraph + sectPr-only paragraph removed", patched === `<w:body>${CONTENT_P}</w:body>`, patched);

  const emptyTagVariant = `<w:body>${EMPTY_P.replace("<w:t/>", '<w:t xml:space="preserve"></w:t>')}${SECT_ONLY_P}${CONTENT_P}</w:body>`;
  check(
    "empty <w:t></w:t> variant also removed",
    patchDocumentXml(emptyTagVariant) === `<w:body>${CONTENT_P}</w:body>`,
  );

  console.log("\nnon-matching bodies (untouched):");
  const contentFirst = `<w:body>${CONTENT_P}${EMPTY_P}${SECT_ONLY_P}</w:body>`;
  check("mid-body section break kept", patchDocumentXml(contentFirst) === contentFirst);

  const noSectPr = `<w:body>${EMPTY_P}${CONTENT_P}</w:body>`;
  check("leading empty paragraph without sectPr kept", patchDocumentXml(noSectPr) === noSectPr);

  const nonEmptyLead = `<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p>${SECT_ONLY_P}${CONTENT_P}</w:body>`;
  check("non-empty leading paragraph kept", patchDocumentXml(nonEmptyLead) === nonEmptyLead);

  const ok = failures === 0;
  console.log(ok ? `\nSection prefix guard passed (${checksRun} checks).` : `\n${failures} check(s) failed.`);
  if (!ok) process.exitCode = 1;

  void writeGuardResult({
    id: "section-prefix",
    label: "Blank section prefix",
    passed: ok ? checksRun : checksRun - failures,
    total: checksRun,
    ok,
    unit: "document.xml prefix strip",
    command: "npm run guard:section-prefix",
  });
}

main();
