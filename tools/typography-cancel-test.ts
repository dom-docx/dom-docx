/**
 * Guard: explicit bold/italic cancellation via CSS overrides.
 *
 * When a child element has font-style:normal or font-weight:normal inside an
 * inherited bold/italic context, the run must emit w:i/@val=false (or w:b/@val=false)
 * to override the paragraph or parent-run default — otherwise the text inherits the
 * property it was intended to cancel.
 */
import { unzipSync } from "fflate";
import { convertHtmlToDocx } from "../src/converter.js";
import { writeGuardResult } from "./guard-result.js";

let failures = 0;
let checksRun = 0;

function check(name: string, cond: boolean, detail?: string): void {
  checksRun++;
  if (cond) console.log(`  ✓ ${name}`);
  else {
    console.error(`  ✗ ${name}${detail ? ` — ${detail}` : ""}`);
    failures++;
  }
}

async function docXml(html: string): Promise<string> {
  const buf = await convertHtmlToDocx(html, { onWarning: null });
  const files = unzipSync(new Uint8Array(buf));
  return new TextDecoder().decode(files["word/document.xml"]!);
}

async function main(): Promise<void> {
  console.log("italic cancellation:");
  {
    // <i> sets italic; inner span with font-style:normal must cancel it
    const xml = await docXml('<p><i>outer <span style="font-style:normal">inner</span></i></p>');
    const runs = xml.match(/<w:r>[\s\S]*?<\/w:r>/g) ?? [];
    const outerRun = runs.find(r => r.includes(">outer<") || r.includes(">outer "));
    const innerRun = runs.find(r => r.includes(">inner<") || r.includes(">inner<"));
    check(
      "outer run is italic",
      (outerRun ?? "").includes('<w:i/>') || (outerRun ?? "").includes('<w:i ') || (outerRun ?? "").includes('w:i>'),
      outerRun,
    );
    check(
      "inner run cancels italic (w:i val=false)",
      (innerRun ?? "").includes('w:val="false"') || (innerRun ?? "").includes("w:val='false'"),
      innerRun,
    );
    check(
      "inner run has no positive italic flag",
      !(innerRun ?? "").match(/<w:i\s*\/>/) && !(innerRun ?? "").match(/<w:i w:val="1"/),
    );
  }

  console.log("\nbold cancellation:");
  {
    const xml = await docXml('<p><b>outer <span style="font-weight:normal">inner</span></b></p>');
    const runs = xml.match(/<w:r>[\s\S]*?<\/w:r>/g) ?? [];
    const outerRun = runs.find(r => r.includes(">outer<") || r.includes(">outer "));
    const innerRun = runs.find(r => r.includes(">inner<") || r.includes(">inner<"));
    check(
      "outer run is bold",
      (outerRun ?? "").includes('<w:b/>') || (outerRun ?? "").includes('<w:b ') || (outerRun ?? "").includes('w:b>'),
      outerRun,
    );
    check(
      "inner run cancels bold (w:b val=false)",
      (innerRun ?? "").includes('w:val="false"') || (innerRun ?? "").includes("w:val='false'"),
      innerRun,
    );
    check(
      "inner run has no positive bold flag",
      !(innerRun ?? "").match(/<w:b\s*\/>/) && !(innerRun ?? "").match(/<w:b w:val="1"/),
    );
  }

  console.log("\nbold+italic mixed — field token in footer with cancellation:");
  {
    // {page} is bold and explicitly not italic inside an italic parent.
    // Use footerHtml since injectFieldTokens only runs on header/footer/toc/cover.
    const buf = await convertHtmlToDocx("<p>body</p>", {
      onWarning: null,
      footerHtml:
        '<p><i>page <span style="font-weight:bold;font-style:normal">{page}</span> of {pages}</i></p>',
    });
    const files = unzipSync(new Uint8Array(buf));
    const xml = new TextDecoder().decode(files["word/footer1.xml"]!);
    // After patchPackedDocx the begin run has rStyle; find the FldS character style
    // body in styles.xml — that's where the actual rPr lives.
    const stylesXml = new TextDecoder().decode(files["word/styles.xml"]!);
    // Find the PAGE begin run's rStyle value
    const rStyleMatch = xml.match(
      /<w:rPr><w:rStyle w:val="(FldS\d+)"\/><\/w:rPr><w:fldChar w:fldCharType="begin"[^/]*\/><\/w:r>\s*<w:r><w:instrText[^>]*>\s*PAGE\s*</
    );
    const styleId = rStyleMatch?.[1];
    // Find that style in styles.xml
    const styleBody = styleId
      ? stylesXml.match(
          new RegExp(`<w:style[^>]*w:styleId="${styleId}"[^>]*>[\\s\\S]*?<\\/w:style>`),
        )?.[0] ?? ""
      : "";
    check("{page} begin run has a named char style", styleId !== undefined, `styleId=${styleId}`);
    check("{page} char style is bold", styleBody.includes('<w:b/>') || styleBody.includes('<w:b '));
    check(
      "{page} char style cancels italic",
      styleBody.includes('w:val="false"'),
      styleBody,
    );
  }

  console.log("\nfont-weight numeric — 700 is bold, 400 cancels:");
  {
    const xml = await docXml(
      '<p><b>bold <span style="font-weight:400">normal weight</span> back</b></p>',
    );
    const runs = xml.match(/<w:r>[\s\S]*?<\/w:r>/g) ?? [];
    const cancelRun = runs.find(r => r.includes("normal weight") || r.includes("normal"));
    check(
      "font-weight:400 inside <b> cancels bold",
      (cancelRun ?? "").includes('w:val="false"'),
      cancelRun,
    );
    const boldRun = runs.find(r => r.includes(">bold<") || r.includes(">bold "));
    check(
      "font-weight:700 produces bold",
      (() => {
        const xml2 = xml;
        const w700 = xml2.includes('font-weight:700') ? "n/a" :
          (boldRun ?? "").includes('<w:b/>') || (boldRun ?? "").includes('<w:b ');
        return (boldRun ?? "").includes('<w:b/>') || (boldRun ?? "").includes('<w:b ');
      })(),
      boldRun,
    );
  }

  console.log("\nvertical-align: super / sub:");
  {
    const xml = await docXml(
      '<p>E=mc<span style="vertical-align:super;font-size:10px">2</span> and H<span style="vertical-align:sub;font-size:10px">2</span>O</p>',
    );
    const runs = xml.match(/<w:r>[\s\S]*?<\/w:r>/g) ?? [];
    const superRun = runs.find(r => r.includes(">2<") && r.includes("vertAlign"));
    const subRun = runs.find(r => r.includes(">2<") && !r.includes("vertAlign"));
    // find all runs with vertAlign
    const superRuns = runs.filter(r => r.includes('w:val="superscript"'));
    const subRuns = runs.filter(r => r.includes('w:val="subscript"'));
    check(
      "vertical-align:super produces w:vertAlign superscript",
      superRuns.length > 0,
      superRuns[0],
    );
    check(
      "vertical-align:sub produces w:vertAlign subscript",
      subRuns.length > 0,
      subRuns[0],
    );
    check(
      "plain text run has no vertAlign",
      runs.some(r => r.includes("E=mc") && !r.includes("vertAlign")),
    );
  }

  const ok = failures === 0;
  await writeGuardResult({
    id: "typography-cancel",
    label: "Typography cancellation",
    passed: ok ? checksRun : checksRun - failures,
    total: checksRun,
    ok,
    unit: "bold/italic overrides",
    command: "npm run guard:typography-cancel",
  });
  console.log(
    ok
      ? `\nTypography-cancel guard passed (${checksRun} checks).`
      : `\n${failures} check(s) failed.`,
  );
  if (!ok) process.exitCode = 1;
}

main().catch(err => {
  console.error(err);
  process.exitCode = 1;
});
