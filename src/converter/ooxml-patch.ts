/**
 * Post-pack OOXML tweaks for primitives the docx npm API does not expose.
 */

/**
 * LibreOffice ignores tentative numbering levels and needs w:tab/@w:val="num"
 * (not "left") for decimal list markers to appear in PDF export.
 */
export function patchNumberingXml(numberingXml: string): string {
  let xml = numberingXml.replace(/\s*w15:tentative="1"/g, "");

  return xml.replace(
    /(<w:lvl[\s\S]*?<w:pPr>[\s\S]*?<w:tabs>\s*)<w:tab w:val="left"/g,
    '$1<w:tab w:val="num"',
  );
}

/**
 * Vertically center text inside shaded paragraphs that use EXACT line spacing
 * to simulate block padding. LibreOffice PDF export paints spacing.before outside
 * w:shd, so padding is folded into w:spacing/@w:line; w:textAlignment centers
 * glyphs within that shaded band.
 */
export function patchShadedParagraphVerticalAlign(documentXml: string): string {
  return documentXml.replace(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g, (paragraph, body: string) => {
    // Tall AT_LEAST line boxes for raster images must stay top-aligned; centering
    // vertically offsets the drawing and recreates whitespace above/below charts.
    if (body.includes("<w:drawing>")) return paragraph;
    return paragraph.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/g, (full, inner: string) => {
      // Shaded EXACT paragraphs (padding folded into the line) and AT_LEAST
      // paragraphs (CSS line-height) both need glyphs centered in the line box:
      // LO otherwise stacks ALL extra leading above the text, while browsers
      // split it half above / half below.
      const shadedExact = inner.includes("<w:shd") && inner.includes('w:lineRule="exact"');
      const atLeastLine = inner.includes('w:lineRule="atLeast"');
      if (!shadedExact && !atLeastLine) return full;
      if (inner.includes("<w:textAlignment")) return full;
      return `<w:pPr>${inner}<w:textAlignment w:val="center"/></w:pPr>`;
    });
  });
}

/**
 * The docx library appends w:tblCellSpacing after w:tblLayout, but CT_TblPrBase
 * requires it before w:tblInd/w:tblBorders (right after w:tblW/w:jc) — schema
 * validation fails otherwise and Word may drop the spacing.
 */
export function patchTableCellSpacingOrder(documentXml: string): string {
  return documentXml.replace(
    /<w:tblPr>([\s\S]*?)<\/w:tblPr>/g,
    (full, inner: string) => {
      const spacing = inner.match(/<w:tblCellSpacing[^/]*\/>/);
      if (!spacing) return full;
      const rest = inner.replace(spacing[0], "");
      const anchor = rest.match(/<w:jc [^/]*\/>/) ?? rest.match(/<w:tblW [^/]*\/>/);
      if (!anchor) return `<w:tblPr>${spacing[0]}${rest}</w:tblPr>`;
      const at = rest.indexOf(anchor[0]) + anchor[0].length;
      return `<w:tblPr>${rest.slice(0, at)}${spacing[0]}${rest.slice(at)}</w:tblPr>`;
    },
  );
}

/**
 * Replace `w:fldSimple` fields with the proper 5-run complex field structure
 * (begin → instrText → separate → cached display run → end).
 *
 * `w:fldSimple` is simpler to emit but LibreOffice's OOXML importer drops the
 * inner run's `w:rPr` — it creates a bare page-number element with no character
 * style. Word-style 5-run complex fields preserve run properties when combined
 * with named character styles (see patchChromeFieldFiles).
 */
export function patchFldSimple(xml: string): string {
  return xml.replace(/<w:fldSimple w:instr="([^"]+)">([\s\S]*?)<\/w:fldSimple>/g, (_, instr, inner) => {
    const rPrMatch = inner.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
    const rPr = rPrMatch ? `<w:rPr>${rPrMatch[1]}</w:rPr>` : "";
    const tMatch = inner.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/);
    const cachedVal = tMatch ? tMatch[1] : "1";
    const instrTrimmed = instr.trim();
    return (
      `<w:r>${rPr}<w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> ${instrTrimmed} </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r>${rPr}<w:t>${cachedVal}</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
    );
  });
}

/**
 * Patch header/footer parts: fldSimple → complex fields, then promote field run
 * rPr to FldS* named character styles so LibreOffice applies typography.
 */
export function patchChromeFieldFiles(files: Record<string, Uint8Array>): void {
  const dec = new TextDecoder();
  const enc = new TextEncoder();
  const chromeKeys = Object.keys(files).filter((k) => /^word\/(footer|header)\d*\.xml$/.test(k));
  if (chromeKeys.length === 0) return;

  const compactToId = new Map<string, string>();
  const idToRpr = new Map<string, string>();
  let counter = 0;
  const patchedChrome = new Map<string, string>();
  for (const key of chromeKeys) {
    const xml = patchFldSimple(dec.decode(files[key]!));
    patchedChrome.set(key, xml);
    for (const m of xml.matchAll(/<w:rPr>((?:[^<]|<(?!\/w:rPr>))*)<\/w:rPr><w:fldChar w:fldCharType="begin"/g)) {
      const rPr = m[1]!;
      const compact = rPr.replace(/\s+/g, "");
      if (compact && !compactToId.has(compact)) {
        const id = `FldS${counter++}`;
        compactToId.set(compact, id);
        idToRpr.set(id, rPr);
      }
    }
  }

  if (idToRpr.size === 0) {
    for (const [key, xml] of patchedChrome) files[key] = enc.encode(xml);
    return;
  }

  const rStyleRpr = (rPr: string): string => {
    const id = compactToId.get(rPr.replace(/\s+/g, ""));
    return id ? `<w:rPr><w:rStyle w:val="${id}"/></w:rPr>` : `<w:rPr>${rPr}</w:rPr>`;
  };
  for (const [key, xml] of patchedChrome) {
    let patched = xml;
    patched = patched.replace(
      /<w:rPr>((?:[^<]|<(?!\/w:rPr>))*)<\/w:rPr>(<w:fldChar w:fldCharType="begin")/g,
      (_, rPr, fldChar) => `${rStyleRpr(rPr)}${fldChar}`,
    );
    patched = patched.replace(
      /(<w:r>)<w:rPr>((?:[^<]|<(?!\/w:rPr>))*)<\/w:rPr>(<w:t[^>]*>[\s\S]*?<\/w:t><\/w:r>(?=<w:r><w:fldChar w:fldCharType="end"))/g,
      (_, open, rPr, rest) => `${open}${rStyleRpr(rPr)}${rest}`,
    );
    files[key] = enc.encode(patched);
  }

  if (!files["word/styles.xml"]) return;
  const charStyles = [...idToRpr.entries()]
    .map(
      ([id, rPr]) =>
        `<w:style w:type="character" w:customStyle="1" w:styleId="${id}">` +
        `<w:name w:val="${id}"/>` +
        `<w:basedOn w:val="DefaultParagraphFont"/>` +
        `<w:rPr>${rPr}</w:rPr>` +
        `</w:style>`,
    )
    .join("");
  files["word/styles.xml"] = enc.encode(
    dec.decode(files["word/styles.xml"]).replace("</w:styles>", `${charStyles}</w:styles>`),
  );
}

/**
 * Multi-section docs from docx can start with an empty paragraph followed by a
 * paragraph whose only content is <w:sectPr>. Word may render this prefix as an
 * empty first page. Drop only this exact leading prefix inside <w:body>.
 */
export function patchLeadingBlankSectionPrefix(documentXml: string): string {
  return documentXml.replace(
    /(<w:body[^>]*>)\s*<w:p>\s*<w:r>\s*(?:<w:t\b[^>]*\/>|<w:t\b[^>]*>\s*<\/w:t>)\s*<\/w:r>\s*<\/w:p>\s*<w:p>\s*<w:pPr>\s*<w:sectPr>[\s\S]*?<\/w:sectPr>\s*<\/w:pPr>\s*<\/w:p>/,
    "$1",
  );
}

export function patchDocumentXml(documentXml: string): string {
  return patchLeadingBlankSectionPrefix(
    patchTableCellSpacingOrder(patchShadedParagraphVerticalAlign(documentXml)),
  );
}
