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
  return documentXml.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/g, (full, inner: string) => {
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
  return patchLeadingBlankSectionPrefix(patchTableCellSpacingOrder(patchShadedParagraphVerticalAlign(documentXml)));
}
