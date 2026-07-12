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
 * The docx library's built-in Heading1..9 styles carry run properties but no
 * `<w:outlineLvl>`. Word still collects them into a TOC via its built-in
 * "Heading N" convention, but LibreOffice's "create from outline" TOC evaluates
 * the *explicit* outline level — without it, updating/regenerating the field finds
 * no entries and empties the table. Stamp each `HeadingN` style with
 * `outlineLvl = N − 1` (Heading1 → 0) so the TOC survives a refresh everywhere.
 */
export function patchHeadingOutlineLevels(stylesXml: string): string {
  return stylesXml.replace(
    /<w:style\b[^>]*\bw:styleId="Heading([1-9])"[^>]*>[\s\S]*?<\/w:style>/g,
    (styleXml, level: string) => {
      if (styleXml.includes("<w:outlineLvl")) return styleXml;
      const outline = `<w:outlineLvl w:val="${Number(level) - 1}"/>`;
      // CT_Style order is (…qFormat) pPr, rPr. Fold into an existing pPr (outlineLvl
      // sits last in CT_PPrBase), else insert a fresh pPr right before rPr.
      if (/<w:pPr>[\s\S]*?<\/w:pPr>/.test(styleXml)) {
        return styleXml.replace("</w:pPr>", `${outline}</w:pPr>`);
      }
      return styleXml.replace("<w:rPr>", `<w:pPr>${outline}</w:pPr><w:rPr>`);
    },
  );
}

/** Parse a Word heading range (`"1-3"`, `"2"`) into an inclusive `[min, max]`. */
function parseHeadingRange(range: string): [number, number] {
  const span = range.match(/(\d)\s*-\s*(\d)/);
  if (span) {
    const lo = Number(span[1]);
    const hi = Number(span[2]);
    return lo <= hi ? [lo, hi] : [hi, lo];
  }
  const single = range.match(/\d/);
  const n = single ? Number(single[0]) : 1;
  return [n, n];
}

/** Concatenated visible text of every `<w:t>` run inside one paragraph's XML. */
function paragraphText(paragraphXml: string): string {
  return [...paragraphXml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)]
    .map((m) => m[1])
    .join("");
}

/**
 * Per-level type styling for a page-number-less TOC entry. Hierarchy comes from a
 * muted teal/blue ramp (darker + bold at the top, lighter deeper) plus indent and
 * space, since there are no page numbers or leaders to carry structure.
 */
function tocEntryStyle(level: number): {
  bold: boolean;
  color: string;
  spaceBefore: number;
  spaceAfter: number;
} {
  // NB: Word/LibreOffice collapse adjacent before+after to the MAX (not the sum),
  // so these are the actual inter-entry gaps, not half of them.
  if (level <= 1) return { bold: true, color: "20464D", spaceBefore: 200, spaceAfter: 200 };
  if (level === 2) return { bold: false, color: "2D5C63", spaceBefore: 160, spaceAfter: 160 };
  if (level === 3) return { bold: false, color: "537479", spaceBefore: 160, spaceAfter: 160 };
  return { bold: false, color: "6E868A", spaceBefore: 160, spaceAfter: 160 };
}

/** Left indent (twips) for a TOC entry: a small base gutter + per-level step. */
function tocEntryIndent(level: number): number {
  return 160 + (level - 1) * 320;
}

/**
 * Turn the TOC field into a clickable, page-number-less table of contents that is
 * final at creation — no field update, no "update fields" prompt.
 *
 * A page-numbered TOC has to be a live field: the numbers depend on layout, so the
 * word processor computes them on update (which needs the alarming `updateFields`
 * prompt, and only refreshes correctly on a full "Update entire table"). Dropping
 * page numbers removes that dependency entirely — the entry list is complete the
 * moment we write it. So here we, in a single pass over `document.xml`:
 *
 *  1. Wrap each in-range heading paragraph in a `_TocN` bookmark (the link target).
 *  2. Inject one clean entry per heading into the field's cached result — a
 *     hyperlink to that bookmark, indented by level, no dot leader, no number.
 *  3. Add the `\n` switch so that if a user ever *does* refresh the field, it stays
 *     number-less instead of growing a page-number column.
 *
 * Deriving the entries and the bookmarks from the same scan means they cannot drift
 * out of sync. When `hyperlink` is false, entries are plain text and no bookmarks
 * are emitted.
 */
export function patchTableOfContents(
  documentXml: string,
  headingRange: string,
  hyperlink: boolean,
): string {
  if (!/<w:instrText[^>]*>\s*TOC\b/.test(documentXml)) return documentXml;

  const [min, max] = parseHeadingRange(headingRange);
  const entries: { level: number; text: string; anchor: string }[] = [];
  let seq = 0;

  // 1. Bookmark each in-range, non-empty heading; collect its text + anchor in order.
  const withBookmarks = documentXml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (para) => {
    const styled = para.match(/<w:pStyle w:val="Heading([1-6])"\s*\/>/);
    if (!styled) return para;
    const level = Number(styled[1]);
    if (level < min || level > max) return para;
    const text = paragraphText(para);
    if (!text.trim()) return para;

    seq += 1;
    const anchor = `_Toc${9000000 + seq}`;
    entries.push({ level, text, anchor });
    if (!hyperlink) return para; // no jump target needed for a plain-text TOC

    const id = 9000000 + seq;
    const start = `<w:bookmarkStart w:id="${id}" w:name="${anchor}"/>`;
    const end = `<w:bookmarkEnd w:id="${id}"/>`;
    const opened = /<\/w:pPr>/.test(para)
      ? para.replace("</w:pPr>", `</w:pPr>${start}`)
      : para.replace(/(<w:p\b[^>]*>)/, `$1${start}`);
    return opened.replace(/<\/w:p>$/, `${end}</w:p>`);
  });

  if (!entries.length) return withBookmarks;

  // 2. Build one styled entry paragraph per heading. A page-number-less TOC has no
  // dot leaders to carry structure, so the hierarchy has to come from type: bold
  // near-black top levels, lighter/greyer deeper levels, indent per level, and a
  // little space above each top-level entry to group sections. Entries stay
  // hyperlinked (clickable) but are NOT the default blue-underline — that reads as
  // a wall of links rather than a contents list (matches how Word styles a TOC).
  const entryXml = entries
    .map(({ level, text, anchor }) => {
      const style = tocEntryStyle(level);
      const rPr = `<w:rPr>${style.bold ? "<w:b/><w:bCs/>" : ""}<w:color w:val="${style.color}"/></w:rPr>`;
      const textRun = `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
      const run = hyperlink
        ? `<w:hyperlink w:anchor="${anchor}" w:history="1">${textRun}</w:hyperlink>`
        : textRun;
      const pPr =
        `<w:pPr>` +
        `<w:spacing w:before="${style.spaceBefore}" w:after="${style.spaceAfter}"/>` +
        `<w:ind w:left="${tocEntryIndent(level)}"/>` +
        `</w:pPr>`;
      return `<w:p>${pPr}${run}</w:p>`;
    })
    .join("");

  // 3. Keep updates number-less (\n); clear the field's dirty flag so Word treats
  // the cached result as final (no grey field, no "update this field" nudge); then
  // splice the entries into the result (after "separate", before "end").
  return withBookmarks
    .replace(
      /(<w:instrText[^>]*>)\s*TOC\b([^<]*)(<\/w:instrText>)/,
      (_full, open, rest, close) => `${open}TOC \\n &quot;${min}-${max}&quot;${rest}${close}`,
    )
    .replace(
      /<w:fldChar w:fldCharType="begin"[^>]*\/>(\s*<w:instrText[^>]*>\s*TOC)/,
      '<w:fldChar w:fldCharType="begin"/>$1',
    )
    .replace(/(<w:fldChar w:fldCharType="separate"\s*\/><\/w:r><\/w:p>)/, `$1${entryXml}`);
}

export function patchDocumentXml(documentXml: string): string {
  return patchTableCellSpacingOrder(patchShadedParagraphVerticalAlign(documentXml));
}
