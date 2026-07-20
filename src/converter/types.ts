import type { Paragraph, Table } from "docx";
import { BODY_FONT_HALF_POINTS } from "./constants.js";
import type { StyleResolver } from "./style-resolver.js";
import { INLINE_STYLE_RESOLVER } from "./style-resolver.js";

export type DocxBlock = Paragraph | Table;

export interface BlockBorderSide {
  color: string;
  size: number;
  space: number;
}

export interface BlockBorders {
  top?: BlockBorderSide;
  bottom?: BlockBorderSide;
  left?: BlockBorderSide;
  right?: BlockBorderSide;
}

/** Paragraph-level layout: shading, borders, spacing, indentation. */
export interface BlockLayout {
  alignment?: (typeof import("docx").AlignmentType)[keyof typeof import("docx").AlignmentType];
  shading?: { fill?: string; color?: string };
  spacingBefore?: number;
  spacingAfter?: number;
  indentLeft?: number;
  indentRight?: number;
  /** Hanging indent (twips) — first line outdents left of indentLeft (list markers). */
  hangingIndent?: number;
  paddingTop?: number;
  paddingBottom?: number;
  paddingLeft?: number;
  paddingRight?: number;
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  borders?: BlockBorders;
  /** Line box height (px) for EXACT shaded paragraphs; defaults to body line height. */
  shadedContentLinePx?: number;
  /** Font size (half-points) for shaded padding tab runs. */
  shadedTabHalfPoints?: number;
  /** Force a page break before this paragraph (CSS break-before / class). */
  pageBreakBefore?: boolean;
  /** Request a page break before the next block sibling (CSS break-after / class). */
  pageBreakAfter?: boolean;
}

export interface FlexLayout {
  direction: "row" | "column";
  gap?: number;
}
export interface RunTypography {
  bold?: boolean;
  italics?: boolean;
  underline?: boolean;
  /** Display-only uppercase (`w:caps`) — CSS `text-transform: uppercase`. */
  allCaps?: boolean;
  /** Extra tracking between characters in twips — CSS `letter-spacing`. */
  characterSpacing?: number;
  /** CSS `vertical-align: super` → `w:vertAlign superscript`. */
  superScript?: boolean;
  /** CSS `vertical-align: sub` → `w:vertAlign subscript`. */
  subScript?: boolean;
  color?: string;
  fontSize?: number;
  font?: string;
  style?: string;
  shading?: { fill?: string; color?: string };
}
export type BlockStyle = BlockLayout;

/** @deprecated alias — use RunTypography */
export type RunStyle = RunTypography;

export interface VisitorContext {
  blockquoteDepth: number;
  /** Numbering reference (see LIST_STYLE_REFERENCES); ordered refs start with "numbers". */
  listRef?: string;
  listLevel: number;
  inheritedLayout?: BlockLayout;
  styleResolver: StyleResolver;
  /** Default run size (half-points) for text with no explicit CSS font-size. */
  defaultSizeHalfPoints: number;
  /**
   * Inside a flex-item's block children: use EXACT per-font line boxes (CSS line-height)
   * instead of AUTO, so stacked card lines don't inflate vertically vs the browser.
   */
  flexBlockContent?: boolean;
  /** When true, the next visited block starts on a new page (prior sibling had break-after). */
  pageBreakBeforeNext?: boolean;
}

export const DEFAULT_VISITOR_CONTEXT: VisitorContext = {
  blockquoteDepth: 0,
  listLevel: 0,
  styleResolver: INLINE_STYLE_RESOLVER,
  defaultSizeHalfPoints: BODY_FONT_HALF_POINTS,
};
