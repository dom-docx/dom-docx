import {
  ExternalHyperlink,
  InternalHyperlink,
  TextRun,
  UnderlineType,
  type ParagraphChild,
} from "docx";
import type { AnyNode, Element } from "domhandler";
import { internalAnchorFromHref, wrapWithBookmark } from "./bookmarks.js";
import { BODY_FONT_HALF_POINTS, HYPERLINK_COLOR } from "./constants.js";
import { imageRunFromElement } from "./image.js";
import {
  cssToInlineRunTypography,
  isBlockElement,
  isHiddenElement,
  runShadingForFill,
  typographyFromBlockElement,
} from "./css.js";
import type { StyleResolver } from "./style-resolver.js";
import { INLINE_STYLE_RESOLVER } from "./style-resolver.js";
import type { BlockLayout, RunTypography } from "./types.js";

function mergeTypography(base: RunTypography, overlay: RunTypography): RunTypography {
  return {
    bold: overlay.bold !== undefined ? overlay.bold : base.bold,
    italics: overlay.italics !== undefined ? overlay.italics : base.italics,
    underline: overlay.underline !== undefined ? overlay.underline : base.underline,
    allCaps: overlay.allCaps !== undefined ? overlay.allCaps : base.allCaps,
    characterSpacing: overlay.characterSpacing ?? base.characterSpacing,
    color: overlay.color ?? base.color,
    shading: overlay.shading ?? base.shading,
    style: overlay.style ?? base.style,
    fontSize: overlay.fontSize ?? base.fontSize,
    font: overlay.font ?? base.font,
  };
}

function tagTypographyFlags(tag: string): RunTypography {
  switch (tag) {
    case "strong":
    case "b":
      return { bold: true };
    case "em":
    case "i":
      return { italics: true };
    case "u":
      return { underline: true };
    case "code":
      // Font size inherits from the surrounding text (matches browsers).
      // Consolas over Courier New: full-weight like the browser's monospace.
      return { font: "Consolas" };
    case "a":
      return { underline: true, style: "Hyperlink" };
    default:
      return {};
  }
}

export function typographyToTextRunOptions(
  style: RunTypography,
  defaultSize: number = BODY_FONT_HALF_POINTS,
): Record<string, unknown> {
  const defaultColor = style.style === "Hyperlink" ? HYPERLINK_COLOR : undefined;
  const resolvedColor = style.color ?? defaultColor;
  const options: Record<string, unknown> = {
    size: style.fontSize ?? defaultSize,
  };
  if (resolvedColor) {
    options.color = resolvedColor;
  }
  if (style.bold !== undefined) options.bold = style.bold;
  if (style.italics !== undefined) options.italics = style.italics;
  if (style.allCaps !== undefined) options.allCaps = style.allCaps;
  if (style.characterSpacing) options.characterSpacing = style.characterSpacing;
  if (style.underline === true) {
    options.underline = {
      type: UnderlineType.SINGLE,
      ...(style.style === "Hyperlink" ? { color: style.color ?? HYPERLINK_COLOR } : {}),
    };
  } else if (style.underline === false) {
    options.underline = { type: UnderlineType.NONE };
  }
  if (style.font) options.font = style.font;
  if (style.style) options.style = style.style;
  if (style.shading?.fill) {
    options.shading = runShadingForFill(style.shading.fill);
  }
  return options;
}

const RUN_TEXT = new WeakMap<TextRun, string>();

function detachTrailingSpaces(
  runs: ParagraphChild[],
  style: RunTypography,
  defaultSize: number,
): void {
  if (runs.length === 0) return;
  const last = runs[runs.length - 1];
  if (!(last instanceof TextRun)) return;
  const text = RUN_TEXT.get(last) ?? "";
  const match = text.match(/^(.*?)( +)$/);
  if (!match || !match[1]) return;
  runs[runs.length - 1] = textRun(match[1], style, defaultSize);
  runs.push(textRun(match[2], style, defaultSize));
}

function textRun(text: string, style: RunTypography, defaultSize: number = BODY_FONT_HALF_POINTS): TextRun {
  const run = new TextRun({ text, ...typographyToTextRunOptions(style, defaultSize) });
  RUN_TEXT.set(run, text);
  return run;
}

function pushTextRun(
  runs: ParagraphChild[],
  text: string,
  style: RunTypography,
  defaultSize: number,
  state: InlineRunState,
): void {
  runs.push(textRun(text, style, defaultSize));
  state.lastUnderlineCancelled = style.underline === false;
}

function isElement(node: AnyNode): node is Element {
  return node.type === "tag";
}

interface InlineRunState {
  leadingTextDone: boolean;
  needsSpaceBeforeNext: boolean;
  /** After `<br>` — trim HTML source indentation on the next line. */
  afterLineBreak: boolean;
  /** Previous text run explicitly cancelled underline (e.g. `text-decoration:none`). */
  lastUnderlineCancelled?: boolean;
}

function prependPendingSpace(
  text: string,
  inherited: RunTypography,
  state: InlineRunState,
): string {
  if (!state.needsSpaceBeforeNext || !text || text.startsWith(" ")) {
    state.needsSpaceBeforeNext = false;
    return text;
  }
  state.needsSpaceBeforeNext = false;
  return ` ${text}`;
}

function markPendingSpace(state: InlineRunState, runs: ParagraphChild[]): void {
  if (state.afterLineBreak) return;
  if (state.leadingTextDone && runs.length > 0) {
    state.needsSpaceBeforeNext = true;
  }
}

export function collectInlineRunsFromNodes(
  childNodes: AnyNode[],
  inherited: RunTypography = {},
  state: InlineRunState = { leadingTextDone: false, needsSpaceBeforeNext: false, afterLineBreak: false },
  styleResolver: StyleResolver = INLINE_STYLE_RESOLVER,
  defaultSize: number = BODY_FONT_HALF_POINTS,
): ParagraphChild[] {
  const runs: ParagraphChild[] = [];

  for (const node of childNodes) {
    if (node.type === "text") {
      const raw = node.data ?? "";
      let normalized = raw.replace(/\s+/g, " ");
      if (state.afterLineBreak) {
        // Pretty-printed HTML adds newline+indent after `<br>` — browsers drop that,
        // but a literal `<br> Line two` space after the break should stay.
        if (/^\s*[\r\n]/.test(raw)) {
          normalized = normalized.trimStart();
        }
        if (!normalized) continue;
        state.afterLineBreak = false;
        state.leadingTextDone = true;
      } else if (!state.leadingTextDone) {
        normalized = normalized.trimStart();
        if (normalized) state.leadingTextDone = true;
      }
      if (!normalized.trim()) {
        markPendingSpace(state, runs);
        continue;
      }
      normalized = prependPendingSpace(normalized, inherited, state);
      if (normalized.startsWith(" ") && runs.length > 0) {
        // Resuming `<u>` after `text-decoration:none`: keep the leading space on the
        // following underlined run. A standalone underlined space-only run becomes
        // an empty `<w:t/>` in the packer and LibreOffice PDF export paints its
        // underline backward through the preceding plain run.
        const resumeAfterCancel = inherited.underline === true && state.lastUnderlineCancelled;
        if (!resumeAfterCancel) {
          pushTextRun(runs, " ", inherited, defaultSize, state);
          normalized = normalized.trimStart();
        }
      }
      pushTextRun(runs, normalized, inherited, defaultSize, state);
      continue;
    }

    if (!isElement(node)) continue;
    if (isHiddenElement(node, styleResolver)) continue;

    const tag = node.name.toLowerCase();
    if (tag === "br") {
      state.needsSpaceBeforeNext = false;
      state.afterLineBreak = true;
      runs.push(
        new TextRun({
          break: 1,
          size: inherited.fontSize ?? defaultSize,
        }),
      );
      continue;
    }

    if (tag === "img") {
      const image = imageRunFromElement(node);
      if (image) {
        runs.push(image);
      } else if (node.attribs?.alt) {
        runs.push(textRun(node.attribs.alt, inherited, defaultSize));
      }
      state.needsSpaceBeforeNext = false;
      continue;
    }

    if (state.needsSpaceBeforeNext && runs.length > 0) {
      pushTextRun(runs, " ", inherited, defaultSize, state);
      state.needsSpaceBeforeNext = false;
    }
    if (state.afterLineBreak) {
      state.afterLineBreak = false;
    }

    const cssTypography = cssToInlineRunTypography(styleResolver.getCss(node));
    const typography = mergeTypography(
      mergeTypography(inherited, tagTypographyFlags(tag)),
      cssTypography,
    );

    if (tag === "a") {
      const href = node.attribs?.href ?? "";
      const internalAnchor = internalAnchorFromHref(href);
      // `id` or legacy `name` makes this element a jump target as well.
      const bookmarkId = node.attribs?.id || node.attribs?.name;
      detachTrailingSpaces(runs, inherited, defaultSize);
      const linkTypography = {
        ...typography,
        color: cssTypography.color,
        style: "Hyperlink" as const,
      };
      const childRuns = collectInlineRunsFromNodes(
        node.children ?? [],
        {
          ...linkTypography,
          underline: true,
        },
        state,
        styleResolver,
        defaultSize,
      );
      // Anchors with no inline content render as nothing in browsers (e.g. a
      // link that only wrapped an image) — unless they are a named target.
      if (childRuns.length === 0) {
        runs.push(...wrapWithBookmark(bookmarkId, []));
        continue;
      }
      let linked: ParagraphChild[];
      if (internalAnchor) {
        linked = [new InternalHyperlink({ anchor: internalAnchor, children: childRuns })];
      } else {
        linked = [new ExternalHyperlink({ link: href, children: childRuns })];
      }
      runs.push(...wrapWithBookmark(bookmarkId, linked));
      continue;
    }

    const childRuns = collectInlineRunsFromNodes(
      node.children ?? [],
      typography,
      state,
      styleResolver,
      defaultSize,
    );
    // Any inline element with `id` is a fragment target (e.g. `<span id="n">`).
    runs.push(...wrapWithBookmark(node.attribs?.id, childRuns));
  }

  return runs;
}

export function hasDirectBlockChild(
  element: Element,
  styleResolver: StyleResolver = INLINE_STYLE_RESOLVER,
): boolean {
  return (element.children ?? []).some((node) => isElement(node) && isBlockElement(node, styleResolver));
}

export function getDirectBlockChildren(
  element: Element,
  styleResolver: StyleResolver = INLINE_STYLE_RESOLVER,
): Element[] {
  return (element.children ?? []).filter(
    (node): node is Element => isElement(node) && isBlockElement(node, styleResolver),
  );
}

export function getInlineOrTextNodes(
  element: Element,
  styleResolver: StyleResolver = INLINE_STYLE_RESOLVER,
): AnyNode[] {
  return (element.children ?? []).filter((node) => {
    if (node.type === "text") return (node.data ?? "").trim().length > 0;
    if (!isElement(node)) return false;
    return !isBlockElement(node, styleResolver);
  });
}

export { isBlockElement, typographyFromBlockElement };
