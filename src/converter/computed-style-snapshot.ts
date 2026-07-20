export interface ComputedStyleSnapshot {
  path: string;
  styles: Record<string, string>;
}

const ELEMENT_NODE = 1;

function isElement(node: Node): node is Element {
  return node.nodeType === ELEMENT_NODE;
}

/** Stable path from `el` up to (but not including) `root`, or `body` when `root` is omitted. */
function elementStylePathFrom(el: Element, root?: Element | null): string {
  const stopBefore = root ?? el.ownerDocument.body;
  const parts: string[] = [];
  let current: Element | null = el;

  while (current && current !== stopBefore) {
    const tagName = current.tagName.toLowerCase();
    const parent: Element | null = current.parentElement;
    if (!parent) break;

    const siblings = Array.from(parent.children).filter(
      (c: Element) => c.tagName.toLowerCase() === tagName,
    );
    const index = siblings.indexOf(current);
    parts.unshift(`${tagName}[${index}]`);
    current = parent;
  }

  return parts.join("/");
}

/**
 * Batch-read `getComputedStyle` for elements under `root` (or `document.body` when omitted).
 * Paths are relative to `root` so they match cheerio paths for `root.innerHTML` fragments.
 */
export function snapshotComputedStylesFromDocument(
  doc: Document = document,
  root?: Element | null,
): ComputedStyleSnapshot[] {
  const props = [
    "color",
    "backgroundColor",
    "display",
    "flexDirection",
    "gap",
    "columnGap",
    "rowGap",
    "textAlign",
    "fontSize",
    "fontWeight",
    "fontStyle",
    "marginTop",
    "marginRight",
    "marginBottom",
    "marginLeft",
    "paddingTop",
    "paddingRight",
    "paddingBottom",
    "paddingLeft",
    "height",
    "width",
    "maxWidth",
    "borderTopWidth",
    "borderTopColor",
    "borderRightWidth",
    "borderRightColor",
    "borderBottomWidth",
    "borderBottomColor",
    "borderLeftWidth",
    "borderLeftColor",
    "breakBefore",
    "breakAfter",
    "writingMode",
    "textOrientation",
  ] as const;

  const results: ComputedStyleSnapshot[] = [];

  const walk = (el: Element): void => {
    const cs = doc.defaultView?.getComputedStyle(el);
    if (!cs) return;
    const styles: Record<string, string> = {};
    for (const prop of props) {
      styles[prop] = cs[prop as keyof CSSStyleDeclaration] as string;
    }
    results.push({ path: elementStylePathFrom(el, root), styles });
    for (const child of el.children) {
      if (isElement(child)) walk(child);
    }
  };

  if (root) {
    for (const child of root.children) {
      if (isElement(child)) walk(child);
    }
  } else {
    for (const child of doc.body.children) {
      if (isElement(child)) walk(child);
    }
  }
  return results;
}
