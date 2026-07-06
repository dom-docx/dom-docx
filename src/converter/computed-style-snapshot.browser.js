(root) => {
  const elementStylePathFrom = (el, rootEl) => {
    const stopBefore = rootEl || document.body;
    const parts = [];
    let current = el;
    while (current && current !== stopBefore) {
      const tagName = current.tagName.toLowerCase();
      const parent = current.parentElement;
      if (!parent) break;
      const siblings = Array.from(parent.children).filter(
        (c) => c.tagName.toLowerCase() === tagName,
      );
      const index = siblings.indexOf(current);
      parts.unshift(tagName + "[" + index + "]");
      current = parent;
    }
    return parts.join("/");
  };

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
  ];

  const results = [];
  const walk = (el) => {
    const cs = getComputedStyle(el);
    const styles = {};
    for (let i = 0; i < props.length; i++) {
      const prop = props[i];
      styles[prop] = cs[prop];
    }
    results.push({ path: elementStylePathFrom(el, root), styles: styles });
    const children = el.children;
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      if (child instanceof Element) walk(child);
    }
  };

  if (root) {
    const rootChildren = root.children;
    for (let i = 0; i < rootChildren.length; i++) {
      const child = rootChildren[i];
      if (child instanceof Element) walk(child);
    }
  } else {
    const bodyChildren = document.body.children;
    for (let i = 0; i < bodyChildren.length; i++) {
      const child = bodyChildren[i];
      if (child instanceof Element) walk(child);
    }
  }
  return results;
}
