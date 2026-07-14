import type { ParsedCss } from "./css.js";

/**
 * DOCXs are a light document canvas (white page, dark text). Computed styles from a
 * browser tab in dark mode often yield near-white `color` with a transparent fill —
 * invisible in Word. These helpers detect that case so we can drop the light
 * foreground and fall back to the document default.
 */

/** Relative luminance of a 6-char RRGGBB hex (sRGB), 0 = black, 1 = white. */
export function relativeLuminance(hex: string): number {
  const raw = hex.replace(/^#/, "").toLowerCase();
  if (!/^[0-9a-f]{6}$/.test(raw)) return 0;
  const linear = [0, 2, 4].map((i) => {
    const c = parseInt(raw.slice(i, i + 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2];
}

/** Foregrounds at or above this luminance read as "light" (white / pale grey / pastel). */
const LIGHT_FG_LUMA = 0.55;
/** Backgrounds below this luminance count as a dark fill that may host light text. */
const DARK_BG_LUMA = 0.35;

/**
 * Drop near-white text colors when there is no dark block background — Word's page
 * is white, so those colors would vanish. Light text on a dark shaded block is kept.
 */
export function remapComputedColorsForDocumentCanvas(css: ParsedCss): ParsedCss {
  if (!css.color) return css;
  if (relativeLuminance(css.color) < LIGHT_FG_LUMA) return css;
  if (css.backgroundColor && relativeLuminance(css.backgroundColor) < DARK_BG_LUMA) {
    return css;
  }
  return { ...css, color: undefined };
}
