/**
 * Unit conversions for the editor's page geometry — the single source of truth
 * so the ruler, the page and the pagination all agree (previously the ruler
 * used 37.795 px/cm while the page used 37.8, a ~0.01% drift).
 *
 * The physical page is laid out at the CSS reference of 96 CSS px per inch
 * (exact: 1 cm = 96 / 2.54 = 37.7953 px). The editor's working constant is the
 * rounded **37.8 px/cm** used consistently across the pagination engine; it is
 * centralised here rather than sprinkled as a literal.
 *
 * IMPORTANT: these are PHYSICAL-size conversions at 100%. User zoom is a
 * separate CSS `transform: scale()` applied on top of the page and must NEVER
 * be folded into these values — that keeps the document model and the DOCX
 * export independent of the on-screen zoom (page size stays what the DOCX
 * section defines).
 */

/** Working screen density: CSS px per centimetre at 100% (96 DPI, rounded). */
export const CSS_PX_PER_CM = 37.8;

/** CSS reference pixels per inch (CSS spec). */
export const CSS_PX_PER_INCH = 96;

/** Twips (twentieths of a point) per centimetre: 1440 twip/inch ÷ 2.54. */
export const TWIPS_PER_CM = 566.9291338582677;

/** Twips per point. */
export const TWIPS_PER_PT = 20;

export const cmToPx = (cm: number): number => cm * CSS_PX_PER_CM;
export const pxToCm = (px: number): number => px / CSS_PX_PER_CM;
export const twipsToCm = (twips: number): number => twips / TWIPS_PER_CM;
export const twipsToPx = (twips: number): number => twipsToCm(twips) * CSS_PX_PER_CM;
export const ptToPx = (pt: number): number => (pt / 72) * CSS_PX_PER_INCH;

/** Round to a fixed number of decimals (default 2). */
export const roundTo = (n: number, decimals = 2): number => {
  const f = 10 ** decimals;
  return Math.round(n * f) / f;
};
