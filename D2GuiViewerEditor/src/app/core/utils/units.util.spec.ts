import {
  CSS_PX_PER_CM,
  TWIPS_PER_CM,
  cmToPx,
  pxToCm,
  twipsToCm,
  twipsToPx,
  ptToPx,
  roundTo,
} from './units.util';

/**
 * Item 1 — page geometry. The physical page size must match the DOCX section
 * geometry at 100% and be independent of zoom. These tests pin the single
 * conversion constant so the ruler, page and pagination cannot drift apart.
 */
describe('units.util (item 1 — page geometry)', () => {
  it('uses one 96-DPI constant (37.8 px/cm) for the whole editor', () => {
    expect(CSS_PX_PER_CM).toBe(37.8);
    expect(cmToPx(1)).toBe(37.8);
    expect(pxToCm(37.8)).toBeCloseTo(1, 6);
  });

  it('renders A4 (21 × 29.7 cm) at the expected physical pixel size at 100%', () => {
    expect(cmToPx(21)).toBeCloseTo(793.8, 1);
    expect(cmToPx(29.7)).toBeCloseTo(1122.66, 1);
  });

  it('converts twips (DOCX unit) to cm/px correctly', () => {
    // A4 width = 11906 twips ≈ 21 cm.
    expect(twipsToCm(11906)).toBeCloseTo(21, 2);
    expect(TWIPS_PER_CM).toBeCloseTo(566.929, 3);
    expect(twipsToPx(11906)).toBeCloseTo(cmToPx(21), 1);
  });

  it('converts points to pixels at 96 DPI', () => {
    expect(ptToPx(72)).toBe(96); // 72 pt = 1 inch = 96 px
    expect(ptToPx(12)).toBeCloseTo(16, 6);
  });

  it('rounds to a fixed number of decimals', () => {
    expect(roundTo(1.23456, 2)).toBe(1.23);
    expect(roundTo(37.795, 1)).toBe(37.8);
  });
});
