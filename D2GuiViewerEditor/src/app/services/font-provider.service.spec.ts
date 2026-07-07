import { FontProviderService } from './font-provider.service';

/**
 * Item 7 — a single shared font source. The corporate font read from the
 * deployment CSS variable must appear in the list, and normalisation must be
 * consistent (so both toolbars show the same name for the same run).
 */

/** Test double that lets us inject a deterministic corporate font family. */
class StubFontProvider extends FontProviderService {
  raw = '';
  protected override readCorporateFontRaw(): string {
    return this.raw;
  }
}

describe('FontProviderService (item 7 — shared font source)', () => {
  it('exposes the curated system fonts with no corporate font by default', () => {
    const p = new StubFontProvider();
    p.raw = '';
    p.refresh();
    const names = p.displayNames();
    expect(names).toContain('Calibri');
    expect(names).toContain('Times New Roman');
    // Fallback stack starting with Calibri is not a distinct corporate font.
    expect(p.corporateFontFamily()).toBeNull();
    expect(names.filter((n) => n === 'Calibri').length).toBe(1);
  });

  it('surfaces a configured corporate font first, exactly once', () => {
    const p = new StubFontProvider();
    p.raw = "'CorporateSans', 'Calibri', 'Segoe UI', Arial, sans-serif";
    p.refresh();

    expect(p.corporateFontFamily()).toBe('CorporateSans');
    const names = p.displayNames();
    expect(names[0]).toBe('CorporateSans');
    expect(names.filter((n) => n === 'CorporateSans').length).toBe(1);

    const def = p.fonts().find((f) => f.displayName === 'CorporateSans');
    expect(def?.source).toBe('application');
    expect(def?.available).toBe(true);
  });

  it('ignores a generic-only corporate variable', () => {
    const p = new StubFontProvider();
    p.raw = 'sans-serif';
    p.refresh();
    expect(p.corporateFontFamily()).toBeNull();
  });

  it('normalises names consistently (shared by both toolbars)', () => {
    const p = new StubFontProvider();
    p.raw = '';
    p.refresh();
    expect(p.normalize('ARIAL')).toBe('Arial');
    expect(p.normalize('"Times New Roman", serif')).toBe('Times New Roman');
    // "Calibri Light" must win over "Calibri" (longest-name-first match).
    expect(p.normalize('Calibri Light')).toBe('Calibri Light');
    // Unknown/document font is shown as-is, never silently replaced.
    expect(p.normalize('Some Unknown Font')).toBe('Some Unknown Font');
    expect(p.normalize(null)).toBe('Calibri');
  });
});
