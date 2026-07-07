import { Injectable, computed, signal } from '@angular/core';

/**
 * Where a font comes from. Mirrors the `FontDefinition.source` contract in the
 * task spec so both toolbars and any future embedded-font logic agree.
 */
export type FontSource = 'system' | 'application' | 'document' | 'embedded';

export interface FontDefinition {
  /** Stable identifier (normalised family name, lower-case). */
  id: string;
  /** Human-facing name shown in pickers. */
  displayName: string;
  /** Value used for CSS `font-family` / `w:rFonts`. */
  cssFamily: string;
  /** Alternative spellings that should normalise to this font. */
  aliases: readonly string[];
  source: FontSource;
  available: boolean;
}

/**
 * Single source of truth for the fonts offered across the editor.
 *
 * Both the main toolbar and the contextual (mini) toolbar consume this service
 * instead of keeping their own arrays, so the corporate font — and any other
 * entry — is always available in both places and normalised the same way
 * (item 7). The corporate font is read at runtime from the deployment's
 * `--corporate-font-family` CSS variable (configured in
 * `assets/fonts/_corporate-font.scss` and mirrored by the API's
 * `DocumentDefaults.FontFamily`).
 */
@Injectable({ providedIn: 'root' })
export class FontProviderService {
  /** Curated system fonts (union of the two lists this service replaces). */
  private readonly systemFonts: readonly string[] = [
    'Calibri',
    'Calibri Light',
    'Arial',
    'Arial Narrow',
    'Times New Roman',
    'Cambria',
    'Georgia',
    'Verdana',
    'Tahoma',
    'Trebuchet MS',
    'Helvetica',
    'Comic Sans MS',
    'Courier New',
    'Lucida Console',
    'Palatino Linotype',
    'Garamond',
    'Book Antiqua',
    'Impact',
  ];

  private readonly _fonts = signal<readonly FontDefinition[]>(this.buildFonts());

  /** All fonts, corporate first, then system. */
  readonly fonts = this._fonts.asReadonly();

  /** Display names only — convenient for `@for` in templates. */
  readonly displayNames = computed<readonly string[]>(() =>
    this._fonts().map((f) => f.displayName),
  );

  /** Recompute the list (e.g. after the corporate font CSS variable changes). */
  refresh(): void {
    this._fonts.set(this.buildFonts());
  }

  /**
   * Corporate font family name if one is configured for this deployment, else
   * null. Derived from the first family in `--corporate-font-family`; when it
   * equals the built-in fallback (Calibri…) there is no distinct corporate font.
   */
  corporateFontFamily(): string | null {
    const first = this.firstFamily(this.readCorporateFontRaw());
    if (!first) return null;
    if (this.isGenericFamily(first)) return null;
    // The fallback stack starts with Calibri; treat that as "no corporate font".
    if (first.toLowerCase() === 'calibri') return null;
    return first;
  }

  /**
   * Normalise an incoming font-family string (possibly a CSS stack with quotes)
   * to a canonical display name. Falls back to the raw first family when the
   * font is not in the list, so unknown/document fonts are shown as-is rather
   * than silently replaced. Shared by both toolbars (item 7).
   */
  normalize(raw?: string | null): string {
    const first = this.firstFamily(raw);
    if (!first) return 'Calibri';
    const incoming = first.toLowerCase();

    // Exact match on display name or alias.
    for (const font of this._fonts()) {
      if (font.id === incoming) return font.displayName;
      if (font.aliases.some((a) => a.toLowerCase() === incoming)) return font.displayName;
    }

    // Prefix/contains match, longest name first so "Calibri Light" beats "Calibri".
    const byLength = [...this._fonts()].sort(
      (a, b) => b.displayName.length - a.displayName.length,
    );
    const partial = byLength.find((f) => incoming.includes(f.displayName.toLowerCase()));
    if (partial) return partial.displayName;

    return first;
  }

  /** True when the (normalised) name matches a known font in the list. */
  isKnown(name?: string | null): boolean {
    const first = this.firstFamily(name);
    if (!first) return false;
    const lower = first.toLowerCase();
    return this._fonts().some(
      (f) => f.id === lower || f.aliases.some((a) => a.toLowerCase() === lower),
    );
  }

  private buildFonts(): readonly FontDefinition[] {
    const list: FontDefinition[] = [];

    const corporate = this.corporateFontFamily();
    if (corporate) {
      list.push({
        id: corporate.toLowerCase(),
        displayName: corporate,
        cssFamily: corporate,
        aliases: [],
        source: 'application',
        available: true,
      });
    }

    for (const name of this.systemFonts) {
      // Guard against the corporate font also appearing in the system list.
      if (list.some((f) => f.id === name.toLowerCase())) continue;
      list.push({
        id: name.toLowerCase(),
        displayName: name,
        cssFamily: name,
        aliases: [],
        source: 'system',
        available: true,
      });
    }

    return list;
  }

  /** Overridable seam for tests; reads the corporate font CSS variable. */
  protected readCorporateFontRaw(): string {
    if (typeof document === 'undefined' || !document.documentElement) return '';
    try {
      return getComputedStyle(document.documentElement)
        .getPropertyValue('--corporate-font-family')
        .trim();
    } catch {
      return '';
    }
  }

  private firstFamily(raw?: string | null): string {
    if (!raw) return '';
    return raw
      .split(',')[0]
      .trim()
      .replace(/^['"]|['"]$/g, '')
      .trim();
  }

  private isGenericFamily(name: string): boolean {
    return ['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui'].includes(
      name.toLowerCase(),
    );
  }
}
