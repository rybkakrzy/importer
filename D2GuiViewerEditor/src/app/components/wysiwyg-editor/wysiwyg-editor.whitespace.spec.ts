import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Regresja: wielokrotne spacje z DOCX kolapsowały wizualnie do jednej (dane w DOM/pliku
 * były poprawne — `xml:space="preserve"` w zapisie — ale edytor renderował inaczej niż
 * Word, bo `.editor-content` dziedziczył `white-space: normal`).
 *
 * Strzeżemy skompilowanego CSS: kontenery treści (body + pasma nagłówka/stopki) muszą
 * mieć `white-space: pre-wrap`, a bezklasowy nośnik tabulatora (`min-width:2em` z literalnym
 * \t) musi wracać do `normal`, żeby pod pre-wrap nie zmieniał szerokości boksu.
 */
describe('WysiwygEditorComponent — skompilowany CSS zachowuje wielokrotne spacje (pre-wrap)', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({ imports: [WysiwygEditorComponent] }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    fixture.detectChanges();
  });

  function collectStyleText(): string {
    return Array.from(document.querySelectorAll('style'))
      .map((s) => s.textContent ?? '')
      .join('\n');
  }

  it('.editor-content ma white-space: pre-wrap', () => {
    const css = collectStyleText();
    expect(css).toMatch(/\.editor-content\s*\{[^}]*white-space:\s*pre-wrap/);
  });

  it('pasma nagłówka/stopki mają white-space: pre-wrap', () => {
    const css = collectStyleText();
    expect(css).toMatch(/\.header-editor-content[\s\S]{0,200}?white-space:\s*pre-wrap/);
    expect(css).toMatch(/\.footer-display[\s\S]{0,200}?white-space:\s*pre-wrap/);
  });

  it('nośnik tabulatora (min-width:2em) wraca do white-space: normal !important', () => {
    // !important: reader ścieżki leadera emituje nośnik z inline white-space:pre,
    // które bez !important wygrywałoby z arkuszem (literalny \t skakał do stopu).
    const css = collectStyleText();
    expect(css).toMatch(
      /span\[style\*=['"]min-width:2em['"]\][^{]*\{[^}]*white-space:\s*normal\s*!important/
    );
  });
});
