import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Regresja: „treść nagłówka/stopki nachodzi na treść dokumentu" — textbox pasma (albo obiekt
 * body z ujemnym offsetem) przegrywał/wygrywał walkę warstw wbrew Wordowi.
 *
 * Model warstw Worda: CAŁA warstwa nagłówka/stopki leży POD warstwą główną. Kluczowy mechanizm:
 * `.editor-content` ma `position:relative` + z-index, więc tworzy stacking context — dzieci
 * (textboxy z inline z-index z readera) są w nim UWIĘZIONE i o kolejności względem pasma
 * decyduje wyłącznie para (pasmo, editor-content). Poprzedni układ (pasmo z=5, content z=0)
 * przykrywał obiekty body niezależnie od ich z-index; podbijanie dzieciom z-index (stary fix
 * z=6) nie mogło zadziałać.
 *
 * jsdom nie liczy layoutu, więc wizualne nachodzenie pinuje się w headless Chrome (skill
 * `verify`). Tu strzeżemy skompilowanego CSS: relacja pasmo < content < pasmo-w-edycji
 * nie może się odwrócić, a content nie może dostać nieprzezroczystego tła (zasłoniłoby
 * treść pasma całkowicie — Word pokazuje ją pod tekstem body jak znak wodny).
 */
describe('WysiwygEditorComponent — warstwy pasm nagłówka/stopki vs treść (model Worda)', () => {
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

  /**
   * z-index z pierwszej reguły pasującej do wzorca. Wymaganie „z-index w tym samym bloku" pomija reguły
   * bez z-index (np. bazowy `.editor-content { width… }`).
   */
  function zIndexOf(pattern: RegExp, css: string): number | null {
    const match = pattern.exec(css);
    return match ? parseInt(match[1], 10) : null;
  }

  /* Komponent ma encapsulation None — selektory w skompilowanym CSS są gołe (bez [_ngcontent]).
     Wzorce celują w konkretne reguły: pasmo (lista .page-header,.page-footer), tryb edycji
     oraz SAMODZIELNĄ regułę .editor-content (w liście selektorów po nazwie stoi przecinek,
     w regułach potomnych — spacja, więc `\s*\{` je odrzuca). */
  const bandPattern = /\.page-header,\s*\.page-footer\s*\{[^}]*?z-index:\s*(-?\d+)/;
  const editingPattern = /\.page-header\.editing,\s*\.page-footer\.editing\s*\{[^}]*?z-index:\s*(-?\d+)/;
  const contentPattern = /\.editor-content\s*\{[^}]*?z-index:\s*(-?\d+)/;

  it('kompiluje style komponentu do dokumentu (asercja ma na czym pracować)', () => {
    expect(collectStyleText()).toMatch(/\.page-header/);
  });

  it('pasmo nagłówka/stopki leży POD warstwą główną dokumentu', () => {
    const css = collectStyleText();
    const band = zIndexOf(bandPattern, css);
    const content = zIndexOf(contentPattern, css);

    expect(band).not.toBeNull();
    expect(content).not.toBeNull();
    expect(band!).toBeLessThan(content!);
  });

  it('pasmo w trybie edycji wychodzi PONAD warstwę główną', () => {
    const css = collectStyleText();
    const editing = zIndexOf(editingPattern, css);
    const content = zIndexOf(contentPattern, css);

    expect(editing).not.toBeNull();
    expect(content).not.toBeNull();
    expect(editing!).toBeGreaterThan(content!);
  });

  it('editor-content pozostaje stacking contextem (position:relative z z-index)', () => {
    // Bez stacking contextu inline z-index textboxów body porównywałby się bezpośrednio
    // z pasmem i wynik zależałby od wartości z readera — a ma zależeć wyłącznie od warstwy.
    const css = collectStyleText();

    expect(css).toMatch(/\.editor-content[^{]*\{[^}]*position:\s*relative/);
  });

  it('warstwa główna nad pasmem ma przezroczyste tło (treść pasma widoczna pod tekstem body)', () => {
    const css = collectStyleText();
    // Reguła podnosząca content (z z-index) musi w TYM SAMYM bloku deklarować transparent —
    // nieprzezroczyste tło zasłoniłoby treść pasma zamiast pokazać ją pod tekstem.
    const raisedContentBlock = /\.editor-content\s*\{[^}]*?z-index:[^}]*?\}/.exec(css)?.[0] ?? '';

    expect(raisedContentBlock).toMatch(/background:\s*transparent/);
  });
});
