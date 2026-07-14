import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Regresja: „duży tytuł na początku dokumentu ma obciętą górną część znaków"
 * (wygląda jak nachodzenie nagłówka strony na tytuł).
 *
 * Root cause (potwierdzony w realnym Chrome ≥133): reguła
 *   `.editor-content p { text-box-trim: trim-start; text-box-edge: text alphabetic; }`
 * przesuwa TUSZ pierwszej linii akapitu w górę o połowę interlinii (half-leading),
 * poza box akapitu. Dla PIERWSZEGO akapitu strony (np. duży „Title" tuż pod nagłówkiem)
 * tusz trafiał ~37px (przy 40pt) POWYŻEJ górnej krawędzi `.editor-content`, którą
 * `overflow:hidden` (paginacja, ADR-0038) przycina → górna część znaków znikała.
 * Ten sam mechanizm nasuwał duży tytuł w ŚRODKU strony na akapit powyżej (nakładanie).
 *
 * `text-box-trim` nie jest layoutowane przez jsdom, więc WIZUALNY test przycięcia
 * pinuje się w headless Chrome (patrz skill `verify`). Tu strzeżemy skompilowanego
 * CSS komponentu: reguła przycinająca nie może wrócić na akapity edytora. Standardowe
 * rozkładanie interlinii (bez trim) renderuje tytuł zgodnie z Wordem i nie przycina glifów.
 */
describe('WysiwygEditorComponent — skompilowany CSS nie stosuje text-box-trim (regresja obciętego tytułu)', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({ imports: [WysiwygEditorComponent] }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    fixture.detectChanges();
  });

  /** Cały surowy CSS komponentu wstrzyknięty do dokumentu (encapsulation emulated =
   *  <style> w <head>). Czytamy tekst BEZ parsowania, bo jsdom pomija nieznane
   *  właściwości w `cssRules` — a chcemy wykryć właśnie taką (text-box-trim). */
  function collectStyleText(): string {
    return Array.from(document.querySelectorAll('style'))
      .map((s) => s.textContent ?? '')
      .join('\n');
  }

  it('kompiluje style komponentu do dokumentu (asercja ma na czym pracować)', () => {
    const css = collectStyleText();
    // Kotwica: charakterystyczna reguła edytora musi być obecna, inaczej brak
    // `text-box-trim` byłby fałszywie zielony (styl w ogóle się nie wstrzyknął).
    expect(css).toMatch(/\.editor-content/);
  });

  it('nie zawiera aktywnej deklaracji text-box-trim ani text-box-edge', () => {
    const css = collectStyleText();
    expect(css).not.toMatch(/text-box-trim\s*:/);
    expect(css).not.toMatch(/text-box-edge\s*:/);
  });
});
