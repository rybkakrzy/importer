import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * ADR-0108 r.7: paginacja bierze szerokość pomiaru ze strony 1 w DOM (skala DOM/geometria).
 * Gdy strona w DOM ma jeszcze geometrię POPRZEDNIEGO dokumentu (przebieg przed change
 * detection po setContent), skala fałszowała szerokość measurera (605 px zamiast 687 px →
 * węższe łamanie → +1 strona, fixture table-text-layout PB02). Kontrakt: rozjazd szerokości
 * strony z geometrią bazową = DOM nieaktualny → pomiar wg geometrii + JEDEN przebieg po renderze.
 */
describe('WysiwygEditorComponent — nieaktualna geometria strony w DOM (ADR-0108 r.7)', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;
  const mounted: HTMLElement[] = [];

  beforeEach(async () => {
    await TestBed.configureTestingModule({ imports: [WysiwygEditorComponent] }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  afterEach(() => {
    mounted.splice(0).forEach(el => el.remove());
  });

  /** Strona z jawnym clientWidth (jsdom nie ma layoutu) + editor-content o szerokości treści. */
  function pageWith(pageClientWidth: number, editorClientWidth: number): void {
    const page = document.createElement('div');
    page.className = 'page';
    const editor = document.createElement('div');
    editor.className = 'editor-content';
    editor.style.padding = '0px'; // jsdom: bez inline paddingu getComputedStyle daje '' → innerW NaN
    editor.innerHTML = '<p>a</p><p>b</p>';
    page.appendChild(editor);
    document.body.appendChild(page);
    mounted.push(page);
    Object.defineProperty(page, 'clientWidth', { value: pageClientWidth, configurable: true });
    Object.defineProperty(editor, 'clientWidth', { value: editorClientWidth, configurable: true });
    (component as any).pageEditorRefs = { toArray: () => [{ nativeElement: editor }] };
    (component as any)._measureBlockRunHeights = (_m: HTMLElement, blocks: HTMLElement[]) => blocks.map(() => 20);
    (component as any)._blockMarginTopPx = () => 0;
  }

  function measurerWidths(): number[] {
    const widths: number[] = [];
    const orig = (component as any)._createBlockMeasurer.bind(component);
    (component as any)._createBlockMeasurer = (cs: CSSStyleDeclaration, w: number) => {
      widths.push(w);
      return orig(cs, w);
    };
    return widths;
  }

  const CSS_PX_PER_CM = 37.8;
  const geometryPageWidth = () => component.baseGeometry().widthCm * CSS_PX_PER_CM;
  const geometryContentWidth = () => {
    const g = component.baseGeometry();
    return (g.widthCm - g.margins.left - g.margins.right) * CSS_PX_PER_CM;
  };

  it('strona w DOM o INNEJ szerokości niż geometria → pomiar wg geometrii + jeden przebieg po renderze', () => {
    pageWith(geometryPageWidth() - 100, 500); // stara geometria; editor 500px zaniżałby skalę
    const widths = measurerWidths();
    const soon = vi.fn();
    (component as any)._flushPaginateSoon = soon;

    (component as any)._repaginateNow();

    expect(widths[0]).toBeCloseTo(geometryContentWidth(), 0);
    expect(soon).toHaveBeenCalledTimes(1);

    // Drugi przebieg przy wciąż nieaktualnym DOM nie kolejkuje kolejnych (brak pętli rAF).
    (component as any)._repaginateNow();
    expect(soon).toHaveBeenCalledTimes(1);
  });

  it('strona w DOM zgodna z geometrią → skala z DOM jak dotąd, bez dodatkowego przebiegu', () => {
    pageWith(geometryPageWidth(), 500);
    const widths = measurerWidths();
    const soon = vi.fn();
    (component as any)._flushPaginateSoon = soon;

    (component as any)._repaginateNow();

    expect(widths[0]).toBeCloseTo(500, 0); // innerW (500) / szerokość treści geometrii × szerokość treści
    expect(soon).not.toHaveBeenCalled();
  });
});
