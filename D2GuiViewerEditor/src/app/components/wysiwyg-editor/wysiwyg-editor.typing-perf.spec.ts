import { vi } from 'vitest';
import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Regresja „laga pod klawiszami" (zgłoszenie: pisanie tnie się, kursor na ułamek
 * sekundy skacze na początek tekstu, litery giną — „Ala" → „Aa").
 *
 * Dwa kontrakty:
 *
 * 1. SYNCHRONICZNY rebind: gdy repaginacja zmienia rozkład stron, podmiana
 *    [innerHTML], `_syncPageEditorDom` i `_restoreGlobalCaret` dzieją się w TYM SAMYM
 *    tasku co `_repaginateNow` (jawny `detectChanges()`), a nie przez `setTimeout(0)`.
 *    Poprzednio między renderem Angulara (selekcja skasowana) a odtworzeniem karetki
 *    istniało okno ~30 ms — keystroke obsłużony w tym oknie wstawiał tekst, po czym
 *    `_syncPageEditorDom` nadpisywał go treścią policzoną BEZ tego znaku (dowód: sonda
 *    CDP na żywym edytorze — znak wstrzyknięty w okno znikał, naturalne pisanie przy
 *    60 ms/znak gubiło litery).
 *
 * 2. CACHE pomiarów: `_measureTableRowsHeight` (podzbiory wierszy mierzone narastająco —
 *    O(wierszy²) klonów per tabela per repaginacja) i `_measureBlockRunHeights` (wszystkie
 *    bloki od zera co przebieg) memoizują wyniki po (styl measurera + HTML). Profil CDP
 *    na ~34-stronicowym dokumencie: repaginacja 280–350 ms co ≤600 ms przy pisaniu,
 *    z czego ~70% w ponownych pomiarach niezmienionych tabel; po cache 34–75 ms.
 *    Inwalidacja: `setContent` (nowy dokument) i `_repaginateAfterResources` (fonty/obrazy
 *    zmieniają metryki bez zmiany HTML).
 */
describe('WysiwygEditorComponent — pisanie bez utraty znaków i bez pełnych re-pomiarów', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  beforeEach(async () => {
    if (typeof (document as any).queryCommandState !== 'function') {
      (document as any).queryCommandState = () => false;
    }
    await TestBed.configureTestingModule({ imports: [WysiwygEditorComponent] }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  afterEach(() => {
    window.getSelection()?.removeAllRanges();
  });

  // ---------- 1. synchroniczny rebind ----------

  it('rebind [innerHTML] + sync DOM dzieją się synchronicznie w _repaginateNow (bez okna setTimeout)', () => {
    fixture.detectChanges();
    // Żywy DOM z fragmentacją split-para: pre-merge repaginacji scali fragmenty,
    // więc policzony rozkład ≠ żywy DOM → ścieżka rebindu.
    component.pageContents.set(['<p>Abc </p><p data-split-para="cont">def</p>']);
    fixture.detectChanges();
    const editor = fixture.nativeElement.querySelector('.editor-content') as HTMLElement;
    expect(editor.innerHTML).toContain('data-split-para');

    (component as any)._repaginateNow();

    // Bez flushowania timerów: DOM strony musi już być zgodny z paginacją.
    expect(component.pageContents()[0]).toContain('>Abc def</p>');
    expect(editor.innerHTML).toContain('>Abc def</p>');
    expect(editor.innerHTML).not.toContain('data-split-para');
  });

  // ---------- 2. cache pomiarów ----------

  function makeMeasurer(): HTMLElement {
    const m = document.createElement('div');
    m.style.cssText = 'position:absolute;width:700px;';
    document.body.appendChild(m);
    return m;
  }

  it('_measureTableRowsHeight: drugi identyczny pomiar trafia w cache (bez klonowania do measurera)', () => {
    const measurer = makeMeasurer();
    const table = document.createElement('table');
    table.innerHTML = '<tbody><tr><td><p>A</p></td></tr><tr><td><p>B</p></td></tr></tbody>';
    const rows = Array.from(table.querySelectorAll('tr')) as HTMLTableRowElement[];
    const appendSpy = vi.spyOn(measurer, 'appendChild');

    const h1 = (component as any)._measureTableRowsHeight(table, rows, measurer);
    const callsAfterFirst = appendSpy.mock.calls.length;
    const h2 = (component as any)._measureTableRowsHeight(table, rows, measurer);

    expect(h2).toBe(h1);
    expect(appendSpy.mock.calls.length).toBe(callsAfterFirst); // hit — measurer nietknięty
    // zmiana treści wiersza = inny klucz → świeży pomiar
    rows[0].querySelector('p')!.textContent = 'Zmienione';
    (component as any)._measureTableRowsHeight(table, rows, measurer);
    expect(appendSpy.mock.calls.length).toBeGreaterThan(callsAfterFirst);
    measurer.remove();
  });

  it('_measureBlockRunHeights: identyczna sekwencja bloków trafia w cache', () => {
    const measurer = makeMeasurer();
    const p1 = document.createElement('p'); p1.textContent = 'Pierwszy akapit';
    const p2 = document.createElement('p'); p2.textContent = 'Drugi akapit';
    const appendSpy = vi.spyOn(measurer, 'appendChild');

    const out1 = (component as any)._measureBlockRunHeights(measurer, [p1, p2]);
    const callsAfterFirst = appendSpy.mock.calls.length;
    // te same treści, ale INNE elementy (jak klony kolejnej repaginacji)
    const out2 = (component as any)._measureBlockRunHeights(
      measurer, [p1.cloneNode(true) as HTMLElement, p2.cloneNode(true) as HTMLElement]);

    expect(out2).toEqual(out1);
    expect(appendSpy.mock.calls.length).toBe(callsAfterFirst); // hit
    measurer.remove();
  });

  it('setContent czyści oba cache (nowy dokument = nowe metryki)', () => {
    (component as any)._tableMeasureCache.set('k', 1);
    (component as any)._blockRunMeasureCache.set('k', [1]);

    component.setContent('<p>Nowy dokument</p>');

    expect((component as any)._tableMeasureCache.size).toBe(0);
    expect((component as any)._blockRunMeasureCache.size).toBe(0);
  });
});
