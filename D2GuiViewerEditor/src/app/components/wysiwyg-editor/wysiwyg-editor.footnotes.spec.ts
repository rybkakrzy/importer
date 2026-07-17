import { vi } from 'vitest';
import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';
import { Footnote } from '../../models/document.model';

/**
 * Footnote rendering + editing in the REAL editor component, asserted against the REAL DOM:
 *  - references render as semantic <sup class="footnote-ref"> with the visible number, a stable
 *    data-footnote-id and an aria-label,
 *  - footnote contents render in a per-page region at the BOTTOM of the page holding the reference
 *    (like MS Word, not a form panel under the document), each linked to its reference,
 *  - editing footnote content emits an updated model through footnotesChange (save serialization),
 *  - deleting a footnote removes its reference + content and renumbers the rest,
 *  - adding a footnote inserts a reference + a content entry,
 *  - footnotes reserve space at the bottom of the page — a footnote-bearing block that no longer
 *    fits with its footnote is pushed to the next page (region follows the reference).
 */
describe('WysiwygEditorComponent — przypisy dolne', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  const BODY =
    '<div class="document-content">' +
    '<p>Alfa<sup class="footnote-ref" data-footnote-id="fn-1" aria-label="Przypis 1">1</sup> beta.</p>' +
    '<p>Gamma<sup class="footnote-ref" data-footnote-id="fn-2" aria-label="Przypis 2">2</sup>.</p>' +
    '</div>';

  const FOOTNOTES: Footnote[] = [
    { id: 'fn-1', html: '<p>Pierwszy przypis.</p>' },
    { id: 'fn-2', html: '<p>Drugi przypis.</p>' }
  ];

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  /** Rozkład regionu przypisów liczy repaginacja (debounce) — wymuś przebieg synchronicznie. */
  function flushPagination(): void {
    (component as unknown as { _flushPaginateNow(): void })._flushPaginateNow();
    fixture.detectChanges();
  }

  function load(body = BODY, footnotes = FOOTNOTES): HTMLElement {
    component.content = body;
    component.footnotes = footnotes.map(f => ({ ...f }));
    fixture.detectChanges();
    flushPagination();
    // Numeracja (dziesiętna) i przypisanie odwołań do stron ustala się po renderze stron —
    // wymuś synchronicznie, żeby markery i region były deterministyczne (jak w endnotes.spec).
    component.syncFootnotesWithBody();
    fixture.detectChanges();
    return fixture.nativeElement as HTMLElement;
  }

  it('renderuje odwołania jako <sup class="footnote-ref"> z numerem, id i aria-label', () => {
    const host = load();

    const refs = Array.from(host.querySelectorAll('sup.footnote-ref')) as HTMLElement[];
    expect(refs.length).toBe(2);

    expect(refs[0].tagName).toBe('SUP');
    expect(refs[0].getAttribute('data-footnote-id')).toBe('fn-1');
    expect(refs[0].textContent).toBe('1');
    expect(refs[0].getAttribute('aria-label')).toBe('Przypis 1');

    expect(refs[1].getAttribute('data-footnote-id')).toBe('fn-2');
    expect(refs[1].textContent).toBe('2');
    expect(refs[1].getAttribute('aria-label')).toBe('Przypis 2');
  });

  it('renderuje treści przypisów w regionie WEWNĄTRZ strony (nie pod dokumentem), z separatorem', () => {
    const host = load();

    const region = host.querySelector('.footnotes-region');
    expect(region).not.toBeNull();
    expect(region!.closest('.page')).not.toBeNull();
    // Panel-karta pod dokumentem został usunięty (był stylizowany jak formularz).
    expect(host.querySelector('.footnotes-panel')).toBeNull();
    const separator = region!.querySelector('.footnotes-separator');
    expect(separator).not.toBeNull();
    expect(separator!.classList.contains('footnotes-separator-continuation')).toBe(false);

    const items = Array.from(host.querySelectorAll('.footnotes-region .footnote-item')) as HTMLElement[];
    expect(items.length).toBe(2);

    // Kolejność + numeracja dziesiętna (jak odwołanie w treści).
    expect(items[0].getAttribute('data-footnote-id')).toBe('fn-1');
    expect(items[1].getAttribute('data-footnote-id')).toBe('fn-2');
    expect(items[0].querySelector('.footnote-item-number')?.textContent).toBe('1');
    expect(items[1].querySelector('.footnote-item-number')?.textContent).toBe('2');

    // Treść.
    expect(items[0].querySelector('.footnote-item-content')?.textContent).toContain('Pierwszy przypis.');
    expect(items[1].querySelector('.footnote-item-content')?.textContent).toContain('Drugi przypis.');

    // Każde odwołanie prowadzi do istniejącej treści (powiązanie po id).
    const refIds = Array.from(host.querySelectorAll('sup.footnote-ref'))
      .map(el => el.getAttribute('data-footnote-id'));
    const itemIds = items.map(el => el.getAttribute('data-footnote-id'));
    refIds.forEach(id => expect(itemIds).toContain(id));
  });

  it('edycja treści przypisu emituje zmieniony model przez footnotesChange', () => {
    const host = load();
    let emitted: Footnote[] | null = null;
    component.footnotesChange.subscribe(v => (emitted = v));

    const content = host.querySelector(
      '[data-testid="footnote-content-fn-1"]'
    ) as HTMLElement;
    content.innerHTML = '<p>Zmieniona treść.</p>';
    content.dispatchEvent(new Event('blur'));

    expect(emitted).not.toBeNull();
    const model = emitted as unknown as Footnote[];
    expect(model.find(f => f.id === 'fn-1')?.html).toContain('Zmieniona treść.');
    // Druga treść nietknięta — jedno źródło prawdy, brak duplikacji.
    expect(model.find(f => f.id === 'fn-2')?.html).toContain('Drugi przypis.');
  });

  it('usunięcie przypisu kasuje odwołanie i treść oraz przelicza numerację', () => {
    const host = load();
    let emitted: Footnote[] | null = null;
    component.footnotesChange.subscribe(v => (emitted = v));

    component.removeFootnote('fn-1');
    flushPagination();

    // Odwołanie i treść zniknęły.
    expect(host.querySelector('sup.footnote-ref[data-footnote-id="fn-1"]')).toBeNull();
    expect(host.querySelector('.footnote-item[data-footnote-id="fn-1"]')).toBeNull();

    // Pozostały przypis przenumerowany na 1 (odwołanie w treści + region).
    const remainingRef = host.querySelector('sup.footnote-ref[data-footnote-id="fn-2"]') as HTMLElement;
    expect(remainingRef.textContent).toBe('1');
    expect(remainingRef.getAttribute('aria-label')).toBe('Przypis 1');
    const remainingItem = host.querySelector('.footnote-item[data-footnote-id="fn-2"] .footnote-item-number');
    expect(remainingItem?.textContent).toBe('1');

    // Model bez osieroconych.
    const model = emitted as unknown as Footnote[];
    expect(model.map(f => f.id)).toEqual(['fn-2']);
  });

  it('dodanie przypisu w pozycji kursora wstawia odwołanie i wpis treści', () => {
    const host = load();

    // Ustaw aktywny edytor (jak focusin na stronie).
    component.setActivePage(0, new Event('focusin'));

    const newId = component.addFootnoteAtCursor();
    flushPagination();

    expect(newId).not.toBeNull();
    expect(host.querySelector(`sup.footnote-ref[data-footnote-id="${newId}"]`)).not.toBeNull();
    expect(host.querySelector(`.footnote-item[data-footnote-id="${newId}"]`)).not.toBeNull();
    // Trzy przypisy razem.
    expect(host.querySelectorAll('.footnotes-region .footnote-item').length).toBe(3);
  });

  it('ponowne ustawienie inputu footnotes odświeża region', () => {
    const host = load();
    expect(host.querySelectorAll('.footnotes-region .footnote-item').length).toBe(2);

    component.content =
      '<div class="document-content"><p>X<sup class="footnote-ref" data-footnote-id="fn-9">9</sup></p></div>';
    component.footnotes = [{ id: 'fn-9', html: '<p>Jedyny.</p>' }];
    fixture.detectChanges();
    flushPagination();
    component.syncFootnotesWithBody();
    flushPagination();

    const items = Array.from(host.querySelectorAll('.footnotes-region .footnote-item')) as HTMLElement[];
    expect(items.length).toBe(1);
    expect(items[0].getAttribute('data-footnote-id')).toBe('fn-9');
    expect(items[0].querySelector('.footnote-item-content')?.textContent).toContain('Jedyny.');
  });

  it('dokument bez przypisów nie renderuje regionu', () => {
    const host = load('<div class="document-content"><p>Bez przypisów.</p></div>', []);
    expect(host.querySelector('.footnotes-region')).toBeNull();
    expect(host.querySelector('.footnotes-panel')).toBeNull();
    expect(host.querySelector('sup.footnote-ref')).toBeNull();
  });

  it('klik w odnośnik w treści przenosi do wpisu przypisu (scroll + fokus + podświetlenie)', () => {
    const host = load();

    const item = host.querySelector('.footnote-item[data-footnote-id="fn-1"]') as HTMLElement;
    expect(item).not.toBeNull();
    const scrollSpy = vi.fn();
    item.scrollIntoView = scrollSpy;

    const ref = host.querySelector('sup.footnote-ref[data-footnote-id="fn-1"]') as HTMLElement;
    ref.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(scrollSpy).toHaveBeenCalled();
    expect(item.classList.contains('note-item-flash')).toBe(true);
  });

  it('przypis rezerwuje miejsce na dole strony — blok, który nie mieści się z przypisem, spływa na kolejną stronę (region idzie za odwołaniem)', () => {
    // Wysoki przypis (600) nie mieści się razem z drugim na jednej stronie → drugi akapit
    // z odwołaniem trafia na własną stronę wraz ze swoim przypisem.
    const measureStub = (_m: HTMLElement, blocks: HTMLElement[]): number[] =>
      blocks.map(b => {
        if (b.classList?.contains('footnote-item')) return 600;
        if (b.classList?.contains('footnotes-separator')) return 20;
        return 10;
      });
    (component as unknown as { _measureBlockRunHeights: typeof measureStub })._measureBlockRunHeights = measureStub;

    const body =
      '<div class="document-content">' +
      '<p>Pierwszy<sup class="footnote-ref" data-footnote-id="fn-1">1</sup>.</p>' +
      '<p>Drugi<sup class="footnote-ref" data-footnote-id="fn-2">2</sup>.</p>' +
      '</div>';
    const notes: Footnote[] = [
      { id: 'fn-1', html: '<p>Pierwszy przypis.</p>' },
      { id: 'fn-2', html: '<p>Drugi przypis.</p>' },
    ];
    const host = load(body, notes);

    // Rezerwacja rozepchnęła treść na 2 strony, każda z własnym regionem przypisów.
    const pages = Array.from(host.querySelectorAll('.page'));
    expect(pages.length).toBeGreaterThan(1);

    const regions = Array.from(host.querySelectorAll('.footnotes-region'));
    expect(regions.length).toBe(2);

    // Przypis renderuje się na stronie SWOJEGO odwołania.
    const idsOnPage = (pageEl: Element): string[] =>
      Array.from(pageEl.querySelectorAll('.footnotes-region .footnote-item'))
        .map(i => i.getAttribute('data-footnote-id') ?? '');
    expect(idsOnPage(pages[0])).toEqual(['fn-1']);
    expect(idsOnPage(pages[1])).toEqual(['fn-2']);

    // Treść przypisów żyje POZA contenteditable body — nie wchodzi do zapisu treści.
    const saved = component.getContent();
    expect(saved).toContain('Pierwszy');
    expect(saved).not.toContain('Pierwszy przypis.');
    expect(saved).not.toContain('footnote-item');
  });

  it('numeracja wg formatu z dokumentu (w:numFmt): upperLetter → A, B w odwołaniach i wpisach', () => {
    component.footnoteNumberFormat = 'upperLetter';
    const host = load();

    // Odwołania w treści (<sup>) przeliczone przez syncFootnotesWithBody na format dokumentu.
    const refs = Array.from(host.querySelectorAll('sup.footnote-ref')) as HTMLElement[];
    expect(refs.map(r => r.textContent)).toEqual(['A', 'B']);
    expect(refs[0].getAttribute('aria-label')).toBe('Przypis A');

    // Etykiety wpisów regionu też wg formatu.
    const items = Array.from(host.querySelectorAll('.footnotes-region .footnote-item')) as HTMLElement[];
    expect(items.map(i => i.querySelector('.footnote-item-number')?.textContent)).toEqual(['A', 'B']);
  });

  it('bez formatu z dokumentu przypisy dolne domyślnie cyframi (jak MS Word)', () => {
    const host = load();
    const refs = Array.from(host.querySelectorAll('sup.footnote-ref')) as HTMLElement[];
    expect(refs.map(r => r.textContent)).toEqual(['1', '2']);
  });

  it('_formatNoteLabel + _toWordLetters: rzymskie/litery jak Word', () => {
    const fmt = (component as unknown as {
      _formatNoteLabel(n: number, f: string | undefined, e: boolean): string;
    })._formatNoteLabel.bind(component);
    expect(fmt(4, 'lowerRoman', false)).toBe('iv');
    expect(fmt(4, 'upperRoman', false)).toBe('IV');
    expect(fmt(1, 'lowerLetter', false)).toBe('a');
    expect(fmt(27, 'lowerLetter', false)).toBe('aa'); // Word: powtórzona litera, nie „aa" bijektywne
    expect(fmt(28, 'upperLetter', false)).toBe('BB');
    // Fallback wg typu: dolne = cyfry, końcowe = rzymskie.
    expect(fmt(3, undefined, false)).toBe('3');
    expect(fmt(3, undefined, true)).toBe('iii');
    expect(fmt(3, 'wtf-nieznany', false)).toBe('3');
  });
});
