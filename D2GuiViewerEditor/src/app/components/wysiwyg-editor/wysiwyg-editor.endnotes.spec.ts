import { vi } from 'vitest';
import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';
import { Endnote, Footnote } from '../../models/document.model';

/**
 * Endnote rendering + editing in the REAL editor component, asserted against the REAL DOM.
 */
describe('WysiwygEditorComponent — przypisy końcowe', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  const BODY =
    '<div class="document-content">' +
    '<p>Alfa<sup class="endnote-ref" data-endnote-id="en-1" aria-label="Przypis końcowy 1">1</sup> beta.</p>' +
    '<p>Gamma<sup class="endnote-ref" data-endnote-id="en-2" aria-label="Przypis końcowy 2">2</sup>.</p>' +
    '</div>';

  const ENDNOTES: Endnote[] = [
    { id: 'en-1', html: '<p>Pierwszy przypis końcowy.</p>' },
    { id: 'en-2', html: '<p>Drugi przypis końcowy.</p>' }
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

  function load(body = BODY, endnotes = ENDNOTES): HTMLElement {
    component.content = body;
    component.endnotes = endnotes.map(e => ({ ...e }));
    fixture.detectChanges();
    flushPagination();
    // Note numbering (Roman for endnotes) is applied by sync, which runs asynchronously on a
    // real document load; trigger it explicitly so the rendered markers are deterministic.
    component.syncEndnotesWithBody();
    fixture.detectChanges();
    return fixture.nativeElement as HTMLElement;
  }

  it('numeruje przypisy końcowe małymi cyframi rzymskimi (jak MS Word)', () => {
    const toRoman = (component as unknown as { _toLowerRoman(n: number): string })._toLowerRoman.bind(component);
    expect([1, 2, 3, 4, 5, 9, 10, 14, 40, 90, 100].map(toRoman)).toEqual(
      ['i', 'ii', 'iii', 'iv', 'v', 'ix', 'x', 'xiv', 'xl', 'xc', 'c']
    );
  });

  it('renderuje odwołania jako <sup class="endnote-ref"> z numerem, id i aria-label', () => {
    const host = load();

    const refs = Array.from(host.querySelectorAll('sup.endnote-ref')) as HTMLElement[];
    expect(refs.length).toBe(2);
    expect(refs[0].getAttribute('data-endnote-id')).toBe('en-1');
    // Endnotes use lowercase Roman numerals (i, ii, iii…) like MS Word, to distinguish
    // them from footnotes (Arabic).
    expect(refs[0].textContent).toBe('i');
    expect(refs[0].getAttribute('aria-label')).toBe('Przypis końcowy i');
    expect(refs[1].getAttribute('data-endnote-id')).toBe('en-2');
    expect(refs[1].getAttribute('aria-label')).toBe('Przypis końcowy ii');
  });

  it('renderuje region przypisów WEWNĄTRZ ostatniej strony (nie pod dokumentem), z separatorem', () => {
    const host = load();

    const region = host.querySelector('.endnotes-region');
    expect(region).not.toBeNull();
    expect(region!.closest('.page')).not.toBeNull();
    expect(host.querySelector('.endnotes-panel')).toBeNull();
    const separator = region!.querySelector('.endnotes-separator');
    expect(separator).not.toBeNull();
    expect(separator!.classList.contains('endnotes-separator-continuation')).toBe(false);

    const items = Array.from(host.querySelectorAll('.endnotes-region .footnote-item')) as HTMLElement[];
    expect(items.length).toBe(2);
    expect(items[0].getAttribute('data-endnote-id')).toBe('en-1');
    expect(items[1].getAttribute('data-endnote-id')).toBe('en-2');
    expect(items[0].querySelector('.footnote-item-content')?.textContent).toContain('Pierwszy przypis końcowy.');
    expect(items[1].querySelector('.footnote-item-content')?.textContent).toContain('Drugi przypis końcowy.');
  });

  it('edycja treści przypisu końcowego emituje zmieniony model przez endnotesChange', () => {
    const host = load();
    let emitted: Endnote[] | null = null;
    component.endnotesChange.subscribe(v => (emitted = v));

    const content = host.querySelector('[data-testid="endnote-content-en-1"]') as HTMLElement;
    content.innerHTML = '<p>Zmieniona treść końcowa.</p>';
    content.dispatchEvent(new Event('blur'));

    const model = emitted as unknown as Endnote[];
    expect(model).not.toBeNull();
    expect(model.find(e => e.id === 'en-1')?.html).toContain('Zmieniona treść końcowa.');
    expect(model.find(e => e.id === 'en-2')?.html).toContain('Drugi przypis końcowy.');
  });

  it('usunięcie przypisu końcowego kasuje odwołanie i treść oraz przelicza numerację', () => {
    const host = load();
    let emitted: Endnote[] | null = null;
    component.endnotesChange.subscribe(v => (emitted = v));

    component.removeEndnote('en-1');
    flushPagination();

    expect(host.querySelector('sup.endnote-ref[data-endnote-id="en-1"]')).toBeNull();
    expect(host.querySelector('.footnote-item[data-endnote-id="en-1"]')).toBeNull();

    const remainingRef = host.querySelector('sup.endnote-ref[data-endnote-id="en-2"]') as HTMLElement;
    expect(remainingRef.textContent).toBe('i');
    expect(remainingRef.getAttribute('aria-label')).toBe('Przypis końcowy i');
    const remainingItem = host.querySelector('.footnote-item[data-endnote-id="en-2"] .footnote-item-number');
    expect(remainingItem?.textContent).toBe('i');

    const model = emitted as unknown as Endnote[];
    expect(model.map(e => e.id)).toEqual(['en-2']);
  });

  it('dodanie przypisu końcowego w pozycji kursora wstawia odwołanie i wpis treści', () => {
    const host = load();
    component.setActivePage(0, new Event('focusin'));

    const newId = component.addEndnoteAtCursor();
    flushPagination();

    expect(newId).not.toBeNull();
    expect(host.querySelector(`sup.endnote-ref[data-endnote-id="${newId}"]`)).not.toBeNull();
    expect(host.querySelector(`.footnote-item[data-endnote-id="${newId}"]`)).not.toBeNull();
    expect(host.querySelectorAll('.endnotes-region .footnote-item').length).toBe(3);
  });

  it('dokument bez przypisów końcowych nie renderuje regionu', () => {
    const host = load('<div class="document-content"><p>Bez przypisów.</p></div>', []);
    expect(host.querySelector('.endnotes-region')).toBeNull();
    expect(host.querySelector('sup.endnote-ref')).toBeNull();
  });

  it('footnotes i endnotes współistnieją bez kolizji (region dolnych + region końcowych)', () => {
    const body =
      '<div class="document-content"><p>Tekst' +
      '<sup class="footnote-ref" data-footnote-id="fn-1" aria-label="Przypis 1">1</sup>' +
      '<sup class="endnote-ref" data-endnote-id="en-1" aria-label="Przypis końcowy 1">1</sup>.</p></div>';
    component.content = body;
    component.footnotes = [{ id: 'fn-1', html: '<p>Dolny.</p>' } as Footnote];
    component.endnotes = [{ id: 'en-1', html: '<p>Końcowy.</p>' }];
    fixture.detectChanges();
    flushPagination();
    const host = fixture.nativeElement as HTMLElement;

    expect(host.querySelector('.footnotes-region')).not.toBeNull();
    expect(host.querySelector('.endnotes-region')).not.toBeNull();
    expect(host.querySelector('.footnotes-region [data-testid="footnote-content-fn-1"]')?.textContent).toContain('Dolny.');
    expect(host.querySelector('.endnotes-region [data-testid="endnote-content-en-1"]')?.textContent).toContain('Końcowy.');
  });

  it('klik w odnośnik w treści przenosi do wpisu przypisu (scroll + fokus + podświetlenie)', () => {
    const host = load();

    const item = host.querySelector('.footnote-item[data-endnote-id="en-1"]') as HTMLElement;
    expect(item).not.toBeNull();
    const scrollSpy = vi.fn();
    item.scrollIntoView = scrollSpy;

    const ref = host.querySelector('sup.endnote-ref[data-endnote-id="en-1"]') as HTMLElement;
    ref.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(scrollSpy).toHaveBeenCalled();
    expect(item.classList.contains('note-item-flash')).toBe(true);
  });

  it('klik w numer wpisu wraca do odwołania w treści', () => {
    const host = load();

    const ref = host.querySelector('sup.endnote-ref[data-endnote-id="en-2"]') as HTMLElement;
    const scrollSpy = vi.fn();
    ref.scrollIntoView = scrollSpy;

    const numberEl = host.querySelector(
      '.footnote-item[data-endnote-id="en-2"] .footnote-item-number'
    ) as HTMLElement;
    numberEl.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(scrollSpy).toHaveBeenCalled();
  });

  it('przypisy niemieszczące się na ostatniej stronie przelewają się na strony tylko-przypisowe z pustym body (poza zapisem)', async () => {
    const measureStub = (_m: HTMLElement, blocks: HTMLElement[]): number[] =>
      blocks.map(b => {
        if (b.classList?.contains('footnote-item')) return 700;
        if (b.classList?.contains('endnotes-separator')) return 20;
        return 10;
      });
    (component as unknown as { _measureBlockRunHeights: typeof measureStub })._measureBlockRunHeights = measureStub;

    const body =
      '<div class="document-content">' +
      '<p>Treść<sup class="endnote-ref" data-endnote-id="en-1">1</sup>' +
      '<sup class="endnote-ref" data-endnote-id="en-2">2</sup>' +
      '<sup class="endnote-ref" data-endnote-id="en-3">3</sup>.</p>' +
      '</div>';
    const notes: Endnote[] = [
      { id: 'en-1', html: '<p>Pierwszy.</p>' },
      { id: 'en-2', html: '<p>Drugi.</p>' },
      { id: 'en-3', html: '<p>Trzeci.</p>' },
    ];
    const host = load(body, notes);
    await new Promise(resolve => setTimeout(resolve, 0));
    fixture.detectChanges();

    const pages = Array.from(host.querySelectorAll('.page'));
    expect(pages.length).toBeGreaterThan(1);

    const items = Array.from(host.querySelectorAll('.endnotes-region .footnote-item')) as HTMLElement[];
    expect(items.map(i => i.getAttribute('data-endnote-id'))).toEqual(['en-1', 'en-2', 'en-3']);
    expect(items.map(i => i.querySelector('.footnote-item-number')?.textContent)).toEqual(['i', 'ii', 'iii']);

    const regions = Array.from(host.querySelectorAll('.endnotes-region'));
    expect(regions.length).toBeGreaterThan(1);
    const contSeparators = regions.slice(1).map(r => r.querySelector('.endnotes-separator'));
    contSeparators.forEach(sep =>
      expect(sep?.classList.contains('endnotes-separator-continuation')).toBe(true));

    const saved = component.getContent();
    expect(saved).toContain('Treść');
    expect(saved).not.toContain('Pierwszy.');
    expect(saved).not.toContain('footnote-item');
    const editors = Array.from(host.querySelectorAll('.editor-content')) as HTMLElement[];
    expect(editors[editors.length - 1].textContent?.trim()).toBe('');
  });
});
