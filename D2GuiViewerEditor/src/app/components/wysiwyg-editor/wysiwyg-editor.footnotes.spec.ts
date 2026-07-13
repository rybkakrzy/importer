import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';
import { Footnote } from '../../models/document.model';

/**
 * Footnote rendering + editing in the REAL editor component, asserted against the REAL DOM:
 *  - references render as semantic <sup class="footnote-ref"> with the visible number, a stable
 *    data-footnote-id and an aria-label,
 *  - footnote contents render in a dedicated panel in reference order, each linked to its reference,
 *  - editing footnote content emits an updated model through footnotesChange (save serialization),
 *  - deleting a footnote removes its reference + content and renumbers the rest,
 *  - adding a footnote inserts a reference + a content entry,
 *  - re-binding the footnotes input re-renders the panel.
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

  function load(body = BODY, footnotes = FOOTNOTES): HTMLElement {
    component.content = body;
    component.footnotes = footnotes.map(f => ({ ...f }));
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

  it('renderuje treści przypisów w panelu, w kolejności odwołań, powiązane z odwołaniami', () => {
    const host = load();

    const items = Array.from(host.querySelectorAll('.footnote-item')) as HTMLElement[];
    expect(items.length).toBe(2);

    // Kolejność + numeracja panelu.
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
    fixture.detectChanges();

    // Odwołanie i treść zniknęły.
    expect(host.querySelector('sup.footnote-ref[data-footnote-id="fn-1"]')).toBeNull();
    expect(host.querySelector('.footnote-item[data-footnote-id="fn-1"]')).toBeNull();

    // Pozostały przypis przenumerowany na 1 (odwołanie w treści + panel).
    const remainingRef = host.querySelector('sup.footnote-ref[data-footnote-id="fn-2"]') as HTMLElement;
    expect(remainingRef.textContent).toBe('1');
    expect(remainingRef.getAttribute('aria-label')).toBe('Przypis 1');

    // Model bez osieroconych.
    const model = emitted as unknown as Footnote[];
    expect(model.map(f => f.id)).toEqual(['fn-2']);
  });

  it('dodanie przypisu w pozycji kursora wstawia odwołanie i wpis treści', () => {
    const host = load();

    // Ustaw aktywny edytor (jak focusin na stronie).
    component.setActivePage(0, new Event('focusin'));

    const newId = component.addFootnoteAtCursor();
    fixture.detectChanges();

    expect(newId).not.toBeNull();
    expect(host.querySelector(`sup.footnote-ref[data-footnote-id="${newId}"]`)).not.toBeNull();
    expect(host.querySelector(`.footnote-item[data-footnote-id="${newId}"]`)).not.toBeNull();
    // Trzy przypisy razem.
    expect(host.querySelectorAll('.footnote-item').length).toBe(3);
  });

  it('ponowne ustawienie inputu footnotes odświeża panel', () => {
    const host = load();
    expect(host.querySelectorAll('.footnote-item').length).toBe(2);

    component.footnotes = [{ id: 'fn-9', html: '<p>Jedyny.</p>' }];
    // Body ref dla fn-9, żeby panel miał sens numeracji (choć panel renderuje z modelu).
    fixture.detectChanges();

    const items = Array.from(host.querySelectorAll('.footnote-item')) as HTMLElement[];
    expect(items.length).toBe(1);
    expect(items[0].getAttribute('data-footnote-id')).toBe('fn-9');
    expect(items[0].querySelector('.footnote-item-content')?.textContent).toContain('Jedyny.');
  });

  it('dokument bez przypisów nie renderuje panelu', () => {
    const host = load('<div class="document-content"><p>Bez przypisów.</p></div>', []);
    expect(host.querySelector('.footnotes-panel')).toBeNull();
    expect(host.querySelector('sup.footnote-ref')).toBeNull();
  });
});
