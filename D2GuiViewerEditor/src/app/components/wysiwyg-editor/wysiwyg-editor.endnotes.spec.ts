import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';
import { Endnote, Footnote } from '../../models/document.model';

/**
 * Endnote rendering + editing in the REAL editor component, asserted against the REAL DOM. Mirrors
 * the footnote spec but on the SEPARATE endnote channel (class="endnote-ref", data-endnote-id,
 * endnotesChange, .endnotes-panel). Also asserts footnotes and endnotes coexist without collision.
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

  function load(body = BODY, endnotes = ENDNOTES): HTMLElement {
    component.content = body;
    component.endnotes = endnotes.map(e => ({ ...e }));
    fixture.detectChanges();
    return fixture.nativeElement as HTMLElement;
  }

  it('renderuje odwołania jako <sup class="endnote-ref"> z numerem, id i aria-label', () => {
    const host = load();

    const refs = Array.from(host.querySelectorAll('sup.endnote-ref')) as HTMLElement[];
    expect(refs.length).toBe(2);
    expect(refs[0].getAttribute('data-endnote-id')).toBe('en-1');
    expect(refs[0].textContent).toBe('1');
    expect(refs[0].getAttribute('aria-label')).toBe('Przypis końcowy 1');
    expect(refs[1].getAttribute('data-endnote-id')).toBe('en-2');
    expect(refs[1].getAttribute('aria-label')).toBe('Przypis końcowy 2');
  });

  it('renderuje treści przypisów końcowych w panelu, w kolejności odwołań', () => {
    const host = load();

    const panel = host.querySelector('.endnotes-panel');
    expect(panel).not.toBeNull();

    const items = Array.from(host.querySelectorAll('.endnotes-panel .footnote-item')) as HTMLElement[];
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
    fixture.detectChanges();

    expect(host.querySelector('sup.endnote-ref[data-endnote-id="en-1"]')).toBeNull();
    expect(host.querySelector('.footnote-item[data-endnote-id="en-1"]')).toBeNull();

    const remainingRef = host.querySelector('sup.endnote-ref[data-endnote-id="en-2"]') as HTMLElement;
    expect(remainingRef.textContent).toBe('1');
    expect(remainingRef.getAttribute('aria-label')).toBe('Przypis końcowy 1');

    const model = emitted as unknown as Endnote[];
    expect(model.map(e => e.id)).toEqual(['en-2']);
  });

  it('dodanie przypisu końcowego w pozycji kursora wstawia odwołanie i wpis treści', () => {
    const host = load();
    component.setActivePage(0, new Event('focusin'));

    const newId = component.addEndnoteAtCursor();
    fixture.detectChanges();

    expect(newId).not.toBeNull();
    expect(host.querySelector(`sup.endnote-ref[data-endnote-id="${newId}"]`)).not.toBeNull();
    expect(host.querySelector(`.footnote-item[data-endnote-id="${newId}"]`)).not.toBeNull();
    expect(host.querySelectorAll('.endnotes-panel .footnote-item').length).toBe(3);
  });

  it('dokument bez przypisów końcowych nie renderuje panelu', () => {
    const host = load('<div class="document-content"><p>Bez przypisów.</p></div>', []);
    expect(host.querySelector('.endnotes-panel')).toBeNull();
    expect(host.querySelector('sup.endnote-ref')).toBeNull();
  });

  it('footnotes i endnotes współistnieją bez kolizji (osobne panele i odwołania)', () => {
    const body =
      '<div class="document-content"><p>Tekst' +
      '<sup class="footnote-ref" data-footnote-id="fn-1" aria-label="Przypis 1">1</sup>' +
      '<sup class="endnote-ref" data-endnote-id="en-1" aria-label="Przypis końcowy 1">1</sup>.</p></div>';
    component.content = body;
    component.footnotes = [{ id: 'fn-1', html: '<p>Dolny.</p>' } as Footnote];
    component.endnotes = [{ id: 'en-1', html: '<p>Końcowy.</p>' }];
    fixture.detectChanges();
    const host = fixture.nativeElement as HTMLElement;

    expect(host.querySelector('.footnotes-panel')).not.toBeNull();
    expect(host.querySelector('.endnotes-panel')).not.toBeNull();
    expect(host.querySelector('.footnotes-panel [data-testid="footnote-content-fn-1"]')?.textContent).toContain('Dolny.');
    expect(host.querySelector('.endnotes-panel [data-testid="endnote-content-en-1"]')?.textContent).toContain('Końcowy.');
  });
});
