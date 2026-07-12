import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Warstwa etykiet list DOCX w edytorze (ADR-0036, prezentacja):
 *  - refreshListLabels liczy etykiety silnikiem z list-label.util (szablony %1.%2,
 *    kontynuacja fragmentów, startOverride) i zapisuje w data-list-label na li,
 *  - render robi CSS ::before — etykieta jest POZA edytowalnym tekstem,
 *  - getContent() zdejmuje atrybuty prezentacyjne (nie trafiają do zapisu/undo).
 */
describe('WysiwygEditorComponent — etykiety list liczone silnikiem', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;
  let editor: HTMLDivElement;

  const MULTILEVEL_LIST =
    '<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1.">' +
    '<li>jeden' +
    '<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="1" data-num-fmt="lowerLetter" data-lvl-text="%1.%2)">' +
    '<li>pod</li></ol></li>' +
    '<li>dwa</li></ol>';

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;

    editor = document.createElement('div');
    document.body.appendChild(editor);
    (component as any).editorContent = { nativeElement: editor };
  });

  afterEach(() => {
    editor.remove();
  });

  it('refreshListLabels nadaje etykiety z szablonu wielopoziomowego', () => {
    editor.innerHTML = MULTILEVEL_LIST;

    component.refreshListLabels();

    const labels = Array.from(editor.querySelectorAll('li')).map(li =>
      li.getAttribute('data-list-label'),
    );
    expect(labels).toEqual(['1.', '1.a)', '2.']);
  });

  it('fragmenty tej samej listy rozdzielone akapitem kontynuują numerację', () => {
    editor.innerHTML =
      '<ol data-num-id="7" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1."><li>a</li><li>b</li></ol>' +
      '<p>przerwa</p>' +
      '<ol data-num-id="7" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1."><li>c</li></ol>';

    component.refreshListLabels();

    const labels = Array.from(editor.querySelectorAll('li')).map(li =>
      li.getAttribute('data-list-label'),
    );
    expect(labels).toEqual(['1.', '2.', '3.']);
  });

  it('przeliczenie po edycji aktualizuje numery kolejnych fragmentów (statyczny <ol start> tego nie umie)', () => {
    editor.innerHTML =
      '<ol data-num-id="7" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1."><li>a</li></ol>' +
      '<p>przerwa</p>' +
      '<ol start="2" data-num-id="7" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1."><li>c</li></ol>';
    component.refreshListLabels();

    // Edycja: nowy element w PIERWSZYM fragmencie — dalszy fragment musi się przesunąć.
    const firstOl = editor.querySelector('ol')!;
    const li = document.createElement('li');
    li.textContent = 'wstawiony';
    firstOl.appendChild(li);
    component.refreshListLabels();

    const labels = Array.from(editor.querySelectorAll('li')).map(li2 =>
      li2.getAttribute('data-list-label'),
    );
    expect(labels).toEqual(['1.', '2.', '3.']);
  });

  it('getContent() zdejmuje atrybuty prezentacyjne, zostawiając kontrakt data-* writera', () => {
    editor.innerHTML = MULTILEVEL_LIST;
    component.refreshListLabels();
    expect(editor.querySelectorAll('[data-list-label]').length).toBeGreaterThan(0);

    const html = component.getContent();

    expect(html).not.toContain('data-list-label');
    expect(html).not.toContain('data-list-suffix');
    expect(html).toContain('data-num-id="1"');
    expect(html).toContain('data-lvl-text');
    // DOM edytora zachowuje etykiety (strip działa na klonie serializacji).
    expect(editor.querySelectorAll('[data-list-label]').length).toBeGreaterThan(0);
  });

  it('lista utworzona w edytorze (bez data-num-id) zostaje przy natywnych markerach', () => {
    editor.innerHTML = '<ol><li>natywna</li></ol>';

    component.refreshListLabels();

    expect(editor.querySelectorAll('[data-list-label]').length).toBe(0);
  });
});
