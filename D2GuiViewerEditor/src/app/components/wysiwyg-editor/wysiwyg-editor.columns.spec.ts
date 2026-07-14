import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent, parseColumnDataAttributes } from './wysiwyg-editor';

/**
 * ADR-0039: układ wielokolumnowy (w:cols). Front: parsowanie data-col-*, geometria kolumn
 * (baza + marker sekcji), render CSS (column-count/gap/rule), atomowy podział kolumny w
 * paginacji oraz round-trip atrybutów kontenera przez getContent/_wrapWithDocumentContainer.
 */
describe('WysiwygEditorComponent — kolumny (ADR-0039)', () => {
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

  afterEach(() => window.getSelection()?.removeAllRanges());

  // ---------- parseColumnDataAttributes ----------

  it('parsuje 2 równe kolumny z separatorem (twipy → cm)', () => {
    const el = document.createElement('div');
    el.setAttribute('data-col-count', '2');
    el.setAttribute('data-col-space-tw', '720');
    el.setAttribute('data-col-equal', '1');
    el.setAttribute('data-col-sep', '1');

    const geo = parseColumnDataAttributes(el)!;
    expect(geo.count).toBe(2);
    expect(geo.equalWidth).toBe(true);
    expect(geo.separator).toBe(true);
    expect(geo.spaceCm).toBeCloseTo(720 / (1440 / 2.54), 3); // 1.27 cm
  });

  it('jedna kolumna (lub brak) → undefined (nie renderuje kolumn)', () => {
    const one = document.createElement('div');
    one.setAttribute('data-col-count', '1');
    expect(parseColumnDataAttributes(one)).toBeUndefined();
    expect(parseColumnDataAttributes(document.createElement('div'))).toBeUndefined();
  });

  it('parsuje kolumny nierówne (widths/spaces CSV → cm)', () => {
    const el = document.createElement('div');
    el.setAttribute('data-col-count', '2');
    el.setAttribute('data-col-equal', '0');
    el.setAttribute('data-col-widths-tw', '3402,5000');
    el.setAttribute('data-col-spaces-tw', '425,0');

    const geo = parseColumnDataAttributes(el)!;
    expect(geo.equalWidth).toBe(false);
    expect(geo.widthsCm!.length).toBe(2);
    expect(geo.widthsCm![0]).toBeCloseTo(3402 / (1440 / 2.54), 2);
  });

  // ---------- geometria bazowa i markera ----------

  it('_captureDocumentDefaults wczytuje kolumny sekcji bazowej → baseGeometry', () => {
    const html = '<div class="document-content" data-col-count="2" data-col-space-tw="708" data-col-equal="1"><p>A</p></div>';
    (component as any)._captureDocumentDefaults(html);

    const cols = component.baseGeometry().columns;
    expect(cols).toBeTruthy();
    expect(cols!.count).toBe(2);
    expect(component.getBaseColumnCount()).toBe(2);
  });

  it('_parseSectionGeometry czyta kolumny z markera (bez dziedziczenia)', () => {
    const marker = document.createElement('div');
    marker.className = 'docx-section-break';
    marker.setAttribute('data-col-count', '3');
    marker.setAttribute('data-col-space-tw', '360');

    const base = component.baseGeometry();
    const geo = (component as any)._parseSectionGeometry(marker, base);
    expect(geo.columns.count).toBe(3);

    // Marker bez data-col-* = jednokolumnowa (nie dziedziczy z current).
    const plain = document.createElement('div');
    plain.className = 'docx-section-break';
    expect((component as any)._parseSectionGeometry(plain, geo).columns).toBeUndefined();
  });

  // ---------- render CSS ----------

  it('pageColumnCount/Gap/Rule wynikają z geometrii strony', () => {
    const html = '<div class="document-content" data-col-count="2" data-col-space-tw="567" data-col-sep="1"><p>A</p></div>';
    (component as any)._captureDocumentDefaults(html);

    expect(component.pageColumnCount(0)).toBe(2);
    expect(component.pageColumnGapPx(0)).toBeGreaterThan(0);
    expect(component.pageColumnRule(0)).toContain('solid');

    // Po zniesieniu kolumn (setBaseColumns(1)) strona wraca do jednej kolumny, bez reguły.
    component.setBaseColumns(1);
    expect(component.pageColumnCount(0)).toBe(1);
    expect(component.pageColumnRule(0)).toBe('none');
  });

  // ---------- paginacja: podział kolumny atomowy ----------

  it('_flattenTopBlocks NIE rozwija docx-column-break', () => {
    const host = document.createElement('div');
    host.innerHTML = '<div class="docx-column-break"></div>';
    const blocks = (component as any)._flattenTopBlocks(host);
    expect(blocks.length).toBe(1);
    expect((blocks[0] as HTMLElement).classList.contains('docx-column-break')).toBe(true);
  });

  // ---------- round-trip atrybutów kontenera ----------

  it('kolumny bazowe round-tripują przez _captureDocumentDefaults + _wrapWithDocumentContainer', () => {
    const html = '<div class="document-content" data-col-count="2" data-col-space-tw="708" data-col-equal="1" data-col-sep="1"><p>A</p></div>';
    const inner = (component as any)._captureDocumentDefaults(html);
    const wrapped = (component as any)._wrapWithDocumentContainer(inner);
    expect(wrapped).toContain('data-col-count="2"');
    expect(wrapped).toContain('data-col-space-tw="708"');
    expect(wrapped).toContain('data-col-sep="1"');
  });

  // ---------- edycja: setBaseColumns ----------

  it('setBaseColumns ustawia kolumny na dokumencie bez wrappera (tworzy kontener) i znosi je przy 1', () => {
    // Start bez kontenera (dokument jednokolumnowy z importu).
    (component as any)._captureDocumentDefaults('<p>A</p>');
    expect(component.getBaseColumnCount()).toBe(1);

    component.setBaseColumns(2);
    expect(component.getBaseColumnCount()).toBe(2);
    const wrapped = (component as any)._wrapWithDocumentContainer('<p>A</p>');
    expect(wrapped).toContain('document-content');
    expect(wrapped).toContain('data-col-count="2"');

    component.setBaseColumns(1);
    expect(component.getBaseColumnCount()).toBe(1);
    expect((component as any)._wrapWithDocumentContainer('<p>A</p>')).not.toContain('data-col-count');
  });
});
