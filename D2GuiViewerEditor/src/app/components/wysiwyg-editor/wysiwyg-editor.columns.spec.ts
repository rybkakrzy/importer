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

    // Po zniesieniu kolumn (setBaseColumns(1)) strona wraca do jednej kolumny: null zdejmuje
    // column-count/rule z bindingu — strona nie jest wtedy kontenerem multicol.
    component.setBaseColumns(1);
    expect(component.pageColumnCount(0)).toBeNull();
    expect(component.pageColumnRule(0)).toBeNull();
  });

  // ---------- paginacja: podział kolumny atomowy ----------

  it('_flattenTopBlocks NIE rozwija docx-column-break', () => {
    const host = document.createElement('div');
    host.innerHTML = '<div class="docx-column-break"></div>';
    const blocks = (component as any)._flattenTopBlocks(host);
    expect(blocks.length).toBe(1);
    expect((blocks[0] as HTMLElement).classList.contains('docx-column-break')).toBe(true);
  });

  // ---------- paginacja: pojemność strony wielokolumnowej ----------

  const bodyEditors: HTMLElement[] = [];
  afterEach(() => bodyEditors.splice(0).forEach(el => el.remove()));

  function pageWith(html: string): HTMLDivElement {
    const editor = document.createElement('div');
    editor.innerHTML = html;
    document.body.appendChild(editor);
    bodyEditors.push(editor);
    (component as any).pageEditorRefs = { toArray: () => [{ nativeElement: editor }] };
    return editor;
  }

  /** Stub pomiaru (jsdom bez layoutu): wysokość bloku z atrybutu data-h. */
  function stubBlockHeights(): void {
    (component as any)._measureBlockRunHeights = (_m: HTMLElement, blocks: HTMLElement[]) =>
      blocks.map(b => parseFloat(b.getAttribute('data-h') ?? '0') || 0);
  }

  function setTwoColumnBase(): void {
    (component as any)._captureDocumentDefaults(
      '<div class="document-content" data-col-count="2" data-col-space-tw="708" data-col-equal="1"><p>x</p></div>'
    );
  }

  it('strona 2-kolumnowa mieści ~2× więcej treści niż 1-kolumnowa (pojemność = n×kolumna)', () => {
    // A4, marginesy domyślne → kolumna ~933 px. 5 bloków × 600 px:
    // 1 kolumna: 2 bloki > 933 → blok na stronę (5 stron);
    // 2 kolumny: pojemność ~1847 px → 3 bloki na stronę 1, 2 na stronę 2.
    setTwoColumnBase();
    pageWith('<p data-h="600">1</p><p data-h="600">2</p><p data-h="600">3</p><p data-h="600">4</p><p data-h="600">5</p>');
    stubBlockHeights();

    (component as any)._repaginateNow();

    const pages = component.pageContents();
    expect(pages.length).toBe(2);
    expect(pages[0]).toContain('<p data-h="600">3</p>');
    expect(pages[0]).not.toContain('<p data-h="600">4</p>');
    expect(pages.join('')).toContain('<p data-h="600">5</p>');
  });

  it('bez kolumn te same bloki łamią się per kolumna strony (regresja 1-kolumnowa)', () => {
    (component as any)._captureDocumentDefaults('<p>x</p>');
    pageWith('<p data-h="600">1</p><p data-h="600">2</p><p data-h="600">3</p>');
    stubBlockHeights();

    (component as any)._repaginateNow();

    expect(component.pageContents().length).toBe(3);
  });

  it('docx-column-break konsumuje resztę kolumny (treść za nim liczy się od następnej kolumny)', () => {
    setTwoColumnBase();
    // 600 + skok do kolumny 2 (933) + 600 = ~1533 ≤ pojemność (~1847) → wszystko na 1 stronie;
    // bez skoku trzeci blok też by się zmieścił — pinujemy, że skok jest KONSUMOWANY,
    // dokładając blok, który przez skok już się nie mieści.
    pageWith(
      '<p data-h="600">A</p><div class="docx-column-break"></div>' +
      '<p data-h="600">B</p><p data-h="600">C</p>'
    );
    stubBlockHeights();

    (component as any)._repaginateNow();

    const pages = component.pageContents();
    expect(pages.length).toBe(2);
    expect(pages[0]).toContain('>A</p>');
    expect(pages[0]).toContain('docx-column-break');
    expect(pages[0]).toContain('>B</p>');
    expect(pages[1]).toContain('>C</p>');
  });

  it('docx-column-break w ostatniej kolumnie otwiera nową stronę (marker jedzie z dalszą treścią)', () => {
    setTwoColumnBase();
    pageWith(
      '<p data-h="600">A</p><div class="docx-column-break"></div>' +
      '<p data-h="600">B</p><div class="docx-column-break"></div>' +
      '<p data-h="600">C</p>'
    );
    stubBlockHeights();

    (component as any)._repaginateNow();

    const pages = component.pageContents();
    expect(pages.length).toBe(2);
    expect(pages[0]).toContain('>B</p>');
    expect(pages[0]).not.toContain('>C</p>');
    expect(pages[1]).toContain('docx-column-break');
    expect(pages[1]).toContain('>C</p>');
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
