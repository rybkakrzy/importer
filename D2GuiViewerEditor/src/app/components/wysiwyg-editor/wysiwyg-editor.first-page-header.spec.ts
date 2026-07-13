import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';
import { HeaderFooterContent, SectionHeaderFooter } from '../../models/document.model';

/**
 * "Inna pierwsza strona" (w:titlePg) — semantyka DOCX w edytorze.
 *
 * Regresje pilnowane tutaj:
 * 1. Flagi wariantów są PER PASMO: binding [footerContent] aplikowany po [headerContent]
 *    nie może wyzerować flagi nagłówka (dokument z nagłówkiem first bez stopki first
 *    renderował na stronie 1 wariant default — zgłoszenie DEV).
 * 2. Wariant first dotyczy pierwszej strony KAŻDEJ sekcji (pageSectionIndexes), nie tylko
 *    strony 0 dokumentu; wpisy sekcyjne niosą własne warianty first/even i dziedziczą
 *    brakujące (null) z wcześniejszych sekcji jak Word.
 * 3. Jawnie pusty wariant first ('') = puste pasmo, NIE fallback do default.
 */
describe('WysiwygEditorComponent — titlePg / first-page variant selection', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  const hf = (overrides: Partial<HeaderFooterContent>): HeaderFooterContent => ({
    html: 'DEFAULT', height: 1.27, ...overrides
  });

  const headerFor = (page: number): string => (component as any)._computeHeaderContent(page);
  const footerFor = (page: number): string => (component as any)._computeFooterContent(page);

  it('footer input applied AFTER header input must not clobber the header first-page flag', () => {
    // Kolejność jak w szablonie document-editor: [headerContent] przed [footerContent].
    component.headerContent = hf({ html: 'DEF-H', differentFirstPage: true, firstPageHtml: 'FIRST-H' });
    component.footerContent = hf({ html: 'DEF-F', differentFirstPage: false });

    expect(headerFor(0)).toBe('FIRST-H');
    expect(headerFor(1)).toBe('DEF-H');
    expect(footerFor(0)).toBe('DEF-F');
  });

  it('header input applied AFTER footer input must not clobber the footer first-page flag', () => {
    component.footerContent = hf({ html: 'DEF-F', differentFirstPage: true, firstPageHtml: 'FIRST-F' });
    component.headerContent = hf({ html: 'DEF-H', differentFirstPage: false });

    expect(footerFor(0)).toBe('FIRST-F');
    expect(footerFor(1)).toBe('DEF-F');
    expect(headerFor(0)).toBe('DEF-H');
  });

  it('an explicitly EMPTY first-page variant renders blank, not the default', () => {
    component.headerContent = hf({ html: 'DEF-H', differentFirstPage: true, firstPageHtml: '' });

    expect(headerFor(0)).toBe('');
    expect(headerFor(1)).toBe('DEF-H');
  });

  it('emits per-band flags (header true, footer false) after an edit', () => {
    component.headerContent = hf({ html: 'DEF-H', differentFirstPage: true, firstPageHtml: 'FIRST-H' });
    component.footerContent = hf({ html: 'DEF-F', differentFirstPage: false });
    const headers: HeaderFooterContent[] = [];
    const footers: HeaderFooterContent[] = [];
    component.headerChange.subscribe(v => headers.push(v));
    component.footerChange.subscribe(v => footers.push(v));

    component.onHeaderInput({ target: { innerHTML: 'EDITED-FIRST' } } as unknown as Event);

    expect(headers[0].differentFirstPage).toBe(true);
    expect(headers[0].firstPageHtml).toBe('EDITED-FIRST');
    expect(footers[0].differentFirstPage).toBe(false);
    expect(footers[0].html).toBe('DEF-F');
  });

  it('toggleDifferentFirstPage flips BOTH bands to the same value (Word section setting)', () => {
    component.headerContent = hf({ html: 'DEF-H', differentFirstPage: true, firstPageHtml: 'FIRST-H' });
    component.footerContent = hf({ html: 'DEF-F', differentFirstPage: false });

    component.toggleDifferentFirstPage();
    expect(component.differentFirstPage()).toBe(false);
    expect(headerFor(0)).toBe('DEF-H');

    component.toggleDifferentFirstPage();
    expect(component.differentFirstPage()).toBe(true);
    expect(headerFor(0)).toBe('FIRST-H');
    expect(footerFor(0)).toBe('');
  });

  describe('multi-section documents', () => {
    const sectionEntries: SectionHeaderFooter[] = [
      {
        sectionIndex: 1,
        header: { html: 'S2-DEF', height: 1.27, differentFirstPage: true, firstPageHtml: 'S2-FIRST' }
      }
    ];

    beforeEach(() => {
      component.headerContent = hf({ html: 'BASE-DEF', differentFirstPage: false });
      component.sectionHeadersFooters = sectionEntries;
      // Strony 0-1 = sekcja 0, strony 2-3 = sekcja 1.
      (component as any).pageSectionIndexes.set([0, 0, 1, 1]);
    });

    it('uses the FIRST variant on the first page of EACH section, not only page 0', () => {
      expect(headerFor(0)).toBe('BASE-DEF');   // sekcja 0 bez titlePg
      expect(headerFor(1)).toBe('BASE-DEF');
      expect(headerFor(2)).toBe('S2-FIRST');   // pierwsza strona sekcji 1
      expect(headerFor(3)).toBe('S2-DEF');     // kolejna strona sekcji 1
    });

    it('inherits a missing first variant (null) from the base like Word links sections', () => {
      component.headerContent = hf({ html: 'BASE-DEF', differentFirstPage: true, firstPageHtml: 'BASE-FIRST' });
      component.sectionHeadersFooters = [
        { sectionIndex: 1, header: { html: 'S2-DEF', height: 1.27, differentFirstPage: true } }
      ];

      expect(headerFor(0)).toBe('BASE-FIRST');
      // Brak własnego first w sekcji 1 → dziedziczy z bazy.
      expect(headerFor(2)).toBe('BASE-FIRST');
      expect(headerFor(3)).toBe('S2-DEF');
    });

    it('routes an edit on the section first page into the entry firstPageHtml', () => {
      const emitted: SectionHeaderFooter[][] = [];
      component.sectionHeadersFootersChange.subscribe(v => emitted.push(v));
      (component as any).editingHfPageIndex.set(2);

      component.onHeaderInput({ target: { innerHTML: 'S2-FIRST-EDITED' } } as unknown as Event);

      expect(emitted.length).toBe(1);
      const entry = emitted[0].find(e => e.sectionIndex === 1)!;
      expect(entry.header!.firstPageHtml).toBe('S2-FIRST-EDITED');
      expect(entry.header!.html).toBe('S2-DEF');
      expect(headerFor(2)).toBe('S2-FIRST-EDITED');
    });

    it('routes an edit on an ordinary section page into the entry default html', () => {
      (component as any).editingHfPageIndex.set(3);

      component.onHeaderInput({ target: { innerHTML: 'S2-DEF-EDITED' } } as unknown as Event);

      expect(headerFor(3)).toBe('S2-DEF-EDITED');
      // Wariant first nietknięty.
      expect(headerFor(2)).toBe('S2-FIRST');
    });
  });
});
