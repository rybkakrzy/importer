import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';
import { HeaderFooterContent } from '../../models/document.model';

/**
 * Variant routing for header/footer editing — guards rule 10 ("no apparent
 * editing"): when the user edits page 0 of a document with `differentFirstPage`
 * set, the change must land in `firstPageHtml`, not silently overwrite the
 * default content. Same for footer.
 *
 * The component has heavy DOM/ViewChild wiring; we exercise the public Input
 * setter and the input/blur handlers directly without detectChanges (mirrors
 * document-editor.spec).
 */
describe('WysiwygEditorComponent — variant editing routing', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  function inputEvent(html: string): Event {
    return { target: { innerHTML: html } } as unknown as Event;
  }

  const hf = (overrides: Partial<HeaderFooterContent>): HeaderFooterContent => ({
    html: 'DEFAULT', height: 1.27, ...overrides
  });

  it('writes default header content to _headerHtml when no variant is active', () => {
    component.headerContent = hf({ html: 'DEFAULT' });

    component.onHeaderInput(inputEvent('EDITED'));

    expect((component as any)._headerHtml()).toBe('EDITED');
    expect((component as any)._headerFirstPageHtml()).toBe('');
  });

  it('writes first-page header content to _headerFirstPageHtml when differentFirstPage is set', () => {
    component.headerContent = hf({
      html: 'DEFAULT',
      differentFirstPage: true,
      firstPageHtml: 'FIRST'
    });

    component.onHeaderInput(inputEvent('EDITED-FIRST'));

    expect((component as any)._headerFirstPageHtml()).toBe('EDITED-FIRST');
    // Default content must be untouched — that is the bug rule 10 forbids.
    expect((component as any)._headerHtml()).toBe('DEFAULT');
  });

  it('routes footer input the same way (first-page → _footerFirstPageHtml)', () => {
    component.footerContent = hf({
      html: 'DEFAULT',
      differentFirstPage: true,
      firstPageHtml: 'FIRST'
    });

    component.onFooterInput(inputEvent('EDITED-FIRST'));

    expect((component as any)._footerFirstPageHtml()).toBe('EDITED-FIRST');
    expect((component as any)._footerHtml()).toBe('DEFAULT');
  });

  it('emits the FULL header object (incl. firstPageHtml/evenHtml) on every change', () => {
    component.headerContent = hf({
      html: 'DEFAULT',
      differentFirstPage: true, firstPageHtml: 'FIRST',
      differentOddEven: true, evenHtml: 'EVEN'
    });
    const seen: HeaderFooterContent[] = [];
    component.headerChange.subscribe(v => seen.push(v));

    component.onHeaderInput(inputEvent('EDITED-FIRST'));

    expect(seen.length).toBe(1);
    expect(seen[0].firstPageHtml).toBe('EDITED-FIRST');
    expect(seen[0].html).toBe('DEFAULT');
    expect(seen[0].differentFirstPage).toBe(true);
    expect(seen[0].differentOddEven).toBe(true);
    expect(seen[0].evenHtml).toBe('EVEN');
  });

  it('odd page uses canonical _headerHtml (default = odd in OOXML)', () => {
    component.headerContent = hf({
      html: 'ODD-DEFAULT',
      differentOddEven: true, evenHtml: 'EVEN'
    });

    // page 0 is odd → must surface the canonical default.
    const odd = (component as any)._computeHeaderContent(0);
    const even = (component as any)._computeHeaderContent(1);

    expect(odd).toBe('ODD-DEFAULT');
    expect(even).toBe('EVEN');
  });

  it('an empty document (no header/footer) does not crash on input', () => {
    component.headerContent = hf({ html: '' });

    expect(() => component.onHeaderInput(inputEvent(''))).not.toThrow();
    expect((component as any)._headerHtml()).toBe('');
  });

  it('stopEditingHeaderFooter returns the editor to body mode', () => {
    (component as any).editingSection.set('header');

    component.stopEditingHeaderFooter();

    expect(component.editingSection()).toBe('body');
  });
});
