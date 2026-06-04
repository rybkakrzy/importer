import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Two behaviours that the per-page contenteditable design needs explicit support for:
 *
 * 1. Cross-page caret navigation — each page is a separate contenteditable, so the browser
 *    won't carry the caret from the bottom of one page to the next with ArrowDown/Up. We move
 *    it ourselves at the boundary, but must NOT hijack Shift+Arrow (text selection).
 * 2. Page-number font-size — an inserted page number must inherit the header/footer's font-size,
 *    not the editor container default.
 */
describe('WysiwygEditorComponent — cross-page caret + page-number font', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  // --- caret navigation -----------------------------------------------------

  function setupTwoPages(): { page0: HTMLDivElement; page1: HTMLDivElement } {
    const page0 = document.createElement('div');
    page0.textContent = 'Tekst pierwszej strony';
    const page1 = document.createElement('div');
    page1.textContent = 'Tekst drugiej strony';
    document.body.appendChild(page0);
    document.body.appendChild(page1);
    const refs = [{ nativeElement: page0 }, { nativeElement: page1 }];
    (component as any).pageEditorRefs = { toArray: () => refs };
    return { page0, page1 };
  }

  function collapseCaretIn(node: Node, offset: number): void {
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
  }

  // Edge-line detection (_isCaretOnEdgeLine) is layout-based and not reproducible in jsdom
  // (no Range.getClientRects); we stub it true to exercise the navigation logic it gates.
  it('ArrowDown moves the caret from the bottom of a page to the next page', () => {
    const { page0 } = setupTwoPages();
    (component as any)._isCaretOnEdgeLine = () => true;
    collapseCaretIn(page0.firstChild!, 3);

    const moved = (component as any)._tryMoveCaretAcrossPages('down');

    expect(moved).toBe(true);
    expect(component.activePageIndex()).toBe(1);
  });

  it('ArrowUp moves the caret from the top of a page to the previous page', () => {
    const { page1 } = setupTwoPages();
    (component as any)._isCaretOnEdgeLine = () => true;
    component.activePageIndex.set(1);
    collapseCaretIn(page1.firstChild!, 2);

    const moved = (component as any)._tryMoveCaretAcrossPages('up');

    expect(moved).toBe(true);
    expect(component.activePageIndex()).toBe(0);
  });

  it('does not move when the selection is a range (Shift+Arrow keeps selecting)', () => {
    const { page0 } = setupTwoPages();
    const range = document.createRange();
    range.setStart(page0.firstChild!, 0);
    range.setEnd(page0.firstChild!, 5); // non-collapsed
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);

    const moved = (component as any)._tryMoveCaretAcrossPages('down');

    expect(moved).toBe(false);
    expect(component.activePageIndex()).toBe(0);
  });

  it('does not move past the last page on ArrowDown', () => {
    const { page1 } = setupTwoPages();
    component.activePageIndex.set(1);
    collapseCaretIn(page1.firstChild!, 3);

    const moved = (component as any)._tryMoveCaretAcrossPages('down');

    expect(moved).toBe(false);
  });

  it('does nothing with a single page', () => {
    const page0 = document.createElement('div');
    page0.textContent = 'Jedyna strona';
    document.body.appendChild(page0);
    (component as any).pageEditorRefs = { toArray: () => [{ nativeElement: page0 }] };
    collapseCaretIn(page0.firstChild!, 2);

    expect((component as any)._tryMoveCaretAcrossPages('down')).toBe(false);
  });

  // --- page-number font-size ------------------------------------------------

  it('page number inherits an inline font-size from the footer content', () => {
    const footer = document.createElement('div');
    footer.innerHTML = '<span style="font-size:8pt">Strona </span>';

    const html = (component as any)._pageNumberHtml(footer);

    expect(html).toContain('class="page-number"');
    expect(html).toContain('font-size:8pt');
    expect(html).toContain('{page}');
  });

  it('_inlineFieldFontSize reads the first inline-sized span', () => {
    const footer = document.createElement('div');
    footer.innerHTML = 'Strona <span style="font-size:9pt">x</span>';

    expect((component as any)._inlineFieldFontSize(footer)).toBe('9pt');
  });
});
