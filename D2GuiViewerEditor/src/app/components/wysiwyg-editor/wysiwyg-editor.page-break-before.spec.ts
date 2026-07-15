import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Właściwość „podział strony przed" (w:pageBreakBefore — checkbox dialogu akapitu Worda,
 * style nagłówków rozdziałów). Reader emituje ją jako `page-break-before:always` w stylu
 * INLINE akapitu (nie marker div.page-break — ten reprezentuje ręczny w:br type=page).
 * Paginacja musi łamać stronę PRZED takim blokiem, a blok jedzie na nową stronę razem
 * ze swoim stylem (round-trip: writer odtwarza w:pageBreakBefore w pPr).
 */
describe('WysiwygEditorComponent — właściwość page-break-before łamie stronę', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;
  const bodyEditors: HTMLElement[] = [];

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  afterEach(() => {
    bodyEditors.splice(0).forEach(el => el.remove());
  });

  function pageWith(html: string): HTMLDivElement {
    const editor = document.createElement('div');
    editor.innerHTML = html;
    document.body.appendChild(editor);
    bodyEditors.push(editor);
    (component as any).pageEditorRefs = { toArray: () => [{ nativeElement: editor }] };
    return editor;
  }

  /** Stub pomiaru: stała wysokość na blok (jsdom nie ma layoutu). */
  function stubBlockHeights(heightPerBlock: number): void {
    (component as any)._measureBlockRunHeights = (_m: HTMLElement, blocks: HTMLElement[]) =>
      blocks.map(() => heightPerBlock);
  }

  it('blok z page-break-before:always zaczyna nową stronę (mimo że treść mieści się na jednej)', () => {
    pageWith('<p>a</p><p style="page-break-before:always;">b</p><p>c</p>');
    stubBlockHeights(20); // wszystko zmieściłoby się na stronie 1

    (component as any)._repaginateNow();

    const pages = component.pageContents();
    expect(pages.length).toBe(2);
    expect(pages[0]).toContain('<p>a</p>');
    expect(pages[0]).not.toContain('>b<');
    // Blok jedzie na nową stronę RAZEM ze swoim stylem (żaden marker nie jest dokładany).
    expect(pages[1]).toContain('page-break-before');
    expect(pages[1]).toContain('>b<');
    expect(pages[1]).toContain('<p>c</p>');
    expect(pages.join('')).not.toContain('class="page-break"');
  });

  it('właściwość na PIERWSZYM bloku dokumentu nie tworzy pustej strony (jak w Wordzie)', () => {
    pageWith('<p style="page-break-before:always;">a</p><p>b</p>');
    stubBlockHeights(20);

    (component as any)._repaginateNow();

    expect(component.pageContents().length).toBe(1);
  });

  it('nie gubi ani nie duplikuje treści przy łamaniu właściwością', () => {
    pageWith('<p>1</p><p style="page-break-before:always;">2</p><p>3</p><p style="page-break-before:always;">4</p>');
    stubBlockHeights(20);

    (component as any)._repaginateNow();

    const pages = component.pageContents();
    expect(pages.length).toBe(3);
    const text = pages.join('').replace(/<[^>]+>/g, '');
    expect(text).toBe('1234');
  });

  it('modern break-before:page też łamie; break-before:column (podział kolumny) NIE', () => {
    const has = (style: string) => {
      const el = document.createElement('p');
      el.setAttribute('style', style);
      return (component as any)._hasPageBreakBeforeStyle(el);
    };

    expect(has('break-before:page;')).toBe(true);
    expect(has('page-break-before: always;')).toBe(true);
    expect(has('page-break-before:auto;')).toBe(false);
    expect(has('break-before:column;')).toBe(false);
    expect(has('')).toBe(false);
  });
});
