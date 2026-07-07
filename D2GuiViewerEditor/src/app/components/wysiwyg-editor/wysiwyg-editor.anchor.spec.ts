import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Znacznik kotwicy (Word-like) + stany pola tekstowego:
 *  - zaznaczenie PŁYWAJĄCEGO elementu (obraz front/behind, textbox) pokazuje ikonę kotwicy
 *    przy akapicie-kotwicy — jako overlay w .page (POZA contenteditable, nie serializuje się),
 *  - zaznaczenie elementu inline kotwicy nie pokazuje,
 *  - zmiana zaznaczenia przenosi znacznik, usunięcie zaznaczenia / elementu go chowa,
 *  - klasy stanu edycyjnego textboxa (tb-selected/tb-dragging/tb-edge) nigdy nie trafiają
 *    do zapisu, a klasa docx-textbox i data-* kotwicy zostają (kontrakt writera DOCX).
 */
describe('WysiwygEditorComponent — znacznik kotwicy i stany pola tekstowego', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  /** Strona jak w szablonie edytora: .page > .editor-content z treścią. */
  function buildPage(innerHtml: string): { page: HTMLElement; editor: HTMLElement } {
    const page = document.createElement('div');
    page.className = 'page';
    const editor = document.createElement('div');
    editor.className = 'editor-content';
    editor.innerHTML = innerHtml;
    page.appendChild(editor);
    document.body.appendChild(page);
    return { page, editor };
  }

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('zaznaczenie pływającego obrazu pokazuje znacznik kotwicy w .page z aria-label', () => {
    const { page, editor } = buildPage(
      '<p id="kotwica">tekst<span class="editor-image-wrapper" data-pos-mode="front"><img/></span></p>'
    );
    const wrapper = editor.querySelector('.editor-image-wrapper') as HTMLElement;

    (component as any).selectImageWrapper(wrapper);

    const badge = page.querySelector('.anchor-badge') as HTMLElement;
    expect(badge).toBeTruthy();
    expect(badge.getAttribute('aria-label')).toContain('Kotwica');
    expect(badge.getAttribute('role')).toBe('img');
    // Overlay poza treścią: znacznik NIE jest w contenteditable.
    expect(editor.contains(badge)).toBe(false);
    expect(badge.querySelector('svg')).toBeTruthy();
  });

  it('zaznaczenie obrazu INLINE nie pokazuje kotwicy', () => {
    const { page, editor } = buildPage(
      '<p>tekst<span class="editor-image-wrapper"><img/></span></p>'
    );
    const wrapper = editor.querySelector('.editor-image-wrapper') as HTMLElement;

    (component as any).selectImageWrapper(wrapper);

    expect(page.querySelector('.anchor-badge')).toBeNull();
  });

  it('usunięcie zaznaczenia obrazu chowa znacznik', () => {
    const { page, editor } = buildPage(
      '<p>x<span class="editor-image-wrapper" data-pos-mode="behind"><img/></span></p>'
    );
    const wrapper = editor.querySelector('.editor-image-wrapper') as HTMLElement;
    (component as any).selectImageWrapper(wrapper);
    expect(page.querySelector('.anchor-badge')).toBeTruthy();

    (component as any).clearSelectedImage();

    expect(page.querySelector('.anchor-badge')).toBeNull();
  });

  it('zaznaczenie pływającego textboxa: klasa tb-selected + kotwica przy NASTĘPNYM akapicie', () => {
    const { page, editor } = buildPage(
      '<p>przed</p>'
      + '<div class="docx-textbox" data-textbox="1" data-pos-mode="front"'
      + ' style="position:absolute;left:96px;top:48px;"><p>pole</p></div>'
      + '<p id="kotwica">akapit kotwica</p>'
    );
    const textbox = editor.querySelector('.docx-textbox') as HTMLElement;

    (component as any).selectTextBox(textbox);

    expect(textbox.classList.contains('tb-selected')).toBe(true);
    expect(page.querySelector('.anchor-badge')).toBeTruthy();
  });

  it('przełączenie zaznaczenia na inny element przenosi znacznik (jest tylko jeden)', () => {
    const { page, editor } = buildPage(
      '<p>a<span class="editor-image-wrapper" data-pos-mode="front"><img/></span></p>'
      + '<div class="docx-textbox" data-pos-mode="front" style="position:absolute;"><p>pole</p></div>'
      + '<p>b</p>'
    );
    const wrapper = editor.querySelector('.editor-image-wrapper') as HTMLElement;
    const textbox = editor.querySelector('.docx-textbox') as HTMLElement;

    (component as any).selectImageWrapper(wrapper);
    (component as any).selectTextBox(textbox);

    expect(page.querySelectorAll('.anchor-badge').length).toBe(1);
    // Obraz stracił zaznaczenie przy przełączeniu na textbox.
    expect(wrapper.classList.contains('selected')).toBe(false);

    (component as any).clearSelectedTextBox();
    expect(page.querySelector('.anchor-badge')).toBeNull();
    expect(textbox.classList.contains('tb-selected')).toBe(false);
  });

  it('usunięcie zakotwiczonego elementu z DOM chowa znacznik przy odświeżeniu', () => {
    const { page, editor } = buildPage(
      '<div class="docx-textbox" data-pos-mode="front" style="position:absolute;"><p>pole</p></div>'
      + '<p>kotwica</p>'
    );
    const textbox = editor.querySelector('.docx-textbox') as HTMLElement;
    (component as any).selectTextBox(textbox);
    expect(page.querySelector('.anchor-badge')).toBeTruthy();

    textbox.remove();
    (component as any).refreshAnchorBadge();

    expect(page.querySelector('.anchor-badge')).toBeNull();
  });

  it('serializacja usuwa klasy stanu edycyjnego, zachowuje kontrakt writera (docx-textbox + data-*)', () => {
    const editor = document.createElement('div') as HTMLDivElement;
    editor.innerHTML =
      '<div class="docx-textbox tb-selected tb-edge tb-dragging" data-textbox="1"'
      + ' data-pos-mode="front" data-x-emu="914400" data-y-emu="457200"'
      + ' style="position:absolute;left:96px;top:48px;"><p>pole</p></div>'
      + '<p>kotwica</p>';

    const html = (component as any)._serializeSingleEditor(editor) as string;

    expect(html).not.toContain('tb-selected');
    expect(html).not.toContain('tb-edge');
    expect(html).not.toContain('tb-dragging');
    expect(html).toContain('docx-textbox');
    expect(html).toContain('data-x-emu="914400"');
    expect(html).toContain('data-y-emu="457200"');
    expect(html).toContain('data-pos-mode="front"');
  });

  it('znacznik kotwicy nigdy nie trafia do serializowanej treści', () => {
    const editor = document.createElement('div') as HTMLDivElement;
    editor.innerHTML = '<p>tekst</p><div class="anchor-badge" role="img"><svg></svg></div>';

    const html = (component as any)._serializeSingleEditor(editor) as string;

    expect(html).not.toContain('anchor-badge');
    expect(html).toContain('<p>tekst</p>');
  });

  it('drag textboxa aktualizuje offsety EMU i zachowuje kotwicę (pozycję w DOM)', () => {
    const { editor } = buildPage(
      '<div class="docx-textbox" data-textbox="1" data-pos-mode="front"'
      + ' style="position:absolute;left:100px;top:50px;"><p>pole</p></div>'
      + '<p id="kotwica">akapit</p>'
    );
    const textbox = editor.querySelector('.docx-textbox') as HTMLElement;
    (component as any).onContentChange = () => {};

    (component as any).startTextBoxDrag(
      new MouseEvent('mousedown', { clientX: 200, clientY: 200 }), textbox);
    document.dispatchEvent(new MouseEvent('mousemove', { clientX: 230, clientY: 220 }));
    document.dispatchEvent(new MouseEvent('mouseup', {}));

    // jsdom: brak transformu → skala 1, delta 1:1.
    expect(textbox.style.left).toBe('130px');
    expect(textbox.style.top).toBe('70px');
    expect(textbox.getAttribute('data-x-emu')).toBe(String(130 * 9525));
    expect(textbox.getAttribute('data-y-emu')).toBe(String(70 * 9525));
    // Kotwica bez zmian: div nadal bezpośrednio przed swoim akapitem.
    expect(textbox.nextElementSibling?.id).toBe('kotwica');
    expect(textbox.classList.contains('tb-dragging')).toBe(false);
  });
});
