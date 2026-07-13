import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Bug UAT — każda komórka wstawionej tabeli zaczynała się od twardej spacji (&nbsp;
 * jako placeholder pustej komórki). Po wpisaniu tekstu spacja zostawała przed nim,
 * przesuwając zawartość względem MS Word i trafiając do zapisanego DOCX.
 *
 * Poprawka jest dwuczęściowa:
 *  - nowe komórki dostają placeholder <br> (eksport: goły <br> w td → pusty akapit),
 *  - onEditorBeforeInput zaznacza samotny placeholder U+00A0 tuż przed wstawieniem
 *    tekstu, więc pisanie go zastępuje — to obejmuje też komórki/akapity z importu
 *    DOCX (<td><p>&nbsp;</p></td>) i dokumenty zapisane przed poprawką.
 */
describe('WysiwygEditorComponent — komórki tabeli bez wiodącej twardej spacji', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;
  let editor: HTMLDivElement;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;

    editor = document.createElement('div');
    document.body.appendChild(editor);
    (component as any).editorContent = { nativeElement: editor };
    (component as any).onContentChange = () => {};
    window.getSelection()?.removeAllRanges();
  });

  afterEach(() => {
    editor.remove();
  });

  function collapseIn(node: Node, offset: number): void {
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
  }

  it('insertTable tworzy komórki z placeholderem <br>, bez \\u00A0', () => {
    component.insertTable('2x2');

    const cells = Array.from(editor.querySelectorAll('td'));
    expect(cells.length).toBe(4);
    for (const td of cells) {
      expect(td.innerHTML).toBe('<br>');
      expect(td.textContent).not.toContain('\u00A0');
    }
  });

  it('beforeinput zaznacza samotny placeholder \\u00A0, żeby tekst go zastąpił', () => {
    // Komórka z importu DOCX: <td><p>&nbsp;</p></td>
    editor.innerHTML = '<table><tbody><tr><td><p>\u00A0</p></td></tr></tbody></table>';
    const textNode = editor.querySelector('p')!.firstChild!;
    collapseIn(textNode, 1); // kliknięcie stawia kursor za placeholderem

    component.onEditorBeforeInput({ inputType: 'insertText' } as InputEvent);

    const sel = window.getSelection()!;
    expect(sel.isCollapsed).toBe(false);
    expect(sel.anchorNode).toBe(textNode);
    expect(sel.anchorOffset).toBe(0);
    expect(sel.focusOffset).toBe(1);
  });

  it('beforeinput nie rusza selekcji, gdy blok ma prawdziwą treść', () => {
    editor.innerHTML = '<table><tbody><tr><td><p>abc</p></td></tr></tbody></table>';
    const textNode = editor.querySelector('p')!.firstChild!;
    collapseIn(textNode, 1);

    component.onEditorBeforeInput({ inputType: 'insertText' } as InputEvent);

    const sel = window.getSelection()!;
    expect(sel.isCollapsed).toBe(true);
    expect(sel.anchorOffset).toBe(1);
  });
});
