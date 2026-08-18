import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/** Built via fromCharCode on purpose: a literal ZWS pasted into source is invisible. */
const ZWS = String.fromCharCode(0x200b);

/**
 * Problem 10 — sticky formatting. Reported bug: pick a font at a collapsed caret, type,
 * delete the typed text — the active font silently resets to the document default, and the
 * next typed characters genuinely use the wrong font. Root cause: the ZWS span created by
 * the picker was the ONLY carrier of the choice; when Backspace emptied it, the browser
 * removed the span and `updateFormattingState` re-derived everything from computed styles.
 *
 * The fix is a pending-inline-style record anchored at the caret:
 *  - pickers arm it alongside the ZWS span,
 *  - beforeinput(delete) captures the effective style when a styled span is about to be
 *    emptied; the following input event re-anchors it at the post-delete caret,
 *  - `updateFormattingState` reports the pending values while the caret sits at the anchor,
 *  - beforeinput(insertText) re-creates the styled span so the typed char lands inside it,
 *  - moving the caret away / blur / undo clears the record.
 *
 * jsdom does not execute contenteditable editing, so the browser mutations (insert/delete)
 * are simulated manually between the beforeinput/input calls — same convention as
 * `wysiwyg-editor.table-cell-placeholder.spec.ts` and `wysiwyg-editor.font-family.spec.ts`.
 */
describe('WysiwygEditorComponent — sticky font po skasowaniu wpisanego tekstu (Problem 10)', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;
  const mounted: HTMLElement[] = [];

  beforeEach(async () => {
    // jsdom implements neither queryCommandState nor focus/selection layout — stub what the
    // formatting read-back touches (same as the font-family spec).
    if (typeof (document as any).queryCommandState !== 'function') {
      (document as any).queryCommandState = () => false;
    }
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  afterEach(() => {
    mounted.splice(0).forEach(el => el.remove());
    window.getSelection()?.removeAllRanges();
  });

  function mountEditor(html: string): HTMLDivElement {
    const editor = document.createElement('div');
    editor.setAttribute('contenteditable', 'true');
    editor.innerHTML = html;
    document.body.appendChild(editor);
    mounted.push(editor);
    (component as any).editorContent = { nativeElement: editor };
    (component as any).onContentChange = () => {};
    return editor;
  }

  function caretIn(node: Node, offset: number): void {
    const sel = window.getSelection()!;
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
  }

  const pending = () => (component as any)._pendingInlineStyle as
    { fontFamily?: string; fontSize?: string } | null;

  it('pełny scenariusz: wybór → pisanie → Backspace → stan raportuje wybrany font → pisanie odtwarza span', () => {
    const editor = mountEditor('<p><br></p>');
    const p = editor.querySelector('p')!;
    caretIn(p, 0);

    // 1. Pick a font at a collapsed caret → ZWS carrier span + armed pending record.
    component.setFontFamily('Arial');
    const span = editor.querySelector('span')!;
    expect(span.style.fontFamily).toContain('Arial');
    expect(pending()?.fontFamily).toBe('Arial');

    // 2. Type a character into the span. beforeinput sees the caret at the pending anchor;
    //    the live ZWS span carries the style, so pending is consumed (cleared).
    component.onEditorBeforeInput({ inputType: 'insertText' } as InputEvent);
    expect(pending()).toBeNull();
    const textNode = span.firstChild as Text;
    textNode.data = ZWS + 'x';
    caretIn(textNode, 2);

    // 3. Backspace: beforeinput captures the effective style of the about-to-be-emptied
    //    span; the browser then drops the span; the input phase re-anchors the capture.
    component.onEditorBeforeInput({ inputType: 'deleteContentBackward' } as InputEvent);
    expect((component as any)._pendingDeleteCapture?.fontFamily).toBe('Arial');
    span.remove();
    caretIn(p, 0);
    (component as any)._consumePendingDeleteCapture();
    expect(pending()?.fontFamily).toBe('Arial');

    // 4. The toolbar read-back still reports the picked font, not the document default.
    (component as any).updateFormattingState();
    expect(component.editorState().currentStyle.fontFamily).toBe('Arial');

    // 5. Typing again re-creates the styled span at the caret BEFORE the browser inserts,
    //    so the typed character lands inside it.
    component.onEditorBeforeInput({ inputType: 'insertText' } as InputEvent);
    const newSpan = editor.querySelector('span');
    expect(newSpan).not.toBeNull();
    expect(newSpan!.style.fontFamily).toContain('Arial');
    const sel = window.getSelection()!;
    expect(newSpan!.contains(sel.getRangeAt(0).startContainer)).toBe(true);
    expect(pending()).toBeNull();
  });

  it('setFontSize (zwinięta karetka) uzbraja pending i read-path raportuje wybrany rozmiar', () => {
    const editor = mountEditor('<p><br></p>');
    caretIn(editor.querySelector('p')!, 0);

    component.setFontSize(14);

    expect(pending()?.fontSize).toBe('14pt');
    // jsdom would derive 11pt from computed styles here — 14 proves the pending read path.
    (component as any).updateFormattingState();
    expect(component.editorState().currentStyle.fontSize).toBe(14);
  });

  it('Backspace w spanie z większą ilością tekstu NIE uzbraja przechwycenia (tani warunek)', () => {
    const editor = mountEditor('<p><span style="font-family: Arial">abc</span></p>');
    const text = editor.querySelector('span')!.firstChild as Text;
    caretIn(text, 3);

    component.onEditorBeforeInput({ inputType: 'deleteContentBackward' } as InputEvent);

    expect((component as any)._pendingDeleteCapture).toBeNull();
  });

  it('przeniesienie karetki gdzie indziej czyści pending (selectionchange) — stan wraca do DOM', () => {
    const editor = mountEditor('<p>abc</p><p>def</p>');
    const paragraphs = Array.from(editor.querySelectorAll('p'));
    caretIn(paragraphs[0].firstChild!, 0);
    (component as any)._pendingInlineStyle = { fontFamily: 'Arial' };
    (component as any)._pendingStyleAnchor = { node: paragraphs[0].firstChild!, offset: 0 };

    // User clicks/arrows into the second paragraph.
    caretIn(paragraphs[1].firstChild!, 1);
    (component as any).onSelectionChange();

    expect(pending()).toBeNull();
    (component as any).updateFormattingState();
    expect(component.editorState().currentStyle.fontFamily).not.toBe('Arial');
  });

  it('selectionchange na karetce RÓWNEJ kotwicy nie czyści pending', () => {
    const editor = mountEditor('<p>abc</p>');
    const text = editor.querySelector('p')!.firstChild!;
    caretIn(text, 1);
    (component as any)._pendingInlineStyle = { fontFamily: 'Arial' };
    (component as any)._pendingStyleAnchor = { node: text, offset: 1 };

    (component as any).onSelectionChange();

    expect(pending()?.fontFamily).toBe('Arial');
  });

  it('undo czyści pending (przywrócony DOM unieważnia kotwicę)', () => {
    mountEditor('<p>abc</p>');
    (component as any)._pendingInlineStyle = { fontFamily: 'Arial' };

    component.undo();

    expect(pending()).toBeNull();
  });
});
