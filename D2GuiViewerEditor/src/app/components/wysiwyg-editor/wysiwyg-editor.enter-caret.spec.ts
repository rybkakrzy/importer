import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Issue 1 — after ENTER the caret must stay in the NEW (empty) paragraph, not jump back to the
 * previous line. The editor repaginates ~600ms after input, rebuilding the page DOM; the caret
 * is preserved via a {block, offset} anchor. A pure global text-offset anchor was ambiguous at
 * block boundaries (an empty new paragraph has 0 chars → restore landed at the end of the
 * previous paragraph). These tests pin the block-aware save/restore that fixes it.
 */
describe('WysiwygEditorComponent — caret survives repagination across block boundaries', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  function pageWith(html: string): HTMLDivElement {
    const editor = document.createElement('div');
    editor.innerHTML = html;
    document.body.appendChild(editor);
    (component as any).pageEditorRefs = { toArray: () => [{ nativeElement: editor }] };
    return editor;
  }

  function collapseIn(node: Node, offset: number): void {
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
  }

  it('anchors the caret to the empty new paragraph (block 1), not the previous one', () => {
    const editor = pageWith('<p>Hello</p><p></p>');
    const emptyPara = editor.querySelectorAll('p')[1];
    collapseIn(emptyPara, 0);

    const anchor = (component as any)._saveGlobalCaret([{ nativeElement: editor }]);

    expect(anchor).toEqual({ block: 1, offset: 0 });
  });

  it('restores the caret into the empty new paragraph (not end of previous line)', () => {
    const editor = pageWith('<p>Hello</p><p></p>');
    const emptyPara = editor.querySelectorAll('p')[1];

    (component as any)._restoreGlobalCaret({ block: 1, offset: 0 });

    const sel = window.getSelection()!;
    expect(emptyPara.contains(sel.anchorNode) || sel.anchorNode === emptyPara).toBe(true);
  });

  it('still restores a mid-paragraph caret to the right block + offset', () => {
    const editor = pageWith('<p>Hello</p><p>World</p>');
    const first = editor.querySelectorAll('p')[0];
    collapseIn(first.firstChild!, 3); // caret after "Hel"

    const anchor = (component as any)._saveGlobalCaret([{ nativeElement: editor }]);
    expect(anchor).toEqual({ block: 0, offset: 3 });

    (component as any)._restoreGlobalCaret(anchor);
    const sel = window.getSelection()!;
    expect(first.contains(sel.anchorNode)).toBe(true);
    expect(sel.anchorOffset).toBe(3);
  });
});
