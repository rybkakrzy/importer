import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Regression coverage for "Wklej bez formatowania" (paste plain text).
 *
 * The reported bug: pasting via the menu kept the SOURCE formatting and landed at
 * the SOURCE position instead of the caret. Both symptoms shared one root cause —
 * the async menu path inserted at a stale/live selection (the source range), and
 * inserting into the source run made the text inherit the source style.
 *
 * pastePlainTextAt() fixes this by inserting a bare text node at a bookmark captured
 * at command-invocation time. These tests exercise it against real jsdom DOM Ranges,
 * so they assert the MODEL (target range, replaced content, marks, caret, one undo),
 * not just the visual DOM.
 */
describe('WysiwygEditorComponent — pastePlainTextAt (Wklej bez formatowania)', () => {
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
    editor.setAttribute('contenteditable', 'true');
    document.body.appendChild(editor);
    (component as any).getActiveEditor = () => editor;
    // isolate from repagination/emit; count undo snapshots via the spy below.
    (component as any).onContentChange = vi.fn();
    (component as any).isSelectionInEditor = (sel: Selection) =>
      !!sel.anchorNode && editor.contains(sel.anchorNode);
  });

  afterEach(() => {
    editor.remove();
    window.getSelection()?.removeAllRanges();
  });

  /** Places the caret inside `node` at `offset` and returns the collapsed range. */
  function caretIn(node: Node, offset: number): Range {
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
    return range;
  }

  it('inserts plain text at the captured TARGET bookmark, not the live selection', () => {
    // Source (formatted) in paragraph 1, empty target paragraph 2.
    editor.innerHTML =
      '<p><b style="color:red">SOURCE</b></p><p id="target"><br></p>';
    const target = editor.querySelector('#target')!;

    // Bookmark captured earlier points at the target paragraph.
    const bookmark = document.createRange();
    bookmark.setStart(target, 0);
    bookmark.collapse(true);

    // Meanwhile the LIVE selection is back inside the formatted source run.
    const sourceText = editor.querySelector('b')!.firstChild!;
    caretIn(sourceText, 6);

    component.pastePlainTextAt(bookmark, 'plain');

    // Text landed in the target, and NOT inside the source <b>.
    expect(target.textContent).toContain('plain');
    expect(editor.querySelector('b')!.textContent).toBe('SOURCE');
  });

  it('drops source formatting: inserted run has no bold/color/font marks', () => {
    editor.innerHTML = '<p id="target"><br></p>';
    const target = editor.querySelector('#target')!;
    const bookmark = document.createRange();
    bookmark.setStart(target, 0);
    bookmark.collapse(true);

    component.pastePlainTextAt(bookmark, 'hello');

    // The pasted text is a bare text node with no formatting wrapper.
    const inserted = [...target.childNodes].find(
      (n) => n.nodeType === Node.TEXT_NODE && n.textContent === 'hello',
    );
    expect(inserted).toBeTruthy();
    expect(target.querySelector('b, strong, span[style], font')).toBeNull();
  });

  it('escapes an inline formatting run at the caret (drops inherited color/bold)', () => {
    // Reproduces the reported bug: after copying colored text the source selection is
    // still the paste target, so the caret sits inside the colored/bold run. The pasted
    // plain text must NOT inherit that run's formatting.
    editor.innerHTML =
      '<p id="target">a<b style="color:red"><span style="color:red">RUN</span></b>b</p>';
    const runText = editor.querySelector('span')!.firstChild!; // "RUN"
    const bookmark = document.createRange();
    bookmark.setStart(runText, 1); // between R|UN, deep inside <b><span style=color>
    bookmark.collapse(true);

    component.pastePlainTextAt(bookmark, 'plain');

    // The inserted text node must not have any b/span[style]/font ancestor.
    const inserted = [...editor.querySelectorAll('#target')][0]!;
    const walker = document.createTreeWalker(inserted, NodeFilter.SHOW_TEXT);
    let plainNode: Node | null = null;
    while (walker.nextNode()) {
      if (walker.currentNode.textContent === 'plain') { plainNode = walker.currentNode; break; }
    }
    expect(plainNode).toBeTruthy();
    let ancestor = plainNode!.parentElement;
    while (ancestor && ancestor.id !== 'target') {
      expect(['B', 'STRONG', 'FONT'].includes(ancestor.tagName)).toBe(false);
      expect(ancestor.getAttribute('style')).toBeNull();
      ancestor = ancestor.parentElement;
    }
    // Original run text survives, split around the insertion.
    expect(editor.textContent).toBe('aRplainUNb');
  });

  it('escapes the run even when the plain paste REPLACES the whole colored run', () => {
    // Select the colored run and paste-plain over it: the emptied wrapper must not
    // swallow the new text.
    editor.innerHTML = '<p id="target"><span style="color:red">RED</span></p>';
    const runText = editor.querySelector('span')!.firstChild!;
    const bookmark = document.createRange();
    bookmark.setStart(runText, 0);
    bookmark.setEnd(runText, 3); // whole "RED"

    component.pastePlainTextAt(bookmark, 'plain');

    expect(editor.textContent).toBe('plain');
    // No styled span should wrap the pasted text.
    const target = editor.querySelector('#target')!;
    const styledSpan = target.querySelector('span[style]');
    expect(styledSpan?.textContent ?? '').not.toContain('plain');
  });

  it('replaces a non-empty selection', () => {
    editor.innerHTML = '<p id="target">keepXXXkeep</p>';
    const target = editor.querySelector('#target')!;
    const textNode = target.firstChild!;
    const bookmark = document.createRange();
    bookmark.setStart(textNode, 4); // before XXX
    bookmark.setEnd(textNode, 7); // after XXX

    component.pastePlainTextAt(bookmark, 'YES');

    expect(target.textContent).toBe('keepYESkeep');
  });

  it('leaves the caret directly after the inserted text', () => {
    editor.innerHTML = '<p id="target"><br></p>';
    const target = editor.querySelector('#target')!;
    const bookmark = document.createRange();
    bookmark.setStart(target, 0);
    bookmark.collapse(true);

    component.pastePlainTextAt(bookmark, 'abc');

    // Assert on the caret the component computed (savedSelection is a direct clone of it).
    // jsdom's global Selection does not faithfully round-trip a contenteditable caret, so we
    // check the deterministic bookmark instead of window.getSelection().
    const caret: Range = (component as any).savedSelection;
    expect(caret).toBeTruthy();
    expect(caret.collapsed).toBe(true);
    const after = document.createRange();
    after.setStartAfter(target.firstChild!); // the inserted 'abc' text node
    after.collapse(true);
    expect(caret.compareBoundaryPoints(Range.START_TO_START, after)).toBe(0);
  });

  it('multi-line text becomes soft <br> breaks within the paragraph', () => {
    editor.innerHTML = '<p id="target"><br></p>';
    const target = editor.querySelector('#target')!;
    const bookmark = document.createRange();
    bookmark.setStart(target, 0);
    bookmark.collapse(true);

    component.pastePlainTextAt(bookmark, 'line1\nline2');

    expect(target.querySelectorAll('br').length).toBeGreaterThanOrEqual(1);
    expect(target.textContent).toContain('line1');
    expect(target.textContent).toContain('line2');
  });

  it('keeps HTML-looking text literal (no injection)', () => {
    editor.innerHTML = '<p id="target"><br></p>';
    const target = editor.querySelector('#target')!;
    const bookmark = document.createRange();
    bookmark.setStart(target, 0);
    bookmark.collapse(true);

    component.pastePlainTextAt(bookmark, '<img src=x onerror=alert(1)>');

    expect(target.querySelector('img')).toBeNull();
    expect(target.textContent).toContain('<img src=x onerror=alert(1)>');
  });

  it('produces exactly one content change (one undo entry)', () => {
    editor.innerHTML = '<p id="target"><br></p>';
    const target = editor.querySelector('#target')!;
    const bookmark = document.createRange();
    bookmark.setStart(target, 0);
    bookmark.collapse(true);

    component.pastePlainTextAt(bookmark, 'once');

    expect((component as any).onContentChange).toHaveBeenCalledTimes(1);
  });

  it('aborts safely when the bookmark no longer lives in the editor', () => {
    editor.innerHTML = '<p id="target">unchanged</p>';
    // Bookmark points into a detached node (async doc swap / unmount).
    const orphan = document.createElement('p');
    orphan.textContent = 'gone';
    const bookmark = document.createRange();
    bookmark.setStart(orphan.firstChild!, 0);
    bookmark.collapse(true);

    expect(() => component.pastePlainTextAt(bookmark, 'x')).not.toThrow();
    expect(editor.textContent).toBe('unchanged');
    expect((component as any).onContentChange).not.toHaveBeenCalled();
  });

  it('is a no-op for empty/whitespace-only clipboard text', () => {
    editor.innerHTML = '<p id="target">stay</p>';
    const target = editor.querySelector('#target')!;
    const bookmark = document.createRange();
    bookmark.setStart(target.firstChild!, 0);
    bookmark.collapse(true);

    component.pastePlainTextAt(bookmark, '   \n  ');

    expect(target.textContent).toBe('stay');
    expect((component as any).onContentChange).not.toHaveBeenCalled();
  });

  it('captureSelectionBookmark clones the live editor selection', () => {
    editor.innerHTML = '<p id="target">abc</p>';
    const textNode = editor.querySelector('#target')!.firstChild!;
    const live = caretIn(textNode, 1);

    const bookmark = component.captureSelectionBookmark();
    expect(bookmark).not.toBeNull();
    expect(bookmark!.startContainer).toBe(textNode);
    expect(bookmark!.startOffset).toBe(1);
    // A clone: mutating the live selection must not move the bookmark.
    live.setStart(textNode, 3);
    expect(bookmark!.startOffset).toBe(1);
  });
});
