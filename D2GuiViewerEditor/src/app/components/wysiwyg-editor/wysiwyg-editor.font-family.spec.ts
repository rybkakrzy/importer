import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Font-family for the CURRENT caret and for a selection.
 *
 * Reported bug: after picking a non-default font ("ING Me" is the corporate default), newly
 * typed text kept the default. The document flow is: pick a font at a collapsed caret →
 * `setFontFamily` drops a zero-width-space (ZWS) span carrying the chosen `font-family` and
 * puts the caret inside it, so the next typed characters inherit the font. These tests assert
 * the model that makes typing inherit the choice — the span, the caret anchor, and the saved
 * selection — rather than the pixels (jsdom does not lay out contenteditable typing).
 *
 * Regression guard for the parity fix: the fresh-span branch must update `savedSelection` and
 * refresh the toolbar state, exactly like the sibling `setFontSize` does — otherwise a
 * follow-up combobox pick (which restores `savedSelection` after focus leaves the editor)
 * would land BEFORE the span and the second font would be lost.
 */
describe('WysiwygEditorComponent — font-family at caret and selection', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;
  const mounted: HTMLElement[] = [];

  beforeEach(async () => {
    // jsdom implements neither queryCommandState nor focus/selection layout — stub what the
    // formatting read-back touches so setFontFamily's state refresh survives.
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
    // Both getActiveEditor() and isSelectionInEditor() resolve through editorContent.
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

  function fontFamilyOf(el: Element | null): string {
    return (el as HTMLElement | null)?.style.fontFamily ?? '';
  }

  it('caret without selection: wraps a ZWS span in the chosen font and anchors the caret inside it', () => {
    const editor = mountEditor('<p><br></p>');
    caretIn(editor.querySelector('p')!, 0);

    component.setFontFamily('Arial');

    const span = editor.querySelector('span');
    expect(span).not.toBeNull();
    expect(fontFamilyOf(span)).toContain('Arial');
    expect(span!.textContent).toBe('​');

    // Caret sits inside the span (offset 1, right after the ZWS) so typed text inherits Arial.
    const sel = window.getSelection()!;
    expect(sel.rangeCount).toBe(1);
    expect(span!.contains(sel.getRangeAt(0).startContainer)).toBe(true);
  });

  it('caret path keeps savedSelection + toolbar state in sync (parity with setFontSize)', () => {
    const editor = mountEditor('<p><br></p>');
    caretIn(editor.querySelector('p')!, 0);
    const stateSpy = vi.spyOn(component as any, 'updateFormattingState');

    component.setFontFamily('Verdana');

    const span = editor.querySelector('span')!;
    // savedSelection must point INTO the fresh span (so a later restore lands inside it).
    const saved = (component as any).savedSelection as Range | null;
    expect(saved).not.toBeNull();
    expect(span.contains(saved!.startContainer)).toBe(true);
    expect(stateSpy).toHaveBeenCalled();
  });

  it('second pick at the same caret reuses the ZWS span instead of nesting a new one', () => {
    const editor = mountEditor('<p><br></p>');
    caretIn(editor.querySelector('p')!, 0);

    component.setFontFamily('Arial');
    component.setFontFamily('Times New Roman');

    const spans = editor.querySelectorAll('span');
    expect(spans.length).toBe(1);
    expect(fontFamilyOf(spans[0])).toContain('Times New Roman');
    expect(spans[0].textContent).toBe('​');
  });

  // The wrap logic is exercised directly: jsdom's editor.focus() collapses a live selection
  // (a jsdom quirk absent in real browsers), so the public setFontFamily selection branch is
  // covered by the headless-Chrome scenarios in .ai/EDITOR_KEYBOARD.md instead.
  it('selection: wraps the selected text in the chosen font (existing text kept)', () => {
    const editor = mountEditor('<p>alfa beta</p>');
    const text = editor.querySelector('p')!.firstChild!;
    const sel = window.getSelection()!;
    const range = document.createRange();
    range.setStart(text, 0);
    range.setEnd(text, 4); // "alfa"
    sel.removeAllRanges();
    sel.addRange(range);

    (component as any).applyFontFamilyToSelection('Georgia', sel, range);

    const span = Array.from(editor.querySelectorAll('span'))
      .find(s => (s.style.fontFamily || '').includes('Georgia'));
    expect(span).toBeTruthy();
    expect(span!.textContent).toBe('alfa');
    expect(editor.textContent).toBe('alfa beta');
  });

  it('does nothing without an active editor (guarded)', () => {
    (component as any).editorContent = null;
    expect(() => component.setFontFamily('Arial')).not.toThrow();
  });
});
