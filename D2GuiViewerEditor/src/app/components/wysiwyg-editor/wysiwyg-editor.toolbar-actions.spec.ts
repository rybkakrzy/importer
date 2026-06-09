import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Toolbar action guards that protect the document from invalid input:
 *  - font-size: a non-finite/out-of-range value must NOT reach the DOM as `0pt`/`NaNpt`
 *    (that renders text invisible — the reported "tekst znika").
 *  - insert link: URLs are normalised and labels are HTML-escaped.
 *
 * The DOM-mutating parts (execCommand createLink/insertHTML, applyFontSizeToSelection) rely on
 * contenteditable + Selection, which jsdom does not implement; those are covered by the manual
 * scenarios in .ai/EDITOR_KEYBOARD.md. Here we assert the deterministic guard logic.
 */
describe('WysiwygEditorComponent — toolbar action guards', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
    // setFontSize/insertLink need a non-null active editor; the guards run before any DOM work.
    (component as any).getActiveEditor = () => document.createElement('div');
  });

  // --- Issue 2: font-size --------------------------------------------------

  it('rejects font-size 0 / NaN / out-of-range (text must not get 0pt)', () => {
    (component as any).currentFontSize = 11;

    component.setFontSize(0);
    expect((component as any).currentFontSize).toBe(11);

    component.setFontSize(NaN);
    expect((component as any).currentFontSize).toBe(11);

    component.setFontSize(500);
    expect((component as any).currentFontSize).toBe(11);

    component.setFontSize(-3);
    expect((component as any).currentFontSize).toBe(11);
  });

  it('accepts a valid font-size', () => {
    (component as any).currentFontSize = 11;
    component.setFontSize(14);
    expect((component as any).currentFontSize).toBe(14);
  });

  // --- Issue 6: link URL normalisation + escaping --------------------------

  it('normalises link URLs', () => {
    const n = (u: string) => (component as any).normalizeLinkUrl(u);
    expect(n('')).toBeNull();
    expect(n('   ')).toBeNull();
    expect(n('https://x.pl')).toBe('https://x.pl');
    expect(n('http://x.pl')).toBe('http://x.pl');
    expect(n('mailto:a@b.pl')).toBe('mailto:a@b.pl');
    expect(n('#sekcja')).toBe('#sekcja');
    expect(n('/lokalna/sciezka')).toBe('/lokalna/sciezka');
    expect(n('www.example.com')).toBe('https://www.example.com');
    expect(n('example.com/path?x=1')).toBe('https://example.com/path?x=1');
  });

  it('escapes link label HTML', () => {
    expect((component as any).escapeHtml('a<b>&"')).toBe('a&lt;b&gt;&amp;&quot;');
  });

  it('insertLink with empty URL is a no-op (does not throw)', () => {
    expect(() => component.insertLink('')).not.toThrow();
  });

  // --- Font-family: restore the lost caret BEFORE focus --------------------

  it('setFontFamily restores the saved selection when focus left the editor (font <select> click)', () => {
    const editor = document.createElement('div');
    (component as any).getActiveEditor = () => editor;
    (component as any).onContentChange = () => {}; // isolate from repaginate/emit
    (component as any).savedSelection = document.createRange();
    const spy = vi.spyOn(component as any, 'restoreSelection').mockImplementation(() => true);

    // Picking a font from the native <select> blurs the contenteditable and clears the
    // selection. setFontFamily must restore the saved caret BEFORE focus, otherwise the
    // font span lands at the document start and newly typed text keeps the default font.
    window.getSelection()?.removeAllRanges();
    component.setFontFamily('Arial');

    expect(spy).toHaveBeenCalled();
    expect((component as any).currentFontFamily).toBe('Arial');
  });

  // --- Issue 4: plain paste restores the lost editor selection -------------

  it('insertText restores the saved selection when focus left the editor (menu paste)', () => {
    const editor = document.createElement('div');
    (component as any).getActiveEditor = () => editor;
    (component as any).onContentChange = () => {}; // isolate from repaginate/emit
    (document as any).execCommand = vi.fn(); // jsdom has no execCommand
    (component as any).savedSelection = document.createRange();
    const spy = vi.spyOn(component as any, 'restoreSelection').mockImplementation(() => true);

    // No live selection in the editor (as after a menu click) → must restore the saved one.
    window.getSelection()?.removeAllRanges();
    component.insertText('plain text');

    expect(spy).toHaveBeenCalled();
  });
});
