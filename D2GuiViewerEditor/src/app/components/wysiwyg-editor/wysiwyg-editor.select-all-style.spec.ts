import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Regression coverage for the "Ctrl+A + change color prepends a phantom page" bug.
 *
 * Root cause: the inline-style appliers (color / font family / font size) used
 * Range.extractContents() + insertNode(). That splits the boundary blocks (leaving
 * empty <p> shells at both ends of the selection) and — for the cross-page range
 * produced by Ctrl+A → selectAllContent() (first .editor-content → last one) — it
 * ripped the page scaffolding apart and re-inserted its clones as content, doubling
 * the page count with an empty first page on every use.
 *
 * The fix (_applyInlineStyleToRange) wraps the intersecting TEXT NODES in place:
 * structure is never extracted, page chrome between editors is never touched.
 *
 * Like wysiwyg-editor.font-family.spec.ts, the appliers are exercised DIRECTLY:
 * jsdom's editor.focus() collapses a live selection (a jsdom quirk absent in real
 * browsers), so the public setTextColor path is covered by the headless-Chrome
 * probe (cdp-ctrl-a-color2.mjs) instead.
 */
describe('WysiwygEditorComponent — style na zaznaczeniu (in-place, bez extractContents)', () => {
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
    editor.className = 'editor-content';
    editor.setAttribute('contenteditable', 'true');
    document.body.appendChild(editor);
    (component as any).getActiveEditor = () => editor;
    (component as any).onContentChange = vi.fn();
    (component as any).updateFormattingState = vi.fn();
  });

  afterEach(() => {
    editor.remove();
    document.querySelectorAll('.pages-container').forEach(el => el.remove());
    window.getSelection()?.removeAllRanges();
  });

  function selectRange(setup: (r: Range) => void): Range {
    const range = document.createRange();
    setup(range);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
    return range;
  }

  function applyColor(color: string): void {
    const sel = window.getSelection()!;
    (component as any).applyColorToSelection(color, sel, sel.getRangeAt(0));
  }

  it('zaznaczenie na poziomie kontenera: bez pustych skorup <p>, tekst w spanach z kolorem', () => {
    editor.innerHTML = '<p>Pierwszy</p><p>Drugi</p><ul><li>punkt</li></ul>';
    selectRange(r => {
      r.setStart(editor, 0);
      r.setEnd(editor, editor.childNodes.length);
    });

    applyColor('#ff0000');

    const blocks = Array.from(editor.children).map(c => c.tagName);
    expect(blocks).toEqual(['P', 'P', 'UL']); // no empty shells prepended/appended
    expect(editor.querySelectorAll('p:empty').length).toBe(0);
    const spans = Array.from(editor.querySelectorAll('span'));
    expect(spans.length).toBe(3);
    expect(spans.every(s => (s as HTMLElement).style.color === 'rgb(255, 0, 0)')).toBe(true);
    expect(editor.textContent).toBe('PierwszyDrugipunkt');
  });

  it('granice w środku tekstu: prefiks/sufiks poza zaznaczeniem zostają niepokolorowane', () => {
    editor.innerHTML = '<p id="a">ABCD</p><p id="b">WXYZ</p>';
    const aText = editor.querySelector('#a')!.firstChild as Text;
    const bText = editor.querySelector('#b')!.firstChild as Text;
    selectRange(r => {
      r.setStart(aText, 2); // AB|CD
      r.setEnd(bText, 2);   // WX|YZ
    });

    applyColor('#0000ff');

    // Block structure untouched — the old extract path left empty <p> shells here.
    expect(Array.from(editor.children).map(c => c.id)).toEqual(['a', 'b']);
    const a = editor.querySelector('#a')!;
    const b = editor.querySelector('#b')!;
    expect(a.textContent).toBe('ABCD');
    expect(b.textContent).toBe('WXYZ');
    const aSpan = a.querySelector('span') as HTMLElement;
    const bSpan = b.querySelector('span') as HTMLElement;
    expect(aSpan.textContent).toBe('CD');
    expect(bSpan.textContent).toBe('WX');
    expect(aSpan.style.color).toBe('rgb(0, 0, 255)');
    // Prefix "AB" and suffix "YZ" stay outside any color span.
    expect((a.firstChild as Text).textContent).toBe('AB');
    expect((b.lastChild as Text).textContent).toBe('YZ');
  });

  it('zakres CROSS-PAGE (selectAllContent): struktura stron nietknięta, chrome między stronami niepokolorowany', () => {
    // Simulate two pages: editors inside page scaffolding with a footer between them.
    const container = document.createElement('div');
    container.className = 'pages-container';
    container.innerHTML =
      '<div class="page-wrapper"><div class="page">' +
      '<div class="editor-content" contenteditable="true"><p>Strona pierwsza</p></div>' +
      '<div class="page-footer"><div class="footer-display"><p>Stopka</p></div></div>' +
      '</div></div>' +
      '<div class="page-wrapper"><div class="page">' +
      '<div class="editor-content" contenteditable="true"><p>Strona druga</p></div>' +
      '</div></div>';
    document.body.appendChild(container);
    const eds = Array.from(container.querySelectorAll('.editor-content'));
    (component as any).getActiveEditor = () => eds[0];
    const wrappersBefore = container.querySelectorAll('.page-wrapper').length;

    // The exact range selectAllContent() builds: first editor start → last editor end.
    selectRange(r => {
      r.setStart(eds[0], 0);
      r.setEnd(eds[1], eds[1].childNodes.length);
    });

    applyColor('#ff0000');

    // Page scaffolding intact — the old path doubled pages / emptied the first one.
    expect(container.querySelectorAll('.page-wrapper').length).toBe(wrappersBefore);
    expect(eds[0].textContent).toBe('Strona pierwsza');
    expect(eds[1].textContent).toBe('Strona druga');
    expect(eds[0].querySelector('span')).not.toBeNull();
    expect(eds[1].querySelector('span')).not.toBeNull();
    // Footer chrome between the editors is NOT part of document text — left alone.
    expect(container.querySelector('.footer-display span')).toBeNull();
    // No block nested inside another paragraph, no page-wrapper clones in content.
    expect(container.querySelector('.editor-content .page-wrapper, .editor-content p p')).toBeNull();
  });

  it('ponowne nadanie stylu reużywa istniejący span (bez narastania zagnieżdżenia)', () => {
    editor.innerHTML = '<p>Tekst</p>';
    const selectAll = () => selectRange(r => {
      r.setStart(editor, 0);
      r.setEnd(editor, editor.childNodes.length);
    });

    selectAll();
    applyColor('#ff0000');
    selectAll();
    applyColor('#00ff00');
    selectAll();
    const sel = window.getSelection()!;
    (component as any).applyFontSizeToSelection(14, sel, sel.getRangeAt(0));

    const spans = editor.querySelectorAll('span');
    expect(spans.length).toBe(1);
    const span = spans[0] as HTMLElement;
    expect(span.style.color).toBe('rgb(0, 255, 0)');
    expect(span.style.fontSize).toBe('14pt');
    expect(editor.querySelector('span span')).toBeNull();
  });

  it('wyspy contenteditable=false (markery) nie są owijane', () => {
    editor.innerHTML =
      '<p><span class="list-marker" contenteditable="false">1.</span>Treść punktu</p>';
    selectRange(r => {
      r.setStart(editor, 0);
      r.setEnd(editor, editor.childNodes.length);
    });

    applyColor('#ff0000');

    const marker = editor.querySelector('.list-marker')!;
    expect(marker.querySelector('span')).toBeNull();
    expect((marker as HTMLElement).style.color).toBe('');
    const contentSpan = Array.from(editor.querySelectorAll('p > span'))
      .find(s => s.textContent === 'Treść punktu') as HTMLElement;
    expect(contentSpan).toBeTruthy();
    expect(contentSpan.style.color).toBe('rgb(255, 0, 0)');
  });
});
