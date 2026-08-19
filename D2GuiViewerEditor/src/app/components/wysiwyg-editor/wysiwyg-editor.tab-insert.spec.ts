import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Tab jak w MS Word (Problem 1): Tab z kolapsowaną karetką W ŚRODKU akapitu wstawia
 * znak tabulacji (ten sam nośnik, który emituje reader DOCX: inline-block span
 * z literalnym \t i min-width:2em — writer konwertuje go na w:tab), a NIE wcina
 * całego akapitu. Wcięcie zostaje tylko: na początku bloku, przy zaznaczeniu
 * (niekolapsowanym) oraz dla Shift+Tab (outdent). Tab w tabeli — bez zmian:
 * nawigacja po komórkach (_handleTableTab konsumuje zdarzenie).
 */
describe('WysiwygEditorComponent — Tab wstawia znak tabulacji w środku akapitu', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;
  let editor: HTMLDivElement;
  let contentChanged: ReturnType<typeof vi.fn>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({ imports: [WysiwygEditorComponent] }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;

    editor = document.createElement('div');
    editor.className = 'editor-content';
    editor.setAttribute('contenteditable', 'true');
    document.body.appendChild(editor);
    contentChanged = vi.fn();
    (component as any).getActiveEditor = () => editor;
    (component as any).isSelectionInEditor = () => true;
    (component as any).onContentChange = contentChanged;
    (component as any).updateFormattingState = () => {};
  });

  afterEach(() => {
    editor.remove();
    window.getSelection()?.removeAllRanges();
  });

  function caretIn(node: Node, offset = 0): void {
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
  }

  function pressTab(shiftKey = false): void {
    (component as any).handleKeyboard(new KeyboardEvent('keydown', { key: 'Tab', shiftKey }));
  }

  const carrier = () => editor.querySelector<HTMLElement>("span[style*='min-width:2em']");

  it('karetka w środku akapitu: Tab wstawia nośnik tabulatora, NIE wcina akapitu', () => {
    editor.innerHTML = '<p id="p">Tekst akapitu</p>';
    const p = editor.querySelector<HTMLElement>('#p')!;
    caretIn(p.firstChild!, 5);
    const exec = vi.spyOn(component, 'executeCommand');

    pressTab();

    const span = carrier();
    expect(span).not.toBeNull();
    expect(span!.textContent).toBe('\t');
    expect(span!.getAttribute('style')).toContain('display:inline-block');
    // Nośnik dokładnie w miejscu karetki (po "Tekst").
    expect(p.textContent).toBe('Tekst\t akapitu');
    expect(exec).not.toHaveBeenCalled();
    expect(p.style.marginLeft).toBe('');
    expect(contentChanged).toHaveBeenCalled();
  });

  it('karetka ZA nośnikiem po wstawieniu (kolejny Tab dokłada następny nośnik za pierwszym)', () => {
    editor.innerHTML = '<p id="p">ab</p>';
    const p = editor.querySelector<HTMLElement>('#p')!;
    caretIn(p.firstChild!, 1);

    pressTab();
    pressTab();

    const spans = editor.querySelectorAll("span[style*='min-width:2em']");
    expect(spans.length).toBe(2);
    // Oba między "a" i "b" — drugi za pierwszym.
    expect(p.textContent).toBe('a\t\tb');
  });

  it('karetka na początku akapitu: Tab wcina akapit (bez nośnika)', () => {
    editor.innerHTML = '<p id="p">Tekst</p>';
    const p = editor.querySelector<HTMLElement>('#p')!;
    caretIn(p.firstChild!, 0);
    const exec = vi.spyOn(component, 'executeCommand');

    pressTab();

    expect(carrier()).toBeNull();
    expect(exec).toHaveBeenCalledWith('indent');
  });

  it('początek bloku liczy się przez puste/zero-width poprzedniki (span z ZWSP przed karetką)', () => {
    editor.innerHTML = '<p id="p"><span>' + String.fromCharCode(0x200b) + '</span><span id="s">Tekst</span></p>';
    const s = editor.querySelector<HTMLElement>('#s')!;
    caretIn(s.firstChild!, 0);
    const exec = vi.spyOn(component, 'executeCommand');

    pressTab();

    expect(carrier()).toBeNull();
    expect(exec).toHaveBeenCalledWith('indent');
  });

  it('zaznaczenie (niekolapsowane): Tab wcina, nie wstawia nośnika', () => {
    editor.innerHTML = '<p id="p">Tekst</p>';
    const p = editor.querySelector<HTMLElement>('#p')!;
    const range = document.createRange();
    range.setStart(p.firstChild!, 1);
    range.setEnd(p.firstChild!, 3);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
    const exec = vi.spyOn(component, 'executeCommand');

    pressTab();

    expect(carrier()).toBeNull();
    expect(exec).toHaveBeenCalledWith('indent');
  });

  it('Shift+Tab w środku akapitu: outdent jak dotąd, bez nośnika', () => {
    editor.innerHTML = '<p id="p" style="margin-left:48px">Tekst</p>';
    const p = editor.querySelector<HTMLElement>('#p')!;
    caretIn(p.firstChild!, 3);
    const exec = vi.spyOn(component, 'executeCommand');

    pressTab(true);

    expect(carrier()).toBeNull();
    expect(exec).toHaveBeenCalledWith('outdent');
  });

  it('karetka na KOŃCU akapitu: Tab wstawia nośnik, a karetka ląduje ZA nim (pisanie kontynuuje po tabie)', () => {
    editor.innerHTML = '<p id="p">Koniec akapitu</p>';
    const p = editor.querySelector<HTMLElement>('#p')!;
    caretIn(p.firstChild!, p.firstChild!.textContent!.length);

    pressTab();

    const span = carrier();
    expect(span).not.toBeNull();
    // Atomowy jak tab Worda: bez contenteditable=false Chrome kanonizował karetkę PRZED span
    // (trailing \t = kolapsowany biały znak) i pisanie po Tabie lądowało przed tabulatorem.
    expect(span!.getAttribute('contenteditable')).toBe('false');
    const sel = window.getSelection()!;
    const r = sel.getRangeAt(0);
    expect(r.collapsed).toBe(true);
    expect(r.startContainer).toBe(p);
    expect(r.startOffset).toBe(2); // za spanem (dzieci: [tekst, span])
  });

  it('_placeCaretAtTextOffset nie stawia karetki WEWNĄTRZ atomowego nośnika (contenteditable=false)', () => {
    editor.innerHTML =
      '<p id="p">ab<span style="display:inline-block;min-width:2em" contenteditable="false">\t</span>cd</p>';
    const p = editor.querySelector<HTMLElement>('#p')!;

    // Offset 3 = koniec \t — granica atomu: karetka ma wylądować ZA spanem, nie w jego tekście.
    (component as any)._placeCaretAtTextOffset(p, 3);

    const r = window.getSelection()!.getRangeAt(0);
    expect(r.startContainer).toBe(p);
    expect(r.startOffset).toBe(2); // za spanem
  });

  it('Tab w tabeli: _handleTableTab konsumuje zdarzenie — bez nośnika i bez indentu', () => {
    editor.innerHTML = '<p id="p">Tekst</p>';
    const p = editor.querySelector<HTMLElement>('#p')!;
    caretIn(p.firstChild!, 3);
    const tableTab = vi.fn().mockReturnValue(true);
    (component as any)._handleTableTab = tableTab;
    const exec = vi.spyOn(component, 'executeCommand');

    pressTab();

    expect(tableTab).toHaveBeenCalledWith(false);
    expect(carrier()).toBeNull();
    expect(exec).not.toHaveBeenCalled();
  });
});
