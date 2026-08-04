import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Regresja bug 13184834: „kursor wraca na początek tekstu po cofnięciu zmian".
 *
 * undo()/redo() podmieniają całe pageContents — rebind [innerHTML] przebudowuje DOM
 * stron i kasuje selekcję, więc karetka lądowała na początku dokumentu. Fix: kotwica
 * { blok, offset } (ta sama co przy repaginacji) zapisywana PRZED podmianą treści
 * i przywracana po re-renderze, z domknięciem offsetu/bloku do przywróconej treści.
 *
 * Testy symulują re-render Angulara ręcznym nadpisaniem innerHTML + skasowaniem
 * selekcji (dokładnie to robi rebind), a potem flushują setTimeout(0) undo/redo.
 */
describe('WysiwygEditorComponent — undo/redo przywraca pozycję kursora', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;
  let editor: HTMLDivElement;

  beforeEach(async () => {
    // jsdom nie ma queryCommandState — updateFormattingState (wołane po restore karetki
    // w setTimeout undo/redo) rzucałoby unhandled TypeError zaśmiecającym raport suity.
    if (typeof (document as any).queryCommandState !== 'function') {
      (document as any).queryCommandState = () => false;
    }
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;

    editor = document.createElement('div');
    editor.setAttribute('contenteditable', 'true');
    document.body.appendChild(editor);
    (component as any).editorContent = { nativeElement: editor };
    (component as any).pageEditorRefs = { toArray: () => [{ nativeElement: editor }] };
    // Izolacja od repaginacji (mierzy realny layout, którego jsdom nie ma).
    (component as any)._schedulePaginate = vi.fn();
  });

  afterEach(() => {
    editor.remove();
    window.getSelection()?.removeAllRanges();
  });

  function caretIn(node: Node, offset: number): void {
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
  }

  /** Symuluje rebind [innerHTML] po zmianie pageContents: nowy DOM, selekcja skasowana. */
  function simulateRerender(html: string): void {
    editor.innerHTML = html;
    window.getSelection()!.removeAllRanges();
  }

  /** Czeka aż odpali się setTimeout(0) zaplanowany w undo()/redo(). */
  function flushTimeout(): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, 0));
  }

  /** Offset tekstowy karetki liczony od początku bloku `block`. */
  function caretTextOffset(block: Element): number {
    const sel = window.getSelection()!;
    expect(sel.rangeCount).toBe(1);
    const range = sel.getRangeAt(0);
    expect(block.contains(range.startContainer)).toBe(true);
    const pre = range.cloneRange();
    pre.selectNodeContents(block);
    pre.setEnd(range.startContainer, range.startOffset);
    return pre.toString().length;
  }

  it('po undo kursor zostaje w miejscu edycji (koniec „Hello Worl"), nie na początku dokumentu', async () => {
    editor.innerHTML = '<p>Hello World</p>';
    (component as any).undoStack = ['<p>Hello Worl</p>', '<p>Hello World</p>'];
    (component as any).redoStack = [];
    caretIn(editor.querySelector('p')!.firstChild!, 11); // za "World"

    component.undo();
    simulateRerender('<p>Hello Worl</p>');
    await flushTimeout();

    // Offset 11 domknięty do długości przywróconego tekstu (10 = za "l").
    expect(caretTextOffset(editor.querySelector('p')!)).toBe(10);
  });

  it('undo w środku tekstu przywraca karetkę na ten sam offset', async () => {
    editor.innerHTML = '<p>Hello brave World</p>';
    (component as any).undoStack = ['<p>Hello World</p>', '<p>Hello brave World</p>'];
    (component as any).redoStack = [];
    caretIn(editor.querySelector('p')!.firstChild!, 6); // przed "brave"

    component.undo();
    simulateRerender('<p>Hello World</p>');
    await flushTimeout();

    expect(caretTextOffset(editor.querySelector('p')!)).toBe(6);
  });

  it('gdy fokus zabrał toolbar (brak żywej selekcji w edytorze), kotwica pochodzi z savedSelection', async () => {
    editor.innerHTML = '<p>Hello World</p>';
    (component as any).undoStack = ['<p>Hello</p>', '<p>Hello World</p>'];
    (component as any).redoStack = [];
    const saved = document.createRange();
    saved.setStart(editor.querySelector('p')!.firstChild!, 8);
    saved.collapse(true);
    (component as any).savedSelection = saved;
    window.getSelection()!.removeAllRanges(); // klik w przycisk toolbara

    component.undo();
    simulateRerender('<p>Hello</p>');
    await flushTimeout();

    // Offset 8 domknięty do końca „Hello" (5).
    expect(caretTextOffset(editor.querySelector('p')!)).toBe(5);
  });

  it('undo usuwające blok z kursorem domyka karetkę do końca dokumentu', async () => {
    editor.innerHTML = '<p>One</p><p>Two</p>';
    (component as any).undoStack = ['<p>One</p>', '<p>One</p><p>Two</p>'];
    (component as any).redoStack = [];
    caretIn(editor.querySelectorAll('p')[1].firstChild!, 2); // w "Two"

    component.undo();
    simulateRerender('<p>One</p>');
    await flushTimeout();

    expect(caretTextOffset(editor.querySelector('p')!)).toBe(3); // koniec "One"
  });

  it('klik „Cofnij" w oknie debounce\'a cofa świeżo wpisany tekst (flush wiszącego snapshotu)', async () => {
    // Bug 13184834 (komentarz QA): „przy pierwszym kliknięciu cofaj nic się nie wydarzyło" —
    // snapshot jechał z debounce 500 ms, więc stos nie znał ostatniej porcji pisania.
    editor.innerHTML = '<p>Hello World</p>';
    (component as any).undoStack = ['<p>Hello Worl</p>'];
    (component as any).redoStack = [];
    caretIn(editor.querySelector('p')!.firstChild!, 11);
    (component as any)._schedulePersist(); // timer wisi — jak w trakcie pisania

    component.undo(); // PIERWSZE kliknięcie — musi cofnąć, nie czekać na timer

    expect((component as any).undoStack.length).toBe(1);
    expect(component.pageContents()[0]).toContain('Hello Worl');
    expect(component.pageContents()[0]).not.toContain('Hello World');
    expect((component as any).redoStack.length).toBe(1);
    expect((component as any)._persistTimer).toBeNull();
  });

  it('canUndo jest true w oknie debounce\'a (przycisk nie może wyglądać na martwy)', () => {
    editor.innerHTML = '<p>Hello World</p>';
    (component as any).undoStack = ['<p>Hello Worl</p>'];
    (component as any)._schedulePersist();

    (component as any).updateState();

    expect(component.editorState().canUndo).toBe(true);
  });

  it('redo ustawia karetkę PO ponownie wstawionym tekście', async () => {
    editor.innerHTML = '<p>Hello Worl</p>';
    (component as any).undoStack = ['<p>Hello Worl</p>'];
    // New redo-entry shape: the caret stored with the snapshot is the position AFTER that
    // edit's inserted text (end of "Hello World" = offset 11), which is where redo must land —
    // not the pre-redo caret (offset 10) which sits before the re-inserted "d".
    (component as any).redoStack = [{ html: '<p>Hello World</p>', caret: { block: 0, offset: 11 } }];
    caretIn(editor.querySelector('p')!.firstChild!, 10);

    component.redo();
    simulateRerender('<p>Hello World</p>');
    await flushTimeout();

    expect(caretTextOffset(editor.querySelector('p')!)).toBe(11);
  });
});
