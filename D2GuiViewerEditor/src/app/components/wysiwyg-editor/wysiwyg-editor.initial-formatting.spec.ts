import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';
import { EditorToolbarComponent } from '../editor-toolbar/editor-toolbar';

/**
 * Regresja: „po otwarciu dokumentu 9 pt toolbar pokazuje 11 aż do 1. kliknięcia".
 *
 * Root cause: `updateFormattingState` liczyło formatowanie WYŁĄCZNIE z `window.getSelection()`.
 * Po załadowaniu dokumentu nie ma jeszcze selekcji w edytorze, więc metoda pomijała odczyt
 * i `editorState.currentStyle` zostawał pusty — toolbar trzymał swój sztywny default 11.
 * Dopiero klik (selectionchange) wyliczał rozmiar z runu.
 *
 * Fix (Wariant B — synchronizacja BEZ widocznego kursora): gdy nie ma selekcji w edytorze,
 * `updateFormattingState` rozwiązuje kontekst z PIERWSZEGO edytowalnego runu
 * (`_firstEditableContextElement`) i liczy z niego rozmiar/krój/pogrubienie itd. Bez fokusa,
 * bez selekcji, bez dirty/undo. Klik dalej działa (ścieżka queryCommandState).
 *
 * Uwaga jsdom: getComputedStyle NIE konwertuje pt→px ani nie kaskaduje font-size z rodzica,
 * dlatego testy zadają rozmiar INLINE w px na samym runie (12px = 9pt, 16px = 12pt).
 */
describe('WysiwygEditorComponent — początkowa synchronizacja rozmiaru czcionki (bez kliknięcia)', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;
  const mounted: HTMLElement[] = [];

  beforeEach(async () => {
    // jsdom nie implementuje queryCommandState — ścieżka selekcji musi przeżyć.
    if (typeof (document as any).queryCommandState !== 'function') {
      (document as any).queryCommandState = () => false;
    }
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent, EditorToolbarComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  afterEach(() => {
    mounted.splice(0).forEach(el => el.remove());
    window.getSelection()?.removeAllRanges();
  });

  /** Montuje treść jako aktywną stronę edytora (stub pageEditorRefs jak w innych specach). */
  function mountEditor(html: string): HTMLElement {
    const editor = document.createElement('div');
    editor.className = 'editor-content';
    editor.innerHTML = html;
    document.body.appendChild(editor);
    mounted.push(editor);
    (component as any).pageEditorRefs = { toArray: () => [{ nativeElement: editor }] };
    return editor;
  }

  function caretIn(node: Node, offset = 0): void {
    const sel = window.getSelection()!;
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
  }

  function syncedFontSize(): number | undefined {
    window.getSelection()?.removeAllRanges();
    (component as any).updateFormattingState();
    return component.editorState().currentStyle.fontSize;
  }

  it('dokument 9 pt → toolbar dostaje 9 BEZ kliknięcia (a nie domyślne 11)', () => {
    mountEditor('<p><span style="font-size:12px">tekst 9pt</span></p>');
    expect(syncedFontSize()).toBe(9);
  });

  it('dokument 11 pt → 11 (nie „przypadkowo" 9)', () => {
    mountEditor('<p><span style="font-size:14.6667px">tekst 11pt</span></p>');
    expect(syncedFontSize()).toBe(11);
  });

  it('rozmiar z bezpośredniego formatowania runu wygrywa (span > akapit)', () => {
    mountEditor('<p style="font-size:16px"><span style="font-size:12px">run</span></p>');
    // Kontekst = pierwszy run (12px = 9pt), nie akapit.
    expect(syncedFontSize()).toBe(9);
  });

  it('pusty pierwszy akapit → kontekst to blok (bez wyjątku, rozmiar z bloku)', () => {
    const editor = mountEditor('<p style="font-size:12px"><br></p>');
    const ctx = (component as any)._firstEditableContextElement() as HTMLElement;
    expect(ctx).toBe(editor.querySelector('p'));
    expect(syncedFontSize()).toBe(9);
  });

  it('dokument zaczyna się tabelą → kontekst to tekst pierwszej komórki', () => {
    mountEditor(
      '<table><tr><td><p><span style="font-size:12px">komórka</span></p></td></tr></table>',
    );
    const ctx = (component as any)._firstEditableContextElement() as HTMLElement;
    expect(ctx.textContent).toBe('komórka');
    expect(syncedFontSize()).toBe(9);
  });

  it('dokument zaczyna się obrazem (brak tekstu) → kontekst to blok, brak wyjątku', () => {
    const editor = mountEditor('<p style="font-size:16px"><img src="x"></p>');
    const ctx = (component as any)._firstEditableContextElement() as HTMLElement;
    expect(ctx).toBe(editor.querySelector('p'));
    expect(syncedFontSize()).toBe(12);
  });

  it('całkowicie pusty dokument → brak crashu, default 11', () => {
    mountEditor('');
    expect(syncedFontSize()).toBe(11);
  });

  it('przełączenie 9 pt → 12 pt aktualizuje wartość (bez kliknięcia w każdym)', () => {
    mountEditor('<p><span style="font-size:12px">A</span></p>');
    expect(syncedFontSize()).toBe(9);

    mountEditor('<p><span style="font-size:16px">B</span></p>');
    expect(syncedFontSize()).toBe(12);
  });

  it('klik w run 9 pt po inicjalizacji nadal daje 9 (ścieżka selekcji nie regresuje)', () => {
    const editor = mountEditor('<p><span style="font-size:12px">klik</span></p>');
    // Stub isSelectionInEditor: selekcja jest w naszym mount-edytorze.
    (component as any).isSelectionInEditor = () => true;
    caretIn(editor.querySelector('span')!.firstChild!, 1);

    (component as any).updateFormattingState();
    expect(component.editorState().currentStyle.fontSize).toBe(9);
  });

  it('pogrubienie pierwszego runu odczytane bez selekcji (spójny cały toolbar)', () => {
    mountEditor('<p><span style="font-size:12px;font-weight:700">bold</span></p>');
    window.getSelection()?.removeAllRanges();
    (component as any).updateFormattingState();

    const f = component.editorState().currentFormatting;
    expect(f.bold).toBe(true);
    expect(f.italic).toBe(false);
  });

  it('synchronizacja NIE oznacza dokumentu jako zmodyfikowany', () => {
    mountEditor('<p><span style="font-size:12px">x</span></p>');
    (component as any)._isDirty = false;
    window.getSelection()?.removeAllRanges();
    (component as any).updateFormattingState();

    expect((component as any)._isDirty).toBe(false);
    expect(component.editorState().isModified).toBe(false);
  });

  it('toolbar pokazuje 9 po odebraniu początkowego stanu (integracja z EditorToolbar)', () => {
    mountEditor('<p><span style="font-size:12px">9pt</span></p>');
    window.getSelection()?.removeAllRanges();
    (component as any).updateFormattingState();

    const toolbar = TestBed.createComponent(EditorToolbarComponent).componentInstance;
    toolbar.editorState = component.editorState();
    expect(toolbar.selectedFontSize()).toBe(9);
  });
});
