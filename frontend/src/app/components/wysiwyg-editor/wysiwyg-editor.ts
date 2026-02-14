import { 
  Component, 
  ElementRef, 
  EventEmitter, 
  Input, 
  Output, 
  ViewChild, 
  AfterViewInit,
  OnDestroy,
  inject,
  signal,
  computed
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { 
  EditorCommand, 
  EditorState, 
  HeadingLevel, 
  TextFormatting,
  ParagraphStyle,
  PageMargins
} from '../../models/document.model';

/**
 * Komponent edytora WYSIWYG
 * Własna implementacja edytora contenteditable
 */
@Component({
  selector: 'app-wysiwyg-editor',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './wysiwyg-editor.html',
  styleUrl: './wysiwyg-editor.css'
})
export class WysiwygEditorComponent implements AfterViewInit, OnDestroy {
  @ViewChild('editorContent') editorContent!: ElementRef<HTMLDivElement>;
  
  @Input() set content(value: string) {
    // Nie aktualizuj innerHTML jeśli wartość pochodzi z tego samego edytora
    // (zapobiega resetowaniu kursora podczas pisania)
    if (this._content() === value) {
      return;
    }
    
    this._content.set(value);
    if (this.editorContent?.nativeElement && !this._isInternalUpdate) {
      this.editorContent.nativeElement.innerHTML = value;
    }
  }
  
  @Input() pageMargins: PageMargins = { top: 2.5, bottom: 2.5, left: 2.5, right: 2.5 };
  @Input() pageOrientation: 'portrait' | 'landscape' = 'portrait';
  @Input() showMarginGuides = true;
  
  @Output() contentChange = new EventEmitter<string>();
  @Output() stateChange = new EventEmitter<EditorState>();
  @Output() selectionChange = new EventEmitter<Selection | null>();

  private _content = signal<string>('');
  private _isInternalUpdate = false;
  private undoStack: string[] = [];
  private redoStack: string[] = [];
  private lastSavedContent = '';
  private autoSaveInterval?: ReturnType<typeof setInterval>;

  // Stan edytora
  editorState = signal<EditorState>({
    isModified: false,
    canUndo: false,
    canRedo: false,
    wordCount: 0,
    characterCount: 0,
    currentFormatting: {
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      subscript: false,
      superscript: false
    },
    currentStyle: {}
  });

  ngAfterViewInit(): void {
    this.initializeEditor();
    this.setupEventListeners();
    
    // Autosave co 30 sekund (zapisuje do localStorage)
    this.autoSaveInterval = setInterval(() => {
      this.autoSave();
    }, 30000);
  }

  ngOnDestroy(): void {
    if (this.autoSaveInterval) {
      clearInterval(this.autoSaveInterval);
    }
  }

  /**
   * Inicjalizuje edytor
   */
  private initializeEditor(): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    editor.innerHTML = this._content() || '<p>&nbsp;</p>';
    this.lastSavedContent = editor.innerHTML;
    this.saveToUndoStack();
    this.updateState();
  }

  /**
   * Konfiguruje nasłuchiwanie zdarzeń
   */
  private setupEventListeners(): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    // Input - gdy użytkownik zmienia zawartość
    editor.addEventListener('input', () => {
      this.onContentChange();
    });

    // Selekcja - gdy użytkownik zaznacza tekst
    document.addEventListener('selectionchange', () => {
      this.onSelectionChange();
    });

    // Paste - specjalna obsługa wklejania
    editor.addEventListener('paste', (e) => {
      this.handlePaste(e);
    });

    // Keyboard shortcuts
    editor.addEventListener('keydown', (e) => {
      this.handleKeyboard(e);
    });

    // Drop - obsługa przeciągania
    editor.addEventListener('drop', (e) => {
      this.handleDrop(e);
    });
  }

  /**
   * Obsługa zmiany zawartości
   */
  private onContentChange(): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    const html = editor.innerHTML;
    this._isInternalUpdate = true;
    this._content.set(html);
    this.contentChange.emit(html);
    this._isInternalUpdate = false;
    this.saveToUndoStack();
    this.updateState();
  }

  /**
   * Obsługa zmiany selekcji
   */
  private onSelectionChange(): void {
    const selection = window.getSelection();
    
    if (selection && this.isSelectionInEditor(selection)) {
      this.updateFormattingState();
      this.selectionChange.emit(selection);
    }
  }

  /**
   * Sprawdza czy selekcja jest w edytorze
   */
  private isSelectionInEditor(selection: Selection): boolean {
    const editor = this.editorContent?.nativeElement;
    if (!editor || !selection.anchorNode) return false;
    return editor.contains(selection.anchorNode);
  }

  /**
   * Obsługa wklejania
   */
  private handlePaste(e: ClipboardEvent): void {
    e.preventDefault();
    
    const clipboardData = e.clipboardData;
    if (!clipboardData) return;

    // Spróbuj pobrać HTML
    let html = clipboardData.getData('text/html');
    
    if (html) {
      // Oczyść HTML z niechcianych elementów
      html = this.sanitizeHtml(html);
      this.insertHtml(html);
    } else {
      // Wklej jako zwykły tekst
      const text = clipboardData.getData('text/plain');
      this.insertText(text);
    }
  }

  /**
   * Oczyszcza HTML z niechcianych elementów
   */
  private sanitizeHtml(html: string): string {
    // Usuń skrypty
    html = html.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
    // Usuń style globalne (zachowaj inline)
    html = html.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '');
    // Usuń komentarze
    html = html.replace(/<!--[\s\S]*?-->/g, '');
    // Usuń atrybuty onclick, onerror itp.
    html = html.replace(/\s*on\w+\s*=\s*["'][^"']*["']/gi, '');
    
    return html;
  }

  /**
   * Obsługa skrótów klawiszowych
   */
  private handleKeyboard(e: KeyboardEvent): void {
    if (e.ctrlKey || e.metaKey) {
      switch (e.key.toLowerCase()) {
        case 'b':
          e.preventDefault();
          this.executeCommand('bold');
          break;
        case 'i':
          e.preventDefault();
          this.executeCommand('italic');
          break;
        case 'u':
          e.preventDefault();
          this.executeCommand('underline');
          break;
        case 'z':
          e.preventDefault();
          if (e.shiftKey) {
            this.redo();
          } else {
            this.undo();
          }
          break;
        case 'y':
          e.preventDefault();
          this.redo();
          break;
        case 's':
          e.preventDefault();
          // Emituj zdarzenie zapisu (obsługiwane przez komponent nadrzędny)
          break;
      }
    }

    // Tab - wcięcie
    if (e.key === 'Tab') {
      e.preventDefault();
      if (e.shiftKey) {
        this.executeCommand('outdent');
      } else {
        this.executeCommand('indent');
      }
    }
  }

  /**
   * Obsługa upuszczania plików
   */
  private handleDrop(e: DragEvent): void {
    e.preventDefault();
    
    const files = e.dataTransfer?.files;
    if (!files || files.length === 0) return;

    // Obsłuż obrazy
    Array.from(files).forEach(file => {
      if (file.type.startsWith('image/')) {
        this.insertImageFromFile(file);
      }
    });
  }

  /**
   * Wstawia obraz z pliku
   */
  private insertImageFromFile(file: File): void {
    const reader = new FileReader();
    reader.onload = (e) => {
      const base64 = e.target?.result as string;
      if (base64) {
        this.insertImage(base64);
      }
    };
    reader.readAsDataURL(file);
  }

  /**
   * Wykonuje komendę edytora
   */
  executeCommand(command: EditorCommand, value?: string): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    // Upewnij się, że edytor ma focus
    editor.focus();

    switch (command) {
      case 'bold':
        document.execCommand('bold', false);
        break;
      case 'italic':
        document.execCommand('italic', false);
        break;
      case 'underline':
        document.execCommand('underline', false);
        break;
      case 'strikethrough':
        document.execCommand('strikeThrough', false);
        break;
      case 'subscript':
        document.execCommand('subscript', false);
        break;
      case 'superscript':
        document.execCommand('superscript', false);
        break;
      case 'alignLeft':
        document.execCommand('justifyLeft', false);
        break;
      case 'alignCenter':
        document.execCommand('justifyCenter', false);
        break;
      case 'alignRight':
        document.execCommand('justifyRight', false);
        break;
      case 'alignJustify':
        document.execCommand('justifyFull', false);
        break;
      case 'indent':
        document.execCommand('indent', false);
        break;
      case 'outdent':
        document.execCommand('outdent', false);
        break;
      case 'bulletList':
        document.execCommand('insertUnorderedList', false);
        break;
      case 'numberedList':
        document.execCommand('insertOrderedList', false);
        break;
      case 'removeFormat':
        document.execCommand('removeFormat', false);
        break;
      case 'undo':
        this.undo();
        return;
      case 'redo':
        this.redo();
        return;
      case 'heading1':
      case 'heading2':
      case 'heading3':
      case 'heading4':
      case 'heading5':
      case 'heading6':
        const level = command.replace('heading', '');
        document.execCommand('formatBlock', false, `h${level}`);
        break;
      case 'paragraph':
        document.execCommand('formatBlock', false, 'p');
        break;
      case 'insertLink':
        if (value) {
          document.execCommand('createLink', false, value);
        }
        break;
      case 'insertImage':
        if (value) {
          this.insertImage(value);
        }
        break;
      case 'insertTable':
        if (value) {
          this.insertTable(value);
        }
        break;
    }

    this.onContentChange();
    this.updateFormattingState();
  }

  /**
   * Ustawia rozmiar czcionki
   */
  setFontSize(size: number): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    editor.focus();
    
    // execCommand fontSize używa wartości 1-7, więc używamy innego podejścia
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      
      if (!range.collapsed) {
        const span = document.createElement('span');
        span.style.fontSize = `${size}pt`;
        
        try {
          range.surroundContents(span);
        } catch {
          // Jeśli selekcja zawiera wiele elementów
          document.execCommand('fontSize', false, '7');
          // Znajdź i zamień font size
          const fonts = editor.querySelectorAll('font[size="7"]');
          fonts.forEach(font => {
            const newSpan = document.createElement('span');
            newSpan.style.fontSize = `${size}pt`;
            newSpan.innerHTML = font.innerHTML;
            font.parentNode?.replaceChild(newSpan, font);
          });
        }
      }
    }

    this.onContentChange();
  }

  /**
   * Ustawia rodzinę czcionki
   */
  setFontFamily(fontFamily: string): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    editor.focus();
    document.execCommand('fontName', false, fontFamily);
    this.onContentChange();
  }

  /**
   * Ustawia kolor tekstu
   */
  setTextColor(color: string): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    editor.focus();
    document.execCommand('foreColor', false, color);
    this.onContentChange();
  }

  /**
   * Ustawia kolor tła
   */
  setBackgroundColor(color: string): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    editor.focus();
    document.execCommand('hiliteColor', false, color);
    this.onContentChange();
  }

  /**
   * Wstawia tekst
   */
  insertText(text: string): void {
    document.execCommand('insertText', false, text);
    this.onContentChange();
  }

  /**
   * Wstawia HTML
   */
  insertHtml(html: string): void {
    document.execCommand('insertHTML', false, html);
    this.onContentChange();
  }

  /**
   * Wstawia obraz
   */
  insertImage(src: string, alt: string = ''): void {
    const img = `<img src="${src}" alt="${alt}" style="max-width:100%;height:auto;" />`;
    this.insertHtml(img);
  }

  /**
   * Wstawia tabelę
   */
  insertTable(config: string): void {
    const [rows, cols] = config.split('x').map(Number);
    
    let tableHtml = '<table style="border-collapse:collapse;width:100%;margin:10px 0;">';
    
    for (let i = 0; i < rows; i++) {
      tableHtml += '<tr>';
      for (let j = 0; j < cols; j++) {
        tableHtml += '<td style="border:1px solid #ccc;padding:8px;min-width:50px;">&nbsp;</td>';
      }
      tableHtml += '</tr>';
    }
    
    tableHtml += '</table><p>&nbsp;</p>';
    this.insertHtml(tableHtml);
  }

  /**
   * Wstawia link
   */
  insertLink(url: string, text?: string): void {
    const selection = window.getSelection();
    
    if (selection && !selection.isCollapsed) {
      document.execCommand('createLink', false, url);
    } else if (text) {
      const link = `<a href="${url}" target="_blank">${text}</a>`;
      this.insertHtml(link);
    }
    
    this.onContentChange();
  }

  /**
   * Wstawia poziomą linię
   */
  insertHorizontalRule(): void {
    document.execCommand('insertHorizontalRule', false);
    this.onContentChange();
  }

  /**
   * Wstawia podział strony
   */
  insertPageBreak(): void {
    this.insertHtml('<div class="page-break" style="page-break-after:always;border-top:2px dashed #ccc;margin:20px 0;"></div>');
  }

  /**
   * Undo
   */
  undo(): void {
    if (this.undoStack.length > 1) {
      const current = this.undoStack.pop()!;
      this.redoStack.push(current);
      
      const previous = this.undoStack[this.undoStack.length - 1];
      if (this.editorContent?.nativeElement) {
        this.editorContent.nativeElement.innerHTML = previous;
        this._content.set(previous);
        this.contentChange.emit(previous);
      }
      
      this.updateState();
    }
  }

  /**
   * Redo
   */
  redo(): void {
    if (this.redoStack.length > 0) {
      const next = this.redoStack.pop()!;
      this.undoStack.push(next);
      
      if (this.editorContent?.nativeElement) {
        this.editorContent.nativeElement.innerHTML = next;
        this._content.set(next);
        this.contentChange.emit(next);
      }
      
      this.updateState();
    }
  }

  /**
   * Zapisuje do stosu undo
   */
  private saveToUndoStack(): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    const html = editor.innerHTML;
    
    // Nie zapisuj jeśli to samo co ostatni wpis
    if (this.undoStack.length > 0 && this.undoStack[this.undoStack.length - 1] === html) {
      return;
    }

    this.undoStack.push(html);
    
    // Ogranicz rozmiar stosu
    if (this.undoStack.length > 100) {
      this.undoStack.shift();
    }

    // Wyczyść redo po nowej akcji
    this.redoStack = [];
  }

  /**
   * Aktualizuje stan formatowania
   */
  private updateFormattingState(): void {
    const formatting: TextFormatting = {
      bold: document.queryCommandState('bold'),
      italic: document.queryCommandState('italic'),
      underline: document.queryCommandState('underline'),
      strikethrough: document.queryCommandState('strikeThrough'),
      subscript: document.queryCommandState('subscript'),
      superscript: document.queryCommandState('superscript')
    };

    const currentBlock = document.queryCommandValue('formatBlock');
    const fontSize = document.queryCommandValue('fontSize');
    const fontName = document.queryCommandValue('fontName');
    const foreColor = document.queryCommandValue('foreColor');

    this.editorState.update(state => ({
      ...state,
      currentFormatting: formatting,
      currentStyle: {
        fontFamily: fontName || 'Calibri',
        fontSize: fontSize ? parseInt(fontSize) : 11,
        textColor: foreColor || '#000000'
      }
    }));

    this.stateChange.emit(this.editorState());
  }

  /**
   * Aktualizuje ogólny stan edytora
   */
  private updateState(): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    const text = editor.innerText || '';
    const words = text.trim().split(/\s+/).filter(w => w.length > 0);

    this.editorState.update(state => ({
      ...state,
      isModified: editor.innerHTML !== this.lastSavedContent,
      canUndo: this.undoStack.length > 1,
      canRedo: this.redoStack.length > 0,
      wordCount: words.length,
      characterCount: text.length
    }));

    this.stateChange.emit(this.editorState());
  }

  /**
   * Autosave do localStorage
   */
  private autoSave(): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    const html = editor.innerHTML;
    if (html && html !== '<p>&nbsp;</p>') {
      localStorage.setItem('editor_autosave', html);
      localStorage.setItem('editor_autosave_time', new Date().toISOString());
    }
  }

  /**
   * Przywraca z autosave
   */
  restoreFromAutoSave(): boolean {
    const saved = localStorage.getItem('editor_autosave');
    if (saved && this.editorContent?.nativeElement) {
      this.editorContent.nativeElement.innerHTML = saved;
      this._content.set(saved);
      this.contentChange.emit(saved);
      this.saveToUndoStack();
      this.updateState();
      return true;
    }
    return false;
  }

  /**
   * Czyści autosave
   */
  clearAutoSave(): void {
    localStorage.removeItem('editor_autosave');
    localStorage.removeItem('editor_autosave_time');
  }

  /**
   * Pobiera zawartość HTML
   */
  getContent(): string {
    return this.editorContent?.nativeElement?.innerHTML || '';
  }

  /**
   * Ustawia zawartość HTML
   */
  setContent(html: string): void {
    if (this.editorContent?.nativeElement) {
      this.editorContent.nativeElement.innerHTML = html;
      this._content.set(html);
      this.lastSavedContent = html;
      this.undoStack = [html];
      this.redoStack = [];
      this.updateState();
    }
  }

  /**
   * Oznacza dokument jako zapisany
   */
  markAsSaved(): void {
    this.lastSavedContent = this.getContent();
    this.updateState();
  }

  /**
   * Focus na edytorze
   */
  focus(): void {
    this.editorContent?.nativeElement?.focus();
  }
}
