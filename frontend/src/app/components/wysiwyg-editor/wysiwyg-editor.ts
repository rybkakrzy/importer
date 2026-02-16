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
  computed,
  ViewChildren,
  QueryList
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { 
  EditorCommand, 
  EditorState, 
  HeadingLevel, 
  TextFormatting,
  ParagraphStyle,
  PageMargins,
  HeaderFooterContent
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
  styleUrl: './wysiwyg-editor.scss'
})
export class WysiwygEditorComponent implements AfterViewInit, OnDestroy {
  @ViewChild('editorContent') editorContent!: ElementRef<HTMLDivElement>;
  @ViewChild('headerContent') headerContentEl?: ElementRef<HTMLDivElement>;
  @ViewChild('footerContent') footerContentEl?: ElementRef<HTMLDivElement>;
  
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
  
  // Nagłówek i stopka
  @Input() set headerContent(value: HeaderFooterContent | undefined) {
    if (value) {
      this._headerHtml.set(value.html || '');
      this._headerHeight.set(value.height || 1.27);
      if (value.differentFirstPage !== undefined) {
        this._differentFirstPage.set(value.differentFirstPage);
      }
      if (value.firstPageHtml !== undefined) {
        this._headerFirstPageHtml.set(value.firstPageHtml);
      }
      if (value.differentOddEven !== undefined) {
        this._differentOddEven.set(value.differentOddEven);
      }
      if (value.oddHtml !== undefined) {
        this._headerOddHtml.set(value.oddHtml);
      }
      if (value.evenHtml !== undefined) {
        this._headerEvenHtml.set(value.evenHtml);
      }
    }
  }
  
  @Input() set footerContent(value: HeaderFooterContent | undefined) {
    if (value) {
      this._footerHtml.set(value.html || '');
      this._footerHeight.set(value.height || 1.27);
      if (value.differentFirstPage !== undefined) {
        this._differentFirstPage.set(value.differentFirstPage);
      }
      if (value.firstPageHtml !== undefined) {
        this._footerFirstPageHtml.set(value.firstPageHtml);
      }
      if (value.differentOddEven !== undefined) {
        this._differentOddEven.set(value.differentOddEven);
      }
      if (value.oddHtml !== undefined) {
        this._footerOddHtml.set(value.oddHtml);
      }
      if (value.evenHtml !== undefined) {
        this._footerEvenHtml.set(value.evenHtml);
      }
    }
  }
  
  @Output() contentChange = new EventEmitter<string>();
  @Output() stateChange = new EventEmitter<EditorState>();
  @Output() selectionChange = new EventEmitter<Selection | null>();
  @Output() pagesChange = new EventEmitter<number>();
  @Output() headerChange = new EventEmitter<HeaderFooterContent>();
  @Output() footerChange = new EventEmitter<HeaderFooterContent>();
  @Output() openHeaderFooterSettings = new EventEmitter<{
    headerMargin: number;
    footerMargin: number;
    differentFirstPage: boolean;
    differentOddEven: boolean;
  }>();

  private _content = signal<string>('');
  private _isInternalUpdate = false;
  private undoStack: string[] = [];
  private redoStack: string[] = [];
  private lastSavedContent = '';
  private pageCheckInterval?: ReturnType<typeof setInterval>;

  // Strony dokumentu - pierwsza strona to edytor, pozostałe to overflow
  pages = signal<string[]>(['']);
  
  // Wysokość strony A4 w pikselach (bez marginesów)
  private readonly PAGE_HEIGHT_PX = 1122; // ~29.7cm at 96 DPI

  // Bieżący rozmiar czcionki (dla nowego tekstu gdy nie ma zaznaczenia)
  private currentFontSize = 11;
  private currentFontFamily = 'Calibri';
  private pendingFontSize: number | null = null;
  private pendingFontFamily: string | null = null;

  // Nagłówek i stopka - stan
  private _headerHtml = signal<string>('');
  private _footerHtml = signal<string>('');
  private _headerFirstPageHtml = signal<string>('');
  private _footerFirstPageHtml = signal<string>('');
  private _headerOddHtml = signal<string>('');
  private _headerEvenHtml = signal<string>('');
  private _footerOddHtml = signal<string>('');
  private _footerEvenHtml = signal<string>('');
  private _headerHeight = signal<number>(1.27); // domyślnie 1.27 cm (jak w Google Docs)
  private _footerHeight = signal<number>(1.27); // domyślnie 1.27 cm
  private _differentFirstPage = signal<boolean>(false);
  private _differentOddEven = signal<boolean>(false);
  editingSection = signal<'header' | 'footer' | 'body'>('body');
  
  // Menu opcji nagłówka/stopki
  showHeaderOptionsMenu = signal<boolean>(false);
  showFooterOptionsMenu = signal<boolean>(false);
  
  // Publiczne gettery dla template
  headerHeight = computed(() => this._headerHeight());
  footerHeight = computed(() => this._footerHeight());
  differentFirstPage = computed(() => this._differentFirstPage());
  differentOddEven = computed(() => this._differentOddEven());

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
    
    // Oblicz strony przy starcie
    setTimeout(() => {
      this.calculatePages();
      this.pagesChange.emit(this.pages().length);
    }, 100);
    
    // Sprawdzaj podział na strony co 500ms
    this.pageCheckInterval = setInterval(() => {
      this.calculatePages();
    }, 500);
  }

  ngOnDestroy(): void {
    if (this.pageCheckInterval) {
      clearInterval(this.pageCheckInterval);
    }
  }

  /**
   * Oblicza liczbę stron na podstawie wysokości zawartości
   */
  private calculatePages(): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;
    
    const contentHeight = editor.scrollHeight;
    const marginTop = this.pageMargins.top * 37.8;
    const marginBottom = this.pageMargins.bottom * 37.8;
    const availableHeight = this.PAGE_HEIGHT_PX - marginTop - marginBottom;
    
    const pageCount = Math.max(1, Math.ceil(contentHeight / availableHeight));
    
    // Aktualizuj liczbę stron tylko jeśli się zmieniła
    if (this.pages().length !== pageCount) {
      const newPages = Array(pageCount).fill('');
      this.pages.set(newPages);
      this.pagesChange.emit(pageCount);
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
    this.updateFormattingState(); // Aktualizuj też stan formatowania
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
      case 'justifyLeft':
        document.execCommand('justifyLeft', false);
        break;
      case 'alignCenter':
      case 'justifyCenter':
        document.execCommand('justifyCenter', false);
        break;
      case 'alignRight':
      case 'justifyRight':
        document.execCommand('justifyRight', false);
        break;
      case 'alignJustify':
      case 'justifyFull':
        document.execCommand('justifyFull', false);
        break;
      case 'indent':
        document.execCommand('indent', false);
        break;
      case 'outdent':
        document.execCommand('outdent', false);
        break;
      case 'bulletList':
      case 'insertUnorderedList':
        document.execCommand('insertUnorderedList', false);
        break;
      case 'numberedList':
      case 'insertOrderedList':
        document.execCommand('insertOrderedList', false);
        break;
      case 'removeFormat':
        document.execCommand('removeFormat', false);
        break;
      case 'selectAll':
        document.execCommand('selectAll', false);
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

    this.currentFontSize = size;
    editor.focus();
    
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      // Brak selekcji - ustaw dla następnego tekstu
      this.pendingFontSize = size;
      return;
    }

    const range = selection.getRangeAt(0);
    
    if (range.collapsed) {
      // Kursor bez zaznaczenia - wstaw pusty span z rozmiarem dla kolejnego tekstu
      this.pendingFontSize = size;
      
      // Wstaw zero-width space w span z odpowiednim rozmiarem
      const span = document.createElement('span');
      span.style.fontSize = `${size}pt`;
      span.innerHTML = '\u200B'; // Zero-width space
      
      range.insertNode(span);
      
      // Ustaw kursor wewnątrz spana
      const newRange = document.createRange();
      newRange.setStart(span.firstChild!, 1);
      newRange.setEnd(span.firstChild!, 1);
      selection.removeAllRanges();
      selection.addRange(newRange);
      
      return;
    }

    // Jest zaznaczenie - zastosuj rozmiar do zaznaczonego tekstu
    this.applyFontSizeToSelection(size, selection, range);
    this.onContentChange();
  }

  /**
   * Aplikuje rozmiar czcionki do zaznaczenia - bez execCommand
   */
  private applyFontSizeToSelection(size: number, selection: Selection, range: Range): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    // Wyodrębnij zawartość zaznaczenia
    const fragment = range.extractContents();
    
    // Funkcja pomocnicza do rekurencyjnego przetwarzania węzłów
    const processNode = (node: Node): Node => {
      if (node.nodeType === Node.TEXT_NODE) {
        // Węzeł tekstowy - opakuj w span z nowym rozmiarem
        const span = document.createElement('span');
        span.style.fontSize = `${size}pt`;
        span.textContent = node.textContent;
        return span;
      }
      
      if (node.nodeType === Node.ELEMENT_NODE) {
        const element = node as HTMLElement;
        
        // Jeśli to span lub font - nadpisz font-size i zachowaj inne style
        if (element.tagName === 'SPAN' || element.tagName === 'FONT') {
          const newSpan = document.createElement('span');
          
          // Skopiuj wszystkie style oprócz font-size
          if (element.style.cssText) {
            newSpan.style.cssText = element.style.cssText;
          }
          // Nadpisz font-size
          newSpan.style.fontSize = `${size}pt`;
          
          // Kopiuj inne atrybuty font (face -> font-family, color)
          if (element.tagName === 'FONT') {
            const fontEl = element as HTMLFontElement;
            if (fontEl.face) {
              newSpan.style.fontFamily = fontEl.face;
            }
            if (fontEl.color) {
              newSpan.style.color = fontEl.color;
            }
          }
          
          // Przetwórz dzieci
          Array.from(element.childNodes).forEach(child => {
            newSpan.appendChild(processNode(child));
          });
          
          return newSpan;
        }
        
        // Dla innych elementów (b, i, u, etc.) - zachowaj i przetwórz dzieci
        const clone = element.cloneNode(false) as HTMLElement;
        Array.from(element.childNodes).forEach(child => {
          clone.appendChild(processNode(child));
        });
        return clone;
      }
      
      return node.cloneNode(true);
    };
    
    // Przetwórz fragment - zbierz węzły do późniejszego zaznaczenia
    const newFragment = document.createDocumentFragment();
    const insertedNodes: Node[] = [];
    Array.from(fragment.childNodes).forEach(child => {
      const processed = processNode(child);
      insertedNodes.push(processed);
      newFragment.appendChild(processed);
    });
    
    // Wstaw przetworzony fragment
    range.insertNode(newFragment);
    
    // Przywróć zaznaczenie na wstawionej zawartości
    if (insertedNodes.length > 0) {
      const newRange = document.createRange();
      const firstNode = insertedNodes[0];
      const lastNode = insertedNodes[insertedNodes.length - 1];
      
      newRange.setStartBefore(firstNode);
      newRange.setEndAfter(lastNode);
      
      selection.removeAllRanges();
      selection.addRange(newRange);
    }
    
    // Normalizuj edytor (połącz sąsiadujące węzły tekstowe)
    editor.normalize();
  }

  /**
   * Ustawia rodzinę czcionki
   */
  setFontFamily(fontFamily: string): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    this.currentFontFamily = fontFamily;
    editor.focus();
    
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      this.pendingFontFamily = fontFamily;
      return;
    }

    const range = selection.getRangeAt(0);
    
    if (range.collapsed) {
      // Kursor bez zaznaczenia - wstaw pusty span
      this.pendingFontFamily = fontFamily;
      
      const span = document.createElement('span');
      span.style.fontFamily = fontFamily;
      span.innerHTML = '\u200B';
      
      range.insertNode(span);
      
      const newRange = document.createRange();
      newRange.setStart(span.firstChild!, 1);
      newRange.setEnd(span.firstChild!, 1);
      selection.removeAllRanges();
      selection.addRange(newRange);
      
      return;
    }

    // Użyj tej samej logiki co dla font-size
    this.applyFontFamilyToSelection(fontFamily, selection, range);
    this.onContentChange();
  }

  /**
   * Aplikuje rodzinę czcionki do zaznaczenia
   */
  private applyFontFamilyToSelection(fontFamily: string, selection: Selection, range: Range): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    // Wyodrębnij zawartość zaznaczenia
    const fragment = range.extractContents();
    
    // Funkcja pomocnicza do rekurencyjnego przetwarzania węzłów
    const processNode = (node: Node): Node => {
      if (node.nodeType === Node.TEXT_NODE) {
        const span = document.createElement('span');
        span.style.fontFamily = fontFamily;
        span.textContent = node.textContent;
        return span;
      }
      
      if (node.nodeType === Node.ELEMENT_NODE) {
        const element = node as HTMLElement;
        
        if (element.tagName === 'SPAN' || element.tagName === 'FONT') {
          const newSpan = document.createElement('span');
          
          if (element.style.cssText) {
            newSpan.style.cssText = element.style.cssText;
          }
          newSpan.style.fontFamily = fontFamily;
          
          if (element.tagName === 'FONT') {
            const fontEl = element as HTMLFontElement;
            if (fontEl.size) {
              const sizeMap: Record<string, number> = {
                '1': 8, '2': 10, '3': 12, '4': 14, '5': 18, '6': 24, '7': 36
              };
              newSpan.style.fontSize = `${sizeMap[fontEl.size] || 11}pt`;
            }
            if (fontEl.color) {
              newSpan.style.color = fontEl.color;
            }
          }
          
          Array.from(element.childNodes).forEach(child => {
            newSpan.appendChild(processNode(child));
          });
          
          return newSpan;
        }
        
        const clone = element.cloneNode(false) as HTMLElement;
        Array.from(element.childNodes).forEach(child => {
          clone.appendChild(processNode(child));
        });
        return clone;
      }
      
      return node.cloneNode(true);
    };
    
    const newFragment = document.createDocumentFragment();
    const insertedNodes: Node[] = [];
    Array.from(fragment.childNodes).forEach(child => {
      const processed = processNode(child);
      insertedNodes.push(processed);
      newFragment.appendChild(processed);
    });
    
    range.insertNode(newFragment);
    
    // Przywróć zaznaczenie
    if (insertedNodes.length > 0) {
      const newRange = document.createRange();
      newRange.setStartBefore(insertedNodes[0]);
      newRange.setEndAfter(insertedNodes[insertedNodes.length - 1]);
      selection.removeAllRanges();
      selection.addRange(newRange);
    }
    
    editor.normalize();
  }

  /**
   * Ustawia kolor tekstu
   */
  setTextColor(color: string): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    editor.focus();
    
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;
    
    const range = selection.getRangeAt(0);
    if (range.collapsed) return;
    
    this.applyColorToSelection(color, selection, range);
    this.onContentChange();
  }

  /**
   * Aplikuje kolor do zaznaczenia
   */
  private applyColorToSelection(color: string, selection: Selection, range: Range): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    const fragment = range.extractContents();
    
    const processNode = (node: Node): Node => {
      if (node.nodeType === Node.TEXT_NODE) {
        const span = document.createElement('span');
        span.style.color = color;
        span.textContent = node.textContent;
        return span;
      }
      
      if (node.nodeType === Node.ELEMENT_NODE) {
        const element = node as HTMLElement;
        
        if (element.tagName === 'SPAN' || element.tagName === 'FONT') {
          const newSpan = document.createElement('span');
          
          if (element.style.cssText) {
            newSpan.style.cssText = element.style.cssText;
          }
          newSpan.style.color = color;
          
          if (element.tagName === 'FONT') {
            const fontEl = element as HTMLFontElement;
            if (fontEl.size) {
              const sizeMap: Record<string, number> = {
                '1': 8, '2': 10, '3': 12, '4': 14, '5': 18, '6': 24, '7': 36
              };
              newSpan.style.fontSize = `${sizeMap[fontEl.size] || 11}pt`;
            }
            if (fontEl.face) {
              newSpan.style.fontFamily = fontEl.face;
            }
          }
          
          Array.from(element.childNodes).forEach(child => {
            newSpan.appendChild(processNode(child));
          });
          
          return newSpan;
        }
        
        const clone = element.cloneNode(false) as HTMLElement;
        Array.from(element.childNodes).forEach(child => {
          clone.appendChild(processNode(child));
        });
        return clone;
      }
      
      return node.cloneNode(true);
    };
    
    const newFragment = document.createDocumentFragment();
    const insertedNodes: Node[] = [];
    Array.from(fragment.childNodes).forEach(child => {
      const processed = processNode(child);
      insertedNodes.push(processed);
      newFragment.appendChild(processed);
    });
    
    range.insertNode(newFragment);
    
    // Przywróć zaznaczenie
    if (insertedNodes.length > 0) {
      const newRange = document.createRange();
      newRange.setStartBefore(insertedNodes[0]);
      newRange.setEndAfter(insertedNodes[insertedNodes.length - 1]);
      selection.removeAllRanges();
      selection.addRange(newRange);
    }
    
    editor.normalize();
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

  // Zapisana selekcja (do użycia gdy selekcja jest tracona przez kliknięcie na toolbar)
  private savedSelection: Range | null = null;

  /**
   * Zapisuje aktualną selekcję - wywoływane przed focusout
   */
  saveSelection(): void {
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      const editor = this.editorContent?.nativeElement;
      if (editor && editor.contains(range.commonAncestorContainer)) {
        this.savedSelection = range.cloneRange();
        console.log('[saveSelection] Zapisano selekcję:', range.toString());
      }
    }
  }

  /**
   * Przywraca zapisaną selekcję
   */
  restoreSelection(): boolean {
    if (!this.savedSelection) {
      console.warn('[restoreSelection] Brak zapisanej selekcji');
      return false;
    }
    
    const selection = window.getSelection();
    if (selection) {
      selection.removeAllRanges();
      selection.addRange(this.savedSelection);
      console.log('[restoreSelection] Przywrócono selekcję:', this.savedSelection.toString());
      return true;
    }
    return false;
  }

  /**
   * Aplikuje styl dokumentu TYLKO do zaznaczonego fragmentu tekstu.
   * Jeśli nic nie jest zaznaczone, nie robi nic.
   * Działa jak formatowanie w Word - aplikuje czcionkę, rozmiar, kolor, bold, italic, underline do selekcji.
   */
  applyDocumentStyle(style: {
    id: string;
    name: string;
    fontFamily?: string;
    fontSize?: number;
    color?: string;
    isBold?: boolean;
    isItalic?: boolean;
    isUnderline?: boolean;
    alignment?: string;
    outlineLevel?: number;
  }): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) {
      console.warn('[applyDocumentStyle] Brak elementu edytora');
      return;
    }

    // Debug - sprawdź co przychodzi w stylu
    console.log('[applyDocumentStyle] Otrzymany styl:', JSON.stringify(style, null, 2));

    // Najpierw spróbuj przywrócić zapisaną selekcję (bo mogła być utracona przez kliknięcie na toolbar)
    let selection = window.getSelection();
    let range: Range | null = null;
    
    if (selection && selection.rangeCount > 0) {
      range = selection.getRangeAt(0);
      // Sprawdź czy selekcja jest w edytorze i nie jest pusta
      if (!editor.contains(range.commonAncestorContainer) || range.collapsed) {
        range = null;
      }
    }
    
    // Jeśli nie ma aktualnej selekcji, spróbuj użyć zapisanej
    if (!range && this.savedSelection) {
      console.log('[applyDocumentStyle] Używam zapisanej selekcji');
      this.restoreSelection();
      selection = window.getSelection();
      if (selection && selection.rangeCount > 0) {
        range = selection.getRangeAt(0);
      }
    }
    
    if (!range || range.collapsed) {
      console.warn('[applyDocumentStyle] Brak zaznaczonego tekstu');
      return;
    }

    // Sprawdź czy selekcja jest wewnątrz edytora
    if (!editor.contains(range.commonAncestorContainer)) {
      console.warn('[applyDocumentStyle] Selekcja poza edytorem');
      return;
    }

    editor.focus();

    // Pobierz zaznaczony tekst
    const selectedText = range.toString();
    if (!selectedText || selectedText.trim().length === 0) {
      console.warn('[applyDocumentStyle] Pusty zaznaczony tekst');
      return;
    }

    console.log('[applyDocumentStyle] Zaznaczony tekst:', selectedText);

    // Wyodrębnij zawartość zaznaczenia
    const fragment = range.extractContents();
    
    // Spłaszcz zagnieżdżone elementy - wyciągnij tylko tekst
    const flattenFragment = (node: Node): string => {
      if (node.nodeType === Node.TEXT_NODE) {
        return node.textContent || '';
      }
      let text = '';
      node.childNodes.forEach(child => {
        text += flattenFragment(child);
      });
      return text;
    };
    const plainText = flattenFragment(fragment);
    
    // Tworzę element SPAN z wszystkimi stylami
    const styledSpan = document.createElement('span');
    
    // Buduj style inline
    const styles: string[] = [];
    
    if (style.fontFamily) {
      styles.push(`font-family: "${style.fontFamily}"`);
      console.log('[applyDocumentStyle] Ustawiam font-family:', style.fontFamily);
    }
    
    if (style.fontSize) {
      styles.push(`font-size: ${style.fontSize}pt`);
      console.log('[applyDocumentStyle] Ustawiam font-size:', style.fontSize + 'pt');
    }
    
    if (style.color) {
      styles.push(`color: ${style.color}`);
      console.log('[applyDocumentStyle] Ustawiam color:', style.color);
    }
    
    if (style.isBold === true) {
      styles.push('font-weight: bold');
      console.log('[applyDocumentStyle] Ustawiam bold');
    } else if (style.isBold === false) {
      styles.push('font-weight: normal');
    }
    
    if (style.isItalic === true) {
      styles.push('font-style: italic');
      console.log('[applyDocumentStyle] Ustawiam italic');
    } else if (style.isItalic === false) {
      styles.push('font-style: normal');
    }
    
    if (style.isUnderline === true) {
      styles.push('text-decoration: underline');
      console.log('[applyDocumentStyle] Ustawiam underline');
    } else if (style.isUnderline === false) {
      styles.push('text-decoration: none');
    }
    
    // Zastosuj style do span
    if (styles.length > 0) {
      styledSpan.setAttribute('style', styles.join('; '));
      console.log('[applyDocumentStyle] Finalne style:', styles.join('; '));
    }
    
    // Wstaw czysty tekst do span (bez zagnieżdżonych elementów)
    styledSpan.textContent = plainText;
    
    // Wstaw span w miejsce zaznaczenia
    range.insertNode(styledSpan);
    
    // Ustaw kursor na końcu wstawionego elementu
    const newRange = document.createRange();
    newRange.selectNodeContents(styledSpan);
    newRange.collapse(false);
    selection!.removeAllRanges();
    selection!.addRange(newRange);
    
    // Wyczyść zapisaną selekcję
    this.savedSelection = null;

    console.log('[applyDocumentStyle] Styl został zastosowany');
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

    // Pobierz rzeczywisty rozmiar i czcionkę z computed styles
    const selection = window.getSelection();
    let fontSize = 11;
    let fontFamily = 'Calibri';
    let textColor = '#000000';
    let currentBlockFormat = 'p';

    if (selection && selection.rangeCount > 0) {
      let element = selection.anchorNode as HTMLElement;
      if (element?.nodeType === Node.TEXT_NODE) {
        element = element.parentElement!;
      }

      if (element) {
        const computedStyle = window.getComputedStyle(element);
        
        // Rozmiar czcionki - konwersja px na pt
        const fontSizePx = parseFloat(computedStyle.fontSize);
        fontSize = Math.round(fontSizePx * 0.75); // px to pt (96dpi / 72pt)
        
        // Czcionka - usuń cudzysłowy i weź pierwszą
        fontFamily = computedStyle.fontFamily.replace(/['"]/g, '').split(',')[0].trim();
        
        // Kolor tekstu
        textColor = this.rgbToHex(computedStyle.color);

        // Znajdź blok nadrzędny (p, h1, h2, etc.)
        let blockElement = element;
        while (blockElement && !['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'DIV', 'LI'].includes(blockElement.tagName)) {
          blockElement = blockElement.parentElement!;
        }
        if (blockElement) {
          currentBlockFormat = blockElement.tagName.toLowerCase();
        }
      }
    }

    // Debug log
    console.log('[updateFormattingState] fontSize:', fontSize, 'fontFamily:', fontFamily, 'blockFormat:', currentBlockFormat);

    this.editorState.update(state => ({
      ...state,
      currentFormatting: formatting,
      currentStyle: {
        fontFamily: fontFamily,
        fontSize: fontSize,
        textColor: textColor,
        blockFormat: currentBlockFormat
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

  /**
   * Pobiera aktualne formatowanie z zaznaczenia
   */
  getCurrentFormatting(): any {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      return null;
    }

    // Pobierz element z którego kopiujemy formatowanie
    let element = selection.anchorNode as HTMLElement;
    if (element?.nodeType === Node.TEXT_NODE) {
      element = element.parentElement!;
    }

    if (!element) return null;

    const computedStyle = window.getComputedStyle(element);
    
    return {
      bold: document.queryCommandState('bold'),
      italic: document.queryCommandState('italic'),
      underline: document.queryCommandState('underline'),
      strikethrough: document.queryCommandState('strikeThrough'),
      subscript: document.queryCommandState('subscript'),
      superscript: document.queryCommandState('superscript'),
      fontFamily: computedStyle.fontFamily.replace(/['"]/g, '').split(',')[0].trim(),
      fontSize: parseInt(computedStyle.fontSize),
      textColor: this.rgbToHex(computedStyle.color),
      backgroundColor: computedStyle.backgroundColor === 'rgba(0, 0, 0, 0)' ? '' : this.rgbToHex(computedStyle.backgroundColor)
    };
  }

  /**
   * Aplikuje formatowanie do zaznaczenia
   */
  applyFormatting(format: any): void {
    if (!format) return;

    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0 || selection.isCollapsed) {
      return;
    }

    // Aplikuj formatowanie tekstu
    if (format.bold) document.execCommand('bold', false);
    if (format.italic) document.execCommand('italic', false);
    if (format.underline) document.execCommand('underline', false);
    if (format.strikethrough) document.execCommand('strikeThrough', false);
    if (format.subscript) document.execCommand('subscript', false);
    if (format.superscript) document.execCommand('superscript', false);

    // Aplikuj czcionkę i rozmiar
    if (format.fontFamily) {
      document.execCommand('fontName', false, format.fontFamily);
    }
    if (format.fontSize) {
      // Zawijamy zaznaczenie w span z rozmiarem czcionki
      const range = selection.getRangeAt(0);
      const span = document.createElement('span');
      span.style.fontSize = format.fontSize + 'px';
      
      try {
        const contents = range.extractContents();
        span.appendChild(contents);
        range.insertNode(span);
        selection.selectAllChildren(span);
      } catch (e) {
        // Fallback dla złożonych zakresów
        document.execCommand('fontSize', false, '7');
        const fontElements = this.editorContent?.nativeElement.querySelectorAll('font[size="7"]');
        fontElements?.forEach((el: Element) => {
          (el as HTMLElement).removeAttribute('size');
          (el as HTMLElement).style.fontSize = format.fontSize + 'px';
        });
      }
    }

    // Aplikuj kolory
    if (format.textColor) {
      document.execCommand('foreColor', false, format.textColor);
    }
    if (format.backgroundColor) {
      document.execCommand('hiliteColor', false, format.backgroundColor);
    }

    // Emituj zmiany
    const html = this.editorContent?.nativeElement?.innerHTML || '';
    this._content.set(html);
    this.contentChange.emit(html);
    this.updateFormattingState();
  }

  /**
   * Konwertuje RGB na HEX
   */
  private rgbToHex(rgb: string): string {
    if (rgb.startsWith('#')) return rgb;
    if (rgb === 'transparent' || rgb === 'rgba(0, 0, 0, 0)') return '';
    
    const match = rgb.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    if (!match) return '#000000';
    
    const r = parseInt(match[1]).toString(16).padStart(2, '0');
    const g = parseInt(match[2]).toString(16).padStart(2, '0');
    const b = parseInt(match[3]).toString(16).padStart(2, '0');
    
    return `#${r}${g}${b}`;
  }

  // ================================
  // Nagłówek i Stopka
  // ================================

  /**
   * Rozpoczyna edycję nagłówka
   */
  startEditingHeader(): void {
    this.editingSection.set('header');
    // Ustaw focus na edytorze nagłówka po renderowaniu
    setTimeout(() => {
      if (this.headerContentEl?.nativeElement) {
        this.headerContentEl.nativeElement.innerHTML = this._headerHtml();
        this.headerContentEl.nativeElement.focus();
      }
    }, 0);
  }

  /**
   * Rozpoczyna edycję stopki
   */
  startEditingFooter(): void {
    this.editingSection.set('footer');
    setTimeout(() => {
      if (this.footerContentEl?.nativeElement) {
        this.footerContentEl.nativeElement.innerHTML = this._footerHtml();
        this.footerContentEl.nativeElement.focus();
      }
    }, 0);
  }

  /**
   * Kończy edycję nagłówka/stopki i wraca do głównej treści
   */
  stopEditingHeaderFooter(): void {
    if (this.editingSection() !== 'body') {
      this.editingSection.set('body');
    }
  }

  /**
   * Obsługa blur nagłówka
   */
  onHeaderBlur(): void {
    const content = this.headerContentEl?.nativeElement?.innerHTML || '';
    this._headerHtml.set(content);
    this.headerChange.emit({
      html: content,
      height: this._headerHeight()
    });
    // Nie kończymy edycji od razu, pozwalamy na kliknięcie poza nagłówkiem
  }

  /**
   * Obsługa input nagłówka
   */
  onHeaderInput(event: Event): void {
    const content = (event.target as HTMLDivElement).innerHTML;
    this._headerHtml.set(content);
  }

  /**
   * Obsługa blur stopki
   */
  onFooterBlur(): void {
    const content = this.footerContentEl?.nativeElement?.innerHTML || '';
    this._footerHtml.set(content);
    this.footerChange.emit({
      html: content,
      height: this._footerHeight()
    });
  }

  /**
   * Obsługa input stopki
   */
  onFooterInput(event: Event): void {
    const content = (event.target as HTMLDivElement).innerHTML;
    this._footerHtml.set(content);
  }

  /**
   * Pobiera zawartość nagłówka dla danej strony
   */
  getHeaderContent(pageIndex: number): string {
    // Pierwsza strona z inną treścią
    if (this._differentFirstPage() && pageIndex === 0) {
      return this._headerFirstPageHtml();
    }
    // Różne parzyste/nieparzyste
    if (this._differentOddEven()) {
      const isOdd = (pageIndex + 1) % 2 === 1;
      return isOdd ? this._headerOddHtml() : this._headerEvenHtml();
    }
    return this._headerHtml();
  }

  /**
   * Pobiera zawartość stopki dla danej strony
   */
  getFooterContent(pageIndex: number): string {
    let content: string;
    // Pierwsza strona z inną treścią
    if (this._differentFirstPage() && pageIndex === 0) {
      content = this._footerFirstPageHtml();
    } 
    // Różne parzyste/nieparzyste
    else if (this._differentOddEven()) {
      const isOdd = (pageIndex + 1) % 2 === 1;
      content = isOdd ? this._footerOddHtml() : this._footerEvenHtml();
    } else {
      content = this._footerHtml();
    }
    // Zamień placeholder na numer strony
    content = content.replace(/\{page\}/gi, String(pageIndex + 1));
    content = content.replace(/\{pages\}/gi, String(this.pages().length));
    return content;
  }

  /**
   * Oblicza dostępną wysokość dla treści głównej (bez nagłówka i stopki)
   */
  getContentAreaHeight(): number {
    const pageHeight = this.pageOrientation === 'landscape' ? 816 : 1122;
    const headerHeightPx = this._headerHeight() * 37.8;
    const footerHeightPx = this._footerHeight() * 37.8;
    return pageHeight - headerHeightPx - footerHeightPx;
  }

  /**
   * Ustawia wysokość nagłówka
   */
  setHeaderHeight(heightCm: number): void {
    this._headerHeight.set(Math.max(0.5, Math.min(5, heightCm)));
    this.headerChange.emit({
      html: this._headerHtml(),
      height: this._headerHeight()
    });
  }

  /**
   * Ustawia wysokość stopki
   */
  setFooterHeight(heightCm: number): void {
    this._footerHeight.set(Math.max(0.5, Math.min(5, heightCm)));
    this.footerChange.emit({
      html: this._footerHtml(),
      height: this._footerHeight()
    });
  }

  /**
   * Pobiera pełną zawartość dokumentu z nagłówkiem i stopką
   */
  getFullDocumentContent(): { body: string; header: HeaderFooterContent; footer: HeaderFooterContent } {
    return {
      body: this.editorContent?.nativeElement?.innerHTML || '',
      header: {
        html: this._headerHtml(),
        height: this._headerHeight(),
        differentFirstPage: this._differentFirstPage(),
        firstPageHtml: this._headerFirstPageHtml()
      },
      footer: {
        html: this._footerHtml(),
        height: this._footerHeight(),
        differentFirstPage: this._differentFirstPage(),
        firstPageHtml: this._footerFirstPageHtml()
      }
    };
  }

  // ================================
  // Menu Opcji Nagłówka/Stopki (styl Google Docs)
  // ================================

  /**
   * Toggle menu opcji nagłówka
   */
  toggleHeaderOptionsMenu(event: Event): void {
    event.stopPropagation();
    this.showHeaderOptionsMenu.update(v => !v);
    this.showFooterOptionsMenu.set(false);
    
    if (this.showHeaderOptionsMenu()) {
      // Zamknij menu po kliknięciu poza nim
      setTimeout(() => {
        const closeHandler = () => {
          this.showHeaderOptionsMenu.set(false);
          document.removeEventListener('click', closeHandler);
        };
        document.addEventListener('click', closeHandler);
      }, 0);
    }
  }

  /**
   * Toggle menu opcji stopki
   */
  toggleFooterOptionsMenu(event: Event): void {
    event.stopPropagation();
    this.showFooterOptionsMenu.update(v => !v);
    this.showHeaderOptionsMenu.set(false);
    
    if (this.showFooterOptionsMenu()) {
      setTimeout(() => {
        const closeHandler = () => {
          this.showFooterOptionsMenu.set(false);
          document.removeEventListener('click', closeHandler);
        };
        document.addEventListener('click', closeHandler);
      }, 0);
    }
  }

  /**
   * Toggle "Inna pierwsza strona"
   */
  toggleDifferentFirstPage(): void {
    this._differentFirstPage.update(v => !v);
    this.emitHeaderFooterChanges();
  }

  /**
   * Otwiera dialog formatu nagłówka
   */
  openHeaderFormatDialog(): void {
    this.showHeaderOptionsMenu.set(false);
    this.openHeaderFooterFormatDialog();
  }

  /**
   * Otwiera dialog formatu stopki
   */
  openFooterFormatDialog(): void {
    this.showFooterOptionsMenu.set(false);
    this.openHeaderFooterFormatDialog();
  }

  /**
   * Otwiera dialog formatowania nagłówka i stopki
   */
  openHeaderFooterFormatDialog(): void {
    // Emituj event do rodzica - dialog zostanie wyświetlony w document-editor
    this.openHeaderFooterSettings.emit({
      headerMargin: this._headerHeight(),
      footerMargin: this._footerHeight(),
      differentFirstPage: this._differentFirstPage(),
      differentOddEven: this._differentOddEven()
    });
  }

  /**
   * Aplikuje ustawienia nagłówka/stopki z zewnątrz (z document-editor)
   */
  applyHeaderFooterSettings(settings: {
    headerMargin: number;
    footerMargin: number;
    differentFirstPage: boolean;
    differentOddEven: boolean;
  }): void {
    this._headerHeight.set(settings.headerMargin);
    this._footerHeight.set(settings.footerMargin);
    this._differentFirstPage.set(settings.differentFirstPage);
    this._differentOddEven.set(settings.differentOddEven);
    
    this.emitHeaderFooterChanges();
  }

  /**
   * Wstawia numer strony do nagłówka
   */
  insertPageNumbers(): void {
    this.showHeaderOptionsMenu.set(false);
    
    if (this.headerContentEl?.nativeElement) {
      const pageNumber = '<span class="page-number">{page}</span>';
      this.headerContentEl.nativeElement.focus();
      document.execCommand('insertHTML', false, pageNumber);
      this.onHeaderInput({ target: this.headerContentEl.nativeElement } as any);
    }
  }

  /**
   * Wstawia numer strony do stopki
   */
  insertPageNumbersFooter(): void {
    this.showFooterOptionsMenu.set(false);
    
    if (this.footerContentEl?.nativeElement) {
      const pageNumber = '<span class="page-number">{page}</span>';
      this.footerContentEl.nativeElement.focus();
      document.execCommand('insertHTML', false, pageNumber);
      this.onFooterInput({ target: this.footerContentEl.nativeElement } as any);
    }
  }

  /**
   * Usuwa nagłówek
   */
  removeHeader(): void {
    this.showHeaderOptionsMenu.set(false);
    this._headerHtml.set('');
    this._headerFirstPageHtml.set('');
    if (this.headerContentEl?.nativeElement) {
      this.headerContentEl.nativeElement.innerHTML = '';
    }
    this.emitHeaderFooterChanges();
    this.stopEditingHeaderFooter();
  }

  /**
   * Usuwa stopkę
   */
  removeFooter(): void {
    this.showFooterOptionsMenu.set(false);
    this._footerHtml.set('');
    this._footerFirstPageHtml.set('');
    if (this.footerContentEl?.nativeElement) {
      this.footerContentEl.nativeElement.innerHTML = '';
    }
    this.emitHeaderFooterChanges();
    this.stopEditingHeaderFooter();
  }

  /**
   * Emituje zmiany nagłówka i stopki
   */
  private emitHeaderFooterChanges(): void {
    this.headerChange.emit({
      html: this._headerHtml(),
      height: this._headerHeight(),
      differentFirstPage: this._differentFirstPage(),
      firstPageHtml: this._headerFirstPageHtml(),
      differentOddEven: this._differentOddEven(),
      oddHtml: this._headerOddHtml(),
      evenHtml: this._headerEvenHtml()
    });
    this.footerChange.emit({
      html: this._footerHtml(),
      height: this._footerHeight(),
      differentFirstPage: this._differentFirstPage(),
      firstPageHtml: this._footerFirstPageHtml(),
      differentOddEven: this._differentOddEven(),
      oddHtml: this._footerOddHtml(),
      evenHtml: this._footerEvenHtml()
    });
  }
}
