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
  QueryList,
  ViewEncapsulation
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
  selector: 'd2-wysiwyg-editor',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './wysiwyg-editor.html',
  styleUrl: './wysiwyg-editor.scss',
  encapsulation: ViewEncapsulation.None
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
  private selectedImageWrapper: HTMLElement | null = null;
  private draggedImageWrapper: HTMLElement | null = null;
  private imageResizeState: {
    wrapper: HTMLElement;
    startX: number;
    startY: number;
    startWidth: number;
    startHeight: number;
    axis: 'x' | 'y' | 'both';
  } | null = null;

  // Stan resize tabeli
  private tableResizeState: {
    type: 'col' | 'row' | 'table';
    table: HTMLTableElement;
    startX: number;
    startY: number;
    colIndex: number;
    rowIndex: number;
    startWidths: number[];
    startHeight: number;
    startTableWidth: number;
  } | null = null;

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

  // Computed: zawartość nagłówka/stopki per strona (reaktywna na zmiany sygnałów)
  headerContents = computed(() => {
    const pagesArr = this.pages();
    return pagesArr.map((_, i) => this._computeHeaderContent(i));
  });
  footerContents = computed(() => {
    const pagesArr = this.pages();
    return pagesArr.map((_, i) => this._computeFooterContent(i));
  });

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

    // Dla pustego dokumentu zawsze pokazuj 1 stronę
    const plainText = (editor.textContent || '').replace(/\u00A0/g, '').trim();
    const hasMedia = editor.querySelector('img, table, hr, .page-break') !== null;
    if (!plainText && !hasMedia) {
      if (this.pages().length !== 1) {
        this.pages.set(['']);
        this.pagesChange.emit(1);
      }
      return;
    }
    
    const contentHeight = editor.scrollHeight;
    const marginTop = this.pageMargins.top * 37.8;
    const marginBottom = this.pageMargins.bottom * 37.8;
    const availableHeight = this.PAGE_HEIGHT_PX - marginTop - marginBottom;

    // Tolerancja na różnice renderowania (1-4px), które potrafią sztucznie dodać stronę
    const adjustedHeight = Math.max(0, contentHeight - 4);
    const pageCount = Math.max(1, Math.ceil(adjustedHeight / availableHeight));
    
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

    editor.innerHTML = this._content() || '<p></p>';
    this.wrapExistingImages();
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

    // Kliknięcia w obrazy (zaznaczanie)
    editor.addEventListener('click', (e) => {
      this.handleEditorClick(e);
    });

    // Resize obrazów
    editor.addEventListener('mousedown', (e) => {
      this.handleEditorMouseDown(e);
    });

    // Drag&drop obrazów wewnątrz dokumentu
    editor.addEventListener('dragstart', (e) => {
      this.handleEditorDragStart(e);
    });

    editor.addEventListener('dragover', (e) => {
      this.handleEditorDragOver(e);
    });

    editor.addEventListener('dragend', () => {
      this.draggedImageWrapper = null;
    });

    // Kursor resize tabel (wykrywanie krawędzi kolumn/wierszy)
    editor.addEventListener('mousemove', (e) => {
      this.handleTableResizeCursor(e);
    });
  }

  private handleEditorClick(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    const imageWrapper = target.closest('.editor-image-wrapper') as HTMLElement | null;

    if (imageWrapper) {
      // Kliknięcie na wrapper - nie rób nic, selekcja jest obsługiwana w mousedown
      return;
    }

    this.clearSelectedImage();
  }

  private handleEditorMouseDown(event: MouseEvent): void {
    const target = event.target as HTMLElement;

    // --- Resize tabeli ---
    const tableHit = this.detectTableResizeHit(event);
    if (tableHit) {
      event.preventDefault();
      event.stopPropagation();
      this.startTableResize(tableHit, event);
      return;
    }

    // Sprawdź czy kliknięto na wrapper obrazu lub jego zawartość
    const wrapper = target.closest('.editor-image-wrapper') as HTMLElement | null;
    if (!wrapper) {
      return;
    }

    // Resize handle
    if (target.classList.contains('image-resize-handle')) {
      event.preventDefault();
      event.stopPropagation();

      this.selectImageWrapper(wrapper);
      wrapper.setAttribute('draggable', 'false');

      const rect = wrapper.getBoundingClientRect();
      const img = wrapper.querySelector('img') as HTMLImageElement | null;
      let axis: 'x' | 'y' | 'both' = 'both';
      if (target.classList.contains('resize-handle-right')) {
        axis = 'x';
      } else if (target.classList.contains('resize-handle-bottom')) {
        axis = 'y';
      }

      this.imageResizeState = {
        wrapper,
        startX: event.clientX,
        startY: event.clientY,
        startWidth: rect.width,
        startHeight: rect.height,
        axis
      };

      const onMouseMove = (moveEvent: MouseEvent) => {
        if (!this.imageResizeState) return;

        const editor = this.editorContent?.nativeElement;
        const editorMaxWidth = (editor?.clientWidth || 900) - 30;
        const st = this.imageResizeState;
        const currentImg = st.wrapper.querySelector('img') as HTMLImageElement | null;
        if (!currentImg) return;

        if (st.axis === 'x') {
          // Rozciąganie tylko szerokości — wysokość zostaje stała
          const deltaX = moveEvent.clientX - st.startX;
          const newWidth = Math.max(60, Math.min(editorMaxWidth, st.startWidth + deltaX));
          st.wrapper.style.width = `${newWidth}px`;
          st.wrapper.style.maxWidth = '100%';
          currentImg.style.width = '100%';
          currentImg.style.height = `${st.startHeight}px`;
        } else if (st.axis === 'y') {
          // Rozciąganie tylko wysokości — szerokość zostaje stała
          const deltaY = moveEvent.clientY - st.startY;
          const newHeight = Math.max(30, st.startHeight + deltaY);
          st.wrapper.style.width = `${st.startWidth}px`;
          st.wrapper.style.maxWidth = '100%';
          currentImg.style.width = '100%';
          currentImg.style.height = `${newHeight}px`;
        } else {
          // Proporcjonalne skalowanie z narożnika
          const deltaX = moveEvent.clientX - st.startX;
          const newWidth = Math.max(60, Math.min(editorMaxWidth, st.startWidth + deltaX));
          st.wrapper.style.width = `${newWidth}px`;
          st.wrapper.style.maxWidth = '100%';
          currentImg.style.width = '100%';
          currentImg.style.height = 'auto';
        }
      };

      const onMouseUp = () => {
        if (this.imageResizeState?.wrapper) {
          this.imageResizeState.wrapper.setAttribute('draggable', 'true');
        }

        this.imageResizeState = null;
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
        this.onContentChange();
      };

      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
      return;
    }

    // Kliknięcie na wrapper obrazu (nie resize handle) - zaznacz obraz
    event.preventDefault();
    this.selectImageWrapper(wrapper);
  }

  // ======= RESIZE TABEL =======

  private readonly TABLE_EDGE_THRESHOLD = 6; // px od krawędzi

  /**
   * Wykrywa czy kursor jest nad krawędzią kolumny, wiersza lub narożnikiem tabeli
   */
  private detectTableResizeHit(event: MouseEvent): {
    type: 'col' | 'row' | 'table';
    table: HTMLTableElement;
    colIndex: number;
    rowIndex: number;
  } | null {
    const target = event.target as HTMLElement;
    const td = target.closest('td, th') as HTMLTableCellElement | null;
    const table = target.closest('table') as HTMLTableElement | null;

    if (!table) return null;

    const t = this.TABLE_EDGE_THRESHOLD;

    // Sprawdź narożnik tabeli (prawy dolny)
    const tableRect = table.getBoundingClientRect();
    if (
      Math.abs(event.clientX - tableRect.right) < t + 2 &&
      Math.abs(event.clientY - tableRect.bottom) < t + 2
    ) {
      return { type: 'table', table, colIndex: -1, rowIndex: -1 };
    }

    if (!td) return null;
    const cellRect = td.getBoundingClientRect();

    // Prawa krawędź komórki = granica kolumny
    if (Math.abs(event.clientX - cellRect.right) < t) {
      return {
        type: 'col',
        table,
        colIndex: td.cellIndex,
        rowIndex: (td.parentElement as HTMLTableRowElement).rowIndex
      };
    }

    // Lewa krawędź (granica z poprzednią kolumną)
    if (td.cellIndex > 0 && Math.abs(event.clientX - cellRect.left) < t) {
      return {
        type: 'col',
        table,
        colIndex: td.cellIndex - 1,
        rowIndex: (td.parentElement as HTMLTableRowElement).rowIndex
      };
    }

    // Dolna krawędź komórki = granica wiersza
    if (Math.abs(event.clientY - cellRect.bottom) < t) {
      return {
        type: 'row',
        table,
        colIndex: td.cellIndex,
        rowIndex: (td.parentElement as HTMLTableRowElement).rowIndex
      };
    }

    return null;
  }

  /**
   * Zmienia kursor nad krawędziami tabeli
   */
  private handleTableResizeCursor(event: MouseEvent): void {
    // Nie zmieniaj kursora podczas aktywnego resize
    if (this.tableResizeState || this.imageResizeState) return;

    const target = event.target as HTMLElement;
    const td = target.closest('td, th') as HTMLTableCellElement | null;
    const table = target.closest('table') as HTMLTableElement | null;

    if (!table || !td) {
      // Reset kursora na komórkach które miały zmieniony kursor
      if (this._lastCursorCell) {
        this._lastCursorCell.style.cursor = '';
        this._lastCursorCell = null;
      }
      return;
    }

    const hit = this.detectTableResizeHit(event);

    // Wyczyść poprzednią komórkę
    if (this._lastCursorCell && this._lastCursorCell !== td) {
      this._lastCursorCell.style.cursor = '';
    }

    if (hit) {
      if (hit.type === 'col') {
        td.style.cursor = 'col-resize';
      } else if (hit.type === 'row') {
        td.style.cursor = 'row-resize';
      } else {
        td.style.cursor = 'nwse-resize';
      }
      this._lastCursorCell = td;
    } else {
      td.style.cursor = '';
      this._lastCursorCell = null;
    }
  }

  private _lastCursorCell: HTMLElement | null = null;

  /**
   * Normalizuje tabele - ustawia stałe szerokości kolumn jeśli brak
   */
  private ensureTableColWidths(table: HTMLTableElement): void {
    const firstRow = table.rows[0];
    if (!firstRow) return;

    // Sprawdź czy kolumny mają już ustawione szerokości
    const hasWidths = Array.from(firstRow.cells).every(c => !!c.style.width);
    if (hasWidths) return;

    // Zmierz aktualne i ustaw px
    const cells = Array.from(firstRow.cells);
    const widths = cells.map(c => c.getBoundingClientRect().width);
    cells.forEach((c, i) => {
      c.style.width = `${widths[i]}px`;
    });

    // Ustaw table-layout: fixed
    table.style.tableLayout = 'fixed';
  }

  /**
   * Rozpoczyna resize tabeli
   */
  private startTableResize(
    hit: { type: 'col' | 'row' | 'table'; table: HTMLTableElement; colIndex: number; rowIndex: number },
    event: MouseEvent
  ): void {
    const table = hit.table;
    this.ensureTableColWidths(table);

    const firstRow = table.rows[0];
    const startWidths = firstRow ? Array.from(firstRow.cells).map(c => c.getBoundingClientRect().width) : [];
    const startTableWidth = table.getBoundingClientRect().width;

    let startHeight = 0;
    if (hit.type === 'row' && table.rows[hit.rowIndex]) {
      startHeight = table.rows[hit.rowIndex].getBoundingClientRect().height;
    }

    this.tableResizeState = {
      type: hit.type,
      table,
      startX: event.clientX,
      startY: event.clientY,
      colIndex: hit.colIndex,
      rowIndex: hit.rowIndex,
      startWidths,
      startHeight,
      startTableWidth
    };

    // Zablokuj zaznaczanie tekstu i wymusz kursor na całym dokumencie
    document.body.classList.add('table-resizing');
    const cursorType = hit.type === 'col' ? 'col-resize' : hit.type === 'row' ? 'row-resize' : 'nwse-resize';
    document.body.style.cursor = cursorType;

    const onMouseMove = (moveEvent: MouseEvent) => {
      moveEvent.preventDefault();
      if (!this.tableResizeState) return;
      const st = this.tableResizeState;

      if (st.type === 'col') {
        this.resizeTableColumn(st, moveEvent);
      } else if (st.type === 'row') {
        this.resizeTableRow(st, moveEvent);
      } else {
        this.resizeWholeTable(st, moveEvent);
      }
    };

    const onMouseUp = () => {
      this.tableResizeState = null;
      document.body.classList.remove('table-resizing');
      document.body.style.cursor = '';
      if (this._lastCursorCell) {
        this._lastCursorCell.style.cursor = '';
        this._lastCursorCell = null;
      }
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
      this.onContentChange();
    };

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  }

  /**
   * Resize kolumny - przesuwa granicę między kolumnami
   */
  private resizeTableColumn(
    st: NonNullable<typeof this.tableResizeState>,
    moveEvent: MouseEvent
  ): void {
    const deltaX = moveEvent.clientX - st.startX;
    const ci = st.colIndex;
    const totalCols = st.startWidths.length;

    // Rozszerzamy kolumnę ci, zwężamy ci+1 (lub rozszerzamy tabelę jeśli ostatnia)
    const minW = 40;
    const newLeft = Math.max(minW, st.startWidths[ci] + deltaX);

    const allRows = st.table.rows;

    if (ci < totalCols - 1) {
      // Środkowa kolumna - zabierz z sąsiedniej
      const newRight = Math.max(minW, st.startWidths[ci + 1] - deltaX);
      for (let r = 0; r < allRows.length; r++) {
        const cells = allRows[r].cells;
        if (cells[ci]) cells[ci].style.width = `${newLeft}px`;
        if (cells[ci + 1]) cells[ci + 1].style.width = `${newRight}px`;
      }
    } else {
      // Ostatnia kolumna - zmień szerokość tabeli
      for (let r = 0; r < allRows.length; r++) {
        const cells = allRows[r].cells;
        if (cells[ci]) cells[ci].style.width = `${newLeft}px`;
      }
      const totalWidth = st.startWidths.reduce((s, w, i) => s + (i === ci ? newLeft : w), 0);
      st.table.style.width = `${totalWidth}px`;
    }
  }

  /**
   * Resize wiersza - zmienia wysokość
   */
  private resizeTableRow(
    st: NonNullable<typeof this.tableResizeState>,
    moveEvent: MouseEvent
  ): void {
    const deltaY = moveEvent.clientY - st.startY;
    const newHeight = Math.max(20, Math.round(st.startHeight + deltaY));
    const row = st.table.rows[st.rowIndex];
    if (row) {
      row.style.height = `${newHeight}px`;
    }
  }

  /**
   * Resize całej tabeli - skaluje proporcjonalnie
   */
  private resizeWholeTable(
    st: NonNullable<typeof this.tableResizeState>,
    moveEvent: MouseEvent
  ): void {
    const deltaX = moveEvent.clientX - st.startX;
    const editor = this.editorContent?.nativeElement;
    const maxW = (editor?.clientWidth || 900) - 20;
    const newTableWidth = Math.max(200, Math.min(maxW, st.startTableWidth + deltaX));
    const ratio = newTableWidth / st.startTableWidth;

    st.table.style.width = `${newTableWidth}px`;

    const firstRow = st.table.rows[0];
    if (firstRow) {
      const allRows = st.table.rows;
      for (let r = 0; r < allRows.length; r++) {
        const cells = allRows[r].cells;
        for (let c = 0; c < cells.length && c < st.startWidths.length; c++) {
          cells[c].style.width = `${st.startWidths[c] * ratio}px`;
        }
      }
    }
  }

  // ======= KONIEC RESIZE TABEL =======

  private handleEditorDragStart(event: DragEvent): void {
    const target = event.target as HTMLElement;
    const wrapper = target.closest('.editor-image-wrapper') as HTMLElement | null;
    if (!wrapper) {
      return;
    }

    this.draggedImageWrapper = wrapper;
    this.selectImageWrapper(wrapper);

    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/editor-image', '1');
    }
  }

  private handleEditorDragOver(event: DragEvent): void {
    if (!this.draggedImageWrapper) {
      return;
    }

    event.preventDefault();
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = 'move';
    }
  }

  private selectImageWrapper(wrapper: HTMLElement): void {
    if (this.selectedImageWrapper && this.selectedImageWrapper !== wrapper) {
      this.selectedImageWrapper.classList.remove('selected');
    }

    this.selectedImageWrapper = wrapper;
    this.selectedImageWrapper.classList.add('selected');
  }

  private clearSelectedImage(): void {
    if (this.selectedImageWrapper) {
      this.selectedImageWrapper.classList.remove('selected');
      this.selectedImageWrapper = null;
    }
  }

  private getRangeFromPoint(x: number, y: number): Range | null {
    const docWithCaret = document as Document & {
      caretRangeFromPoint?: (x: number, y: number) => Range | null;
      caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
    };

    if (docWithCaret.caretRangeFromPoint) {
      return docWithCaret.caretRangeFromPoint(x, y);
    }

    if (docWithCaret.caretPositionFromPoint) {
      const pos = docWithCaret.caretPositionFromPoint(x, y);
      if (pos) {
        const range = document.createRange();
        range.setStart(pos.offsetNode, pos.offset);
        range.collapse(true);
        return range;
      }
    }

    return null;
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
    if ((e.key === 'Delete' || e.key === 'Backspace') && this.selectedImageWrapper) {
      e.preventDefault();
      this.selectedImageWrapper.remove();
      this.selectedImageWrapper = null;
      this.onContentChange();
      return;
    }

    if (e.key === 'Escape' && this.selectedImageWrapper) {
      e.preventDefault();
      this.clearSelectedImage();
      return;
    }

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

    // Przenoszenie obrazu wewnątrz dokumentu
    if (this.draggedImageWrapper) {
      const editor = this.editorContent?.nativeElement;
      if (!editor) return;

      const dropRange = this.getRangeFromPoint(e.clientX, e.clientY);
      if (dropRange && editor.contains(dropRange.startContainer) && !this.draggedImageWrapper.contains(dropRange.startContainer)) {
        this.draggedImageWrapper.remove();
        dropRange.insertNode(this.draggedImageWrapper);
        this.selectImageWrapper(this.draggedImageWrapper);
        this.onContentChange();
      }

      this.draggedImageWrapper = null;
      return;
    }

    const files = e.dataTransfer?.files;
    if (!files || files.length === 0) return;

    // Obsłuż obrazy z systemu
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
   * Ustawia fokus na edytorze
   */
  focus(): void {
    this.editorContent?.nativeElement?.focus();
  }

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
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    const imageId = `img-${Date.now()}-${Math.floor(Math.random() * 10000)}`;

    // Buduj DOM bezpośrednio zamiast insertHTML (które nie działa po utracie focusu)
    const wrapper = document.createElement('span');
    wrapper.className = 'editor-image-wrapper';
    wrapper.setAttribute('data-image-id', imageId);
    wrapper.setAttribute('contenteditable', 'false');
    wrapper.setAttribute('draggable', 'true');
    wrapper.style.maxWidth = '100%';

    const img = document.createElement('img');
    img.src = src;
    img.alt = alt;
    img.style.maxWidth = '100%';
    img.style.height = 'auto';
    img.setAttribute('draggable', 'false');
    wrapper.appendChild(img);

    // Uchwyty resize: prawy (szerokość), dolny (wysokość), narożnik (proporcjonalnie)
    ['right', 'bottom', 'corner'].forEach(type => {
      const h = document.createElement('span');
      h.className = `image-resize-handle resize-handle-${type}`;
      h.title = type === 'right' ? 'Zmień szerokość' : type === 'bottom' ? 'Zmień wysokość' : 'Zmień rozmiar';
      wrapper.appendChild(h);
    });

    // Spróbuj wstawić w miejsce kursora / zapisanej selekcji
    let inserted = false;

    // Najpierw próbuj przywrócić selekcję
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      if (editor.contains(range.commonAncestorContainer)) {
        range.deleteContents();
        range.insertNode(wrapper);
        // Ustaw kursor za wstawionym obrazem
        range.setStartAfter(wrapper);
        range.collapse(true);
        selection.removeAllRanges();
        selection.addRange(range);
        inserted = true;
      }
    }

    // Jeśli nie udało się wstawić w selekcji, próbuj savedSelection
    if (!inserted && this.savedSelection) {
      editor.focus();
      const sel = window.getSelection();
      if (sel) {
        sel.removeAllRanges();
        sel.addRange(this.savedSelection);
        const range = sel.getRangeAt(0);
        if (editor.contains(range.commonAncestorContainer)) {
          range.deleteContents();
          range.insertNode(wrapper);
          range.setStartAfter(wrapper);
          range.collapse(true);
          sel.removeAllRanges();
          sel.addRange(range);
          inserted = true;
        }
      }
    }

    // Fallback: dołącz na koniec edytora
    if (!inserted) {
      editor.focus();
      const p = document.createElement('p');
      p.appendChild(wrapper);
      editor.appendChild(p);
    }

    this.selectImageWrapper(wrapper);
    this.onContentChange();
  }

  /**
   * Wstawia kod kreskowy z wartością tekstową pod spodem
   */
  insertBarcodeWithValue(src: string, valueText: string): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    const imageId = `img-${Date.now()}-${Math.floor(Math.random() * 10000)}`;

    // Kontener na kod kreskowy z wartością
    const container = document.createElement('div');
    container.className = 'barcode-container';
    container.setAttribute('contenteditable', 'false');
    container.style.display = 'inline-block';
    container.style.textAlign = 'center';

    const wrapper = document.createElement('span');
    wrapper.className = 'editor-image-wrapper';
    wrapper.setAttribute('data-image-id', imageId);
    wrapper.setAttribute('contenteditable', 'false');
    wrapper.setAttribute('draggable', 'true');
    wrapper.style.maxWidth = '100%';
    wrapper.style.display = 'block';

    const img = document.createElement('img');
    img.src = src;
    img.alt = 'barcode';
    img.style.maxWidth = '100%';
    img.style.height = 'auto';
    img.setAttribute('draggable', 'false');
    wrapper.appendChild(img);

    // Uchwyty resize
    ['right', 'bottom', 'corner'].forEach(type => {
      const h = document.createElement('span');
      h.className = `image-resize-handle resize-handle-${type}`;
      h.title = type === 'right' ? 'Zmień szerokość' : type === 'bottom' ? 'Zmień wysokość' : 'Zmień rozmiar';
      wrapper.appendChild(h);
    });

    container.appendChild(wrapper);

    // Tekst wartości pod kodem
    const valueDiv = document.createElement('div');
    valueDiv.className = 'barcode-value-text';
    valueDiv.style.cssText = 'font-size: 12px; font-family: monospace; color: #333; margin-top: 4px; text-align: center; word-break: break-all;';
    valueDiv.textContent = valueText;
    container.appendChild(valueDiv);

    // Wstaw w miejsce kursora
    let inserted = false;
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      if (editor.contains(range.commonAncestorContainer)) {
        range.deleteContents();
        range.insertNode(container);
        range.setStartAfter(container);
        range.collapse(true);
        selection.removeAllRanges();
        selection.addRange(range);
        inserted = true;
      }
    }

    if (!inserted && this.savedSelection) {
      editor.focus();
      const sel = window.getSelection();
      if (sel) {
        sel.removeAllRanges();
        sel.addRange(this.savedSelection);
        const range = sel.getRangeAt(0);
        if (editor.contains(range.commonAncestorContainer)) {
          range.deleteContents();
          range.insertNode(container);
          range.setStartAfter(container);
          range.collapse(true);
          sel.removeAllRanges();
          sel.addRange(range);
          inserted = true;
        }
      }
    }

    if (!inserted) {
      editor.focus();
      const p = document.createElement('p');
      p.appendChild(container);
      editor.appendChild(p);
    }

    this.selectImageWrapper(wrapper);
    this.onContentChange();
  }

  /**
   * Wstawia tabelę
   */
  insertTable(config: string): void {
    const [rows, cols] = config.split('x').map(Number);
    const editor = this.editorContent?.nativeElement;
    if (!editor || rows <= 0 || cols <= 0) return;

    const colWidth = Math.floor(100 / cols);

    // Buduj tabelę jako element DOM (zamiast execCommand insertHTML, który nie działa bez focusu)
    const table = document.createElement('table');
    table.style.cssText = 'border-collapse:collapse;width:100%;margin:10px 0;table-layout:fixed;position:relative;';

    for (let i = 0; i < rows; i++) {
      const tr = document.createElement('tr');
      for (let j = 0; j < cols; j++) {
        const td = document.createElement('td');
        td.style.cssText = `border:1px solid #ccc;padding:8px;min-width:30px;width:${colWidth}%;`;
        td.innerHTML = '&nbsp;';
        tr.appendChild(td);
      }
      table.appendChild(tr);
    }

    const afterParagraph = document.createElement('p');
    afterParagraph.innerHTML = '&nbsp;';

    const fragment = document.createDocumentFragment();
    fragment.appendChild(table);
    fragment.appendChild(afterParagraph);

    // Próbuj wstawić w miejsce kursora / selekcji
    let inserted = false;

    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      if (editor.contains(range.commonAncestorContainer)) {
        range.deleteContents();
        range.insertNode(fragment);
        const newRange = document.createRange();
        newRange.setStartAfter(afterParagraph);
        newRange.collapse(true);
        selection.removeAllRanges();
        selection.addRange(newRange);
        inserted = true;
      }
    }

    // Fallback: użyj zapisanej selekcji (utraconej przez kliknięcie dialogu)
    if (!inserted && this.savedSelection) {
      editor.focus();
      const sel = window.getSelection();
      if (sel) {
        sel.removeAllRanges();
        sel.addRange(this.savedSelection);
        const range = sel.getRangeAt(0);
        if (editor.contains(range.commonAncestorContainer)) {
          range.deleteContents();
          range.insertNode(fragment);
          const newRange = document.createRange();
          newRange.setStartAfter(afterParagraph);
          newRange.collapse(true);
          sel.removeAllRanges();
          sel.addRange(newRange);
          inserted = true;
        }
      }
    }

    // Fallback: dołącz na koniec edytora
    if (!inserted) {
      editor.focus();
      editor.appendChild(table);
      editor.appendChild(afterParagraph);
    }

    this.savedSelection = null;
    this.onContentChange();
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
    const editor = this.editorContent?.nativeElement;
    if (!editor) return '';

    // Przechwytuj rzeczywiste wysokości wierszy tabel z DOM (przed klonowaniem)
    // Dzięki temu wysokości są ZAWSZE w HTML - niezależnie czy pochodzą z drag-resize czy z treści
    const tableRowHeights: Map<number, { heights: number[] }>  = new Map();
    const liveTables = editor.querySelectorAll('table');
    liveTables.forEach((table, tableIdx) => {
      const rows = table.querySelectorAll('tr');
      const heights: number[] = [];
      rows.forEach(tr => {
        heights.push(Math.round((tr as HTMLElement).getBoundingClientRect().height));
      });
      tableRowHeights.set(tableIdx, { heights });
    });

    const clone = editor.cloneNode(true) as HTMLDivElement;

    // Zastosuj przechwycone wysokości do sklonowanych wierszy
    const cloneTables = clone.querySelectorAll('table');
    cloneTables.forEach((table, tableIdx) => {
      const data = tableRowHeights.get(tableIdx);
      if (!data) return;
      const rows = table.querySelectorAll('tr');
      rows.forEach((tr, rowIdx) => {
        const h = data.heights[rowIdx];
        if (h && h > 0) {
          (tr as HTMLElement).style.height = `${h}px`;
        }
      });
    });

    clone.querySelectorAll('.editor-image-wrapper').forEach(wrapperEl => {
      const wrapper = wrapperEl as HTMLElement;
      wrapper.classList.remove('selected');
      wrapper.removeAttribute('contenteditable');
      wrapper.removeAttribute('draggable');

      wrapper.querySelectorAll('.image-resize-handle').forEach(h => h.remove());

      const img = wrapper.querySelector('img');
      if (img) {
        const imgEl = img as HTMLImageElement;
        const wrapperWidth = wrapper.style.width;
        if (wrapperWidth && wrapperWidth !== 'auto') {
          imgEl.style.width = wrapperWidth;
        }
        imgEl.style.maxWidth = '100%';
        // Zachowaj height jeśli był ustawiony (nie 'auto')
        if (!imgEl.style.height || imgEl.style.height === 'auto') {
          imgEl.style.height = 'auto';
        }
        imgEl.removeAttribute('draggable');
        wrapper.replaceWith(imgEl);
      }
    });

    return clone.innerHTML;
  }

  /**
   * Ustawia zawartość HTML
   */
  setContent(html: string): void {
    if (this.editorContent?.nativeElement) {
      this.editorContent.nativeElement.innerHTML = html;
      this.wrapExistingImages();
      this._content.set(this.editorContent.nativeElement.innerHTML);
      this.lastSavedContent = this.editorContent.nativeElement.innerHTML;
      this.undoStack = [this.editorContent.nativeElement.innerHTML];
      this.redoStack = [];
      this.updateState();
    }
  }

  /**
   * Opakowuje istniejące elementy <img> (bez wrappera) w editor-image-wrapper
   */
  private wrapExistingImages(): void {
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    const images = editor.querySelectorAll('img');
    images.forEach((img: HTMLImageElement) => {
      // Pomiń obrazy które już mają wrapper
      if (img.parentElement?.classList.contains('editor-image-wrapper')) {
        return;
      }

      const imageId = `img-${Date.now()}-${Math.floor(Math.random() * 10000)}`;

      const wrapper = document.createElement('span');
      wrapper.className = 'editor-image-wrapper';
      wrapper.setAttribute('data-image-id', imageId);
      wrapper.setAttribute('contenteditable', 'false');
      wrapper.setAttribute('draggable', 'true');

      // Przenieś width z img na wrapper
      if (img.style.width) {
        wrapper.style.width = img.style.width;
      }
      wrapper.style.maxWidth = '100%';

      img.style.maxWidth = '100%';
      img.style.height = 'auto';
      img.setAttribute('draggable', 'false');

      img.parentNode?.insertBefore(wrapper, img);
      wrapper.appendChild(img);

      ['right', 'bottom', 'corner'].forEach(type => {
        const h = document.createElement('span');
        h.className = `image-resize-handle resize-handle-${type}`;
        h.title = type === 'right' ? 'Zmień szerokość' : type === 'bottom' ? 'Zmień wysokość' : 'Zmień rozmiar';
        wrapper.appendChild(h);
      });
    });
  }

  /**
   * Oznacza dokument jako zapisany
   */
  markAsSaved(): void {
    this.lastSavedContent = this.getContent();
    this.updateState();
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
   * Wylicza zawartość nagłówka dla danej strony (używane wewnętrznie przez computed)
   */
  private _computeHeaderContent(pageIndex: number): string {
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
   * Pobiera zawartość nagłówka dla danej strony (używane w szablonie)
   */
  getHeaderContent(pageIndex: number): string {
    const contents = this.headerContents();
    return contents[pageIndex] ?? this._headerHtml();
  }

  /**
   * Wylicza zawartość stopki dla danej strony (używane wewnętrznie przez computed)
   */
  private _computeFooterContent(pageIndex: number): string {
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
   * Pobiera zawartość stopki dla danej strony (używane w szablonie)
   */
  getFooterContent(pageIndex: number): string {
    const contents = this.footerContents();
    return contents[pageIndex] ?? this._footerHtml();
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

  // ========== Wyszukiwanie i zamiana ==========

  private searchHighlights: HTMLElement[] = [];
  private currentHighlightIndex = -1;

  /**
   * Wyszukuje tekst w dokumencie i podświetla wyniki
   */
  searchText(text: string, direction: 'next' | 'previous'): { count: number; currentIndex: number } {
    this.clearSearchHighlights();
    
    const editor = this.editorContent?.nativeElement;
    if (!editor || !text) return { count: 0, currentIndex: -1 };

    const searchLower = text.toLowerCase();
    const treeWalker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT, null);
    const matches: { node: Text; index: number }[] = [];

    while (treeWalker.nextNode()) {
      const node = treeWalker.currentNode as Text;
      const content = node.textContent || '';
      let idx = content.toLowerCase().indexOf(searchLower);
      while (idx !== -1) {
        matches.push({ node, index: idx });
        idx = content.toLowerCase().indexOf(searchLower, idx + 1);
      }
    }

    if (matches.length === 0) return { count: 0, currentIndex: -1 };

    // Podświetl wszystkie wyniki (od końca żeby nie psuć indeksów)
    for (let i = matches.length - 1; i >= 0; i--) {
      const { node, index } = matches[i];
      const range = document.createRange();
      range.setStart(node, index);
      range.setEnd(node, index + text.length);

      const highlight = document.createElement('mark');
      highlight.className = 'search-highlight';
      highlight.style.backgroundColor = '#fff3a8';
      highlight.style.color = 'inherit';
      highlight.style.padding = '0';
      highlight.dataset['searchHighlight'] = 'true';

      try {
        range.surroundContents(highlight);
        this.searchHighlights.unshift(highlight);
      } catch {
        // Jeśli range obejmuje wiele elementów, pomiń
      }
    }

    // Ustaw aktualny indeks
    if (this.searchHighlights.length > 0) {
      this.currentHighlightIndex = 0;
      if (direction === 'previous') {
        this.currentHighlightIndex = this.searchHighlights.length - 1;
      }
      this.highlightCurrent();
    }

    return { count: this.searchHighlights.length, currentIndex: this.currentHighlightIndex };
  }

  /**
   * Przechodzi do następnego wyniku wyszukiwania
   */
  findNext(): { count: number; currentIndex: number } {
    if (this.searchHighlights.length === 0) return { count: 0, currentIndex: -1 };
    this.currentHighlightIndex = (this.currentHighlightIndex + 1) % this.searchHighlights.length;
    this.highlightCurrent();
    return { count: this.searchHighlights.length, currentIndex: this.currentHighlightIndex };
  }

  /**
   * Przechodzi do poprzedniego wyniku wyszukiwania
   */
  findPrevious(): { count: number; currentIndex: number } {
    if (this.searchHighlights.length === 0) return { count: 0, currentIndex: -1 };
    this.currentHighlightIndex = (this.currentHighlightIndex - 1 + this.searchHighlights.length) % this.searchHighlights.length;
    this.highlightCurrent();
    return { count: this.searchHighlights.length, currentIndex: this.currentHighlightIndex };
  }

  /**
   * Podświetla aktualny wynik wyszukiwania
   */
  private highlightCurrent(): void {
    this.searchHighlights.forEach((el, i) => {
      if (i === this.currentHighlightIndex) {
        el.style.backgroundColor = '#ff9632';
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      } else {
        el.style.backgroundColor = '#fff3a8';
      }
    });
  }

  /**
   * Zamienia bieżący wynik wyszukiwania
   */
  replaceCurrentMatch(replaceText: string): { count: number; currentIndex: number } {
    if (this.searchHighlights.length === 0 || this.currentHighlightIndex < 0) {
      return { count: 0, currentIndex: -1 };
    }

    const highlight = this.searchHighlights[this.currentHighlightIndex];
    const textNode = document.createTextNode(replaceText);
    highlight.parentNode?.replaceChild(textNode, highlight);
    this.searchHighlights.splice(this.currentHighlightIndex, 1);

    if (this.currentHighlightIndex >= this.searchHighlights.length) {
      this.currentHighlightIndex = 0;
    }
    if (this.searchHighlights.length > 0) {
      this.highlightCurrent();
    }

    this.emitContentChange();
    return { count: this.searchHighlights.length, currentIndex: this.currentHighlightIndex };
  }

  /**
   * Zamienia wszystkie wyniki wyszukiwania
   */
  replaceAllMatches(replaceText: string): { count: number; currentIndex: number } {
    for (const highlight of this.searchHighlights) {
      const textNode = document.createTextNode(replaceText);
      highlight.parentNode?.replaceChild(textNode, highlight);
    }
    this.searchHighlights = [];
    this.currentHighlightIndex = -1;
    this.emitContentChange();
    return { count: 0, currentIndex: -1 };
  }

  /**
   * Czyści podświetlenia wyszukiwania
   */
  clearSearchHighlights(): void {
    for (const highlight of this.searchHighlights) {
      const parent = highlight.parentNode;
      if (parent) {
        while (highlight.firstChild) {
          parent.insertBefore(highlight.firstChild, highlight);
        }
        parent.removeChild(highlight);
        parent.normalize();
      }
    }
    this.searchHighlights = [];
    this.currentHighlightIndex = -1;
  }

  private emitContentChange(): void {
    const editor = this.editorContent?.nativeElement;
    if (editor) {
      this._content.set(editor.innerHTML);
      this.contentChange.emit(editor.innerHTML);
    }
  }
}
