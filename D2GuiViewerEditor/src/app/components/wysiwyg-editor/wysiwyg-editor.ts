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
  ViewEncapsulation,
  input
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DomSanitizer, SafeHtml } from '@angular/platform-browser';
import { 
  EditorCommand, 
  EditorState, 
  HeadingLevel, 
  TextFormatting,
  ParagraphStyle,
  PageMargins,
  HeaderFooterContent
} from '../../models/document.model';
import { normalizeWhitespace, resolvePlainText } from '../../core/utils/paste-text.util';

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
  /**
   * Aktywny edytor strony — ustawiany ręcznie przy `focusin` na konkretnej stronie.
   * Cała istniejąca logika (toolbar, paste, undo, table edit, image drag, search)
   * pracuje na tym refie. Dzięki temu MVP multi-page nie wymaga zmian
   * w setkach miejsc kodu.
   */
  editorContent!: ElementRef<HTMLDivElement>;

  @ViewChildren('pageEditor') pageEditorRefs!: QueryList<ElementRef<HTMLDivElement>>;
  @ViewChild('headerContent') headerContentEl?: ElementRef<HTMLDivElement>;
  @ViewChild('footerContent') footerContentEl?: ElementRef<HTMLDivElement>;

  /**
   * Zwraca AKTUALNIE edytowany contenteditable: header / footer / aktywna strona body.
   * Toolbar musi kierować komendy (B/I/U/align/color/font/size/insertImage) tu,
   * a nie zawsze na body — inaczej formatowanie w header/footer nie zadziała.
   */
  private getActiveEditor(): HTMLDivElement | null {
    const section = this.editingSection();
    if (section === 'header' && this.headerContentEl?.nativeElement) {
      return this.headerContentEl.nativeElement;
    }
    if (section === 'footer' && this.footerContentEl?.nativeElement) {
      return this.footerContentEl.nativeElement;
    }
    return this.editorContent?.nativeElement ?? null;
  }
  
  @Input() set content(value: string) {
    // Nie aktualizuj innerHTML jeśli wartość pochodzi z tego samego edytora
    // (zapobiega resetowaniu kursora podczas pisania)
    if (this._content() === value) {
      return;
    }
    
    this._content.set(value);
    if (!this._isInternalUpdate) {
      // Rozbij na strony po znacznikach <div class="page-break">
      const splitPages = this._splitHtmlIntoPages(value || '<p></p>');
      this.pageContents.set(splitPages.length ? splitPages : ['<p></p>']);
      // Po Angular re-render zaktualizuj aktywny edytor i zrepaginuj
      this._schedulePaginate('content-input');
    }
  }
  
  pageMargins = input<PageMargins>({ top: 2.5, bottom: 2.5, left: 2.5, right: 2.5 });
  @Input() pageOrientation: 'portrait' | 'landscape' = 'portrait';
  /** Tryb tylko-do-odczytu (Krok 2) — blokuje edycję contenteditable. */
  @Input() readOnly = false;
  @Input() showMarginGuides = false;
  
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
  /** Emituje aktualnie edytowaną sekcję (treść / nagłówek / stopka) — używane przez pionową linijkę. */
  @Output() editingSectionChange = new EventEmitter<'header' | 'footer' | 'body'>();
  /**
   * Emits a snapshot of the currently selected image (or null when nothing is selected).
   * Drives d2-image-properties-panel — the parent owns the signal and the visibility logic.
   */
  @Output() imageSelectionChange = new EventEmitter<{
    widthPx: number;
    heightPx: number;
    aspectRatio: number;
    alignment: 'left' | 'center' | 'right' | null;
    positionMode: 'inline' | 'square' | 'topBottom' | 'front' | 'behind';
    border: { enabled: boolean; color: string; widthPx: number; style: 'solid' | 'dashed' | 'dotted' };
    crop: { left: number; right: number; top: number; bottom: number };
  } | null>();
  /**
   * Emituje ZMIERZONĄ geometrię edytowanego pasma nagłówka/stopki (cm od górnej krawędzi
   * strony 1). Pasmo ma `min-height` i rośnie z treścią (np. obraz), więc pionowa linijka
   * musi odzwierciedlać faktyczne położenie, a nie wyliczone z marginesów cm.
   */
  @Output() sectionGeometryChange = new EventEmitter<{ section: 'header' | 'footer'; topCm: number; bottomCm: number }>();

  /** Obserwator rozmiaru aktywnego pasma nagłówka/stopki (re-emisja geometrii przy zmianie wysokości). */
  private _sectionResizeObserver?: ResizeObserver;
  @Output() openHeaderFooterSettings = new EventEmitter<{
    headerMargin: number;
    footerMargin: number;
    differentFirstPage: boolean;
    differentOddEven: boolean;
  }>();

  private _content = signal<string>('');
  private _isInternalUpdate = false;
  /** Flaga ustawiana na input, czyszczona przy save — pozwala uniknąć ciężkiego getContent() w updateState. */
  private _isDirty = false;
  /** Debouncer dla saveToUndoStack + emitContent — nie wykonujemy ich na każde naciśnięcie klawisza. */
  private _persistTimer: ReturnType<typeof setTimeout> | null = null;
  private undoStack: string[] = [];
  private redoStack: string[] = [];
  private lastSavedContent = '';
  private pageCheckInterval?: ReturnType<typeof setInterval>;
  private selectedImageWrapper: HTMLElement | null = null;
  private draggedImageWrapper: HTMLElement | null = null;
  private imageDragCaret: HTMLElement | null = null;
  private imageMoveState: {
    wrapper: HTMLElement;
    startX: number;
    startY: number;
    isDragging: boolean;
  } | null = null;
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

  // Multi-page MVP (Wariant A):
  // pełna treść HTML per strona; każda strona renderuje własny contenteditable
  pageContents = signal<string[]>(['<p></p>']);
  // która strona ma aktualnie focus (do operacji toolbar/undo/paste)
  activePageIndex = signal<number>(0);

  private _sanitizer = inject(DomSanitizer);
  /** Cache trusted-HTML per strona — KLUCZOWE dla wydajności i contenteditable.
   *  Bez tego każde change detection tworzy nowy obiekt SafeHtml, Angular widzi
   *  „zmianę" i rebinduje innerHTML co kasuje kursor + uniemożliwia pisanie. */
  private _safeHtmlCache: Array<{ html: string; safe: SafeHtml }> = [];
  getPageContentSafe(index: number): SafeHtml {
    const html = this.pageContents()[index] ?? '';
    const cached = this._safeHtmlCache[index];
    if (cached && cached.html === html) {
      return cached.safe;
    }
    const safe = this._sanitizer.bypassSecurityTrustHtml(html);
    this._safeHtmlCache[index] = { html, safe };
    return safe;
  }

  /** Cache SafeHtml dla nagłówków/stopek — żeby preview zachował formatowanie
   *  inline (color/font-size/text-align/img style="width:..."). Bez tego Angular
   *  strippuje atrybuty `style` i obraz/rozmiar/kolor "ginie" w trybie podglądu. */
  private _safeHeaderCache = new Map<number, { html: string; safe: SafeHtml }>();
  private _safeFooterCache = new Map<number, { html: string; safe: SafeHtml }>();

  getHeaderContentSafe(pageIndex: number): SafeHtml {
    const html = this.getHeaderContent(pageIndex) ?? '';
    const cached = this._safeHeaderCache.get(pageIndex);
    if (cached && cached.html === html) return cached.safe;
    const safe = this._sanitizer.bypassSecurityTrustHtml(html);
    this._safeHeaderCache.set(pageIndex, { html, safe });
    return safe;
  }

  getFooterContentSafe(pageIndex: number): SafeHtml {
    const html = this.getFooterContent(pageIndex) ?? '';
    const cached = this._safeFooterCache.get(pageIndex);
    if (cached && cached.html === html) return cached.safe;
    const safe = this._sanitizer.bypassSecurityTrustHtml(html);
    this._safeFooterCache.set(pageIndex, { html, safe });
    return safe;
  }

  /** Inwalidacja cache nagłówka/stopki — wołać po każdej edycji */
  private invalidateHeaderFooterCache(): void {
    this._safeHeaderCache.clear();
    this._safeFooterCache.clear();
  }

  // Wysokość strony A4 w pikselach (bez marginesów)
  private readonly PAGE_HEIGHT_PX = 1122; // ~29.7cm at 96 DPI

  // Paginator: debounce + safety flag
  private _paginateTimer: ReturnType<typeof setTimeout> | null = null;
  private _isRepaginating = false;

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

  // Computed: efektywne paddingi treści (margines minus wysokość nagłówka/stopki)
  // W MS Word, nagłówek/stopka zajmują CZĘŚĆ marginesu, nie dodają się do niego
  contentPaddingTop = computed(() => {
    const topMargin = this.pageMargins().top; // cm
    const headerH = this._headerHeight(); // cm
    const effectivePadding = Math.max(0, topMargin - headerH);
    return effectivePadding * 37.8; // px
  });
  
  contentPaddingBottom = computed(() => {
    const bottomMargin = this.pageMargins().bottom; // cm
    const footerH = this._footerHeight(); // cm
    const effectivePadding = Math.max(0, bottomMargin - footerH);
    return effectivePadding * 37.8; // px
  });

  // Stan edytora
  editorState = signal<EditorState>({
    isModified: false,
    canUndo: false,
    canRedo: false,
    wordCount: 0,
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

  // Ostatnia policzona liczba stron (do emisji pagesChange bez rerenderu DOM)
  private lastEmittedPageCount = 1;

  ngAfterViewInit(): void {
    // Ustaw editorContent na pierwszej (aktywnej) stronie i obserwuj zmiany
    // (np. po repaginacji liczba stron się zmienia).
    const syncActiveEditor = () => {
      const refs = this.pageEditorRefs?.toArray() ?? [];
      const idx = Math.min(this.activePageIndex(), Math.max(0, refs.length - 1));
      if (refs[idx]) {
        this.editorContent = refs[idx];
      }
    };
    syncActiveEditor();
    this.pageEditorRefs?.changes.subscribe(() => {
      syncActiveEditor();
      // Po repaginacji ponownie podpinamy listenery do nowych edytorów
      this.setupEventListeners();
    });

    this.initializeEditor();
    this.setupEventListeners();
    
    // Oblicz strony przy starcie
    setTimeout(() => {
      this.calculatePages();
      this._schedulePaginate('init');
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
    this._sectionResizeObserver?.disconnect();
  }

  /**
   * Zaczyna obserwować pasmo aktywnej sekcji (header/footer) i emituje jego zmierzoną
   * geometrię w cm od górnej krawędzi strony 1. Re-emituje przy zmianie wysokości pasma
   * (np. po załadowaniu obrazu w nagłówku).
   */
  private observeActiveSectionGeometry(): void {
    this._sectionResizeObserver?.disconnect();
    const section = this.editingSection();
    if (section !== 'header' && section !== 'footer') return;

    const inner = section === 'header'
      ? this.headerContentEl?.nativeElement
      : this.footerContentEl?.nativeElement;
    const band = inner?.closest(section === 'header' ? '.page-header' : '.page-footer') as HTMLElement | null;
    if (!band) return;

    const emit = () => this.emitSectionGeometry(section, band);
    emit();
    this._sectionResizeObserver = new ResizeObserver(() => emit());
    this._sectionResizeObserver.observe(band);
  }

  /** Mierzy pasmo względem strony (uwzględnia skalę zoomu) i emituje cm od góry strony. */
  private emitSectionGeometry(section: 'header' | 'footer', band: HTMLElement): void {
    const page = band.closest('.page') as HTMLElement | null;
    if (!page) return;
    const pr = page.getBoundingClientRect();
    const br = band.getBoundingClientRect();
    // Skala niezależna od wzrostu pasma: z szerokości strony (stała: A4 21cm / landscape 29.7cm).
    const expectedWidthPx = (this.pageOrientation === 'portrait' ? 21 : 29.7) * 37.8;
    const scale = pr.width > 0 ? pr.width / expectedWidthPx : 1;
    const topCm = ((br.top - pr.top) / scale) / 37.8;
    const bottomCm = ((br.bottom - pr.top) / scale) / 37.8;
    this.sectionGeometryChange.emit({ section, topCm, bottomCm });
  }

  private stopObservingSectionGeometry(): void {
    this._sectionResizeObserver?.disconnect();
    this._sectionResizeObserver = undefined;
  }

  /**
   * Emituje aktualną liczbę stron (= długość `pageContents`, czyli zgodną
   * z wizualnym podziałem po repaginacji Wariantu A).
   */
  private calculatePages(): void {
    const pageCount = Math.max(1, this.pageContents().length);
    if (pageCount !== this.lastEmittedPageCount) {
      this.lastEmittedPageCount = pageCount;
      this.pagesChange.emit(pageCount);
    }
  }

  /** Zwraca innerHTML edytora (legacy helper – wcześniej usuwał wstrzykiwane separatory stron). */
  private _getCleanEditorHtml(editor: HTMLElement): string {
    return editor.innerHTML;
  }

  /**
   * Inicjalizuje edytor
   */
  private initializeEditor(): void {
    // pageContents jest już zainicjalizowany (np. przez setter content/setContent).
    // Po renderze Angular nakłada [innerHTML] na każdą stronę.
    // Tu opakowujemy obrazki i robimy snapshot dla isModified.
    const editor = this.editorContent?.nativeElement;
    if (!editor) return;

    this.wrapExistingImages();
    this.lastSavedContent = this._getCleanEditorHtml(editor);
    this.saveToUndoStack();
    this.updateState();
  }

  /**
   * Konfiguruje nasłuchiwanie zdarzeń
   */
  private setupEventListeners(): void {
    const refs = this.pageEditorRefs?.toArray() ?? [];
    // Globalne nasłuchiwanie selekcji (idempotentne — flag na document)
    if (!(document as any).__wysiwygSelListener) {
      document.addEventListener('selectionchange', () => {
        this.onSelectionChange();
      });
      (document as any).__wysiwygSelListener = true;
    }

    for (const ref of refs) {
      this.attachEditorListeners(ref.nativeElement);
    }
  }

  /**
   * Rejestruje pełen zestaw listenerów (input/paste/keydown/click/mousedown/drag/resize)
   * dla danego contenteditable. Używane dla każdej strony body oraz dla header/footer
   * (po pierwszym wejściu w edycję). Idempotentne — flag `__wysiwygBound`.
   */
  private attachEditorListeners(editorEl: HTMLDivElement | null | undefined): void {
    const editor = editorEl as (HTMLDivElement & { __wysiwygBound?: boolean }) | null | undefined;
    if (!editor || editor.__wysiwygBound) return;
    editor.__wysiwygBound = true;

    editor.addEventListener('input', () => {
      this.onContentChange();
    });
    editor.addEventListener('paste', (e) => {
      this.handlePaste(e);
    });
    editor.addEventListener('keydown', (e) => {
      this.handleKeyboard(e);
    });
    editor.addEventListener('blur', () => {
      this.saveSelection();
    });
    editor.addEventListener('drop', (e) => {
      this.handleDrop(e);
    });
    editor.addEventListener('click', (e) => {
      this.handleEditorClick(e);
    });
    editor.addEventListener('mousedown', (e) => {
      this.handleEditorMouseDown(e);
    });
    editor.addEventListener('dragstart', (e) => {
      this.handleEditorDragStart(e);
    });
    editor.addEventListener('dragover', (e) => {
      this.handleEditorDragOver(e);
    });
    editor.addEventListener('dragend', () => {
      this.draggedImageWrapper = null;
    });
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

        // KONTENER = aktywny edytor (body / header / footer) — wrapper.closest
        const editor = wrapper.closest('.editor-content, .header-editor-content, .footer-editor-content') as HTMLElement | null;
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

          // Po zakończeniu skalowania zapisujemy realny rozmiar na <img>
          // i aktualizujemy `data-width-emu` / `data-height-emu`, których
          // używa eksporter DOCX. Inaczej eksport używa ORYGINALNYCH wymiarów
          // EMU (z importu), ignorując zmianę w edytorze — i obraz w pliku
          // .docx jest dużo większy niż widać w edytorze.
          const finalWrapper = this.imageResizeState.wrapper;
          const finalImg = finalWrapper.querySelector('img') as HTMLImageElement | null;
          if (finalImg) {
            const rect = finalImg.getBoundingClientRect();
            const widthPx = Math.round(rect.width);
            const heightPx = Math.round(rect.height);
            if (widthPx > 0 && heightPx > 0) {
              finalImg.style.width = `${widthPx}px`;
              finalImg.style.height = `${heightPx}px`;
              // 1 px = 9525 EMU (przybliżenie używane też po stronie API)
              const EMU_PER_PX = 9525;
              finalImg.setAttribute('data-width-emu', String(widthPx * EMU_PER_PX));
              finalImg.setAttribute('data-height-emu', String(heightPx * EMU_PER_PX));
            }
          }
        }

        this.imageResizeState = null;
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
        this.onContentChange();
        // Snapshot the new dimensions so the side panel reflects the post-resize state.
        this.emitImageSelectionState();
      };

      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
      return;
    }

    // Kliknięcie / przeciąganie na wrapper obrazu (nie resize handle)
    event.preventDefault();
    this.selectImageWrapper(wrapper);

    const startX = event.clientX;
    const startY = event.clientY;
    const isFloating = wrapper.dataset['posMode'] === 'front' || wrapper.dataset['posMode'] === 'behind';
    this.imageMoveState = { wrapper, startX, startY, isDragging: false };

    // Floating drag — image stays absolutely positioned, drag updates left/top.
    if (isFloating) {
      const startLeft = parseInt(wrapper.style.left || '0', 10);
      const startTop = parseInt(wrapper.style.top || '0', 10);
      const onFloatingMove = (moveEvent: MouseEvent) => {
        const dx = moveEvent.clientX - startX;
        const dy = moveEvent.clientY - startY;
        if (!this.imageMoveState!.isDragging && Math.hypot(dx, dy) > 3) {
          this.imageMoveState!.isDragging = true;
          wrapper.classList.add('image-dragging');
        }
        if (this.imageMoveState!.isDragging) {
          const newLeft = Math.max(0, startLeft + dx);
          const newTop = Math.max(0, startTop + dy);
          wrapper.style.left = `${newLeft}px`;
          wrapper.style.top = `${newTop}px`;
        }
      };
      const onFloatingUp = () => {
        if (this.imageMoveState?.isDragging) {
          const xPx = parseInt(wrapper.style.left || '0', 10);
          const yPx = parseInt(wrapper.style.top || '0', 10);
          wrapper.dataset['xPx'] = String(xPx);
          wrapper.dataset['yPx'] = String(yPx);
          const img = wrapper.querySelector('img') as HTMLImageElement | null;
          if (img) {
            const EMU_PER_PX = 9525;
            img.setAttribute('data-x-emu', String(xPx * EMU_PER_PX));
            img.setAttribute('data-y-emu', String(yPx * EMU_PER_PX));
          }
          wrapper.classList.remove('image-dragging');
          this.onContentChange();
          this.emitImageSelectionState();
        }
        this.imageMoveState = null;
        document.removeEventListener('mousemove', onFloatingMove);
        document.removeEventListener('mouseup', onFloatingUp);
      };
      document.addEventListener('mousemove', onFloatingMove);
      document.addEventListener('mouseup', onFloatingUp);
      return;
    }

    const onImageMouseMove = (moveEvent: MouseEvent) => {
      if (!this.imageMoveState) return;
      const dx = moveEvent.clientX - this.imageMoveState.startX;
      const dy = moveEvent.clientY - this.imageMoveState.startY;
      if (!this.imageMoveState.isDragging && Math.hypot(dx, dy) > 5) {
        this.imageMoveState.isDragging = true;
        document.body.classList.add('image-moving');
        wrapper.classList.add('image-dragging');
        // Wyłącz pointer-events na wrapperze żeby getRangeFromPoint trafiał w tekst pod grafiką
        wrapper.style.pointerEvents = 'none';
        // Utwórz element wskazujący miejsce upuszczenia (kursor edytora)
        this.imageDragCaret = document.createElement('div');
        this.imageDragCaret.className = 'image-drop-caret';
        document.body.appendChild(this.imageDragCaret);
      }

      if (this.imageMoveState.isDragging && this.imageDragCaret) {
        const editor = wrapper.closest('.editor-content, .header-editor-content, .footer-editor-content') as HTMLElement | null;
        const range = editor ? this.getRangeFromPoint(moveEvent.clientX, moveEvent.clientY) : null;
        if (range && editor && editor.contains(range.startContainer) && !wrapper.contains(range.startContainer)) {
          const rect = range.getBoundingClientRect();
          if (rect.height > 0) {
            this.imageDragCaret.style.display = 'block';
            this.imageDragCaret.style.left = `${rect.left}px`;
            this.imageDragCaret.style.top = `${rect.top}px`;
            this.imageDragCaret.style.height = `${rect.height}px`;
          } else {
            this.imageDragCaret.style.display = 'none';
          }
        } else {
          this.imageDragCaret.style.display = 'none';
        }
      }
    };

    const onImageMouseUp = (upEvent: MouseEvent) => {
      if (this.imageMoveState?.isDragging) {
        // Przywróć pointer-events i usuń kursor upuszczenia
        wrapper.style.pointerEvents = '';
        this.imageDragCaret?.remove();
        this.imageDragCaret = null;

        const editor = wrapper.closest('.editor-content, .header-editor-content, .footer-editor-content') as HTMLElement | null;
        if (editor) {
          wrapper.classList.remove('image-dragging');
          const dropRange = this.getRangeFromPoint(upEvent.clientX, upEvent.clientY);
          if (dropRange && editor.contains(dropRange.startContainer) && !wrapper.contains(dropRange.startContainer)) {
            wrapper.remove();
            dropRange.insertNode(wrapper);
            this.selectImageWrapper(wrapper);
            this.onContentChange();
            // Snapshot in case alignment changed because the drop landed in a paragraph
            // with different text-align.
            this.emitImageSelectionState();
          }
        }
      }
      // Sprzątanie po każdym mouse-up (także gdy nie było faktycznego drag)
      document.body.classList.remove('image-moving');
      wrapper.classList.remove('image-dragging');
      wrapper.style.pointerEvents = '';
      this.imageDragCaret?.remove();
      this.imageDragCaret = null;
      this.imageMoveState = null;
      document.removeEventListener('mousemove', onImageMouseMove);
      document.removeEventListener('mouseup', onImageMouseUp);
    };

    document.addEventListener('mousemove', onImageMouseMove);
    document.addEventListener('mouseup', onImageMouseUp);
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
    // Nie zmieniaj kursora podczas aktywnego resize / przenoszenia obrazu
    if (this.tableResizeState || this.imageResizeState || this.imageMoveState) return;

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
    const rawTarget = event.target as Node | null;
    // event.target może być węzłem tekstowym (np. podczas drag-zaznaczenia tekstu)
    // — wtedy nie ma metody closest(). Bierzemy najbliższy element nadrzędny.
    const target: HTMLElement | null = rawTarget instanceof HTMLElement
      ? rawTarget
      : (rawTarget?.parentElement ?? null);
    const wrapper = target?.closest('.editor-image-wrapper') as HTMLElement | null;
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
    this.emitImageSelectionState();
  }

  private clearSelectedImage(): void {
    if (this.selectedImageWrapper) {
      this.selectedImageWrapper.classList.remove('selected');
      this.selectedImageWrapper = null;
      this.imageSelectionChange.emit(null);
    }
  }

  /**
   * Snapshots the currently selected image wrapper into the wire shape consumed by
   * d2-image-properties-panel. Reads dimensions from the rendered &lt;img&gt;'s bounding
   * rect (covers both inline width/height and zoom). Alignment is detected by inspecting
   * the parent paragraph's text-align (Word-like alignment lives on the paragraph, not
   * the image element).
   */
  private emitImageSelectionState(): void {
    const wrapper = this.selectedImageWrapper;
    if (!wrapper) {
      this.imageSelectionChange.emit(null);
      return;
    }
    const img = wrapper.querySelector('img') as HTMLImageElement | null;
    if (!img) {
      this.imageSelectionChange.emit(null);
      return;
    }
    const rect = img.getBoundingClientRect();
    const widthPx = Math.max(1, Math.round(rect.width));
    const heightPx = Math.max(1, Math.round(rect.height));
    const aspectRatio = heightPx > 0 ? widthPx / heightPx : 1;
    const para = wrapper.closest('p, div, h1, h2, h3, h4, h5, h6') as HTMLElement | null;
    const align = para?.style.textAlign as 'left' | 'center' | 'right' | '' | undefined;
    const rawMode = (wrapper.dataset['posMode'] ?? 'inline');
    const allowed = new Set(['inline', 'square', 'topBottom', 'front', 'behind']);
    const positionMode = (allowed.has(rawMode) ? rawMode : 'inline') as
      'inline' | 'square' | 'topBottom' | 'front' | 'behind';

    const borderWidthPx = parseInt(img.dataset['borderWidth'] ?? '0', 10);
    const borderColor = img.dataset['borderColor'] || '#000000';
    const borderStyleAttr = img.dataset['borderStyle'] as 'solid' | 'dashed' | 'dotted' | undefined;
    const border = {
      enabled: borderWidthPx > 0,
      color: borderColor,
      widthPx: borderWidthPx > 0 ? borderWidthPx : 1,
      style: borderStyleAttr ?? 'solid',
    };

    const crop = {
      left: parseFloat(img.dataset['cropL'] ?? '0') || 0,
      right: parseFloat(img.dataset['cropR'] ?? '0') || 0,
      top: parseFloat(img.dataset['cropT'] ?? '0') || 0,
      bottom: parseFloat(img.dataset['cropB'] ?? '0') || 0,
    };

    this.imageSelectionChange.emit({
      widthPx,
      heightPx,
      aspectRatio,
      alignment: align === 'left' || align === 'center' || align === 'right' ? align : null,
      positionMode,
      border,
      crop,
    });
  }

  /**
   * Switches the selected image between in-text (inline) and floating (front/behind).
   * Word-like semantics — floating uses position: absolute within the closest .page
   * container, with z-index controlling the front/behind layering. The mode is mirrored
   * to both the wrapper (CSS hook) and the inner &lt;img&gt; (so the DOCX exporter sees it
   * on the round-trippable element) so import / export can reproduce it.
   */
  setSelectedImagePositionMode(mode: 'inline' | 'square' | 'topBottom' | 'front' | 'behind'): void {
    const wrapper = this.selectedImageWrapper;
    if (!wrapper) return;
    const img = wrapper.querySelector('img') as HTMLImageElement | null;
    if (!img) return;

    // Always clear previous floating-only state, then re-apply per the new mode.
    delete wrapper.dataset['xPx'];
    delete wrapper.dataset['yPx'];
    img.removeAttribute('data-x-emu');
    img.removeAttribute('data-y-emu');
    wrapper.style.removeProperty('position');
    wrapper.style.removeProperty('left');
    wrapper.style.removeProperty('top');
    wrapper.style.removeProperty('z-index');
    wrapper.style.removeProperty('pointer-events');
    wrapper.style.removeProperty('float');
    wrapper.style.removeProperty('clear');
    wrapper.style.removeProperty('display');
    wrapper.style.removeProperty('margin');

    if (mode === 'inline') {
      delete wrapper.dataset['posMode'];
      delete img.dataset['posMode'];
    } else if (mode === 'square') {
      // Float-based wrap — text fills around the rectangular bounding box.
      wrapper.dataset['posMode'] = 'square';
      img.dataset['posMode'] = 'square';
      wrapper.style.float = 'left';
      wrapper.style.margin = '0 12px 8px 0';
    } else if (mode === 'topBottom') {
      // Block + clear:both — text sits above and below the image, not beside it.
      wrapper.dataset['posMode'] = 'topBottom';
      img.dataset['posMode'] = 'topBottom';
      wrapper.style.display = 'block';
      wrapper.style.clear = 'both';
      wrapper.style.margin = '8px auto';
    } else {
      // Anchor the wrapper at its current rendered position so the switch is visually
      // stable — no jump to (0,0). Coords are relative to the nearest paginated page.
      const page = wrapper.closest('.page, .editor-content, .header-editor-content, .footer-editor-content') as HTMLElement | null;
      const pageRect = page?.getBoundingClientRect();
      const rect = wrapper.getBoundingClientRect();
      const xPx = Math.max(0, Math.round(rect.left - (pageRect?.left ?? 0)));
      const yPx = Math.max(0, Math.round(rect.top - (pageRect?.top ?? 0)));
      this.applyFloatingPosition(wrapper, img, mode as 'front' | 'behind', xPx, yPx);
    }
    this.onContentChange();
    this.emitImageSelectionState();
  }

  /**
   * Border state — when enabled, applied as inline CSS on the img so it survives the
   * HTML round-trip; mirrored to data-border-* attributes for the DOCX exporter, which
   * maps them to a:ln in pic:spPr.
   */
  setSelectedImageBorder(border: {
    enabled: boolean; color: string; widthPx: number; style: 'solid' | 'dashed' | 'dotted';
  }): void {
    const wrapper = this.selectedImageWrapper;
    if (!wrapper) return;
    const img = wrapper.querySelector('img') as HTMLImageElement | null;
    if (!img) return;
    if (!border.enabled || border.widthPx <= 0) {
      img.style.removeProperty('border');
      img.style.removeProperty('border-width');
      img.style.removeProperty('border-style');
      img.style.removeProperty('border-color');
      delete img.dataset['borderWidth'];
      delete img.dataset['borderColor'];
      delete img.dataset['borderStyle'];
    } else {
      const w = Math.max(1, Math.min(20, Math.round(border.widthPx)));
      const color = /^#[0-9a-fA-F]{6}$/.test(border.color) ? border.color : '#000000';
      img.style.border = `${w}px ${border.style} ${color}`;
      img.dataset['borderWidth'] = String(w);
      img.dataset['borderColor'] = color;
      img.dataset['borderStyle'] = border.style;
    }
    this.onContentChange();
    this.emitImageSelectionState();
  }

  /**
   * Crop state — % from each side. Rendered with CSS clip-path: inset() so the
   * original raster stays intact. Persisted via data-crop-* attrs the DOCX exporter
   * maps to a:srcRect (Word's "trim from each side" model).
   */
  setSelectedImageCrop(crop: { left: number; right: number; top: number; bottom: number }): void {
    const wrapper = this.selectedImageWrapper;
    if (!wrapper) return;
    const img = wrapper.querySelector('img') as HTMLImageElement | null;
    if (!img) return;
    const clamp = (v: number) => Math.max(0, Math.min(95, Math.round(v)));
    const l = clamp(crop.left), r = clamp(crop.right), t = clamp(crop.top), b = clamp(crop.bottom);
    if (l === 0 && r === 0 && t === 0 && b === 0) {
      img.style.removeProperty('clip-path');
      delete img.dataset['cropL'];
      delete img.dataset['cropR'];
      delete img.dataset['cropT'];
      delete img.dataset['cropB'];
    } else {
      img.style.clipPath = `inset(${t}% ${r}% ${b}% ${l}%)`;
      img.dataset['cropL'] = String(l);
      img.dataset['cropR'] = String(r);
      img.dataset['cropT'] = String(t);
      img.dataset['cropB'] = String(b);
    }
    this.onContentChange();
    this.emitImageSelectionState();
  }

  resetSelectedImageCrop(): void {
    this.setSelectedImageCrop({ left: 0, right: 0, top: 0, bottom: 0 });
  }

  private applyFloatingPosition(
    wrapper: HTMLElement, img: HTMLImageElement,
    mode: 'front' | 'behind', xPx: number, yPx: number,
  ): void {
    wrapper.dataset['posMode'] = mode;
    img.dataset['posMode'] = mode;
    wrapper.dataset['xPx'] = String(xPx);
    wrapper.dataset['yPx'] = String(yPx);
    const EMU_PER_PX = 9525;
    img.setAttribute('data-x-emu', String(xPx * EMU_PER_PX));
    img.setAttribute('data-y-emu', String(yPx * EMU_PER_PX));
    wrapper.style.position = 'absolute';
    wrapper.style.left = `${xPx}px`;
    wrapper.style.top = `${yPx}px`;
    if (mode === 'behind') {
      wrapper.style.zIndex = '-1';
      // Allow text clicks to land through the wrapper when it sits visually behind.
      wrapper.style.pointerEvents = 'auto';
    } else {
      wrapper.style.zIndex = '10';
      wrapper.style.pointerEvents = 'auto';
    }
  }

  /** Apply a new width (px) to the selected image; height follows aspect when locked. */
  setSelectedImageWidth(widthPx: number, lockAspect = true): void {
    const wrapper = this.selectedImageWrapper;
    if (!wrapper) return;
    const img = wrapper.querySelector('img') as HTMLImageElement | null;
    if (!img) return;
    const safeWidth = Math.max(16, Math.round(widthPx));
    const aspect = img.naturalWidth > 0 && img.naturalHeight > 0
      ? img.naturalWidth / img.naturalHeight
      : (img.clientWidth > 0 && img.clientHeight > 0 ? img.clientWidth / img.clientHeight : 1);
    const newHeight = lockAspect ? Math.max(16, Math.round(safeWidth / aspect)) : img.clientHeight;
    this.applyImageSize(img, wrapper, safeWidth, newHeight);
  }

  /** Apply a new height (px); width follows aspect when locked. */
  setSelectedImageHeight(heightPx: number, lockAspect = true): void {
    const wrapper = this.selectedImageWrapper;
    if (!wrapper) return;
    const img = wrapper.querySelector('img') as HTMLImageElement | null;
    if (!img) return;
    const safeHeight = Math.max(16, Math.round(heightPx));
    const aspect = img.naturalWidth > 0 && img.naturalHeight > 0
      ? img.naturalWidth / img.naturalHeight
      : (img.clientWidth > 0 && img.clientHeight > 0 ? img.clientWidth / img.clientHeight : 1);
    const newWidth = lockAspect ? Math.max(16, Math.round(safeHeight * aspect)) : img.clientWidth;
    this.applyImageSize(img, wrapper, newWidth, safeHeight);
  }

  /** Re-stretches the image to its intrinsic aspect ratio at the current width. */
  resetSelectedImageAspect(): void {
    const wrapper = this.selectedImageWrapper;
    if (!wrapper) return;
    const img = wrapper.querySelector('img') as HTMLImageElement | null;
    if (!img || img.naturalWidth <= 0 || img.naturalHeight <= 0) return;
    const aspect = img.naturalWidth / img.naturalHeight;
    const widthPx = Math.max(16, Math.round(img.clientWidth));
    const heightPx = Math.max(16, Math.round(widthPx / aspect));
    this.applyImageSize(img, wrapper, widthPx, heightPx);
  }

  /** Align the paragraph containing the selected image (null = clear text-align). */
  setSelectedImageAlignment(value: 'left' | 'center' | 'right' | null): void {
    const wrapper = this.selectedImageWrapper;
    if (!wrapper) return;
    const para = wrapper.closest('p, div, h1, h2, h3, h4, h5, h6') as HTMLElement | null;
    if (!para) return;
    if (value === null) {
      para.style.removeProperty('text-align');
    } else {
      para.style.textAlign = value;
    }
    this.onContentChange();
    this.emitImageSelectionState();
  }

  /** Removes the currently selected image from the DOM and notifies the editor. */
  removeSelectedImage(): void {
    const wrapper = this.selectedImageWrapper;
    if (!wrapper) return;
    wrapper.remove();
    this.selectedImageWrapper = null;
    this.imageSelectionChange.emit(null);
    this.onContentChange();
  }

  private applyImageSize(
    img: HTMLImageElement, wrapper: HTMLElement, widthPx: number, heightPx: number,
  ): void {
    img.style.width = `${widthPx}px`;
    img.style.height = `${heightPx}px`;
    // Keep the EMU data-attributes in sync so the DOCX exporter re-emits the new size.
    const EMU_PER_PX = 9525;
    img.setAttribute('data-width-emu', String(widthPx * EMU_PER_PX));
    img.setAttribute('data-height-emu', String(heightPx * EMU_PER_PX));
    wrapper.style.width = `${widthPx}px`;
    this.onContentChange();
    this.emitImageSelectionState();
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
   * Publiczny trigger zmiany zawartości — używany przez parenta po edycji DOM,
   * której wysywig nie obserwuje (np. linijka modyfikuje margin-left bloku).
   */
  triggerContentChange(): void {
    this.onContentChange();
  }

  /**
   * Obsługa zmiany zawartości — automatycznie kieruje na body lub header/footer
   * zależnie od aktualnie edytowanej sekcji.
   */
  private onContentChange(): void {
    const section = this.editingSection();

    // Header / footer — emituj headerChange / footerChange
    if (section === 'header' && this.headerContentEl?.nativeElement) {
      const html = this.headerContentEl.nativeElement.innerHTML;
      this._headerHtml.set(html);
      this.emitHeaderFooterChanges();
      this.updateFormattingState();
      return;
    }
    if (section === 'footer' && this.footerContentEl?.nativeElement) {
      const html = this.footerContentEl.nativeElement.innerHTML;
      this._footerHtml.set(html);
      this.emitHeaderFooterChanges();
      this.updateFormattingState();
      return;
    }

    // Body
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
   * Sprawdza czy selekcja jest w edytorze (body lub header/footer).
   * Sprawdzamy WSZYSTKIE strony (multi-page) — nie tylko aktywną, bo
   * `editorContent` ref aktualizuje się dopiero na focusin, a `selectionchange`
   * fire'uje też dla strony, na której kursor już jest.
   */
  private isSelectionInEditor(selection: Selection): boolean {
    if (!selection.anchorNode) return false;
    const refs = this.pageEditorRefs?.toArray() ?? [];
    for (const ref of refs) {
      if (ref.nativeElement.contains(selection.anchorNode)) return true;
    }
    const body = this.editorContent?.nativeElement;
    if (body && body.contains(selection.anchorNode)) return true;
    const header = this.headerContentEl?.nativeElement;
    if (header && header.contains(selection.anchorNode)) return true;
    const footer = this.footerContentEl?.nativeElement;
    if (footer && footer.contains(selection.anchorNode)) return true;
    return false;
  }

  /**
   * Znacznik czasu, do którego najbliższe zdarzenie `paste` ma zostać potraktowane
   * jako „Wklej tylko tekst" (ustawiany skrótem Ctrl/Cmd+Shift+V w handleKeyboard).
   * Używamy okna czasowego zamiast bool, żeby nieużyty skrót nie „zatruł" kolejnego
   * zwykłego Ctrl+V.
   */
  private plainTextPasteUntil = 0;

  /** Zwraca true i konsumuje żądanie, jeśli bieżące wklejenie ma być czystym tekstem. */
  private consumePlainTextPasteRequest(): boolean {
    const requested = Date.now() < this.plainTextPasteUntil;
    this.plainTextPasteUntil = 0;
    return requested;
  }

  /**
   * Obsługa wklejania.
   *
   * Ctrl/Cmd+Shift+V → zawsze czysty tekst (preferuj text/plain, fallback z HTML).
   * Ctrl/Cmd+V → zachowaj formatowanie po sanitizacji; gdy brak HTML, zwykły tekst.
   */
  private handlePaste(e: ClipboardEvent): void {
    e.preventDefault();

    const clipboardData = e.clipboardData;
    if (!clipboardData) return;

    const html = clipboardData.getData('text/html');
    const plain = clipboardData.getData('text/plain');

    if (this.consumePlainTextPasteRequest()) {
      this.insertText(resolvePlainText(plain, html));
      return;
    }

    if (html) {
      this.insertHtml(this.sanitizeHtml(html));
    } else {
      this.insertText(normalizeWhitespace(plain));
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

    // Cross-page caret navigation: pages are separate contenteditable elements, so the browser
    // cannot move the caret to the next/previous page with Arrow keys. At a page boundary (last/
    // first line, collapsed caret, no modifiers) move it ourselves; otherwise let the browser
    // handle normal in-page navigation and selections (Shift/Ctrl/Alt untouched).
    if ((e.key === 'ArrowDown' || e.key === 'ArrowUp')
        && !e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey
        && this._tryMoveCaretAcrossPages(e.key === 'ArrowDown' ? 'down' : 'up')) {
      e.preventDefault();
      return;
    }

    if (e.ctrlKey || e.metaKey) {
      // Ctrl/Cmd+Shift+V — oznacz najbliższe wklejenie jako „tylko tekst".
      // Nie wołamy preventDefault: pozwalamy przeglądarce wywołać zdarzenie `paste`,
      // które obsłuży handlePaste z uwzględnieniem tej flagi.
      if (e.shiftKey && e.key.toLowerCase() === 'v') {
        this.plainTextPasteUntil = Date.now() + 1000;
        return;
      }

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
    const editor = this.getActiveEditor();
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
        this.selectAllContent();
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
   * Ustawia rozmiar czcionki.
   *
   * Wywoływane z toolbara (input + Enter / blur / +/-). Musi działać poprawnie
   * w dwóch scenariuszach:
   *  1. fokus jest w edytorze (klik na +/− z `preventDefault` na mousedown) — selekcja w edytorze istnieje,
   *  2. fokus przed chwilą był na inputcie toolbara — w `window.getSelection()` jest selekcja inputa
   *     (NIE w edytorze); musimy odtworzyć selekcję z `savedSelection` ZANIM zawołamy `editor.focus()`,
   *     bo `focus()` na contenteditable po utracie kursora ustawia caret na początku — i wstawienie
   *     spana lądowało na początku dokumentu.
   */
  setFontSize(size: number): void {
    const editor = this.getActiveEditor();
    if (!editor) return;

    this.currentFontSize = size;

    // Odtwórz selekcję ZANIM dotkniemy fokusu — jeżeli live selection jest poza edytorem
    // (np. user kliknął w input rozmiaru czcionki), a mamy zapisaną ostatnią pozycję.
    const live = window.getSelection();
    const liveInEditor = !!live && live.rangeCount > 0 && this.isSelectionInEditor(live);
    if (!liveInEditor && this.savedSelection) {
      this.restoreSelection();
    }

    editor.focus();

    let selection = window.getSelection();
    if ((!selection || selection.rangeCount === 0) && this.savedSelection) {
      this.restoreSelection();
      selection = window.getSelection();
    }

    if (!selection || selection.rangeCount === 0) {
      // Brak selekcji - ustaw dla następnego tekstu
      this.pendingFontSize = size;
      return;
    }

    const range = selection.getRangeAt(0);

    if (range.collapsed) {
      // Kursor bez zaznaczenia - ustaw rozmiar dla następnie wpisywanego tekstu.
      this.pendingFontSize = size;

      // Jeśli kursor siedzi wewnątrz istniejącego ZWS-spana (wstawionego przez
      // poprzedni klik +/-), aktualizuj jego font-size zamiast zagnieżdżać nowy.
      // Dzięki temu nie powstają stosy spanów z różnymi rozmiarami, które utrzymują
      // duży line-height nawet po powrocie do małego fontu.
      const containerEl = range.startContainer.nodeType === Node.TEXT_NODE
        ? range.startContainer.parentElement
        : range.startContainer as HTMLElement;
      const isInZwsSpan = containerEl instanceof HTMLSpanElement
        && containerEl.textContent === '\u200B'
        && containerEl.style.fontSize !== '';

      if (isInZwsSpan && containerEl) {
        // Zaktualizuj rozmiar istniejącego ZWS-spana.
        containerEl.style.fontSize = `${size}pt`;
        // Umieść kursor za ZWS (pozycja 1) — bez zmiany.
        const newRange = document.createRange();
        newRange.setStart(containerEl.firstChild!, Math.min(1, containerEl.firstChild!.textContent!.length));
        newRange.setEnd(containerEl.firstChild!, Math.min(1, containerEl.firstChild!.textContent!.length));
        selection.removeAllRanges();
        selection.addRange(newRange);
        this.savedSelection = newRange.cloneRange();
        this.updateFormattingState();
        return;
      }

      // Wstaw nowy ZWS-span z żądanym rozmiarem.
      const span = document.createElement('span');
      span.style.fontSize = `${size}pt`;
      span.innerHTML = '\u200B'; // Zero-width space

      range.insertNode(span);

      // Usuń stale ZWS-spany w tym samym bloku (z poprzednich kliknięć +/-).
      // Zostawiamy tylko właśnie wstawiony i ewentualnie spany z prawdziwą treścią.
      this.removeStaleZwsSpans(span);

      // Ustaw kursor wewnątrz spana
      const newRange = document.createRange();
      newRange.setStart(span.firstChild!, 1);
      newRange.setEnd(span.firstChild!, 1);
      selection.removeAllRanges();
      selection.addRange(newRange);

      // Zapisz nową pozycję karetki, żeby kolejne klik +/- znalazły żywą selekcję
      // a nie zdezaktualizowaną z poprzedniego zapisu.
      this.savedSelection = newRange.cloneRange();
      this.updateFormattingState();
      return;
    }

    // Jest zaznaczenie - zastosuj rozmiar do zaznaczonego tekstu
    this.applyFontSizeToSelection(size, selection, range);
    this.onContentChange();

    // Po zmianie selekcji w applyFontSizeToSelection — zaktualizuj zapisaną
    // selekcję, żeby kolejne kliknięcie +/− trafiało dokładnie na ten sam zakres.
    const after = window.getSelection();
    if (after && after.rangeCount > 0 && this.isSelectionInEditor(after)) {
      this.savedSelection = after.getRangeAt(0).cloneRange();
    }
    this.updateFormattingState();
  }

  /**
   * Aplikuje rozmiar czcionki do zaznaczenia - bez execCommand
   */
  private applyFontSizeToSelection(size: number, selection: Selection, range: Range): void {
    const editor = this.getActiveEditor();
    if (!editor) return;

    // Wyodrębnij zawartość zaznaczenia
    const fragment = range.extractContents();

    // insideStyledSpan = true gdy węzeł jest już dzieckiem spana z font-size;
    // węzły tekstowe w tym kontekście NIE powinny być owijane kolejnym spanem —
    // inaczej każdy cykl zwiększ/zmniejsz dodaje kolejną warstwę zagnieżdżenia,
    // a każda warstwa wnosi swój line-height do wysokości linii (ogromny odstęp).
    const processNode = (node: Node, insideStyledSpan = false): Node => {
      if (node.nodeType === Node.TEXT_NODE) {
        if (insideStyledSpan) {
          // Rodzic span już ma ustawiony font-size — klonuj tekst bez owijania.
          return node.cloneNode(true);
        }
        // Tekst na poziomie bloku (bezpośrednio w <p>, <li> itp.) — opakuj w span.
        const span = document.createElement('span');
        span.style.fontSize = `${size}pt`;
        span.textContent = node.textContent;
        return span;
      }

      if (node.nodeType === Node.ELEMENT_NODE) {
        const element = node as HTMLElement;

        // Jeśli to span lub font — nadpisz font-size, zachowaj inne style.
        if (element.tagName === 'SPAN' || element.tagName === 'FONT') {
          const newSpan = document.createElement('span');

          if (element.style.cssText) {
            newSpan.style.cssText = element.style.cssText;
          }
          newSpan.style.fontSize = `${size}pt`;

          if (element.tagName === 'FONT') {
            const fontEl = element as HTMLFontElement;
            if (fontEl.face) newSpan.style.fontFamily = fontEl.face;
            if (fontEl.color) newSpan.style.color = fontEl.color;
          }

          // Dzieci spana: przekaż flagę insideStyledSpan=true, żeby teksty
          // nie były ponownie owijane (eliminuje rosnące zagnieżdżenie).
          Array.from(element.childNodes).forEach(child => {
            newSpan.appendChild(processNode(child, true));
          });

          return newSpan;
        }

        // Dla innych elementów (b, i, u, sub, sup itp.) — zachowaj, przetwórz dzieci.
        // Dziedzicz flagę insideStyledSpan (gdy jesteśmy już w strefie spana z fontem).
        const clone = element.cloneNode(false) as HTMLElement;
        Array.from(element.childNodes).forEach(child => {
          clone.appendChild(processNode(child, insideStyledSpan));
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
    const editor = this.getActiveEditor();
    if (!editor) return;

    this.currentFontFamily = fontFamily;
    editor.focus();

    // Jeśli selekcja zaginęła (np. klik w select toolbara) – przywróć zapisaną
    let selection = window.getSelection();
    if ((!selection || selection.rangeCount === 0) && this.savedSelection) {
      this.restoreSelection();
      selection = window.getSelection();
    }

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
   * Usuwa "stale" ZWS-spany (zero-width space, font-size ustawiony, bez innej treści)
   * z tego samego bloku co `keepSpan`. Takie spany powstają przy klikaniu +/− bez
   * zaznaczenia — każde kliknięcie wstawiało nowy span, stare pozostawały w DOM
   * i utrzymywały line-height linii na poziomie największego fontu.
   */
  private removeStaleZwsSpans(keepSpan: HTMLElement): void {
    // Znajdź blok nadrzędny (p, li, div itp.)
    let block: HTMLElement | null = keepSpan.parentElement;
    while (block && !['P', 'LI', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'TD', 'TH'].includes(block.tagName)) {
      block = block.parentElement;
    }
    if (!block) return;

    const stale: HTMLElement[] = [];
    block.querySelectorAll<HTMLElement>('span[style*="font-size"]').forEach(el => {
      if (el === keepSpan) return;
      if (el.textContent === '\u200B' && (el.childNodes.length === 0 || (el.childNodes.length === 1 && el.firstChild!.nodeType === Node.TEXT_NODE))) {
        stale.push(el);
      }
    });
    stale.forEach(el => el.remove());
  }

  /**
   * Aplikuje rodzinę czcionki do zaznaczenia
   */
  private applyFontFamilyToSelection(fontFamily: string, selection: Selection, range: Range): void {
    const editor = this.getActiveEditor();
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
    const editor = this.getActiveEditor();
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
    const editor = this.getActiveEditor();
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
    const editor = this.getActiveEditor();
    if (!editor) return;

    editor.focus();
    document.execCommand('hiliteColor', false, color);
    this.onContentChange();
  }

  // Zapisana selekcja (do użycia gdy selekcja jest tracona przez kliknięcie na toolbar)
  private savedSelection: Range | null = null;

  /**
   * Ustawia fokus na edytorze (header/footer jeśli edytowany, inaczej body).
   */
  focus(): void {
    const editor = this.getActiveEditor();
    editor?.focus();
  }

  /**
   * Zapisuje aktualną selekcję - wywoływane przed focusout.
   * Akceptuje selekcje z body lub header/footer.
   */
  saveSelection(): void {
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      if (this.isSelectionInEditor(selection)) {
        this.savedSelection = range.cloneRange();
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
    const editor = this.getActiveEditor();
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
    const editor = this.getActiveEditor();
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
      const pages = this._splitHtmlIntoPages(previous);
      this.pageContents.set(pages);
      this._content.set(previous);
      this.contentChange.emit(previous);
      this._schedulePaginate('undo');
      
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
      
      const pages = this._splitHtmlIntoPages(next);
      this.pageContents.set(pages);
      this._content.set(next);
      this.contentChange.emit(next);
      this._schedulePaginate('redo');
      
      this.updateState();
    }
  }

  /**
   * Zapisuje do stosu undo
   */
  private saveToUndoStack(): void {
    const html = this.getContent();
    if (!html) return;
    
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
      // Wyznacz „element pod karetką" tak, żeby na granicach spanów
      // (np. selection.anchorNode wskazuje na sam <h1> z offsetem dziecka)
      // wejść w głąb do faktycznego text-node/span — inaczej odczyt computed
      // font-family/size wraca z <h1>/<p> zamiast z konkretnego runa.
      const range = selection.getRangeAt(0);
      let node: Node | null = range.startContainer;
      if (node && node.nodeType === Node.ELEMENT_NODE) {
        const el = node as HTMLElement;
        // Jeżeli zaznaczenie jest niezwinięte, weź dziecko z prawej strony granicy
        // (start zaznaczenia), żeby trafić w pierwszy zaznaczony span.
        // Dla zwiniętej karetki też wolimy „następne" dziecko — odpowiada wpisywaniu.
        const idx = Math.min(range.startOffset, el.childNodes.length - 1);
        node = el.childNodes[Math.max(idx, 0)] ?? el.lastChild ?? el;
        // Zejdź do pierwszego liścia (tekst lub element bez dzieci)
        while (node && node.nodeType === Node.ELEMENT_NODE && (node as HTMLElement).firstChild) {
          node = (node as HTMLElement).firstChild;
        }
      }
      let element: HTMLElement | null = null;
      if (node?.nodeType === Node.TEXT_NODE) {
        element = node.parentElement;
      } else if (node instanceof HTMLElement) {
        element = node;
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

    // Debug log — szczegółowy, żeby diagnozować mismatch toolbar vs DOM.
    {
      const dbg: Record<string, unknown> = { fontSize, fontFamily, blockFormat: currentBlockFormat };
      if (selection && selection.rangeCount > 0) {
        const r = selection.getRangeAt(0);
        const sc = r.startContainer;
        dbg['range'] = {
          collapsed: r.collapsed,
          startContainerType: sc.nodeType === Node.TEXT_NODE ? 'TEXT' : sc.nodeType === Node.ELEMENT_NODE ? `EL(${(sc as Element).tagName})` : sc.nodeType,
          startOffset: r.startOffset,
          startContainerParent: sc.parentElement ? `<${sc.parentElement.tagName.toLowerCase()} style="${sc.parentElement.getAttribute('style') ?? ''}">` : null,
        };
      }
      console.log('[updateFormattingState]', dbg);
    }

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
   * Aktualizuje ogólny stan edytora (LEKKI — bez getContent()).
   * isModified ustalamy na podstawie flagi `_isDirty`, którą czyścimy przy save/setContent.
   */
  private updateState(): void {
    const refs = this.pageEditorRefs?.toArray() ?? [];
    const fallback = this.editorContent?.nativeElement;
    if (refs.length === 0 && !fallback) return;

    let text = '';
    if (refs.length > 0) {
      text = refs.map(r => r.nativeElement.innerText || '').join('\n');
    } else if (fallback) {
      text = fallback.innerText || '';
    }
    // Word liczy te\u017c zawarto\u015b\u0107 nag\u0142\u00f3wka i stopki (raz, nie razy liczba stron),
    // dlatego do\u0142\u0105czamy je tu jednorazowo.
    const headerText = this.headerContentEl?.nativeElement?.innerText || '';
    const footerText = this.footerContentEl?.nativeElement?.innerText || '';
    if (headerText) text += '\n' + headerText;
    if (footerText) text += '\n' + footerText;

    // Algorytm zliczania s\u0142\u00f3w zbli\u017cony do MS Word:
    // - traktuje l\u0105czniki (-), apostrofy (', \u2019) i podkre\u015blniki wewn\u0105trz wyrazu jako spoiwo
    //   (np. "e-mail", "don\u2019t" \u2192 1 s\u0142owo)
    // - kropki i przecinki mi\u0119dzy cyframi traktuje jako cz\u0119\u015b\u0107 liczby ("3.14", "1,000" \u2192 1)
    // - separatorami s\u0105 m.in. spacje, tabulatory, my\u015blniki en/em (\u2013, \u2014), uko\u015bniki, znaki interpunkcyjne
    // - usuwa znaki o zerowej szeroko\u015bci oraz nasze sztuczne placeholdery p\u00f3l ({page}, {pages})
    const normalized = text
      .replace(/[\u200B-\u200D\uFEFF]/g, '')
      .replace(/\{page\}|\{pages\}/g, ' ');
    const matches = normalized.match(/[\p{L}\p{N}]+(?:[\-_'\u2019][\p{L}\p{N}]+|[.,]\p{N}+)*/gu);
    const wordCount = matches ? matches.length : 0;

    // Diagnostyka dostępna na żądanie z konsoli: window.__wcDebug() / window.__wcCopy().
    try {
      const w = window as unknown as Record<string, unknown>;
      w['__wcDebug'] = () => {
        const lines = normalized.split('\n');
        const perLine = lines.map((line, i) => {
          const m = line.match(/[\p{L}\p{N}]+(?:[\-_'\u2019][\p{L}\p{N}]+|[.,]\p{N}+)*/gu);
          return { i, count: m ? m.length : 0, text: line };
        });
        console.table(perLine.filter(p => p.count > 0));
        console.log('TOTAL:', wordCount);
        console.log('TEXT LENGTH:', normalized.length);
        return { wordCount, totalLines: lines.length, perLine, text: normalized };
      };
      w['__wcCopy'] = async () => {
        await navigator.clipboard.writeText(normalized);
        console.log('Skopiowano', normalized.length, 'znaków do schowka.');
      };
    } catch { /* ignore */ }

    this.editorState.update(state => ({
      ...state,
      isModified: this._isDirty,
      canUndo: this.undoStack.length > 1,
      canRedo: this.redoStack.length > 0,
      wordCount
    }));

    this.stateChange.emit(this.editorState());
  }

  // ===== MULTI-PAGE PAGINATION (MVP - Wariant A) =====

  /** Aktywacja strony przy focusin — przełącza editorContent ref i indeks aktywnej strony. */
  setActivePage(index: number, _ev: Event): void {
    const refs = this.pageEditorRefs?.toArray() ?? [];
    if (refs[index]) {
      this.editorContent = refs[index];
      this.activePageIndex.set(index);
    }
  }

  /** Input na konkretnej stronie — NIE ustawia pageContents (bo to rebinduje innerHTML wszystkich stron
   *  i kasuje kursor). Zmienione DOM żyje samo do czasu repaginacji. */
  onPageInput(index: number, _ev: Event): void {
    if (this._isRepaginating) return;
    this._isDirty = true;
    this._schedulePaginate('input');
    this._schedulePersist();
    // lekki update stanu (bez getContent)
    this.updateState();
    this.updateFormattingState();
  }

  /** Debounce ciężkich operacji (undo snapshot + emit contentChange) — 500 ms. */
  private _schedulePersist(): void {
    if (this._persistTimer) clearTimeout(this._persistTimer);
    this._persistTimer = setTimeout(() => {
      this._persistTimer = null;
      const html = this.getContent();
      this._isInternalUpdate = true;
      this._content.set(html);
      this.contentChange.emit(html);
      this._isInternalUpdate = false;
      // undo snapshot — tylko jeśli różni się od ostatniego
      if (this.undoStack.length === 0 || this.undoStack[this.undoStack.length - 1] !== html) {
        this.undoStack.push(html);
        if (this.undoStack.length > 100) this.undoStack.shift();
        this.redoStack = [];
      }
    }, 500);
  }

  /**
   * Dzieli HTML wejściowy na strony po znacznikach page-break, ale ZACHOWUJE marker na końcu
   * każdej strony (poza ostatnią). Bez tego `_repaginateNow` (re-paginacja wg wysokości) nie
   * widziała już bloku page-break i scalała treść z powrotem — manualny podział ginął wizualnie
   * (np. „PROTOKÓŁ…" lądował pod podpisami). Marker przeżywa też zapis (getContent → writer → w:br).
   */
  private _splitHtmlIntoPages(html: string): string[] {
    if (!html) return ['<p></p>'];
    const marker = '<div class="page-break"></div>';
    const parts = html.split(/<div[^>]*class=["'][^"']*\bpage-break\b[^"']*["'][^>]*>\s*<\/div>/gi);
    const pages = parts
      .map((p, i) => (i < parts.length - 1 ? p + marker : p))
      .map(p => p.trim())
      .filter(p => p.length > 0);
    return pages.length ? pages : ['<p></p>'];
  }

  /** Schedule paginacji z debouncingiem 300 ms. */
  private _schedulePaginate(_reason: string): void {
    if (this._paginateTimer) clearTimeout(this._paginateTimer);
    this._paginateTimer = setTimeout(() => {
      this._paginateTimer = null;
      this._repaginateNow();
    }, 600);
  }

  /**
   * Główna paginacja: bierze zawartość każdej strony, łączy, dzieli na kartki A4
   * z zachowaniem reguły "block-atomic" (paragraf w całości na 1 stronie),
   * z wyjątkiem tabel — te dzielimy między wierszami.
   */
  private _repaginateNow(): void {
    if (this._isRepaginating) return;
    const refs = this.pageEditorRefs?.toArray() ?? [];
    if (refs.length === 0) {
      return;
    }

    this._isRepaginating = true;
    try {
      const caret = this._saveGlobalCaret(refs);

      const allBlocks: HTMLElement[] = [];
      // Flatten: jeśli DOCX import wsadził treść w jeden wrapper <div>/<section>/<article>,
      // weź jego dzieci. Powtórz dla zagnieżdżonych wrapperów (max 3 poziomy).
      const flattenChildren = (el: HTMLElement, depth = 0): HTMLElement[] => {
        const kids = Array.from(el.children) as HTMLElement[];
        if (depth >= 3) return kids;
        // Jeśli mamy dokładnie jedno generyczne dziecko (DIV/SECTION/ARTICLE) bez
        // znaczących stylów blokowych, schodzimy w głąb.
        if (kids.length === 1) {
          const c = kids[0];
          if (c.tagName === 'DIV' || c.tagName === 'SECTION' || c.tagName === 'ARTICLE') {
            return flattenChildren(c, depth + 1);
          }
        }
        return kids;
      };
      for (const ref of refs) {
        const el = ref.nativeElement;
        const kids = flattenChildren(el);
        kids.forEach(child => {
          allBlocks.push(child.cloneNode(true) as HTMLElement);
        });
      }
      if (allBlocks.length === 0) {
        allBlocks.push(document.createElement('p'));
      }

      const marginTop = this.pageMargins().top * 37.8;
      const marginBottom = this.pageMargins().bottom * 37.8;
      const headerHpx = this._headerHeight() * 37.8;
      const footerHpx = this._footerHeight() * 37.8;
      const padTop = Math.max(0, marginTop - headerHpx);
      const padBottom = Math.max(0, marginBottom - footerHpx);
      const availableHeight = Math.max(100, this.PAGE_HEIGHT_PX - headerHpx - footerHpx - padTop - padBottom);

      const probeEd = refs[0].nativeElement;
      const cs = getComputedStyle(probeEd);
      const measurer = document.createElement('div');
      const innerW = probeEd.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
      measurer.style.cssText = `position:absolute;left:-99999px;top:0;width:${innerW}px;font-family:${cs.fontFamily};font-size:${cs.fontSize};line-height:${cs.lineHeight};visibility:hidden;`;
      document.body.appendChild(measurer);

      const measureBlock = (block: HTMLElement): number => {
        measurer.innerHTML = '';
        measurer.appendChild(block.cloneNode(true));
        return measurer.firstElementChild?.getBoundingClientRect().height ?? 0;
      };

      const pages: HTMLElement[][] = [[]];
      let currentHeight = 0;

      const pushBlock = (block: HTMLElement) => {
        const h = measureBlock(block);
        if (currentHeight + h > availableHeight && pages[pages.length - 1].length > 0) {
          pages.push([]);
          currentHeight = 0;
        }
        pages[pages.length - 1].push(block);
        currentHeight += h;
      };

      for (const block of allBlocks) {
        if (this._isPageBreakBlock(block)) {
          // Manualny page break: wymuś nową stronę. Marker zostaje na końcu bieżącej strony,
          // żeby przeżył zapis (getContent → writer → w:br type=page).
          pages[pages.length - 1].push(block);
          pages.push([]);
          currentHeight = 0;
          continue;
        }
        if (block.tagName === 'TABLE') {
          const split = this._splitTableForPagination(
            block as HTMLTableElement,
            Math.max(80, availableHeight - currentHeight),
            availableHeight,
            measurer
          );
          for (let i = 0; i < split.length; i++) {
            if (i > 0) {
              pages.push([]);
              currentHeight = 0;
            }
            pages[pages.length - 1].push(split[i]);
            currentHeight += measureBlock(split[i]);
          }
        } else {
          pushBlock(block);
        }
      }

      measurer.remove();

      const newPageContents = pages.map(blocks => {
        const tmp = document.createElement('div');
        blocks.forEach(b => tmp.appendChild(b));
        return tmp.innerHTML || '<p></p>';
      });

      const current = this.pageContents();
      const identical = current.length === newPageContents.length
        && current.every((v, i) => v === newPageContents[i]);
      if (!identical) {
        this.pageContents.set(newPageContents);
        setTimeout(() => this._restoreGlobalCaret(caret), 0);
      }
      this.calculatePages();
    } finally {
      this._isRepaginating = false;
    }
  }

  /** Dzieli tabelę między wierszami; zwraca array <table> dla kolejnych stron. */
  private _splitTableForPagination(
    table: HTMLTableElement,
    firstAvail: number,
    fullAvail: number,
    measurer: HTMLElement
  ): HTMLTableElement[] {
    const rows = Array.from(table.querySelectorAll('tr')) as HTMLTableRowElement[];
    if (rows.length === 0) return [table];

    const measureRows = (subset: HTMLTableRowElement[]): number => {
      const t = table.cloneNode(false) as HTMLTableElement;
      const tbody = document.createElement('tbody');
      subset.forEach(r => tbody.appendChild(r.cloneNode(true)));
      t.appendChild(tbody);
      measurer.innerHTML = '';
      measurer.appendChild(t);
      return t.getBoundingClientRect().height;
    };

    const chunks: HTMLTableRowElement[][] = [];
    let bucket: HTMLTableRowElement[] = [];
    let avail = firstAvail;
    for (const row of rows) {
      const tentative = [...bucket, row];
      const h = measureRows(tentative);
      if (h > avail && bucket.length > 0) {
        chunks.push(bucket);
        bucket = [row];
        avail = fullAvail;
      } else {
        bucket = tentative;
      }
    }
    if (bucket.length > 0) chunks.push(bucket);

    // Tag every fragment of THIS split with a shared logical id so serialization can merge them
    // back into one table (R-17). The colgroup (column widths) is cloned into each fragment so a
    // fragment renders with correct columns and the merged result keeps them.
    const splitId = chunks.length > 1 ? `st-${++this._splitTableSeq}` : null;
    const colgroup = table.querySelector('colgroup');

    return chunks.map(subset => {
      const t = table.cloneNode(false) as HTMLTableElement;
      if (colgroup) t.appendChild(colgroup.cloneNode(true));
      const tbody = document.createElement('tbody');
      subset.forEach(r => tbody.appendChild(r.cloneNode(true)));
      t.appendChild(tbody);
      if (splitId) t.setAttribute('data-split-table-id', splitId);
      return t;
    });
  }

  /** Sekwencja id dla fragmentów jednej logicznie podzielonej tabeli (R-17). */
  private _splitTableSeq = 0;

  /** Czy blok to manualny page break (div.page-break albo akapit zawierający tylko page-break). */
  private _isPageBreakBlock(el: HTMLElement): boolean {
    if (!el || el.nodeType !== 1) return false;
    if (el.classList?.contains('page-break')) return true;
    const nested = el.querySelector?.('.page-break');
    return !!nested && (el.textContent ?? '').trim().length === 0;
  }

  /**
   * Scala sąsiednie fragmenty tej samej logicznej tabeli (te same data-split-table-id) w jedną
   * tabelę — paginacja widoku dzieli tabelę między strony, ale zapis ma zawierać jedną tabelę.
   * Niezależne sąsiednie tabele (bez wspólnego id) NIE są scalane (R-17 / reguła 11).
   */
  private _mergeSplitTables(html: string): string {
    if (!html.includes('data-split-table-id')) return html;

    const tmp = document.createElement('div');
    tmp.innerHTML = html;

    const handled = new Set<Element>();
    tmp.querySelectorAll('table[data-split-table-id]').forEach(el => {
      const first = el as HTMLTableElement;
      if (handled.has(first)) return;
      const id = first.getAttribute('data-split-table-id');
      const targetBody = first.querySelector('tbody') ?? first;

      let next = first.nextElementSibling;
      while (next && next.tagName === 'TABLE' && next.getAttribute('data-split-table-id') === id) {
        handled.add(next);
        next.querySelectorAll('tr').forEach(tr => targetBody.appendChild(tr));
        const toRemove = next;
        next = next.nextElementSibling;
        toRemove.remove();
      }
      first.removeAttribute('data-split-table-id');
    });

    return tmp.innerHTML;
  }

  /**
   * Przenosi kursor na sąsiednią stronę (osobny contenteditable) na granicy strony. Zwraca true
   * (i obsłużono nawigację) tylko gdy: jest 2+ stron, karetka jest zwinięta, leży na skrajnej
   * linii bieżącej strony i istnieje sąsiednia strona. Inaczej false → przeglądarka robi normalną
   * nawigację wewnątrz strony. Nie ingeruje w zaznaczenia ani w nawigację wewnątrz strony.
   */
  private _tryMoveCaretAcrossPages(dir: 'down' | 'up'): boolean {
    const refs = this.pageEditorRefs?.toArray() ?? [];
    if (refs.length < 2) return false;
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || !sel.isCollapsed || !sel.anchorNode) return false;

    const node = sel.anchorNode;
    const curIdx = refs.findIndex(r => r.nativeElement === node || r.nativeElement.contains(node));
    if (curIdx < 0) return false;
    const cur = refs[curIdx].nativeElement;

    if (dir === 'down') {
      if (curIdx >= refs.length - 1 || !this._isCaretOnEdgeLine(cur, 'bottom')) return false;
      this._placeCaretAtEditorEdge(refs[curIdx + 1].nativeElement, 'start', curIdx + 1);
      return true;
    }
    if (curIdx <= 0 || !this._isCaretOnEdgeLine(cur, 'top')) return false;
    this._placeCaretAtEditorEdge(refs[curIdx - 1].nativeElement, 'end', curIdx - 1);
    return true;
  }

  /** Czy zwinięta karetka leży na górnej/dolnej skrajnej linii danego edytora strony. */
  private _isCaretOnEdgeLine(editor: HTMLElement, edge: 'top' | 'bottom'): boolean {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return false;
    const caret = sel.getRangeAt(0).cloneRange();
    caret.collapse(true);
    const cr = caret.getClientRects()[0] ?? caret.getBoundingClientRect();

    const bound = document.createRange();
    bound.selectNodeContents(editor);
    bound.collapse(edge === 'top');
    const rects = bound.getClientRects();
    const br = rects.length ? rects[edge === 'top' ? 0 : rects.length - 1] : bound.getBoundingClientRect();

    return edge === 'bottom' ? cr.bottom >= br.bottom - 2 : cr.top <= br.top + 2;
  }

  /** Ustawia karetkę na początku/końcu wskazanego edytora strony i aktywuje tę stronę. */
  private _placeCaretAtEditorEdge(editor: HTMLElement, edge: 'start' | 'end', index: number): void {
    editor.focus();
    const range = document.createRange();
    range.selectNodeContents(editor);
    range.collapse(edge === 'start');
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);

    const refs = this.pageEditorRefs?.toArray() ?? [];
    if (refs[index]) this.editorContent = refs[index];
    this.activePageIndex.set(index);
  }

  /** Zapamiętuje pozycję kursora jako globalny offset tekstowy (po wszystkich stronach). */
  private _saveGlobalCaret(refs: ElementRef<HTMLDivElement>[]): { offset: number } | null {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return null;
    const range = sel.getRangeAt(0);
    for (let i = 0; i < refs.length; i++) {
      const editor = refs[i].nativeElement;
      if (editor.contains(range.endContainer)) {
        const pre = range.cloneRange();
        pre.selectNodeContents(editor);
        pre.setEnd(range.endContainer, range.endOffset);
        let total = 0;
        for (let j = 0; j < i; j++) total += refs[j].nativeElement.innerText.length;
        return { offset: total + pre.toString().length };
      }
    }
    return null;
  }

  /** Przywraca kursor po globalnym offsetcie tekstowym. */
  private _restoreGlobalCaret(caret: { offset: number } | null): void {
    if (!caret) return;
    const refs = this.pageEditorRefs?.toArray() ?? [];
    let remaining = caret.offset;
    for (let i = 0; i < refs.length; i++) {
      const editor = refs[i].nativeElement;
      const len = editor.innerText.length;
      if (remaining <= len) {
        const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
        let node: Node | null;
        let r = remaining;
        while ((node = walker.nextNode())) {
          const nl = node.textContent?.length ?? 0;
          if (r <= nl) {
            const range = document.createRange();
            range.setStart(node, r);
            range.collapse(true);
            const sel = window.getSelection();
            if (sel) {
              sel.removeAllRanges();
              sel.addRange(range);
            }
            editor.focus();
            this.editorContent = refs[i];
            this.activePageIndex.set(i);
            return;
          }
          r -= nl;
        }
        return;
      }
      remaining -= len;
    }
  }

  /**
   * Pobiera zawartość HTML do ZAPISU — scala strony BEZ wstawiania znaczników
   * <div class="page-break"> na granicach stron.
   *
   * Granice stron w edytorze pochodzą z auto-paginacji wg wysokości (_repaginateNow),
   * a nie z intencji użytkownika. Wstawianie tu page-breaków materializowało paginację
   * widoku jako twarde <w:br type=page> w DOCX — dokument rósł (np. 4 → 7 stron), a Word
   * i tak paginuje sam. Akapity są block-atomic (nie dzielone), więc czysta konkatenacja
   * odtwarza treść; jawne page-breaki użytkownika (insertPageBreak) przeżywają jako
   * <div class="page-break"> wewnątrz treści strony. Patrz analiza orginał_GOOD vs zapisany_BAD.
   */
  getContent(): string {
    const refs = this.pageEditorRefs?.toArray() ?? [];
    if (refs.length === 0) {
      const fallback = this.editorContent?.nativeElement;
      return fallback ? this._serializeSingleEditor(fallback) : '';
    }

    const parts = refs.map(r => this._serializeSingleEditor(r.nativeElement));
    const merged = parts.filter(p => p && p.trim().length > 0).join('');
    // Scal fragmenty tej samej logicznej tabeli rozdzielonej przez paginację (R-17).
    return this._mergeSplitTables(merged);
  }

  /** Serializuje pojedynczy edytor strony do HTML (z zachowaniem wysokości tabel i odwijaniem image-wrapperów). */
  private _serializeSingleEditor(editor: HTMLDivElement): string {
    const tableRowHeights: Map<number, { heights: number[] }> = new Map();
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
   * Ustawia zawartość HTML — rozbija na strony po znacznikach <div class="page-break">.
   */
  setContent(html: string): void {
    const pages = this._splitHtmlIntoPages(html || '<p></p>');
    this.pageContents.set(pages);
    // Pages already carry their own page-break markers (see _splitHtmlIntoPages); plain join
    // avoids doubling them.
    this._content.set(pages.join(''));
    this._isDirty = false;
    // Po Angular re-render: opakuj obrazki, zapisz snapshot, repaginuj
    setTimeout(() => {
      this.wrapExistingImages();
      const merged = this.getContent();
      this.lastSavedContent = merged;
      this.undoStack = [merged];
      this.redoStack = [];
      this.updateState();
      this._schedulePaginate('setContent');
    }, 0);
  }

  /**
   * Opakowuje istniejące elementy <img> (bez wrappera) w editor-image-wrapper.
   * Domyślnie operuje na aktywnym edytorze body; można podać kontener (header/footer/strona).
   */
  private wrapExistingImages(container?: HTMLElement | null): void {
    const editor = container ?? this.editorContent?.nativeElement;
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

      // Keep the img's inline width/height untouched so the editor renders the
      // image at the same size as the read-only display (rule 14 — no apparent
      // resize on entering edit mode). The wrapper is inline-block, so it
      // shrinks to fit the image; clamping to container width is delegated to
      // the img's own `max-width: 100%`.
      wrapper.style.maxWidth = '100%';
      img.style.maxWidth = '100%';
      img.setAttribute('draggable', 'false');

      img.parentNode?.insertBefore(wrapper, img);
      wrapper.appendChild(img);

      // Restore floating positioning from the img's data attributes (set by the DOCX
      // importer when the source had wp:anchor, or persisted from a previous editing
      // session). Inline images are the default and need no further setup.
      const posMode = img.dataset['posMode'];
      if (posMode === 'front' || posMode === 'behind') {
        const xPx = Math.round((Number(img.getAttribute('data-x-emu') ?? 0)) / 9525);
        const yPx = Math.round((Number(img.getAttribute('data-y-emu') ?? 0)) / 9525);
        this.applyFloatingPosition(wrapper, img, posMode, xPx, yPx);
      } else if (posMode === 'square') {
        wrapper.dataset['posMode'] = 'square';
        wrapper.style.float = 'left';
        wrapper.style.margin = '0 12px 8px 0';
      } else if (posMode === 'topBottom') {
        wrapper.dataset['posMode'] = 'topBottom';
        wrapper.style.display = 'block';
        wrapper.style.clear = 'both';
        wrapper.style.margin = '8px auto';
      }

      // Restore border + crop from data attributes set by the DOCX importer or a
      // previous edit. They're applied as inline CSS on the <img> for round-trip.
      const bw = parseInt(img.dataset['borderWidth'] ?? '0', 10);
      if (bw > 0) {
        const bc = img.dataset['borderColor'] || '#000000';
        const bs = (img.dataset['borderStyle'] as 'solid' | 'dashed' | 'dotted') ?? 'solid';
        img.style.border = `${bw}px ${bs} ${bc}`;
      }
      const cl = parseFloat(img.dataset['cropL'] ?? '0') || 0;
      const cr = parseFloat(img.dataset['cropR'] ?? '0') || 0;
      const ct = parseFloat(img.dataset['cropT'] ?? '0') || 0;
      const cb = parseFloat(img.dataset['cropB'] ?? '0') || 0;
      if (cl > 0 || cr > 0 || ct > 0 || cb > 0) {
        img.style.clipPath = `inset(${ct}% ${cr}% ${cb}% ${cl}%)`;
      }

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
    this._isDirty = false;
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
   * Rozpoczyna edycję nagłówka. Jeśli przekazano event — kursor zostanie ustawiony
   * w miejscu kliknięcia. W przeciwnym wypadku trafi na koniec zawartości.
   */
  startEditingHeader(event?: MouseEvent): void {
    // Jeśli już edytujemy nagłówek, nie restartuj kursora (pozwól natywnemu klikowi go ustawić)
    if (this.editingSection() === 'header') return;
    const clickX = event?.clientX;
    const clickY = event?.clientY;
    this.editingSection.set('header');
    this.editingSectionChange.emit('header');
    setTimeout(() => {
      const el = this.headerContentEl?.nativeElement;
      if (el) {
        // Load the variant currently shown on page 0 so the editor matches the
        // displayed content (rule 10 — no apparent editing).
        el.innerHTML = this._differentFirstPage()
          ? this._headerFirstPageHtml()
          : this._headerHtml();
        this.wrapExistingImages(el);
        this.attachEditorListeners(el);
        el.focus();
        this.placeCaretAtPoint(el, clickX, clickY);
        this.observeActiveSectionGeometry();
      }
    }, 0);
  }

  /**
   * Rozpoczyna edycję stopki
   */
  startEditingFooter(event?: MouseEvent): void {
    if (this.editingSection() === 'footer') return;
    const clickX = event?.clientX;
    const clickY = event?.clientY;
    this.editingSection.set('footer');
    this.editingSectionChange.emit('footer');
    setTimeout(() => {
      const el = this.footerContentEl?.nativeElement;
      if (el) {
        el.innerHTML = this._differentFirstPage()
          ? this._footerFirstPageHtml()
          : this._footerHtml();
        this.wrapExistingImages(el);
        this.attachEditorListeners(el);
        el.focus();
        this.placeCaretAtPoint(el, clickX, clickY);
        this.observeActiveSectionGeometry();
      }
    }, 0);
  }

  /**
   * Ustawia kursor w punkcie (x,y) jeśli trafia w content edytora; w przeciwnym
   * wypadku ustawia kursor na końcu.
   */
  private placeCaretAtPoint(el: HTMLElement, x?: number, y?: number): void {
    if (typeof x === 'number' && typeof y === 'number') {
      const range = this.getRangeFromPoint(x, y);
      if (range && el.contains(range.startContainer)) {
        const sel = window.getSelection();
        sel?.removeAllRanges();
        sel?.addRange(range);
        return;
      }
    }
    this.placeCaretAtEnd(el);
  }

  /**
   * Ustawia kursor na końcu danego contenteditable.
   */
  private placeCaretAtEnd(el: HTMLElement): void {
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(false);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
  }

  /**
   * Kończy edycję nagłówka/stopki i wraca do głównej treści
   */
  stopEditingHeaderFooter(): void {
    if (this.editingSection() !== 'body') {
      this.editingSection.set('body');
      this.editingSectionChange.emit('body');
      this.stopObservingSectionGeometry();
    }
  }

  /**
   * Obsługa blur nagłówka
   */
  onHeaderBlur(): void {
    const content = this.headerContentEl?.nativeElement?.innerHTML || '';
    if (this._differentFirstPage()) {
      this._headerFirstPageHtml.set(content);
    } else {
      this._headerHtml.set(content);
    }
    // Emit the full header (incl. firstPage/even variants) — partial emit on blur
    // would drop the other variants in the parent's signal.
    this.emitHeaderFooterChanges();
    // Nie kończymy edycji od razu, pozwalamy na kliknięcie poza nagłówkiem
  }

  /**
   * Obsługa input nagłówka — emituje zmiany do parent (saveDocument używa headerContent)
   */
  onHeaderInput(event: Event): void {
    const content = (event.target as HTMLDivElement).innerHTML;
    // Write to the variant the user is actually editing on page 0; otherwise the
    // first-page edit would silently overwrite the default content.
    if (this._differentFirstPage()) {
      this._headerFirstPageHtml.set(content);
    } else {
      this._headerHtml.set(content);
    }
    this.invalidateHeaderFooterCache();
    this.emitHeaderFooterChanges();
  }

  /**
   * Obsługa blur stopki
   */
  onFooterBlur(): void {
    const content = this.footerContentEl?.nativeElement?.innerHTML || '';
    if (this._differentFirstPage()) {
      this._footerFirstPageHtml.set(content);
    } else {
      this._footerHtml.set(content);
    }
    this.emitHeaderFooterChanges();
  }

  /**
   * Obsługa input stopki — emituje zmiany do parent
   */
  onFooterInput(event: Event): void {
    const content = (event.target as HTMLDivElement).innerHTML;
    if (this._differentFirstPage()) {
      this._footerFirstPageHtml.set(content);
    } else {
      this._footerHtml.set(content);
    }
    this.invalidateHeaderFooterCache();
    this.emitHeaderFooterChanges();
  }

  /**
   * Wylicza zawartość nagłówka dla danej strony (używane wewnętrznie przez computed)
   */
  private _computeHeaderContent(pageIndex: number): string {
    // First-page variant wins over odd/even for page 0.
    if (this._differentFirstPage() && pageIndex === 0) {
      return this._headerFirstPageHtml();
    }
    // In odd/even mode, "default" (Word's reference type=default) IS the odd content
    // — the canonical source is _headerHtml. Even pages use the dedicated even signal.
    if (this._differentOddEven()) {
      const isOdd = (pageIndex + 1) % 2 === 1;
      return isOdd ? this._headerHtml() : this._headerEvenHtml();
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
    if (this._differentFirstPage() && pageIndex === 0) {
      content = this._footerFirstPageHtml();
    }
    // See _computeHeaderContent: "default" = odd; canonical source is _footerHtml.
    else if (this._differentOddEven()) {
      const isOdd = (pageIndex + 1) % 2 === 1;
      content = isOdd ? this._footerHtml() : this._footerEvenHtml();
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
   * Zapobiega utracie zaznaczenia w nagłówku/stopce przy klikaniu w pasek narzędzi.
   * Pozwala na fokus tylko dla pól input/select (np. checkbox "Inna pierwsza strona").
   */
  onHeaderFooterToolbarMouseDown(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    if (target.tagName !== 'INPUT' && target.tagName !== 'SELECT' && target.tagName !== 'TEXTAREA') {
      event.preventDefault();
    }
  }

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
   * Otwiera file picker i wstawia wybrany obrazek do aktualnie edytowanej
   * sekcji (header/footer/body). Używane przez menu opcji header/footer.
   */
  insertImageIntoActive(): void {
    this.showHeaderOptionsMenu.set(false);
    this.showFooterOptionsMenu.set(false);

    // Upewnij się że focus jest w aktywnym edytorze przed otwarciem dialogu
    const editor = this.getActiveEditor();
    if (editor) {
      editor.focus();
      this.saveSelection();
    }

    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    input.onchange = (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        const base64 = ev.target?.result as string;
        if (!base64) return;
        // Przywróć focus i selekcję — file dialog je gubi
        const editor2 = this.getActiveEditor();
        if (editor2) {
          editor2.focus();
          this.restoreSelection();
        }
        this.insertImage(base64, file.name);
      };
      reader.readAsDataURL(file);
    };
    input.click();
  }

  /**
   * Wstawia numer strony do nagłówka
   */
  insertPageNumbers(): void {
    this.showHeaderOptionsMenu.set(false);
    const el = this.headerContentEl?.nativeElement;
    if (!el) return;
    el.focus();
    document.execCommand('insertHTML', false, this._pageNumberHtml(el));
    this.onHeaderInput({ target: el } as any);
  }

  /**
   * Wstawia numer strony do stopki
   */
  insertPageNumbersFooter(): void {
    this.showFooterOptionsMenu.set(false);
    const el = this.footerContentEl?.nativeElement;
    if (!el) return;
    el.focus();
    document.execCommand('insertHTML', false, this._pageNumberHtml(el));
    this.onFooterInput({ target: el } as any);
  }

  /**
   * Buduje znacznik numeru strony dziedziczący rozmiar czcionki z treści nagłówka/stopki, zamiast
   * domyślnego rozmiaru edytora. Inaczej numer strony bywa większy niż reszta stopki (np. 10.5pt
   * kontenera vs 8pt runów stopki). Bez własnego rozmiaru ".page-number" dziedziczy przez CSS.
   */
  private _pageNumberHtml(editor: HTMLElement): string {
    const size = this._inlineFieldFontSize(editor);
    const style = size ? ` style="font-size:${size};"` : '';
    return `<span class="page-number"${style}>{page}</span>`;
  }

  private _inlineFieldFontSize(editor: HTMLElement): string | null {
    const span = editor.querySelector('span[style*="font-size"]') as HTMLElement | null;
    if (span?.style.fontSize) return span.style.fontSize;
    const cs = getComputedStyle(span ?? editor).fontSize;
    return cs && cs !== '0px' ? cs : null;
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

  /**
   * Zaznacza WYŁĄCZNIE treść dokumentu (edytory stron), a nie całą stronę przeglądarki.
   * `document.execCommand('selectAll')` w trybie read-only (brak fokusu w contenteditable)
   * zaznaczał całe `body` — łącznie z menu, toolbarem i paskiem statusu. Tutaj tworzymy
   * jeden ciągły zakres od początku pierwszego do końca ostatniego edytora strony.
   */
  selectAllContent(): void {
    const sel = window.getSelection();
    if (!sel) return;

    const refs = this.pageEditorRefs?.toArray() ?? [];
    const editors = refs.length > 0
      ? refs.map(r => r.nativeElement)
      : (this.editorContent?.nativeElement ? [this.editorContent.nativeElement] : []);
    if (editors.length === 0) return;

    const first = editors[0];
    const last = editors[editors.length - 1];
    const range = document.createRange();
    range.setStart(first, 0);
    range.setEnd(last, last.childNodes.length);

    sel.removeAllRanges();
    sel.addRange(range);
  }

  /**
   * Zwraca faktyczny układ stron (zmierzony z DOM, w px ze skalą zoomu): wysokość każdej
   * kartki i odstęp do następnej (separator). Strona z wysokim nagłówkiem rośnie ponad
   * min-height 1122px, więc do zrównania pionowej linijki per-strona potrzebne są realne
   * wymiary, nie stałe 1122.
   */
  getPageLayout(): { top: number; height: number }[] {
    const refs = this.pageEditorRefs?.toArray() ?? [];
    const pages = refs
      .map(r => r.nativeElement.closest('.page') as HTMLElement | null)
      .filter((p): p is HTMLElement => !!p);
    if (pages.length === 0) return [];
    const scrollEl = pages[0].closest('.editor-scroll-container') as HTMLElement | null;
    if (!scrollEl) return [];
    // Pozycja Y odpowiadająca offsetowi 0 w zawartości scrolla (niezależna od przewinięcia).
    const base = scrollEl.getBoundingClientRect().top - scrollEl.scrollTop;
    return pages.map(p => {
      const r = p.getBoundingClientRect();
      return { top: r.top - base, height: r.height };
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

    // Przeszukujemy WSZYSTKIE strony (nie tylko aktywną) — w kolejności dokumentu.
    const editors = (this.pageEditorRefs?.toArray() ?? []).map(r => r.nativeElement);
    if (editors.length === 0 && this.editorContent?.nativeElement) {
      editors.push(this.editorContent.nativeElement);
    }
    if (editors.length === 0 || !text) return { count: 0, currentIndex: -1 };

    const searchLower = text.toLowerCase();
    const matches: { node: Text; index: number }[] = [];

    for (const editor of editors) {
      const treeWalker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT, null);
      while (treeWalker.nextNode()) {
        const node = treeWalker.currentNode as Text;
        const content = node.textContent || '';
        let idx = content.toLowerCase().indexOf(searchLower);
        while (idx !== -1) {
          matches.push({ node, index: idx });
          idx = content.toLowerCase().indexOf(searchLower, idx + 1);
        }
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
   * Zwraca listę wyników wyszukiwania jako fragmenty tekstu (kontekst wokół trafienia)
   * — do panelu „Wyszukiwanie". Kolejność zgodna z `searchHighlights` (kolejność dokumentu).
   */
  getSearchSnippets(ctx = 32): { before: string; match: string; after: string }[] {
    return this.searchHighlights.map(mark => {
      const block = (mark.closest('p,li,h1,h2,h3,h4,h5,h6,td,th,blockquote') as HTMLElement | null)
        ?? mark.parentElement;
      const full = block?.textContent ?? mark.textContent ?? '';
      const match = mark.textContent ?? '';
      let prefixLen = 0;
      if (block) {
        const r = document.createRange();
        r.setStart(block, 0);
        r.setEndBefore(mark);
        prefixLen = r.toString().length;
      }
      const start = Math.max(0, prefixLen - ctx);
      const before = (start > 0 ? '…' : '') + full.substring(start, prefixLen);
      const afterEnd = prefixLen + match.length + ctx;
      const after = full.substring(prefixLen + match.length, afterEnd) + (afterEnd < full.length ? '…' : '');
      return { before, match, after };
    });
  }

  /**
   * Skacze do wyniku o danym indeksie (klik na liście w panelu) — podświetla i przewija.
   */
  goToMatch(index: number): { count: number; currentIndex: number } {
    if (index < 0 || index >= this.searchHighlights.length) {
      return { count: this.searchHighlights.length, currentIndex: this.currentHighlightIndex };
    }
    this.currentHighlightIndex = index;
    this.highlightCurrent();
    return { count: this.searchHighlights.length, currentIndex: index };
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
    // Agreguj WSZYSTKIE strony — zamiana może dotknąć innej strony niż aktywna.
    const html = this.getContent();
    if (!html && !this.editorContent?.nativeElement) return;
    this._isInternalUpdate = true;
    this._content.set(html);
    this.contentChange.emit(html);
    this._isInternalUpdate = false;
  }
}
