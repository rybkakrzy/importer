import { 
  Component, 
  ViewChild, 
  inject,
  signal,
  HostListener
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { WysiwygEditorComponent } from '../wysiwyg-editor/wysiwyg-editor';
import { EditorToolbarComponent } from '../editor-toolbar/editor-toolbar';
import { BarcodeDialogComponent } from '../barcode-dialog/barcode-dialog';
import { DocumentService } from '../../services/document.service';
import { 
  DocumentContent, 
  DocumentMetadata, 
  EditorState,
  EditorCommand,
  DocumentTemplate,
  PageMargins,
  PageSettings,
  MARGIN_PRESETS,
  DocumentStyle,
  HeaderFooterContent,
  DigitalSignatureInfo,
  SignDocumentRequest
} from '../../models/document.model';

/**
 * Główny komponent edytora dokumentów Word Online
 */
@Component({
  selector: 'app-document-editor',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    WysiwygEditorComponent,
    EditorToolbarComponent,
    BarcodeDialogComponent
  ],
  templateUrl: './document-editor.html',
  styleUrl: './document-editor.scss'
})
export class DocumentEditorComponent {
  @ViewChild(WysiwygEditorComponent) editor!: WysiwygEditorComponent;
  @ViewChild(EditorToolbarComponent) toolbar!: EditorToolbarComponent;

  private documentService = inject(DocumentService);

  // Stan dokumentu
  documentContent = signal<string>('<p></p>');
  documentMetadata = signal<DocumentMetadata>({
    title: 'Nowy dokument',
    created: new Date().toISOString(),
    modified: new Date().toISOString()
  });
  documentStyles = signal<DocumentStyle[]>([]);
  originalFileName = signal<string>('');
  
  // Nagłówek i stopka
  headerContent = signal<HeaderFooterContent>({ html: '', height: 1.25 });
  footerContent = signal<HeaderFooterContent>({ html: '', height: 1.25 });
  
  // Stan edytora
  editorState = signal<EditorState | null>(null);
  
  // Stan UI
  isLoading = signal(false);
  showMenu = signal(false);
  showEditMenu = signal(false);
  showFormatMenu = signal(false);
  showInsertMenu = signal(false);
  activeSubmenu = signal<string | null>(null);
  showTemplates = signal(false);
  showAbout = signal(false);
  showFindReplace = signal(false);
  showBarcodeDialog = signal(false);
  errorMessage = signal<string | null>(null);
  successMessage = signal<string | null>(null);

  // Menu kontekstowe
  showContextMenu = signal(false);
  contextMenuX = signal(0);
  contextMenuY = signal(0);
  contextSubmenu = signal<string | null>(null);
  contextMenuTargetCell = signal<HTMLElement | null>(null);

  // Menu Narzędzia
  showToolsMenu = signal(false);

  // Dialog Akapit
  showParagraphDialog = signal(false);
  paragraphDialogTab = signal<'indents' | 'breaks'>('indents');
  paragraphData = {
    alignment: 'left' as string,
    outlineLevel: 'body' as string,
    indentLeft: 0,
    indentRight: 0,
    specialIndent: 'none' as string,
    specialIndentBy: 1.27,
    mirrorIndents: false,
    spaceBefore: 0,
    spaceAfter: 8,
    lineSpacingType: 'multiple' as string,
    lineSpacingValue: 1.08,
    dontAddSpaceBetweenSameStyle: false,
    widowOrphanControl: true,
    keepWithNext: false,
    keepLinesTogether: false,
    pageBreakBefore: false
  };

  // Dialog Wstawianie tabeli
  showInsertTableDialog = signal(false);
  tableDialogData = {
    columns: 5,
    rows: 2,
    autoFitBehavior: 'fixed' as string,
    fixedWidth: 0, // 0 = Auto
    rememberDimensions: false
  };
  private savedTableDimensions: { columns: number; rows: number } | null = null;
  
  // Szablony
  templates = signal<DocumentTemplate[]>([]);

  // Zoom
  zoomLevel = signal(100);
  zoomLevels = [50, 75, 100, 125, 150, 200];

  // Toolbar tabeli
  isInTable = signal(false);
  activeTableCell = signal<HTMLTableCellElement | null>(null);
  activeTable = signal<HTMLTableElement | null>(null);

  // Zaznaczanie komórek tabeli (custom cell selection)
  selectedCells = signal<Set<HTMLTableCellElement>>(new Set());
  private cellSelectionStartCell: HTMLTableCellElement | null = null;
  private isCellSelecting = false;

  // Cieniowanie (shading) dropdown w toolbarze tabeli
  showShadingDropdown = signal(false);

  // Strony
  currentPage = signal(1);
  totalPages = signal(1);
  
  // Tooltip ze stronami przy scrollowaniu
  showPageIndicator = signal(false);
  private pageIndicatorTimeout?: ReturnType<typeof setTimeout>;

  // Ustawienia strony
  showPageSetup = signal(false);
  showMarginGuides = signal(true);
  pageSettings = signal<PageSettings>({
    margins: { top: 2.5, bottom: 2.5, left: 2.5, right: 2.5 },
    orientation: 'portrait',
    paperSize: 'a4'
  });
  marginPresets = MARGIN_PRESETS;

  // Dialog nagłówka i stopki
  showHeaderFooterDialog = signal(false);
  headerFooterDialogData = signal<{
    headerMargin: number;
    footerMargin: number;
    differentFirstPage: boolean;
    differentOddEven: boolean;
  }>({
    headerMargin: 1.27,
    footerMargin: 1.27,
    differentFirstPage: false,
    differentOddEven: false
  });

  // Dialog Właściwości dokumentu
  showPropertiesDialog = signal(false);
  propertiesData = signal<DocumentMetadata>({});

  // Dialog Podpisów cyfrowych
  showSignatureDialog = signal(false);
  signatureDialogTab = signal<'list' | 'sign'>('list');
  signatureData = {
    signerName: '' as string,
    signerTitle: '' as string,
    signerEmail: '' as string,
    reason: '' as string,
    certificateBase64: '' as string,
    certificatePassword: '' as string,
    certificateFileName: '' as string
  };

  // Baner podpisów
  documentSignatures = signal<DigitalSignatureInfo[]>([]);

  // Math dla template
  protected readonly Math = Math;

  constructor() {
    // Załaduj szablony
    this.loadTemplates();
  }

  /**
   * Aktualizuje tytuł dokumentu
   */
  updateTitle(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.documentMetadata.update(m => ({ ...m, title: input.value }));
  }

  /**
   * Ładuje dostępne szablony
   */
  private loadTemplates(): void {
    this.documentService.getTemplates().subscribe({
      next: (templates) => this.templates.set(templates),
      error: (err) => console.error('Błąd ładowania szablonów:', err)
    });
  }

  /**
   * Tworzy nowy dokument
   */
  newDocument(): void {
    if (this.editorState()?.isModified) {
      if (!confirm('Masz niezapisane zmiany. Czy na pewno chcesz utworzyć nowy dokument?')) {
        return;
      }
    }
    
    this.documentContent.set('<p></p>');
    this.documentMetadata.set({
      title: 'Nowy dokument',
      created: new Date().toISOString(),
      modified: new Date().toISOString()
    });
    this.documentStyles.set([]); // Reset stylów - toolbar użyje domyślnych
    this.originalFileName.set('');
    this.editor?.setContent('<p></p>');
    this.showMenu.set(false);
  }

  /**
   * Otwiera dokument z pliku
   */
  openDocument(): void {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.docx';
    
    input.onchange = (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (file) {
        this.loadDocument(file);
      }
    };
    
    input.click();
    this.showMenu.set(false);
  }

  /**
   * Ładuje dokument z pliku
   */
  private loadDocument(file: File): void {
    this.isLoading.set(true);
    this.errorMessage.set(null);
    
    this.documentService.openDocument(file).subscribe({
      next: (content) => {
        this.documentContent.set(content.html);
        this.documentMetadata.set(content.metadata);
        this.documentStyles.set(content.styles || []);
        this.originalFileName.set(file.name);
        
        // Wczytaj nagłówek i stopkę
        if (content.header) {
          this.headerContent.set({
            html: content.header.html || '',
            height: content.header.height || 1.25
          });
        }
        if (content.footer) {
          this.footerContent.set({
            html: content.footer.html || '',
            height: content.footer.height || 1.25
          });
        }
        
        if (this.editor) {
          this.editor.setContent(content.html);
        }
        
        // Wczytaj podpisy
        this.documentSignatures.set(content.metadata.signatures || []);
        
        this.showSuccess(`Otwarto dokument: ${file.name}`);
        this.isLoading.set(false);
      },
      error: (err) => {
        this.showError(err.message || 'Nie udało się otworzyć dokumentu');
        this.isLoading.set(false);
      }
    });
  }

  /**
   * Zapisuje dokument
   */
  saveDocument(): void {
    const html = this.editor?.getContent() || this.documentContent();
    const fileName = this.originalFileName() || `${this.documentMetadata().title || 'dokument'}.docx`;
    
    this.isLoading.set(true);
    
    this.documentService.downloadDocument(
      {
        html,
        originalFileName: fileName,
        metadata: this.documentMetadata(),
        header: this.headerContent(),
        footer: this.footerContent()
      },
      fileName
    );
    
    this.editor?.markAsSaved();
    this.showSuccess('Dokument został zapisany');
    this.isLoading.set(false);
    this.showMenu.set(false);
  }

  /**
   * Zapisuje dokument jako nowy plik
   */
  saveDocumentAs(): void {
    const newName = prompt('Podaj nazwę pliku:', this.documentMetadata().title || 'dokument');
    
    if (newName) {
      this.originalFileName.set(newName.endsWith('.docx') ? newName : `${newName}.docx`);
      this.documentMetadata.update(m => ({ ...m, title: newName }));
      this.saveDocument();
    }
  }

  /**
   * Otwiera szablon
   */
  openTemplate(templateId: string): void {
    this.isLoading.set(true);
    this.showTemplates.set(false);
    
    this.documentService.getTemplate(templateId).subscribe({
      next: (content) => {
        this.documentContent.set(content.html);
        this.documentMetadata.set(content.metadata);
        this.originalFileName.set('');
        
        if (this.editor) {
          this.editor.setContent(content.html);
        }
        
        this.isLoading.set(false);
      },
      error: (err) => {
        this.showError('Nie udało się załadować szablonu');
        this.isLoading.set(false);
      }
    });
  }

  /**
   * Obsługa komendy z toolbara
   */
  onCommand(event: { command: EditorCommand; value?: string }): void {
    this.editor?.executeCommand(event.command, event.value);
  }

  /**
   * Obsługa zmiany rozmiaru czcionki
   */
  onFontSizeChange(size: number): void {
    this.editor?.setFontSize(size);
  }

  /**
   * Obsługa zmiany rodziny czcionki
   */
  onFontFamilyChange(family: string): void {
    this.editor?.setFontFamily(family);
  }

  /**
   * Obsługa zmiany koloru tekstu
   */
  onTextColorChange(color: string): void {
    this.editor?.setTextColor(color);
  }

  /**
   * Obsługa zmiany koloru tła
   */
  onBackgroundColorChange(color: string): void {
    this.editor?.setBackgroundColor(color);
  }

  /**
   * Wstawia link
   */
  onInsertLink(event: { url: string; text?: string }): void {
    this.editor?.insertLink(event.url, event.text);
  }

  /**
   * Wstawia obraz
   */
  onInsertImage(): void {
    // Zapisz selekcję edytora przed otwarciem dialogu pliku (który zabiera focus)
    this.editor?.saveSelection();

    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    
    input.onchange = (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (file) {
        this.uploadAndInsertImage(file);
      }
    };
    
    input.click();
  }

  /**
   * Wgrywa i wstawia obraz
   */
  private uploadAndInsertImage(file: File): void {
    // Konwertuj lokalnie do base64 (bez wysyłania na serwer)
    const reader = new FileReader();
    reader.onload = (e) => {
      const base64 = e.target?.result as string;
      if (base64 && this.editor) {
        // Przywróć fokus i selekcję w edytorze przed wstawieniem
        this.editor.focus();
        this.editor.restoreSelection();
        this.editor.insertImage(base64, file.name);
      }
    };
    reader.readAsDataURL(file);
  }

  /**
   * Wstawia tabelę (szybkie wstawianie z podmenu)
   */
  onInsertTable(config: string): void {
    if (this.editor) {
      this.editor.insertTable(config);
      this.applyTableAutoFit(this.tableDialogData.autoFitBehavior, this.tableDialogData.fixedWidth);
    }
  }

  /**
   * Obsługa zmiany stylu dokumentu
   */
  onStyleChange(style: DocumentStyle): void {
    // Zastosuj pełny styl do zaznaczenia
    if (this.editor) {
      this.editor.applyDocumentStyle(style);
    }
  }

  // Przechowywane formatowanie do kopiowania
  private copiedFormat: any = null;

  /**
   * Kopiuje formatowanie z bieżącego zaznaczenia
   */
  onCopyFormat(): void {
    if (this.editor) {
      this.copiedFormat = this.editor.getCurrentFormatting();
    }
  }

  /**
   * Aplikuje skopiowane formatowanie do zaznaczenia
   */
  onPasteFormat(): void {
    if (this.editor && this.copiedFormat) {
      this.editor.applyFormatting(this.copiedFormat);
    }
  }

  /**
   * Wyszukiwanie tekstu w dokumencie
   */
  private lastSearchText = '';

  onSearchInDocument(event: { text: string; direction: 'next' | 'previous' }): void {
    if (!this.editor) return;

    let result: { count: number; currentIndex: number };

    if (event.text !== this.lastSearchText) {
      // Nowe wyszukiwanie
      this.lastSearchText = event.text;
      result = this.editor.searchText(event.text, event.direction);
    } else {
      // Nawigacja po istniejących wynikach
      result = event.direction === 'next' ? this.editor.findNext() : this.editor.findPrevious();
    }

    if (this.toolbar) {
      this.toolbar.updateSearchResults(result.count, result.currentIndex);
    }
  }

  onReplaceInDocument(event: { searchText: string; replaceText: string; all: boolean }): void {
    if (!this.editor) return;

    let result: { count: number; currentIndex: number };

    if (event.all) {
      result = this.editor.replaceAllMatches(event.replaceText);
    } else {
      result = this.editor.replaceCurrentMatch(event.replaceText);
    }

    if (this.toolbar) {
      this.toolbar.updateSearchResults(result.count, result.currentIndex);
    }
  }

  onClearSearch(): void {
    if (this.editor) {
      this.editor.clearSearchHighlights();
    }
    this.lastSearchText = '';
  }

  /**
   * Obsługa zmiany zawartości
   */
  onContentChange(html: string): void {
    this.documentContent.set(html);
    this.documentMetadata.update(m => ({
      ...m,
      modified: new Date().toISOString()
    }));
  }

  /**
   * Obsługa zmiany nagłówka
   */
  onHeaderChange(header: HeaderFooterContent): void {
    this.headerContent.set(header);
    this.documentMetadata.update(m => ({
      ...m,
      modified: new Date().toISOString()
    }));
  }

  /**
   * Obsługa zmiany stopki
   */
  onFooterChange(footer: HeaderFooterContent): void {
    this.footerContent.set(footer);
    this.documentMetadata.update(m => ({
      ...m,
      modified: new Date().toISOString()
    }));
  }

  /**
   * Obsługa zmiany stanu edytora
   */
  onStateChange(state: EditorState): void {
    this.editorState.set(state);
    this.detectTableContext();
  }

  /**
   * Wykrywa czy kursor jest wewnątrz tabeli
   */
  private detectTableContext(): void {
    const selection = window.getSelection();
    const editorEl = this.editor?.editorContent?.nativeElement;
    if (!selection || !editorEl || !selection.anchorNode) {
      this.isInTable.set(false);
      this.activeTableCell.set(null);
      this.activeTable.set(null);
      return;
    }
    const node = selection.anchorNode instanceof HTMLElement
      ? selection.anchorNode
      : selection.anchorNode.parentElement;
    if (!node || !editorEl.contains(node)) {
      this.isInTable.set(false);
      this.activeTableCell.set(null);
      this.activeTable.set(null);
      return;
    }
    const cell = node.closest('td, th') as HTMLTableCellElement | null;
    const table = node.closest('table') as HTMLTableElement | null;
    this.isInTable.set(!!cell && !!table);
    this.activeTableCell.set(cell);
    this.activeTable.set(table);
  }

  /**
   * Zmienia zoom
   */
  setZoom(level: number): void {
    this.zoomLevel.set(level);
  }

  /**
   * Obsługuje scroll aby aktualizować bieżącą stronę
   */
  onEditorScroll(event: Event): void {
    const container = event.target as HTMLElement;
    const scrollTop = container.scrollTop;
    const scale = this.zoomLevel() / 100;
    
    // Wysokość strony A4 w pikselach + margines
    const PAGE_HEIGHT = 1122;
    const PAGE_GAP = 40; // gap między stronami + separator
    const PADDING_TOP = 20; // padding containera
    
    // Oblicz wysokość strony z uwzględnieniem skali
    const scaledPageHeight = PAGE_HEIGHT * scale;
    const scaledGap = PAGE_GAP * scale;
    
    // Oblicz pozycję środka widocznego obszaru
    const viewportCenter = scrollTop + (container.clientHeight / 2) - (PADDING_TOP * scale);
    
    // Oblicz bieżącą stronę
    const currentPageNum = Math.floor(viewportCenter / (scaledPageHeight + scaledGap)) + 1;
    const maxPages = this.totalPages();
    
    this.currentPage.set(Math.min(Math.max(1, currentPageNum), maxPages));
    
    // Pokaż wskaźnik stron przy scrollowaniu (jeśli jest więcej niż 1 strona)
    if (maxPages > 1) {
      this.showPageIndicator.set(true);
      
      // Ukryj wskaźnik po 1.5 sekundy bez scrollowania
      if (this.pageIndicatorTimeout) {
        clearTimeout(this.pageIndicatorTimeout);
      }
      this.pageIndicatorTimeout = setTimeout(() => {
        this.showPageIndicator.set(false);
      }, 1500);
    }
  }

  /**
   * Obsługuje zmianę liczby stron
   */
  onPagesChange(pageCount: number): void {
    this.totalPages.set(pageCount);
    // Upewnij się, że currentPage nie jest większa niż totalPages
    if (this.currentPage() > pageCount) {
      this.currentPage.set(pageCount);
    }
  }

  /**
   * Drukuje dokument
   */
  printDocument(): void {
    window.print();
    this.showMenu.set(false);
  }

  /**
   * Pokazuje komunikat sukcesu
   */
  private showSuccess(message: string): void {
    this.successMessage.set(message);
    setTimeout(() => this.successMessage.set(null), 3000);
  }

  /**
   * Pokazuje komunikat błędu
   */
  private showError(message: string): void {
    this.errorMessage.set(message);
    setTimeout(() => this.errorMessage.set(null), 5000);
  }

  /**
   * Toggle menu
   */
  toggleMenu(): void {
    const wasOpen = this.showMenu();
    this.closeAllMenus();
    this.showMenu.set(!wasOpen);
  }

  /**
   * Toggle menu Edytuj
   */
  toggleEditMenu(): void {
    const wasOpen = this.showEditMenu();
    this.closeAllMenus();
    this.showEditMenu.set(!wasOpen);
  }

  /**
   * Toggle menu Format
   */
  toggleFormatMenu(): void {
    const wasOpen = this.showFormatMenu();
    this.closeAllMenus();
    this.showFormatMenu.set(!wasOpen);
  }

  /**
   * Toggle menu Wstaw
   */
  toggleInsertMenu(): void {
    const wasOpen = this.showInsertMenu();
    this.closeAllMenus();
    this.showInsertMenu.set(!wasOpen);
  }

  /**
   * Zamyka menu po kliknięciu poza obszarem menu
   */
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    // Sprawdź czy kliknięto w obszarze menu
    const isMenuArea = target.closest('.menu-bar') || 
                       target.closest('.dropdown-menu');
    const isShadingArea = target.closest('.shading-dropdown') || target.closest('.table-toolbar-btn-shading');
    // Jeśli kliknięto poza menu i poza cieniowaniem - zamknij
    if (!isMenuArea && !isShadingArea) {
      this.closeAllMenus();
    }
    // Zamknij dropdown cieniowania jeśli kliknięto poza nim
    if (!isShadingArea) {
      this.showShadingDropdown.set(false);
    }
  }

  // =====================
  // ZAZNACZANIE KOMÓREK TABELI (MULTI-CELL SELECTION)
  // =====================

  @HostListener('mousedown', ['$event'])
  onCellMouseDown(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    const editorEl = this.editor?.editorContent?.nativeElement;
    if (!editorEl) return;

    const cell = target.closest('td, th') as HTMLTableCellElement | null;
    if (cell && editorEl.contains(cell)) {
      this.cellSelectionStartCell = cell;
      this.isCellSelecting = false;
      // Wyczyść zaznaczenie jeśli nie trzymamy Shift
      if (!event.shiftKey) {
        this.clearCellSelection();
      }
    } else if (!target.closest('.table-toolbar') && !target.closest('.context-menu') && !target.closest('.shading-dropdown')) {
      this.cellSelectionStartCell = null;
      this.clearCellSelection();
    }
  }

  @HostListener('document:mousemove', ['$event'])
  onCellMouseMove(event: MouseEvent): void {
    if (!this.cellSelectionStartCell || !(event.buttons & 1)) return;

    const target = event.target as HTMLElement;
    const cell = target.closest('td, th') as HTMLTableCellElement | null;
    const editorEl = this.editor?.editorContent?.nativeElement;

    if (cell && editorEl && editorEl.contains(cell) && cell !== this.cellSelectionStartCell) {
      // Sprawdź czy obie komórki są w tej samej tabeli
      const startTable = this.cellSelectionStartCell.closest('table');
      const endTable = cell.closest('table');
      if (startTable && startTable === endTable) {
        this.isCellSelecting = true;
        event.preventDefault();
        // Wyczyść selekcję tekstową przeglądarki
        window.getSelection()?.removeAllRanges();
        this.selectCellRange(this.cellSelectionStartCell, cell);
      }
    }
  }

  @HostListener('document:mouseup', ['$event'])
  onCellMouseUp(event: MouseEvent): void {
    if (this.isCellSelecting) {
      this.isCellSelecting = false;
      // Wyczyść selekcję tekstową przeglądarki - zostawiamy custom cell selection
      window.getSelection()?.removeAllRanges();
    }
    this.cellSelectionStartCell = null;
  }

  /**
   * Zaznacza prostokątny zakres komórek od start do end
   */
  private selectCellRange(start: HTMLTableCellElement, end: HTMLTableCellElement): void {
    const table = start.closest('table') as HTMLTableElement;
    if (!table) return;

    const startPos = this.getCellPosition(start);
    const endPos = this.getCellPosition(end);
    if (!startPos || !endPos) return;

    const minRow = Math.min(startPos.rowIndex, endPos.rowIndex);
    const maxRow = Math.max(startPos.rowIndex, endPos.rowIndex);
    const minCol = Math.min(startPos.colIndex, endPos.colIndex);
    const maxCol = Math.max(startPos.colIndex, endPos.colIndex);

    const newSelection = new Set<HTMLTableCellElement>();
    for (let r = minRow; r <= maxRow; r++) {
      const row = table.rows[r];
      if (!row) continue;
      for (let c = minCol; c <= maxCol; c++) {
        if (c < row.cells.length) {
          newSelection.add(row.cells[c]);
        }
      }
    }
    this.applyCellSelection(newSelection);
  }

  /**
   * Stosuje wizualne zaznaczenie na podanych komórkach
   */
  private applyCellSelection(cells: Set<HTMLTableCellElement>): void {
    // Usuń stare zaznaczenie
    const prev = this.selectedCells();
    prev.forEach(c => c.classList.remove('table-cell-selected'));
    // Zaznacz nowe
    cells.forEach(c => c.classList.add('table-cell-selected'));
    this.selectedCells.set(cells);
  }

  /**
   * Czyści zaznaczenie komórek
   */
  clearCellSelection(): void {
    const prev = this.selectedCells();
    prev.forEach(c => c.classList.remove('table-cell-selected'));
    this.selectedCells.set(new Set());
  }

  /**
   * Obsługuje prawy przycisk myszy - menu kontekstowe
   */
  @HostListener('contextmenu', ['$event'])
  onContextMenu(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    // Pokaż menu kontekstowe tylko w obszarze edytora
    const isEditorArea = target.closest('.editor-main') || 
                         target.closest('app-wysiwyg-editor') ||
                         target.closest('.paper-container');
    if (isEditorArea) {
      event.preventDefault();
      this.closeAllMenus();
      this.contextSubmenu.set(null);

      // Wykryj czy kliknięto w komórkę tabeli
      const cellTarget = target.closest('td, th') as HTMLElement | null;
      this.contextMenuTargetCell.set(cellTarget);

      // Oblicz pozycję — upewnij się, że menu nie wychodzi poza ekran
      const menuWidth = 260;
      const menuHeight = 420;
      let x = event.clientX;
      let y = event.clientY;
      if (x + menuWidth > window.innerWidth) {
        x = window.innerWidth - menuWidth - 8;
      }
      if (y + menuHeight > window.innerHeight) {
        y = window.innerHeight - menuHeight - 8;
      }

      this.contextMenuX.set(x);
      this.contextMenuY.set(y);
      this.showContextMenu.set(true);
    }
  }

  /**
   * Zamyka wszystkie menu
   */
  closeAllMenus(): void {
    this.showMenu.set(false);
    this.showEditMenu.set(false);
    this.showFormatMenu.set(false);
    this.showInsertMenu.set(false);
    this.showToolsMenu.set(false);
    this.activeSubmenu.set(null);
    this.showTemplates.set(false);
    this.showContextMenu.set(false);
    this.contextSubmenu.set(null);
    this.showShadingDropdown.set(false);
  }

  /**
   * Ustawia aktywne podmenu
   */
  setActiveSubmenu(submenu: string | null): void {
    this.activeSubmenu.set(submenu);
  }

  // =====================
  // MENU EDYTUJ
  // =====================

  /**
   * Cofnij
   */
  undo(): void {
    this.editor?.executeCommand('undo');
    this.closeAllMenus();
  }

  /**
   * Ponów
   */
  redo(): void {
    this.editor?.executeCommand('redo');
    this.closeAllMenus();
  }

  /**
   * Wytnij
   */
  cut(): void {
    document.execCommand('cut');
    this.closeAllMenus();
  }

  /**
   * Kopiuj
   */
  copy(): void {
    document.execCommand('copy');
    this.closeAllMenus();
  }

  /**
   * Wklej
   */
  paste(): void {
    navigator.clipboard.readText().then(text => {
      this.editor?.insertText(text);
    }).catch(() => {
      document.execCommand('paste');
    });
    this.closeAllMenus();
  }

  /**
   * Wklej bez formatowania
   */
  pasteWithoutFormatting(): void {
    navigator.clipboard.readText().then(text => {
      this.editor?.insertText(text);
    });
    this.closeAllMenus();
  }

  /**
   * Zaznacz wszystko
   */
  selectAll(): void {
    this.editor?.executeCommand('selectAll');
    this.closeAllMenus();
  }

  /**
   * Usuń zaznaczenie
   */
  deleteSelection(): void {
    document.execCommand('delete');
    this.closeAllMenus();
  }

  /**
   * Otwiera dialog Znajdź i zamień
   */
  openFindReplace(): void {
    this.showFindReplace.set(true);
    this.closeAllMenus();
  }

  // =====================
  // MENU FORMATUJ
  // =====================

  /**
   * Pogrubienie
   */
  toggleBold(): void {
    this.editor?.executeCommand('bold');
    this.closeAllMenus();
  }

  /**
   * Kursywa
   */
  toggleItalic(): void {
    this.editor?.executeCommand('italic');
    this.closeAllMenus();
  }

  /**
   * Podkreślenie
   */
  toggleUnderline(): void {
    this.editor?.executeCommand('underline');
    this.closeAllMenus();
  }

  /**
   * Przekreślenie
   */
  toggleStrikethrough(): void {
    this.editor?.executeCommand('strikethrough');
    this.closeAllMenus();
  }

  /**
   * Indeks górny
   */
  toggleSuperscript(): void {
    this.editor?.executeCommand('superscript');
    this.closeAllMenus();
  }

  /**
   * Indeks dolny
   */
  toggleSubscript(): void {
    this.editor?.executeCommand('subscript');
    this.closeAllMenus();
  }

  /**
   * Zwiększ rozmiar czcionki
   */
  increaseFontSize(): void {
    const currentSize = this.editorState()?.fontSize || 11;
    this.editor?.setFontSize(currentSize + 1);
    this.closeAllMenus();
  }

  /**
   * Zmniejsz rozmiar czcionki
   */
  decreaseFontSize(): void {
    const currentSize = this.editorState()?.fontSize || 11;
    if (currentSize > 1) {
      this.editor?.setFontSize(currentSize - 1);
    }
    this.closeAllMenus();
  }

  /**
   * Zmień na wielkie litery
   */
  toUpperCase(): void {
    const selection = window.getSelection();
    if (selection && selection.toString()) {
      const text = selection.toString().toUpperCase();
      document.execCommand('insertText', false, text);
    }
    this.closeAllMenus();
  }

  /**
   * Zmień na małe litery
   */
  toLowerCase(): void {
    const selection = window.getSelection();
    if (selection && selection.toString()) {
      const text = selection.toString().toLowerCase();
      document.execCommand('insertText', false, text);
    }
    this.closeAllMenus();
  }

  /**
   * Zmień na Kapitaliki (każde słowo z wielkiej litery)
   */
  toTitleCase(): void {
    const selection = window.getSelection();
    if (selection && selection.toString()) {
      const text = selection.toString().replace(/\b\w/g, l => l.toUpperCase());
      document.execCommand('insertText', false, text);
    }
    this.closeAllMenus();
  }

  /**
   * Wyrównaj do lewej
   */
  alignLeft(): void {
    this.editor?.executeCommand('justifyLeft');
    this.closeAllMenus();
  }

  /**
   * Wyrównaj do środka
   */
  alignCenter(): void {
    this.editor?.executeCommand('justifyCenter');
    this.closeAllMenus();
  }

  /**
   * Wyrównaj do prawej
   */
  alignRight(): void {
    this.editor?.executeCommand('justifyRight');
    this.closeAllMenus();
  }

  /**
   * Wyjustuj
   */
  alignJustify(): void {
    this.editor?.executeCommand('justifyFull');
    this.closeAllMenus();
  }

  /**
   * Zwiększ wcięcie
   */
  increaseIndent(): void {
    this.editor?.executeCommand('indent');
    this.closeAllMenus();
  }

  /**
   * Zmniejsz wcięcie
   */
  decreaseIndent(): void {
    this.editor?.executeCommand('outdent');
    this.closeAllMenus();
  }

  /**
   * Interlinia pojedyncza
   */
  setLineSpacingSingle(): void {
    this.setLineSpacing(1);
  }

  /**
   * Interlinia 1.15
   */
  setLineSpacing115(): void {
    this.setLineSpacing(1.15);
  }

  /**
   * Interlinia 1.5
   */
  setLineSpacing15(): void {
    this.setLineSpacing(1.5);
  }

  /**
   * Interlinia podwójna
   */
  setLineSpacingDouble(): void {
    this.setLineSpacing(2);
  }

  /**
   * Ustawia interlinię
   */
  private setLineSpacing(value: number): void {
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      let block = range.startContainer as Node;
      if (block.nodeType === Node.TEXT_NODE) {
        block = block.parentNode!;
      }
      // Znajdź blok
      while (block && !['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI'].includes((block as HTMLElement).tagName)) {
        block = block.parentNode!;
      }
      if (block) {
        (block as HTMLElement).style.lineHeight = value.toString();
      }
    }
    this.closeAllMenus();
  }

  /**
   * Dodaj odstęp przed akapitem
   */
  addSpaceBefore(): void {
    this.setBlockSpacing('marginTop', '12pt');
  }

  /**
   * Usuń odstęp przed akapitem
   */
  removeSpaceBefore(): void {
    this.setBlockSpacing('marginTop', '0');
  }

  /**
   * Dodaj odstęp po akapicie
   */
  addSpaceAfter(): void {
    this.setBlockSpacing('marginBottom', '12pt');
  }

  /**
   * Usuń odstęp po akapicie
   */
  removeSpaceAfter(): void {
    this.setBlockSpacing('marginBottom', '0');
  }

  /**
   * Ustawia odstęp bloku
   */
  private setBlockSpacing(property: 'marginTop' | 'marginBottom', value: string): void {
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      let block = range.startContainer as Node;
      if (block.nodeType === Node.TEXT_NODE) {
        block = block.parentNode!;
      }
      while (block && !['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI'].includes((block as HTMLElement).tagName)) {
        block = block.parentNode!;
      }
      if (block) {
        (block as HTMLElement).style[property] = value;
      }
    }
    this.closeAllMenus();
  }

  /**
   * Lista punktowana
   */
  insertBulletList(): void {
    this.editor?.executeCommand('insertUnorderedList');
    this.closeAllMenus();
  }

  /**
   * Lista numerowana
   */
  insertNumberedList(): void {
    this.editor?.executeCommand('insertOrderedList');
    this.closeAllMenus();
  }

  /**
   * Wyczyść formatowanie
   */
  clearFormatting(): void {
    this.editor?.executeCommand('removeFormat');
    this.closeAllMenus();
  }

  /**
   * Otwiera dialog wstawiania kodu kreskowego / QR
   */
  openBarcodeDialog(): void {
    // Zapisz selekcję przed otwarciem dialogu - dialog zabierze fokus z edytora
    this.editor?.saveSelection();
    this.showBarcodeDialog.set(true);
    this.closeAllMenus();
  }

  /**
   * Wstawia kod kreskowy / QR do edytora
   */
  onInsertBarcode(event: { base64Image: string; content: string; showValueBelow: boolean }): void {
    if (this.editor) {
      // Przywróć fokus i selekcję w edytorze przed wstawieniem
      this.editor.focus();
      this.editor.restoreSelection();
      if (event.showValueBelow) {
        this.editor.insertBarcodeWithValue(event.base64Image, event.content);
      } else {
        this.editor.insertImage(event.base64Image, 'barcode');
      }
    }
    this.showBarcodeDialog.set(false);
  }

  /**
   * Zamyka dialog kodu kreskowego
   */
  closeBarcodeDialog(): void {
    this.showBarcodeDialog.set(false);
  }

  /**
   * Wstawia linię poziomą
   */
  insertHorizontalLine(): void {
    this.editor?.insertHorizontalRule();
    this.closeAllMenus();
  }

  /**
   * Wstawia podział strony
   */
  insertPageBreak(): void {
    this.editor?.insertPageBreak();
    this.closeAllMenus();
  }

  /**
   * Rozpoczyna edycję nagłówka
   */
  editHeader(): void {
    this.editor?.startEditingHeader();
  }

  /**
   * Rozpoczyna edycję stopki
   */
  editFooter(): void {
    this.editor?.startEditingFooter();
  }

  /**
   * Przełącza widok marginesów
   */
  toggleMarginGuides(): void {
    this.showMarginGuides.update(v => !v);
    this.closeAllMenus();
  }

  // =====================
  // ZNAJDŹ I ZAMIEŃ
  // =====================
  findText = signal('');
  replaceText = signal('');

  /**
   * Znajdź następny
   */
  findNext(): void {
    const text = this.findText();
    if (!text) return;

    // Użyj natywnej funkcji window.find
    (window as any).find(text);
  }

  /**
   * Zamień
   */
  replaceOne(): void {
    const findStr = this.findText();
    const replaceStr = this.replaceText();
    if (!findStr) return;

    const selection = window.getSelection();
    if (selection && selection.toString() === findStr) {
      document.execCommand('insertText', false, replaceStr);
      this.findNext();
    } else {
      this.findNext();
    }
  }

  /**
   * Zamień wszystko
   */
  replaceAll(): void {
    const findStr = this.findText();
    const replaceStr = this.replaceText();
    if (!findStr) return;

    const content = this.editor?.getContent() || '';
    const newContent = content.split(findStr).join(replaceStr);
    this.editor?.setContent(newContent);
    this.showSuccess(`Zamieniono wszystkie wystąpienia "${findStr}"`);
  }

  /**
   * Zamyka menu po kliknięciu poza
   */
  closeMenuOnOutsideClick(event: MouseEvent): void {
    this.showMenu.set(false);
    this.showTemplates.set(false);
  }

  /**
   * Otwiera dialog ustawień strony
   */
  openPageSetup(): void {
    this.showPageSetup.set(true);
    this.showMenu.set(false);
  }

  /**
   * Ustawia preset marginesów
   */
  applyMarginPreset(preset: { name: string; margins: PageMargins }): void {
    this.pageSettings.update(s => ({
      ...s,
      margins: { ...preset.margins }
    }));
  }

  /**
   * Aktualizuje pojedynczy margines
   */
  updateMargin(side: keyof PageMargins, value: number): void {
    this.pageSettings.update(s => ({
      ...s,
      margins: { ...s.margins, [side]: value }
    }));
  }

  /**
   * Pobiera style marginesów w pikselach
   */
  getMarginStyles(): { [key: string]: string } {
    const m = this.pageSettings().margins;
    // 1 cm = 37.8 px (przy 96 DPI)
    const cmToPx = 37.8;
    return {
      'padding-top': `${m.top * cmToPx}px`,
      'padding-bottom': `${m.bottom * cmToPx}px`,
      'padding-left': `${m.left * cmToPx}px`,
      'padding-right': `${m.right * cmToPx}px`
    };
  }

  /**
   * Zmienia orientację strony
   */
  setOrientation(orientation: 'portrait' | 'landscape'): void {
    this.pageSettings.update(s => ({ ...s, orientation }));
  }

  /**
   * Sprawdza czy preset marginesów jest aktywny
   */
  isPresetActive(preset: { name: string; margins: PageMargins }): boolean {
    const current = this.pageSettings().margins;
    return current.top === preset.margins.top &&
           current.bottom === preset.margins.bottom &&
           current.left === preset.margins.left &&
           current.right === preset.margins.right;
  }

  /**
   * Pobiera style dla podglądu presetu
   */
  getPresetPreviewStyle(preset: { name: string; margins: PageMargins }): { [key: string]: string } {
    const m = preset.margins;
    const scale = 2; // Skala dla miniaturki
    return {
      'padding': `${m.top * scale}px ${m.right * scale}px ${m.bottom * scale}px ${m.left * scale}px`
    };
  }

  /**
   * Pobiera style dla podglądu strony
   */
  getPreviewStyle(): { [key: string]: string } {
    const settings = this.pageSettings();
    const isLandscape = settings.orientation === 'landscape';
    
    return {
      'width': isLandscape ? '140px' : '100px',
      'height': isLandscape ? '100px' : '140px'
    };
  }

  /**
   * Pobiera style dla obszaru zawartości w podglądzie
   */
  getContentPreviewStyle(): { [key: string]: string } {
    const m = this.pageSettings().margins;
    const scale = 4; // Skala dla podglądu
    return {
      'padding-top': `${m.top * scale}px`,
      'padding-bottom': `${m.bottom * scale}px`,
      'padding-left': `${m.left * scale}px`,
      'padding-right': `${m.right * scale}px`
    };
  }

  /**
   * Aplikuje ustawienia strony do edytora
   */
  applyPageSettings(): void {
    // Marginesy zostaną przekazane do edytora przez style
    this.showPageSetup.set(false);
    this.showSuccess('Zastosowano ustawienia strony');
  }

  // ================================
  // Dialog Nagłówka i Stopki
  // ================================

  /**
   * Otwiera dialog nagłówka i stopki
   */
  onOpenHeaderFooterSettings(data: {
    headerMargin: number;
    footerMargin: number;
    differentFirstPage: boolean;
    differentOddEven: boolean;
  }): void {
    this.headerFooterDialogData.set(data);
    this.showHeaderFooterDialog.set(true);
  }

  /**
   * Zamyka dialog nagłówka i stopki
   */
  closeHeaderFooterDialog(): void {
    this.showHeaderFooterDialog.set(false);
  }

  /**
   * Aktualizuje dane dialogu
   */
  updateHeaderFooterDialogData(field: string, value: number | boolean): void {
    this.headerFooterDialogData.update(data => ({
      ...data,
      [field]: value
    }));
  }

  /**
   * Zatwierdza ustawienia nagłówka i stopki
   */
  applyHeaderFooterSettings(): void {
    const data = this.headerFooterDialogData();
    this.editor?.applyHeaderFooterSettings(data);
    this.closeHeaderFooterDialog();
  }

  // =====================
  // MENU KONTEKSTOWE
  // =====================

  closeContextMenu(): void {
    this.showContextMenu.set(false);
    this.contextSubmenu.set(null);
  }

  contextMenuToggleBold(): void {
    this.editor?.executeCommand('bold');
    this.closeContextMenu();
  }

  contextMenuToggleItalic(): void {
    this.editor?.executeCommand('italic');
    this.closeContextMenu();
  }

  contextMenuToggleUnderline(): void {
    this.editor?.executeCommand('underline');
    this.closeContextMenu();
  }

  contextMenuAlignLeft(): void {
    this.editor?.executeCommand('justifyLeft');
    this.closeContextMenu();
  }

  contextMenuAlignCenter(): void {
    this.editor?.executeCommand('justifyCenter');
    this.closeContextMenu();
  }

  contextMenuAlignRight(): void {
    this.editor?.executeCommand('justifyRight');
    this.closeContextMenu();
  }

  contextMenuAlignJustify(): void {
    this.editor?.executeCommand('justifyFull');
    this.closeContextMenu();
  }

  contextMenuSetLineSpacing(value: number): void {
    this.setLineSpacing(value);
    this.closeContextMenu();
  }

  contextMenuIncreaseIndent(): void {
    this.editor?.executeCommand('indent');
    this.closeContextMenu();
  }

  contextMenuDecreaseIndent(): void {
    this.editor?.executeCommand('outdent');
    this.closeContextMenu();
  }

  /**
   * Ustawia kolor tła komórki tabeli (context menu + toolbar)
   */
  setCellColor(color: string): void {
    // Użyj custom zaznaczenia lub aktywnej/target komórki
    const customSelected = this.selectedCells();
    if (customSelected.size > 0) {
      customSelected.forEach(c => (c as HTMLElement).style.backgroundColor = color);
    } else {
      const cell = this.contextMenuTargetCell() || this.activeTableCell();
      if (cell) {
        (cell as HTMLElement).style.backgroundColor = color;
      }
    }
    this.closeContextMenu();
    this.showShadingDropdown.set(false);
    this.notifyEditorChange();
  }

  /**
   * Czyści kolor tła komórki tabeli
   */
  clearCellColor(): void {
    const customSelected = this.selectedCells();
    if (customSelected.size > 0) {
      customSelected.forEach(c => (c as HTMLElement).style.backgroundColor = '');
    } else {
      const cell = this.contextMenuTargetCell() || this.activeTableCell();
      if (cell) {
        (cell as HTMLElement).style.backgroundColor = '';
      }
    }
    this.closeContextMenu();
    this.showShadingDropdown.set(false);
    this.notifyEditorChange();
  }

  /**
   * Pobiera zaznaczone komórki tabeli (z custom cell selection)
   */
  private getSelectedCells(selection: Selection, editor: HTMLElement): HTMLElement[] {
    // Użyj custom zaznaczenia komórek
    const customSelected = this.selectedCells();
    if (customSelected.size > 0) {
      return Array.from(customSelected);
    }

    // Fallback: aktywna komórka
    const cell = this.activeTableCell();
    return cell ? [cell] : [];
  }

  // =====================
  // MENU NARZĘDZIA
  // =====================

  toggleToolsMenu(): void {
    const wasOpen = this.showToolsMenu();
    this.closeAllMenus();
    this.showToolsMenu.set(!wasOpen);
  }

  // =====================
  // DIALOG AKAPIT
  // =====================

  openParagraphDialog(): void {
    this.closeAllMenus();
    this.readCurrentParagraphSettings();
    this.paragraphDialogTab.set('indents');
    this.showParagraphDialog.set(true);
  }

  closeParagraphDialog(): void {
    this.showParagraphDialog.set(false);
  }

  /**
   * Odczytuje bieżące ustawienia akapitu z zaznaczenia
   */
  private readCurrentParagraphSettings(): void {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;

    const range = selection.getRangeAt(0);
    let block = range.startContainer as Node;
    if (block.nodeType === Node.TEXT_NODE) {
      block = block.parentNode!;
    }
    while (block && !['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI'].includes((block as HTMLElement).tagName)) {
      block = block.parentNode!;
    }

    if (block) {
      const el = block as HTMLElement;
      const style = window.getComputedStyle(el);

      // Wyrównanie
      const textAlign = style.textAlign;
      if (textAlign === 'center') this.paragraphData.alignment = 'center';
      else if (textAlign === 'right' || textAlign === 'end') this.paragraphData.alignment = 'right';
      else if (textAlign === 'justify') this.paragraphData.alignment = 'justify';
      else this.paragraphData.alignment = 'left';

      // Wcięcia (px -> cm, 1cm ≈ 37.8px)
      const pxToCm = (px: number) => Math.round(px / 37.8 * 10) / 10;
      this.paragraphData.indentLeft = pxToCm(parseFloat(style.paddingLeft) || 0);
      this.paragraphData.indentRight = pxToCm(parseFloat(style.paddingRight) || 0);

      // Text-indent (wcięcie specjalne)
      const textIndent = parseFloat(style.textIndent) || 0;
      if (textIndent > 0) {
        this.paragraphData.specialIndent = 'firstLine';
        this.paragraphData.specialIndentBy = pxToCm(textIndent);
      } else if (textIndent < 0) {
        this.paragraphData.specialIndent = 'hanging';
        this.paragraphData.specialIndentBy = pxToCm(Math.abs(textIndent));
      } else {
        this.paragraphData.specialIndent = 'none';
      }

      // Odstępy (px -> pt, 1pt ≈ 1.333px)
      const pxToPt = (px: number) => Math.round(px / 1.333);
      this.paragraphData.spaceBefore = pxToPt(parseFloat(style.marginTop) || 0);
      this.paragraphData.spaceAfter = pxToPt(parseFloat(style.marginBottom) || 0);

      // Interlinia
      const lineHeight = style.lineHeight;
      if (lineHeight === 'normal') {
        this.paragraphData.lineSpacingType = 'single';
        this.paragraphData.lineSpacingValue = 1;
      } else {
        const lhValue = parseFloat(lineHeight);
        const fontSize = parseFloat(style.fontSize);
        const ratio = Math.round(lhValue / fontSize * 100) / 100;
        this.paragraphData.lineSpacingType = 'multiple';
        this.paragraphData.lineSpacingValue = ratio;
      }
    }
  }

  /**
   * Stosuje ustawienia akapitu
   */
  applyParagraphSettings(): void {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      this.closeParagraphDialog();
      return;
    }

    const range = selection.getRangeAt(0);
    let block = range.startContainer as Node;
    if (block.nodeType === Node.TEXT_NODE) {
      block = block.parentNode!;
    }
    while (block && !['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI'].includes((block as HTMLElement).tagName)) {
      block = block.parentNode!;
    }

    if (block) {
      const el = block as HTMLElement;
      const cmToPx = (cm: number) => cm * 37.8;

      // Wyrównanie
      el.style.textAlign = this.paragraphData.alignment;

      // Wcięcia
      el.style.paddingLeft = cmToPx(this.paragraphData.indentLeft) + 'px';
      el.style.paddingRight = cmToPx(this.paragraphData.indentRight) + 'px';

      // Wcięcie specjalne
      if (this.paragraphData.specialIndent === 'firstLine') {
        el.style.textIndent = cmToPx(this.paragraphData.specialIndentBy) + 'px';
      } else if (this.paragraphData.specialIndent === 'hanging') {
        el.style.textIndent = '-' + cmToPx(this.paragraphData.specialIndentBy) + 'px';
        el.style.paddingLeft = cmToPx(this.paragraphData.indentLeft + this.paragraphData.specialIndentBy) + 'px';
      } else {
        el.style.textIndent = '0';
      }

      // Odstępy
      const ptToPx = (pt: number) => pt * 1.333;
      el.style.marginTop = ptToPx(this.paragraphData.spaceBefore) + 'px';
      el.style.marginBottom = ptToPx(this.paragraphData.spaceAfter) + 'px';

      // Interlinia
      switch (this.paragraphData.lineSpacingType) {
        case 'single':
          el.style.lineHeight = '1';
          break;
        case '1.5':
          el.style.lineHeight = '1.5';
          break;
        case 'double':
          el.style.lineHeight = '2';
          break;
        case 'multiple':
          el.style.lineHeight = this.paragraphData.lineSpacingValue.toString();
          break;
        case 'atLeast':
          el.style.lineHeight = this.paragraphData.lineSpacingValue + 'pt';
          break;
        case 'exactly':
          el.style.lineHeight = this.paragraphData.lineSpacingValue + 'pt';
          break;
      }

      // Podziały strony
      if (this.paragraphData.pageBreakBefore) {
        el.style.pageBreakBefore = 'always';
      } else {
        el.style.pageBreakBefore = 'auto';
      }
    }

    this.closeParagraphDialog();
  }

  /**
   * Resetuje do domyślnych
   */
  resetParagraphDefaults(): void {
    this.paragraphData.alignment = 'left';
    this.paragraphData.outlineLevel = 'body';
    this.paragraphData.indentLeft = 0;
    this.paragraphData.indentRight = 0;
    this.paragraphData.specialIndent = 'none';
    this.paragraphData.specialIndentBy = 1.27;
    this.paragraphData.mirrorIndents = false;
    this.paragraphData.spaceBefore = 0;
    this.paragraphData.spaceAfter = 8;
    this.paragraphData.lineSpacingType = 'multiple';
    this.paragraphData.lineSpacingValue = 1.08;
    this.paragraphData.dontAddSpaceBetweenSameStyle = false;
    this.paragraphData.widowOrphanControl = true;
    this.paragraphData.keepWithNext = false;
    this.paragraphData.keepLinesTogether = false;
    this.paragraphData.pageBreakBefore = false;
  }

  /**
   * Jednostka interlinii
   */
  getLineSpacingUnit(): string {
    switch (this.paragraphData.lineSpacingType) {
      case 'atLeast':
      case 'exactly':
        return 'pkt';
      case 'multiple':
        return '';
      default:
        return '';
    }
  }

  /**
   * Interlinia dla podglądu
   */
  getPreviewLineHeight(): string {
    switch (this.paragraphData.lineSpacingType) {
      case 'single': return '1';
      case '1.5': return '1.5';
      case 'double': return '2';
      case 'multiple': return this.paragraphData.lineSpacingValue.toString();
      default: return '1.15';
    }
  }

  // =====================
  // DIALOG WSTAWIANIE TABELI
  // =====================

  openInsertTableDialog(): void {
    this.closeAllMenus();
    if (this.savedTableDimensions) {
      this.tableDialogData.columns = this.savedTableDimensions.columns;
      this.tableDialogData.rows = this.savedTableDimensions.rows;
    } else {
      this.tableDialogData.columns = 5;
      this.tableDialogData.rows = 2;
    }
    this.tableDialogData.autoFitBehavior = 'fixed';
    this.tableDialogData.fixedWidth = 0;
    this.showInsertTableDialog.set(true);
  }

  closeInsertTableDialog(): void {
    this.showInsertTableDialog.set(false);
  }

  onFixedWidthChange(value: string): void {
    if (value.toLowerCase() === 'auto' || value === '') {
      this.tableDialogData.fixedWidth = 0;
    } else {
      const num = parseFloat(value);
      if (!isNaN(num)) {
        this.tableDialogData.fixedWidth = num;
      }
    }
  }

  applyInsertTable(): void {
    const cols = Math.max(1, Math.min(63, this.tableDialogData.columns));
    const rows = Math.max(1, Math.min(500, this.tableDialogData.rows));

    if (this.tableDialogData.rememberDimensions) {
      this.savedTableDimensions = { columns: cols, rows: rows };
    }

    const config = `${cols}x${rows}`;

    if (this.editor) {
      this.editor.insertTable(config);
      this.applyTableAutoFit(this.tableDialogData.autoFitBehavior, this.tableDialogData.fixedWidth);
    }

    this.closeInsertTableDialog();
  }

  // =====================
  // TOOLBAR TABELI - OPERACJE
  // =====================

  /**
   * Powiadamia edytor o zmianach w DOM (wywołuje contentChange)
   */
  private notifyEditorChange(): void {
    const el = this.editor?.editorContent?.nativeElement;
    if (el) {
      el.dispatchEvent(new Event('input', { bubbles: true }));
    }
  }

  /** Pobiera indeks wiersza i kolumny aktywnej komórki */
  private getCellPosition(cell: HTMLTableCellElement): { rowIndex: number; colIndex: number } | null {
    const row = cell.parentElement as HTMLTableRowElement;
    if (!row) return null;
    const table = row.closest('table');
    if (!table) return null;
    const rows = Array.from(table.rows);
    const rowIndex = rows.indexOf(row);
    const colIndex = Array.from(row.cells).indexOf(cell);
    return { rowIndex, colIndex };
  }

  /** Wstaw wiersz powyżej */
  tableInsertRowAbove(): void {
    const cell = this.activeTableCell();
    const table = this.activeTable();
    if (!cell || !table) return;
    const pos = this.getCellPosition(cell);
    if (!pos) return;
    const colCount = table.rows[pos.rowIndex]?.cells.length || 1;
    const newRow = table.insertRow(pos.rowIndex);
    for (let i = 0; i < colCount; i++) {
      const td = newRow.insertCell();
      td.innerHTML = '&nbsp;';
      td.style.cssText = 'border:1px solid #ccc;padding:8px;min-width:30px;';
    }
    this.notifyEditorChange();
  }

  /** Wstaw wiersz poniżej */
  tableInsertRowBelow(): void {
    const cell = this.activeTableCell();
    const table = this.activeTable();
    if (!cell || !table) return;
    const pos = this.getCellPosition(cell);
    if (!pos) return;
    const colCount = table.rows[pos.rowIndex]?.cells.length || 1;
    const insertAt = pos.rowIndex + 1;
    const newRow = table.insertRow(insertAt < table.rows.length ? insertAt : -1);
    for (let i = 0; i < colCount; i++) {
      const td = newRow.insertCell();
      td.innerHTML = '&nbsp;';
      td.style.cssText = 'border:1px solid #ccc;padding:8px;min-width:30px;';
    }
    this.notifyEditorChange();
  }

  /** Wstaw kolumnę z lewej */
  tableInsertColLeft(): void {
    const cell = this.activeTableCell();
    const table = this.activeTable();
    if (!cell || !table) return;
    const pos = this.getCellPosition(cell);
    if (!pos) return;
    Array.from(table.rows).forEach(row => {
      const td = row.insertCell(Math.min(pos.colIndex, row.cells.length));
      td.innerHTML = '&nbsp;';
      td.style.cssText = 'border:1px solid #ccc;padding:8px;min-width:30px;';
    });
    this.notifyEditorChange();
  }

  /** Wstaw kolumnę z prawej */
  tableInsertColRight(): void {
    const cell = this.activeTableCell();
    const table = this.activeTable();
    if (!cell || !table) return;
    const pos = this.getCellPosition(cell);
    if (!pos) return;
    const insertAt = pos.colIndex + 1;
    Array.from(table.rows).forEach(row => {
      const td = row.insertCell(Math.min(insertAt, row.cells.length));
      td.innerHTML = '&nbsp;';
      td.style.cssText = 'border:1px solid #ccc;padding:8px;min-width:30px;';
    });
    this.notifyEditorChange();
  }

  /** Usuń wiersz */
  tableDeleteRow(): void {
    const cell = this.activeTableCell();
    const table = this.activeTable();
    if (!cell || !table) return;
    const pos = this.getCellPosition(cell);
    if (!pos) return;
    if (table.rows.length <= 1) {
      this.tableDeleteTable();
      return;
    }
    table.deleteRow(pos.rowIndex);
    this.notifyEditorChange();
  }

  /** Usuń kolumnę */
  tableDeleteCol(): void {
    const cell = this.activeTableCell();
    const table = this.activeTable();
    if (!cell || !table) return;
    const pos = this.getCellPosition(cell);
    if (!pos) return;
    if (table.rows[0]?.cells.length <= 1) {
      this.tableDeleteTable();
      return;
    }
    Array.from(table.rows).forEach(row => {
      if (pos.colIndex < row.cells.length) {
        row.deleteCell(pos.colIndex);
      }
    });
    this.notifyEditorChange();
  }

  /** Usuń tabelę */
  tableDeleteTable(): void {
    const table = this.activeTable();
    if (!table) return;
    table.parentNode?.removeChild(table);
    this.isInTable.set(false);
    this.activeTableCell.set(null);
    this.activeTable.set(null);
    this.notifyEditorChange();
  }

  /** Scal zaznaczone komórki */
  tableMergeCells(): void {
    const customSelected = this.selectedCells();
    const cells = customSelected.size > 0
      ? Array.from(customSelected)
      : (() => {
          const selection = window.getSelection();
          const editorEl = this.editor?.editorContent?.nativeElement;
          return (selection && editorEl) ? this.getSelectedCells(selection, editorEl) : [];
        })();
    if (cells.length < 2) return;

    // Zbierz treść i usuń komórki oprócz pierwszej
    const firstCell = cells[0] as HTMLTableCellElement;
    let mergedContent = '';
    let minRow = Infinity, maxRow = -1, minCol = Infinity, maxCol = -1;

    cells.forEach((c) => {
      const td = c as HTMLTableCellElement;
      const row = td.parentElement as HTMLTableRowElement;
      const table = row.closest('table')!;
      const ri = Array.from(table.rows).indexOf(row);
      const ci = Array.from(row.cells).indexOf(td);
      minRow = Math.min(minRow, ri);
      maxRow = Math.max(maxRow, ri);
      minCol = Math.min(minCol, ci);
      maxCol = Math.max(maxCol, ci + (td.colSpan || 1) - 1);
    });

    // Zbierz treści
    cells.forEach(c => {
      const txt = c.innerHTML.trim();
      if (txt && txt !== '&nbsp;' && txt !== '<br>') {
        mergedContent += (mergedContent ? ' ' : '') + txt;
      }
    });

    // Ustaw colspan/rowspan na pierwszej komórce
    const colSpan = maxCol - minCol + 1;
    const rowSpan = maxRow - minRow + 1;
    firstCell.colSpan = colSpan;
    firstCell.rowSpan = rowSpan;
    firstCell.innerHTML = mergedContent || '&nbsp;';

    // Usuń nadmiarowe komórki
    const table = this.activeTable();
    if (!table) return;
    for (let r = minRow; r <= maxRow; r++) {
      const row = table.rows[r];
      if (!row) continue;
      for (let c = row.cells.length - 1; c >= 0; c--) {
        const cell = row.cells[c];
        if (cell !== firstCell && cells.includes(cell)) {
          row.removeChild(cell);
        }
      }
    }
    this.clearCellSelection();
    this.notifyEditorChange();
  }

  /** Podziel komórkę */
  tableSplitCell(): void {
    const cell = this.activeTableCell();
    if (!cell) return;
    const table = this.activeTable();
    if (!table) return;

    const cs = cell.colSpan || 1;
    const rs = cell.rowSpan || 1;

    if (cs <= 1 && rs <= 1) {
      // Komórka nie jest scalona - podziel na 2 kolumny
      const pos = this.getCellPosition(cell);
      if (!pos) return;
      cell.colSpan = 1;
      Array.from(table.rows).forEach((row, ri) => {
        if (ri === pos.rowIndex) {
          const newTd = row.insertCell(pos.colIndex + 1);
          newTd.innerHTML = '&nbsp;';
          newTd.style.cssText = 'border:1px solid #ccc;padding:8px;min-width:30px;';
        } else {
          const newTd = row.insertCell(Math.min(pos.colIndex + 1, row.cells.length));
          newTd.innerHTML = '&nbsp;';
          newTd.style.cssText = 'border:1px solid #ccc;padding:8px;min-width:30px;';
        }
      });
    } else {
      // Komórka jest scalona - cofnij scalenie
      const pos = this.getCellPosition(cell);
      if (!pos) return;
      cell.colSpan = 1;
      cell.rowSpan = 1;
      // Dodaj brakujące komórki w bieżącym wierszu
      const row = cell.parentElement as HTMLTableRowElement;
      for (let c = 1; c < cs; c++) {
        const newTd = row.insertCell(Array.from(row.cells).indexOf(cell) + 1);
        newTd.innerHTML = '&nbsp;';
        newTd.style.cssText = 'border:1px solid #ccc;padding:8px;min-width:30px;';
      }
      // Dodaj brakujące komórki w kolejnych wierszach
      for (let r = 1; r < rs; r++) {
        const targetRow = table.rows[pos.rowIndex + r];
        if (!targetRow) continue;
        for (let c = 0; c < cs; c++) {
          const insertIdx = Math.min(pos.colIndex, targetRow.cells.length);
          const newTd = targetRow.insertCell(insertIdx);
          newTd.innerHTML = '&nbsp;';
          newTd.style.cssText = 'border:1px solid #ccc;padding:8px;min-width:30px;';
        }
      }
    }
    this.notifyEditorChange();
  }

  /** Podziel tabelę (dzieli nad bieżącym wierszem) */
  tableSplitTable(): void {
    const cell = this.activeTableCell();
    const table = this.activeTable();
    if (!cell || !table) return;
    const pos = this.getCellPosition(cell);
    if (!pos || pos.rowIndex === 0) return;

    // Utwórz nową tabelę z wierszami od bieżącego w dół
    const newTable = document.createElement('table');
    newTable.style.cssText = table.style.cssText;
    const rowsToMove = Array.from(table.rows).slice(pos.rowIndex);
    rowsToMove.forEach(row => newTable.appendChild(row));

    // Wstaw paragraf separator i nową tabelę po starej
    const separator = document.createElement('p');
    separator.innerHTML = '&nbsp;';
    table.parentNode?.insertBefore(separator, table.nextSibling);
    separator.parentNode?.insertBefore(newTable, separator.nextSibling);
    this.notifyEditorChange();
  }

  /** Autodopasowanie - do zawartości */
  tableAutoFitContents(): void {
    const table = this.activeTable();
    if (!table) return;
    table.style.width = 'auto';
    table.style.tableLayout = 'auto';
    table.querySelectorAll('td, th').forEach(c => {
      (c as HTMLElement).style.width = '';
    });
    this.notifyEditorChange();
  }

  /** Autodopasowanie - do okna */
  tableAutoFitWindow(): void {
    const table = this.activeTable();
    if (!table) return;
    table.style.width = '100%';
    table.style.tableLayout = 'auto';
    table.querySelectorAll('td, th').forEach(c => {
      (c as HTMLElement).style.width = '';
    });
    this.notifyEditorChange();
  }

  /** Stała szerokość kolumn */
  tableFixedWidth(): void {
    const table = this.activeTable();
    if (!table) return;
    table.style.width = '100%';
    table.style.tableLayout = 'fixed';
    this.notifyEditorChange();
  }

  /** Rozłóż wiersze równomiernie */
  tableDistributeRows(): void {
    const table = this.activeTable();
    if (!table) return;
    Array.from(table.rows).forEach(row => {
      row.style.height = '';
      Array.from(row.cells).forEach(cell => {
        cell.style.height = '';
      });
    });
    this.notifyEditorChange();
  }

  /** Rozłóż kolumny równomiernie */
  tableDistributeCols(): void {
    const table = this.activeTable();
    if (!table) return;
    const colCount = table.rows[0]?.cells.length || 1;
    const w = Math.floor(100 / colCount);
    Array.from(table.rows).forEach(row => {
      Array.from(row.cells).forEach(cell => {
        cell.style.width = w + '%';
      });
    });
    this.notifyEditorChange();
  }

  /** Wyświetl/ukryj linie siatki */
  showTableGridLines = signal(true);
  tableToggleGridLines(): void {
    this.showTableGridLines.update(v => !v);
    const table = this.activeTable();
    if (!table) return;
    if (this.showTableGridLines()) {
      table.querySelectorAll('td, th').forEach(c => {
        (c as HTMLElement).style.borderColor = '#ccc';
      });
    } else {
      table.querySelectorAll('td, th').forEach(c => {
        (c as HTMLElement).style.borderColor = 'transparent';
      });
    }
  }

  /** Pozycja dropdown cieniowania */
  shadingDropdownX = signal(0);
  shadingDropdownY = signal(0);

  /** Toggle dropdown cieniowania w toolbarze tabeli */
  toggleShadingDropdown(event: MouseEvent): void {
    const btn = (event.target as HTMLElement).closest('.table-toolbar-btn-shading') as HTMLElement;
    if (btn) {
      const rect = btn.getBoundingClientRect();
      this.shadingDropdownX.set(rect.left);
      this.shadingDropdownY.set(rect.bottom + 4);
    }
    this.showShadingDropdown.update(v => !v);
  }

  /** Kolory do palety cieniowania */
  shadingColors = [
    '#FFFFFF', '#F2F2F2', '#D9D9D9', '#BFBFBF', '#A6A6A6', '#808080', '#595959', '#404040', '#262626', '#000000',
    '#FFF2CC', '#FFE599', '#FFD966', '#FFC000', '#BF9000', '#806000', '#FCE4D6', '#F8CBAD', '#F4B084', '#ED7D31',
    '#C55A11', '#833C0B', '#D6E4F0', '#B4C6E7', '#8DB4E2', '#4472C4', '#2F5597', '#1F3864', '#E2EFDA', '#C6EFCE',
    '#A9D18E', '#70AD47', '#548235', '#375623', '#F8D7DA', '#F5C6CB', '#E8A0A0', '#FF0000', '#C00000', '#800000',
    '#E6D5F5', '#D0B8E8', '#B490D0', '#7030A0', '#5B259A', '#3B1770'
  ];

  /**
   * Stosuje autodopasowanie do ostatnio wstawionej tabeli
   */
  private applyTableAutoFit(behavior: string, fixedWidth: number): void {
    setTimeout(() => {
      const editorEl = this.editor?.editorContent?.nativeElement;
      const tables = editorEl?.querySelectorAll('table');
      if (tables && tables.length > 0) {
        const lastTable = tables[tables.length - 1] as HTMLTableElement;
        switch (behavior) {
          case 'fixed':
            if (fixedWidth > 0) {
              lastTable.style.width = '';
              lastTable.style.tableLayout = 'fixed';
              const widthPx = fixedWidth * 37.8;
              lastTable.querySelectorAll('td, th').forEach((cell) => {
                (cell as HTMLElement).style.width = widthPx + 'px';
              });
            } else {
              lastTable.style.width = '100%';
              lastTable.style.tableLayout = 'fixed';
            }
            break;
          case 'contents':
            lastTable.style.width = 'auto';
            lastTable.style.tableLayout = 'auto';
            break;
          case 'window':
            lastTable.style.width = '100%';
            lastTable.style.tableLayout = 'auto';
            break;
        }
      }
    }, 50);
  }

  // ===== WŁAŚCIWOŚCI DOKUMENTU =====

  /**
   * Otwiera dialog właściwości dokumentu
   */
  openPropertiesDialog(): void {
    this.propertiesData.set({ ...this.documentMetadata() });
    this.showPropertiesDialog.set(true);
    this.closeAllMenus();
  }

  /**
   * Zapisuje właściwości dokumentu
   */
  saveProperties(): void {
    const props = this.propertiesData();
    this.documentMetadata.update(m => ({
      ...m,
      title: props.title,
      author: props.author,
      subject: props.subject,
      keywords: props.keywords,
      description: props.description,
      category: props.category,
      company: props.company,
      manager: props.manager,
      contentStatus: props.contentStatus,
      lastModifiedBy: props.lastModifiedBy,
      revision: props.revision,
      version: props.version,
      modified: new Date().toISOString()
    }));
    this.showPropertiesDialog.set(false);
    this.showSuccess('Właściwości dokumentu zostały zaktualizowane');
  }

  /**
   * Zamyka dialog właściwości
   */
  closePropertiesDialog(): void {
    this.showPropertiesDialog.set(false);
  }

  /**
   * Aktualizuje pojedynczą właściwość w propertiesData
   */
  updateProperty(key: string, value: string): void {
    this.propertiesData.update(p => ({ ...p, [key]: value }));
  }

  // ===== PODPISY CYFROWE =====

  /**
   * Otwiera dialog podpisów cyfrowych
   */
  openSignatureDialog(): void {
    this.signatureDialogTab.set(
      this.documentSignatures().length > 0 ? 'list' : 'sign'
    );
    this.signatureData.signerName = '';
    this.signatureData.signerTitle = '';
    this.signatureData.signerEmail = '';
    this.signatureData.reason = '';
    this.signatureData.certificateBase64 = '';
    this.signatureData.certificatePassword = '';
    this.signatureData.certificateFileName = '';
    this.showSignatureDialog.set(true);
    this.closeAllMenus();
  }

  /**
   * Zamyka dialog podpisów
   */
  closeSignatureDialog(): void {
    this.showSignatureDialog.set(false);
  }

  /**
   * Obsługuje wybranie pliku certyfikatu PFX
   */
  onCertificateFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    this.signatureData.certificateFileName = file.name;

    const reader = new FileReader();
    reader.onload = () => {
      const arrayBuffer = reader.result as ArrayBuffer;
      const bytes = new Uint8Array(arrayBuffer);
      let binary = '';
      bytes.forEach(b => binary += String.fromCharCode(b));
      this.signatureData.certificateBase64 = btoa(binary);
    };
    reader.readAsArrayBuffer(file);
  }

  /**
   * Podpisuje dokument
   */
  signDocument(): void {
    if (!this.signatureData.certificateBase64) {
      this.showError('Wybierz plik certyfikatu (.pfx/.p12)');
      return;
    }
    if (!this.signatureData.signerName.trim()) {
      this.showError('Podaj imię i nazwisko podpisującego');
      return;
    }
    if (!this.signatureData.certificatePassword) {
      this.showError('Podaj hasło do certyfikatu');
      return;
    }

    const html = this.editor?.getContent() || this.documentContent();
    const fileName = this.originalFileName() || `${this.documentMetadata().title || 'dokument'}.docx`;

    this.isLoading.set(true);

    const request: SignDocumentRequest = {
      html,
      originalFileName: fileName,
      metadata: this.documentMetadata(),
      header: this.headerContent(),
      footer: this.footerContent(),
      certificateBase64: this.signatureData.certificateBase64,
      certificatePassword: this.signatureData.certificatePassword,
      signerName: this.signatureData.signerName,
      signerTitle: this.signatureData.signerTitle || undefined,
      signerEmail: this.signatureData.signerEmail || undefined,
      signatureReason: this.signatureData.reason || undefined
    };

    this.documentService.signDocument(request).subscribe({
      next: (blob) => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName.endsWith('.docx') ? fileName : `${fileName}.docx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);

        this.showSignatureDialog.set(false);
        this.showSuccess('Dokument został podpisany i pobrany');
        this.isLoading.set(false);
      },
      error: (err) => {
        this.showError(err.message || 'Nie udało się podpisać dokumentu');
        this.isLoading.set(false);
      }
    });
  }

  /**
   * Wstawia wizualny blok podpisu do dokumentu
   */
  insertSignatureLine(): void {
    const name = this.signatureData.signerName || '________________________';
    const title = this.signatureData.signerTitle || '';
    const date = new Date().toLocaleDateString('pl-PL');

    const html = `
      <div style="margin: 24px 0; padding: 16px; border: 1px solid #999; width: 300px; font-family: Calibri, sans-serif;">
        <div style="border-bottom: 1px solid #333; padding-bottom: 40px; margin-bottom: 8px; font-size: 10px; color: #999;">
          ✕ Podpis
        </div>
        <div style="font-size: 12px; font-weight: bold;">${name}</div>
        ${title ? `<div style="font-size: 11px; color: #555;">${title}</div>` : ''}
        <div style="font-size: 10px; color: #888; margin-top: 4px;">Data: ${date}</div>
      </div>
    `;

    this.editor?.insertHtml(html);
    this.showSignatureDialog.set(false);
    this.notifyEditorChange();
  }
}
