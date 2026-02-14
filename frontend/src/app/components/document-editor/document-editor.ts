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
  DocumentStyle
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
    EditorToolbarComponent
  ],
  templateUrl: './document-editor.html',
  styleUrl: './document-editor.css'
})
export class DocumentEditorComponent {
  @ViewChild(WysiwygEditorComponent) editor!: WysiwygEditorComponent;

  private documentService = inject(DocumentService);

  // Stan dokumentu
  documentContent = signal<string>('<p>&nbsp;</p>');
  documentMetadata = signal<DocumentMetadata>({
    title: 'Nowy dokument',
    created: new Date().toISOString(),
    modified: new Date().toISOString()
  });
  documentStyles = signal<DocumentStyle[]>([]);
  originalFileName = signal<string>('');
  
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
  errorMessage = signal<string | null>(null);
  successMessage = signal<string | null>(null);
  
  // Szablony
  templates = signal<DocumentTemplate[]>([]);

  // Zoom
  zoomLevel = signal(100);
  zoomLevels = [50, 75, 100, 125, 150, 200];

  // Strony
  currentPage = signal(1);
  totalPages = signal(1);

  // Ustawienia strony
  showPageSetup = signal(false);
  showMarginGuides = signal(true);
  pageSettings = signal<PageSettings>({
    margins: { top: 2.5, bottom: 2.5, left: 2.5, right: 2.5 },
    orientation: 'portrait',
    paperSize: 'a4'
  });
  marginPresets = MARGIN_PRESETS;

  // Math dla template
  protected readonly Math = Math;

  constructor() {
    // Załaduj szablony
    this.loadTemplates();
    
    // Sprawdź autosave
    this.checkAutoSave();
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
   * Sprawdza czy jest zapisany autosave
   */
  private checkAutoSave(): void {
    const autoSaveTime = localStorage.getItem('editor_autosave_time');
    const autoSaveContent = localStorage.getItem('editor_autosave');
    
    if (autoSaveTime && autoSaveContent) {
      const savedTime = new Date(autoSaveTime);
      const now = new Date();
      const hoursDiff = (now.getTime() - savedTime.getTime()) / (1000 * 60 * 60);
      
      // Jeśli zapisano w ciągu ostatnich 24 godzin
      if (hoursDiff < 24) {
        const restore = confirm(
          `Znaleziono niezapisaną pracę z ${savedTime.toLocaleString()}.\n\nCzy chcesz ją przywrócić?`
        );
        
        if (restore) {
          this.documentContent.set(autoSaveContent);
          this.showSuccess('Przywrócono niezapisaną pracę');
        } else {
          localStorage.removeItem('editor_autosave');
          localStorage.removeItem('editor_autosave_time');
        }
      }
    }
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
    
    this.documentContent.set('<p>&nbsp;</p>');
    this.documentMetadata.set({
      title: 'Nowy dokument',
      created: new Date().toISOString(),
      modified: new Date().toISOString()
    });
    this.documentStyles.set([]); // Reset stylów - toolbar użyje domyślnych
    this.originalFileName.set('');
    this.editor?.setContent('<p>&nbsp;</p>');
    this.editor?.clearAutoSave();
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
        
        if (this.editor) {
          this.editor.setContent(content.html);
          this.editor.clearAutoSave();
        }
        
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
        metadata: this.documentMetadata()
      },
      fileName
    );
    
    this.editor?.markAsSaved();
    this.editor?.clearAutoSave();
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
          this.editor.clearAutoSave();
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
      if (base64) {
        this.editor?.insertImage(base64, file.name);
      }
    };
    reader.readAsDataURL(file);
  }

  /**
   * Wstawia tabelę
   */
  onInsertTable(config: string): void {
    this.editor?.insertTable(config);
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
   * Obsługa zmiany stanu edytora
   */
  onStateChange(state: EditorState): void {
    this.editorState.set(state);
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
    // Jeśli kliknięto poza menu - zamknij
    if (!isMenuArea) {
      this.closeAllMenus();
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
    this.activeSubmenu.set(null);
    this.showTemplates.set(false);
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
}
