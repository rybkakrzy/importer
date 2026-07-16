import {
  Component,
  EventEmitter,
  Input,
  Output,
  signal,
  computed,
  effect,
  inject,
  HostListener
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { EditorCommand, EditorState, HeadingLevel, DocumentStyle } from '../../models/document.model';
import { FontProviderService } from '../../services/font-provider.service';

/** Domyślne style Word */
const DEFAULT_WORD_STYLES: DocumentStyle[] = [
  {
    id: 'Title',
    name: 'Tytuł',
    type: 'paragraph',
    fontFamily: 'Calibri Light',
    fontSize: 28,
    color: '#000000',
    isBold: false,
    isItalic: false,
    isUnderline: false
  },
  {
    id: 'Subtitle',
    name: 'Podtytuł',
    type: 'paragraph',
    fontFamily: 'Calibri',
    fontSize: 14,
    color: '#5A5A5A',
    isBold: false,
    isItalic: true,
    isUnderline: false
  },
  {
    id: 'Normal',
    name: 'Normalny',
    type: 'paragraph',
    fontFamily: 'Calibri',
    fontSize: 11,
    color: '#000000',
    isBold: false,
    isItalic: false,
    isUnderline: false,
    alignment: 'left',
    spaceAfter: 8,
    lineSpacing: 1.08
  },
  {
    id: 'Heading1',
    name: 'Nagłówek 1',
    type: 'paragraph',
    fontFamily: 'Calibri Light',
    fontSize: 16,
    color: '#2F5496',
    isBold: true,
    isItalic: false,
    isUnderline: false,
    spaceBefore: 12,
    outlineLevel: 1
  },
  {
    id: 'Heading2',
    name: 'Nagłówek 2',
    type: 'paragraph',
    fontFamily: 'Calibri Light',
    fontSize: 13,
    color: '#2F5496',
    isBold: true,
    isItalic: false,
    isUnderline: false,
    spaceBefore: 2,
    outlineLevel: 2
  },
  {
    id: 'Heading3',
    name: 'Nagłówek 3',
    type: 'paragraph',
    fontFamily: 'Calibri Light',
    fontSize: 12,
    color: '#1F3763',
    isBold: true,
    isItalic: false,
    isUnderline: false,
    spaceBefore: 2,
    outlineLevel: 3
  },
  {
    id: 'Heading4',
    name: 'Nagłówek 4',
    type: 'paragraph',
    fontFamily: 'Calibri Light',
    fontSize: 11,
    color: '#2F5496',
    isBold: true,
    isItalic: true,
    isUnderline: false,
    outlineLevel: 4
  },
  {
    id: 'Heading5',
    name: 'Nagłówek 5',
    type: 'paragraph',
    fontFamily: 'Calibri Light',
    fontSize: 11,
    color: '#2F5496',
    isBold: false,
    isItalic: false,
    isUnderline: false,
    outlineLevel: 5
  },
  {
    id: 'Heading6',
    name: 'Nagłówek 6',
    type: 'paragraph',
    fontFamily: 'Calibri Light',
    fontSize: 11,
    color: '#1F3763',
    isBold: false,
    isItalic: true,
    isUnderline: false,
    outlineLevel: 6
  }
];

/**
 * Komponent paska narzędzi edytora
 */
@Component({
  selector: 'd2-editor-toolbar',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './editor-toolbar.html',
  styleUrl: './editor-toolbar.scss'
})
export class EditorToolbarComponent {
  private readonly fontProvider = inject(FontProviderService);
  private _editorState: EditorState | null = null;

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent): void {
    // Nie zamykaj dropdowna gdy klik jest wewnątrz toolbara
    if ((event.target as HTMLElement).closest('d2-editor-toolbar')) {
      return;
    }
    this.showStyleDropdown.set(false);
  }
  
  @Input() set editorState(state: EditorState | null) {
    this._editorState = state;
    this.updateFromEditorState(state);
  }
  
  get editorState(): EditorState | null {
    return this._editorState;
  }
  
  /** Gdy true (read-only / dokument zajęty), ukrywamy edycyjne kontrolki — zostaje wyszukiwarka. */
  @Input() readOnly = false;

  @Input() set documentStyles(styles: DocumentStyle[] | null) {
    if (styles && styles.length > 0) {
      this._documentStyles.set(styles);
    } else {
      this._documentStyles.set(DEFAULT_WORD_STYLES);
    }
  }
  
  @Output() command = new EventEmitter<{ command: EditorCommand; value?: string }>();
  @Output() fontSizeChange = new EventEmitter<number>();
  @Output() fontFamilyChange = new EventEmitter<string>();
  @Output() textColorChange = new EventEmitter<string>();
  @Output() backgroundColorChange = new EventEmitter<string>();
  @Output() insertLink = new EventEmitter<{ url: string; text?: string }>();
  @Output() insertImage = new EventEmitter<void>();
  @Output() insertTable = new EventEmitter<string>();
  @Output() openTableDialog = new EventEmitter<void>();
  @Output() insertFootnote = new EventEmitter<void>();
  @Output() insertEndnote = new EventEmitter<void>();
  @Output() insertBarcode = new EventEmitter<void>();
  @Output() styleChange = new EventEmitter<DocumentStyle>();
  @Output() copyFormat = new EventEmitter<void>();
  @Output() pasteFormat = new EventEmitter<void>();
  @Output() searchInDocument = new EventEmitter<{ text: string; direction: 'next' | 'previous' }>();
  @Output() replaceInDocument = new EventEmitter<{ searchText: string; replaceText: string; all: boolean }>();
  /**
   * Emitowane gdy użytkownik klika element toolbara, który przejmie fokus
   * (input/select). Rodzic powinien wtedy zachować selekcję edytora, żeby
   * po blur/Enter można było ją przywrócić.
   */
  @Output() preserveSelection = new EventEmitter<void>();
  @Output() clearSearch = new EventEmitter<void>();
  /** Klik w lupę — otwórz panel „Wyszukiwanie" (po lewej), zamiast paska pod toolbarem. */
  @Output() openSearch = new EventEmitter<void>();
  /** Klik „Akapit" — otwórz okno ustawień akapitu (ta sama akcja co menu „Narzędzia"). */
  @Output() openParagraph = new EventEmitter<void>();

  // Style dokumentu
  private _documentStyles = signal<DocumentStyle[]>(DEFAULT_WORD_STYLES);
  
  // Style do wyświetlenia w dropdown
  blockFormats = computed(() => {
    return this._documentStyles().map(style => ({
      value: this.styleIdToCommand(style.id),
      label: style.name,
      style: style
    }));
  });

  /** Shared font list (main + contextual toolbars, incl. corporate font — item 7). */
  readonly fontFamilies = this.fontProvider.displayNames;

  /** True when the current selection spans more than one font family (item 6). */
  readonly fontMixed = signal(false);

  /** Value shown in the font combobox — blank on a mixed selection. */
  readonly fontInputValue = computed(() =>
    this.fontMixed() ? '' : this.selectedFontFamily(),
  );

  /** True while the user is actively editing the font input (guards read-back). */
  private fontEditing = false;

  // Dostępne rozmiary czcionki
  fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

  // Aktualnie wybrane wartości
  selectedFontFamily = signal('Calibri');
  selectedFontSize = signal(11);
  selectedTextColor = signal('#000000');
  selectedBgColor = signal('#ffffff');

  // Stan format painter
  formatPainterActive = signal(false);
  private copiedFormat: Partial<EditorState['currentFormatting']> | null = null;

  /**
   * Znacznik czasu ostatniej manualnej zmiany rozmiaru czcionki (klik +/- lub input).
   * Przez krótki czas (300 ms) ignorujemy aktualizacje fontSize z editorState — inaczej
   * read-back z edytora (computed style w pustym ZWS-spanie po wstawieniu) nadpisuje
   * naszą świeżą wartość starym rozmiarem i input wraca do poprzedniej wartości.
   */
  private lastManualFontSizeChange = 0;

  // Stan dialogów
  showLinkDialog = signal(false);
  showStyleDropdown = signal(false);
  showSearchBar = signal(false);
  showReplaceRow = signal(false);
  searchText = '';
  replaceText = '';
  searchResultCount = signal(0);
  currentSearchIndex = signal(0);
  linkUrl = '';
  linkText = '';

  selectedBlockFormat = signal('paragraph');

  /**
   * Pobiera label wybranego stylu
   */
  getSelectedStyleLabel(): string {
    const format = this.blockFormats().find(f => f.value === this.selectedBlockFormat());
    return format?.label || 'Normalny';
  }

  /**
   * Przełącza dropdown stylów
   */
  toggleStyleDropdown(): void {
    this.showStyleDropdown.update(v => !v);
  }

  /**
   * Zamyka dropdown stylów
   */
  closeStyleDropdown(): void {
    this.showStyleDropdown.set(false);
  }

  /**
   * Wybiera styl z dropdown
   */
  selectStyle(format: { value: string; label: string; style: DocumentStyle }): void {
    this.selectedBlockFormat.set(format.value);
    this.styleChange.emit(format.style);
    this.showStyleDropdown.set(false);
  }

  /**
   * Oblicza rozmiar podglądu stylu (skalowany dla dropdown)
   */
  getStylePreviewSize(originalSize: number | undefined): number {
    if (!originalSize) return 11;
    // Skaluj rozmiary aby zmieściły się w dropdown
    // Tytuł (28pt) -> 18pt, Normalny (11pt) -> 11pt
    if (originalSize >= 24) return 18;
    if (originalSize >= 16) return 14;
    if (originalSize >= 13) return 12;
    return 11;
  }

  /**
   * Aktualizuje toolbar na podstawie stanu edytora
   */
  private updateFromEditorState(state: EditorState | null): void {
    if (!state?.currentStyle) return;

    // Aktualizuj rozmiar czcionki — pomijamy jeśli user właśnie kliknął +/-/wpisał wartość
    // (read-back z edytora bywa stary, bo karetka leży w pustym ZWS-spanie).
    if (state.currentStyle.fontSize && state.currentStyle.fontSize > 0) {
      const sinceManual = Date.now() - this.lastManualFontSizeChange;
      if (sinceManual > 300) {
        this.selectedFontSize.set(state.currentStyle.fontSize);
      }
    }

    // Update the font name from the caret/selection. Skip while the user is
    // typing in the combobox, otherwise a selectionChange read-back would stomp
    // the draft. Normalisation is delegated to the shared provider (item 7).
    if (!this.fontEditing) {
      this.fontMixed.set(!!state.fontMixed);
      const rawFont = state.currentStyle.fontFamily ?? state.fontFamily;
      if (!state.fontMixed && rawFont) {
        this.selectedFontFamily.set(this.fontProvider.normalize(rawFont));
      }
    }

    // Aktualizuj kolor tekstu
    if (state.currentStyle.textColor) {
      this.selectedTextColor.set(state.currentStyle.textColor);
    }

    // Aktualizuj format bloku - dopasuj na podstawie tagu LUB właściwości stylu
    this.updateBlockFormatFromState(state);
  }

  /**
   * Dopasowuje format bloku na podstawie stanu edytora
   */
  private updateBlockFormatFromState(state: EditorState): void {
    const blockFormat = state.currentStyle?.blockFormat;
    const fontSize = state.currentStyle?.fontSize || 11;
    const isBold = state.currentFormatting?.bold || false;
    const isItalic = state.currentFormatting?.italic || false;

    let format = 'paragraph';

    // Najpierw sprawdź tag HTML dla nagłówków
    if (blockFormat === 'h1') {
      format = 'heading1';
    } else if (blockFormat === 'h2') {
      format = 'heading2';
    } else if (blockFormat === 'h3') {
      format = 'heading3';
    } else if (blockFormat === 'h4') {
      format = 'heading4';
    } else if (blockFormat === 'h5') {
      format = 'heading5';
    } else if (blockFormat === 'h6') {
      format = 'heading6';
    } else {
      // Dopasuj styl na podstawie porównania z definicjami stylów
      // Używamy tolerancji ±2pt dla fontSize
      const tolerance = 2;
      
      // Tytuł: fontSize ~28pt (26-30)
      if (fontSize >= 26) {
        format = 'title';
      }
      // Nagłówek 1: fontSize ~16pt, bold, kolor niebieski
      else if (fontSize >= 15 && fontSize <= 18 && isBold) {
        format = 'heading1';
      }
      // Podtytuł: fontSize ~14pt, italic, nie bold
      else if (fontSize >= 13 && fontSize <= 15 && isItalic && !isBold) {
        format = 'subtitle';
      }
      // Nagłówek 2: fontSize ~13pt, bold
      else if (fontSize >= 12 && fontSize <= 14 && isBold && !isItalic) {
        format = 'heading2';
      }
      // Nagłówek 3: fontSize ~12pt, bold
      else if (fontSize >= 11 && fontSize <= 13 && isBold && !isItalic) {
        format = 'heading3';
      }
      // Nagłówek 4: fontSize ~11pt, bold i italic
      else if (fontSize >= 10 && fontSize <= 12 && isBold && isItalic) {
        format = 'heading4';
      }
      // Dla tekstu większego niż normalny (>14pt) ale bez innych cech - traktuj jako Tytuł
      else if (fontSize >= 18) {
        format = 'title';
      }
      // Normalny: fontSize ~11pt lub inne
      else {
        format = 'paragraph';
      }
    }

    this.selectedBlockFormat.set(format);
  }

  /**
   * Konwertuje ID stylu na komendę edytora
   */
  private styleIdToCommand(styleId: string): string {
    const id = styleId.toLowerCase();
    if (id === 'normal') return 'paragraph';
    if (id === 'title') return 'title';
    if (id === 'subtitle') return 'subtitle';
    if (id.startsWith('heading')) {
      const level = id.replace('heading', '');
      return `heading${level}`;
    }
    return styleId.toLowerCase();
  }

  /**
   * Wykonuje komendę edytora
   */
  executeCommand(cmd: EditorCommand, value?: string): void {
    this.command.emit({ command: cmd, value });
  }

  /**
   * Zmienia format bloku (ngModel)
   */
  onBlockFormatSelect(format: string): void {
    this.selectedBlockFormat.set(format);
    
    // Znajdź styl i wyemituj go - applyDocumentStyle zajmie się wszystkim
    const selectedFormat = this.blockFormats().find(f => f.value === format);
    if (selectedFormat) {
      this.styleChange.emit(selectedFormat.style);
    }
  }

  /**
   * Zmienia format bloku (event)
   */
  onBlockFormatChange(event: Event): void {
    const select = event.target as HTMLSelectElement;
    this.onBlockFormatSelect(select.value);
  }

  /**
   * Font combobox (item 6). Backed by an `<input list=…>` + `<datalist>` so the
   * user can read the effective font, type a name, and search — without the
   * value ever blanking on focus (the old native `<select>` blanked because it
   * reset `selectedIndex` to -1 on mousedown). The document font is NEVER
   * overwritten until the user confirms (change/Enter/blur with a real value).
   */
  onFontFocus(event: FocusEvent): void {
    this.fontEditing = true;
    // Clear the field so the native datalist shows the FULL list. If we left the
    // current font name in place, the browser would filter the options down to
    // that single entry and no other font could be picked without backspacing.
    // The current font is restored on blur (see onFontBlur) if nothing is chosen.
    (event.target as HTMLInputElement).value = '';
    // Let the parent snapshot the editor selection before focus moves here.
    this.preserveSelection.emit();
  }

  /** Restore the visible font when the field is left empty (no pick made). */
  onFontBlur(event: FocusEvent): void {
    const input = event.target as HTMLInputElement;
    this.fontEditing = false;
    if (!input.value.trim()) {
      input.value = this.fontInputValue();
    }
  }

  onFontKeydown(event: KeyboardEvent): void {
    const input = event.target as HTMLInputElement;
    if (event.key === 'Enter') {
      event.preventDefault();
      this.commitFont(input.value, input);
      input.blur();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      // Cancel — restore the current effective font, do not overwrite.
      input.value = this.fontInputValue();
      input.blur();
    }
  }

  /** Fires on datalist pick or on blur after a change. */
  onFontCommit(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.fontEditing = false;
    this.commitFont(input.value, input);
  }

  private commitFont(raw: string, input: HTMLInputElement): void {
    this.fontEditing = false;
    const value = raw.trim();
    if (!value) {
      // Empty input → keep the current font (no silent overwrite).
      input.value = this.fontInputValue();
      return;
    }
    const canonical = this.fontProvider.normalize(value);
    input.value = canonical;
    if (this.fontMixed() || canonical !== this.selectedFontFamily()) {
      this.selectedFontFamily.set(canonical);
      this.fontMixed.set(false);
      this.fontFamilyChange.emit(canonical);
    }
  }

  /**
   * Zmienia rozmiar czcionki
   */
  onFontSizeChange(event: Event): void {
    const select = event.target as HTMLSelectElement;
    const size = parseInt(select.value, 10);
    this.selectedFontSize.set(size);
    this.lastManualFontSizeChange = Date.now();
    this.fontSizeChange.emit(size);
  }

  /**
   * Zwiększa rozmiar czcionki — skacze do następnej wartości ze standardowej listy
   * (jak w MS Word: 11→12→14→16→18…). Powyżej 72 dorzucamy +2pt liniowo.
   */
  increaseFontSize(): void {
    const currentSize = this.selectedFontSize();
    const next = this.fontSizes.find(s => s > currentSize);
    const newSize = next ?? Math.min(currentSize + 2, 400);
    this.selectedFontSize.set(newSize);
    this.lastManualFontSizeChange = Date.now();
    this.fontSizeChange.emit(newSize);
  }

  /**
   * Zmniejsza rozmiar czcionki — skacze do poprzedniej wartości ze standardowej listy.
   */
  decreaseFontSize(): void {
    const currentSize = this.selectedFontSize();
    const prev = [...this.fontSizes].reverse().find(s => s < currentSize);
    const newSize = prev ?? Math.max(currentSize - 2, 1);
    this.selectedFontSize.set(newSize);
    this.lastManualFontSizeChange = Date.now();
    this.fontSizeChange.emit(newSize);
  }

  /**
   * Obsługa Enter w input rozmiaru czcionki
   */
  onFontSizeInputEnter(event: Event): void {
    // ENTER: zablokuj domyślną akcję i ODŁÓŻ aplikację na po zakończeniu zdarzenia.
    // Synchroniczny blur() w trakcie obsługi ENTER powoduje, że `setFontSize` przywraca fokus
    // i zaznaczenie do edytora JESZCZE w trakcie tego zdarzenia — domyślna akcja Enter (nowa
    // linia) trafia wtedy w przywrócone zaznaczenie i KASUJE zaznaczony tekst. Klik poza pole
    // (blur myszką) nie ma tego problemu, bo nie ma zdarzenia Enter. preventDefault + setTimeout
    // rozdzielają aplikację od zdarzenia Enter. `onFontSizeInputBlur` aplikuje raz.
    event.preventDefault();
    const input = event.target as HTMLInputElement;
    setTimeout(() => input.blur(), 0);
  }

  /**
   * Obsługa blur w input rozmiaru czcionki
   */
  onFontSizeInputBlur(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.applyFontSizeFromInput(input);
  }

  /**
   * Aplikuje rozmiar czcionki z inputa
   */
  private applyFontSizeFromInput(input: HTMLInputElement): void {
    const value = parseInt(input.value, 10);
    if (!isNaN(value) && value >= 1 && value <= 400) {
      this.selectedFontSize.set(value);
      this.lastManualFontSizeChange = Date.now();
      this.fontSizeChange.emit(value);
    } else {
      // Przywróć poprzednią wartość
      input.value = this.selectedFontSize().toString();
    }
  }

  /**
   * Przełącza tryb kopiowania formatowania
   */
  toggleFormatPainter(): void {
    if (this.formatPainterActive()) {
      // Wyłącz format painter
      this.formatPainterActive.set(false);
    } else {
      // Kopiuj bieżące formatowanie
      this.copyFormat.emit();
      this.formatPainterActive.set(true);
    }
  }

  /**
   * Aplikuje skopiowane formatowanie
   */
  applyFormatPainter(): void {
    if (this.formatPainterActive()) {
      this.pasteFormat.emit();
      this.formatPainterActive.set(false);
    }
  }

  /**
   * Wyłącza format painter
   */
  deactivateFormatPainter(): void {
    this.formatPainterActive.set(false);
  }

  /**
   * Zmienia kolor tekstu
   */
  onTextColorChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.selectedTextColor.set(input.value);
    this.textColorChange.emit(input.value);
  }

  /**
   * Zmienia kolor tła
   */
  onBgColorChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.selectedBgColor.set(input.value);
    this.backgroundColorChange.emit(input.value);
  }

  /**
   * Otwiera dialog linku
   */
  openLinkDialog(): void {
    this.linkUrl = '';
    this.linkText = '';
    this.showLinkDialog.set(true);
  }

  /**
   * Zamyka dialog linku
   */
  closeLinkDialog(): void {
    this.showLinkDialog.set(false);
  }

  /**
   * Wstawia link
   */
  confirmInsertLink(): void {
    if (this.linkUrl) {
      this.insertLink.emit({ 
        url: this.linkUrl, 
        text: this.linkText || undefined 
      });
    }
    this.closeLinkDialog();
  }

  /**
   * Otwiera dialog tabeli (deleguje do document-editor)
   */
  onOpenTableDialog(): void {
    this.openTableDialog.emit();
  }

  /**
   * Wstawia obraz
   */
  onInsertImage(): void {
    this.insertImage.emit();
  }

  onInsertFootnote(): void {
    this.insertFootnote.emit();
  }

  onInsertEndnote(): void {
    this.insertEndnote.emit();
  }

  /**
   * Otwiera dialog wstawiania kodu kreskowego / QR
   */
  onInsertBarcode(): void {
    this.insertBarcode.emit();
  }

  /**
   * Sprawdza czy formatowanie jest aktywne
   */
  isActive(format: keyof EditorState['currentFormatting']): boolean {
    return this.editorState?.currentFormatting?.[format] === true;
  }

  /**
   * Stan przycisków wyrównania — jak w Wordzie dokładnie jeden jest aktywny,
   * domyślnie „do lewej" (brak jawnego text-align = lewa).
   */
  isAlignActive(align: 'left' | 'center' | 'right' | 'justify'): boolean {
    return (this.editorState?.currentFormatting?.alignment ?? 'left') === align;
  }

  /**
   * Przełącza pasek wyszukiwania
   */
  toggleSearchBar(): void {
    const newValue = !this.showSearchBar();
    this.showSearchBar.set(newValue);
    if (!newValue) {
      this.searchText = '';
      this.replaceText = '';
      this.searchResultCount.set(0);
      this.currentSearchIndex.set(0);
      this.showReplaceRow.set(false);
      this.clearSearch.emit();
    }
  }

  /**
   * Zamyka pasek wyszukiwania
   */
  closeSearchBar(): void {
    this.showSearchBar.set(false);
    this.searchText = '';
    this.replaceText = '';
    this.searchResultCount.set(0);
    this.currentSearchIndex.set(0);
    this.showReplaceRow.set(false);
    this.clearSearch.emit();
  }

  /**
   * Przełącza wiersz zamiany
   */
  toggleReplaceRow(): void {
    this.showReplaceRow.update(v => !v);
  }

  /**
   * Reaguje na zmianę tekstu w polu wyszukiwania
   */
  onSearchInput(): void {
    if (this.searchText.length > 0) {
      this.searchInDocument.emit({ text: this.searchText, direction: 'next' });
    } else {
      this.searchResultCount.set(0);
      this.currentSearchIndex.set(0);
      this.clearSearch.emit();
    }
  }

  /**
   * Znajduje następne wystąpienie
   */
  findNext(): void {
    if (this.searchText) {
      this.searchInDocument.emit({ text: this.searchText, direction: 'next' });
    }
  }

  /**
   * Znajduje poprzednie wystąpienie
   */
  findPrevious(): void {
    if (this.searchText) {
      this.searchInDocument.emit({ text: this.searchText, direction: 'previous' });
    }
  }

  /**
   * Zamienia następne wystąpienie
   */
  replaceNext(): void {
    if (this.searchText) {
      this.replaceInDocument.emit({ searchText: this.searchText, replaceText: this.replaceText, all: false });
    }
  }

  /**
   * Zamienia wszystkie wystąpienia
   */
  replaceAll(): void {
    if (this.searchText) {
      this.replaceInDocument.emit({ searchText: this.searchText, replaceText: this.replaceText, all: true });
    }
  }

  /**
   * Aktualizuje wyniki wyszukiwania (wywoływane z zewnątrz)
   */
  updateSearchResults(count: number, currentIndex: number): void {
    this.searchResultCount.set(count);
    this.currentSearchIndex.set(currentIndex);
  }

  /**
   * Zapobiega utracie fokusa z edytora przy klikaniu w toolbar
   * (oprócz inputów, które muszą otrzymać fokus).
   * Dla input/select emitujemy `preserveSelection` — rodzic zapisuje selekcję
   * edytora ZANIM fokus przeskoży na pole tekstowe, dzięki czemu po Enter/blur
   * można ją przywrócić.
   */
  onToolbarMouseDown(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    // Pozwól na fokus tylko dla inputów i selectów
    if (target.tagName !== 'INPUT' && target.tagName !== 'SELECT') {
      event.preventDefault();
    } else {
      this.preserveSelection.emit();
    }
  }
}
