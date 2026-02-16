import { 
  Component, 
  EventEmitter, 
  Input, 
  Output,
  signal,
  computed,
  effect,
  HostListener
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { EditorCommand, EditorState, HeadingLevel, DocumentStyle } from '../../models/document.model';

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
  selector: 'app-editor-toolbar',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './editor-toolbar.html',
  styleUrl: './editor-toolbar.scss'
})
export class EditorToolbarComponent {
  private _editorState: EditorState | null = null;

  @HostListener('document:click')
  onDocumentClick(): void {
    this.showStyleDropdown.set(false);
  }
  
  @Input() set editorState(state: EditorState | null) {
    this._editorState = state;
    this.updateFromEditorState(state);
  }
  
  get editorState(): EditorState | null {
    return this._editorState;
  }
  
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
  @Output() styleChange = new EventEmitter<DocumentStyle>();
  @Output() copyFormat = new EventEmitter<void>();
  @Output() pasteFormat = new EventEmitter<void>();

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

  // Dostępne czcionki
  fontFamilies = [
    'Calibri',
    'Calibri Light',
    'Arial',
    'Times New Roman',
    'Georgia',
    'Verdana',
    'Tahoma',
    'Trebuchet MS',
    'Comic Sans MS',
    'Courier New',
    'Impact'
  ];

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

  // Stan dialogów
  showLinkDialog = signal(false);
  showTableDialog = signal(false);
  showStyleDropdown = signal(false);
  linkUrl = '';
  linkText = '';
  tableRows = 3;
  tableCols = 3;

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

    // Aktualizuj rozmiar czcionki
    if (state.currentStyle.fontSize && state.currentStyle.fontSize > 0) {
      this.selectedFontSize.set(state.currentStyle.fontSize);
    }

    // Aktualizuj czcionkę
    if (state.currentStyle.fontFamily) {
      // Normalizuj nazwę czcionki
      const fontFamily = state.currentStyle.fontFamily;
      const matchedFont = this.fontFamilies.find(f => 
        fontFamily.toLowerCase().includes(f.toLowerCase())
      );
      if (matchedFont) {
        this.selectedFontFamily.set(matchedFont);
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

    // Debug log
    console.log('[updateBlockFormatFromState] fontSize:', fontSize, 'bold:', isBold, 'italic:', isItalic, 'blockFormat:', blockFormat);

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

    console.log('[updateBlockFormatFromState] Wybrany format:', format);
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
   * Zmienia rodzinę czcionki (ngModel)
   */
  onFontFamilySelect(fontFamily: string): void {
    this.selectedFontFamily.set(fontFamily);
    this.fontFamilyChange.emit(fontFamily);
  }

  /**
   * Zmienia rodzinę czcionki (event)
   */
  onFontFamilyChange(event: Event): void {
    const select = event.target as HTMLSelectElement;
    this.onFontFamilySelect(select.value);
  }

  /**
   * Zmienia rozmiar czcionki
   */
  onFontSizeChange(event: Event): void {
    const select = event.target as HTMLSelectElement;
    const size = parseInt(select.value, 10);
    this.selectedFontSize.set(size);
    this.fontSizeChange.emit(size);
  }

  /**
   * Zwiększa rozmiar czcionki
   */
  increaseFontSize(): void {
    const currentSize = this.selectedFontSize();
    const newSize = Math.min(currentSize + 1, 400);
    this.selectedFontSize.set(newSize);
    this.fontSizeChange.emit(newSize);
  }

  /**
   * Zmniejsza rozmiar czcionki
   */
  decreaseFontSize(): void {
    const currentSize = this.selectedFontSize();
    const newSize = Math.max(currentSize - 1, 1);
    this.selectedFontSize.set(newSize);
    this.fontSizeChange.emit(newSize);
  }

  /**
   * Obsługa Enter w input rozmiaru czcionki
   */
  onFontSizeInputEnter(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.applyFontSizeFromInput(input);
    input.blur();
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
   * Otwiera dialog tabeli
   */
  openTableDialog(): void {
    this.tableRows = 3;
    this.tableCols = 3;
    this.showTableDialog.set(true);
  }

  /**
   * Zamyka dialog tabeli
   */
  closeTableDialog(): void {
    this.showTableDialog.set(false);
  }

  /**
   * Wstawia tabelę
   */
  confirmInsertTable(): void {
    if (this.tableRows > 0 && this.tableCols > 0) {
      this.insertTable.emit(`${this.tableRows}x${this.tableCols}`);
    }
    this.closeTableDialog();
  }

  /**
   * Wstawia obraz
   */
  onInsertImage(): void {
    this.insertImage.emit();
  }

  /**
   * Sprawdza czy formatowanie jest aktywne
   */
  isActive(format: keyof EditorState['currentFormatting']): boolean {
    return this.editorState?.currentFormatting?.[format] ?? false;
  }

  /**
   * Zapobiega utracie fokusa z edytora przy klikaniu w toolbar
   * (oprócz inputów, które muszą otrzymać fokus)
   */
  onToolbarMouseDown(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    // Pozwól na fokus tylko dla inputów i selectów
    if (target.tagName !== 'INPUT' && target.tagName !== 'SELECT') {
      event.preventDefault();
    }
  }
}
