import { 
  Component, 
  EventEmitter, 
  Input, 
  Output,
  signal,
  computed,
  effect
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { EditorCommand, EditorState, HeadingLevel, DocumentStyle } from '../../models/document.model';

/** Domyślne style Word */
const DEFAULT_WORD_STYLES: DocumentStyle[] = [
  {
    id: 'Normal',
    name: 'Normalny',
    type: 'paragraph',
    fontFamily: 'Calibri',
    fontSize: 11,
    color: '#000000',
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
    outlineLevel: 4
  },
  {
    id: 'Heading5',
    name: 'Nagłówek 5',
    type: 'paragraph',
    fontFamily: 'Calibri Light',
    fontSize: 11,
    color: '#2F5496',
    outlineLevel: 5
  },
  {
    id: 'Heading6',
    name: 'Nagłówek 6',
    type: 'paragraph',
    fontFamily: 'Calibri Light',
    fontSize: 11,
    color: '#1F3763',
    isItalic: true,
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
  styleUrl: './editor-toolbar.css'
})
export class EditorToolbarComponent {
  @Input() editorState: EditorState | null = null;
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

  // Stan dialogów
  showLinkDialog = signal(false);
  showTableDialog = signal(false);
  linkUrl = '';
  linkText = '';
  tableRows = 3;
  tableCols = 3;

  selectedBlockFormat = signal('paragraph');

  /**
   * Konwertuje ID stylu na komendę edytora
   */
  private styleIdToCommand(styleId: string): string {
    const id = styleId.toLowerCase();
    if (id === 'normal') return 'paragraph';
    if (id.startsWith('heading')) {
      const level = id.replace('heading', '');
      return `heading${level}`;
    }
    return 'paragraph';
  }

  /**
   * Wykonuje komendę edytora
   */
  executeCommand(cmd: EditorCommand, value?: string): void {
    this.command.emit({ command: cmd, value });
  }

  /**
   * Zmienia format bloku
   */
  onBlockFormatChange(event: Event): void {
    const select = event.target as HTMLSelectElement;
    const format = select.value as EditorCommand;
    this.selectedBlockFormat.set(format);
    
    // Znajdź styl i wyemituj go
    const selectedFormat = this.blockFormats().find(f => f.value === format);
    if (selectedFormat) {
      this.styleChange.emit(selectedFormat.style);
    }
    
    this.executeCommand(format);
  }

  /**
   * Zmienia rodzinę czcionki
   */
  onFontFamilyChange(event: Event): void {
    const select = event.target as HTMLSelectElement;
    this.selectedFontFamily.set(select.value);
    this.fontFamilyChange.emit(select.value);
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
}
