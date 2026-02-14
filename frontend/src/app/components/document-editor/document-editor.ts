import { 
  Component, 
  ViewChild, 
  inject,
  signal 
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
  showTemplates = signal(false);
  showAbout = signal(false);
  errorMessage = signal<string | null>(null);
  successMessage = signal<string | null>(null);
  
  // Szablony
  templates = signal<DocumentTemplate[]>([]);

  // Zoom
  zoomLevel = signal(100);
  zoomLevels = [50, 75, 100, 125, 150, 200];

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
    // Zastosuj styl do zaznaczenia
    if (this.editor) {
      // Zastosuj właściwości czcionki ze stylu
      if (style.fontFamily) {
        this.editor.setFontFamily(style.fontFamily);
      }
      if (style.fontSize) {
        this.editor.setFontSize(style.fontSize);
      }
      if (style.color) {
        this.editor.setTextColor(style.color);
      }
      if (style.isBold !== undefined) {
        // Jeśli styl wymaga bold a tekst nie jest bold - lub odwrotnie
        const currentState = this.editorState()?.currentFormatting?.bold ?? false;
        if (style.isBold !== currentState) {
          this.editor.executeCommand('bold');
        }
      }
      if (style.isItalic !== undefined) {
        const currentState = this.editorState()?.currentFormatting?.italic ?? false;
        if (style.isItalic !== currentState) {
          this.editor.executeCommand('italic');
        }
      }
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
    this.showMenu.update(v => !v);
    this.showTemplates.set(false);
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
