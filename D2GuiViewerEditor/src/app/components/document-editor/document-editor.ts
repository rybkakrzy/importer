import {
  Component,
  ViewChild,
  ElementRef,
  inject,
  signal,
  computed,
  HostListener,
  OnInit,
  OnDestroy
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ActivatedRoute, Router, RouterLink } from '@angular/router';
import { switchMap, map, filter, distinctUntilChanged } from 'rxjs/operators';
import { from, Observable, Subscription, timer } from 'rxjs';
import { environment } from '../../../environments/environment';
import { WysiwygEditorComponent } from '../wysiwyg-editor/wysiwyg-editor';
import { EditorToolbarComponent } from '../editor-toolbar/editor-toolbar';
import { BarcodeDialogComponent } from '../barcode-dialog/barcode-dialog';
import { RulerComponent } from '../ruler/ruler';
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
import { BuildInfoService } from '../../core/services/build-info.service';
import { DocumentStorageService } from '../../services/document-storage.service';

/**
 * Główny komponent edytora dokumentów Word Online
 */
@Component({
  selector: 'd2-document-editor',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    RouterLink,
    WysiwygEditorComponent,
    EditorToolbarComponent,
    BarcodeDialogComponent,
    RulerComponent
  ],
  templateUrl: './document-editor.html',
  styleUrl: './document-editor.scss'
})
export class DocumentEditorComponent implements OnInit, OnDestroy {
  @ViewChild(WysiwygEditorComponent) editor!: WysiwygEditorComponent;
  @ViewChild(EditorToolbarComponent) toolbar!: EditorToolbarComponent;
  @ViewChild('verticalRulerBar') verticalRulerBar?: ElementRef<HTMLDivElement>;
  @ViewChild('horizontalRulerInner') horizontalRulerInner?: ElementRef<HTMLDivElement>;

  private documentService = inject(DocumentService);
  private documentStorageService = inject(DocumentStorageService);
  private route = inject(ActivatedRoute);
  private router = inject(Router);
  readonly buildInfo = inject(BuildInfoService);

  // Stan dokumentu
  documentContent = signal<string>('<p></p>');
  documentMasterId = signal<string | null>(null);
  /** GUID wersji edytowalnej (v2) — obecny tylko w trybie edycji (?versionId=...). Cel auto-save. */
  documentVersionId = signal<string | null>(null);
  /** Tryb tylko-do-odczytu (Krok 2): brak versionId → ładujemy wersję bazową i blokujemy edycję. */
  readOnly = signal<boolean>(false);

  /**
   * Dokument jest aktualnie edytowany przez kogoś innego (status `Editing` na liście).
   * HOOK: podłączyć pod backendowy `DocumentStatus`, gdy dotrze do edytora — wtedy ustawić
   * `true`, by ukryć narzędzia edycyjne tak samo jak w trybie tylko-do-odczytu.
   */
  lockedByOther = signal<boolean>(false);

  /**
   * Edycja zablokowana: tryb tylko-do-odczytu LUB dokument zajęty przez kogoś innego.
   * Steruje ukrywaniem edycyjnych funkcji w toolbarze i menu.
   */
  editingDisabled = computed(() => this.readOnly() || this.lockedByOther());

  // Auto-save (nadpisuje wersję edytowalną w miejscu)
  autoSaveEnabled = signal<boolean>(environment.autoSave?.enabled ?? true);
  /** Status ostatniego auto-save dla wskaźnika w UI. */
  autoSaveStatus = signal<'idle' | 'saving' | 'saved' | 'error'>('idle');
  lastAutoSaveAt = signal<Date | null>(null);
  private autoSaveSub?: Subscription;
  private isAutoSaving = false;
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
  showFindReplace = signal(false);
  showBarcodeDialog = signal(false);
  errorMessage = signal<string | null>(null);
  successMessage = signal<string | null>(null);
  documentNotFound = signal(false);

  // Menu kontekstowe
  showContextMenu = signal(false);
  contextMenuX = signal(0);
  contextMenuY = signal(0);
  contextSubmenu = signal<string | null>(null);
  contextMenuTargetCell = signal<HTMLElement | null>(null);
  contextMenuTargetImage = signal<HTMLImageElement | null>(null);

  // Mini toolbar nad zaznaczeniem
  showMiniToolbar = signal(false);
  miniToolbarX = signal(0);
  miniToolbarY = signal(0);

  readonly commonFonts = [
    'Calibri', 'Arial', 'Arial Narrow', 'Times New Roman', 'Cambria',
    'Georgia', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Helvetica',
    'Courier New', 'Lucida Console', 'Palatino Linotype', 'Garamond', 'Book Antiqua'
  ];

  /** Czcionka do pokazania w mini-toolbarze — czyta currentStyle i normalizuje
   *  do najbliższej pozycji z commonFonts (tak jak robi to main toolbar). */
  readonly miniToolbarFontFamily = computed(() => {
    const raw = this.editorState()?.currentStyle?.fontFamily;
    if (!raw) return 'Calibri';
    const incoming = raw.trim().toLowerCase();
    // Exact match
    const exact = this.commonFonts.find(f => f.toLowerCase() === incoming);
    if (exact) return exact;
    // Prefix/includes — od najdłuższych, żeby "Calibri Light" > "Calibri"
    const byLength = [...this.commonFonts].sort((a, b) => b.length - a.length);
    const partial = byLength.find(f => incoming.includes(f.toLowerCase()));
    return partial ?? raw;
  });

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
  showRuler = signal(true);

  /**
   * Wcięcie paragrafu/bloku dla aktualnie zaznaczonego fragmentu (cm względem marginesu strony).
   * Wczytywane przy każdej zmianie zaznaczenia z `margin-left`/`margin-right` bieżącego bloku
   * (P/H/UL/OL/LI/TABLE/FIGURE/IMG-wrapper). Używane przez poziomą linijkę — dragowanie
   * uchwytu modyfikuje TYLKO ten blok, jak w MS Word, a nie marginesy całego dokumentu.
   */
  currentBlockIndent = signal<{ start: number; end: number }>({ start: 0, end: 0 });

  /**
   * Stan linii prowadzącej linijki (jak w MS Word). Renderowana nad kartką podczas
   * przeciągania uchwytu — w samej linijce (overflow:hidden, 22px) byłaby przycięta.
   * `offsetPx` to NIEZSKALOWANA odległość od krawędzi strony (lewej/górnej).
   */
  rulerGuide = signal<{ active: boolean; axis: 'horizontal' | 'vertical'; offsetPx: number }>({
    active: false,
    axis: 'horizontal',
    offsetPx: 0
  });

  /** Aktualnie edytowana sekcja (treść / nagłówek / stopka) — z d2-wysiwyg-editor. */
  editingSection = signal<'header' | 'footer' | 'body'>('body');

  /**
   * ZMIERZONA geometria edytowanego pasma nagłówka/stopki (cm od górnej krawędzi strony),
   * z d2-wysiwyg-editor. Pasmo rośnie z treścią (min-height + obraz), więc położenie na
   * pionowej linijce musi pochodzić z pomiaru DOM, nie z marginesów cm.
   */
  sectionGeometry = signal<{ section: 'header' | 'footer'; topCm: number; bottomCm: number } | null>(null);

  /**
   * Marginesy dla PIONOWEJ linijki. Gdy edytowany jest nagłówek/stopka, biały (aktywny)
   * obszar linijki odzwierciedla FAKTYCZNE pasmo nagłówka/stopki (z pomiaru DOM) — jak
   * w MS Word — zamiast globalnego marginesu treści.
   */
  verticalRulerMargins = computed<PageMargins>(() => {
    const m = this.pageSettings().margins;
    const pageH = this.pageSettings().orientation === 'portrait' ? 29.7 : 21;
    const section = this.editingSection();
    const geo = this.sectionGeometry();
    if ((section === 'header' || section === 'footer') && geo && geo.section === section) {
      // białe pasmo linijki = [topCm … bottomCm], reszta = szary margines
      return { ...m, top: Math.max(0, geo.topCm), bottom: Math.max(0, pageH - geo.bottomCm) };
    }
    return m;
  });

  /** Konwersja cm ↔ px (96 DPI). */
  private static readonly CM_TO_PX = 37.795;

  // Menu Widok
  showViewMenu = signal(false);
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

  // Dialog opuszczania edytora
  showLeaveDialog = signal(false);

  // Math dla template
  protected readonly Math = Math;

  constructor() {
    // Załaduj szablony
    this.loadTemplates();
  }

  ngOnInit(): void {
    // Czytamy masterId i versionId RAZEM: versionId decyduje o trybie (edycja vs read-only),
    // więc musi być znany zanim zdecydujemy, którą wersję załadować.
    this.route.queryParams.pipe(
      map(params => ({
        masterId: params['masterId'] as string | undefined,
        versionId: params['versionId'] as string | undefined
      })),
      filter(p => !!p.masterId),
      distinctUntilChanged((a, b) => a.masterId === b.masterId && a.versionId === b.versionId)
    ).subscribe(({ masterId, versionId }) => {
      this.documentVersionId.set(versionId ?? null);
      this.readOnly.set(!versionId);
      this.loadFromStorage(masterId!, versionId ?? null);
    });

    this.startAutoSave();
  }

  ngOnDestroy(): void {
    this.stopAutoSave();
  }

  /**
   * Uruchamia cykliczny auto-save. Interwał z konfiguracji (environment.autoSave.intervalSeconds).
   * Każdy tick nadpisuje wersję edytowalną tylko gdy: auto-save włączony, jest versionId (tryb edycji)
   * oraz w edytorze są niezapisane zmiany.
   */
  private startAutoSave(): void {
    this.stopAutoSave();
    const intervalMs = (environment.autoSave?.intervalSeconds ?? 30) * 1000;
    this.autoSaveSub = timer(intervalMs, intervalMs).subscribe(() => {
      if (!this.autoSaveEnabled()) return;
      if (this.isAutoSaving) return;
      if (!this.documentVersionId() || !this.documentMasterId()) return;
      if (!this.editorState()?.isModified) return;
      this.performAutoSave();
    });
  }

  private stopAutoSave(): void {
    this.autoSaveSub?.unsubscribe();
    this.autoSaveSub = undefined;
  }

  /**
   * Przełącza auto-save (switch w UI). Wyłączenie zatrzymuje cykliczne zapisy.
   */
  toggleAutoSave(): void {
    const next = !this.autoSaveEnabled();
    this.autoSaveEnabled.set(next);
    this.autoSaveStatus.set('idle');
  }

  /**
   * Buduje request zapisu z bieżącego stanu edytora (HTML + metadane + nagłówek/stopka + marginesy).
   */
  private buildSaveRequest() {
    const html = this.editor?.getContent() || this.documentContent();
    const fileName = this.originalFileName() || `${this.documentMetadata().title || 'dokument'}.docx`;
    return {
      html,
      originalFileName: fileName,
      metadata: this.documentMetadata(),
      header: this.headerContent(),
      footer: this.footerContent(),
      margins: this.pageSettings().margins
    };
  }

  /**
   * Serializuje zawartość do DOCX i utrwala przez API.
   * PUT (nadpisanie wersji edytowalnej v2 w miejscu) gdy jest versionId, inaczej POST (nowa wersja).
   */
  private persistDocument(): Observable<unknown> {
    const masterId = this.documentMasterId()!;
    const versionId = this.documentVersionId();
    return this.documentService.saveDocument(this.buildSaveRequest()).pipe(
      switchMap(blob => from(this.blobToBase64(blob))),
      switchMap(base64 => versionId
        ? this.documentStorageService.updateDocumentVersion(masterId, versionId, { content: base64 })
        : this.documentStorageService.saveDocumentVersion(masterId, { content: base64 })
      )
    );
  }

  /**
   * Auto-save: nadpisuje wersję edytowalną w tle (ta sama ścieżka co ręczny „Zapisz").
   */
  private performAutoSave(): void {
    this.isAutoSaving = true;
    this.autoSaveStatus.set('saving');

    this.persistDocument().subscribe({
      next: () => {
        this.editor?.markAsSaved();
        this.lastAutoSaveAt.set(new Date());
        this.autoSaveStatus.set('saved');
        this.isAutoSaving = false;
      },
      error: () => {
        this.autoSaveStatus.set('error');
        this.isAutoSaving = false;
      }
    });
  }

  private static readonly DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  private static readonly DOC_MIME = 'application/msword';
  private static readonly PDF_MIME = 'application/pdf';

  /**
   * Ładuje dokument z bazy.
   * - Tryb edycji (versionId): ładuje WSKAZANĄ wersję edytowalną (Krok 3).
   * - Tryb read-only (brak versionId): ładuje wersję BAZOWĄ — oryginał (Krok 2, bardzo ważne!).
   * PDF nie jest obsługiwany w edytorze DOCX → przekierowanie do /viewer.
   */
  private loadFromStorage(masterId: string, versionId: string | null): void {
    this.isLoading.set(true);
    this.errorMessage.set(null);
    this.documentMasterId.set(masterId);

    this.documentStorageService.getDocumentMetadata(masterId).pipe(
      switchMap(meta => {
        const mime = (meta.mimeType || '').toLowerCase();

        // PDF: edytor DOCX nie renderuje PDF — kieruj do PDFViewer (tryb podglądu).
        if (mime === DocumentEditorComponent.PDF_MIME) {
          this.router.navigate(['/viewer'], { queryParams: { masterId } });
          return from(Promise.reject({ handled: true } as const));
        }

        // Bajty: edycja → wskazana wersja; read-only → wersja bazowa (oryginał).
        const bytes$ = versionId
          ? this.documentStorageService.downloadVersion(masterId, versionId)
          : this.documentStorageService.downloadBaseVersion(masterId);

        const ext = mime === DocumentEditorComponent.DOC_MIME ? '.doc' : '.docx';
        const fileName = `dokument${ext}`;

        return bytes$.pipe(
          switchMap(blob => {
            const file = new File([blob], fileName, { type: mime || DocumentEditorComponent.DOCX_MIME });
            return this.documentService.openDocument(file).pipe(
              map(content => ({ content, fileName }))
            );
          })
        );
      })
    ).subscribe({
      next: ({ content, fileName }) => {
        this.documentContent.set(content.html);
        this.documentMetadata.set(content.metadata);
        this.documentStyles.set(content.styles || []);
        this.originalFileName.set(fileName);
        this.headerContent.set({
          html: content.header?.html || '',
          height: content.header?.height || 1.25
        });
        this.footerContent.set({
          html: content.footer?.html || '',
          height: content.footer?.height || 1.25
        });
        if (content.margins) {
          this.pageSettings.update(s => ({ ...s, margins: content.margins! }));
        }
        if (this.editor) {
          this.editor.setContent(content.html);
        }
        this.documentSignatures.set(content.metadata.signatures || []);
        this.isLoading.set(false);
      },
      error: (err) => {
        if (err?.handled) {
          // przekierowanie do /viewer — nic nie pokazujemy
          return;
        }
        if (err.status === 404) {
          this.documentNotFound.set(true);
        } else {
          this.showError(err.message || 'Nie udało się otworzyć dokumentu z bazy danych');
        }
        this.isLoading.set(false);
      }
    });
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
   * Wraca do dashboardu — pokazuje ładny dialog jeśli jest otwarty dokument lub niezapisane zmiany
   */
  goToDashboard(): void {
    const hasDocument = !!this.documentMasterId();
    const hasChanges = !!this.editorState()?.isModified;

    if (hasDocument || hasChanges) {
      this.showLeaveDialog.set(true);
    } else {
      this.router.navigate(['/']);
    }
  }

  confirmLeave(): void {
    this.showLeaveDialog.set(false);
    this.router.navigate(['/']);
  }

  cancelLeave(): void {
    this.showLeaveDialog.set(false);
  }

  /**
   * Tworzy nowy dokument — zapisuje pusty dokument do bazy, nawiguje do edytora z nowym masterId
   */
  newDocument(): void {
    if (this.editorState()?.isModified) {
      if (!confirm('Masz niezapisane zmiany. Czy na pewno chcesz utworzyć nowy dokument?')) {
        return;
      }
    }

    this.showMenu.set(false);
    this.isLoading.set(true);
    this.errorMessage.set(null);

    this.documentService.newDocument().pipe(
      switchMap(content =>
        this.documentService.saveDocument({ html: content.html, metadata: content.metadata }).pipe(
          switchMap(blob =>
            from(this.blobToBase64(blob)).pipe(
              switchMap(base64 =>
                this.documentStorageService.uploadDocument({
                  name: 'Nowy dokument.docx',
                  mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                  content: base64
                }).pipe(
                  // Nowy dokument = zamiar edycji → twórz wersję edytowalną (v2).
                  switchMap(result =>
                    this.documentStorageService.saveDocumentVersion(result.masterId, { content: base64 }).pipe(
                      map(saved => ({ masterId: result.masterId, versionId: saved.versionId }))
                    )
                  )
                )
              )
            )
          )
        )
      )
    ).subscribe({
      next: ({ masterId, versionId }) => {
        this.router.navigate(['/editor'], { queryParams: { masterId, versionId } });
      },
      error: () => {
        this.showError('Nie udało się utworzyć nowego dokumentu');
        this.isLoading.set(false);
      }
    });
  }

  /**
   * Otwiera dokument z pliku — DOCX: upload do bazy, nawiguje z nowym masterId; PDF: strona konserwacji
   */
  openDocument(): void {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.docx,.pdf';

    input.onchange = async (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (!file) return;

      this.showMenu.set(false);

      if (file.name.toLowerCase().endsWith('.pdf')) {
        this.router.navigate(['/pdf-maintenance']);
        return;
      }

      this.isLoading.set(true);
      this.errorMessage.set(null);

      try {
        const base64 = await this.documentStorageService.fileToBase64(file);
        this.documentStorageService.uploadDocument({
          name: file.name,
          mimeType: file.type || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          content: base64
        }).pipe(
          // Ręczne wczytanie z dysku = zamiar edycji → od razu twórz wersję edytowalną (v2)
          // i otwórz ją w trybie edycji. Oryginał (v1) pozostaje niezmienny; auto-save
          // nadpisuje tylko v2. (Flow aplikacji zewnętrznej bez zmian: master → read-only,
          // master+version → edycja.)
          switchMap(result =>
            this.documentStorageService.saveDocumentVersion(result.masterId, { content: base64 }).pipe(
              map(saved => ({ masterId: result.masterId, versionId: saved.versionId }))
            )
          )
        ).subscribe({
          next: ({ masterId, versionId }) => {
            this.router.navigate(['/editor'], { queryParams: { masterId, versionId } });
          },
          error: () => {
            this.showError('Nie udało się zapisać dokumentu w bazie danych');
            this.isLoading.set(false);
          }
        });
      } catch {
        this.showError('Nie udało się odczytać pliku');
        this.isLoading.set(false);
      }
    };

    input.click();
  }

  private blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(blob);
      reader.onload = () => resolve((reader.result as string).split(',')[1]);
      reader.onerror = reject;
    });
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
        
        // Wczytaj nagłówek i stopkę (resetuj jeśli brak w dokumencie)
        this.headerContent.set({
          html: content.header?.html || '',
          height: content.header?.height || 1.25
        });
        this.footerContent.set({
          html: content.footer?.html || '',
          height: content.footer?.height || 1.25
        });

        // Wczytaj marginesy strony
        if (content.margins) {
          this.pageSettings.update(s => ({ ...s, margins: content.margins! }));
        }

        if (this.editor) {
          this.editor.setContent(content.html);
        }
        
        // Wczytaj podpisy
        this.documentSignatures.set(content.metadata.signatures || []);
        
        this.showSuccess(`Otwarto dokument: ${file.name}`);
        this.isLoading.set(false);

        // masterId już ustawiony przed wywołaniem loadDocument()
      },
      error: (err) => {
        this.showError(err.message || 'Nie udało się otworzyć dokumentu');
        this.isLoading.set(false);
      }
    });
  }

  /**
   * Zapisuje dokument przez API (ujednolicony zapis — ta sama ścieżka co auto-save).
   * Nadpisuje wersję edytowalną (v2) w miejscu gdy jest versionId; w przeciwnym razie tworzy nową wersję.
   */
  saveDocument(): void {
    if (this.readOnly()) {
      // Tryb podglądu (Krok 2) — wersja bazowa jest nietykalna. Zapis przez API zablokowany.
      this.showError('Tryb podglądu — dokument jest tylko do odczytu. Użyj „Pobierz dokument", aby zapisać kopię lokalnie.');
      this.showMenu.set(false);
      return;
    }

    const masterId = this.documentMasterId();
    if (!masterId) {
      // Brak mastera (np. dokument z szablonu, jeszcze nie utrwalony) — pozwól pobrać plik zamiast cichego nic.
      this.showError('Dokument nie jest powiązany z bazą — użyj „Pobierz dokument".');
      this.showMenu.set(false);
      return;
    }

    this.showMenu.set(false);
    this.isLoading.set(true);
    this.autoSaveStatus.set('saving');

    this.persistDocument().subscribe({
      next: () => {
        this.editor?.markAsSaved();
        this.lastAutoSaveAt.set(new Date());
        this.autoSaveStatus.set('saved');
        this.showSuccess('Dokument został zapisany');
        this.isLoading.set(false);
      },
      error: () => {
        this.autoSaveStatus.set('error');
        this.showError('Nie udało się zapisać dokumentu');
        this.isLoading.set(false);
      }
    });
  }

  /**
   * Pobiera dokument jako plik DOCX do przeglądarki (dawne „Zapisz").
   * Nie utrwala w bazie — to lokalna kopia dla użytkownika.
   */
  downloadDocument(): void {
    const request = this.buildSaveRequest();
    const fileName = request.originalFileName;

    // === DIAGNOSTYKA (do debugowania zgubionych stylów / formatowania) ===
    try {
      const html = request.html;
      const tmp = document.createElement('div');
      tmp.innerHTML = html;
      const counts = {
        chars: html.length,
        h1: tmp.querySelectorAll('h1').length,
        h2: tmp.querySelectorAll('h2').length,
        h3: tmp.querySelectorAll('h3').length,
        p: tmp.querySelectorAll('p').length,
        b_strong: tmp.querySelectorAll('b,strong').length,
        i_em: tmp.querySelectorAll('i,em').length,
        u: tmp.querySelectorAll('u').length,
        img: tmp.querySelectorAll('img').length,
        table: tmp.querySelectorAll('table').length,
        tr: tmp.querySelectorAll('tr').length,
        td: tmp.querySelectorAll('td').length,
        pageBreaks: tmp.querySelectorAll('div.page-break').length,
        inlineStyles: tmp.querySelectorAll('[style]').length,
        fontSizeAttrs: Array.from(tmp.querySelectorAll('[style*="font-size"]')).slice(0, 5).map(e => (e as HTMLElement).style.fontSize),
        fontFamilyAttrs: Array.from(tmp.querySelectorAll('[style*="font-family"]')).slice(0, 5).map(e => (e as HTMLElement).style.fontFamily),
      };
      console.group('[downloadDocument] DIAGNOSTYKA HTML wysyłanego do API');
      console.log('fileName:', fileName);
      console.log('counts:', counts);
      console.log('first 2000 chars:', html.substring(0, 2000));
      console.log('header:', this.headerContent());
      console.log('footer:', this.footerContent());
      console.log('margins:', this.pageSettings().margins);
      (window as unknown as { __lastSaveHtml?: string }).__lastSaveHtml = html;
      console.log('Pełny HTML dostępny w window.__lastSaveHtml');
      console.groupEnd();
    } catch (e) {
      console.warn('[downloadDocument] diagnostyka failed', e);
    }

    this.isLoading.set(true);
    this.documentService.downloadDocument(request, fileName);
    this.showSuccess('Pobrano dokument');
    this.isLoading.set(false);
    this.showMenu.set(false);
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

    // Synchronizuj pionową linijkę ze scrollem (pionowym)
    if (this.verticalRulerBar?.nativeElement) {
      this.verticalRulerBar.nativeElement.scrollTop = scrollTop;
    }

    // Synchronizuj poziomą linijkę ze scrollem (poziomym) — przesuwamy ją razem z kartką,
    // żeby podziałka pokrywała się z dokumentem także przy przewijaniu w bok / dużym zoomie.
    if (this.horizontalRulerInner?.nativeElement) {
      this.horizontalRulerInner.nativeElement.style.transform = `translateX(${-container.scrollLeft}px)`;
    }
    
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
                         target.closest('d2-wysiwyg-editor') ||
                         target.closest('.paper-container');
    if (isEditorArea) {
      event.preventDefault();
      this.closeAllMenus();
      this.contextSubmenu.set(null);

      // Wykryj czy kliknięto w komórkę tabeli
      const cellTarget = target.closest('td, th') as HTMLElement | null;
      this.contextMenuTargetCell.set(cellTarget);

      // Wykryj czy kliknięto w obraz
      const imgTarget = (target.tagName === 'IMG' ? target : target.closest('img')) as HTMLImageElement | null;
      this.contextMenuTargetImage.set(imgTarget);

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
    this.showViewMenu.set(false);
    this.activeSubmenu.set(null);
    this.showTemplates.set(false);
    this.showContextMenu.set(false);
    this.contextSubmenu.set(null);
    this.showShadingDropdown.set(false);
  }

  finishDocument(): void {
    // TODO: implement finish logic
  }

  openReportEmail(): void {
    const masterId = this.documentMasterId() ?? '—';
    const version = this.documentMetadata()?.version ?? '—';
    const date = new Date().toLocaleString('pl-PL');
    const url = window.location.href;
    const buildNumber = this.buildInfo.buildNumber();
    const environment = this.buildInfo.environment();

    const subject = encodeURIComponent('[Doc2 Editor] Zgłoszenie');

    // Wyrównane etykiety dla czytelnej kolumny "key: value"
    const rows: Array<[string, string]> = [
      ['Data zgłoszenia',   date],
      ['Master ID',         masterId],
      ['Wersja dokumentu',  version],
      ['Wersja aplikacji',  buildNumber],
      ['Środowisko',        environment],
      ['URL',               url],
    ];
    const labelWidth = Math.max(...rows.map(([k]) => k.length));
    const formatted = rows
      .map(([k, v]) => `  ${k.padEnd(labelWidth)} : ${v}`)
      .join('\n');

    const body = encodeURIComponent(
      `Dzień dobry,\n\n` +
      `proszę o opis problemu poniżej:\n\n` +
      `\n\n\n` +
      `────────────────────────────────────────────\n` +
      `  INFORMACJE DIAGNOSTYCZNE — proszę nie usuwać\n` +
      `────────────────────────────────────────────\n` +
      `${formatted}\n` +
      `────────────────────────────────────────────\n`
    );

    window.open(`mailto:?subject=${subject}&body=${body}`, '_self');
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
   * Globalny skrót wyszukiwania: Ctrl/Cmd+F → Znajdź (oba tryby, w read-only bez zamiany),
   * Ctrl/Cmd+H → Znajdź i zamień (tylko gdy edycja dozwolona).
   */
  @HostListener('document:keydown', ['$event'])
  onGlobalKeydown(e: KeyboardEvent): void {
    if (!(e.ctrlKey || e.metaKey)) return;
    const key = e.key.toLowerCase();
    if (key === 'f') {
      e.preventDefault();
      this.openFindReplace();
    } else if (key === 'h' && !this.editingDisabled()) {
      e.preventDefault();
      this.openFindReplace();
    } else if (key === 'a') {
      // Ctrl+A → zaznacz tylko treść dokumentu (nie całe body z menu/paskami).
      // Pomijamy pola formularzy, by nie psuć natywnego zaznaczania w inputach.
      const tag = (e.target as HTMLElement)?.tagName;
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
      e.preventDefault();
      this.selectAll();
    }
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
    const currentSize = this.editorState()?.currentStyle?.fontSize || 11;
    this.editor?.setFontSize(currentSize + 1);
    this.closeAllMenus();
  }

  /**
   * Zmniejsz rozmiar czcionki
   */
  decreaseFontSize(): void {
    const currentSize = this.editorState()?.currentStyle?.fontSize || 11;
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

  // =====================
  // ZNAJDŹ I ZAMIEŃ
  // =====================
  findText = signal('');
  replaceText = signal('');
  /** Liczba trafień i indeks bieżącego (do wyświetlenia „x/y" w dialogu). */
  findResultCount = signal(0);
  findCurrentIndex = signal(-1);

  /**
   * Wyszukiwanie na żywo podczas wpisywania — podświetla wszystkie trafienia i przewija
   * do pierwszego (przez prawdziwe API edytora, nie ułomne `window.find`).
   */
  onFindInput(value: string): void {
    this.findText.set(value);
    if (!value || !this.editor) {
      this.editor?.clearSearchHighlights();
      this.lastSearchText = '';
      this.findResultCount.set(0);
      this.findCurrentIndex.set(-1);
      return;
    }
    this.lastSearchText = value;
    const result = this.editor.searchText(value, 'next');
    this.findResultCount.set(result.count);
    this.findCurrentIndex.set(result.currentIndex);
  }

  /** Następne trafienie (pierwsze wyszukanie, jeśli tekst się zmienił). */
  findNext(): void {
    const text = this.findText();
    if (!text || !this.editor) return;
    const result = text !== this.lastSearchText
      ? (this.lastSearchText = text, this.editor.searchText(text, 'next'))
      : this.editor.findNext();
    this.findResultCount.set(result.count);
    this.findCurrentIndex.set(result.currentIndex);
  }

  /** Poprzednie trafienie. */
  findPrev(): void {
    const text = this.findText();
    if (!text || !this.editor) return;
    const result = text !== this.lastSearchText
      ? (this.lastSearchText = text, this.editor.searchText(text, 'previous'))
      : this.editor.findPrevious();
    this.findResultCount.set(result.count);
    this.findCurrentIndex.set(result.currentIndex);
  }

  /** Zamknij dialog i wyczyść podświetlenia. */
  closeFindReplace(): void {
    this.showFindReplace.set(false);
    this.editor?.clearSearchHighlights();
    this.lastSearchText = '';
    this.findResultCount.set(0);
    this.findCurrentIndex.set(-1);
  }

  /**
   * Zamienia bieżące trafienie (tylko gdy edycja dozwolona).
   */
  replaceOne(): void {
    if (this.editingDisabled() || !this.editor || !this.findText()) return;
    if (this.findText() !== this.lastSearchText) {
      this.findNext();
      return;
    }
    const result = this.editor.replaceCurrentMatch(this.replaceText());
    this.findResultCount.set(result.count);
    this.findCurrentIndex.set(result.currentIndex);
  }

  /**
   * Zamienia wszystkie trafienia (tylko gdy edycja dozwolona).
   */
  replaceAll(): void {
    if (this.editingDisabled() || !this.editor || !this.findText()) return;
    if (this.findText() !== this.lastSearchText) {
      this.lastSearchText = this.findText();
      this.editor.searchText(this.findText(), 'next');
    }
    const result = this.editor.replaceAllMatches(this.replaceText());
    this.findResultCount.set(result.count);
    this.findCurrentIndex.set(result.currentIndex);
    this.showSuccess(`Zamieniono wszystkie wystąpienia "${this.findText()}"`);
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
  // MINI TOOLBAR
  // =====================

  onEditorMouseUp(event: MouseEvent): void {
    // Nie pokazuj jeśli otwarte jest menu kontekstowe
    if (this.showContextMenu()) return;

    setTimeout(() => {
      const selection = window.getSelection();
      if (!selection || selection.isCollapsed || selection.rangeCount === 0) {
        this.showMiniToolbar.set(false);
        return;
      }
      const range = selection.getRangeAt(0);
      const rect = range.getBoundingClientRect();
      if (rect.width === 0) {
        this.showMiniToolbar.set(false);
        return;
      }

      const toolbarWidth = 560;
      const toolbarHeight = 76;
      const margin = 8;

      let x = rect.left + rect.width / 2 - toolbarWidth / 2;
      let y = rect.top - toolbarHeight - margin;

      // Nie wychodź poza lewą/prawą krawędź ekranu
      x = Math.max(margin, Math.min(x, window.innerWidth - toolbarWidth - margin));
      // Jeśli nie mieści się nad — pokaż pod
      if (y < margin) {
        y = rect.bottom + margin;
      }

      this.miniToolbarX.set(x);
      this.miniToolbarY.set(y);
      this.showMiniToolbar.set(true);
    }, 10);
  }

  onEditorMouseDown(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    if (!target.closest('.mini-toolbar')) {
      this.showMiniToolbar.set(false);
    }
  }

  /** Zapobiega utracie selekcji w edytorze przy klikaniu w mini-toolbar,
   *  ale pozwala INPUT i SELECT na normalne działanie.
   *  Dla INPUT/SELECT selekcja jest zapisywana PRZED przeniesieniem focusu
   *  przez przeglądarkę (mousedown odpala się przed blur edytora). */
  onMiniToolbarMouseDown(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    if (target.tagName === 'INPUT' || target.tagName === 'SELECT') {
      // Zapisz selekcję zanim focus przejdzie do kontrolki i edytor ją wyczyści
      this.editor?.saveSelection();
    } else {
      event.preventDefault();
    }
  }

  miniToolbarCommand(command: string): void {
    this.editor?.executeCommand(command as any);
    // Nie zamykaj — użytkownik może kliknąć kolejny przycisk
    setTimeout(() => {
      const selection = window.getSelection();
      if (!selection || selection.isCollapsed) {
        this.showMiniToolbar.set(false);
      }
    }, 50);
  }

  miniToolbarSetFontFamily(family: string): void {
    this.editor?.setFontFamily(family);
  }

  miniToolbarSetFontSize(event: Event): void {
    const val = parseInt((event.target as HTMLInputElement).value, 10);
    if (!isNaN(val) && val > 0) {
      this.editor?.setFontSize(val);
    }
  }

  miniToolbarIncreaseFontSize(): void {
    const current = this.editorState()?.currentStyle?.fontSize ?? 11;
    this.editor?.setFontSize(current + 1);
  }

  miniToolbarDecreaseFontSize(): void {
    const current = this.editorState()?.currentStyle?.fontSize ?? 11;
    if (current > 1) this.editor?.setFontSize(current - 1);
  }

  miniToolbarSetTextColor(color: string): void {
    this.editor?.setTextColor(color);
  }

  miniToolbarSetHighlightColor(color: string): void {
    this.editor?.setBackgroundColor(color);
  }

  miniToolbarCut(): void {
    document.execCommand('cut');
  }

  miniToolbarCopy(): void {
    document.execCommand('copy');
  }

  miniToolbarPaste(): void {
    navigator.clipboard.readText().then(text => this.editor?.insertText(text)).catch(() => document.execCommand('paste'));
  }

  miniToolbarIncreaseIndent(): void {
    this.editor?.executeCommand('indent');
  }

  miniToolbarDecreaseIndent(): void {
    this.editor?.executeCommand('outdent');
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

  // =====================
  // TABELA – MENU KONTEKSTOWE
  // =====================

  private getContextCell(): HTMLElement | null {
    return this.contextMenuTargetCell() || this.activeTableCell();
  }

  contextMenuInsertRowAbove(): void {
    const cell = this.getContextCell();
    if (!cell) { this.closeContextMenu(); return; }
    const row = cell.closest('tr');
    if (!row) { this.closeContextMenu(); return; }
    const table = row.closest('table')!;
    const colspan = row.querySelectorAll('td, th').length;
    const newRow = row.cloneNode(false) as HTMLTableRowElement;
    for (let i = 0; i < colspan; i++) {
      const td = document.createElement('td');
      td.innerHTML = '<br>';
      newRow.appendChild(td);
    }
    table.querySelector('tbody')?.insertBefore(newRow, row) || row.parentNode?.insertBefore(newRow, row);
    this.notifyEditorChange();
    this.closeContextMenu();
  }

  contextMenuInsertRowBelow(): void {
    const cell = this.getContextCell();
    if (!cell) { this.closeContextMenu(); return; }
    const row = cell.closest('tr');
    if (!row) { this.closeContextMenu(); return; }
    const table = row.closest('table')!;
    const colspan = row.querySelectorAll('td, th').length;
    const newRow = row.cloneNode(false) as HTMLTableRowElement;
    for (let i = 0; i < colspan; i++) {
      const td = document.createElement('td');
      td.innerHTML = '<br>';
      newRow.appendChild(td);
    }
    const nextSibling = row.nextSibling;
    nextSibling ? row.parentNode?.insertBefore(newRow, nextSibling) : row.parentNode?.appendChild(newRow);
    this.notifyEditorChange();
    this.closeContextMenu();
  }

  contextMenuInsertColLeft(): void {
    const cell = this.getContextCell();
    if (!cell) { this.closeContextMenu(); return; }
    const table = cell.closest('table');
    if (!table) { this.closeContextMenu(); return; }
    const colIndex = (cell as HTMLTableCellElement).cellIndex;
    table.querySelectorAll('tr').forEach(row => {
      const ref = row.cells[colIndex];
      const newTd = document.createElement('td');
      newTd.innerHTML = '<br>';
      if (ref) row.insertBefore(newTd, ref);
      else row.appendChild(newTd);
    });
    this.notifyEditorChange();
    this.closeContextMenu();
  }

  contextMenuInsertColRight(): void {
    const cell = this.getContextCell();
    if (!cell) { this.closeContextMenu(); return; }
    const table = cell.closest('table');
    if (!table) { this.closeContextMenu(); return; }
    const colIndex = (cell as HTMLTableCellElement).cellIndex;
    table.querySelectorAll('tr').forEach(row => {
      const ref = row.cells[colIndex];
      const newTd = document.createElement('td');
      newTd.innerHTML = '<br>';
      if (ref?.nextSibling) row.insertBefore(newTd, ref.nextSibling);
      else row.appendChild(newTd);
    });
    this.notifyEditorChange();
    this.closeContextMenu();
  }

  contextMenuDeleteRow(): void {
    const cell = this.getContextCell();
    const row = cell?.closest('tr');
    row?.parentNode?.removeChild(row);
    this.notifyEditorChange();
    this.closeContextMenu();
  }

  contextMenuDeleteCol(): void {
    const cell = this.getContextCell();
    if (!cell) { this.closeContextMenu(); return; }
    const table = cell.closest('table');
    const colIndex = (cell as HTMLTableCellElement).cellIndex;
    table?.querySelectorAll('tr').forEach(row => {
      const td = row.cells[colIndex];
      if (td) row.removeChild(td);
    });
    this.notifyEditorChange();
    this.closeContextMenu();
  }

  contextMenuDeleteTable(): void {
    const cell = this.getContextCell();
    const table = cell?.closest('table');
    table?.parentNode?.removeChild(table);
    this.notifyEditorChange();
    this.closeContextMenu();
  }

  // =====================
  // GRAFIKA – MENU KONTEKSTOWE
  // =====================

  contextMenuAlignImageLeft(): void {
    const img = this.contextMenuTargetImage();
    if (img) {
      img.style.display = 'block';
      img.style.marginLeft = '0';
      img.style.marginRight = 'auto';
      img.style.float = '';
    }
    this.notifyEditorChange();
    this.closeContextMenu();
  }

  contextMenuAlignImageCenter(): void {
    const img = this.contextMenuTargetImage();
    if (img) {
      img.style.display = 'block';
      img.style.marginLeft = 'auto';
      img.style.marginRight = 'auto';
      img.style.float = '';
    }
    this.notifyEditorChange();
    this.closeContextMenu();
  }

  contextMenuAlignImageRight(): void {
    const img = this.contextMenuTargetImage();
    if (img) {
      img.style.display = 'block';
      img.style.marginLeft = 'auto';
      img.style.marginRight = '0';
      img.style.float = '';
    }
    this.notifyEditorChange();
    this.closeContextMenu();
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
  // MENU WIDOK
  // =====================

  toggleViewMenu(): void {
    const wasOpen = this.showViewMenu();
    this.closeAllMenus();
    this.showViewMenu.set(!wasOpen);
  }

  toggleRuler(): void {
    this.showRuler.set(!this.showRuler());
    this.closeAllMenus();
  }

  toggleMarginGuides(): void {
    this.showMarginGuides.set(!this.showMarginGuides());
    this.closeAllMenus();
  }

  /**
   * Obsługuje zmianę marginesów z linijki (drag & drop)
   */
  onRulerMarginsChange(margins: PageMargins): void {
    this.pageSettings.update(s => ({
      ...s,
      margins: { ...margins }
    }));
  }

  /**
   * Aktualizuje stan linii prowadzącej linijki (kreska nad kartką podczas drag).
   */
  onRulerDragGuide(e: { active: boolean; axis: 'horizontal' | 'vertical'; offsetPx: number }): void {
    this.rulerGuide.set({ ...e });
  }

  /**
   * Reaguje na zmianę edytowanej sekcji (treść / nagłówek / stopka) — przełącza
   * obrazowanie pionowej linijki na pasmo nagłówka/stopki.
   */
  onEditingSectionChange(section: 'header' | 'footer' | 'body'): void {
    this.editingSection.set(section);
    if (section === 'body') {
      this.sectionGeometry.set(null);
    }
  }

  /** Przyjmuje zmierzoną geometrię pasma nagłówka/stopki do obrazowania pionowej linijki. */
  onSectionGeometryChange(geo: { section: 'header' | 'footer'; topCm: number; bottomCm: number }): void {
    this.sectionGeometry.set(geo);
  }

  /**
   * Zmiana z PIONOWEJ linijki.
   * - tryb `body`: zwykła zmiana górnego/dolnego marginesu strony.
   * - tryb `header`: dolny uchwyt = dolna krawędź pasma → nowa wysokość nagłówka.
   * - tryb `footer`: górny uchwyt = górna krawędź pasma → nowa wysokość stopki.
   * (jak „header/footer from edge" w MS Word). Wysokość spinamy przez setHeaderHeight/
   * setFooterHeight, co emituje headerChange/footerChange i odświeża pasmo + linijkę.
   */
  onVerticalRulerMarginsChange(margins: PageMargins): void {
    const section = this.editingSection();
    if (section === 'body') {
      this.onRulerMarginsChange(margins);
      return;
    }
    const pageH = this.pageSettings().orientation === 'portrait' ? 29.7 : 21;
    const geo = this.sectionGeometry();
    if (section === 'header') {
      const topCm = geo ? Math.max(0, geo.topCm) : 0;
      const newBottomCm = pageH - margins.bottom; // dolna krawędź pasma nagłówka
      const newHeight = Math.round((newBottomCm - topCm) * 100) / 100;
      this.editor?.setHeaderHeight(newHeight);
    } else if (section === 'footer') {
      const bottomCm = geo ? geo.bottomCm : pageH - margins.bottom; // dolna krawędź pasma stopki
      const newTopCm = margins.top; // górna krawędź pasma stopki
      const newHeight = Math.round((bottomCm - newTopCm) * 100) / 100;
      this.editor?.setFooterHeight(newHeight);
    }
  }

  /**
   * Obsługuje zmianę wcięcia paragrafu z poziomej linijki (drag & drop).
   * Zachowuje się jak w MS Word: zmiana dotyczy TYLKO zaznaczonych bloków
   * (paragraf, lista, tabela, obraz/figura), a NIE marginesów ca\u0142ego dokumentu.
   */
  onRulerBlockIndentChange(indent: { start?: number; end?: number }): void {
    const blocks = this.getSelectedBlocks();
    if (blocks.length === 0) return;

    for (const block of blocks) {
      if (indent.start !== undefined) {
        const cm = Math.max(-this.pageSettings().margins.left + 0.1, indent.start);
        if (cm === 0) {
          block.style.removeProperty('margin-left');
        } else {
          block.style.marginLeft = `${cm.toFixed(2)}cm`;
        }
      }
      if (indent.end !== undefined) {
        const cm = Math.max(-this.pageSettings().margins.right + 0.1, indent.end);
        if (cm === 0) {
          block.style.removeProperty('margin-right');
        } else {
          block.style.marginRight = `${cm.toFixed(2)}cm`;
        }
      }
    }

    // Zaktualizuj sygna\u0142 wci\u0119cia (uchwyty linijki natychmiast podskakuj\u0105 do nowej pozycji)
    this.currentBlockIndent.update(prev => ({
      start: indent.start !== undefined ? indent.start : prev.start,
      end: indent.end !== undefined ? indent.end : prev.end
    }));

    // Powiadom edytor o modyfikacji (auto-save / dirty flag)
    this.editor?.triggerContentChange();
  }

  /**
   * Nas\u0142uchuje zmian zaznaczenia, \u017ceby zaktualizowa\u0107 odczyt wci\u0119cia paragrafu
   * dla poziomej linijki.
   */
  @HostListener('document:selectionchange')
  onDocumentSelectionChange(): void {
    this.updateCurrentBlockIndent();
  }

  /**
   * Znajduje wszystkie unikalne bloki nadrz\u0119dne zawarte w aktualnym zaznaczeniu.
   * Blokiem jest: P, H1\u2013H6, UL, OL, LI, TABLE, FIGURE, BLOCKQUOTE, DIV (poza wrapperami).
   * Je\u015bli zaznaczona jest grafika \u2014 zwracamy paragraf w kt\u00f3rym jest osadzona (lub IMG).
   */
  private getSelectedBlocks(): HTMLElement[] {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return [];

    const range = selection.getRangeAt(0);
    const blocks: HTMLElement[] = [];
    const seen = new Set<HTMLElement>();
    const blockTags = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'TABLE', 'FIGURE', 'BLOCKQUOTE', 'LI', 'IMG']);

    const findBlock = (node: Node | null): HTMLElement | null => {
      let n = node;
      while (n && n !== document) {
        if (n.nodeType === Node.ELEMENT_NODE) {
          const el = n as HTMLElement;
          if (blockTags.has(el.tagName)) return el;
        }
        n = n.parentNode;
      }
      return null;
    };

    if (range.collapsed) {
      const b = findBlock(range.startContainer);
      if (b) blocks.push(b);
    } else {
      // Iteruj po w\u0119z\u0142ach mi\u0119dzy startem a ko\u0144cem
      const walker = document.createTreeWalker(
        range.commonAncestorContainer,
        NodeFilter.SHOW_ELEMENT,
        {
          acceptNode: (n: Node) => {
            const el = n as HTMLElement;
            if (!blockTags.has(el.tagName)) return NodeFilter.FILTER_SKIP;
            return range.intersectsNode(el) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
          }
        }
      );
      // Dodaj rodzic\u00f3w start/end na wypadek gdyby walker pomin\u0105\u0142
      const startBlock = findBlock(range.startContainer);
      if (startBlock) blocks.push(startBlock);
      let node = walker.nextNode();
      while (node) {
        const el = node as HTMLElement;
        // Pomi\u0144 LI je\u015bli ma rodzica UL/OL w blocks (bo wci\u0119cie aplikujemy do listy)
        blocks.push(el);
        node = walker.nextNode();
      }
      const endBlock = findBlock(range.endContainer);
      if (endBlock) blocks.push(endBlock);
    }

    // Deduplikacja + filtracja: je\u015bli mamy UL/OL i jego LI \u2014 zostaw UL/OL.
    // Je\u015bli mamy IMG i jego paragraf \u2014 zostaw paragraf (CSS margin na P dzia\u0142a lepiej).
    const result: HTMLElement[] = [];
    for (const b of blocks) {
      if (seen.has(b)) continue;
      seen.add(b);
      result.push(b);
    }
    // Usu\u0144 LI je\u015bli rodzic UL/OL te\u017c jest w secie
    const filtered = result.filter(el => {
      if (el.tagName === 'LI') {
        const parent = el.parentElement;
        if (parent && (parent.tagName === 'UL' || parent.tagName === 'OL') && seen.has(parent)) {
          return false;
        }
      }
      if (el.tagName === 'IMG') {
        // Zamie\u0144 na rodzica paragrafu
        let p = el.parentElement;
        while (p && !['P', 'DIV', 'FIGURE'].includes(p.tagName)) p = p.parentElement;
        if (p) {
          if (!seen.has(p)) {
            seen.add(p);
            result.push(p);
          }
          return false;
        }
      }
      return true;
    });

    return filtered;
  }

  /**
   * Odczytuje wci\u0119cie (margin-left/right) z pierwszego bloku w zaznaczeniu
   * i zapisuje do `currentBlockIndent`. Warto\u015bci w cm (px / 37.795).
   */
  private updateCurrentBlockIndent(): void {
    const blocks = this.getSelectedBlocks();
    if (blocks.length === 0) {
      this.currentBlockIndent.set({ start: 0, end: 0 });
      return;
    }
    const block = blocks[0];
    const style = window.getComputedStyle(block);
    const mlPx = parseFloat(style.marginLeft) || 0;
    const mrPx = parseFloat(style.marginRight) || 0;
    const start = Math.round((mlPx / DocumentEditorComponent.CM_TO_PX) * 100) / 100;
    const end = Math.round((mrPx / DocumentEditorComponent.CM_TO_PX) * 100) / 100;
    const prev = this.currentBlockIndent();
    if (prev.start !== start || prev.end !== end) {
      this.currentBlockIndent.set({ start, end });
    }
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

  // ===== Walidacja rozmiaru tabeli =====
  // type="number" + min/max nie blokuje wpisania ręcznie —0", „-5" czy 9999
  // — atrybuty te wpływają tylko na spinner i :invalid. Dlatego trzymamy
  // jawne sprawdzenie + clamp on blur + disabled na przycisku „Wstaw".
  private static readonly TABLE_MIN_COLS = 1;
  private static readonly TABLE_MAX_COLS = 63;   // limit Worda
  private static readonly TABLE_MIN_ROWS = 1;
  private static readonly TABLE_MAX_ROWS = 500;

  isTableColumnsValid(): boolean {
    const v = this.tableDialogData.columns;
    return Number.isFinite(v) && Number.isInteger(v)
      && v >= DocumentEditorComponent.TABLE_MIN_COLS && v <= DocumentEditorComponent.TABLE_MAX_COLS;
  }

  isTableRowsValid(): boolean {
    const v = this.tableDialogData.rows;
    return Number.isFinite(v) && Number.isInteger(v)
      && v >= DocumentEditorComponent.TABLE_MIN_ROWS && v <= DocumentEditorComponent.TABLE_MAX_ROWS;
  }

  insertTableValidationError(): string | null {
    if (!this.isTableColumnsValid()) {
      return `Liczba kolumn musi być liczbą całkowitą z zakresu ${DocumentEditorComponent.TABLE_MIN_COLS}–${DocumentEditorComponent.TABLE_MAX_COLS}.`;
    }
    if (!this.isTableRowsValid()) {
      return `Liczba wierszy musi być liczbą całkowitą z zakresu ${DocumentEditorComponent.TABLE_MIN_ROWS}–${DocumentEditorComponent.TABLE_MAX_ROWS}.`;
    }
    return null;
  }

  clampTableColumns(): void {
    const v = this.tableDialogData.columns;
    if (!Number.isFinite(v)) {
      this.tableDialogData.columns = DocumentEditorComponent.TABLE_MIN_COLS;
      return;
    }
    this.tableDialogData.columns = Math.max(
      DocumentEditorComponent.TABLE_MIN_COLS,
      Math.min(DocumentEditorComponent.TABLE_MAX_COLS, Math.floor(v))
    );
  }

  clampTableRows(): void {
    const v = this.tableDialogData.rows;
    if (!Number.isFinite(v)) {
      this.tableDialogData.rows = DocumentEditorComponent.TABLE_MIN_ROWS;
      return;
    }
    this.tableDialogData.rows = Math.max(
      DocumentEditorComponent.TABLE_MIN_ROWS,
      Math.min(DocumentEditorComponent.TABLE_MAX_ROWS, Math.floor(v))
    );
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
    // Twarda bramka — nawet jeśli ktoś ominie disabled (np. enter), nic nie wstawimy.
    if (this.insertTableValidationError()) {
      return;
    }
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
