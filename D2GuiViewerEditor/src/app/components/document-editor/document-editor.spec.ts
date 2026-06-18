import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { of, throwError } from 'rxjs';
import { ActivatedRoute, Router } from '@angular/router';
import { DocumentEditorComponent } from './document-editor';
import { DocumentService, OpenDocumentError } from '../../services/document.service';
import { DocumentStorageService } from '../../services/document-storage.service';
import { BuildInfoService } from '../../core/services/build-info.service';

/**
 * Testy reguły: ESC zamyka aktywny boczny panel (Wyszukiwanie / właściwości tabeli)
 * tą samą logiką co przycisk × — bez regresji dla edytora i bez skutków ubocznych,
 * gdy panel jest zamknięty lub gdy ESC obsłużył już bardziej szczegółowy handler.
 *
 * Komponent ma ciężkie zależności (route/auto-save/timer) inicjalizowane w ngOnInit,
 * dlatego NIE wołamy detectChanges() — testujemy logikę HostListenera bezpośrednio
 * na instancji (HostListener i tak deleguje do tej samej metody co realne zdarzenie).
 */
describe('DocumentEditorComponent — ESC zamyka boczny panel', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: { getTemplates: () => of([]) } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: {} },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
  });

  function esc(): KeyboardEvent {
    const ev = new KeyboardEvent('keydown', { key: 'Escape', cancelable: true });
    component.onEscapeKeydown(ev);
    return ev;
  }

  it('zamyka panel Wyszukiwania po ESC', () => {
    component.showFindReplace.set(true);
    const ev = esc();
    expect(component.showFindReplace()).toBe(false);
    expect(ev.defaultPrevented).toBe(true);
  });

  it('zamyka panel właściwości / stylizacji tabeli po ESC', () => {
    component.showTablePanel.set(true);
    const ev = esc();
    expect(component.showTablePanel()).toBe(false);
    expect(ev.defaultPrevented).toBe(true);
  });

  it('gdy żaden panel nie jest otwarty — ESC nie robi nic i nie blokuje zdarzenia', () => {
    component.showFindReplace.set(false);
    component.showTablePanel.set(false);
    const ev = esc();
    expect(component.showFindReplace()).toBe(false);
    expect(component.showTablePanel()).toBe(false);
    expect(ev.defaultPrevented).toBe(false);
  });

  it('nie zamyka panelu, gdy ESC obsłużył już bardziej szczegółowy handler (defaultPrevented)', () => {
    component.showFindReplace.set(true);
    const ev = new KeyboardEvent('keydown', { key: 'Escape', cancelable: true });
    ev.preventDefault(); // np. deselekcja obrazu w edytorze
    component.onEscapeKeydown(ev);
    expect(component.showFindReplace()).toBe(true);
  });

  it('nie zamyka bocznego panelu, gdy otwarty jest dialog (dialog ma pierwszeństwo)', () => {
    component.showFindReplace.set(true);
    component.showInsertTableDialog.set(true);
    esc();
    expect(component.showFindReplace()).toBe(true);
  });

  it('ESC używa tej samej logiki zamykania co × — czyści stan wyszukiwania', () => {
    component.showFindReplace.set(true);
    component.findResultCount.set(5);
    component.findCurrentIndex.set(2);
    esc();
    expect(component.showFindReplace()).toBe(false);
    expect(component.findResultCount()).toBe(0);
    expect(component.findCurrentIndex()).toBe(-1);
  });

  it('zamknięcie panelu tabeli przez ESC oznacza go jako ręcznie zamknięty (jak ×)', () => {
    component.showTablePanel.set(true);
    esc();
    expect(component.showTablePanel()).toBe(false);
    // Kolejny ESC bez otwartego panelu nie rzuca błędu
    expect(() => esc()).not.toThrow();
  });
});

describe('DocumentEditorComponent — przełączanie doku tabela ↔ wyszukiwanie', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: { getTemplates: () => of([]) } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: {} },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
  });

  // syncTablePanel is the single decision point for the docked panel mode.
  const sync = () => (component as any).syncTablePanel();

  it('ponowne zaznaczenie tabeli przy otwartym wyszukiwaniu przełącza dok na formatowanie tabeli', () => {
    component.showFindReplace.set(true);
    component.isInTable.set(true);
    sync();
    expect(component.showTablePanel()).toBe(true);
    expect(component.showFindReplace()).toBe(false);
  });

  it('w tabeli (bez wyszukiwania) pokazuje panel tabeli', () => {
    component.showFindReplace.set(false);
    component.isInTable.set(true);
    sync();
    expect(component.showTablePanel()).toBe(true);
  });

  it('opuszczenie tabeli chowa panel tabeli i resetuje flagę ręcznego zamknięcia', () => {
    component.isInTable.set(true);
    sync();
    component.isInTable.set(false);
    sync();
    expect(component.showTablePanel()).toBe(false);
  });

  it('ręczne zamknięcie panelu tabeli (×) nie otwiera go ponownie póki karetka jest w tabeli', () => {
    component.isInTable.set(true);
    component.closeTablePanel();
    sync();
    expect(component.showTablePanel()).toBe(false);
  });

  it('dok ma jeden aktywny tryb — tabela i wyszukiwanie nie są widoczne jednocześnie', () => {
    component.showFindReplace.set(true);
    component.isInTable.set(true);
    sync();
    expect(component.showTablePanel() && component.showFindReplace()).toBe(false);
  });
});

describe('DocumentEditorComponent — widoczność przycisku „Zakończ" (canFinish)', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: { getTemplates: () => of([]) } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: {} },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
  });

  it('jest widoczny, gdy link zwrotny istnieje i jest poprawny', () => {
    component.returnUrl.set('https://app.example.com/return');
    expect(component.canFinish()).toBe(true);
  });

  it('nie jest renderowany, gdy link jest null', () => {
    component.returnUrl.set(null);
    expect(component.canFinish()).toBe(false);
  });

  it('nie jest renderowany, gdy link jest pustym stringiem', () => {
    component.returnUrl.set('');
    expect(component.canFinish()).toBe(false);
  });

  it('nie jest renderowany, gdy link to same białe znaki', () => {
    component.returnUrl.set('   ');
    expect(component.canFinish()).toBe(false);
  });

  it('nie jest renderowany, gdy link jest niepoprawnym URL', () => {
    component.returnUrl.set('not-a-url');
    expect(component.canFinish()).toBe(false);
  });
});

describe('DocumentEditorComponent — ESC zamyka edycję nagłówka/stopki', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: { getTemplates: () => of([]) } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: {} },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
  });

  function esc(): KeyboardEvent {
    const ev = new KeyboardEvent('keydown', { key: 'Escape', cancelable: true });
    component.onEscapeKeydown(ev);
    return ev;
  }

  it('zamyka tryb edycji nagłówka po ESC (preventDefault + delegacja do editor)', () => {
    component.editingSection.set('header');
    let stopped = false;
    (component as any).editor = { stopEditingHeaderFooter: () => { stopped = true; } };

    const ev = esc();

    expect(stopped).toBe(true);
    expect(ev.defaultPrevented).toBe(true);
  });

  it('zamyka tryb edycji stopki po ESC', () => {
    component.editingSection.set('footer');
    let stopped = false;
    (component as any).editor = { stopEditingHeaderFooter: () => { stopped = true; } };

    const ev = esc();

    expect(stopped).toBe(true);
    expect(ev.defaultPrevented).toBe(true);
  });

  it('tryb nagłówka/stopki ma pierwszeństwo nad panelem wyszukiwania', () => {
    component.editingSection.set('header');
    component.showFindReplace.set(true);
    let stopped = false;
    (component as any).editor = { stopEditingHeaderFooter: () => { stopped = true; } };

    esc();

    expect(stopped).toBe(true);
    // Panel wyszukiwania nie zostaje przedwcześnie zamknięty — nadrzędna akcja to wyjście z header/footer.
    expect(component.showFindReplace()).toBe(true);
  });

  it('w trybie body ESC nie ruga editor.stopEditingHeaderFooter', () => {
    component.editingSection.set('body');
    let stopped = false;
    (component as any).editor = { stopEditingHeaderFooter: () => { stopped = true; } };

    const ev = esc();

    expect(stopped).toBe(false);
    expect(ev.defaultPrevented).toBe(false);
  });
});

describe('DocumentEditorComponent — koordynacja panelu Nagłówek/Stopka z Find/Tables', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: { getTemplates: () => of([]) } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: {} },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
  });

  it('panel nagłówka/stopki jest otwarty, gdy editingSection != body i brak innych paneli', () => {
    component.editingSection.set('header');
    component.showFindReplace.set(false);
    component.showTablePanel.set(false);
    expect(component.showHeaderFooterPanel()).toBe(true);
  });

  it('panel nagłówka/stopki ustępuje miejsca panelowi „Znajdź", gdy oba są aktywne', () => {
    component.editingSection.set('header');
    component.showFindReplace.set(true);
    expect(component.showHeaderFooterPanel()).toBe(false);
  });

  it('panel nagłówka/stopki ustępuje miejsca panelowi tabeli', () => {
    component.editingSection.set('footer');
    component.showTablePanel.set(true);
    expect(component.showHeaderFooterPanel()).toBe(false);
  });

  it('po zamknięciu „Znajdź" panel nagłówka wraca, jeżeli edycja trwa', () => {
    component.editingSection.set('header');
    component.showFindReplace.set(true);
    expect(component.showHeaderFooterPanel()).toBe(false);

    component.showFindReplace.set(false);

    expect(component.showHeaderFooterPanel()).toBe(true);
  });

  it('zamknięcie edycji (editingSection = body) zamyka panel', () => {
    component.editingSection.set('header');
    expect(component.showHeaderFooterPanel()).toBe(true);

    component.editingSection.set('body');

    expect(component.showHeaderFooterPanel()).toBe(false);
  });

  it('panel obrazu ma pierwszeństwo nad panelem nagłówka/stopki', () => {
    component.editingSection.set('header');
    component.selectedImage.set({
      widthPx: 100, heightPx: 50, aspectRatio: 2, alignment: null, positionMode: 'inline',
      border: { enabled: false, color: '#000000', widthPx: 1, style: 'solid' },
      crop: { left: 0, right: 0, top: 0, bottom: 0 },
    });
    expect(component.showImagePanel()).toBe(true);
    expect(component.showHeaderFooterPanel()).toBe(false);
  });
});

describe('DocumentEditorComponent — menu „Pomoc" i akcja „Zgłoś"', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: { getTemplates: () => of([]) } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        // openReportEmail reads buildNumber() / environment() from BuildInfoService.
        { provide: BuildInfoService, useValue: {
            buildNumber: () => '1.0.0',
            environment: () => 'TEST',
        } },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
  });

  it('toggleHelpMenu otwiera dropdown „Pomoc" i zamyka inne menu', () => {
    component.showViewMenu.set(true);

    component.toggleHelpMenu();

    expect(component.showHelpMenu()).toBe(true);
    expect(component.showViewMenu()).toBe(false);
  });

  it('drugi toggleHelpMenu zamyka dropdown', () => {
    component.toggleHelpMenu();
    expect(component.showHelpMenu()).toBe(true);

    component.toggleHelpMenu();

    expect(component.showHelpMenu()).toBe(false);
  });

  it('closeAllMenus zamyka również „Pomoc"', () => {
    component.showHelpMenu.set(true);

    component.closeAllMenus();

    expect(component.showHelpMenu()).toBe(false);
  });

  it('openReportEmail() zamyka otwarte menu „Pomoc" (akcja zamyka dropdown jak inne menu)', () => {
    component.showHelpMenu.set(true);
    // window.open is invoked by openReportEmail — stub it so the test doesn't open a tab.
    const origOpen = window.open;
    (window as any).open = () => null;

    try {
      component.openReportEmail();
      expect(component.showHelpMenu()).toBe(false);
    } finally {
      (window as any).open = origOpen;
    }
  });

  /** Przechwytuje URL mailto przekazany do window.open i zwraca odkodowane body. */
  function captureMailBody(): string {
    let captured = '';
    const origOpen = window.open;
    (window as any).open = (url: string) => { captured = url; return null; };
    try {
      component.openReportEmail();
    } finally {
      (window as any).open = origOpen;
    }
    const body = /[?&]body=([^&]*)/.exec(captured)?.[1] ?? '';
    return decodeURIComponent(body);
  }

  it('openReportEmail() wstawia poprawny Master ID i Version ID w treści maila (Problem 5)', () => {
    component.documentMasterId.set('master-123');
    component.documentVersionId.set('version-456');

    const body = captureMailBody();

    expect(body).toContain('Master ID');
    expect(body).toContain('master-123');
    expect(body).toContain('Version ID');
    expect(body).toContain('version-456');
    // VersionId NIE może być fallbackiem, skoro istnieje.
    expect(body).not.toMatch(/Version ID\s*:\s*—/);
  });

  it('openReportEmail() używa fallbacku „—" dla Version ID tylko gdy wersja faktycznie nie istnieje', () => {
    component.documentMasterId.set('master-123');
    component.documentVersionId.set(null);

    const body = captureMailBody();

    expect(body).toMatch(/Version ID\s*:\s*—/);
  });
});

describe('DocumentEditorComponent — dialog hasła: anulowanie (Problem 1)', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;
  let navigateSpy: ReturnType<typeof vi.fn>;

  beforeEach(async () => {
    navigateSpy = vi.fn();
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: { getTemplates: () => of([]) } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: navigateSpy } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: {} },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
  });

  it('Anuluj bez returnUrl → przekierowanie na dashboard (nie zostawia pustego dokumentu)', () => {
    component.returnUrl.set(null);
    component.showPasswordDialog.set(true);

    component.cancelPasswordDialog();

    expect(component.showPasswordDialog()).toBe(false);
    expect(navigateSpy).toHaveBeenCalledWith(['/']);
  });

  it('Anuluj z returnUrl → przepływ powrotu (zamknięcie karty), bez nawigacji na dashboard', () => {
    component.returnUrl.set('https://app.example.com/return');
    component.showPasswordDialog.set(true);
    const origClose = window.close;
    const closeSpy = vi.fn();
    (window as any).close = closeSpy;

    try {
      component.cancelPasswordDialog();
    } finally {
      (window as any).close = origClose;
    }

    expect(component.showPasswordDialog()).toBe(false);
    expect(closeSpy).toHaveBeenCalled();
    expect(navigateSpy).not.toHaveBeenCalled();
  });
});

describe('DocumentEditorComponent — userDownload (widoczność „Pobierz dokument")', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: { getTemplates: () => of([]) } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: { buildNumber: () => '1', environment: () => 'TEST' } },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
  });

  it('default ⇒ canUserDownload=false (brak metadanych)', () => {
    expect(component.canUserDownload()).toBe(false);
  });

  it('userDownload=true ⇒ canUserDownload=true', () => {
    component.userDownload.set(true);
    expect(component.canUserDownload()).toBe(true);
  });

  it('userDownload=false ⇒ canUserDownload=false', () => {
    component.userDownload.set(false);
    expect(component.canUserDownload()).toBe(false);
  });

  it('downloadDocument() bez canUserDownload pokazuje błąd i nie wywołuje API', () => {
    let downloadCalls = 0;
    (component as any).documentStorageService = {
      downloadEditedDocument: () => { downloadCalls++; return { subscribe: () => {} }; }
    };
    component.documentMasterId.set('master-1');
    component.userDownload.set(false);

    component.downloadDocument();

    expect(downloadCalls).toBe(0);
    expect(component.errorMessage()).not.toBeNull();
  });
});

/**
 * Etap 4 (front): rozmiar/orientacja strony round-tripują — wczytany `pageSize`
 * trafia z powrotem do requestu zapisu, więc landscape/niestandardowy rozmiar nie
 * jest gubiony (backend zapisuje go jako w:pgSz; brak → fallback A4 po stronie API).
 */
describe('DocumentEditorComponent — rozmiar strony (PageSize round-trip)', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: { getTemplates: () => of([]) } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: {} },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
  });

  it('buildSaveRequest niesie wczytany pageSize', () => {
    component.documentPageSize.set({ widthCm: 29.7, heightCm: 21, orientation: 'landscape' });

    const req = (component as any).buildSaveRequest();

    expect(req.pageSize).toEqual({ widthCm: 29.7, heightCm: 21, orientation: 'landscape' });
  });

  it('bez wczytanego pageSize request ma undefined (backend → fallback A4)', () => {
    const req = (component as any).buildSaveRequest();

    expect(req.pageSize).toBeUndefined();
  });
});

/**
 * Qutas-PAR-007: „Ustaw jako domyślne" w ustawieniach akapitu.
 * Regresja: przycisk RESETOWAŁ formularz do wartości bazowych zamiast zapisać
 * bieżące ustawienia jako domyślne. Teraz: zachowuje bieżące wartości jako default
 * sesji i seeduje nimi dialog, gdy nie ma aktywnej selekcji.
 *
 * jsdom nie ma layoutu/contenteditable — `applyParagraphSettings()` przy braku selekcji
 * robi no-op i zamyka dialog; testujemy stan modelu (paragraphData / default sesji).
 */
describe('DocumentEditorComponent — „Ustaw jako domyślne" akapitu (Qutas-PAR-007)', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: { getTemplates: () => of([]) } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: {} },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
  });

  it('NIE resetuje formularza do wartości bazowych (sedno regresji)', () => {
    component.paragraphData.alignment = 'center';
    component.paragraphData.spaceAfter = 24;
    component.paragraphData.lineSpacingValue = 2;

    component.setParagraphAsDefault();

    // Wartości użytkownika zachowane, nie podmienione na bazowe (left / 8 / 1.08).
    expect(component.paragraphData.alignment).toBe('center');
    expect(component.paragraphData.spaceAfter).toBe(24);
    expect(component.paragraphData.lineSpacingValue).toBe(2);
  });

  it('zapisany default seeduje dialog, gdy nie ma aktywnej selekcji', () => {
    component.paragraphData.alignment = 'right';
    component.paragraphData.spaceBefore = 12;
    component.setParagraphAsDefault();

    // Symuluj inny stan formularza, a następnie ponowne otwarcie dialogu bez selekcji.
    component.paragraphData.alignment = 'left';
    component.paragraphData.spaceBefore = 0;
    (component as any).readCurrentParagraphSettings();

    expect(component.paragraphData.alignment).toBe('right');
    expect(component.paragraphData.spaceBefore).toBe(12);
  });

  it('default jest niezależną kopią — późniejsza edycja formularza go nie zmienia', () => {
    component.paragraphData.indentLeft = 3;
    component.setParagraphAsDefault();

    component.paragraphData.indentLeft = 99;
    (component as any).readCurrentParagraphSettings(); // brak selekcji → seed z defaultu

    expect(component.paragraphData.indentLeft).toBe(3);
  });
});

/**
 * Qutas-UI-006: menu kontekstowe nie może zasłaniać UI (toolbar) ani wychodzić poza viewport.
 * Regresja: clamp dolnej krawędzi bez `Math.max(8, …)` dawał ujemne `y` na niskim oknie →
 * menu wjeżdżało nad viewport, zasłaniając toolbar.
 */
describe('DocumentEditorComponent — pozycja menu kontekstowego (Qutas-UI-006)', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;
  let origW: number;
  let origH: number;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: { getTemplates: () => of([]) } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: {} },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
    origW = window.innerWidth;
    origH = window.innerHeight;
  });

  afterEach(() => {
    Object.defineProperty(window, 'innerWidth', { value: origW, configurable: true });
    Object.defineProperty(window, 'innerHeight', { value: origH, configurable: true });
  });

  function ctxEventAt(clientX: number, clientY: number): MouseEvent {
    const target = document.createElement('div');
    target.className = 'paper-container';
    document.body.appendChild(target);
    return { target, clientX, clientY, preventDefault: () => {} } as unknown as MouseEvent;
  }

  it('na niskim oknie clamp nie daje ujemnego y (menu nie zasłania toolbara)', () => {
    Object.defineProperty(window, 'innerWidth', { value: 1200, configurable: true });
    Object.defineProperty(window, 'innerHeight', { value: 300, configurable: true });

    component.onContextMenu(ctxEventAt(600, 280));

    expect(component.contextMenuY()).toBeGreaterThanOrEqual(8);
    expect(component.showContextMenu()).toBe(true);
  });

  it('przy prawej krawędzi przesuwa menu w lewo, by zmieściło się w viewport', () => {
    Object.defineProperty(window, 'innerWidth', { value: 1000, configurable: true });
    Object.defineProperty(window, 'innerHeight', { value: 900, configurable: true });

    component.onContextMenu(ctxEventAt(995, 100));

    expect(component.contextMenuX()).toBeLessThanOrEqual(1000 - 260 - 8);
    expect(component.contextMenuX()).toBeGreaterThanOrEqual(8);
  });

  it('w typowym miejscu otwiera się pod kursorem (bez zbędnego przesuwania)', () => {
    Object.defineProperty(window, 'innerWidth', { value: 1600, configurable: true });
    Object.defineProperty(window, 'innerHeight', { value: 1000, configurable: true });

    component.onContextMenu(ctxEventAt(400, 200));

    expect(component.contextMenuX()).toBe(400);
    expect(component.contextMenuY()).toBe(200);
  });
});

/**
 * Flow „Zakończ" → modal „Trwa wysyłanie pliku" + odliczanie 30 s + próba zamknięcia karty.
 *
 * Nie wołamy detectChanges()/ngOnInit (jak inne testy tego pliku), więc nie startuje auto-save
 * ani route-load. Eksperymentalny runner (Vitest) nie wspiera `fakeAsync`, więc odliczanie
 * testujemy deterministycznie wołając wydzielony `onFinishCountdownTick(n)` (zamiast czekać na
 * realny timer), a łańcuch save→finish (microtask) domykamy `vi.waitFor`. blobToBase64
 * (FileReader) jest stubowane na Promise.resolve. afterEach woła ngOnDestroy → sprząta timer.
 */
describe('DocumentEditorComponent — flow „Zakończ" (modal + odliczanie + zamknięcie karty)', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;
  let finishCalls: number;
  let finishResult: { next?: unknown; error?: unknown };

  beforeEach(async () => {
    finishCalls = 0;
    finishResult = { next: { deliveryId: 'd-1', status: 'Pending', statusUrl: '/x' } };

    const storageMock = {
      finishAndSend: () => {
        finishCalls++;
        return finishResult.error
          ? throwError(() => finishResult.error)
          : of(finishResult.next);
      },
    };

    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: {
            getTemplates: () => of([]),
            saveDocument: () => of(new Blob(['<p></p>'], { type: 'text/html' })),
        } },
        { provide: DocumentStorageService, useValue: storageMock },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: {} },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;

    // Wymagane do przejścia guardów + deterministyczny base64 (bez realnego FileReadera).
    component.documentMasterId.set('m-1');
    component.documentVersionId.set('v-1');
    component.returnUrl.set('https://app.example.com/return');
    (component as any).blobToBase64 = () => Promise.resolve('PHA+PC9wPg==');
  });

  afterEach(() => {
    // Sprzątnij realny timer odliczania uruchomiony przez finishDocument() (brak wycieku setInterval).
    component.ngOnDestroy();
  });

  /** Zatrzymuje realny timer i steruje odliczaniem ręcznie (deterministycznie, bez fake-timerów). */
  const stopRealCountdownTimer = () => (component as any).finishCountdownSub?.unsubscribe();
  const tickCountdown = (n: number) => (component as any).onFinishCountdownTick(n);

  it('kliknięcie „Zakończ" pokazuje modal z tekstem i odliczaniem od 30', () => {
    component.finishDocument();

    expect(component.showFinishModal()).toBe(true);
    expect(component.finishCountdown()).toBe(30);
    expect(component.isFinishing()).toBe(true);
  });

  it('odliczanie maleje co sekundę (30 → 29 → 28 ...)', () => {
    component.finishDocument();
    stopRealCountdownTimer();
    expect(component.finishCountdown()).toBe(30);

    tickCountdown(0);
    expect(component.finishCountdown()).toBe(29);
    tickCountdown(1);
    expect(component.finishCountdown()).toBe(28);
  });

  it('po dojściu odliczania do 0 modal się zamyka i próbuje zamknąć kartę', () => {
    const origClose = window.close;
    let closeAttempts = 0;
    (window as any).close = () => { closeAttempts++; };

    try {
      component.finishDocument();
      stopRealCountdownTimer();

      // 30. tyk (indeks 29) → remaining = 0 → ta sama akcja co przycisk „Zamknij".
      tickCountdown(29);

      expect(component.showFinishModal()).toBe(false);
      expect(component.finishCountdown()).toBe(0);
      expect(closeAttempts).toBe(1);
    } finally {
      (window as any).close = origClose;
    }
  });

  it('kliknięcie „Zamknij" przed końcem odliczania robi to samo: zamyka modal + próbuje zamknąć kartę', () => {
    const origClose = window.close;
    let closeAttempts = 0;
    (window as any).close = () => { closeAttempts++; };

    try {
      component.finishDocument();

      component.closeFinishModalAndExit();

      expect(component.showFinishModal()).toBe(false);
      expect(closeAttempts).toBe(1);
      expect(component.isFinishing()).toBe(false);
      // Subskrypcja odliczania została zatrzymana (brak wycieku).
      expect((component as any).finishCountdownSub).toBeUndefined();
    } finally {
      (window as any).close = origClose;
    }
  });

  it('po „Zamknij" wchodzi w stan końcowy (workFinished) i wyłącza auto-save — brak powrotu do edycji', () => {
    const origClose = window.close;
    // Symulacja przeglądarki, która NIE zamyka karty (typowe dla kart nieotwartych skryptem).
    (window as any).close = () => { /* brak efektu — karta zostaje */ };

    try {
      component.autoSaveEnabled.set(true);
      component.finishDocument();

      component.closeFinishModalAndExit();

      expect(component.workFinished()).toBe(true);     // blokujący ekran końcowy
      expect(component.autoSaveEnabled()).toBe(false);  // auto-save zatrzymany
      expect(component.showFinishModal()).toBe(false);
    } finally {
      (window as any).close = origClose;
    }
  });

  it('błąd natychmiastowej wysyłki pokazuje komunikat o ponowieniu w tle, a modal zostaje otwarty', async () => {
    finishResult = { error: new Error('network') };
    // mockImplementation, bo realny showError planuje setTimeout(5 s) na wyczyszczenie toasta —
    // wyciekłby poza teardown testu. Asercja sprawdza dokładny komunikat.
    const errSpy = vi.spyOn(component as any, 'showError').mockImplementation(() => {});

    component.finishDocument();

    await vi.waitFor(() =>
      expect(errSpy).toHaveBeenCalledWith('Nie udało się natychmiast wysłać pliku. Ponowimy próbę wysłania w tle.'),
    );
    expect(component.showFinishModal()).toBe(true); // nie blokujemy — użytkownik może zamknąć
  });

  it('wielokrotne kliknięcie „Zakończ" nie tworzy wielu równoległych flow', async () => {
    component.finishDocument();
    component.finishDocument(); // zablokowane guardem isFinishing
    component.finishDocument();

    await vi.waitFor(() => expect(finishCalls).toBe(1));
  });

  it('gdy window.close() rzuci wyjątek, aplikacja nie crashuje (best-effort)', () => {
    const origClose = window.close;
    (window as any).close = () => { throw new Error('blocked by browser'); };

    try {
      component.finishDocument();
      expect(() => component.closeFinishModalAndExit()).not.toThrow();
      expect(component.showFinishModal()).toBe(false);
    } finally {
      (window as any).close = origClose;
    }
  });

  it('w trybie podglądu (readOnly) „Zakończ" nie otwiera modala', () => {
    component.readOnly.set(true);

    component.finishDocument();

    expect(component.showFinishModal()).toBe(false);
    expect(component.isFinishing()).toBe(false);
  });
});

/**
 * Dialog hasła do zaszyfrowanego dokumentu — zastępuje window.prompt.
 */
describe('DocumentEditorComponent — dialog hasła', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: { getTemplates: () => of([]) } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: {} },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
  });

  it('confirm bez hasła nie zamyka i pokazuje błąd', () => {
    component.showPasswordDialog.set(true);
    component.passwordDialogValue = '';

    component.confirmPasswordDialog();

    expect(component.showPasswordDialog()).toBe(true);
    expect(component.passwordDialogError()).toBeTruthy();
  });

  it('cancel zamyka dialog i czyści stan', () => {
    component.showPasswordDialog.set(true);
    component.passwordDialogValue = 'x';
    component.passwordDialogError.set('err');

    component.cancelPasswordDialog();

    expect(component.showPasswordDialog()).toBe(false);
    expect(component.passwordDialogValue).toBe('');
    expect(component.passwordDialogError()).toBeNull();
  });

  it('confirm z hasłem zamyka dialog', () => {
    component.showPasswordDialog.set(true);
    component.passwordDialogValue = 'sezam'; // brak pliku oczekującego → tylko zamknięcie

    component.confirmPasswordDialog();

    expect(component.showPasswordDialog()).toBe(false);
  });
});

/**
 * Kluczowe: dialog hasła wyzwalany w ścieżce KONWERSJI (_convertAndLoad), z której korzysta
 * zarówno otwarcie z dysku, jak i ładowanie wersji z bazy (loadFromStorage = dashboard/odświeżenie).
 */
describe('DocumentEditorComponent — dialog hasła wyzwalany przy konwersji', () => {
  let fixture: ComponentFixture<DocumentEditorComponent>;
  let component: DocumentEditorComponent;
  let openCalls: any[][];
  let openResult: any;

  beforeEach(async () => {
    openCalls = [];
    openResult = of({ html: '', metadata: {}, styles: [] });
    await TestBed.configureTestingModule({
      imports: [DocumentEditorComponent],
      providers: [
        { provide: DocumentService, useValue: {
            getTemplates: () => of([]),
            openDocument: (...args: any[]) => { openCalls.push(args); return openResult; }
        } },
        { provide: DocumentStorageService, useValue: {} },
        { provide: Router, useValue: { navigate: () => {} } },
        { provide: ActivatedRoute, useValue: { queryParams: of({}) } },
        { provide: BuildInfoService, useValue: {} },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentEditorComponent);
    component = fixture.componentInstance;
  });

  it('PASSWORD_REQUIRED z /open otwiera dialog hasła', () => {
    openResult = throwError(() => new OpenDocumentError('zabezpieczony', 'PASSWORD_REQUIRED'));

    (component as any)._convertAndLoad(new File([], 'tajne.docx'), 'tajne.docx');

    expect(component.showPasswordDialog()).toBe(true);
  });

  it('zatwierdzenie hasła ponawia konwersję z podanym hasłem', () => {
    openResult = throwError(() => new OpenDocumentError('zabezpieczony', 'PASSWORD_REQUIRED'));
    (component as any)._convertAndLoad(new File([], 'tajne.docx'), 'tajne.docx');

    openResult = of({ html: '<p>ok</p>', metadata: {}, styles: [] }); // poprawne hasło → sukces
    component.passwordDialogValue = 'sezam';
    component.confirmPasswordDialog();

    // Ostatnie wywołanie openDocument dostało hasło jako drugi argument.
    expect(openCalls[openCalls.length - 1][1]).toBe('sezam');
    expect(component.showPasswordDialog()).toBe(false);
  });

  it('WRONG_PASSWORD pokazuje dialog z komunikatem błędu', () => {
    openResult = throwError(() => new OpenDocumentError('złe', 'WRONG_PASSWORD'));

    (component as any)._convertAndLoad(new File([], 'tajne.docx'), 'tajne.docx');

    expect(component.showPasswordDialog()).toBe(true);
    expect(component.passwordDialogError()).toBeTruthy();
  });
});
