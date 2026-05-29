import { TestBed, ComponentFixture } from '@angular/core/testing';
import { of } from 'rxjs';
import { ActivatedRoute, Router } from '@angular/router';
import { DocumentEditorComponent } from './document-editor';
import { DocumentService } from '../../services/document.service';
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
    component.selectedImage.set({ widthPx: 100, heightPx: 50, aspectRatio: 2, alignment: null });
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
