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
