import { TestBed, ComponentFixture } from '@angular/core/testing';
import { of } from 'rxjs';
import { AdminFilesComponent } from './admin-files';
import { AdminService, DocumentListItem, DocumentVersionListItem } from '../../../services/admin.service';

/**
 * Lista plików admina: filtr ID musi znajdować po MasterID ORAZ VersionID (wersje
 * dociągane raz przy pierwszym wyszukiwaniu), a filtr typu działa po ROZSZERZENIU
 * (doc/docx/pdf — etykieta „Word" nie rozróżniała .doc od .docx i „doc" nie znajdowało nic).
 */

function doc(masterId: string, name: string, mimeType: string): DocumentListItem {
  return {
    masterId,
    name,
    mimeType,
    createdAt: '2026-07-22T12:26:00Z',
    activeVersionId: `active-${masterId}`,
    versionNumber: 1,
    status: 'Saved',
    lastModifiedBy: null,
  };
}

describe('AdminFilesComponent — filtry Master/Version ID i rozszerzeń', () => {
  let fixture: ComponentFixture<AdminFilesComponent>;
  let component: AdminFilesComponent;
  let versionCalls: string[];

  const docs = [
    doc('master-aaa', 'umowa.docx',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
    doc('master-bbb', 'stary.doc', 'application/msword'),
    doc('master-ccc', 'pismo.pdf', 'application/pdf'),
  ];
  const versionsByMaster: Record<string, DocumentVersionListItem[]> = {
    'master-aaa': [{
      versionId: 'ver-hidden-123', versionNumber: 2, createdAt: '2026-07-22T12:30:00Z',
      createdBy: 'x', isActive: false, sizeInBytes: 100,
    }],
    'master-bbb': [],
    'master-ccc': [],
  };

  beforeEach(async () => {
    versionCalls = [];
    await TestBed.configureTestingModule({
      imports: [AdminFilesComponent],
      providers: [{
        provide: AdminService,
        useValue: {
          getAllDocuments: () => of(docs),
          getDocumentVersions: (masterId: string) => {
            versionCalls.push(masterId);
            return of(versionsByMaster[masterId] ?? []);
          },
        },
      }],
    }).compileComponents();
    fixture = TestBed.createComponent(AdminFilesComponent);
    component = fixture.componentInstance;
    component.ngOnInit();
  });

  it('filtr ID znajduje po MasterID', () => {
    component.setFilter('id', 'master-bbb');
    expect(component.documents().map(d => d.masterId)).toEqual(['master-bbb']);
  });

  it('filtr ID znajduje po VersionID aktywnej wersji', () => {
    component.setFilter('id', 'active-master-ccc');
    expect(component.documents().map(d => d.masterId)).toEqual(['master-ccc']);
  });

  it('filtr ID dociąga wersje RAZ i znajduje po VersionID nieaktywnej wersji', () => {
    component.setFilter('id', 'ver-hidden');
    expect(component.documents().map(d => d.masterId)).toEqual(['master-aaa']);
    // Kolejne wpisywanie nie ponawia dociągania (stan preloadu = done).
    component.setFilter('id', 'ver-hidden-1');
    expect(versionCalls.filter(m => m === 'master-aaa').length).toBe(1);
  });

  it('filtr rozszerzenia rozróżnia doc od docx (etykieta „Word" tego nie umiała)', () => {
    component.setFilter('extension', 'doc');
    expect(component.documents().map(d => d.masterId)).toEqual(['master-bbb']);

    component.setFilter('extension', 'docx');
    expect(component.documents().map(d => d.masterId)).toEqual(['master-aaa']);

    component.setFilter('extension', 'pdf');
    expect(component.documents().map(d => d.masterId)).toEqual(['master-ccc']);
  });

  it('opcje selecta = rozszerzenia obecne na liście (posortowane)', () => {
    expect(component.availableExtensions()).toEqual(['doc', 'docx', 'pdf']);
  });

  it('rozszerzenie liczone z nazwy pliku, fallback z MIME; chip pokazuje uppercase', () => {
    expect(component.fileExtension({ name: 'x.DOCX', mimeType: 'application/pdf' })).toBe('docx');
    expect(component.fileExtension({ name: '', mimeType: 'application/msword' })).toBe('doc');
    expect(component.extensionLabel(docs[2])).toBe('PDF');
  });
});
