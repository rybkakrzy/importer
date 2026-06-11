import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { Router } from '@angular/router';
import { AdminOpenDocumentComponent } from './admin-open-document';

describe('AdminOpenDocumentComponent — „Otwórz plik w edytorze"', () => {
  let fixture: ComponentFixture<AdminOpenDocumentComponent>;
  let component: AdminOpenDocumentComponent;
  let navigateSpy: ReturnType<typeof vi.fn>;

  beforeEach(async () => {
    navigateSpy = vi.fn();
    await TestBed.configureTestingModule({
      imports: [AdminOpenDocumentComponent],
      providers: [{ provide: Router, useValue: { navigate: navigateSpy } }],
    }).compileComponents();
    fixture = TestBed.createComponent(AdminOpenDocumentComponent);
    component = fixture.componentInstance;
  });

  it('canOpen jest false dla pustego masterId i true po wpisaniu', () => {
    expect(component.canOpen()).toBe(false);
    component.masterId.set('  ');
    expect(component.canOpen()).toBe(false);
    component.masterId.set('master-1');
    expect(component.canOpen()).toBe(true);
  });

  it('open() z masterId i versionId nawiguje do /editor z oboma query params', () => {
    component.masterId.set(' master-1 ');
    component.versionId.set(' version-9 ');

    component.open();

    expect(navigateSpy).toHaveBeenCalledWith(['/editor'], {
      queryParams: { masterId: 'master-1', versionId: 'version-9' },
    });
  });

  it('open() bez versionId nawiguje tylko z masterId (tryb podglądu)', () => {
    component.masterId.set('master-1');
    component.versionId.set('');

    component.open();

    expect(navigateSpy).toHaveBeenCalledWith(['/editor'], {
      queryParams: { masterId: 'master-1' },
    });
  });

  it('open() nic nie robi, gdy masterId puste', () => {
    component.masterId.set('   ');

    component.open();

    expect(navigateSpy).not.toHaveBeenCalled();
  });
});
