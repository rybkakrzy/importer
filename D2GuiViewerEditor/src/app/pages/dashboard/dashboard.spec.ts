import { TestBed } from '@angular/core/testing';
import { of } from 'rxjs';
import { DashboardComponent } from './dashboard';
import { DocumentService } from '../../services/document.service';
import { DocumentStorageService } from '../../services/document-storage.service';
import { DocumentNavigationService } from '../../core/services/document-navigation.service';
import { ResourceAccessService } from '../../core/services/resource-access.service';
import { MsalService } from '@azure/msal-angular';

function setup(canUseEditor: boolean) {
  const documentService = { newDocument: vi.fn().mockReturnValue(of({ html: '', metadata: '' })) };
  const documentStorageService = {};
  const documentNavigation = { navigateToEditableDocument: vi.fn() };
  const resourceAccess = { hasAccessToResource: vi.fn().mockReturnValue(of(canUseEditor)) };
  const msal = { instance: { getActiveAccount: () => null, getAllAccounts: () => [] } };

  TestBed.resetTestingModule();
  TestBed.configureTestingModule({
    providers: [
      { provide: DocumentService, useValue: documentService },
      { provide: DocumentStorageService, useValue: documentStorageService },
      { provide: DocumentNavigationService, useValue: documentNavigation },
      { provide: ResourceAccessService, useValue: resourceAccess },
      { provide: MsalService, useValue: msal },
    ],
  });
  const component = TestBed.runInInjectionContext(() => new DashboardComponent());
  return { component, documentService };
}

describe('DashboardComponent editor gating', () => {
  it('exposes canUseEditor from the "editor" resource', () => {
    expect(setup(true).component.canUseEditor()).toBe(true);
    expect(setup(false).component.canUseEditor()).toBe(false);
  });

  it('blocks newDocument when the user lacks the editor resource', () => {
    const { component, documentService } = setup(false);
    component.newDocument();
    expect(documentService.newDocument).not.toHaveBeenCalled();
    expect(component.errorMessage()).toBe('Nie masz uprawnień do edycji dokumentów.');
  });

  it('allows newDocument when the user has the editor resource', () => {
    const { component, documentService } = setup(true);
    component.newDocument();
    expect(documentService.newDocument).toHaveBeenCalled();
  });
});
