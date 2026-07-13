import { TestBed } from '@angular/core/testing';
import { provideRouter } from '@angular/router';
import { of } from 'rxjs';
import { MsalService } from '@azure/msal-angular';
import { DashboardComponent } from './dashboard/dashboard';
import { PdfMaintenanceComponent } from './pdf-maintenance/pdf-maintenance';
import { DocumentService } from '../services/document.service';
import { DocumentStorageService } from '../services/document-storage.service';
import { DocumentNavigationService } from '../core/services/document-navigation.service';
import { ResourceAccessService } from '../core/services/resource-access.service';

/**
 * Shell-layout contract: routed page wrappers must fill the area the app shell
 * hands them (100% of their flexed host) — they must NOT size to the viewport
 * (`100vh`). The always-on global banners (environment / offline) sit above the
 * page in normal flow and consume vertical space; a page that asks for the full
 * `100vh` is therefore pushed DOWN by the banner height and its footer is clipped
 * (the dashboard bug). The editor never had this because it uses `height:100%`.
 *
 * jsdom can't lay out, so we assert on the *CSS contract* Angular injects for each
 * component (the thing that actually regressed), not on measured pixels.
 */
describe('Routed page layout — banner must not push content down', () => {
  /** Concatenated text of every injected <style> block that mentions `needle`. */
  function injectedCssFor(needle: string): string {
    return Array.from(document.querySelectorAll('style'))
      .map(s => s.textContent ?? '')
      .filter(t => t.includes(needle))
      .join('\n');
  }

  it('dashboard wrapper fills the host (100%), never the viewport (100vh)', async () => {
    await TestBed.configureTestingModule({
      imports: [DashboardComponent],
      providers: [
        { provide: DocumentService, useValue: {} },
        { provide: DocumentStorageService, useValue: {} },
        { provide: DocumentNavigationService, useValue: {} },
        // Dashboard wstrzykuje gating zasobów + MSAL — stub jak w dashboard.spec.ts,
        // inaczej ścieżka ResourceAccessService → HttpClient wywala NG0201 w jsdom.
        { provide: ResourceAccessService, useValue: { hasAccessToResource: () => of(false) } },
        { provide: MsalService, useValue: { instance: { getActiveAccount: () => null, getAllAccounts: () => [] } } },
      ],
    }).compileComponents();

    const fixture = TestBed.createComponent(DashboardComponent);
    fixture.detectChanges();

    const css = injectedCssFor('dashboard-wrapper');
    expect(css).not.toBe(''); // styles were injected
    expect(css).not.toContain('100vh'); // the banner-pushing antipattern is gone
    expect(css).toMatch(/dashboard-wrapper[^}]*min-height\s*:\s*100%/s);
  });

  it('pdf-maintenance wrapper fills the host (100%), never the viewport (100vh)', async () => {
    await TestBed.configureTestingModule({
      imports: [PdfMaintenanceComponent],
      providers: [provideRouter([])],
    }).compileComponents();

    const fixture = TestBed.createComponent(PdfMaintenanceComponent);
    fixture.detectChanges();

    const css = injectedCssFor('maintenance-wrapper');
    expect(css).not.toBe('');
    expect(css).not.toContain('100vh');
    expect(css).toMatch(/maintenance-wrapper[^}]*min-height\s*:\s*100%/s);
  });
});
