import { TestBed } from '@angular/core/testing';
import { HttpRequest, HttpErrorResponse, HttpHandlerFn } from '@angular/common/http';
import { Router } from '@angular/router';
import { lastValueFrom, throwError } from 'rxjs';
import { MsalService } from '@azure/msal-angular';
import { httpErrorInterceptor } from './http-error.interceptor';
import { NotificationService } from '../services/notification.service';
import { ConnectionStatusService } from '../services/connection-status.service';
import { LastHttpErrorService } from '../services/last-http-error.service';
import { MSAL_CUSTOM_CONFIG, DEFAULT_AUTH_CONFIG } from '../config/runtime-config';
import { resetInteractiveReauthForTests } from '../auth/interactive-reauth';
import { environment } from '../../../environments/environment';

/**
 * Bug 13942097: 401 z chronionego API = wygasła sesja → interceptor inicjuje redirect
 * logowania (bez toastu). 403 zachowuje dotychczasową nawigację na /brak-uprawnien.
 */
describe('httpErrorInterceptor — 401 jako wygasła sesja', () => {
  const apiUrl = new URL(`${environment.apiUrl}/identity/resources`, window.location.origin).toString();

  let redirectCalls: number;
  let errorToasts: string[];
  let navigations: string[];

  function setup(authEnabled: boolean): void {
    resetInteractiveReauthForTests();
    redirectCalls = 0;
    errorToasts = [];
    navigations = [];
    TestBed.configureTestingModule({
      providers: [
        { provide: NotificationService, useValue: { error: (m: string) => errorToasts.push(m) } },
        { provide: ConnectionStatusService, useValue: { reportApiError: () => {} } },
        { provide: LastHttpErrorService, useValue: { record: () => {} } },
        { provide: Router, useValue: { navigateByUrl: (u: string) => (navigations.push(u), Promise.resolve(true)) } },
        {
          provide: MSAL_CUSTOM_CONFIG,
          useValue: { ...DEFAULT_AUTH_CONFIG, enabled: authEnabled, apiScopes: ['User.Read'] },
        },
        {
          provide: MsalService,
          useValue: {
            instance: {
              getActiveAccount: () => ({ homeAccountId: 'acc' }),
              getAllAccounts: () => [{ homeAccountId: 'acc' }],
              acquireTokenRedirect: () => (redirectCalls++, Promise.resolve()),
            },
          },
        },
      ],
    });
  }

  function runWithStatus(status: number, url = apiUrl): Promise<unknown> {
    const next: HttpHandlerFn = () =>
      throwError(() => new HttpErrorResponse({ status, url }));
    return lastValueFrom(
      TestBed.runInInjectionContext(() => httpErrorInterceptor(new HttpRequest('GET', url), next)),
    );
  }

  it('401 z chronionego API → redirect logowania, bez toastu, błąd propagowany', async () => {
    setup(true);

    await expect(runWithStatus(401)).rejects.toBeInstanceOf(HttpErrorResponse);

    expect(redirectCalls).toBe(1);
    expect(errorToasts).toEqual([]);
    expect(navigations).toEqual([]);
  });

  it('401 przy wyłączonym auth (dev bypass) → zwykła ścieżka błędu, bez redirectu', async () => {
    setup(false);

    await expect(runWithStatus(401)).rejects.toBeInstanceOf(HttpErrorResponse);

    expect(redirectCalls).toBe(0);
    expect(errorToasts.length).toBe(1);
  });

  it('403 nadal nawiguje na /brak-uprawnien (regresja)', async () => {
    setup(true);

    await expect(runWithStatus(403)).rejects.toBeInstanceOf(HttpErrorResponse);

    expect(redirectCalls).toBe(0);
    expect(navigations).toEqual(['/brak-uprawnien']);
  });
});
