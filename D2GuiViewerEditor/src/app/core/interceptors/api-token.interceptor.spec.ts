import { TestBed } from '@angular/core/testing';
import { HttpRequest, HttpResponse, HttpEvent, HttpHandlerFn } from '@angular/common/http';
import { firstValueFrom, Observable, of, lastValueFrom } from 'rxjs';
import { toArray } from 'rxjs/operators';
import { MsalService } from '@azure/msal-angular';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { apiTokenInterceptor } from './api-token.interceptor';
import { MSAL_CUSTOM_CONFIG, DEFAULT_AUTH_CONFIG } from '../config/runtime-config';
import { resetInteractiveReauthForTests } from '../auth/interactive-reauth';
import { environment } from '../../../environments/environment';

/**
 * Bug 13942097: interceptor tokenu musi (a) odrzucić przeterminowany idToken z cache MSAL
 * (forceRefresh), (b) przy wygasłej sesji (interaction_required) zainicjować redirect
 * logowania zamiast wysyłać żądanie bez nagłówka i zostawiać użytkownika na 401.
 */
describe('apiTokenInterceptor — świeżość idToken i re-autoryzacja', () => {
  const apiUrl = new URL(`${environment.apiUrl}/document/1`, window.location.origin).toString();
  const futureExp = Math.floor(Date.now() / 1000) + 3600;
  const pastExp = Math.floor(Date.now() / 1000) - 60;

  let silentCalls: Array<Record<string, unknown>>;
  let redirectCalls: number;
  let silentImpl: (req: Record<string, unknown>) => Promise<unknown>;

  beforeEach(() => {
    resetInteractiveReauthForTests();
    silentCalls = [];
    redirectCalls = 0;
    const msalStub = {
      instance: {
        getActiveAccount: () => ({ homeAccountId: 'acc' }),
        getAllAccounts: () => [{ homeAccountId: 'acc' }],
        acquireTokenSilent: (req: Record<string, unknown>) => {
          silentCalls.push(req);
          return silentImpl(req);
        },
        acquireTokenRedirect: () => {
          redirectCalls++;
          return Promise.resolve();
        },
      },
    };
    TestBed.configureTestingModule({
      providers: [
        { provide: MsalService, useValue: msalStub },
        {
          provide: MSAL_CUSTOM_CONFIG,
          useValue: { ...DEFAULT_AUTH_CONFIG, enabled: true, apiScopes: ['User.Read'] },
        },
      ],
    });
  });

  function run(url = apiUrl): Promise<HttpEvent<unknown>[]> {
    const captured: HttpRequest<unknown>[] = [];
    const next: HttpHandlerFn = (req) => {
      captured.push(req);
      return of(new HttpResponse({ status: 200 })) as Observable<HttpEvent<unknown>>;
    };
    const events = TestBed.runInInjectionContext(() =>
      apiTokenInterceptor(new HttpRequest('GET', url), next),
    );
    return lastValueFrom(events.pipe(toArray()), { defaultValue: [] }).then((evts) => {
      (run as unknown as { lastCaptured: HttpRequest<unknown>[] }).lastCaptured = captured;
      return evts;
    });
  }
  const captured = () => (run as unknown as { lastCaptured: HttpRequest<unknown>[] }).lastCaptured;

  it('świeży idToken idzie w nagłówku bez forceRefresh', async () => {
    silentImpl = () => Promise.resolve({ idToken: 'FRESH', idTokenClaims: { exp: futureExp } });

    await run();

    expect(silentCalls.length).toBe(1);
    expect(captured()[0].headers.get('Authorization')).toBe('Bearer FRESH');
    expect(redirectCalls).toBe(0);
  });

  it('przeterminowany idToken z cache wymusza forceRefresh i wysyła odświeżony', async () => {
    silentImpl = (req) =>
      Promise.resolve(
        req['forceRefresh']
          ? { idToken: 'REFRESHED', idTokenClaims: { exp: futureExp } }
          : { idToken: 'STALE', idTokenClaims: { exp: pastExp } },
      );

    await run();

    expect(silentCalls.length).toBe(2);
    expect(silentCalls[1]['forceRefresh']).toBe(true);
    expect(captured()[0].headers.get('Authorization')).toBe('Bearer REFRESHED');
  });

  it('interaction_required → redirect logowania, żądanie wygaszone (bez wysyłki bez tokenu)', async () => {
    silentImpl = () => Promise.reject(new InteractionRequiredAuthError('login_required'));

    const events = await run();

    expect(redirectCalls).toBe(1);
    expect(events).toEqual([]); // EMPTY — strona odpływa do /authorize
    expect(captured().length).toBe(0);
  });

  it('forceRefresh nadal „przeterminowany" lokalnie (skew zegara) → wysyła token, bez redirectu', async () => {
    // Po forceRefresh token jest świeżo wystawiony przez Entra — lokalny exp w przeszłości
    // oznacza przestawiony zegar klienta. O ważności rozstrzyga backend; pętla redirectów
    // logowania na maszynie ze skewem to regresja z pierwszej wersji fixa 13942097.
    silentImpl = () => Promise.resolve({ idToken: 'SKEWED', idTokenClaims: { exp: pastExp } });

    await run();

    expect(silentCalls.length).toBe(2);
    expect(redirectCalls).toBe(0);
    expect(captured().length).toBe(1);
    expect(captured()[0].headers.get('Authorization')).toBe('Bearer SKEWED');
  });

  it('inny błąd silent (sieć) → żądanie bez nagłówka jak dotąd, bez redirectu', async () => {
    silentImpl = () => Promise.reject(new Error('network'));

    await run();

    expect(redirectCalls).toBe(0);
    expect(captured().length).toBe(1);
    expect(captured()[0].headers.has('Authorization')).toBe(false);
  });

  it('endpoint health przechodzi bez tokenu i bez MSAL', async () => {
    silentImpl = () => Promise.reject(new Error('should not be called'));
    const healthUrl = new URL(`${environment.apiUrl}/health`, window.location.origin).toString();

    await run(healthUrl);

    expect(silentCalls.length).toBe(0);
    expect(captured()[0].headers.has('Authorization')).toBe(false);
  });
});
