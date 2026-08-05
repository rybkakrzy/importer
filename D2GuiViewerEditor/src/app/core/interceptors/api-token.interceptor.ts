import { HttpInterceptorFn } from '@angular/common/http';
import { inject } from '@angular/core';
import { EMPTY, from, of, switchMap, catchError } from 'rxjs';
import { MsalService } from '@azure/msal-angular';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { environment } from '../../../environments/environment';
import { MSAL_CUSTOM_CONFIG } from '../config/runtime-config';
import { ensureInteractiveReauth, hasFreshIdToken } from '../auth/interactive-reauth';

// API base resolved once (environment.apiUrl may be scheme-relative, e.g. `//host:15111/api`).
const apiBase = new URL(environment.apiUrl, window.location.origin).toString();
const apiBaseWithSlash = apiBase.endsWith('/') ? apiBase : `${apiBase}/`;
const healthUrl = `${apiBaseWithSlash}health`;

/** localStorage key under which the bearer token sent to the API is cached. */
export const API_ACCESS_TOKEN_KEY = 'api_access_token';

/** True for backend API calls that must carry a token — excludes the anonymous health endpoint. */
export function isProtectedApiRequest(url: string): boolean {
  const absolute = new URL(url, window.location.origin).toString();
  if (absolute === healthUrl || absolute.startsWith(`${healthUrl}/`)) {
    return false;
  }
  return absolute === apiBase || absolute.startsWith(apiBaseWithSlash);
}

/**
 * Auth header for backend API calls.
 *
 * The SPA and the API share a single Entra app registration (clientId == AzureAd:ClientId) and the
 * tenant does not expose a custom API scope. So instead of a Microsoft Graph access token (scope
 * `User.Read`), whose signature this API cannot validate ("The signature is invalid"), we attach the
 * ID token: its `aud` equals the clientId, which `AddMicrosoftIdentityWebApi` accepts. We acquire a
 * token silently for the SAME scopes used at login (already consented and cached, so no /token 400)
 * and use the `idToken` from that result — the Graph access token in the same result is ignored.
 */
export const apiTokenInterceptor: HttpInterceptorFn = (req, next) => {
  const auth = inject(MSAL_CUSTOM_CONFIG);

  // Auth disabled (local dev bypass) or non-API request → pass through untouched.
  if (auth.enabled === false || !isProtectedApiRequest(req.url)) {
    return next(req);
  }

  const msal = inject(MsalService);
  const account = msal.instance.getActiveAccount() ?? msal.instance.getAllAccounts()[0] ?? null;
  if (!account) {
    return next(req);
  }

  // Same scopes as the MSAL login request (the runtime-config apiScopes, e.g. User.Read) so the
  // silent call hits the cache instead of failing at the token endpoint.
  const loginScopes = auth.apiScopes?.length ? auth.apiScopes : ['openid', 'profile'];

  // Bug 13942097: silent musi dostarczyć ŚWIEŻY idToken. Cache hit MSAL waliduje tylko access
  // token — idToken sprzed godzin przechodził do API i wracał 401 bez żadnej próby ponownego
  // logowania. Przeterminowany → forceRefresh; interaction_required (wygasła sesja Entra) →
  // interaktywny redirect do logowania (jedyny moment w aplikacji, który go inicjuje).
  const acquireFreshIdToken = async (): Promise<string | null> => {
    let result = await msal.instance.acquireTokenSilent({ account, scopes: loginScopes });
    if (!hasFreshIdToken(result)) {
      result = await msal.instance.acquireTokenSilent({
        account,
        scopes: loginScopes,
        forceRefresh: true,
      });
    }
    // Po forceRefresh idToken jest świeżo wystawiony przez Entra — jeśli lokalny zegar wciąż
    // uznaje go za przeterminowany, to przestawiony zegar klienta, nie wygasła sesja (realnie
    // wygasła sesja nie dochodzi tutaj: forceRefresh rzuca InteractionRequiredAuthError).
    // Wysyłamy — o ważności rozstrzyga backend, a 401 łapie httpErrorInterceptor. Wcześniejsze
    // rzucanie stale_id_token wpędzało klienta ze skewem zegara w pętlę redirectów logowania.
    return result.idToken || null;
  };

  return from(acquireFreshIdToken()).pipe(
    catchError((error) => {
      if (error instanceof InteractionRequiredAuthError) {
        if (ensureInteractiveReauth(msal.instance, loginScopes)) {
          // Strona odpływa do /authorize — bieżące żądanie wygaszamy bez emisji (żaden toast
          // „nieoczekiwany błąd" nie powinien mignąć w trakcie nawigacji do logowania).
          return EMPTY;
        }
        // Bezpiecznik antypętlowy zatrzymał redirect — żądanie leci bez tokenu, żeby stan
        // (401/„brak uprawnień") był widoczny zamiast wygaszania wszystkiego po cichu.
        return of(null);
      }
      // Inne błędy silent (sieć itp.) → jak dotąd: żądanie bez nagłówka; backend odpowie 401,
      // a obsługa 401 w httpErrorInterceptor jest siatką bezpieczeństwa. Downstream HTTP errors
      // are NOT caught here (no `next(req)` inside catchError), so a 401 never silently retries.
      return of(null);
    }),
    switchMap((token: string | null) => {
      // Persist the token we send to the API so other parts of the app (and debugging) can read it.
      if (token) {
        localStorage.setItem(API_ACCESS_TOKEN_KEY, token);
      }
      const authedReq = token
        ? req.clone({ setHeaders: { Authorization: `Bearer ${token}` } })
        : req;
      return next(authedReq);
    }),
  );
};
