import { HttpInterceptorFn } from '@angular/common/http';
import { inject } from '@angular/core';
import { from, of, switchMap, catchError } from 'rxjs';
import { MsalService } from '@azure/msal-angular';
import { environment } from '../../../environments/environment';
import { MSAL_CUSTOM_CONFIG } from '../config/runtime-config';

// API base resolved once (environment.apiUrl may be scheme-relative, e.g. `//host:15111/api`).
const apiBase = new URL(environment.apiUrl, window.location.origin).toString();
const apiBaseWithSlash = apiBase.endsWith('/') ? apiBase : `${apiBase}/`;
const healthUrl = `${apiBaseWithSlash}health`;

/** True for backend API calls that must carry a token — excludes the anonymous health endpoint. */
function isProtectedApiRequest(url: string): boolean {
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

  return from(msal.instance.acquireTokenSilent({ account, scopes: loginScopes })).pipe(
    // Token acquisition failed (e.g. interaction required) → send without header; the backend
    // returns 401 and the global error handling kicks in. Downstream HTTP errors are NOT caught
    // here (no `next(req)` inside catchError), so a 401 never silently retries.
    catchError(() => of(null)),
    switchMap((result) => {
      const authedReq = result?.idToken
        ? req.clone({ setHeaders: { Authorization: `Bearer ${result.idToken}` } })
        : req;
      return next(authedReq);
    }),
  );
};
