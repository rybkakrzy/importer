import { IPublicClientApplication } from '@azure/msal-browser';
import { AppAuthConfig } from '../config/runtime-config';
import { API_ACCESS_TOKEN_KEY } from '../interceptors/api-token.interceptor';

/**
 * Best-effort priming of the API token in localStorage right after login, so `api_access_token` is
 * available before the first backend request (useful for diagnostics). Mirrors the interceptor: same
 * login scopes, stores the `idToken` under the same key.
 *
 * Never throws and never blocks bootstrap — if there is no account yet (redirect still settling) or
 * silent acquisition needs interaction, the interceptor stores the token on the first protected
 * request instead. Dev bypass (auth disabled) is a no-op, matching the interceptor.
 */
export async function primeApiToken(
  msalInstance: IPublicClientApplication,
  auth: AppAuthConfig,
): Promise<void> {
  if (auth.enabled === false) {
    return;
  }

  const account = msalInstance.getActiveAccount() ?? msalInstance.getAllAccounts()[0] ?? null;
  if (!account) {
    return;
  }

  // Same scopes as the login request so the silent call hits the MSAL cache (no /token round-trip).
  const scopes = auth.apiScopes?.length ? auth.apiScopes : ['openid', 'profile'];

  try {
    const result = await msalInstance.acquireTokenSilent({ account, scopes });
    if (result?.idToken) {
      localStorage.setItem(API_ACCESS_TOKEN_KEY, result.idToken);
    }
  } catch {
    // interaction_required / token not ready → the interceptor stores it on the first API call.
  }
}
