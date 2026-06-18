import {
  IPublicClientApplication,
  PublicClientApplication,
  InteractionType,
  BrowserCacheLocation,
} from '@azure/msal-browser';
import {
  MsalGuardConfiguration,
  MsalInterceptorConfiguration,
} from '@azure/msal-angular';
import { environment } from '../../../environments/environment';
import { AppAuthConfig } from '../config/runtime-config';

/**
 * MSAL setup for Entra ID. Backend remains the source of truth for authorization; MSAL only
 * acquires tokens and gates navigation. Auth values come from the runtime config (Qutas
 * `assets/configs/config.json`, injected as MSAL_CUSTOM_CONFIG) with build-time fallback.
 */
export function msalInstanceFactory(auth: AppAuthConfig): IPublicClientApplication {
  return new PublicClientApplication({
    auth: {
      clientId: auth.clientId,
      authority: auth.authority,
      redirectUri: auth.redirectUri,
      postLogoutRedirectUri: auth.postLogoutRedirectUri,
      // After completing login, return the user to the page they started from (MSAL default,
      // set explicitly for clarity).
      navigateToLoginRequestUrl: true,
    },
    cache: {
      cacheLocation: BrowserCacheLocation.LocalStorage,
    },
  });
}

/**
 * API scope(s) requested for the access token. Prefers explicit `apiScopes` from config.json;
 * when none are configured, falls back to `{clientId}/.default` (all delegated permissions
 * statically granted to the App Registration) — the Qutas pattern.
 */
function apiScopesFor(auth: AppAuthConfig): string[] {
  return auth.apiScopes?.length ? auth.apiScopes : [`${auth.clientId}/.default`];
}

export function msalGuardConfigFactory(auth: AppAuthConfig): MsalGuardConfiguration {
  return {
    interactionType: InteractionType.Redirect,
    authRequest: { scopes: apiScopesFor(auth) },
  };
}

/**
 * Attaches access tokens (with the API scope) to requests hitting the backend API.
 */
export function msalInterceptorConfigFactory(auth: AppAuthConfig): MsalInterceptorConfiguration {
  const protectedResourceMap = new Map<string, Array<string>>();
  if (auth.enabled !== false) {
    protectedResourceMap.set(environment.apiUrl, apiScopesFor(auth));
  }

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap,
  };
}
