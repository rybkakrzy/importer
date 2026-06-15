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
 * acquires tokens and gates navigation. Auth values come from the runtime config (Doc2
 * `assets/configs/config.json`, injected as RUNTIME_AUTH_CONFIG) with build-time fallback.
 */
export function msalInstanceFactory(auth: AppAuthConfig): IPublicClientApplication {
  return new PublicClientApplication({
    auth: {
      clientId: auth.clientId,
      authority: auth.authority,
      redirectUri: auth.redirectUri,
      postLogoutRedirectUri: auth.postLogoutRedirectUri,
    },
    cache: {
      cacheLocation: BrowserCacheLocation.LocalStorage,
    },
  });
}

export function msalGuardConfigFactory(auth: AppAuthConfig): MsalGuardConfiguration {
  return {
    interactionType: InteractionType.Redirect,
    authRequest: { scopes: auth.apiScopes },
  };
}

/**
 * Attaches access tokens (with the API scope) to requests hitting the backend API.
 */
export function msalInterceptorConfigFactory(auth: AppAuthConfig): MsalInterceptorConfiguration {
  const protectedResourceMap = new Map<string, Array<string>>();
  if (auth.enabled !== false) {
    protectedResourceMap.set(environment.apiUrl, auth.apiScopes);
  }

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap,
  };
}
