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
  validateAuthConfig(auth);

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

function isAuthEnabled(auth: AppAuthConfig): boolean {
  return auth.enabled !== false;
}

function validateAuthConfig(auth: AppAuthConfig): void {
  if (!isAuthEnabled(auth)) {
    return;
  }

  const missingFields = [
    ['clientId', auth.clientId],
    ['authority', auth.authority],
    ['redirectUri', auth.redirectUri],
  ].filter(([, value]) => !value || !value.toString().trim());

  if (missingFields.length > 0) {
    const names = missingFields.map(([name]) => name).join(', ');
    throw new Error(`Invalid MSAL config. Missing fields: ${names}.`);
  }

  const scopes = apiScopesFor(auth);
  const invalidScopes = scopes.filter((scope) => scope === '/.default' || scope.startsWith('/'));
  if (invalidScopes.length > 0) {
    throw new Error(`Invalid MSAL scopes: ${invalidScopes.join(', ')}.`);
  }
}

/**
 * API scopes requested for the access token. Scopes must be explicitly configured in runtime
 * config (no implicit `/.default` fallback).
 */
function apiScopesFor(auth: AppAuthConfig): string[] {
  const scopes = auth.apiScopes
    ?.filter((scope): scope is string => typeof scope === 'string')
    .map((scope) => scope.trim())
    .filter(Boolean);

  if (!scopes?.length) {
    throw new Error('Invalid MSAL config. Missing apiScopes.');
  }

  return scopes;
}

export function msalGuardConfigFactory(auth: AppAuthConfig): MsalGuardConfiguration {
  if (!isAuthEnabled(auth)) {
    return {
      interactionType: InteractionType.Redirect,
    };
  }

  validateAuthConfig(auth);

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
  if (isAuthEnabled(auth)) {
    validateAuthConfig(auth);
    const scopes = apiScopesFor(auth);
    const apiBase = new URL(environment.apiUrl, window.location.origin).toString();
    const apiBaseWithSlash = apiBase.endsWith('/') ? apiBase : `${apiBase}/`;
    protectedResourceMap.set(apiBase, scopes);
    protectedResourceMap.set(apiBaseWithSlash, scopes);
  }

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap,
  };
}
