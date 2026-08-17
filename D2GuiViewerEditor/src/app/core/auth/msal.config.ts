import {
  IPublicClientApplication,
  PublicClientApplication,
  InteractionType,
  BrowserCacheLocation,
  LogLevel,
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
    system: {
      loggerOptions: {
        // Błędy/ostrzeżenia MSAL do konsoli (bez PII): bez tego nieudana wymiana kodu na
        // tokeny po powrocie z Entra (#code w URL) umierała bezgłośnie i środowiskowych
        // problemów auth (CSP/proxy na /token, zły redirectUri) nie dało się diagnozować.
        logLevel: LogLevel.Warning,
        piiLoggingEnabled: false,
        loggerCallback: (level: LogLevel, message: string) => {
          if (level === LogLevel.Error) {
            console.error('[msal]', message);
          } else {
            console.warn('[msal]', message);
          }
        },
      },
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
  const protectedResourceMap = new Map<string, Array<string> | null>();
  if (isAuthEnabled(auth)) {
    validateAuthConfig(auth);
    const scopes = apiScopesFor(auth);
    const apiBase = new URL(environment.apiUrl, window.location.origin).toString();
    const apiBaseWithSlash = apiBase.endsWith('/') ? apiBase : `${apiBase}/`;
    // The health endpoint is anonymous (backend HealthController has no [Authorize]) and is polled
    // at bootstrap + every 30s by ConnectionStatusService. It must NOT trigger interactive token
    // acquisition — a background poll before login would otherwise start a redirect loop to
    // /authorize. null = explicitly unprotected; listed before the broad base so the more specific
    // entry wins.
    protectedResourceMap.set(`${apiBaseWithSlash}health`, null);
    protectedResourceMap.set(apiBase, scopes);
    protectedResourceMap.set(apiBaseWithSlash, scopes);
  }

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap,
  };
}
