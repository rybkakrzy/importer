import { InjectionToken } from '@angular/core';

/**
 * Auth configuration consumed by MSAL. In the Doc2 pattern this is the single source of auth
 * values: loaded at runtime from `assets/configs/config.json` (filled per environment at deploy,
 * so one build is deployed everywhere). Auth values are intentionally NOT kept in `environment.*`.
 */
export interface AppAuthConfig {
  clientId: string;
  authority: string;
  redirectUri: string;
  postLogoutRedirectUri: string;
  apiScopes: string[];
  enabled?: boolean;
}

/**
 * Structural fallback only — used when `config.json` is missing/invalid and in tests. The real
 * per-environment values (clientId, authority, apiScopes) live in `config.json`, NOT here.
 * redirectUri/postLogoutRedirectUri default to '/' (MSAL resolves it against the current origin).
 * Authorization is resource-based (backend GET /api/identity/resources) — no role names live on the
 * frontend. config.json overrides these field-by-field.
 */
export const DEFAULT_AUTH_CONFIG: AppAuthConfig = {
  clientId: '',
  authority: '',
  redirectUri: '/',
  postLogoutRedirectUri: '/',
  apiScopes: [],
  enabled: true,
};

/**
 * Provided at bootstrap (see main.ts) after the runtime config.json has been fetched. Has a root
 * factory default (structural fallback) so injection always resolves — in tests and as a safety net
 * — and main.ts simply overrides it with the runtime value.
 */
export const MSAL_CUSTOM_CONFIG = new InjectionToken<AppAuthConfig>('MSAL_CUSTOM_CONFIG', {
  providedIn: 'root',
  factory: () => DEFAULT_AUTH_CONFIG,
});

/** Merges a partial runtime config over the build-time defaults (missing fields keep defaults). */
export function mergeAuthConfig(raw: Partial<AppAuthConfig> | undefined | null): AppAuthConfig {
  const a = raw ?? {};
  return {
    clientId: a.clientId ?? DEFAULT_AUTH_CONFIG.clientId,
    authority: a.authority ?? DEFAULT_AUTH_CONFIG.authority,
    redirectUri: a.redirectUri ?? DEFAULT_AUTH_CONFIG.redirectUri,
    postLogoutRedirectUri: a.postLogoutRedirectUri ?? DEFAULT_AUTH_CONFIG.postLogoutRedirectUri,
    apiScopes: a.apiScopes ?? DEFAULT_AUTH_CONFIG.apiScopes,
    enabled: a.enabled ?? DEFAULT_AUTH_CONFIG.enabled ?? true,
  };
}

/** Auth disabled = local dev bypass (build-time default is enabled). */
export function isAuthDisabled(config: AppAuthConfig): boolean {
  return config.enabled === false;
}
