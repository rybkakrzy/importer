import { InjectionToken } from '@angular/core';
import { environment } from '../../../environments/environment';

/**
 * Auth configuration consumed by MSAL. In the Doc2 pattern this is loaded at runtime from
 * `assets/configs/config.json` (so a single build is deployed to every environment without
 * rebuilding). Build-time `environment.auth` provides the defaults / fallback.
 */
export interface AppAuthConfig {
  clientId: string;
  authority: string;
  redirectUri: string;
  postLogoutRedirectUri: string;
  apiScopes: string[];
  adminRole: string;
  enabled?: boolean;
}

/** Build-time defaults; runtime config.json (when present) overrides these field-by-field. */
export const DEFAULT_AUTH_CONFIG: AppAuthConfig = { ...environment.auth };

/**
 * Provided at bootstrap (see main.ts) after the runtime config.json has been fetched. Has a root
 * factory default (build-time config) so injection always resolves — in tests and as a safety net
 * — and main.ts simply overrides it with the runtime value.
 */
export const RUNTIME_AUTH_CONFIG = new InjectionToken<AppAuthConfig>('RUNTIME_AUTH_CONFIG', {
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
    adminRole: a.adminRole ?? DEFAULT_AUTH_CONFIG.adminRole,
    enabled: a.enabled ?? DEFAULT_AUTH_CONFIG.enabled ?? true,
  };
}

/** Auth disabled = local dev bypass (build-time default is enabled). */
export function isAuthDisabled(config: AppAuthConfig): boolean {
  return config.enabled === false;
}
