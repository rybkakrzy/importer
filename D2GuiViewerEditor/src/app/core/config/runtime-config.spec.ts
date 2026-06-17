import { mergeAuthConfig, DEFAULT_AUTH_CONFIG } from './runtime-config';

/**
 * Runtime config merge: config.json (Doc2) is the source of auth values and overrides the
 * structural DEFAULT_AUTH_CONFIG field-by-field; missing/invalid config falls back to those
 * defaults so the app still bootstraps without config.json.
 */
describe('mergeAuthConfig', () => {
  it('returns the build-time defaults when the runtime config is null/undefined', () => {
    expect(mergeAuthConfig(null)).toEqual(DEFAULT_AUTH_CONFIG);
    expect(mergeAuthConfig(undefined)).toEqual(DEFAULT_AUTH_CONFIG);
    expect(mergeAuthConfig({})).toEqual(DEFAULT_AUTH_CONFIG);
  });

  it('overrides only the fields present in the runtime config', () => {
    const merged = mergeAuthConfig({
      clientId: 'real-client-id',
      authority: 'https://login.microsoftonline.com/real-tenant',
    });

    expect(merged.clientId).toBe('real-client-id');
    expect(merged.authority).toBe('https://login.microsoftonline.com/real-tenant');
    // Untouched fields keep the defaults.
    expect(merged.redirectUri).toBe(DEFAULT_AUTH_CONFIG.redirectUri);
    expect(merged.postLogoutRedirectUri).toBe(DEFAULT_AUTH_CONFIG.postLogoutRedirectUri);
    expect(merged.apiScopes).toEqual(DEFAULT_AUTH_CONFIG.apiScopes);
  });

  it('replaces apiScopes when provided, keeps default array when absent', () => {
    expect(mergeAuthConfig({ apiScopes: ['api://x/access'] }).apiScopes).toEqual(['api://x/access']);
    expect(mergeAuthConfig({}).apiScopes).toEqual(DEFAULT_AUTH_CONFIG.apiScopes);
  });

  it('does not mutate the defaults object', () => {
    const snapshot = JSON.stringify(DEFAULT_AUTH_CONFIG);
    mergeAuthConfig({ clientId: 'x', apiScopes: ['y'] });
    expect(JSON.stringify(DEFAULT_AUTH_CONFIG)).toBe(snapshot);
  });
});
