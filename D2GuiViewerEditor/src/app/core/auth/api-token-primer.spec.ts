import { primeApiToken } from './api-token-primer';
import { API_ACCESS_TOKEN_KEY } from '../interceptors/api-token.interceptor';
import { DEFAULT_AUTH_CONFIG, AppAuthConfig } from '../config/runtime-config';

function msalMock(overrides: Record<string, unknown> = {}) {
  return {
    getActiveAccount: vi.fn().mockReturnValue({ homeAccountId: 'a' }),
    getAllAccounts: vi.fn().mockReturnValue([]),
    acquireTokenSilent: vi.fn().mockResolvedValue({ idToken: 'ID_TOKEN' }),
    ...overrides,
  } as unknown as import('@azure/msal-browser').IPublicClientApplication;
}

const enabled: AppAuthConfig = { ...DEFAULT_AUTH_CONFIG, apiScopes: ['User.Read'], enabled: true };

describe('primeApiToken', () => {
  beforeEach(() => localStorage.removeItem(API_ACCESS_TOKEN_KEY));

  it('stores the idToken under api_access_token when an account is present', async () => {
    const msal = msalMock();
    await primeApiToken(msal, enabled);
    expect(localStorage.getItem(API_ACCESS_TOKEN_KEY)).toBe('ID_TOKEN');
    expect(msal.acquireTokenSilent).toHaveBeenCalledWith(
      expect.objectContaining({ scopes: ['User.Read'] }),
    );
  });

  it('is a no-op when auth is disabled (dev bypass)', async () => {
    const msal = msalMock();
    await primeApiToken(msal, { ...enabled, enabled: false });
    expect(msal.acquireTokenSilent).not.toHaveBeenCalled();
    expect(localStorage.getItem(API_ACCESS_TOKEN_KEY)).toBeNull();
  });

  it('does nothing when there is no account yet (redirect still settling)', async () => {
    const msal = msalMock({
      getActiveAccount: vi.fn().mockReturnValue(null),
      getAllAccounts: vi.fn().mockReturnValue([]),
    });
    await primeApiToken(msal, enabled);
    expect(msal.acquireTokenSilent).not.toHaveBeenCalled();
    expect(localStorage.getItem(API_ACCESS_TOKEN_KEY)).toBeNull();
  });

  it('swallows silent-acquisition failures without storing anything', async () => {
    const msal = msalMock({ acquireTokenSilent: vi.fn().mockRejectedValue(new Error('interaction_required')) });
    await expect(primeApiToken(msal, enabled)).resolves.toBeUndefined();
    expect(localStorage.getItem(API_ACCESS_TOKEN_KEY)).toBeNull();
  });

  it('falls back to the first cached account when none is active', async () => {
    const msal = msalMock({
      getActiveAccount: vi.fn().mockReturnValue(null),
      getAllAccounts: vi.fn().mockReturnValue([{ homeAccountId: 'b' }]),
    });
    await primeApiToken(msal, enabled);
    expect(localStorage.getItem(API_ACCESS_TOKEN_KEY)).toBe('ID_TOKEN');
  });
});
