import {
  ensureInteractiveReauth,
  hasFreshIdToken,
  resetInteractiveReauthForTests,
} from './interactive-reauth';
import { AuthenticationResult, IPublicClientApplication } from '@azure/msal-browser';

/**
 * Bug 13942097: po wygaśnięciu sesji aplikacja nie inicjowała ponownej autoryzacji —
 * helper jest JEDYNYM wyzwalaczem interaktywnego redirectu i musi (a) odpalić go raz
 * na życie strony, (b) rozpoznawać przeterminowany idToken z cache MSAL.
 */
describe('interactive-reauth', () => {
  beforeEach(() => resetInteractiveReauthForTests());

  function msalStub(): { instance: IPublicClientApplication; redirects: unknown[] } {
    const redirects: unknown[] = [];
    const instance = {
      getActiveAccount: () => ({ homeAccountId: 'acc' }),
      getAllAccounts: () => [{ homeAccountId: 'acc' }],
      acquireTokenRedirect: (req: unknown) => {
        redirects.push(req);
        return Promise.resolve();
      },
    } as unknown as IPublicClientApplication;
    return { instance, redirects };
  }

  it('inicjuje redirect logowania dokładnie raz (równoległe 401 nie dublują nawigacji)', () => {
    const { instance, redirects } = msalStub();

    expect(ensureInteractiveReauth(instance, ['User.Read'])).toBe(true);
    expect(ensureInteractiveReauth(instance, ['User.Read'])).toBe(true);

    expect(redirects.length).toBe(1);
    const req = redirects[0] as { scopes: string[]; redirectStartPage?: string };
    expect(req.scopes).toEqual(['User.Read']);
    // Powrót po logowaniu na stronę sprzed wygaśnięcia (nie na /brak-uprawnien).
    expect(req.redirectStartPage).toBe(window.location.href);
  });

  it('połyka błąd MSAL (interaction_in_progress) zamiast wywracać obsługę HTTP', async () => {
    const instance = {
      getActiveAccount: () => null,
      getAllAccounts: () => [],
      acquireTokenRedirect: () => Promise.reject(new Error('interaction_in_progress')),
    } as unknown as IPublicClientApplication;

    expect(ensureInteractiveReauth(instance, ['User.Read'])).toBe(true);
    await Promise.resolve(); // odrzucenie obsłużone wewnętrznie — brak unhandled rejection
  });

  describe('hasFreshIdToken (MSAL waliduje tylko access token — idToken z cache bywa przeterminowany)', () => {
    const result = (idToken: string | null, exp?: number): AuthenticationResult | null =>
      idToken === null
        ? null
        : ({ idToken, idTokenClaims: exp ? { exp } : {} } as unknown as AuthenticationResult);

    it('brak wyniku/tokenu → nieświeży', () => {
      expect(hasFreshIdToken(null)).toBe(false);
      expect(hasFreshIdToken(result(''))).toBe(false);
    });

    it('exp w przeszłości → nieświeży; exp w przyszłości → świeży; brak exp → świeży', () => {
      const now = Math.floor(Date.now() / 1000);
      expect(hasFreshIdToken(result('T', now - 10))).toBe(false);
      // Mniej niż 30 s zapasu = traktuj jak przeterminowany (żądanie zdąży dolecieć z trupem).
      expect(hasFreshIdToken(result('T', now + 10))).toBe(false);
      expect(hasFreshIdToken(result('T', now + 3600))).toBe(true);
      expect(hasFreshIdToken(result('T'))).toBe(true);
    });
  });
});
