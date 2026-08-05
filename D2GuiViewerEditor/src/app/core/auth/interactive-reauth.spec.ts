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

  it('bezpiecznik antypętlowy: powyżej 3 prób w oknie czasowym NIE nawiguję i zwracam false', async () => {
    // Pętla z DEV/UAT: nieudana wymiana kodu po powrocie z Entra → każdy przeładunek strony
    // odpalał kolejny redirect, aż Microsoft ubił sesję throttlingiem. Licznik w sessionStorage
    // przeżywa przeładowania; flaga modułowa symulowana resetem (nowe „życie strony").
    const { instance, redirects } = msalStub();

    for (let i = 0; i < 3; i++) {
      expect(ensureInteractiveReauth(instance, ['User.Read'])).toBe(true);
      resetInteractiveReauthForTestsKeepingAttempts();
    }
    expect(ensureInteractiveReauth(instance, ['User.Read'])).toBe(false);

    expect(redirects.length).toBe(3);
  });

  /** Reset flagi modułowej BEZ czyszczenia licznika prób (symulacja przeładowania strony). */
  function resetInteractiveReauthForTestsKeepingAttempts(): void {
    const attempts = sessionStorage.getItem('d2.reauth.attempts');
    resetInteractiveReauthForTests();
    if (attempts !== null) {
      sessionStorage.setItem('d2.reauth.attempts', attempts);
    }
  }

  it('po nieudanym redirect (failed_to_redirect) następny wyzwalacz ponawia próbę', async () => {
    // Bug z DEV: nawigacja do /authorize nie wyszła (timeout MSAL), a jednorazowa flaga
    // blokowała każdą kolejną próbę do ręcznego F5 — martwy punkt z serii 401.
    let attempts = 0;
    const instance = {
      getActiveAccount: () => ({ homeAccountId: 'acc' }),
      getAllAccounts: () => [{ homeAccountId: 'acc' }],
      acquireTokenRedirect: () => {
        attempts++;
        return Promise.reject(new Error('failed_to_redirect'));
      },
    } as unknown as IPublicClientApplication;

    ensureInteractiveReauth(instance, ['User.Read']);
    await Promise.resolve(); // catch zdejmuje flagę
    await Promise.resolve();
    ensureInteractiveReauth(instance, ['User.Read']);

    expect(attempts).toBe(2);
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
