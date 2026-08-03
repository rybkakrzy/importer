import { IPublicClientApplication, AuthenticationResult } from '@azure/msal-browser';

/**
 * Jednorazowy wyzwalacz interaktywnej re-autoryzacji (redirect do Entra) dla całej aplikacji.
 *
 * Kontekst (bug 13942097): po wygaśnięciu sesji `acquireTokenSilent` rzuca
 * InteractionRequiredAuthError, a żaden element aplikacji nie inicjował ponownego logowania —
 * żądania szły bez tokenu (albo z przeterminowanym idToken z cache MSAL), backend odpowiadał
 * 401, guardy klasyfikowały to jak brak roli (/brak-uprawnien) i użytkownik był w martwym
 * punkcie do ręcznego wyczyszczenia ciasteczek. MsalGuard nie ratuje: konto wciąż jest
 * w cache, więc nawigacja przechodzi bez odświeżenia tokenów.
 *
 * Flaga modułowa (nie sessionStorage): redirect przeładowuje aplikację, więc świeży moduł
 * zaczyna z czystą flagą; w ramach jednego życia strony wystarczy JEDNA próba — kolejne
 * wywołania (równoległe żądania, retry guardów) nie mogą dublować nawigacji do /authorize.
 */
let reauthStarted = false;

/** Reset dla testów jednostkowych (moduł żyje przez cały przebieg suity). */
export function resetInteractiveReauthForTests(): void {
  reauthStarted = false;
}

/**
 * Uruchamia redirect logowania, jeżeli jeszcze nie wystartował. Zwraca true, gdy nawigacja
 * została zainicjowana (caller powinien wygasić bieżące żądanie — strona i tak odpływa).
 * Błędy MSAL (np. interaction_in_progress, gdy MsalGuard równolegle zaczął własny redirect)
 * są połykane — ktoś inny już prowadzi użytkownika do logowania.
 */
export function ensureInteractiveReauth(
  msalInstance: IPublicClientApplication,
  scopes: string[],
): boolean {
  if (reauthStarted) {
    return true;
  }
  reauthStarted = true;
  const account = msalInstance.getActiveAccount() ?? msalInstance.getAllAccounts()[0] ?? undefined;
  // redirectStartPage: adres SPRZED nawigacji guardów — po ponownym logowaniu użytkownik ma
  // wrócić do dokumentu/strony, na której wygasła sesja, a nie na /brak-uprawnien.
  void msalInstance
    .acquireTokenRedirect({ scopes, account, redirectStartPage: window.location.href })
    .catch(() => {
      // interaction_in_progress / zablokowana nawigacja — nie eskalujemy; następne
      // przeładowanie strony zaczyna od świeżej flagi.
    });
  return true;
}

/**
 * Czy wynik silent acquisition niesie idToken nadający się do wysłania do API.
 * MSAL sprawdza świeżość ACCESS tokenu — cache hit potrafi zwrócić idToken sprzed godzin,
 * który backend (AddMicrosoftIdentityWebApi) odrzuci. 30 s zapasu na przelot żądania.
 */
export function hasFreshIdToken(result: AuthenticationResult | null): boolean {
  if (!result?.idToken) {
    return false;
  }
  const exp = (result.idTokenClaims as { exp?: number } | undefined)?.exp;
  return !exp || exp * 1000 > Date.now() + 30_000;
}
