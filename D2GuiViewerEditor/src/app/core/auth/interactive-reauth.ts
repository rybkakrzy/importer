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

/**
 * Bezpiecznik antypętlowy (zgłoszenie z DEV/UAT): gdy wymiana kodu na tokeny po powrocie
 * z Entra zawodzi (np. zablokowany POST na endpoint /token), każde przeładowanie strony
 * kończyło się 401 → kolejnym redirectem → aż Entra ubijała pętlę throttlingiem.
 * Flaga modułowa nie przeżywa przeładowań, więc licznik prób żyje w sessionStorage
 * (per karta, przeżywa redirecty). Po przekroczeniu limitu w oknie czasowym nie nawigujemy —
 * aplikacja zostaje w widocznym stanie 401/„brak uprawnień" zamiast młócić Entra.
 */
const REAUTH_ATTEMPTS_KEY = 'd2.reauth.attempts';
const REAUTH_WINDOW_MS = 60_000;
const REAUTH_MAX_ATTEMPTS = 3;

function countRecentAttemptAndRecord(): number {
  let attempts: number[] = [];
  try {
    attempts = JSON.parse(sessionStorage.getItem(REAUTH_ATTEMPTS_KEY) ?? '[]');
    if (!Array.isArray(attempts)) attempts = [];
  } catch {
    attempts = [];
  }
  const now = Date.now();
  attempts = attempts.filter((t) => typeof t === 'number' && now - t < REAUTH_WINDOW_MS);
  attempts.push(now);
  try {
    sessionStorage.setItem(REAUTH_ATTEMPTS_KEY, JSON.stringify(attempts));
  } catch {
    // sessionStorage niedostępny — bez licznika (zachowanie sprzed bezpiecznika).
  }
  return attempts.length;
}

/** Reset dla testów jednostkowych (moduł żyje przez cały przebieg suity). */
export function resetInteractiveReauthForTests(): void {
  reauthStarted = false;
  try {
    sessionStorage.removeItem(REAUTH_ATTEMPTS_KEY);
  } catch {
    // jak wyżej — brak sessionStorage nie może wywracać testów.
  }
}

/**
 * Uruchamia redirect logowania, jeżeli jeszcze nie wystartował. Zwraca true, gdy nawigacja
 * została zainicjowana (caller powinien wygasić bieżące żądanie — strona i tak odpływa);
 * false, gdy bezpiecznik antypętlowy zatrzymał kolejne wejście na /authorize — wtedy żądania
 * mają lecieć dalej bez tokenu (widoczny 401/„brak uprawnień" zamiast młócenia Entra).
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
  if (countRecentAttemptAndRecord() > REAUTH_MAX_ATTEMPTS) {
    console.error(
      `[auth] Przerwano pętlę logowania: ${REAUTH_MAX_ATTEMPTS} nieudane powroty z Entra w ` +
        `${REAUTH_WINDOW_MS / 1000} s. Najczęstsza przyczyna: wymiana kodu na tokeny nie dochodzi ` +
        'do skutku (POST na …/oauth2/v2.0/token zablokowany przez CSP/proxy) albo błędny ' +
        'redirectUri w assets/configs/config.json. Sprawdź konsolę/Network po powrocie z #code=.',
    );
    return false;
  }
  reauthStarted = true;
  const account = msalInstance.getActiveAccount() ?? msalInstance.getAllAccounts()[0] ?? undefined;
  // redirectStartPage: adres SPRZED nawigacji guardów — po ponownym logowaniu użytkownik ma
  // wrócić do dokumentu/strony, na której wygasła sesja, a nie na /brak-uprawnien.
  void msalInstance
    .acquireTokenRedirect({ scopes, account, redirectStartPage: window.location.href })
    .catch((error: unknown) => {
      // Nawigacja do /authorize nie doszła do skutku (np. failed_to_redirect: strona nie
      // opuściła aplikacji w limicie MSAL — brak sieci/proxy do IdP; albo
      // interaction_in_progress). Zdejmujemy flagę, żeby NASTĘPNY wyzwalacz (kolejne 401,
      // interakcja użytkownika) ponowił próbę — bez tego użytkownik był w martwym punkcie
      // do ręcznego F5, mimo że przyczyna mogła być przejściowa.
      reauthStarted = false;
      console.warn('[auth] acquireTokenRedirect nie powiódł się — ponowię przy następnym wyzwalaczu:', error);
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
