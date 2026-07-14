import { HttpErrorResponse } from '@angular/common/http';

/**
 * Stabilny kod maszynowy zwracany przez backend dla dokumentu bez treści.
 * Występuje w DWÓCH kształtach odpowiedzi:
 *  - POST /open → `{ code, error }`,
 *  - walidacja uploadu (ProblemDetails) → rozszerzenie `code` obok `errors`.
 * GUI rozpoznaje przypadek po TYM kodzie, nigdy po treści komunikatu backendu.
 */
export const DOCUMENT_CONTENT_EMPTY_CODE = 'DOCUMENT_CONTENT_EMPTY';

/**
 * Jednolity, zlokalizowany komunikat dla pustego dokumentu — JEDNO źródło prawdy dla wszystkich
 * ścieżek (strona startowa, edytor „Plik → Otwórz", centralny interceptor). Opisuje przyczynę,
 * nie sugeruje braku pliku ani ponowienia i nie ujawnia szczegółów technicznych backendu.
 */
export const EMPTY_DOCUMENT_MESSAGE =
  'Nie można otworzyć dokumentu, ponieważ nie zawiera on treści.';

/**
 * Rozpoznaje błąd „dokument bez treści" niezależnie od ścieżki i opakowania:
 *  - `HttpErrorResponse` (kod w `error.code`),
 *  - błąd domenowy serwisu z polem `code` (np. `OpenDocumentError`).
 */
export function isEmptyDocumentError(err: unknown): boolean {
  if (!err || typeof err !== 'object') return false;

  const directCode = (err as { code?: unknown }).code;
  if (directCode === DOCUMENT_CONTENT_EMPTY_CODE) return true;

  const body = (err as HttpErrorResponse).error;
  return !!body && typeof body === 'object'
    && (body as { code?: unknown }).code === DOCUMENT_CONTENT_EMPTY_CODE;
}
