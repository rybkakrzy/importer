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
 * Kody i komunikaty defektów PLIKU przy otwieraniu (bug 13625398) — teksty uzgodnione z QA,
 * identyczne na stronie startowej, w edytorze („Plik → Otwórz") i w centralnym interceptorze.
 */
export const DOCUMENT_FORMAT_INVALID_CODE = 'DOCUMENT_FORMAT_INVALID';
export const DOCUMENT_CORRUPTED_CODE = 'DOCUMENT_CORRUPTED';

export const INVALID_FORMAT_MESSAGE = 'Nieprawidłowy format dokumentu.';
export const CORRUPTED_DOCUMENT_MESSAGE = 'Dokument jest uszkodzony.';

const MESSAGE_BY_CODE: Record<string, string> = {
  [DOCUMENT_CONTENT_EMPTY_CODE]: EMPTY_DOCUMENT_MESSAGE,
  [DOCUMENT_FORMAT_INVALID_CODE]: INVALID_FORMAT_MESSAGE,
  [DOCUMENT_CORRUPTED_CODE]: CORRUPTED_DOCUMENT_MESSAGE,
};

/** Kod maszynowy z błędu niezależnie od opakowania (HttpErrorResponse / błąd domenowy). */
function documentErrorCode(err: unknown): string | null {
  if (!err || typeof err !== 'object') return null;

  const directCode = (err as { code?: unknown }).code;
  if (typeof directCode === 'string') return directCode;

  const body = (err as HttpErrorResponse).error;
  if (body && typeof body === 'object' && typeof (body as { code?: unknown }).code === 'string') {
    return (body as { code: string }).code;
  }
  return null;
}

/**
 * Jednolity komunikat dla znanych defektów dokumentu (pusty / zły format / uszkodzony)
 * albo null, gdy błąd nie jest jednym z tych przypadków (wtedy ścieżki pokazują swój
 * dotychczasowy komunikat). GUI rozpoznaje przypadek po KODZIE, nigdy po treści.
 */
export function documentDefectMessage(err: unknown): string | null {
  const code = documentErrorCode(err);
  return code ? MESSAGE_BY_CODE[code] ?? null : null;
}

/**
 * Rozpoznaje błąd „dokument bez treści" niezależnie od ścieżki i opakowania:
 *  - `HttpErrorResponse` (kod w `error.code`),
 *  - błąd domenowy serwisu z polem `code` (np. `OpenDocumentError`).
 */
export function isEmptyDocumentError(err: unknown): boolean {
  return documentErrorCode(err) === DOCUMENT_CONTENT_EMPTY_CODE;
}
