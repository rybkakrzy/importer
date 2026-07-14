import { describe, it, expect } from 'vitest';
import { HttpErrorResponse } from '@angular/common/http';
import {
  DOCUMENT_CONTENT_EMPTY_CODE,
  EMPTY_DOCUMENT_MESSAGE,
  isEmptyDocumentError,
} from './document-error.util';

describe('isEmptyDocumentError', () => {
  it('rozpoznaje kod bezpośrednio na błędzie domenowym serwisu (np. OpenDocumentError)', () => {
    expect(isEmptyDocumentError({ code: DOCUMENT_CONTENT_EMPTY_CODE })).toBe(true);
  });

  it('rozpoznaje kod w ciele { code, error } odpowiedzi POST /open', () => {
    const err = new HttpErrorResponse({
      status: 400,
      error: { code: DOCUMENT_CONTENT_EMPTY_CODE, error: 'Zawartość dokumentu nie może być pusta.' },
    });
    expect(isEmptyDocumentError(err)).toBe(true);
  });

  it('rozpoznaje kod w rozszerzeniu ProblemDetails walidacji uploadu', () => {
    const err = new HttpErrorResponse({
      status: 400,
      error: {
        title: 'Błąd walidacji',
        errors: { Content: ['Zawartość dokumentu nie może być pusta'] },
        code: DOCUMENT_CONTENT_EMPTY_CODE,
      },
    });
    expect(isEmptyDocumentError(err)).toBe(true);
  });

  it('nie rozpoznaje innych błędów (inny kod, brak kodu, string, null)', () => {
    expect(isEmptyDocumentError({ code: 'WRONG_PASSWORD' })).toBe(false);
    expect(isEmptyDocumentError(new HttpErrorResponse({ status: 404, error: 'not found' }))).toBe(false);
    expect(isEmptyDocumentError(null)).toBe(false);
    expect(isEmptyDocumentError('DOCUMENT_CONTENT_EMPTY')).toBe(false);
  });

  it('komunikat nie sugeruje braku pliku ani ponowienia', () => {
    expect(EMPTY_DOCUMENT_MESSAGE.toLowerCase()).not.toContain('przesłano');
    expect(EMPTY_DOCUMENT_MESSAGE.toLowerCase()).not.toContain('ponownie');
    expect(EMPTY_DOCUMENT_MESSAGE).toContain('treści');
  });
});
