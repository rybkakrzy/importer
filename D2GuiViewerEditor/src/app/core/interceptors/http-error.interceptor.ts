import { HttpInterceptorFn, HttpErrorResponse } from '@angular/common/http';
import { inject } from '@angular/core';
import { Router } from '@angular/router';
import { catchError, throwError } from 'rxjs';
import { MsalService } from '@azure/msal-angular';
import { NotificationService } from '../services/notification.service';
import { ConnectionStatusService } from '../services/connection-status.service';
import { LastHttpErrorService } from '../services/last-http-error.service';
import { documentDefectMessage } from '../errors/document-error.util';
import { MSAL_CUSTOM_CONFIG } from '../config/runtime-config';
import { ensureInteractiveReauth } from '../auth/interactive-reauth';
import { isProtectedApiRequest } from './api-token.interceptor';

/**
 * Interceptor HTTP — centralna obsługa błędów API (RFC 7807 ProblemDetails)
 */
export const httpErrorInterceptor: HttpInterceptorFn = (req, next) => {
  const notificationService = inject(NotificationService);
  const connectionStatus = inject(ConnectionStatusService);
  const lastHttpError = inject(LastHttpErrorService);
  const router = inject(Router);
  const auth = inject(MSAL_CUSTOM_CONFIG);
  const msal = auth.enabled !== false ? inject(MsalService) : null;

  return next(req).pipe(
    catchError((error: HttpErrorResponse) => {
      let errorMessage = 'Wystąpił nieoczekiwany błąd';

      if (error.status === 401 && msal && isProtectedApiRequest(req.url)) {
        // Bug 13942097: 401 z chronionego API = sesja wygasła/token odrzucony. Bez tej gałęzi
        // aplikacja NIGDY nie ponawiała autoryzacji (guardy klasyfikowały 401 jak brak roli →
        // /brak-uprawnien; ratunkiem było czyszczenie ciasteczek). Inicjujemy redirect
        // logowania (raz na życie strony) i wyciszamy toast — strona odpływa do /authorize.
        recordLastError(lastHttpError, error, req, 'Sesja wygasła (401) — ponowna autoryzacja');
        const scopes = auth.apiScopes?.length ? auth.apiScopes : ['openid', 'profile'];
        ensureInteractiveReauth(msal.instance, scopes);
        return throwError(() => error);
      }

      if (error.status === 403) {
        // 403 → dedicated "Brak uprawnień" view (not a toast, and distinct from 401/404/500).
        // Route guards may also redirect here; navigating to the same target is idempotent.
        recordLastError(lastHttpError, error, req, 'Brak uprawnień (403)');
        void router.navigateByUrl('/brak-uprawnien');
        return throwError(() => error);
      }

      if (error.status === 0) {
        connectionStatus.reportApiError();
        errorMessage = 'Nie można połączyć się z serwerem. Sprawdź połączenie sieciowe.';
      } else if (documentDefectMessage(error)) {
        // Znany defekt pliku (pusty / zły format / uszkodzony) — jednolity komunikat po stabilnym
        // kodzie, niezależnie od kształtu odpowiedzi ({code,error} z /open lub ProblemDetails).
        errorMessage = documentDefectMessage(error)!;
      } else if (error.error) {
        // ProblemDetails (RFC 7807) from backend ExceptionHandlingMiddleware
        if (error.error.detail) {
          errorMessage = error.error.detail;
        } else if (error.error.errors) {
          // ValidationProblemDetails — lista błędów walidacji
          const validationErrors = Object.values(error.error.errors).flat();
          errorMessage = validationErrors.join(', ');
        } else if (error.error.error) {
          // Legacy error format
          errorMessage = error.error.error;
        } else if (error.error.message) {
          errorMessage = error.error.message;
        } else if (typeof error.error === 'string') {
          errorMessage = error.error;
        }
      } else {
        errorMessage = error.message || errorMessage;
      }

      // Loguj szczegóły do konsoli w trybie deweloperskim
      console.error(`[HTTP ${error.status}] ${req.method} ${req.url}:`, error);

      // Zapamiętaj ostatni błąd na potrzeby diagnostyki „Zgłoś problem".
      recordLastError(lastHttpError, error, req, errorMessage);

      notificationService.error(errorMessage);

      return throwError(() => error);
    })
  );
};

/** Zapisuje migawkę błędu do serwisu diagnostycznego (wspólne dla ścieżki 403 i pozostałych). */
function recordLastError(
  service: LastHttpErrorService,
  error: HttpErrorResponse,
  req: { method: string; url: string },
  detail: string,
): void {
  service.record({
    status: error.status,
    method: req.method,
    url: req.url,
    detail,
    at: new Date().toLocaleString('pl-PL'),
  });
}
