import { HttpInterceptorFn, HttpErrorResponse } from '@angular/common/http';
import { inject } from '@angular/core';
import { catchError, throwError } from 'rxjs';
import { NotificationService } from '../services/notification.service';

/**
 * Interceptor HTTP — centralna obsługa błędów API (RFC 7807 ProblemDetails)
 */
export const httpErrorInterceptor: HttpInterceptorFn = (req, next) => {
  const notificationService = inject(NotificationService);

  return next(req).pipe(
    catchError((error: HttpErrorResponse) => {
      let errorMessage = 'Wystąpił nieoczekiwany błąd';

      if (error.status === 0) {
        errorMessage = 'Nie można połączyć się z serwerem. Sprawdź połączenie sieciowe.';
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

      notificationService.error(errorMessage);

      return throwError(() => error);
    })
  );
};
