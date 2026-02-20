import { ErrorHandler, Injectable, inject, NgZone } from '@angular/core';
import { NotificationService } from '../services/notification.service';

/**
 * Globalny handler błędów Angular — przechwytuje nieobsłużone wyjątki
 */
@Injectable()
export class GlobalErrorHandler implements ErrorHandler {
  private notificationService = inject(NotificationService);
  private zone = inject(NgZone);

  handleError(error: unknown): void {
    // Loguj do konsoli
    console.error('[GlobalErrorHandler]', error);

    // Pokaż powiadomienie w strefie Angular
    this.zone.run(() => {
      const message = error instanceof Error
        ? error.message
        : 'Wystąpił nieoczekiwany błąd aplikacji';

      this.notificationService.error(message);
    });
  }
}
