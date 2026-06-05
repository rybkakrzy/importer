import {
  ApplicationConfig,
  ErrorHandler,
  inject,
  provideAppInitializer,
  provideBrowserGlobalErrorListeners,
} from '@angular/core';
import { provideRouter } from '@angular/router';
import { provideHttpClient, withInterceptors } from '@angular/common/http';

import { routes } from './app.routes';
import { httpErrorInterceptor } from './core/interceptors/http-error.interceptor';
import { GlobalErrorHandler } from './core/error-handling/global-error-handler';
import { ConnectionStatusService } from './core/services/connection-status.service';

export const appConfig: ApplicationConfig = {
  providers: [
    provideBrowserGlobalErrorListeners(),
    provideRouter(routes),
    provideHttpClient(withInterceptors([httpErrorInterceptor])),
    { provide: ErrorHandler, useClass: GlobalErrorHandler },
    // Eagerly start the connection/health monitor at bootstrap so the first
    // environment health-check fires immediately, before any component renders
    // and without waiting for the 30 s poll. We call checkNow() explicitly (rather
    // than relying on the constructor side-effect) so the intent is clear and
    // regression-proof; the service's in-flight guard keeps it a single request.
    // Non-blocking: we don't return/await the response.
    provideAppInitializer(() => {
      inject(ConnectionStatusService).checkNow();
    }),
  ]
};
