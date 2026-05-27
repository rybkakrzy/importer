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
    // (its constructor already polls every 30 s afterwards). Non-blocking: we
    // don't await the response, just trigger instantiation.
    provideAppInitializer(() => {
      inject(ConnectionStatusService);
    }),
  ]
};
