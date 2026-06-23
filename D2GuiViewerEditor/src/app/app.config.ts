import {
  ApplicationConfig,
  ErrorHandler,
  inject,
  provideAppInitializer,
  provideBrowserGlobalErrorListeners,
} from '@angular/core';
import { provideRouter } from '@angular/router';
import {
  provideHttpClient,
  withInterceptors,
} from '@angular/common/http';
import {
  MSAL_INSTANCE,
  MSAL_GUARD_CONFIG,
  MsalService,
  MsalGuard,
  MsalBroadcastService,
} from '@azure/msal-angular';

import { routes } from './app.routes';
import { apiTokenInterceptor } from './core/interceptors/api-token.interceptor';
import { httpErrorInterceptor } from './core/interceptors/http-error.interceptor';
import { GlobalErrorHandler } from './core/error-handling/global-error-handler';
import { ConnectionStatusService } from './core/services/connection-status.service';
import { MSAL_CUSTOM_CONFIG } from './core/config/runtime-config';
import {
  msalInstanceFactory,
  msalGuardConfigFactory,
} from './core/auth/msal.config';

export const appConfig: ApplicationConfig = {
  providers: [
    provideBrowserGlobalErrorListeners(),
    provideRouter(routes),

    // apiToken (functional) attaches the ID token to API calls; httpError handles errors.
    provideHttpClient(withInterceptors([apiTokenInterceptor, httpErrorInterceptor])),

    // MSAL (Entra ID) — auth config comes from MSAL_CUSTOM_CONFIG (Qutas runtime config.json),
    // provided at bootstrap in main.ts with build-time environment fallback.
    { provide: MSAL_INSTANCE, useFactory: msalInstanceFactory, deps: [MSAL_CUSTOM_CONFIG] },
    { provide: MSAL_GUARD_CONFIG, useFactory: msalGuardConfigFactory, deps: [MSAL_CUSTOM_CONFIG] },
    MsalService,
    MsalGuard,
    MsalBroadcastService,

    // MSAL v3+ must be initialized before use. We also set the active account as early as
    // possible so first API calls can acquire a token immediately.
    provideAppInitializer(async () => {
      const msal = inject(MsalService);
      await msal.instance.initialize();

      if (!msal.instance.getActiveAccount()) {
        const firstAccount = msal.instance.getAllAccounts()[0] ?? null;
        if (firstAccount) {
          msal.instance.setActiveAccount(firstAccount);
        }
      }
    }),

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
  ],
};
