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
  withInterceptorsFromDi,
  HTTP_INTERCEPTORS,
} from '@angular/common/http';
import {
  MSAL_INSTANCE,
  MSAL_GUARD_CONFIG,
  MSAL_INTERCEPTOR_CONFIG,
  MsalService,
  MsalGuard,
  MsalBroadcastService,
  MsalInterceptor,
} from '@azure/msal-angular';

import { routes } from './app.routes';
import { httpErrorInterceptor } from './core/interceptors/http-error.interceptor';
import { GlobalErrorHandler } from './core/error-handling/global-error-handler';
import { ConnectionStatusService } from './core/services/connection-status.service';
import { MSAL_CUSTOM_CONFIG } from './core/config/runtime-config';
import {
  msalInstanceFactory,
  msalGuardConfigFactory,
  msalInterceptorConfigFactory,
} from './core/auth/msal.config';

export const appConfig: ApplicationConfig = {
  providers: [
    provideBrowserGlobalErrorListeners(),
    provideRouter(routes),

    // MsalInterceptor (DI-based) attaches the access token; httpError (functional) handles errors.
    provideHttpClient(withInterceptorsFromDi(), withInterceptors([httpErrorInterceptor])),
    { provide: HTTP_INTERCEPTORS, useClass: MsalInterceptor, multi: true },

    // MSAL (Entra ID) — auth config comes from MSAL_CUSTOM_CONFIG (Doc2 runtime config.json),
    // provided at bootstrap in main.ts with build-time environment fallback.
    { provide: MSAL_INSTANCE, useFactory: msalInstanceFactory, deps: [MSAL_CUSTOM_CONFIG] },
    { provide: MSAL_GUARD_CONFIG, useFactory: msalGuardConfigFactory, deps: [MSAL_CUSTOM_CONFIG] },
    { provide: MSAL_INTERCEPTOR_CONFIG, useFactory: msalInterceptorConfigFactory, deps: [MSAL_CUSTOM_CONFIG] },
    MsalService,
    MsalGuard,
    MsalBroadcastService,

    // MSAL v3+ must be initialized before use.
    provideAppInitializer(() => inject(MsalService).instance.initialize()),

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
