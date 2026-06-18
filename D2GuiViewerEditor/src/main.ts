import 'zone.js';
import { bootstrapApplication } from '@angular/platform-browser';
import { MsalRedirectComponent } from '@azure/msal-angular';
import { appConfig } from './app/app.config';
import { App } from './app/app';
import { AppAuthConfig, MSAL_CUSTOM_CONFIG, mergeAuthConfig } from './app/core/config/runtime-config';

/**
 * Qutas runtime config: fetch `assets/configs/config.json` BEFORE bootstrap so a single build can
 * be deployed to every environment (clientId/authority/redirect supplied per environment, no
 * rebuild). config.json is the source of auth values; a missing/invalid file falls back to the
 * structural defaults in runtime-config.ts (DEFAULT_AUTH_CONFIG).
 */
fetch('assets/configs/config.json')
  .then((response) => (response.ok ? response.json() : {}))
  .catch(() => ({}))
  .then((config: { auth?: Partial<AppAuthConfig> } | null) => {
    const auth = mergeAuthConfig(config?.auth);
    bootstrapApplication(App, {
      ...appConfig,
      providers: [{ provide: MSAL_CUSTOM_CONFIG, useValue: auth }, ...appConfig.providers],
    })
      // Bootstrap MsalRedirectComponent into the <app-redirect> host (index.html) so MSAL's
      // dedicated redirect handler runs. No AppModule needed — it shares this app's injector
      // (MsalService) and is the canonical redirect setup for standalone apps.
      .then((appRef) => appRef.bootstrap(MsalRedirectComponent))
      .catch((err) => console.error(err));
  });
