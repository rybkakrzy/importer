import 'zone.js';
import { bootstrapApplication } from '@angular/platform-browser';
import { appConfig } from './app/app.config';
import { App } from './app/app';
import { AppAuthConfig, RUNTIME_AUTH_CONFIG, mergeAuthConfig } from './app/core/config/runtime-config';

/**
 * Doc2 runtime config: fetch `assets/configs/config.json` BEFORE bootstrap so a single build can
 * be deployed to every environment (clientId/authority/redirect supplied per environment, no
 * rebuild). Missing/invalid file falls back to build-time `environment.auth`.
 */
fetch('assets/configs/config.json')
  .then((response) => (response.ok ? response.json() : {}))
  .catch(() => ({}))
  .then((config: { auth?: Partial<AppAuthConfig> } | null) => {
    const auth = mergeAuthConfig(config?.auth);
    bootstrapApplication(App, {
      ...appConfig,
      providers: [{ provide: RUNTIME_AUTH_CONFIG, useValue: auth }, ...appConfig.providers],
    }).catch((err) => console.error(err));
  });
