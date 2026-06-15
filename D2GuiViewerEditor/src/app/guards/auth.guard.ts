import { inject } from '@angular/core';
import { CanActivateFn } from '@angular/router';
import { MsalGuard } from '@azure/msal-angular';
import { RUNTIME_AUTH_CONFIG, isAuthDisabled } from '../core/config/runtime-config';

/**
 * Gate logowania całej aplikacji. Gdy auth jest wyłączony (lokalny dev bypass,
 * `config.json` `auth.enabled=false`) — przepuszcza bez MSAL. W przeciwnym razie deleguje do MsalGuard.
 */
export const authGuard: CanActivateFn = (route, state) => {
  if (isAuthDisabled(inject(RUNTIME_AUTH_CONFIG))) {
    return true;
  }
  return inject(MsalGuard).canActivate(route, state);
};
