import { inject } from '@angular/core';
import { CanActivateFn, Router } from '@angular/router';
import { MsalService } from '@azure/msal-angular';
import { RUNTIME_AUTH_CONFIG, isAuthDisabled } from '../core/config/runtime-config';

/**
 * UX guard for the admin module. Backend enforces RequireAppAdmin (source of truth);
 * this only hides /admin from non-admins and redirects them to the shared access-denied view.
 */
export const appAdminGuard: CanActivateFn = () => {
  const router = inject(Router);
  const msal = inject(MsalService);
  const authConfig = inject(RUNTIME_AUTH_CONFIG);

  // Lokalny dev bypass: brak Entra → admin dostępny (backend i tak ma DevBypass).
  if (isAuthDisabled(authConfig)) {
    return true;
  }

  const account = msal.instance.getActiveAccount() ?? msal.instance.getAllAccounts()[0];
  const roles = (account?.idTokenClaims as { roles?: string[] } | undefined)?.roles ?? [];

  return roles.includes(authConfig.adminRole)
    ? true
    : router.createUrlTree(['/access-denied']);
};
