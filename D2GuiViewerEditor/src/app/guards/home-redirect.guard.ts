import { inject } from '@angular/core';
import { CanActivateFn, Router } from '@angular/router';
import { catchError, map, of, retry, throwError, timer } from 'rxjs';
import { ResourceAccessService } from '../core/services/resource-access.service';
import { MSAL_CUSTOM_CONFIG, isAuthDisabled } from '../core/config/runtime-config';
import { isTransient } from './resource.guard';

/**
 * Landing guard for the root route. Access areas are disjoint (backend ResourcesProvider):
 *   - Operator has "dashboard" → show the dashboard.
 *   - Administrator has "admin" but not "dashboard" → redirect to the admin module.
 *   - no application role → /brak-uprawnien (the dashboard is never shown without access).
 *
 * Backend is the source of truth; this only routes the user to the area they may open. Transient
 * auth-not-ready / network errors are retried (mirrors resourceGuard); a definitive failure fails
 * closed to /brak-uprawnien.
 */
export const homeRedirectGuard: CanActivateFn = () => {
  const router = inject(Router);

  if (isAuthDisabled(inject(MSAL_CUSTOM_CONFIG))) {
    return true; // dev bypass → show dashboard
  }

  return inject(ResourceAccessService).getResources().pipe(
    map((resources) => {
      if (resources.includes('dashboard')) {
        return true;
      }
      if (resources.includes('admin')) {
        return router.createUrlTree(['/admin']);
      }
      return router.createUrlTree(['/brak-uprawnien']);
    }),
    retry({
      count: 3,
      delay: (error) => (isTransient(error) ? timer(300) : throwError(() => error)),
    }),
    catchError(() => of(router.createUrlTree(['/brak-uprawnien']))),
  );
};
