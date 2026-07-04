import { inject } from '@angular/core';
import { HttpErrorResponse } from '@angular/common/http';
import { ActivatedRouteSnapshot, CanActivateFn, Router } from '@angular/router';
import { catchError, map, of, retry, throwError, timer } from 'rxjs';
import { ResourceAccessService } from '../core/services/resource-access.service';

/**
 * UX guard: gates a route by its resource name (the route's path segment, e.g. "editor", "viewer",
 * "admin") against the backend resource list (GET /api/identity/resources). Backend is the source of
 * truth. Dev bypass (auth disabled) passes through via ResourceAccessService.
 *
 * Loading vs denied: right after a redirect/reload MSAL may still be settling, so the first resource
 * fetch can go out before the token is ready (transient 401 / network 0 / 5xx). Those are retried a
 * few times rather than being mistaken for a denial. A definitive 403 (Forbidden) — or exhausted
 * retries — fails closed to /brak-uprawnien.
 *
 * Used for both document routes and the admin module, replacing per-role frontend guards: the
 * frontend no longer needs to know role names — the backend maps roles → resources.
 */
export const resourceGuard: CanActivateFn = (route: ActivatedRouteSnapshot) => {
  const router = inject(Router);
  // Explicit resource name (route data) wins over the path segment — lets routes with empty/ambiguous
  // paths (e.g. maintenance pages) declare which resource they belong to.
  const resource =
    (route.data?.['resource'] as string | undefined) ?? route.routeConfig?.path ?? route.url[0]?.path ?? '';
  const denied = () => router.createUrlTree(['/brak-uprawnien']);

  return inject(ResourceAccessService).hasAccessToResource(resource).pipe(
    map((allowed) => (allowed ? true : denied())),
    retry({
      count: 3,
      delay: (error) => (isTransient(error) ? timer(300) : throwError(() => error)),
    }),
    catchError(() => of(denied())),
  );
};

/** A transient (auth-not-ready / network) error worth retrying — as opposed to a definitive 403. */
export function isTransient(error: unknown): boolean {
  if (!(error instanceof HttpErrorResponse)) {
    return false;
  }
  return error.status === 0 || error.status === 401 || error.status >= 500;
}
