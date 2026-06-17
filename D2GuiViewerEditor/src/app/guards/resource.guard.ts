import { inject } from '@angular/core';
import { ActivatedRouteSnapshot, CanActivateFn, Router } from '@angular/router';
import { catchError, map, of } from 'rxjs';
import { ResourceAccessService } from '../core/services/resource-access.service';

/**
 * UX guard: gates a route by its resource name (the route's path segment, e.g. "editor", "viewer",
 * "admin") against the backend resource list (GET /api/identity/resources). Backend is the source of
 * truth — on denial or error we fail closed and redirect to /access-denied. Dev bypass (auth
 * disabled) passes through via ResourceAccessService.
 *
 * Used for both document routes and the admin module, replacing per-role frontend guards: the
 * frontend no longer needs to know role names — the backend maps roles → resources.
 */
export const resourceGuard: CanActivateFn = (route: ActivatedRouteSnapshot) => {
  const router = inject(Router);
  const resource = route.routeConfig?.path ?? route.url[0]?.path ?? '';
  return inject(ResourceAccessService).hasAccessToResource(resource).pipe(
    map((allowed) => (allowed ? true : router.createUrlTree(['/access-denied']))),
    catchError(() => of(router.createUrlTree(['/access-denied']))),
  );
};
