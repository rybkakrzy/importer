import { Injectable, inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable, of, map, tap } from 'rxjs';
import { environment } from '../../../environments/environment';
import { MSAL_CUSTOM_CONFIG, isAuthDisabled } from '../config/runtime-config';

/**
 * Resource-based authorization (Qutas/D2WebCore pattern). The backend maps the signed-in user's app
 * roles to a set of resource names (route names) and exposes them at GET /api/identity/resources.
 * The frontend never inspects role names — it asks the backend "what may I access" and gates routes
 * by resource. Backend = source of truth; this is UX only.
 *
 * The resource list is cached per session (it does not change mid-session); only successful
 * responses are cached, so a transient error is retried on the next navigation.
 */
@Injectable({ providedIn: 'root' })
export class ResourceAccessService {
  private readonly http = inject(HttpClient);
  private readonly authConfig = inject(MSAL_CUSTOM_CONFIG);
  private cached?: string[];

  /** Resource names the current user may access (cached after the first successful fetch). */
  getResources(): Observable<string[]> {
    if (this.cached) {
      return of(this.cached);
    }
    return this.http
      .get<string[]>(`${environment.apiUrl}/identity/resources`)
      .pipe(tap((resources) => (this.cached = resources)));
  }

  /** True if the user may access the named resource. Dev bypass (auth disabled) → always true. */
  hasAccessToResource(name: string): Observable<boolean> {
    if (isAuthDisabled(this.authConfig)) {
      return of(true);
    }
    return this.getResources().pipe(map((resources) => resources.includes(name)));
  }
}
