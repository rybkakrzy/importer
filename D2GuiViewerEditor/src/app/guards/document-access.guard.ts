import { inject } from '@angular/core';
import { CanActivateFn, Router } from '@angular/router';
import { HttpErrorResponse } from '@angular/common/http';
import { catchError, map, of, retry, throwError, timer } from 'rxjs';
import { DocumentStorageService } from '../services/document-storage.service';
import { isTransient } from './resource.guard';

/**
 * Blocks navigation into a document view (editor / PDF viewer) before the component
 * is created, when the backend denies access (403). Backend remains the source of truth;
 * this guard only prevents the editor/viewer from initializing and loading resources.
 *
 * The metadata endpoint doubles as the access gate (GUI calls it first anyway), so no
 * dedicated check endpoint is needed. Reusable for both DOCX editor and PDF viewer routes.
 *
 * On a cold deep-link the first call can race MSAL settling; transient errors (401/0/5xx) are
 * retried like resourceGuard so the race is not confused with a decision. A definitive 403 fails
 * closed to /brak-uprawnien; a 404 (or exhausted transient retries) lets the target view render its
 * own not-found/error state — the underlying endpoints stay backend-protected regardless.
 */
export const documentAccessGuard: CanActivateFn = (route) => {
  const router = inject(Router);
  const storage = inject(DocumentStorageService);

  const masterId = route.queryParamMap.get('masterId');
  // No document context (e.g. new/blank editor, upload flow) → nothing to guard.
  if (!masterId) {
    return of(true);
  }

  return storage.getDocumentMetadata(masterId).pipe(
    map(() => true),
    retry({
      count: 3,
      delay: (error) => (isTransient(error) ? timer(300) : throwError(() => error)),
    }),
    catchError((err: HttpErrorResponse) =>
      err.status === 403 ? of(router.createUrlTree(['/brak-uprawnien'])) : of(true),
    ),
  );
};
