import { inject } from '@angular/core';
import { CanActivateFn, Router } from '@angular/router';
import { HttpErrorResponse } from '@angular/common/http';
import { catchError, map, of } from 'rxjs';
import { DocumentStorageService } from '../services/document-storage.service';

/**
 * Blocks navigation into a document view (editor / PDF viewer) before the component
 * is created, when the backend denies access (403). Backend remains the source of truth;
 * this guard only prevents the editor/viewer from initializing and loading resources.
 *
 * The metadata endpoint doubles as the access gate (GUI calls it first anyway), so no
 * dedicated check endpoint is needed. Reusable for both DOCX editor and PDF viewer routes.
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
    catchError((err: HttpErrorResponse) => {
      if (err.status === 403) {
        return of(router.createUrlTree(['/access-denied']));
      }
      // 404 / other errors: let the target view render its own not-found/error state.
      return of(true);
    })
  );
};
