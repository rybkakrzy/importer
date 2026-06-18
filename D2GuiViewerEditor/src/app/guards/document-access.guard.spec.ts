import { TestBed } from '@angular/core/testing';
import { provideRouter, Router, UrlTree, ActivatedRouteSnapshot } from '@angular/router';
import { HttpErrorResponse } from '@angular/common/http';
import { of, throwError, firstValueFrom, isObservable, Observable } from 'rxjs';
import { documentAccessGuard } from './document-access.guard';
import { DocumentStorageService } from '../services/document-storage.service';

function snapshotWith(masterId: string | null): ActivatedRouteSnapshot {
  return {
    queryParamMap: { get: (k: string) => (k === 'masterId' ? masterId : null) },
  } as unknown as ActivatedRouteSnapshot;
}

function runGuard(route: ActivatedRouteSnapshot): boolean | UrlTree | Observable<boolean | UrlTree> {
  return TestBed.runInInjectionContext(() =>
    documentAccessGuard(route, {} as never)
  ) as boolean | UrlTree | Observable<boolean | UrlTree>;
}

async function resolve(value: boolean | UrlTree | Observable<boolean | UrlTree>) {
  return isObservable(value) ? firstValueFrom(value) : value;
}

describe('documentAccessGuard', () => {
  let storage: { getDocumentMetadata: ReturnType<typeof vi.fn> };

  beforeEach(() => {
    storage = { getDocumentMetadata: vi.fn() };
    TestBed.configureTestingModule({
      providers: [
        provideRouter([]),
        { provide: DocumentStorageService, useValue: storage },
      ],
    });
  });

  it('allows navigation when there is no masterId (new/blank editor)', async () => {
    const result = await resolve(runGuard(snapshotWith(null)));
    expect(result).toBe(true);
    expect(storage.getDocumentMetadata).not.toHaveBeenCalled();
  });

  it('allows navigation when access is granted (200)', async () => {
    storage.getDocumentMetadata.mockReturnValue(of({ masterId: 'm', mimeType: 'application/pdf' }));
    const result = await resolve(runGuard(snapshotWith('m')));
    expect(result).toBe(true);
  });

  it('redirects to /brak-uprawnien on 403', async () => {
    storage.getDocumentMetadata.mockReturnValue(
      throwError(() => new HttpErrorResponse({ status: 403 }))
    );
    const router = TestBed.inject(Router);
    const result = await resolve(runGuard(snapshotWith('m')));
    expect(result).toBeInstanceOf(UrlTree);
    expect((result as UrlTree).toString()).toBe(router.createUrlTree(['/brak-uprawnien']).toString());
  });

  it('does not block on 404 (target view handles not-found)', async () => {
    storage.getDocumentMetadata.mockReturnValue(
      throwError(() => new HttpErrorResponse({ status: 404 }))
    );
    const result = await resolve(runGuard(snapshotWith('m')));
    expect(result).toBe(true);
  });
});
