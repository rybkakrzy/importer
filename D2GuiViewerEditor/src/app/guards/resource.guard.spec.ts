import { TestBed } from '@angular/core/testing';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { Router } from '@angular/router';
import { ActivatedRouteSnapshot } from '@angular/router';
import { of, throwError } from 'rxjs';
import { resourceGuard } from './resource.guard';
import { ResourceAccessService } from '../core/services/resource-access.service';
import { MSAL_CUSTOM_CONFIG, DEFAULT_AUTH_CONFIG } from '../core/config/runtime-config';

const DENIED = { __denied: true };

function run(path: string, httpGet: () => unknown) {
  TestBed.resetTestingModule();
  const createUrlTree = vi.fn().mockReturnValue(DENIED);
  const get = vi.fn().mockImplementation(httpGet);
  TestBed.configureTestingModule({
    providers: [
      ResourceAccessService,
      { provide: HttpClient, useValue: { get } },
      { provide: MSAL_CUSTOM_CONFIG, useValue: { ...DEFAULT_AUTH_CONFIG } },
      { provide: Router, useValue: { createUrlTree } },
    ],
  });
  const route = { routeConfig: { path }, url: [] } as unknown as ActivatedRouteSnapshot;
  const result = TestBed.runInInjectionContext(() =>
    resourceGuard(route, {} as never),
  );
  return { result, get, createUrlTree };
}

function firstValue<T>(obs: unknown): Promise<T> {
  return new Promise((resolve) => (obs as { subscribe: (o: { next: (v: T) => void }) => void }).subscribe({ next: resolve }));
}

describe('resourceGuard', () => {
  it('allows when the route resource is in the backend list', async () => {
    const { result } = run('editor', () => of(['editor', 'viewer']));
    expect(await firstValue(result)).toBe(true);
  });

  it('denies (redirect to /brak-uprawnien) when the resource is absent', async () => {
    const { result, createUrlTree } = run('admin', () => of(['editor', 'viewer']));
    expect(await firstValue(result)).toBe(DENIED);
    expect(createUrlTree).toHaveBeenCalledWith(['/brak-uprawnien']);
  });

  it('denies on a definitive 403 without retrying', async () => {
    const forbidden = new HttpErrorResponse({ status: 403 });
    const { result, get } = run('editor', () => throwError(() => forbidden));
    expect(await firstValue(result)).toBe(DENIED);
    // 403 is a real "no access" — the guard must not retry it as if it were a race.
    expect(get).toHaveBeenCalledTimes(1);
  });
});
