import { TestBed } from '@angular/core/testing';
import { HttpClient } from '@angular/common/http';
import { Router } from '@angular/router';
import { of } from 'rxjs';
import { homeRedirectGuard } from './home-redirect.guard';
import { ResourceAccessService } from '../core/services/resource-access.service';
import { MSAL_CUSTOM_CONFIG, DEFAULT_AUTH_CONFIG } from '../core/config/runtime-config';

function run(resources: string[], authConfig: Record<string, unknown> = { ...DEFAULT_AUTH_CONFIG }) {
  TestBed.resetTestingModule();
  const createUrlTree = vi.fn().mockImplementation((commands: string[]) => ({ __tree: commands }));
  const get = vi.fn().mockReturnValue(of(resources));
  TestBed.configureTestingModule({
    providers: [
      ResourceAccessService,
      { provide: HttpClient, useValue: { get } },
      { provide: MSAL_CUSTOM_CONFIG, useValue: authConfig },
      { provide: Router, useValue: { createUrlTree } },
    ],
  });
  const result = TestBed.runInInjectionContext(() => homeRedirectGuard({} as never, {} as never));
  return { result, createUrlTree, get };
}

function firstValue<T>(obs: unknown): Promise<T> {
  return new Promise((resolve) =>
    (obs as { subscribe: (o: { next: (v: T) => void }) => void }).subscribe({ next: resolve }),
  );
}

describe('homeRedirectGuard', () => {
  it('Operator (has "dashboard") → shows the dashboard', async () => {
    const { result } = run(['dashboard', 'editor', 'viewer']);
    expect(await firstValue(result)).toBe(true);
  });

  it('Administrator (has "admin", no "dashboard") → redirects to /admin', async () => {
    const { result, createUrlTree } = run(['admin']);
    expect(await firstValue(result)).toEqual({ __tree: ['/admin'] });
    expect(createUrlTree).toHaveBeenCalledWith(['/admin']);
  });

  it('no application role → redirects to /brak-uprawnien (dashboard never shown)', async () => {
    const { result, createUrlTree } = run([]);
    expect(await firstValue(result)).toEqual({ __tree: ['/brak-uprawnien'] });
    expect(createUrlTree).toHaveBeenCalledWith(['/brak-uprawnien']);
  });

  it('dev bypass (auth disabled) → shows the dashboard without calling the API', () => {
    const { result, get } = run([], { ...DEFAULT_AUTH_CONFIG, enabled: false });
    expect(result).toBe(true);
    expect(get).not.toHaveBeenCalled();
  });
});
