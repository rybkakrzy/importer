import { TestBed } from '@angular/core/testing';
import { HttpClient } from '@angular/common/http';
import { of } from 'rxjs';
import { ResourceAccessService } from './resource-access.service';
import { MSAL_CUSTOM_CONFIG, DEFAULT_AUTH_CONFIG, AppAuthConfig } from '../config/runtime-config';

function setup(resources: string[], configOverride: Partial<AppAuthConfig> = {}) {
  const get = vi.fn().mockReturnValue(of(resources));
  TestBed.resetTestingModule();
  TestBed.configureTestingModule({
    providers: [
      { provide: HttpClient, useValue: { get } },
      { provide: MSAL_CUSTOM_CONFIG, useValue: { ...DEFAULT_AUTH_CONFIG, ...configOverride } },
    ],
  });
  return { svc: TestBed.inject(ResourceAccessService), get };
}

describe('ResourceAccessService', () => {
  it('hasAccessToResource: true when the resource is in the backend list', async () => {
    const { svc } = setup(['editor', 'viewer']);
    expect(await firstValue(svc.hasAccessToResource('editor'))).toBe(true);
    expect(await firstValue(svc.hasAccessToResource('admin'))).toBe(false);
  });

  it('caches the resource list (one HTTP call across multiple checks)', async () => {
    const { svc, get } = setup(['editor']);
    await firstValue(svc.hasAccessToResource('editor'));
    await firstValue(svc.hasAccessToResource('viewer'));
    await firstValue(svc.getResources());
    expect(get).toHaveBeenCalledTimes(1);
  });

  it('dev bypass (auth disabled): always allows without calling the backend', async () => {
    const { svc, get } = setup([], { enabled: false });
    expect(await firstValue(svc.hasAccessToResource('admin'))).toBe(true);
    expect(get).not.toHaveBeenCalled();
  });
});

function firstValue<T>(obs: { subscribe: (o: { next: (v: T) => void }) => void }): Promise<T> {
  return new Promise((resolve) => obs.subscribe({ next: resolve }));
}
