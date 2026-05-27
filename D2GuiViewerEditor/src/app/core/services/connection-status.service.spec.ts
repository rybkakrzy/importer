import { TestBed } from '@angular/core/testing';
import { provideHttpClient } from '@angular/common/http';
import {
  provideHttpClientTesting,
  HttpTestingController,
} from '@angular/common/http/testing';
import { ConnectionStatusService, HealthResponse } from './connection-status.service';

/**
 * Guards the environment/health bootstrap behaviour: the first health request
 * must fire immediately on service construction (no waiting for the 30 s poll),
 * the poll must not double-fire at t=0, and timers must be cleaned up on destroy.
 */
describe('ConnectionStatusService — health bootstrap', () => {
  let httpMock: HttpTestingController;

  const health: HealthResponse = {
    status: 'Healthy',
    environment: 'DEV',
    buildNumber: '1.2.3',
    buildDate: '2026-05-28 10:00',
    timestamp: '2026-05-28T10:00:00Z',
  };

  beforeEach(() => {
    vi.useFakeTimers();
    TestBed.configureTestingModule({
      providers: [provideHttpClient(), provideHttpClientTesting()],
    });
  });

  afterEach(() => {
    httpMock.verify();
    vi.useRealTimers();
  });

  function create(): ConnectionStatusService {
    const service = TestBed.inject(ConnectionStatusService);
    httpMock = TestBed.inject(HttpTestingController);
    return service;
  }

  it('fires the first health request immediately on construction (no 30 s wait)', () => {
    const service = create();

    const req = httpMock.expectOne((r) => r.url.endsWith('/health'));
    req.flush(health);

    expect(service.apiEnvironment()).toBe('DEV');
    expect(service.apiBuildNumber()).toBe('1.2.3');
  });

  it('does not issue a duplicate health request at t=0', () => {
    create();

    const reqs = httpMock.match((r) => r.url.endsWith('/health'));
    expect(reqs.length).toBe(1);
    reqs[0].flush(health);
  });

  it('polls again after 30 s (single recurring poll)', () => {
    create();

    httpMock.expectOne((r) => r.url.endsWith('/health')).flush(health);

    vi.advanceTimersByTime(30_000);

    const reqs = httpMock.match((r) => r.url.endsWith('/health'));
    expect(reqs.length).toBe(1);
    reqs[0].flush(health);
  });

  it('handles an API error on the first request without throwing', () => {
    const service = create();

    const req = httpMock.expectOne((r) => r.url.endsWith('/health'));
    req.error(new ProgressEvent('error'), { status: 0 });

    expect(service.isOffline()).toBe(true);
  });

  it('clears the polling interval on destroy', () => {
    const service = create();
    httpMock.expectOne((r) => r.url.endsWith('/health')).flush(health);

    service.ngOnDestroy();
    vi.advanceTimersByTime(60_000);

    httpMock.expectNone((r) => r.url.endsWith('/health'));
  });
});
