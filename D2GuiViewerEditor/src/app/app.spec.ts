import { TestBed } from '@angular/core/testing';
import { signal } from '@angular/core';
import { provideRouter } from '@angular/router';
import { App } from './app';
import { BuildInfoService } from './core/services/build-info.service';
import { ConnectionStatusService } from './core/services/connection-status.service';

describe('App', () => {
  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [App],
      providers: [
        provideRouter([]),
        // Stub the health-check chain so the shell renders without real HTTP.
        { provide: BuildInfoService, useValue: { environment: signal('DEV') } },
        { provide: ConnectionStatusService, useValue: { isOffline: signal(false), dismiss: () => {} } },
      ],
    }).compileComponents();
  });

  it('should create the app', () => {
    const fixture = TestBed.createComponent(App);
    expect(fixture.componentInstance).toBeTruthy();
  });

  it('mounts the global banners container at the top of the shell', () => {
    const fixture = TestBed.createComponent(App);
    fixture.detectChanges();
    const compiled = fixture.nativeElement as HTMLElement;
    expect(compiled.querySelector('d2-global-banners')).not.toBeNull();
  });
});
