import { TestBed, ComponentFixture } from '@angular/core/testing';
import { signal, WritableSignal } from '@angular/core';
import { GlobalBannersComponent } from './global-banners';
import { BuildInfoService } from '../../core/services/build-info.service';
import { ConnectionStatusService } from '../../core/services/connection-status.service';

describe('GlobalBannersComponent', () => {
  let fixture: ComponentFixture<GlobalBannersComponent>;
  let env: WritableSignal<string>;
  let offline: WritableSignal<boolean>;

  beforeEach(async () => {
    env = signal<string>('DEV');
    offline = signal<boolean>(false);
    await TestBed.configureTestingModule({
      imports: [GlobalBannersComponent],
      providers: [
        { provide: BuildInfoService, useValue: { environment: env } },
        { provide: ConnectionStatusService, useValue: { isOffline: offline, dismiss: () => {} } },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(GlobalBannersComponent);
    fixture.detectChanges();
  });

  it('renders the environment banner above the offline banner (fixed order)', () => {
    offline.set(true);
    fixture.detectChanges();

    const html: string = fixture.nativeElement.innerHTML;
    const envIndex = html.indexOf('env-banner');
    const offlineIndex = html.indexOf('offline-banner');
    expect(envIndex).toBeGreaterThanOrEqual(0);
    expect(offlineIndex).toBeGreaterThanOrEqual(0);
    expect(envIndex).toBeLessThan(offlineIndex);
  });

  it('shows the offline banner only while the API is unreachable', () => {
    expect(fixture.nativeElement.querySelector('.offline-banner')).toBeNull();
    offline.set(true);
    fixture.detectChanges();
    expect(fixture.nativeElement.querySelector('.offline-banner')).not.toBeNull();
  });

  it('keeps the environment banner independent of the offline state', () => {
    expect(fixture.nativeElement.querySelector('.env-banner')).not.toBeNull();
    offline.set(true);
    fixture.detectChanges();
    expect(fixture.nativeElement.querySelector('.env-banner')).not.toBeNull();
  });

  it('renders neither banner when on production and online', () => {
    env.set('PRD');
    offline.set(false);
    fixture.detectChanges();
    expect(fixture.nativeElement.querySelector('.env-banner')).toBeNull();
    expect(fixture.nativeElement.querySelector('.offline-banner')).toBeNull();
  });
});
