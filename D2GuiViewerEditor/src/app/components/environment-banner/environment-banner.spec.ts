import { TestBed, ComponentFixture } from '@angular/core/testing';
import { signal, WritableSignal } from '@angular/core';
import { EnvironmentBannerComponent } from './environment-banner';
import { BuildInfoService } from '../../core/services/build-info.service';

describe('EnvironmentBannerComponent', () => {
  let fixture: ComponentFixture<EnvironmentBannerComponent>;
  let env: WritableSignal<string>;

  beforeEach(async () => {
    env = signal<string>('DEV');
    await TestBed.configureTestingModule({
      imports: [EnvironmentBannerComponent],
      providers: [
        // Stub the single source of truth so the matrix is driven without HTTP.
        { provide: BuildInfoService, useValue: { environment: env } },
      ],
    }).compileComponents();
    fixture = TestBed.createComponent(EnvironmentBannerComponent);
  });

  function banner(): HTMLElement | null {
    return fixture.nativeElement.querySelector('.env-banner');
  }

  function renderWith(envName: string): HTMLElement | null {
    env.set(envName);
    fixture.detectChanges();
    return banner();
  }

  it('renders "Wersja Local" with the blue (local) variant', () => {
    const el = renderWith('Local');
    expect(el?.textContent?.trim()).toBe('Wersja Local');
    expect(el?.classList).toContain('env-banner--local');
  });

  it('renders "Wersja DEV" with the green (dev) variant', () => {
    const el = renderWith('DEV');
    expect(el?.textContent?.trim()).toBe('Wersja DEV');
    expect(el?.classList).toContain('env-banner--dev');
  });

  it('renders "Wersja UAT" with the light-yellow (test) variant', () => {
    const el = renderWith('UAT');
    expect(el?.textContent?.trim()).toBe('Wersja UAT');
    expect(el?.classList).toContain('env-banner--test');
  });

  it('renders "Wersja TST" with the light-yellow (test) variant', () => {
    const el = renderWith('TST');
    expect(el?.textContent?.trim()).toBe('Wersja TST');
    expect(el?.classList).toContain('env-banner--test');
  });

  it('renders "Wersja PRE" with the purple (pre) variant', () => {
    const el = renderWith('PRE');
    expect(el?.textContent?.trim()).toBe('Wersja PRE');
    expect(el?.classList).toContain('env-banner--pre');
  });

  it('hides the banner entirely on production', () => {
    expect(renderWith('PRD')).toBeNull();
    expect(renderWith('PROD')).toBeNull();
  });

  it('shows a neutral fallback for an unknown environment', () => {
    const el = renderWith('Sandbox');
    expect(el?.textContent?.trim()).toBe('Wersja Sandbox');
    expect(el?.classList).toContain('env-banner--unknown');
  });

  it('exposes role="status" for assistive technologies', () => {
    const el = renderWith('DEV');
    expect(el?.getAttribute('role')).toBe('status');
  });

  it('reacts to environment changes (e.g. once the API health-check resolves)', () => {
    expect(renderWith('PRD')).toBeNull();
    const el = renderWith('UAT');
    expect(el?.classList).toContain('env-banner--test');
  });
});
