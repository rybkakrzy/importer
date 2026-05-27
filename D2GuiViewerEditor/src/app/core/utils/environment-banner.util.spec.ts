import { resolveEnvironmentBanner } from './environment-banner.util';

describe('resolveEnvironmentBanner', () => {
  it('maps Local to a blue variant with "Wersja Local"', () => {
    expect(resolveEnvironmentBanner('Local')).toEqual({ label: 'Wersja Local', variant: 'local' });
  });

  it('maps DEV to a green variant with "Wersja DEV"', () => {
    expect(resolveEnvironmentBanner('DEV')).toEqual({ label: 'Wersja DEV', variant: 'dev' });
  });

  it('maps UAT to the light-yellow "test" variant with "Wersja UAT"', () => {
    expect(resolveEnvironmentBanner('UAT')).toEqual({ label: 'Wersja UAT', variant: 'test' });
  });

  it('maps TST to the light-yellow "test" variant with "Wersja TST"', () => {
    expect(resolveEnvironmentBanner('TST')).toEqual({ label: 'Wersja TST', variant: 'test' });
  });

  it('maps PRE to the purple variant with "Wersja PRE"', () => {
    expect(resolveEnvironmentBanner('PRE')).toEqual({ label: 'Wersja PRE', variant: 'pre' });
  });

  it('hides the banner on production (PROD / PRD)', () => {
    expect(resolveEnvironmentBanner('PROD')).toBeNull();
    expect(resolveEnvironmentBanner('PRD')).toBeNull();
    expect(resolveEnvironmentBanner('Production')).toBeNull();
  });

  it('hides the banner for empty / nullish environment', () => {
    expect(resolveEnvironmentBanner('')).toBeNull();
    expect(resolveEnvironmentBanner('   ')).toBeNull();
    expect(resolveEnvironmentBanner(null)).toBeNull();
    expect(resolveEnvironmentBanner(undefined)).toBeNull();
  });

  it('normalizes casing and surrounding whitespace', () => {
    expect(resolveEnvironmentBanner('  dev ')).toEqual({ label: 'Wersja DEV', variant: 'dev' });
    expect(resolveEnvironmentBanner('uat')).toEqual({ label: 'Wersja UAT', variant: 'test' });
    expect(resolveEnvironmentBanner('preprod')).toEqual({ label: 'Wersja PRE', variant: 'pre' });
  });

  it('falls back to a neutral variant for an unknown environment, preserving the name', () => {
    expect(resolveEnvironmentBanner('Sandbox')).toEqual({
      label: 'Wersja Sandbox',
      variant: 'unknown',
    });
  });
});
