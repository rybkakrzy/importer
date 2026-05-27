export type EnvironmentBannerVariant = 'local' | 'dev' | 'test' | 'pre' | 'unknown';

export interface EnvironmentBannerView {
  label: string;
  variant: EnvironmentBannerVariant;
}

// Production-equivalent names never surface a banner (rule: hidden on PROD by default).
const HIDDEN_ENVIRONMENTS = new Set(['PROD', 'PRD', 'PRODUCTION']);

// UAT and TST share the light-yellow "test" variant but keep distinct labels.
const KNOWN_ENVIRONMENTS: Record<string, EnvironmentBannerView> = {
  LOCAL: { label: 'Wersja Local', variant: 'local' },
  DEV: { label: 'Wersja DEV', variant: 'dev' },
  DEVELOPMENT: { label: 'Wersja DEV', variant: 'dev' },
  UAT: { label: 'Wersja UAT', variant: 'test' },
  TST: { label: 'Wersja TST', variant: 'test' },
  TEST: { label: 'Wersja TST', variant: 'test' },
  PRE: { label: 'Wersja PRE', variant: 'pre' },
  PREPROD: { label: 'Wersja PRE', variant: 'pre' },
  PREPRODUCTION: { label: 'Wersja PRE', variant: 'pre' },
};

/**
 * Maps a raw environment name (any casing) to a banner view, or null when the
 * banner must stay hidden (production / empty). Unknown names get a neutral
 * fallback so a misconfigured environment is still surfaced, not silently lost.
 */
export function resolveEnvironmentBanner(
  environmentName: string | null | undefined,
): EnvironmentBannerView | null {
  const raw = (environmentName ?? '').trim();
  const normalized = raw.toUpperCase();

  if (!normalized || HIDDEN_ENVIRONMENTS.has(normalized)) {
    return null;
  }

  return KNOWN_ENVIRONMENTS[normalized] ?? { label: `Wersja ${raw}`, variant: 'unknown' };
}
