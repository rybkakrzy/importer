/**
 * Validates the "return URL" used to send a document back to the source app
 * after editing (Krok 4). The "Zakończ" action is only meaningful when this
 * link is present and usable, so the UI must hide the button otherwise.
 *
 * A link counts as valid only when it is a non-empty, absolute http(s) URL.
 * Empty / whitespace-only / null / undefined / non-http values are rejected.
 */
export function isValidReturnUrl(returnUrl: string | null | undefined): boolean {
  const raw = (returnUrl ?? '').trim();
  if (!raw) {
    return false;
  }

  try {
    const parsed = new URL(raw);
    return parsed.protocol === 'http:' || parsed.protocol === 'https:';
  } catch {
    return false;
  }
}
