import { isValidReturnUrl } from './return-url.util';

describe('isValidReturnUrl', () => {
  it('accepts an absolute https URL', () => {
    expect(isValidReturnUrl('https://app.example.com/return')).toBe(true);
  });

  it('accepts an absolute http URL', () => {
    expect(isValidReturnUrl('http://localhost:5000/callback')).toBe(true);
  });

  it('trims surrounding whitespace before validating', () => {
    expect(isValidReturnUrl('  https://example.com/x  ')).toBe(true);
  });

  it('rejects null', () => {
    expect(isValidReturnUrl(null)).toBe(false);
  });

  it('rejects undefined', () => {
    expect(isValidReturnUrl(undefined)).toBe(false);
  });

  it('rejects an empty string', () => {
    expect(isValidReturnUrl('')).toBe(false);
  });

  it('rejects a whitespace-only string', () => {
    expect(isValidReturnUrl('   ')).toBe(false);
  });

  it('rejects a malformed URL', () => {
    expect(isValidReturnUrl('not a url')).toBe(false);
  });

  it('rejects a non-http(s) protocol', () => {
    expect(isValidReturnUrl('javascript:alert(1)')).toBe(false);
    expect(isValidReturnUrl('ftp://example.com/file')).toBe(false);
  });
});
