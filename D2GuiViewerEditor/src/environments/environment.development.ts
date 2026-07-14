/**
 * Konfiguracja środowiska deweloperskiego
 */
export const environment = {
  production: false,
  apiUrl: 'http://localhost:5190/api',
  environmentName: 'DEV',
  buildVersion: '1.0.0-dev',
  buildDate: '26/02/2026 23:00',
  // Adres wsparcia (odbiorca zgłoszeń z funkcji „Zgłoś problem"). TODO: podmienić na docelowy.
  supportEmail: 'test@testowy.pl',
  autoSave: {
    enabled: true,
    intervalSeconds: 30
  }
};
