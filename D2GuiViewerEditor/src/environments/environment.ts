/**
 * Konfiguracja środowiska produkcyjnego
 */
export const environment = {
  production: true,
  apiUrl: 'http://localhost:5190/api',
  environmentName: 'PRD',
  buildVersion: '1.0.0',
  buildDate: '26/02/2026 23:00',
  // Adres wsparcia (odbiorca zgłoszeń z funkcji „Zgłoś problem"). TODO: podmienić na docelowy.
  supportEmail: 'test@testowy.pl',
  autoSave: {
    enabled: true,
    intervalSeconds: 30
  }
};
