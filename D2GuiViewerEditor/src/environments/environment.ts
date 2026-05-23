/**
 * Konfiguracja środowiska produkcyjnego
 */
export const environment = {
  production: true,
  apiUrl: 'http://localhost:5190/api',
  environmentName: 'PRD',
  buildVersion: '1.0.0',
  buildDate: '26/02/2026 23:00',
  // Auto-save edytora: nadpisuje wersję edytowalną (v2) w miejscu co `intervalSeconds`.
  autoSave: {
    enabled: true,
    intervalSeconds: 30
  }
};
