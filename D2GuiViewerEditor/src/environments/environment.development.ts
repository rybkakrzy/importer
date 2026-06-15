/**
 * Konfiguracja środowiska deweloperskiego
 */
export const environment = {
  production: false,
  apiUrl: 'http://localhost:5190/api',
  environmentName: 'DEV',
  buildVersion: '1.0.0-dev',
  buildDate: '26/02/2026 23:00',
  // Auto-save edytora: nadpisuje wersję edytowalną (v2) w miejscu co `intervalSeconds`.
  autoSave: {
    enabled: true,
    intervalSeconds: 30
  },
  // Entra ID (MSAL) — DEV. Uzupełnij wartościami z app registration DEV.
  auth: {
    clientId: '<DEV_SPA_CLIENT_ID>',
    authority: 'https://login.microsoftonline.com/<DEV_TENANT_ID>',
    redirectUri: 'http://localhost:4200/',
    postLogoutRedirectUri: 'http://localhost:4200/',
    apiScopes: ['<DEV_API_SCOPE>'],
    adminRole: 'APP_Admin'
  }
};
