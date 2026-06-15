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
  },
  // Entra ID (MSAL). Wartości per środowisko uzupełniane przy deployu (nie są sekretami).
  auth: {
    clientId: '<SPA_CLIENT_ID>',
    authority: 'https://login.microsoftonline.com/<TENANT_ID>',
    redirectUri: '/',
    postLogoutRedirectUri: '/',
    // Scope eksponowany przez API (api://<API_CLIENT_ID>/access_as_user)
    apiScopes: ['<API_SCOPE>'],
    // Rola admina (claim "roles") — używana tylko do UX (ukrycie menu). Backend = źródło prawdy.
    adminRole: 'APP_Admin',
    // Domyślnie auth włączony. Lokalny dev bypass: ustaw `auth.enabled=false` w runtime config.json.
    enabled: true
  }
};
