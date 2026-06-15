# Analiza integracji Microsoft Entra ID - dokumentacja

## Executive Summary

Niniejszy dokument stanowi kompleksową analizę bezpieczeństwa i architektoniczną implementacji Microsoft Entra ID oraz przedstawia pełny plan wdrożenia dla projektu D2ViewerEditor.

## Kluczowe ustalenia

### ✅ Projekt D2WebCore - działająca implementacja

- Pełna integracja z Microsoft Entra ID.
  - Tenant: `587b6ea1-3db9-4fe1-a9d7-85d4c64ce5cc`
- Dwojakiego rodzaju autentykacja: Keycloak (legacy) oraz Entra ID (nowa).
- Frontend: Angular 19 z MSAL Browser 4.27.0 i MSAL Angular 4.0.23.
- Backend: .NET 8.0 z Microsoft.Identity.Web 3.12.0.
- Hybrydowy model:
  - OpenID Connect dla użytkowników interaktywnych.
  - JWT Bearer dla API.
- Zaawansowane zarządzanie rolami przez mapowanie grup Entra ID.
- Sekrety przechowywane w GCP Secret Manager.
- Integracja z Microsoft Graph API 5.103.0 dla pobierania informacji o użytkownikach.

### ⚠️ Projekt D2ViewerEditor - brak autentykacji

- Obecnie tylko API Key authentication.
- Brak integracji z Entra ID.
- Wymaga pełnego wdrożenia mechanizmu autoryzacji.

---

# Analiza implementacji Entra ID

## Frontend (Angular)

## Znalezione komponenty

### 1. MSAL Configuration Factory

**Lokalizacja:** `D2AngularNew/src/main.ts`

Konfiguracja MSAL ładowana runtime z pliku `config.json`:

```ts
export const MSAL_CUSTOM_CONFIG = new InjectionToken<any>('');

fetch("assets/configs/config.json")
  .then(resp => resp.json())
  .then(config => {
    platformBrowserDynamic([
      { provide: MSAL_CUSTOM_CONFIG, useValue: config }
    ])
      .bootstrapModule(AppModule)
      .catch(err => console.error(err));
  });
```

**Rola:** Umożliwia dynamiczną konfigurację per-środowisko bez przebudowy aplikacji.

> ⚠️ **Ryzyko bezpieczeństwa:** `ClientId` jest publiczny w pliku JSON dostępnym przez HTTP.

---

## 2. MSAL Instance Factory

**Lokalizacja:** `app.module.ts:51-65`

```ts
export function MSALInstanceFactory(): IPublicClientApplication {
  return new PublicClientApplication({
    auth: {
      clientId: msalCustomConfig['clientId'],
      authority: "https://login.microsoftonline.com/587b6ea1-3db9-4fe1-a9d7-85d4c64ce5cc/v2.0",
      redirectUri: msalCustomConfig['redirectUri'],
      navigateToLoginRequestUrl: true
    },
    cache: {
      cacheLocation: BrowserCacheLocation.LocalStorage,
      storeAuthStateInCookie: isIE
    }
  });
}
```

**Kluczowe parametry:**

- **Authority:** hardcoded Tenant ID, brak multi-tenancy.
- **Cache:** `LocalStorage`, persystencja między sesjami.
- **Fallback dla IE:** cookies dla starych przeglądarek.

---

## 3. MSAL Guard Configuration

**Lokalizacja:** `app.module.ts:77-83`

```ts
export function MSALGuardConfigFactory(): MsalGuardConfiguration {
  return {
    interactionType: InteractionType.Redirect,
    authRequest: {
      scopes: [msalCustomConfig['clientId'] + '/.default']
    }
  };
}
```

**Scope:** `{clientId}/.default` — delegated access do wszystkich API permissions zdefiniowanych w App Registration.

---

## 4. MSAL Interceptor Configuration

**Lokalizacja:** `app.module.ts:68-75`

```ts
export function MSALInterceptorConfigFactory(): MsalInterceptorConfiguration {
  const protectedResourceMap = new Map<string, Array<string>>();
  protectedResourceMap.set('https://graph.microsoft.com/v1.0/me', ['user.read']);

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap
  };
}
```

---

# Rekonesans D2WebCore

Widoczny fragment poprzedzający sekcję:

```text
3. Identyfikacja plików konfiguracyjnych
4. Identyfikacja zależności NuGet/npm
```

## [2026-06-10 10:05:00] Rekonesans D2WebCore — wyn[ucięte]

## Znalezione projekty

- **D2WebCore** - główna aplikacja backend .NET 8.0.
- **D2AngularNew** - frontend Angular 19.x.
- **D2ViewerEditor** (`doc2tools/Viewer`) - oddzielna aplikacja bez auth.

## Zależności NuGet związane z auth (`D2WebCore.csproj`)

- `Microsoft.Identity.Web` 3.12.0 ✅
- `Microsoft.Identity.Web.UI` 3.12.0 ✅
- `Microsoft.AspNetCore.Authentication.OpenIdConnect` 8.0.18 ✅
- `IdentityModel` 7.0.0 ✅
- `Microsoft.AspNetCore.Authentication.Negotiate` 8.0.15
- `Google.Cloud.SecretManager.V1` 2.6.0 — dla sekretów GCP

## Zależności npm związane z auth (`D2AngularNew/package.json`)

- `@azure/msal-angular` ^4.0.23 ✅
- `@azure/msal-browser` ^4.27.0 ✅

## Struktura katalogów auth w D2WebCore

```text
D2WebCore/
├── Auth/
│   ├── ClaimsTransformer.cs                  # Transformacja claims AD -> role aplikacji
│   ├── CustomClaimTypes.cs                   # Własne typy claims
│   ├── IsMemberGroupRequirement.cs           # Wymaganie członkostwa w grupie
│   ├── ResourcesProvider.cs                  # Provider zasobów per rola
│   └── Events/
│       ├── Doc2CookieAuthenticationEvents.cs # Obsługa cookie + refresh
│       └── Doc2OpenIdConnectEvents.cs        # Obsługa OIDC events
├── Security/
│   ├── Roles.cs                              # Definicja ról aplikacji
│   ├── MfaAuthorizationFilter.cs             # Filtr MFA
│   └── ExternalRequestFilter.cs
├── Middlewares/
│   └── TokenInjectionMiddleware.cs           # XSRF token injection
├── ConfigurationModels/
│   └── ConfigureAuthentication.Startup.cs    # Konfiguracja auth Entra ID/Keycloak
└── Controllers/
    └── IdentityController.cs                 # API endpoints identity
```

---

## Konfiguracje appsettings

| Plik | Środowisko | Sekcje auth |
|---|---:|---|
| `appsettings.json` | Base | `keyCloakConfig`, `Roles` |
| `appsettings.dev.json` | DEV | `AzureAd`, `GCPSecretManager` |
| `appsettings.tst.json` | TST | `AzureAd`, `GCPSecretManager` |
| `appsettings.acc.json` | ACC | `AzureAd`, `GCPSecretManager` |
| `appsettings.prd.json` | PRD | `AzureAd`, `GCPSecretManager` |
| `appsettings.local.json` | LOCAL | `AzureAd` bez secret |

---

# [2026-06-10 10:10:00] Analiza konfiguracji Entra ID — Backend

## Znalezione konfiguracje AzureAd

### `appsettings.dev.json`

```json
{
  "AzureAd": {
    "Instance": "https://login.microsoftonline.com/",
    "TenantId": "587b6ea1-3db9-4fe1-a9d7-85d4c64ce5cc",
    "ClientId": "fac950da-4c4e-48b3-acd9-afc7fb49c756",
    "Scopes": ["User.Read", "profile"],
    "Proxy": {
      "Url": "http://localhost:3128"
    }
  }
}
```

## GCP Secret Manager

- **Secret Name Pattern:** `SMC01{ENV}2_APP00404_entra_secret`
- **DEV:** `SMC01D2_APP00404_entra_secret`
- **TST:** `SMC01T2_APP00404_entra_secret`
- **ACC:** `SMC01A2_APP00404_entra_secret`
- **PRD:** `SMC01P2_APP00404_entra_secret`

## AuthenticationConfig (`D2BLL/Services/Auth/AuthenticationConfig.cs`)

```csharp
public class AuthenticationConfig
{
    public string Authority { get; init; }
    public string Audience { get; init; }
    public string ClientId { get; init; }
    public string UserNameClaimType { get; init; }
    public int RefreshThresholdMinutes { get; init; }
    public string Instance { get; set; }
    public string TenantId { get; set; }
    public string ClientSecret { get; set; }
    public BusinessProxy Proxy { get; set; }
}
```

## Wnioski

- Aplikacja używa podwójnego mechanizmu auth: Keycloak (legacy) i Entra ID (nowy).
- ClientSecret przechowywany w GCP Secret Manager — **DOBRZE**.
- Proxy konfigurowane dla środowisk enterprise.
- TenantId hardcoded w niektórych miejscach — **UWAGA**.

---

# [2026-06-10 10:15:00] Analiza implementacji auth — ConfigureAuthentication.Startup.cs

## Kluczowe metody

### AddEntraIdAuthentication

**Lokalizacja:** `D2WebCore/ConfigurationModels/ConfigureAuthentication.Startup.cs`

## Funkcjonalność

1. Konfiguracja JWT Bearer authentication dla API.
2. Konfiguracja OpenID Connect dla web app.
3. Token acquisition dla downstream API (Microsoft Graph).
4. In-memory token cache.

## Schemat

```text
AddMicrosoftIdentityWebApi -> JWT Bearer
AddMicrosoftIdentityWebApp -> OpenID Connect + Cookie
EnableTokenAcquisitionToCallDownstreamApi -> Graph calls
AddInMemoryTokenCaches -> Token caching
```

## Wnioski bezpieczeństwa

- ✅ Używa `Microsoft.Identity.Web` — oficjalna biblioteka.
- ✅ PKCE flow przez domyślne ustawienia MSAL.
- ✅ Proxy configuration dla enterprise.
- ⚠️ `ShowPII = true` włączone — **RYZYKO w produkcji**.
- ⚠️ Hardcoded TenantId w niektórych miejscach.
- ⚠️ In-memory token cache — utrata przy restart.

---

# [2026-06-10 10:20:00] Analiza MSAL Angular — Frontend

## Konfiguracja w `app.module.ts`

### MSALInstanceFactory

```ts
return new PublicClientApplication({
  auth: {
    clientId: msalCustomConfig['clientId'],
    authority: "https://login.microsoftonline.com/587b6ea1-3db9-4fe1-a9d7-85d4c64ce5cc/v2.0",
    redirectUri: msalCustomConfig['redirectUri'],
    navigateToLoginRequestUrl: true
  },
  cache: {
    cacheLocation: BrowserCacheLocation.LocalStorage,
    storeAuthStateInCookie: isIE
  }
});
```

### `config.json` (`assets/configs/`)

```json
{
  "clientId": "fac950da-4c4e-48b3-acd9-afc7fb49c756",
  "redirectUri": "https://{hostname}:15333"
}
```

## JWT Interceptor (`jwt.interceptor.ts`)

- Token przechowywany w `localStorage`:
  - `accessToken`
  - `accessTokenExp`
- Silent token refresh przez MSAL.
- Scope: `{clientId}/.default`.

## Auth Guard (`auth.guard.ts`)

- Czeka na token w `localStorage`.
- Sprawdza uprawnienia przez `/api/Identity/Resources`.
- Timeout 10 sekund.

## Wnioski bezpieczeństwa

- ⚠️ Token w `localStorage` — podatny na XSS.
- ⚠️ TenantId hardcoded w `authority`.
- ⚠️ Brak automatycznego logout przy wygaśnięciu sesji.
- ✅ PKCE flow domyślnie w MSAL.
- ✅ Silent token refresh.

---

# [2026-06-10 10:25:00] Analiza RBAC — Role i uprawnienia

## Definicja ról (`Security/Roles.cs`)

| Rola | Opis |
|---|---|
| `UsersStandard` | Podstawowi użytkownicy |
| `Administrator` | Administratorzy |
| `AdministratorKonfiguracja` | Admin konfiguracji |
| `AdministratorUprawnienia` | Admin uprawnień |
| `Developer` | Developerzy |
| `Validator` | Walidatorzy |
| `Loro` | Moduł Loro |

## Mapowanie grup AD → role (`appsettings.json`)

```json
{
  "Roles": {
    "GroupPrefix": "GSAPW4D_DOC2_",
    "Roles": [
      {
        "RoleName": "UsersStandard",
        "GroupNames": [
          "GSAPW4T_DOC2_Pracownik_Non_PRD",
          "GSAPW4D_DOC2_Pracownik",
          "GSAPW4D_DOC2_IW"
        ]
      },
      {
        "RoleName": "Administrator",
        "GroupNames": [
          "GSAPW4D_DOC2_Administrator_Non_PRD",
          "GSAPW4D_DOC2_Administrator"
        ]
      }
    ]
  }
}
```

## ClaimsTransformer

- Transformuje claims:
  - `ClaimTypes.Role`
  - `roles`
  - `groups`
- Mapuje grupy AD na role aplikacji przez `RolesConfig`.

## ResourcesProvider

- Mapuje role na dozwolone zasoby, czyli moduły UI.
- Hardcoded mapowanie rola → zasoby.

## Wnioski

- ✅ Policy-based authorization przez grupy AD.
- ⚠️ Hardcoded mapowanie zasobów w `ResourcesProvider`.
- ⚠️ Brak dynamic permissions z bazy danych.

---

# [2026-06-10 10:30:00] Analiza D2ViewerEditor — brak auth

## Status

**D2ViewerEditor NIE MA implementacji Entra ID!**

## D2ViewerEditor.Api

- Brak `Microsoft.Identity.Web` w dependencies.
- Brak `UseAuthentication` / `UseAuthorization`.
- Brak `[Authorize]` attributes.
- Tylko Swagger i CORS.

## D2GuiViewerEditor (Angular)

- Brak `@azure/msal-angular`.
- Brak `@azure/msal-browser`.
- Brak interceptorów auth.
- Brak guards.

## Wymagane działanie

→ Implementacja pełnej integracji Entra ID wg wzorca D2WebCore.

---

# [2026-06-10 10:35:00] CI/CD i Secrets

## Azure Pipelines

- **Lokalizacja:** `.pipelines/azure-pipelines.yml`
- **Container Registry:** `p09811doc2acr.azurecr.io`
- **Harbor:** `harbor.pl.ing-ad/runtime-commons`
- Checkmarx One scan włączony.

## GCP Secret Manager

- Używany do przechowywania `ClientSecret`.
- **ProjectId:** `gcp-dev-prd-xgkj`
- **Pattern:** `SMC01{ENV}2_APP00404_entra_secret`

## Wnioski

- ✅ Sekrety nie są w repo.
- ✅ Skanowanie bezpieczeństwa: Checkmarx.
- ⚠️ Brak Key Vault rotation policy.
- ⚠️ Brak Managed Identity — używany ClientSecret.

---

# [2026-06-10 10:40:00] Security Findings

## CRITICAL

- Brak auth w D2ViewerEditor — otwarte API.

## HIGH

- `ShowPII = true` w `ConfigureAuthentication`.
- Token w `localStorage` — XSS risk.
- TenantId hardcoded w wielu miejscach.

## MEDIUM

- In-memory token cache — utrata przy restart.
- MFA cookie sprawdzany przez string compare.
- Brak HTTPS enforcement w local dev.

## LOW

- Hardcoded CORS origins.
- Brak rate limiting.
- Brak request signing.

---

# [2026-06-10 10:45:00] Microsoft Graph Integration

## GraphServiceClientProvider

- Używa `ClientSecretCredential` z `Azure.Identity`.
- Scope: `https://graph.microsoft.com/.default`.
- App-only authentication, czyli client credentials flow.

## GraphUserService

- `FindUserAsync` — wyszukiwanie użytkowników w Azure AD.
- Pobiera:
  - `id`
  - `userPrincipalName`
  - `mail`
  - `displayName`
  - `givenName`
  - `surname`
  - `jobTitle`
  - `department`

## Wnioski

- ✅ Integracja z Microsoft Graph dla user lookups.
- ✅ Proper error handling dla 404.
- ⚠️ Sync wrapper `Task.Run` — może blokować.

---

# [2026-06-10 10:50:00] Zakończenie analizy — Podsumowanie

## Obecna architektura D2WebCore

- .NET 8.0 + Angular 19.
- Microsoft.Identity.Web 3.12.0.
- MSAL Angular 4.x.
- OpenID Connect + JWT Bearer.
- GCP Secret Manager dla sekretów.
- Microsoft Graph dla user info.

## D2ViewerEditor - wymaga implementacji

- Brak jakiejkolwiek integracji auth.
- Wymaga pełnego wdrożenia wg wzorca D2WebCore.

## Kluczowe pliki do przeniesienia/adaptacji

1. `ConfigureAuthentication.Startup.cs`
2. `ClaimsTransformer.cs`
3. Auth events — Cookie + OIDC
4. `ResourcesProvider` + `RolesConfig`
5. `TokenInjectionMiddleware`
6. `app.module.ts` — MSAL config
7. `jwt.interceptor.ts`
8. `auth.guard.ts`

---

# Komunikaty widoczne na screenach

```text
Sorry, no response was returned.

Selected "Try Again"
```
