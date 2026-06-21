# AI-Assisted Changelog

Istotne zmiany dla kontynuacji pracy (nie zastępuje changeloga produktu).

## Format

```md
## YYYY-MM-DD — <tytuł>
### Changed
### Verified
### Notes
```

## Entries

## 2026-06-22 — D2Services observability: middleware + correlation id + JSON payload parity
### Changed
- `D2ServicesViewerEditor.Api/Program.cs`:
  - dodano `UseRequestObservability()` i `UseExceptionHandlingMiddleware()` do pipeline,
  - rozszerzono `UseSerilogRequestLogging` o `EnrichDiagnosticContext` (`httpMethod`, `httpPath`, `statusCode`, `requestId`, `correlationId`).
- Dodano `D2ServicesViewerEditor.Api/Middleware/RequestObservabilityMiddleware.cs`:
  - nagłówek `X-Correlation-ID` (echo/generowanie),
  - `HttpContext.Items[CorrelationId]`,
  - scope (`correlationId`, `requestId`) i Serilog `LogContext` dla pełnej korelacji.
- `D2ServicesViewerEditor.Api/Middleware/ExceptionHandlingMiddleware.cs`:
  - `correlationId` w `ProblemDetails.Extensions`,
  - serializacja `ProblemDetails` po typie runtime (zachowanie `errors` dla `ValidationProblemDetails`).
- `D2ServicesViewerEditor.Api/Logging/GcpJsonSerilogFormatter.cs`:
  - ujednolicony payload do standardu D2Api: `severity`, `level`, `timestamp`, `message`, `category`, `service`, `environment`, `traceId`, `spanId`, `exceptionType`, plus spłaszczone właściwości eventu/scope.
### Verified
- `dotnet test D2ServicesViewerEditor/D2ServicesViewerEditor.Api.UnitTests/D2ServicesViewerEditor.Api.UnitTests.csproj` → **28/28** pass.
### Notes
- Zmiana usuwa wcześniejsze luki: brak globalnego middleware wyjątków i brak korelacji requestów w D2Services.

## 2026-06-21 — D2Services: testy HealthController + MiddlewareExtensions
### Changed
- Dodano `HealthControllerTests` w `D2ServicesViewerEditor.Api.UnitTests`:
  - `Get_ReturnsOk_WithExpectedPayloadShape`,
  - `GetDetailed_ReturnsOk_WithDependenciesSection`.
- Dodano `MiddlewareExtensionsTests`:
  - `UseExceptionHandlingMiddleware_ReturnsSameBuilderInstance`.
### Verified
- `dotnet test D2ServicesViewerEditor/D2ServicesViewerEditor.Api.UnitTests/D2ServicesViewerEditor.Api.UnitTests.csproj` → **22/22** pass.
### Notes
- Rozszerzenie domyka podstawowe klasy API po wcześniejszym dodaniu testów dla kontrolera dokumentów, middleware wyjątków i formattera logów.

## 2026-06-21 — Zwiększenie liczby testów jednostkowych (D2Api + D2Services)
### Changed
- **D2ServicesViewerEditor:** utworzono nowy projekt `D2ServicesViewerEditor.Api.UnitTests` (dodany do `D2ServicesViewerEditor.sln`) z testami:
  - `DocumentControllerTests` (walidacja wejścia, wymagany `ReturnUrl` dla DOCX, rozpoznanie MIME po rozszerzeniu, mapowanie odpowiedzi statusu),
  - `ExceptionHandlingMiddlewareTests` (mapowania wyjątków 400/401/500 + pass-through),
  - `GcpJsonSerilogFormatterTests` (mapowanie `LogEventLevel`→`severity`, `category` z `SourceContext`, serializacja wyjątków).
- **D2ApiViewerEditor:** rozszerzono `RequestObservabilityMiddlewareTests` o przypadki brzegowe `X-Correlation-ID` (whitespace, zbyt długi >128); rozszerzono `ExceptionHandlingMiddlewareTests` o propagację `correlationId` do `ProblemDetails.Extensions`.
### Verified
- `dotnet test D2ServicesViewerEditor/D2ServicesViewerEditor.Api.UnitTests/D2ServicesViewerEditor.Api.UnitTests.csproj` → **19/19** pass.
- `dotnet test D2ApiViewerEditor/D2ViewerEditor.Api.UnitTests/D2ViewerEditor.Api.UnitTests.csproj` → **76/76** pass.
### Notes
- W projekcie testowym D2Services wersja pakietu `Serilog` ustawiona na `4.0.0`, aby uniknąć NU1605 (downgrade względem `Serilog.Sinks.Console`).

## 2026-06-21 — „Zakończ" sync + statusy, „Pobierz oryginał", ukrycie „Ostatnia modyfikacja"
### Changed
- **DocumentStatus** +`Queued` („Zlecono do wysyłki"), +`SendAborted` („UzytkownikPrzerwałWysyłkę"); `Document.MarkQueued/MarkSendAborted`.
- **DocumentDelivery** +`BeginInlineAttempt()` (Pending/RetryScheduled→Sending, inline, bez lease) +`HoldAfterFailedInlineAttempt()` (RetryScheduled zaparkowane na DeadlineAt — worker nie przejmie).
- **FinishAndSendDocumentCommandHandler** — synchroniczna pierwsza próba przez `IDeliverySender`; wynik `{DeliveryId,Status,DocumentStatus,Delivered,Error?}`. Endpoint `finish` → 200 (było 202).
- Nowe komendy/endpointy: `POST /{masterId}/abort-send` (Przerwij → delivery `Cancelled` + doc `SendAborted`), `POST /{masterId}/continue-delivery` (Kontynuuj w tle → `Requeue` + doc `Queued`).
- **GUI** `document-editor`: nowa pozycja „Plik → Pobierz oryginał dokumentu" (`canDownloadOriginal = loadedFromDisk || userDownload`, zwraca v1); flow „Zakończ" = „Trwa wysyłanie dokumentu ..." → sukces zamyka kartę / błąd: modal „Wystąpiły problemy..." + Przerwij/Kontynuuj; „Ostatnia modyfikacja" w stopce gated `showSaveState()`.
- `infra/sql/010_extend_document_status.sql` (COMMENT only — kolumna bez CHECK).
### Verified
- `dotnet build` D2Api + D2Services: 0 błędów. Domain 77, Application 282 zielone. GUI Vitest 267 zielone.
- `ng build` AOT kompiluje; błąd tylko z pre-existing budżetu bundla (nietknięte pliki).
### Notes
- Decyzje potwierdzone z użytkownikiem: synchroniczna pierwsza próba (nowy endpoint) + angielski PascalCase enuma (etykiety PL tylko w UI).
- „Pobierz oryginał" z aplikacji zewn. używa `GET {masterId}/download` (view-gated, jak podgląd Krok 2) — oryginał był już pobieralny dla każdego z dostępem do podglądu; `userDownload` bramkuje tylko pobranie EDYTOWANEJ kopii.

## 2026-06-19 — `showSaveState` (ingest) — ukrywanie UI zapisu w edytorze
### Changed
- **D2Services** `DocumentController.CreateDocument`: nowe opcjonalne pole `ShowSaveState: bool?` w `CreateDocumentRequest`; zapisywane do `documents.metadata` jako `showSaveState` **tylko gdy jawne `false`** (inverse-default — brak/`true` → `null` → widoczne).
- **Application**: `ExternalDocumentMetadata` +`ShowSaveState` (pozycyjny param z domyślnym `null`, nie łamie istniejących wywołań) + computed `IsSaveStateVisible => ShowSaveState != false`. `DocumentMetadataDto` +`ShowSaveState`; `GetDocumentMetadataQueryHandler` mapuje `IsSaveStateVisible`.
- **GUI** `document-editor`: sygnał `showSaveState` (domyślnie `true`), ustawiany z `meta.showSaveState !== false`. Szablon ukrywa switch Autozapis, status autozapisu w stopce i przycisk „Zapisz" gdy `false`. `DocumentMetadataDto` (service) +`showSaveState`.
### Verified
- `dotnet build` D2Api + D2Services: 0 błędów. Testy Application: 31/31 (metadata). GUI Vitest: 263/263 (w tym 2 nowe testy `showSaveState`).
### Notes
- Decyzja: `showSaveState=false` ukrywa **tylko UI** — autozapis dalej działa w tle (brak utraty danych), a „Zakończ i wyślij" pozostaje. Reguła prezentacyjna, nie zabezpieczenie.

## 2026-06-19 — Kestrel: globalny limit body 150 MB (duże dokumenty) + AddServerHeader=false
### Changed
- `Program.cs` `ConfigureKestrel`: `MaxRequestBodySize` konfigurowalny (`Kestrel:MaxRequestBodySizeBytes`, domyślnie 150 MB), `MinRequestBodyDataRate`/`MinResponseDataRate=null`, `AddServerHeader=false`. Powód: endpointy `DocumentStorageController` (upload/save/finish, base64) nie miały limitu → domyślny ~30 MB Kestrel groził **413** dla dużych dokumentów. Per-action `[RequestSizeLimit]` (DocumentController) dalej obowiązuje gdzie ostrzejszy.
### Verified
- `dotnet build` 0 błędów.
### Changed (cd.)
- **Wymuszenie `https`** (wzorzec D2WebCore, bezwarunkowo): middleware `context.Request.Scheme = "https"` jako pierwszy w pipeline — poprawne URL-e / OIDC redirect URI za TLS-terminującym proxy (potrzebne dla OIDC z ADR-0019).
### Notes
- Z przeglądu wzorca D2WebCore `Program.cs` **pominięto**: NLog (mamy structured `ILogger` — ADR-0014/0018), `InitializeCulture` (konwertery już używają `InvariantCulture`), `CodePagesEncodingProvider` (kod używa wbudowanego `Encoding.Latin1`, nie code-page; wymagałby nowego pakietu), inline GCP secret load (mamy `EntraSecretLoader`), `UseWindowsService`/`Startup` (kontenery + minimal hosting).

## 2026-06-19 — Parytet z D2WebCore: ClaimsTransformer (szersze typy claimów) + proxy z credentialami
### Changed
- `ClaimsTransformer`: dopasowuje identyfikatory grup/ról niezależnie od typu claimu (`groups`/`roles`/`ClaimTypes.Role`/`*identity/claims/role*`) — jak wzorcowy `ClaimsTransformer`. **Output nadal `roles`** (nasz `RoleClaimType`), nie `ClaimTypes.Role` (inaczej `IsInRole` by nie działał). Tolerancja null `GroupNames`.
- `ProxyOptions` (= `BusinessProxy`): +`Username`/`Password`; `EntraBackchannel` używa `NetworkCredential` gdy username podany, inaczej `UseDefaultCredentials`. Dodane do appsettings (puste; Password z secret store).
- **Nazwy klas bez zmian** (`AzureAdOptions`/`ClaimsTransformer`/`RolesOptions`) — ugruntowane; pominięto martwe pola wzorca (`UserNameClaimType`/`RefreshThresholdMinutes`).
### Verified
- `dotnet build` 0 błędów; Api.UnitTests **73/73** (+`CreateProxy_uses_explicit_credentials`, +`Maps_group_identifier_carried_under_role_claim_type`).

## 2026-06-19 — Entra: jawny `IS_LOCAL_DEV` + serwerowy OIDC `AddMicrosoftIdentityWebApp` (hybryda WebApi+WebApp)
### Changed
- `AddEntraIdAuthentication` (ConfigureAuthentication): **jawny** `IS_LOCAL_DEV` (env var) steruje proxy — lokalnie `UseProxy=false`, inaczej `WebProxy` z `AzureAd:Proxy:Url` (bypass GCS) jako `BackchannelHttpHandler` (WebApi + WebApp) i `HttpClient.DefaultProxy`.
- **Dodano serwerowy OIDC** `AddMicrosoftIdentityWebApp` (scheme `MyAzureAdScheme`): code flow + cookie (`SignInScheme`, `NonceCookie/CorrelationCookie SecurePolicy=Always`, scope `offline_access`/`email`, `ResponseType=Code`) + `EnableTokenAcquisitionToCallDownstreamApi()` + `AddInMemoryTokenCaches()`. WebApi (JWT) pozostaje **domyślnym** schematem dla API SPA. Na wyraźną prośbę — **odwraca** „resource-server-only" z ADR-0011/0012. Patrz ADR-0019.
- Bez nowej zależności (`AddMicrosoftIdentityWebApp` tranzytywnie z `Microsoft.Identity.Web` 3.12.0). `Microsoft.AspNetCore.Authentication.OpenIdConnect` ściągany tranzytywnie.
### Verified
- `dotnet build` 0 błędów; Api.UnitTests **71/71**. **Runtime NIEzweryfikowane** (brak tenanta/ClientSecret lokalnie): faktyczny flow OIDC, cookie, token acquisition — patrz R-27.

## 2026-06-19 — Refaktor: wiring Entra do `ConfigureAuthentication` (parytet z D2WebCore)
### Changed
- Wydzielono auth z `Program.cs` do `Security/ConfigureAuthentication.cs`: `AddDevBypassAuthentication()` + `AddEntraIdAuthentication(configuration, environment)` (bind AzureAd, ShowPII gated, proxy backchannel, Identity.Web JWT, RoleClaimType PostConfigure, grupy→role, polityki, ClaimsCurrentUserProvider, Graph). `Program.cs` = `if (devAuthBypass) AddDevBypass… else AddEntraId…`. Zachowanie 1:1.
- Dodano (z wzorca, gated) `JwtBearerOptions.IncludeErrorDetails = !IsProduction()` — szczegóły 401 w dev/TST, off w PRD.
- ~~Świadomie NIE dodano `AddMicrosoftIdentityWebApp`~~ — **zmienione w nowszym wpisie (powyżej)**: dodane na wyraźną prośbę (ADR-0019).
### Verified
- `dotnet build` 0 błędów; Api.UnitTests **71/71**.

## 2026-06-17 — 3 zgłoszenia: payload wysyłki (master/version/corporateKey), widok „Brak uprawnień", observability ELK
### Changed
- **Z1 (wysyłka):** `HttpDeliverySender` (multipart) dokłada obok `file` pola `masterId`/`versionId`/`corporateKey` + log techniczny (bez treści, corporateKey jako flaga). `DeliveryDispatch` +`MasterId`/`VersionId`/`CorporateKey`; `DeliveryAttemptRunner` wypełnia z encji. `DocumentDelivery` +`CorporateKey` (kolumna `corporate_key`, SQL `009`, EF map), `FinishAndSendDocumentCommandHandler` czyta `ICurrentUserProvider.CorporateKey`. Backward compatible (pole `file` bez zmian).
- **Z2 (frontend):** nowy generyczny widok `AccessForbiddenComponent` „Brak uprawnień" (`/brak-uprawnien`); `resourceGuard`+`documentAccessGuard`+interceptor (403) kierują tu; `/access-denied` → redirect. Usunięto `document-access-denied`. 403 nie mylone z 401/404/500.
- **Z3 (ELK):** `GcpJsonConsoleFormatter` wzbogacony (scope'y→pola, `service`/`environment`/`traceId`/`spanId`/`level`); nowy `RequestObservabilityMiddleware` (correlationId z `X-Correlation-ID` + access-log method/path/status/elapsed/userId, scope na cały request); `correlationId` w ProblemDetails. Dokumentacja `.ai/OBSERVABILITY.md` (pola + KQL Kibana + zasady nie-logowania danych wrażliwych).
### Verified
- Backend: `dotnet build` 0 błędów; unit testy **577** (Api 71, Application 267, Infrastructure 166, Domain 73) — w tym naprawiony wcześniej istniejący `GetDocumentVersionContentQueryHandlerTests`. GUI: `tsc` czysto; `ng test` **261**.
### Notes
- corporateKey = claim `ck` użytkownika kończącego (async wysyłka → utrwalony na delivery); brak → puste pole (bez błędu).
- Integration testy delivery na realnym Postgres niezweryfikowane lokalnie (jak dotąd, R-06).

## 2026-06-17 — Konsumpcja proxy Entra: `AzureAd:Proxy:Url` → JwtBearer backchannel (D2WebCore pattern)
### Changed
- Nowy `EntraBackchannel.CreateProxy(AzureAdOptions)` — buduje `WebProxy` z `AzureAd:Proxy:Url` (bypass `storage.googleapis.com`, `UseDefaultCredentials=true`) lub `null` gdy brak URL. `Program.cs` (gałąź Entra): gdy proxy ustawione → `HttpClient.DefaultProxy = proxy` oraz `JwtBearerOptions.BackchannelHttpHandler` (pobieranie OpenID metadata/JWKS zza korpo-proxy). Brak URL / lokalnie → bez zmian (direct).
### Verified
- `dotnet build` 0 błędów; `dotnet test` Api.UnitTests **67/67** (dodano `EntraBackchannelTests` 3).
### Notes
- `IdentityModelEventSource.ShowPII` z wzorca dodany, ale **gated `!IsProduction()`** (nigdy w PRD — wyciek PII to HIGH w analizie; w PRD off). Gating na `!IsProduction()` a nie `IsDevelopment()`, bo env nazywa się „DEV"/itp., nie „Development".
- **Graph/Azure.Identity** (`ClientSecretCredential` w `GraphUserService`) NIE jest objęty `HttpClient.DefaultProxy` — Azure.Identity używa własnego pipeline'u (proxy via env `HTTPS_PROXY` lub `TokenCredentialOptions.Transport`). Follow-up, jeśli Graph ma działać zza proxy.

## 2026-06-17 — Finalne nazewnictwo ról: `Administrator` / `Operator` (rename z `APP_Admin`/`APP_Pracownik`)
### Changed
- Backend: globalny rename wartości `APP_Admin`→`Administrator`, `APP_Pracownik`→`Operator` w `AzureAdOptions` (default `AdminRole`), `appsettings.json`/`appsettings.DEV.json` (sekcja `Roles`→`RoleName`), `DevAuthHandler` (claimy), testach (`ClaimsTransformerTests`, `IdentityControllerTests`) + komentarzach.
- Rename **identyfikatorów**: właściwość `AzureAdOptions.EmployeeRole`→`OperatorRole` (klucz `AzureAd:OperatorRole`), stała polityki `RequireAppEmployee`→`RequireAppOperator` (nazwa+wartość), `isEmployee`→`isOperator`, nazwy testów. `RequireAppAdmin`/`AdminRole` bez zmian. Semantyka `ResourcesProvider`/polityk bez zmian.
- **GUI:** brak zmian kodu — autoryzacja resource-based (ADR-0016), front nie zna nazw ról (potwierdzone grepem).
- Docs „żywe" (SECURITY/DOMAIN/FEATURES/API_CONTRACTS/CURRENT_STATE) zaktualizowane na nowe nazwy; historyczne ADR-y zachowane.
### Verified
- API: `dotnet build` 0 błędów; `dotnet test` Api.UnitTests **64/64**.
### Notes
- Patrz ADR-0017. R-26: realne Entra appRoles/`RolesOptions` muszą emitować `Administrator`/`Operator` (albo nadpisać `AzureAd:OperatorRole`/`AdminRole`).
- Niezwiązane: `D2ViewerEditor.Application.UnitTests` ma **wcześniej istniejący** błąd kompilacji (`GetDocumentVersionContentQueryHandlerTests` — brak arg. `accessGuard`) — nie z tej zmiany.

## 2026-06-17 — Autoryzacja resource-based (backend /identity/resources + ResourcesProvider, front resourceGuard)
### Changed
- **Backend (D2ApiViewerEditor):** nowy `ResourcesProvider` (rola→zasoby: employee/admin→editor,viewer; admin→+admin) + endpoint `GET /api/identity/resources` (`RequireAppEmployee`). Polityka admina w `IdentityController` przeniesiona z klasy na akcję `users`. DI `AddScoped<ResourcesProvider>()`. `AzureAdOptions` + `Scopes[]`/`Proxy{Url}` (+ appsettings.json/DEV) — powierzchnia konfiguracji (konsumpcja = follow-up).
- **Front (D2GuiViewerEditor):** `ResourceAccessService` (cache + dev-bypass) + `resourceGuard` (gating po `route.path` względem listy z backendu, fail-closed→/access-denied). `resourceGuard` zastąpił `documentRoleGuard`/`appAdminGuard` na /editor,/viewer,/admin. **Usunięto:** `documentRoleGuard`, `appAdminGuard`, `CurrentUserService` (+spec), pola `adminRole`/`viewerRole` z `AppAuthConfig`/`config.json` — front nie zna już nazw ról.
### Verified
- API: `dotnet build` 0 błędów; `dotnet test` Api.UnitTests **64/64** (IdentityController 7). GUI: `tsc` czysto; `ng test` **262/262** (dodano `resource-access.service.spec` 3; usunięto `current-user.service.spec` 5).
### Notes
- Zamyka R-25 (rozjazd nazw ról front↔backend — front już nie zna ról). Otwiera **R-26**: backend `IsInRole` wymaga, by Entra emitowało `APP_Pracownik`/`APP_Admin` (albo nadpisać `AzureAd:EmployeeRole`/`AdminRole`). Patrz ADR-0016.
- `Scopes`/`Proxy` to na razie tylko config — konsumpcja (proxy na Graph/token, downstream scopes) niewdrożona.

## 2026-06-16 — Front Entra: MSAL standalone (redirect-component), config.json jako źródło auth, model ról Administrator/Przeglądający
### Changed
- `main.ts`/`index.html`: `MsalRedirectComponent` (`<app-redirect>`) bootstrapowany przez `appRef.bootstrap(...)` (standalone, bez AppModule); `App` nie woła już `handleRedirectObservable()` (aktywne konto po `inProgress$===None`).
- Wartości auth usunięte z `environment.*` → jedyne źródło to `src/assets/configs/config.json` (przeniesiony z `public/`). `DEFAULT_AUTH_CONFIG` = neutralne defaulty strukturalne. Token DI `RUNTIME_AUTH_CONFIG` → `MSAL_CUSTOM_CONFIG`.
- `msal.config.ts`: helper `apiScopesFor()` (jawne `apiScopes` lub fallback `{clientId}/.default`), jawne `navigateToLoginRequestUrl: true`.
- RBAC (UX): `AppAuthConfig.viewerRole` (domyślnie „Przeglądający"), `adminRole` domyślnie „Administrator". `CurrentUserService` scentralizowany (`isAdmin`/`isViewer`/`canAccessDocuments`, admin=nadzbiór, dev-bypass). Nowy `documentRoleGuard` na `/editor` i `/viewer`; `appAdminGuard` przez `CurrentUserService`.
- Lokalny `config.json` = dev-bypass (`auth.enabled=false`) — aplikacja startuje bez realnej rejestracji Entra.
### Verified
- `tsc --noEmit -p tsconfig.app.json` — czysto. `ng test --watch=false` — **264/264** (dodano `current-user.service.spec` — 5; zaktualizowano `app.spec`, `runtime-config.spec`).
### Notes
- **Rozjazd nazw ról z backendem** (ADR-0011/0012: `APP_Pracownik`/`APP_Admin`): dosynchronizować Entra appRoles/`RolesOptions` lub nadpisać `adminRole`/`viewerRole` w config.json. Patrz ADR-0015.
- Niezweryfikowane runtime (brak tenanta lokalnie): realny redirect + claim `roles`.

## 2026-06-13 — Fix: ProblemDetails serializowany przez typ runtime (errors w body)
### Changed
- `ExceptionHandlingMiddleware` serializował `problemDetails` przez **statyczny typ bazowy** `ProblemDetails` → `ValidationProblemDetails.Errors` ginęło w body (System.Text.Json honoruje typ statyczny). Fix: `JsonSerializer.Serialize(problemDetails, problemDetails.GetType(), options)` — serializacja po typie runtime, więc słownik `errors` (zgrupowany po `PropertyName`) trafia do odpowiedzi 400.
- Test `InvokeAsync_ValidationException_EmitsGroupedErrorsInBody` asercjonuje teraz realną obecność `errors.Name`/`errors.Email` w body (poprzedni test „realnego zachowania bez errors" zastąpiony zachowaniem docelowym).
### Verified
- `dotnet test D2ViewerEditor.Api.UnitTests` — **51** pass (+1). Reszta solucji bez zmian.
### Notes
- Kontrakt API wzbogacony (dodane pole `errors` w 400 dla błędów walidacji) — zgodne z RFC 7807 `ValidationProblemDetails`; klienci dotychczas i tak nie dostawali `errors`, więc brak regresji.

## 2026-06-13 — Pokrycie testami: pipeline behaviours, metadata, exception middleware
### Changed
- **Application.UnitTests** (+19): `ExternalDocumentMetadataTests` (parse tolerancyjny null/empty/malformed/json-null, mapowanie pól, case-insensitive, `IsUserDownloadAllowed` tylko dla jawnego `true`, round-trip serialize→parse, camelCase keys); `ValidationBehaviourTests` (brak walidatorów→next, wszystkie pass→next, fail→`ValidationException`+short-circuit, agregacja błędów z wielu walidatorów); `LoggingBehaviourTests` (przekazanie odpowiedzi, log na wejściu+wyjściu, brak połykania wyjątku).
- **Api.UnitTests** (+7): `ExceptionHandlingMiddlewareTests` — pass-through bez wyjątku, mapowanie `ValidationException`→400 problem+json, `ArgumentException`→400 z detail, `KeyNotFoundException`→404, `OperationCanceledException`→400, nieoczekiwany→500 **bez przeciekania** szczegółów, camelCase ProblemDetails.
### Verified
- `dotnet test D2ViewerEditor.sln` — Domain 67, Api **50** (+7), Application **258** (+19), Infrastructure 164, Integration 6 skip. 0 failures.
### Notes
- Odkryta latentna obserwacja: middleware serializuje `ValidationProblemDetails` przez statyczny typ `ProblemDetails`, więc słownik `errors` NIE trafia do body (test asercjonuje realne zachowanie: 400 + tytuł „Błąd walidacji", bez `errors`). Nie zmieniano produkcyjnego zachowania w ramach dodawania testów.
- Nadal nietestowane (świadomie, wymaga infra/seam): `DeliveryAttemptRunner`, `DocumentDeliveryWorker`, `OoxmlAgileDecryptor`, repo/GCS.

## 2026-06-11 — „Otwórz w edytorze": z osobnej zakładki na przycisk w wierszu wysyłki
### Changed
- **Usunięto** stronę `admin-open-document` (+route `/admin/open`, +pozycję sidebaru w `admin-shell`, +spec).
- **Dodano** w `admin-deliveries` przycisk „Otwórz" per wiersz → `openInEditor(item)`: `router.navigate(['/editor'], { queryParams })` z `masterId=documentId`, `versionId=sourceVersionId` (bez versionId → podgląd). `Router` wstrzyknięty do komponentu.
### Verified
- `admin-deliveries.spec` **13** (+2 nawigacja z/bez sourceVersionId); `ng build` OK.

## 2026-06-11 — Wysyłka na returnUrl jako multipart/form-data
### Changed
- `HttpDeliverySender.SendAsync`: zamiast `ByteArrayContent` (`application/octet-stream`) wysyła `MultipartFormDataContent` — plik w polu **`file`**, filename `document.docx`, part Content-Type = DOCX MIME. Nagłówki `Idempotency-Key`/`X-Content-SHA256` i klasyfikacja statusów bez zmian.
### Verified
- `Infrastructure.UnitTests` +3 `HttpDeliverySenderTests` (multipart + name=file/filename + nagłówki; 422→Permanent; 503→Retryable). Build OK.
### Notes
- **Zmiana kontraktu wysyłki**: odbiorca na `returnUrl` musi odczytać plik z pola form-data `file` (np. `IFormFile file`), nie z surowego body. Nazwa pola jest stała (`file`) — gdyby odbiorca wymagał innej, do sparametryzowania w `HttpDeliverySender`.

## 2026-06-11 — Observability: strukturalne logi JSON dla GCP (severity)
### Changed
- **Internal API** (`D2ViewerEditor.Api`): nowy `Logging/GcpJsonConsoleFormatter` (ConsoleFormatter → JSON z `severity`, `message`+stack trace, `category`, `eventId`, `exceptionType`); `Extensions/LoggingExtensions.AddGcpStructuredLogging()` wpięte w `Program.cs`. Aktywne poza Development lub gdy `Logging:UseGcpFormat=true`.
- **External API** (`D2ServicesViewerEditor.Api`): nowy `Logging/GcpJsonSerilogFormatter : ITextFormatter`; `Program.cs` `WriteTo.Console(new GcpJsonSerilogFormatter())` poza Development (inaczej zwykły tekst). File-sink bez zmian.
### Verified
- `dotnet build` obu hostów OK; `Api.UnitTests` **43** (+8 `GcpJsonConsoleFormatterTests`: mapowanie poziomów, wyjątek=ERROR+stack w message, jedna linia JSON).
- Manualnie do potwierdzenia w GCP: po deployu wyjątki w Logs Explorer mają severity ERROR/CRITICAL i są filtrowalne; Error Reporting grupuje po stack trace.
### Notes
- Root cause: stdout plain-text → Cloud Logging nadaje INFO; rozwiązanie = pole `severity` w JSON. ADR-0014.
- Follow-up: `LoggingBehaviour` loguje `{@Request}` (pełny payload — hałas/dane wrażliwe), do okrojenia osobno.

## 2026-06-11 — Fix: „Zamknij" w modalu wysyłki zostawiał edytowalny dokument
### Changed
- `document-editor`: nowy sygnał `workFinished`; `closeFinishModalAndExit()` zatrzymuje auto-save i ustawia stan końcowy zamiast tylko zamykać modal. Nowy blokujący `work-finished-overlay` (z-index 3000) z przyciskiem „Zamknij kartę" (`closeFinishedTab()`). `window.close()` pozostaje best-effort, ale gdy zawiedzie, użytkownik widzi ekran końcowy i nie wraca do edycji.
### Verified
- `ng build` OK; `document-editor.spec` **62** (+1: po „Zamknij" `workFinished=true`, auto-save off, modal zamknięty).
### Notes
- Root cause: `window.close()` działa tylko dla kart otwartych przez `window.open` — dla zwykłych/otwartych przez link przeglądarka odmawia. Dlatego potrzebny jawny stan końcowy w UI.

## 2026-06-11 — „Pliki do wysłania": edycja adresu odbiorcy (returnUrl)
### Changed
- **Domena:** `DocumentDelivery.UpdateRecipientUrl(url)` — walidacja absolutnego http(s) + blokada dla `Sent`/`Sending`.
- **Application/API:** `UpdateDeliveryRecipientUrlCommand`+handler (przycina URL, mapuje Invalid/Argument → Failure, brak → NotFound); endpoint `PUT /api/documentstorage/deliveries/{id}/recipient-url` (body `{ recipientUrl }`) + `UpdateDeliveryRecipientUrlRequest`.
- **GUI:** `admin-deliveries` — przycisk „Edytuj" (gdy `canEdit`: ≠ Sent/Sending) → modal (`edit-overlay`/`edit-dialog`) z polem URL, walidacją i Zapisz/Anuluj; sygnały `editingId`/`editUrlValue`/`editError`/`savingEdit`; autoodświeżanie pauzowane przy otwartym modalu. `updateDeliveryRecipientUrl()` + `UpdateDeliveryRecipientUrlResult` w `document-storage.service`. `dl-btn-primary` (SCSS).
### Verified
- Backend: `dotnet build` sln OK; Domain `DocumentDelivery` 21 (+5 UpdateRecipientUrl), Application `UpdateDeliveryRecipientUrlCommandHandler` 3.
- GUI: `ng build` OK; `admin-deliveries.spec` 11 (+4 edycja).
### Notes
- Bez migracji SQL (zmiana wartości kolumny `recipient_url`, nie schematu). Po zmianie adresu zadanie zwykle wymaga „Wznów".

## 2026-06-11 — „Pliki do wysłania": autoodświeżanie 3 s + akcje Anuluj/Wznów
### Changed
- **Domena:** `DeliveryStatus.Cancelled` (nowy stan końcowy); `DocumentDelivery.IsTerminal` obejmuje Cancelled; `Cancel()` (Pending/RetryScheduled→Cancelled, czyści lease, ustawia LastError „Anulowano ręcznie"); `Requeue()` rozszerzony o RetryScheduled (wyślij teraz) i Cancelled (poprzednio tylko DeadLettered/FailedPermanently). Worker bez zmian — claim selektuje Pending/RetryScheduled/stuck-Sending, więc Cancelled jest wykluczony.
- **Application/API:** `CancelDeliveryCommand`+handler (anulowanie + `Document.MarkSaved`); endpoint `POST /api/documentstorage/deliveries/{id}/cancel`. „Wznów" reużywa istniejącego `/retry` (Requeue).
- **SQL:** `infra/sql/008_add_delivery_cancelled_status.sql` — DROP+ADD `ck_document_deliveries_status` z wartością `Cancelled` (indeksy częściowe due/active bez zmian).
- **GUI:** `admin-deliveries` — autoodświeżanie co 3 s (`interval`, ciche `fetch(silent)`, przełącznik, `OnDestroy`); przyciski „Wznów" (`canResume`) i „Anuluj" (`canCancel`, `dl-btn-danger`); `cancelDelivery()` w `document-storage.service`; `DeliveryStatus` +`Cancelled`; statusLabel „Anulowano"/statusClass `status-cancelled`; status w dropdownie filtra.
### Verified
- Backend: `dotnet build` sln OK; Domain `DocumentDelivery` 16 (+6 Cancel/Requeue), Application `CancelDeliveryCommandHandler` 3.
- GUI: `ng build` OK; `ng test` **244** (+7 `admin-deliveries.spec`).
### Notes
- **Wymaga uruchomienia migracji `008` na każdym środowisku** (CHECK constraint). Bez niej zapis statusu `Cancelled` odrzuci baza.
- Anulowanie zablokowane dla `Sending` (lease workera) — uniknięcie wyścigu z trwającą próbą.

## 2026-06-11 — 5 zgłoszeń: dialog hasła, .doc, GUI „wysyłanie", EMF→PNG, VersionId w mailu
### Changed
- **(P1) Dialog hasła** — `document-editor.html`: usunięto `(click)=cancelPasswordDialog()` z overlay i `(keydown.escape)` z pola (dialog zamykają tylko `Anuluj`/`Otwórz`; dodano `role=dialog`/`aria-modal`). `document-editor.ts` `cancelPasswordDialog()`: bez `returnUrl` → `router.navigate(['/'])`, z `returnUrl` → `tryCloseBrowserTab()` — koniec pustego dokumentu po anulowaniu.
- **(P3) Modal „Trwa wysyłanie pliku…"** — zamiast ikony papierowego samolotu używa `.loading-spinner` (klasa `.finish-dialog-spinner`, +SCSS `margin-bottom:16px`) — spójny z ekranem „Przetwarzanie…". Logika/countdown bez zmian.
- **(P5) Mail „Zgłoś"** — `openReportEmail()` podstawia `Version ID` = `documentVersionId()` (było `documentMetadata().version` = zwykle `—`); fallback `—` tylko gdy wersji brak; `Master ID` bez zmian.
- **(P4) EMF/WMF → PNG** — `GraphicConversionService`: `TryRasterizeMetafileToPng` + `TryExtractEmfDib` (rekordy STRETCHDIBITS/SETDIBITSTODEVICE/BITBLT/STRETCHBLT/ALPHABLEND) + `TryFindDibGeneric` + `WrapDibInBmpFile` + `DecodeToPng` (SkiaSharp). `DocxToHtmlConverter`: `data-original-src` niesiony dla każdego metafile (drawing + VML), nie tylko placeholdera → eksport wektorowy zachowany.
- **(P2) Binarny .doc → .docx** — nowy `LegacyDocBinaryConverter` (FIB + piece table → tekst+akapity → DOCX/OpenXML); wpięty w `DocumentInputNormalizer` z fallbackiem `UnsupportedLegacyDoc`.
### Verified
- Backend: `dotnet build` sln OK (0 błędów); `Infrastructure.UnitTests` 155 (+1 EMF→PNG `Emf_WithEmbeddedDib_IsRasterizedToPng`, +4 `LegacyDocBinaryConverterTests`); `Api.UnitTests` DocumentController 13.
- GUI: `ng build` OK; `ng test` **233** (+4 w `document-editor.spec`: VersionId obecny/fallback, cancel bez/z returnUrl).
### Notes
- Ograniczenia świadome: czysto wektorowe EMF (bez rastra) wciąż placeholder; `.doc` odzyskuje tekst+akapity, nie formatowanie/tabele/obrazy. Patrz ADR-0013. Pełna wierność = sidecar LibreOffice (roadmapa).
## 2026-06-10 — Pełny wzorzec Qutas/D2WebCore dla Entra ID (Identity.Web + grupy→role + Graph + Secret Manager + Keycloak + runtime config)
### Changed (ADR-0012)
- **Backend — biblioteka:** `JwtBearer` (goły) → **`Microsoft.Identity.Web` 3.12.0** (`AddMicrosoftIdentityWebApi`). Pin `JwtBearer 8.0.12` **usunięty** (Identity.Web dostarcza per-TFM; pin dawał NU1605 na net9.0). Dodano `AzureAd:ClientId`/`ClientSecret`.
- **Backend — grupy→role (Qutas):** `RolesOptions` (sekcja `Roles`: `GroupPrefix` + `Roles[]{RoleName,GroupNames}`) + `ClaimsTransformer : IClaimsTransformation` mapuje claim `groups` → role `APP_Pracownik`/`APP_Admin`. **App Roles zachowane** (współistnieją). Polityki bez zmian. Test `ClaimsTransformerTests` **6/6**.
- **Backend — Microsoft Graph v5 (5.103.0):** `IGraphUserService`/`GraphUserService` (app-only `ClientSecretCredential`) + `GET /api/identity/users?query=` (`IdentityController`, RequireAppAdmin). Aktywny tylko z ClientSecret; inaczej `DisabledGraphUserService`. (NIE `Identity.Web.MicrosoftGraph` = Graph v4.)
- **Backend — GCP Secret Manager** (`Google.Cloud.SecretManager.V1` 2.6.0): `EntraSecretLoader` wstrzykuje `AzureAd:ClientSecret` z `SMC01{ENV}2_APP00404_entra_secret`; guard `Enabled` + try/catch (nigdy nie wywala startu).
- **Backend — Keycloak dual-auth:** gdy `Keycloak:Enabled` — drugi schemat JwtBearer + policy scheme `EntraOrKeycloak` wybierający po issuerze tokena (`AuthSchemes.SelectByIssuer`). Wyłączony → tylko Entra.
- **Frontend — runtime config (Qutas):** `public/assets/configs/config.json` ładowany w `main.ts` przed bootstrapem → `RUNTIME_AUTH_CONFIG` (root-factory fallback = `environment.auth`). Fabryki MSAL (`msal.config`), `appAdminGuard`, `CurrentUserService` czytają z runtime-configu. Jeden build na wszystkie środowiska.
- **Konfiguracja:** `appsettings.json`/`appsettings.DEV.json` rozszerzone o `AzureAd.ClientId/ClientSecret`, `Roles`, `Keycloak`, `GCPSecretManager` — wszystko **placeholdery** (zero realnych sekretów/tenanta).
### Verified
- `dotnet build` API OK (Identity.Web/Graph 5.103.0/Azure.Identity/SecretManager restore). `Api.UnitTests` **41** (+6 `ClaimsTransformerTests`). GUI `tsc` OK, `ng test` **236**, `ng build` (AOT) OK; `config.json` shippowany do `dist/.../assets/configs/`.
- **Niezweryfikowane runtime** (brak tenanta/GCP/Keycloak lokalnie): walidacja tokenów Entra, mapowanie grup z realnego tokena, pobranie sekretu z GCP, selekcja schematu Keycloak. Aktywacja wymaga realnych wartości per środowisko + app consent dla Graph (`User.Read.All`).
### Notes
- Adaptacja Qutas do kształtu SPA+API: **bez** serwerowego OIDC/cookie (`AddMicrosoftIdentityWebApp`) — interaktywny login robi MSAL w przeglądarce, backend = resource server.
- Zmiana wykonana na bieżącym drzewie (niezakończony merge `feature/azure`); `.claude/settings.json` nadal do rozwiązania przez użytkownika.

## 2026-06-09 — „Lista plików" (admin-files): statusy dokumentów po polsku
### Changed
- `admin-files.ts` `statusLabel` — etykiety EN→PL: `Saved`→„Zapisany", `Editing`→„W edycji", `Sending`→„Wysyłanie do odbiorcy", `DeliveryFailed`→„Odbiorca nie odpowiada", `Sent`→„Wysłany". Wartości enuma `DocumentStatus` z API (klucze) bez zmian; `statusClass` (kolory) bez zmian. Filtr statusu jest free-text i matchuje po `statusLabel`, więc działa na polskich etykietach.
### Verified
- `tsc --noEmit` OK. Brak testów odwołujących się do starych etykiet EN.

## 2026-06-09 — Panel administracji (GUI): domyślnie firmowa czcionka
### Changed
- `admin-shell.scss` `:host` — `font-family: var(--corporate-font-family, 'Calibri', 'Segoe UI', Arial, sans-serif)`. Cała zawartość panelu admina (sidebar + strony routowane `admin-files`/`admin-deliveries`) dziedziczy firmowy krój. `--corporate-font-family` to istniejący punkt konfiguracji z `assets/fonts/_corporate-font.scss` (gdy firmowy font nieskonfigurowany → fallback Calibri/Segoe UI; gdy skonfigurowany — automatycznie się podstawi, jak w edytorze).
- `styles.scss` — globalna reguła `d2-admin-shell input, select, button, textarea { font-family: inherit }`. Natywne kontrolki formularzy mają własny font systemowy i nie dziedziczą; reguła musi być globalna (nie component-scoped), bo dotyczy kontrolek w stronach routowanych przez `<router-outlet>`.
### Verified
- `ng build` (AOT) OK; `admin-shell.spec` 3/3. Bez zmian logiki/TS.

## 2026-06-09 — Panel „Pliki do wysłania": rozwijane szczegóły wysyłki po kliknięciu wiersza
### Changed
- **Backend:** `DeliveryListItemDto` rozszerzony o `SourceVersionId` (Guid) i `RecipientUrl` (string) — mapowane w `GetDeliveriesByStatusQueryHandler` z `DocumentDelivery.SourceVersionId`/`RecipientUrl` (oba już istniały na encji; `recipientUrl` = returnUrl, nie jest sekretem — i tak wystawiany przez `…/metadata`).
- **Frontend model:** `DeliveryListItem` +`sourceVersionId`, +`recipientUrl`.
- **`admin-deliveries`:** sygnał `expandedId` + `toggleExpand(id)`/`isExpanded(id)` (jeden wiersz rozwinięty naraz). Klik w `.doc-row` rozwija/zwija; przycisk „Ponów" ma `event.stopPropagation()`, więc nie koliduje. Wiersz `.detail-row` (colspan=11) renderuje grid: **Adres odbiorcy**, **Master ID** (=documentId), **Version ID** (=sourceVersionId), **Status**, **Utworzono**, **Ostatnia próba**, **Komunikat błędu**. Caret ▸ obraca się 90° po rozwinięciu.
### Verified
- Backend `Application.UnitTests` **205** (`GetDeliveriesByStatusQueryHandlerTests` = 8; +1 test mapowania `RecipientUrl`/`SourceVersionId`/`DocumentId`). `dotnet build` OK.
- Frontend `tsc` OK; `ng build` (AOT) OK — `admin-deliveries` kompiluje template z rozwijanym wierszem.

## 2026-06-09 — Panel „Pliki do wysłania": domyślnie wszystkie statusy (zamiast tylko DeadLettered)
### Changed
- **Przyczyna pustego panelu:** domyślny filtr statusu = `DeadLettered`, a świeżo zakolejkowane wysyłki mają `Pending`/`Sending`/`RetryScheduled`/`Sent`. Endpoint `GET …/deliveries` wymagał konkretnego statusu, więc panel startował pusty.
- **Backend:** `GET …/documentstorage/deliveries` — parametr `status` opcjonalny: pusty / `null` / `"all"` (case-insensitive) ⇒ wszystkie statusy. `IDocumentDeliveryRepository.GetAllAsync(skip, take)` (+impl w `DocumentDeliveryRepository`, `OrderByDescending(CreatedAt)` jak `GetByStatusAsync`). `GetDeliveriesByStatusQuery.Status` → `string?` (default `null`); handler rozgałęzia all / konkretny status / nieznany (Failure). Kontroler: domyślny `status` `"DeadLettered"` → `""`.
- **Frontend:** `DocumentStorageService.getDeliveriesByStatus` → **`getDeliveries(status: DeliveryStatus | null = null, …)`** — pomija param `status` w żądaniu, gdy `null`. `admin-deliveries`: `selectedStatus` typu `DeliveryStatus | 'all'`, domyślnie **`'all'`**; w dropdownie nowa opcja **„Wszystkie"** (przekazuje `null` do serwisu).
### Verified
- Backend `D2ViewerEditor.Application.UnitTests` **204** (+7 nowy `GetDeliveriesByStatusQueryHandlerTests`: empty/whitespace/null/all/ALL → `GetAllAsync`; konkretny → `GetByStatusAsync`; nieznany → Failure bez zapytań). `dotnet build` solucji OK (0 błędów).
- Frontend `tsc --noEmit` OK; `ng build` (AOT) OK — `admin-deliveries` kompiluje się z nową opcją.
### Notes
- To zmiana tylko domyślnego widoku + dodanie opcji „wszystkie". Jeśli `document_deliveries` jest pusta (nikt nie dokończył „Zakończ"), panel nadal pokaże „Brak wyników" — to poprawne.

## 2026-06-09 — UX „Zakończ": modal „Trwa wysyłanie pliku" + odliczanie 30 s + best-effort zamknięcie karty
### Changed
- **Frontend `document-editor.ts`** (`finishDocument` przebudowany): po zapisie (`saveDocument`) i zleceniu wysyłki (`finishAndSend`) od razu otwiera modal „Trwa wysyłanie pliku do Twojej aplikacji" z przyciskiem **„Zamknij (NN)"** odliczającym 30→0. Nowe sygnały `showFinishModal`, `finishCountdown`; metody `openFinishModal`, `onFinishCountdownTick` (wydzielony tyk — deterministycznie testowalny bez fake-timerów), `closeFinishModalAndExit` (wspólna akcja dla końca odliczania i kliknięcia „Zamknij"), `tryCloseBrowserTab` (`window.close()` w try/catch — best-effort, bez crashu gdy przeglądarka odmówi). Odliczanie przez `timer(1000,1000)` w `finishCountdownSub`.
- **Usunięto blokujący `pollDeliveryStatus`** (timer + `getDeliveryStatus` + toasty „Dokument został wysłany"/„nie powiodła się"): nie czekamy już na stan końcowy dostarczenia — robi to backendowy `DocumentDeliveryWorker` (retry/backoff do 24h). Usunięty też nieużywany import `takeWhile` oraz pola `deliveryPollSub`/`DELIVERY_POLL_MS`.
- **Błąd natychmiastowego zakolejkowania** (`finishAndSend` rzuca): toast **„Nie udało się natychmiast wysłać pliku. Ponowimy próbę wysłania w tle."** — modal i odliczanie zostają (brak blokady użytkownika).
- **Ochrona przed wielokrotnym „Zakończ"**: guard `if (isFinishing()) return;` + `[disabled]="isFinishing()"` na przycisku; backendowy `finish` jest dodatkowo **idempotentny** (re-klik = to samo zadanie).
- **`template`**: nowy modal `@if (showFinishModal())` (overlay BEZ zamykania kliknięciem w tło — tylko przycisk „Zamknij", by nie dało się przypadkowo zamknąć/ponowić). **`scss`**: `.finish-dialog-overlay` (z-index/blur jak leave-dialog) + `.finish-close-btn` (stała `min-width`, by licznik nie skakał).
- **Higiena**: subskrypcja `finishSendSub` anulowana w `ngOnDestroy` (brak zapisów do sygnałów po zniszczeniu komponentu → koniec „document is not defined" przy teardownie testów).
### Verified
- GUI pełny zestaw **229/229** (`document-editor.spec.ts` **57**, +8 nowych: modal się pokazuje, countdown 30→29→28, dojście do 0 zamyka+`window.close`, klik „Zamknij" robi to samo, błąd→toast retry-w-tle + modal otwarty, wielokrotny klik = 1 flow, wyjątek `window.close` nie crashuje, readOnly bez modala). `tsc --noEmit` OK.
- **Backend bez zmian** — wykorzystany istniejący endpoint `finish` + worker/retry.
### Notes
- Eksperymentalny runner (Vitest unit-test builder) **nie wspiera `fakeAsync`** → odliczanie testowane przez bezpośrednie wołanie `onFinishCountdownTick(n)`; łańcuch async domykany `vi.waitFor`.
- `window.close()` zadziała niezawodnie tylko dla kart otwartych skryptem — przy zwykłym wejściu użytkownika karta może nie zostać zamknięta; to świadomie zaakceptowany, nie-krytyczny stan (modal zamknięty, edytor zostaje).

## 2026-06-09 — Podbicie podatnej zależności System.Security.Cryptography.Xml 8.0.2 → 10.0.6 (PRISMA)
### Changed
- `D2ViewerEditor.Infrastructure.csproj`: dodany **jawny** `PackageReference` na `System.Security.Cryptography.Xml` **10.0.6**. Pakiet nie był nigdzie referencjonowany bezpośrednio — wchodził **tranzytywnie** (przez NPOI/OpenXml) w wersji 8.0.2, flagowanej przez PRISMA. Jawny pin podbija rozwiązaną wersję (wyższa bezpośrednia wygrywa nad tranzytywną).
- Pin wpisany **tylko w Infrastructure**: oba hosty (`D2ViewerEditor.Api`, `D2ServicesViewerEditor.Api`) konsumują Infrastructure przez `ProjectReference`, więc 10.0.6 propaguje się do wszystkich konsumentów.
### Verified
- `dotnet list package --include-transitive` we **wszystkich trzech** projektach: `System.Security.Cryptography.Xml` rozwiązany na **10.0.6** (zamiast 8.0.2).
- `dotnet build -c Release` (Infrastructure + oba API): „Kompilacja powiodła się" — **0 błędów, brak ostrzeżeń NU1605** (downgrade). net8.0 konsumuje pakiet 10.0.x bez problemu (projekt już wcześniej mieszał `Microsoft.Extensions.Configuration.Abstractions 10.0.4` w net8.0).
### Notes
- Zmiana wyłącznie w build/restore — zero zmian kodu. Warto przepuścić pełny `dotnet test` (NUnit) dla potwierdzenia braku regresji w ścieżce podpisów (XML signing).

## 2026-06-09 — ENTER w polu rozmiaru czcionki kasował zaznaczony tekst
### Changed
- **Root cause:** `onFontSizeInputEnter` wołał `input.blur()` synchronicznie w trakcie obsługi ENTER. `setFontSize` przywraca wtedy fokus+zaznaczenie do edytora JESZCZE w trakcie tego zdarzenia, więc domyślna akcja Enter (nowa linia) trafia w przywrócone zaznaczenie i je kasuje. Klik poza pole (blur myszką) nie miał problemu — brak zdarzenia Enter.
- **Fix:** `onFontSizeInputEnter` robi `event.preventDefault()` (zabija domyślny Enter) + `setTimeout(() => input.blur(), 0)` — aplikacja przez blur dzieje się PO zakończeniu zdarzenia Enter, rozdzielona od niego. Ścieżka blur (klik) bez zmian.
### Verified
- GUI **221** (+2: ENTER preventDefault + odroczony blur; blur emituje fontSizeChange od razu).

## 2026-06-09 — Dialog hasła wyzwalany TEŻ przy ładowaniu z bazy (nie tylko „Plik→Otwórz")
### Changed
- **Root cause zgłoszenia „nie widzę gdzie podać hasło":** dialog pojawiał się tylko w ścieżce dyskowej (`loadDocument`), a realny przepływ (dashboard upload → editor `loadFromStorage`, odświeżenie, link z aplikacji zewn.) pobierał bajty i wołał `/open`, ale jego error-handler robił tylko `showError` — kod `PASSWORD_REQUIRED`/`WRONG_PASSWORD` ignorowany.
- Wspólny `_convertAndLoad(file, fileName, password?, announce?)` + `_applyLoadedContent` — używany przez OBIE ścieżki. `loadFromStorage` pobiera bajty → buduje `File` → `_convertAndLoad` (zamiast inline openDocument). Dialog hasła ma teraz callback ponawiający (`_passwordRetry`), więc działa niezależnie od źródła (dysk vs baza).
### Verified
- GUI `tsc` OK + **219** (+3: PASSWORD_REQUIRED z /open otwiera dialog; zatwierdzenie ponawia konwersję z hasłem; WRONG_PASSWORD pokazuje błąd).
### Notes
- Dashboard zapisuje zaszyfrowane bajty surowo (v1/v2), edytor odszyfrowuje przy otwarciu (prompt) — po edycji+autosave v2 staje się odszyfrowany. Read-only base (v1) zaprosi o hasło ponownie. Akceptowalne (v1 immutable = oryginał zaszyfrowany).

## 2026-06-09 — Stylizowany dialog hasła (zamiast window.prompt)
### Changed
- `document-editor`: `window.prompt` zastąpiony modalem `.password-dialog` spójnym z `leave-dialog` (overlay z blur, ikona kłódki, pole `type=password` z focus-ringiem, komunikat błędu, Anuluj/Otwórz). Sygnały `showPasswordDialog`/`passwordDialogValue`/`passwordDialogError` + `openPasswordDialog`/`confirmPasswordDialog`/`cancelPasswordDialog`. Enter=Otwórz, Esc/klik-tło=Anuluj, autofokus pola; przy błędnym haśle dialog wraca z komunikatem i ponawia `loadDocument(file, pwd)`.
### Verified
- GUI `tsc` OK + **216** (+3: confirm bez hasła→błąd, cancel czyści, confirm zamyka).

## 2026-06-09 — Dwa bugi: dekrypcja hasłem (NPOI nie działa→własna) + input font-size gubił selekcję
### Changed
- **Hasło — root cause i realna naprawa:** `NPOI 2.8.0` ma interfejs dekryptora, ale **NIE implementację Agile** (typ `AgileEncryptionInfoBuilder` nieobecny we wszystkich TFM-ach → `EncryptionInfo(fs)` rzuca `EncryptedDocumentException`), więc poprzedni kod mapował każdy plik (nawet z dobrym hasłem) na „WrongPassword". Nowy `OoxmlAgileDecryptor` — **własna, czysto zarządzana dekrypcja Agile** (MS-OFFCRYPTO §2.3.4.10+: SHA512 spinCount + AES-256-CBC, blockKeys, weryfikacja hasła, dekrypcja pakietu segmentami 4096B). NPOI używane już TYLKO do czytania kontenera CFB (`CreateDocumentInputStream`). `DocumentInputNormalizer` wymaga `EncryptionInfo` **i** `EncryptedPackage`; rozróżnia WrongPassword (niezgodność weryfikatora) od Invalid/Unsupported (Standard/CryptoAPI Office 2007 nieobsługiwany).
- **Font-size input (Qutas-FMT-004, naprawa właściwa):** `saveSelection()` zapisywał też ZWINIĘTE selekcje, a leciał na `blur`/`selectionchange` — czyli dokładnie gdy klik w input toolbara zwija zaznaczenie do karetki „wciąż w edytorze" → realne zaznaczenie nadpisywane pustą karetką → „nie ma na czym". Fix: nowy strażnik `editorHasFocus()` (activeElement w obszarze edytowalnym) — `saveSelection` zapisuje TYLKO gdy edytor ma fokus; przy przejściu do toolbara zostaje zaznaczenie z `mouseup`/`keyup`.
### Verified
- Backend Infrastructure **150** (+2: realny round-trip Agile encrypt→decrypt — poprawne hasło → odszyfrowany ważny DOCX otwierany OpenXml SDK; złe hasło → WrongPassword; brak hasła → PasswordRequired; enkryptor testowy też spec-zgodny). `D2ViewerEditor.sln` build OK. GUI **213** (+1: saveSelection nie gubi zaznaczenia po utracie fokusu).
### Notes
- Aktualizacja ADR-0010: NPOI **nie** dekryptuje (vs wcześniejszy zapis) — dekrypcja jest własna; NPOI zostaje tylko jako czytnik CFB. Standard/CryptoAPI (Office 2007) → status Invalid (rzadkie; ewentualny follow-up).

## 2026-06-09 — Otwieranie: DOCX z hasłem + .doc + opcjonalna klasyfikacja (Services)
### Changed
- **Hasło (Task 1):** nowy `IDocumentInputNormalizer` (Domain) + `DocumentInputNormalizer` (Infrastructure) — detekcja formatu po magic-bytes, dekrypcja DOCX zabezpieczonego hasłem przez **NPOI** (POIFS/Crypt, pure-managed, Linux/GCP-safe). Wpięty w `OpenDocumentQueryHandler` (przed `Convert`). `OpenDocumentQuery` + kontroler `/open` przyjmują `password`; sentinele `PASSWORD_REQUIRED`/`WRONG_PASSWORD` → 422 z `code`. GUI: `document.service.openDocument(file, password?)` + typed `OpenDocumentError`; edytor „Plik → Otwórz" przez `/open` z promptem hasła i retry.
- **.doc (Task 2):** normalizer wykrywa CFB: mislabeled-DOCX (ZIP w .doc) → pass-through (działa); binarny .doc (stream WordDocument) → `UNSUPPORTED_LEGACY_DOC` → 400 z instrukcją konwersji. Pickery (`dashboard`, edytor) dopuszczają `.doc`. **Brak wbudowanej konwersji binarnego .doc** — żaden pure-managed NuGet nie konwertuje (NPOI bez HWPF; b2xtranslator NuGet bez DocFileFormat; LibreOffice zakazany) → kontrolowany komunikat, nie udawanie. Patrz DECISIONS ADR-0010.
- **Klasyfikacja (Task 3):** `D2ServicesViewerEditor` `POST /api/v1/document` — pole `Classification` **opcjonalne** (brak/puste = bez klasyfikacji; podana musi być C1..C4). Metadata serializuje `classification` jako null gdy brak.
- **SkiaSharp** 3.116.1 → **3.119.2** (wyrównanie do tranzytywnej zależności NPOI, NU1605).
### Verified
- Backend: Application **8** (OpenDocument, +4 sentinel/password), Infrastructure **148** (+4 normalizer: ZIP/garbage/encrypted-no-pwd/binary-doc, fixtury CFB przez NPOI), Api **OpenDocument 4** (binary-doc→400+code, password→422). `D2ViewerEditor.sln` build OK. Services API build OK. GUI `tsc` OK + **212** testów.
### Notes
- Ścieżka „utwórz nowy z dysku" w dashboardzie zapisuje surowo (bez normalizera) → .doc/hasło kierujemy do edytora „Otwórz" (ścieżka /open). Normalizacja przy ingest/storage = ewentualny follow-up.
- Prompt hasła = `window.prompt` (funkcjonalne); stylizowany dialog = follow-up.

## 2026-06-08 — Zmiana czcionki (font-family) nie wpływała na nowo wpisywany tekst (FMT-005)
### Changed
- **Root cause:** `WysiwygEditorComponent.setFontFamily` wołał `editor.focus()` **przed** odtworzeniem zapisanej selekcji. Wybór czcionki z natywnego `<select>` w toolbarze zabiera fokus i czyści selekcję contenteditable; `focus()` przywracał karetkę na **początek dokumentu** (`rangeCount > 0`, więc strażnik `rangeCount === 0` nie odpalał `restoreSelection`), więc ZWS-span z `font-family` lądował w złym miejscu. Tekst wpisywany w realnej pozycji karetki dalej dziedziczył domyślną/firmową czcionkę. `setFontSize` miał już to zabezpieczenie — `setFontFamily` nie.
- **Fix:** `setFontFamily` odtwarza selekcję **ZANIM** dotknie fokusu (gdy live-selection nie jest w edytorze, a jest `savedSelection`) — lustrzane do `setFontSize`. Dodatkowo: gdy karetka siedzi już w ZWS-spanie z poprzedniego wyboru, aktualizuje jego `font-family` zamiast zagnieżdżać kolejny pusty span; zapisuje nową pozycję karetki + `updateFormattingState()`; po zmianie selekcji na zaznaczeniu odświeża `savedSelection`.
### Verified
- GUI `wysiwyg-editor.toolbar-actions.spec.ts` **7/7** (+1 nowy: `setFontFamily` odtwarza zapisaną selekcję, gdy fokus opuścił edytor). `tsc --noEmit` OK.
- Manualnie do potwierdzenia w przeglądarce (jsdom nie odwzorowuje contenteditable/Selection/focus): wybór czcionki przy zwiniętej karetce → nowo wpisany tekst dostaje wybraną czcionkę i round-tripuje do `w:rFonts`.
### Notes
- `pendingFontFamily`/`pendingFontSize` nadal tylko zapisywane (nieczytane) — ścieżka „brak selekcji" polega na ZWS-spanie; nie ruszane w tej zmianie.

## 2026-06-08 — EMF/eksport + font-size import + page-break import + selekcja font-size
### Changed
- **Qutas-IMG-009 (EMF psuje DOCX) — root cause + fix:** writer pisał placeholder SVG jako goły `a:blip` (SVG bez rastra = NIEPOPRAWNY OOXML → Word „uszkodzony"). Reader (`DocxToHtmlConverter`) niesie teraz oryginalny metafile w `data-original-src`; writer (`ResolveImageSrc`) preferuje go i zapisuje prawdziwy **EMF part** (`ImagePartType.Emf`, Word renderuje natywnie); `BuildImageDrawing` ma **twardy guard**: nigdy nie emituje gołego SVG blip (return null). Mapowanie x-emf/x-wmf dodane.
- **DWA pre-existing bugi schematu w `styles.xml`** (wykryte OpenXmlValidatorem, psuły KAŻDY zapis): heading `w:rPr` miał złą kolejność (`sz`/`color`/`b`/`i`) → poprawione na `rFonts→b→i→color→sz`; heading `w:pPr` miał `spacing` przed `keepNext` → `keepNext→keepLines→spacing→outlineLvl`.
- **Qutas-IMP-005 (14pt → ~10pt) — root cause frontend:** reader poprawny (4 testy: direct/styl/docDefaults/Normal). Bug: `_flattenTopBlocks` ROZWIJA wrapper `.document-content`, na którym reader trzyma default → ginął. Fix: `_captureDocumentDefaults` czyta font-size/family z wrappera; nowe sygnały `documentDefaultFontSize/Family` zbindowane na `[style.font-size/font-family]` contenteditable strony.
- **Qutas-IMP-008 (3 strony → 1):** reader IGNOROWAŁ `w:pageBreakBefore`. Fix: `HasPageBreakBefore` → emit `<div class="page-break">` przed akapitem (ten sam mechanizm co manualny break; writer round-tripuje).
- **Qutas-FMT-004 (input font-size gubił selekcję) — root cause:** `savedSelection` zapisywany tylko na `blur` (zbyt późno — selekcja już znika przy klik w input). Fix: `onSelectionChange` zapisuje selekcję na bieżąco. Dodatkowo Enter w input nie aplikuje podwójnie (był apply + blur→apply → zagnieżdżone spany).
### Verified
- Backend `dotnet test Infrastructure.UnitTests` → **144 passed** (+EMF round-trip z OpenXmlValidator=0 błędów, +4 FontSizeImport, +2 pageBreakBefore). Frontend `ng test` → **211 passed** (+2 captureDocumentDefaults).
### Notes
- EMF: dla osadzonego rastra w EMF/WMF zapis idzie rastrem (PNG/JPEG) — też poprawny. Pełna rasteryzacja EMF w przeglądarce nadal roadmapa (placeholder), ale ZAPIS jest teraz wierny (oryginalny EMF) i niełamiący.

## 2026-06-08 — Edytor: 4 realne naprawy (paginacja / Backspace / context-menu / PDF diag)
### Changed
- **Qutas-PAG-002 (Enter wydłuża stronę):** `_schedulePaginate` debounce **600 → 250 ms** (komentarz mówił 300 — kod zdryfował). Strona `min-height:1122px; overflow:visible` rosła widocznie przez całe okno debounce zanim treść spłynęła; 250 ms eliminuje rozciąganie poza A4.
- **Qutas-PAG-001 (Backspace na 2. stronie):** nowy `_tryMergeAcrossPageBackwards()` wpięty w `handleKeyboard` (`_tryDeletePageBreakBackwards() || _tryMergeAcrossPageBackwards()`). Wcześniej obsługiwany był TYLKO manualny page-break; strony z AUTO-paginacji nie scalały się (osobne contenteditable → przeglądarka nie łączy). Teraz: pusty blok wiodący → usuń; bloki mergeable (P/DIV/H/LI) → scal treść jak Word; tabela/niekompatybilne → tylko nawigacja karetki (treść nietknięta). Po operacji repaginacja spływa treść w górę.
- **Qutas-UI-006 (context-menu zasłania UI):** `onContextMenu` clamp z `Math.max(8, …)` na obu osiach — wcześniej dolny clamp dawał ujemne `y` na niskim oknie → menu nad viewportem zasłaniało toolbar.
- **Qutas-PDF-010 (PDF „prace techniczne"):** `pdf-viewer.ts` `catch {}` → `catch (err)` z `console.error` + rozróżnieniem awarii workera (cold-start/404/MIME `.mjs`) od uszkodzonego pliku. Przestaje maskować przyczynę.
### Verified
- `npx tsc --noEmit` OK; `npx ng test --watch=false` → **209 passed** (+4 merge w `caret-field.spec.ts`, +3 context-menu w `document-editor.spec.ts`).
### Notes
- Diagnoza Problem 4 (font-size input): apply-path JEST poprawny (`applyFontSizeToSelection` re-selektuje wstawioną treść 1996–2006; `setFontSize` odtwarza `savedSelection`). Residualne ryzyko = stała `Range` po repaginacji — wymaga selekcji path-based (większa zmiana, nie ruszane).

## 2026-06-08 — Akapit: „Ustaw jako domyślne" zapisuje zamiast resetować (Qutas-PAR-007)
### Changed
- `document-editor.html` — przycisk „Ustaw jako domyślne" wołał `resetParagraphDefaults()` (reset do wartości bazowych) → teraz `setParagraphAsDefault()`.
- `document-editor.ts` — usunięto `resetParagraphDefaults()`; dodano `setParagraphAsDefault()` (snapshot bieżących ustawień → `_paragraphDefaults` + `applyParagraphSettings()` na bieżącym akapicie). Nowe pole `_paragraphDefaults` (default per-sesja). `readCurrentParagraphSettings()` przy braku selekcji seeduje formularz z `_paragraphDefaults` zamiast zostawiać stałe wartości.
- Model trwałości: per-sesja edytora (reset po odświeżeniu); bez localStorage (brak wzorca w app) i bez zmian API. Nowe akapity dziedziczą styl po bieżącym bloku (contenteditable) → default propaguje się naturalnie.
### Verified
- `npx ng test --watch=false` → **202 passed** (+3 nowe w `document-editor.spec.ts`: brak resetu do bazowych / seed dialogu z defaultu / kopia niezależna). Reszta pakietu bez regresji.
### Notes
- Pełna diagnostyka 11 zgłoszeń edytora w raporcie sesji; pozostałe 10 były już wdrożone w sesjach 2026-06-03/04 (zweryfikowane w drzewie: metody klawiatury/paste/font-size/page-break/grafiki obecne). Ten wpis dotyczy jedynego niewdrożonego wcześniej zgłoszenia.

## 2026-06-04 — Health-check API: pierwszy strzał natychmiast po starcie + dedup
### Problem
- Komunikat „brak komunikacji z API" pojawiał się zbyt późno. Pierwszy check zależał od **implicit constructor side-effect** `ConnectionStatusService` (lazy `providedIn:'root'`) — nieodporne na regresję; brak strażnika duplikatów (start + zdarzenie `online` + interwał mogły wystrzelić równoległe żądania).
### Changed (frontend)
- `connection-status.service.ts`: prywatne `checkApi()` → **publiczne `checkNow()`** z **strażnikiem in-flight** (`_checkInFlight`) — pojedyncze żądanie w locie, brak duplikatów. Guard `if (!apiConfig.baseUrl) return` (brak fałszywego offline zanim URL znany — tu URL statyczny z `environment`). Konstruktor, interwał (30 s) i handler `online` wołają `checkNow()`.
- `app.config.ts`: `provideAppInitializer` woła **`inject(ConnectionStatusService).checkNow()`** (jawnie, nie polega na side-effekcie konstruktora) — pierwszy check zaraz po inicjalizacji aplikacji, przed renderem komponentów. Non-blocking (nie zwraca/await). Strażnik in-flight scala check z konstruktora i z initializera w **jedno** żądanie.
- Subsequent polling bez zmian (30 s `setInterval`, retry/timeout 5 s). Banner/offline-status bez zmian.
### Verified
- GUI **199** (+2: „flags API unreachable quickly on first check"; „checkNow() nie duplikuje w locie"). Istniejące 5 testów `ConnectionStatusService` zielone (konstrukcja nadal = 1 natychmiastowe żądanie).
- Scenariusz manualny: start przy API down → pierwszy `/health` od razu po starcie → status 0 → banner offline „szybko"; start przy API up → `/health` 200 → brak bannera.

## 2026-06-04 — Konwerter grafik VML/EMF/WMF → web (pure-managed, bez LibreOffice)
Pełna referencja: `.ai/GRAPHICS_CONVERSION.md`.
### Problem
- Reader emitował EMF/WMF media parts jako `data:image/x-emf|x-wmf` → przeglądarka nie renderuje → złamany obraz. VML kształty wektorowe nieobsługiwane.
### Architektura (DDD)
- `IGraphicConversionService` (Domain/Interfaces) + modele `GraphicSource`/`WebGraphicRepresentation`/`GraphicConversionResult`/`GraphicConversionDiagnostics` + enumy (Domain/Models).
- `GraphicConversionService` (Infrastructure) — **pure-managed**, bez LibreOffice/GDI/System.Drawing → Linux/GCP-safe. Detekcja (magic bytes + content-type), wymiary z nagłówków (PNG/JPEG/GIF/BMP/EMF rclFrame/WMF placeable), ekstrakcja osadzonego rastra z EMF/EMF+, placeholder SVG, sanitizer SVG, VML rect/oval/line/roundrect→SVG.
- Hook w `DocxToHtmlConverter.WebGraphicForLegacy` — EMF/WMF (a:blip i v:imagedata) → renderowalny `data:URL` (osadzony raster albo placeholder), web-native bez zmian. Atrybut `data-legacy-graphic="placeholder"`.
### Strategia
- EMF/WMF **nie rasteryzowane** (brak bezpiecznej pure-managed ścieżki na Linux) → placeholder + **pass-through oryginalnego partu** (`ConvertPreservingPackage`) — Word renderuje prawdziwą grafikę, dokument nietknięty. Honest fallback, nie udawanie.
- Bezpieczeństwo: `DtdProcessing.Prohibit`/`XmlResolver=null` (XXE blok), SVG sanitizer (script/on*/external/javascript), limity rozmiaru/czasu/wymiarów + cancellation; wyjątki → Fallback (nie wywraca importu).
### Decyzja: zero nowych zależności graficznych
- LibreOffice zakazane; System.Drawing Windows-only; brak permisywnego pure-managed renderera EMF. Rasteryzacja przez out-of-process sidecar = roadmapa.
### USUNIĘTO poprzednią ścieżkę EMF opartą o soffice + System.Drawing (krytyczne dla GCP)
- Reader miał `TryConvertMetafileToPng` → **LibreOffice `soffice`** (zakazane) + fallback **`System.Drawing`** (Windows-only, `PlatformNotSupported` na Linux) → działało tylko na Windows, **nie na GCP**. Usunięto wszystkie te metody + **pakiet `System.Drawing.Common` z .csproj**; oba wywołania (`LoadImageFromPart`, picture-bullet) przepięto na pure-managed `GraphicConversionService`. `SkiaSharp`(+NativeAssets.Linux) zostaje (GCP-safe, tylko barcody). Weryfikacja: `grep System.Drawing|soffice` w .cs = pusto; build OK; Infrastructure 137/137.
### Verified / tests
- Infrastructure **137** (+21: `GraphicConversionServiceTests` 13, `GraphicConversionSecurityTests` 7, `GraphicConversionIntegrationTests` 1 — DOCX z EMF a:blip → reader emituje svg+xml, **nigdy** image/x-emf). Brak regresji istniejących testów obrazów.
- Benchmark `GraphicConversionBenchmarks` (BenchmarkDotNet `[MemoryDiagnoser]`, param GraphicCount 1/10/100) — `dotnet run -c Release -- --filter *GraphicConversion*`.
### Ograniczenia
- EMF/WMF bez podglądu wektorowego w przeglądarce (placeholder; pełny render w Word). VML: tylko bezpieczny podzbiór kształtów; hook kształtów wektorowych do readera + SVG→DOCX PNG-fallback + cache = roadmapa.

## 2026-06-04 — Edytor: 8 błędów (ENTER, font-size, font-family, paste plain, interlinia, link, .doc, Backspace/strzałki)
Pełna referencja: `.ai/EDITOR_KEYBOARD.md`.
### Issue 1 — ENTER cofał kursor do poprzedniej linii (FIX)
- Root cause: po wpisaniu znaku edytor repaginuje (~600 ms) i przebudowuje DOM stron; karetkę odtwarzał **globalny offset tekstowy**, niejednoznaczny na granicy bloków — nowy **pusty** akapit po ENTER ma 0 znaków → restore lądował na końcu poprzedniego akapitu.
- Fix: kotwica karetki **`{ block, offset }`** (`_saveGlobalCaret`/`_restoreGlobalCaret` + współdzielony `_flattenTopBlocks`, `_placeCaretAtTextOffset`). Pusty blok → karetka na początku bloku. Indeks bloku stabilny między repaginacjami.
### Issue 2 — font-size przez input ukrywał tekst (FIX)
- Root cause: wartość `0`/NaN dawała `0pt`/`NaNpt` → tekst niewidoczny. Toolbar walidował, ale publiczne `setFontSize` nie.
- Fix: `setFontSize` odrzuca `!finite`/`<1`/`>400` (defense-in-depth).
### Issue 3 — font-family „nie działała"/wracała do domyślnej (FIX, root cause backend)
- Root cause: `innerHTML` serializuje nazwy wielowyrazowe jako encję `&quot;`; writer (`ApplyRunStyle`) regex `[^,;]+` ucinał nazwę na `;` **wewnątrz `&quot;`** → `w:rFonts ascii="&quot"` → Word wracał do domyślnego fontu. (Single-word jak Arial działały.)
- Fix: **HTML-decode stylu** przed regexem + odcięcie cudzysłowów i fallbacku generycznego. `'Font',serif` / `"Font"` / `Arial` → czysta nazwa.
### Issue 4 — „Wklej bez formatowania" nie działało (FIX)
- Root cause: klik w pozycję menu zabierał fokus edytorowi → `execCommand('insertText')` bez karetki nic nie wstawiał.
- Fix: `insertText` odtwarza zapisaną selekcję + fokus przed wstawieniem; `pasteWithoutFormatting` dostał `.catch`. Ctrl+Shift+V działał wcześniej.
### Issue 5 — interlinia względem Word (ANALIZA + regresja)
- Ustalenie: mapowanie OOXML→CSS jest **poprawne** (`auto`→`line/240` bezjednostkowe; `exact`/`atLeast`→`pt`; before/after→margin pt; bez podwójnego liczenia). Pinned testami `LineSpacingMappingTests`.
- Ograniczenia nieusuwalne mapowaniem (HTML/CSS): „single" ≠ `line-height:1` (Word dolicza line-gap fontu); kolaps sąsiednich marginesów vs sumowanie before+after w Word. Udokumentowane.
### Issue 6 — wstawianie linku nie działało (FIX)
- Root cause: pole URL dialogu zabierało fokus → selekcja kolabowała → `createLink` bez celu.
- Fix: `insertLink` odtwarza zapisaną selekcję, **normalizuje URL** (`normalizeLinkUrl`), escapuje etykietę; zwinięta karetka → `<a … rel=noopener>`. Writer już emituje `w:hyperlink`+relację.
### Issue 7 — brak obsługi `.doc` (DECYZJA: Wariant B — jawne odrzucenie)
- `.doc` = stary binarny format (OLE/CFBF), nie OOXML; brak konwertera DOC→DOCX w pipeline. Odrzucenie z instrukcją konwersji zamiast udawania obsługi: backend `OpenDocument` (`.doc`→400) + frontend `dashboard.openFile`. Wariant A (LibreOffice headless) opisany jako przyszłość.
### Issue 8 — Backspace nie usuwał stron; strzałki „skakały" (FIX częściowy)
- Backspace na **początku strony** usuwa **manualny page-break** poprzedniej strony (`_tryDeletePageBreakBackwards` + `_isCaretAtEditorStart`/`_removeTrailingPageBreak`) i repaginuje; **nie** rusza wrappera strony.
- Nawigacja ArrowDown/Up między stronami była dodana wcześniej (`_tryMoveCaretAcrossPages`); „skakanie" złagodzone przez block-aware karetkę (Issue 1). Detekcja skrajnej linii (`_isCaretOnEdgeLine`) jest layout-zależna — pozostaje obszar do dostrojenia (manualnie).
### Verified / tests
- Backend: Infrastructure **116** (+3 `FontFamilyWriteTests`, +6 `LineSpacingMappingTests`), Api **34** (+1 `.doc`), Application **193**, Domain **57** — zielone.
- GUI: **197** (+3 `enter-caret`, +6 `toolbar-actions` font-size/link/paste, +3 Backspace page-break w `caret-field`).
- jsdom nie pokrywa contenteditable/execCommand/layout → wstawianie linku/tekstu i repaginacja end-to-end mają **scenariusze manualne** (`EDITOR_KEYBOARD.md` §7). Import/eksport/autozapis/undo/tabele/obrazy nietknięte (te same testy zielone).
### Ograniczenia
- Block-index karetki: rzadki edge gdy tabela **przed** karetką zmienia podział w danym przebiegu. Interlinia: różnice metryk fontu i kolapsu marginesów (jw.). `.doc`: brak realnej konwersji (świadomie). Strzałki: edge-line detection layout-zależna.

## 2026-06-04 — Layout: banner środowiska spychał dashboard w dół (fix shell `100vh`→`100%`)
### Diagnoza
- App shell (`d2-root`, global `styles.scss`) jest flex-column `height:100vh; overflow:hidden`; globalne bannery (`d2-global-banners` = environment + offline) mają `flex:0 0 auto` (intrinsic height), a obszar strony flexuje resztę: `d2-root > d2-dashboard/…-editor/…-pdf-viewer/…-pdf-maintenance/…-admin-shell { flex:1 1 0; min-height:0; overflow:hidden }`. **Banner jest w normalnym flow i to jest poprawne** — host strony dostaje `100vh − banner`.
- **Edytor (wzorzec OK):** `:host{height:100%}` → `.document-editor-container{height:100%}` — wypełnia obszar przydzielony przez shell, banner uwzględniony automatycznie.
- **Bug:** `dashboard.scss .dashboard-wrapper{min-height:100vh}` oraz `pdf-maintenance .maintenance-wrapper{min-height:100vh}` wymuszały **pełną wysokość viewportu** wewnątrz hosta wysokiego na `100vh − banner` → wrapper wystawał o wysokość bannera, host (`overflow-y:auto`) scrollował, treść wizualnie zjeżdżała w dół, stopka ucinana. Klasyczny „dashboard obniżony o wysokość labela".
### Changed (frontend)
- `dashboard.scss`: `.dashboard-wrapper` `min-height:100vh` → **`min-height:100%`** (wypełnia host przydzielony przez shell, nie viewport). Komentarz wyjaśniający.
- `pdf-maintenance.ts` (inline styles): `.maintenance-wrapper` `min-height:100vh` → **`min-height:100%` + `height:100%`** (centrowanie w obszarze hosta, nie viewportu).
- **Bez** magicznych marginesów, bez zmiennej `--banner-height`, bez `!important`, bez globalnego CSS — fix to ujednolicenie do banner-agnostycznego idiomu `100%-of-host`, którego shell już używa dla edytora. `app.scss .app-container{min-height:100vh}` to martwy CSS (brak `styleUrl` w `app.ts`, inline template bez tej klasy) — pozostawiony bez zmian.
### Verified / tests
- Nowy `pages/layout-shell.spec.ts` (2): dashboard i pdf-maintenance — wstrzyknięty scoped CSS wrappera **nie zawiera `100vh`** i ma `min-height:100%` (kontrakt layoutu; jsdom nie robi layoutu, więc asercja na faktycznie zregresowanym CSS, nie na pikselach). GUI **185** (było 183, +2). Edytor nietknięty (te same testy zielone).
### Scenariusze manualne
- Dashboard z bannerem DEV: „Witaj w Qutas" wyśrodkowane, stopka `© 2026 ING` widoczna, brak scrolla/skoku. Bez bannera (PROD): identyczny układ. Edytor (screen A): toolbar/panele/stopka bez zmian. Środowiska Local/DEV/TST/PRE: różny kolor bannera, **ta sama wysokość** layoutu. Małe rozdzielczości: gdy treść > obszar, host scrolluje (nic nie ucięte).
### Ograniczenia
- `min-height:100%` wymaga definite-height hosta — zapewnia go shell (`flex:1 1 0` w `d2-root height:100vh`); ten sam mechanizm, na którym opiera się edytor.

## 2026-06-04 — 3 poprawki: font fallback (serif), nawigacja kursora między stronami, rozmiar numeru strony w stopce
### P1 — font fallback wg rodziny (serif/mono/sans) zamiast twardego `sans-serif`
- Diagnoza: reader emitował każdy font jako `font-family:'X',sans-serif`. Gdy `Times New Roman` niedostępny w przeglądarce, przeglądarka spadała na **bezszeryfowy** default (wygląd jak Calibri), mimo że nazwa była zachowana.
- Changed (reader `DocxToHtmlConverter`): `FontFamilyCss(name)` + `GenericFontFallback(name)` — szeryfowe (times/cambria/georgia/garamond/minion/palatino/book antiqua/„serif") → `serif`, courier/consolas/mono → `monospace`, reszta → `sans-serif`. Zastąpiono 3 miejsca emisji `,sans-serif`. **Bez** hardcode konkretnego fontu, bez dosadzania plików czcionek, bez zmiany fontu dokumentu.
### P2 — ArrowDown/ArrowUp nie przechodził między stronami (osobne contenteditable per strona)
- Changed (frontend `wysiwyg-editor`): w `handleKeyboard` (przed blokiem Ctrl) gdy ArrowDown/ArrowUp bez modyfikatorów i `_tryMoveCaretAcrossPages(dir)` zwróci true → `preventDefault`. Nowe: `_tryMoveCaretAcrossPages` (zwinięta karetka na skrajnej linii → przeniesienie do sąsiedniej strony), `_isCaretOnEdgeLine` (porównanie rectów karetki z górną/dolną linią edytora), `_placeCaretAtEditorEdge` (focus + range na start/end + aktywacja strony). Shift/Ctrl/Alt i zaznaczenia (range) nietknięte (guard `isCollapsed` + brak modyfikatorów).
### P3 — numer strony w stopce ignorował rozmiar fontu stopki
- Diagnoza: pole PAGE było emitowane „goło", numer dziedziczył rozmiar kontenera (np. 10.5pt) zamiast runów stopki (np. 8pt).
- Changed (reader): `FieldSpan(cssClass, placeholder, run)` emituje `<span class=… style="{run rPr CSS}">` dla PAGE/NUMPAGES/DATE (simple + complex field) — numer niesie font-size własnego runu.
- Changed (writer): `BuildFieldRun` buduje poprawny `w:fldSimple` z **wewnętrznym runem** niosącym `rPr` (font-size) — zamiast wcześniejszego (schema-niepoprawnego) `SimpleField` wewnątrz `Run`, który po round-tripie gubił właściwości. `FlushPending` w stopce/nagłówku uznaje też paragraf zawierający `SimpleField` (wcześniej wymagał `Run`/`Hyperlink` → pole było pomijane).
- Changed (frontend): `_pageNumberHtml(editor)` + `_inlineFieldFontSize(editor)` — wstawiany `.page-number` dostaje font-size pierwszego inline-sized spanu stopki (fallback computed). CSS `.field-page/.field-numpages/.field-date/.page-number { font-size:inherit; … }`.
### Verified / tests
- Backend: Infrastructure **107** (+4: `DefaultStyleFontTests` serif/sans fallback; `PageFieldFontTests` reader emituje font-size pola + writer round-trip `w:fldSimple` z rPr=16). Application **193**, Domain **57** — zielone. (Api.UnitTests pominięte: lock DLL od działającej instancji API/debuggera — nie dotyczy zmienionego kodu Infrastructure.)
- GUI: **183** (+7: `wysiwyg-editor.caret-field.spec` — ArrowDown/Up przechodzi między stronami, range/Shift nie rusza karetki, brak ruchu poza ostatnią/przy jednej stronie, `_pageNumberHtml`/`_inlineFieldFontSize` dziedziczą font-size).
### Ograniczenia
- `_isCaretOnEdgeLine` jest layout-zależne (rect karetki) — nieodtwarzalne w jsdom (brak `Range.getClientRects`); w testach stubowane, logika nawigacji/granic testowana realnie.
- Font fallback poprawia rendering, ale nie podstawia brakującego kroju — przy braku Times New Roman tekst renderuje się szeryfowym fontem systemowym (świadomie, bez licencjonowania krojów).

## 2026-06-03 — Font treści (Times New Roman→Calibri) + page break „od nowej strony" (display)
### Diagnoza — font
- Oryginał `orginał_GOOD`: domyślny styl akapitu **`Normalny` (`w:default="1"`)** ma `<w:rFonts w:ascii="Times New Roman"/>` (sz=21 → 10.5pt). To **nadpisuje** docDefaults (`asciiTheme=minorHAnsi` → theme minor = Cambria). Runy body mają tylko `w:cs="Times New Roman"` (complex-script), bez `ascii`.
- Bug: reader `LoadDocDefaults` czytał **wyłącznie docDefaults** (→ Cambria), ignorował font domyślnego stylu akapitu → kontener `.document-content` dostawał Cambria, a edytor i tak spadał na własny default (Calibri). Stąd Calibri na ekranie.
### Diagnoza — page break
- To manualny `<w:br w:type="page"/>` we własnym akapicie (po tabeli z podpisami, przed „PROTOKÓŁ…"). Round-trip był już naprawiony (R-15: writer → `w:br`; reader → top-level `<div class="page-break">`). **Pozostał problem DISPLAY**: edytor `_splitHtmlIntoPages` **konsumował** marker przy podziale, więc `_repaginateNow` (re-paginacja wg wysokości) go nie widziała i scalała treść → „PROTOKÓŁ" pod podpisami.
### Changed — backend (reader)
- `DocxToHtmlConverter.ApplyDefaultParagraphStyleFont`: po wczytaniu stylów czyta font domyślnego stylu akapitu (`w:default="1"`, typ paragraph) i ustawia go jako `_defaultFontFamily`/rozmiar kontenera (override docDefaults). Z dokumentu, nie hardcode. **Weryfikacja na realnym pliku**: kontener `font-family:'Times New Roman',sans-serif;font-size:10.5pt`.
### Changed — frontend (display)
- `wysiwyg-editor._splitHtmlIntoPages`: **zachowuje** marker `<div class="page-break">` (doklejony na końcu każdej strony poza ostatnią) → `_repaginateNow` honoruje podział (`_isPageBreakBlock` wymusza nową stronę), marker przeżywa zapis (getContent → writer → `w:br`). `setContent` join plain (bez dublowania); usunięty martwy `_joinPagesWithBreaks`.
### Verified / tests
- Backend **386 pass** (+4 `DefaultStyleFontTests`: default-style override, docDefaults fallback, direct docDefaults, pass-through zachowuje font). GUI **176** (+1 `_splitHtmlIntoPages` zachowuje marker + round-trip). Golden snapshoty bez zmian (syntetyki nie mają default-style fontu). Solucja zielona.
### Ograniczenia
- Font: jeśli `Times New Roman` niedostępny w przeglądarce → fallback `,sans-serif` (nazwa zachowana w CSS; rendering zależny od systemu). Toolbar edytora może pokazywać własny default fontu (kosmetyka) — renderowany tekst używa fontu kontenera.
- Save font: pełna wierność na zapisie zależy od pass-through (R-16); ścieżka autosave bezstanowa (R-20) wciąż regeneruje style.
- Page break: honorowane są **manualne** breaki; naturalna paginacja Worda nadal liczona wg wysokości (nie 1:1).

## 2026-06-03 — Benchmarki + fidelity-checks (perf/pamięć/regresja wierności DOCX)
### Added — backend (`D2ViewerEditor.Benchmarks`, BenchmarkDotNet 0.14)
- `BenchmarkAssets` — syntetyczne DOCX in-memory (simple/tables/large-table/images/header-footer/page-breaks/multi-style) + opcjonalny realny plik regresyjny przez env `D2_BENCH_ASSETS` (domyślnie repo; używany tylko gdy istnieje). Brak commitowania wrażliwych plików.
- `DocxImportBenchmarks`/`DocxExportBenchmarks`/`DocxRoundTripBenchmarks`/`TableBenchmarks` — `[MemoryDiagnoser]`, baseline per grupa, params Rows=100/400 (`[ShortRunJob]` na large-table). Mierzą reader/writer/`ConvertPreservingPackage`/round-trip.
- `DocxPackageReport` (analizator pakietu ZIP/regex) + `FidelityComparison` (PASS/WARN/FAIL, pass-through-aware) + `FidelityReportRunner` (markdown, report-only; `--fidelity [--fail-on-regression]` w `Program.cs`).
### Added — fidelity gate (NUnit, CI)
- `DocxFidelityRegressionTests` (4): R-16 styles/theme zachowane (pass-through), R-15 page breaki, R-17 tabele niemnożone, R-18 brak sztucznych `trHeight`. Parser-niezależny (ZIP/regex).
### Added — frontend perf harness (Vitest, report-only)
- `wysiwyg-editor.perf.spec.ts` (4): timing `getContent`/merge split-table/`setContent` (Performance API, lenient ceiling 4000 ms). Ograniczenie: jsdom nie mierzy layoutu/paginacji (potrzebny Playwright).
### Verified
- Backend **382 pass** (+4 fidelity gate). GUI **175** (+4 perf). Benchmarki uruchamiają się (dry-job: import ~85 ms cold). Fidelity report: syntetyki **PASS**, realny `orginał_GOOD` (pass-through) **WARN** (164 style + Cambria + 6 tabel + 1 page break zachowane; tylko `numbering.xml` nieprzeniesiony — R-19).
### Docs
- Nowy `.ai/BENCHMARKS.md` (jak uruchomić perf/fidelity/frontend, progi, baseline, CI report-only→fail-on-regression, ograniczenia) + `INDEX`.

## 2026-06-03 — Page break „od nowej strony": edytor nie honorował manualnego podziału (display)
### Problem
Manualny page break z DOCX (np. przed „PROTOKÓŁ WYDANIA POJAZDU") nie zaczynał nowej strony w edytorze — treść lądowała tuż pod poprzednią. Round-trip zapisu działał (R-15), ale **paginacja widoku** ignorowała break.
### Root cause
Reader emitował break-only akapit (`<w:p><w:r><w:br type=page/></w:r></w:p>`) jako **zagnieżdżony** `<p><span><div class="page-break"></div></span></p>`. To: (1) psuło `_splitHtmlIntoPages` (regex dzielił string w środku `<p>`), (2) `_repaginateNow` paginował tylko wg wysokości — ignorował break.
### Changed
- **Reader** `DocxToHtmlConverter`: nowy `IsPageBreakOnlyParagraph` — akapit zawierający wyłącznie `w:br type=page` (bez tekstu/grafiki) emitowany jako **top-level** `<div class="page-break"></div>` (zamiast zagnieżdżony). Writer nadal mapuje go na `w:br type=page` (R-15).
- **Frontend** `wysiwyg-editor.ts`: `_repaginateNow` wymusza **nową stronę** na bloku page-break (`_isPageBreakBlock` — top-level lub zagnieżdżony bez tekstu); marker zostaje w treści → przeżywa zapis.
### Verified / tests
- Backend **378 pass** (+1 reader: break-only akapit → top-level block + round-trip=1). GUI **171** (+1 `_isPageBreakBlock`). Golden snapshoty bez zmian (brak page-breaków). Solucja zielona.
### Uwaga
- Wymaga **ponownego wczytania dokumentu** w edytorze, by zobaczyć efekt (paginacja liczona przy load/edycji).

## 2026-06-03 — R-15/R-16/R-17: page break round-trip, pass-through pakietu DOCX, scalanie split-table
### R-15 — manualny page break round-trip (writer)
- `HtmlToDocxConverter`: nowy `IsPageBreakNode` (class `page-break` / `data-docx-break="page"` / CSS `page-break-before`/`break-before:page`). `CreateRunsFromNode` konwertuje **zagnieżdżony** `<div class="page-break">` (reader emituje go wewnątrz akapitu) na `w:br type=page`; blok-level używa tej samej detekcji. Wcześniej break ginął (AFTER=0). Bez duplikacji; dokument bez breaków nie dostaje żadnego. Testy: `PageBreakRoundTripTests` (5).
### R-16 — pass-through oryginalnego pakietu (writer + handler)
- `IHtmlToDocxConverter.ConvertPreservingPackage(html, Stream? original, ...)`: generuje DOCX jak `Convert`, po czym **zachowuje z oryginału** `styles.xml` (pełny zestaw, w tym style tabel), `theme` i `fontTable` (FeedData do istniejących partów). Body/sekcja/nagłówki-stopki/obrazy/numbering pochodzą z konwersji HTML. `null`/pusty oryginał → zachowuje się jak `Convert` (fallback). Best-effort: błąd oryginału → fallback (nie psuje zapisu).
- Wiring: `DownloadEditedDocumentCommandHandler` ładuje **bazową wersję (v1)** przez `IDocumentStorageService.DownloadAsync` i używa pass-through dla DOCX; fallback gdy brak wersji/nie-DOCX. Testy: `PassThroughPackageTests` (5) + handler test pass-through (1).
### R-17 — scalanie fragmentów split-table (frontend)
- `wysiwyg-editor.ts`: `_splitTableForPagination` taguje fragmenty jednej logicznej tabeli `data-split-table-id` + klonuje `colgroup`; `getContent()` woła `_mergeSplitTables` (scala sąsiednie fragmenty o tym samym id, zachowuje wiersze/kolumny, sprząta marker). Niezależne sąsiednie tabele NIE są scalane. Testy: Vitest (2).
### Verified (realny plik: reader → ConvertPreservingPackage(html, oryginał GOOD) → AFTER)
| metryka | GOOD | BAD | AFTER |
|---|---|---|---|
| styles.xml liczba stylów | 164 | 16 | **164** (part `styles2.xml`, relacja+content-type OK — Word czyta po relacji) |
| theme minor font | Cambria | Calibri | **Cambria** |
| twarde page-breaki | 1 | 3 | **1** |
| marginesy / tabele | wzorzec | rozjechane | ≈ GOOD (z poprz. fixów) |
### Tests
- Backend **377 pass / 0 fail / 6 skip** (Infrastructure +10: 5 page-break + 5 pass-through; Application +1). GUI **170** (+2 split-table). Pełna solucja zielona.
### Ograniczenia (R-19..R-21)
- **R-16 niepełny pass-through:** zachowane tylko `styles.xml`/`theme`/`fontTable`; `numbering.xml` NIE jest przenoszony (uniknięcie konfliktu numId z generowanym body) — dla dokumentów z listami custom-bullety mogą się różnić. Part stylów zapisywany jako `styles2.xml` (artefakt SDK; relacja poprawna, Word czyta). **Wiring tylko w DownloadEditedDocument** — ścieżka autosave (`/api/document/save`, bezstanowa) nadal generuje od zera; pełne wpięcie wymaga master-aware endpointu konwersji (next).
- **R-16 model body:** body nadal z konwersji HTML (inline style), nie referencje oryginalnych nazwanych stylów; font dziedziczy z zachowanego docDefaults (Cambria), ale akapity nie wracają do oryginalnych `w:pStyle`.

## 2026-06-03 — Round-trip fidelity: naprawa rozjazdu 4→7 stron i powiększonych tabel (orginał_GOOD vs zapisany_BAD)
### Root causes (z analizy XML dwóch plików)
- **Marginesy napompowane**: writer `AddPageSettings` `Math.Max(topTwips, headerHeightTwips+720)` + hardkod `Header/Footer=720` → top 567→1281, bottom 851→1231 (mniej treści/stronę).
- **Tabele wyższe**: reader `ConvertTableToHtml` domyślny `padding:4px 8px` (4px=60 tw góra/dół) do każdej komórki → BAD +8160 tw (14.4 cm) `tcMar`. Word default = 0 góra/dół.
- **Tabele szersze**: writer domyślnie `tblW pct 5000` (100%) gdy brak/auto width → rozciągnięcie tabel content-sized do pełnej szerokości. + hardkod `tblCellMar 40/80`.
- **+2 twarde page-breaki (4→7)**: edytor `getContent()` materializował auto-paginację wg wysokości (`_repaginateNow`) jako `<div class="page-break">` między każdą stroną → twarde `<w:br type=page>` w DOCX.
### Changed — backend (`HtmlToDocxConverter`/`DocxToHtmlConverter`)
- Reader: domyślny padding komórek = Word default (top/bottom **0**, left/right **108 tw**) zamiast `4px 8px`; fallbacki `tcMar` → 0/108.
- Writer: `tblW` domyślnie **auto** (`{Width="0",Type=Auto}`) zamiast `pct 5000`; tabelowy `tblCellMar` → **0/108/0/108** (Word default) zamiast 40/80.
- Writer `AddPageSettings`: marginesy zapisywane **jak autorskie** (bez `Math.Max(...+720)`); `Header/Footer` distance **rekonstruowane** = `clamp(margin − bandHeight, 0, 720)` (odwrotność wyliczenia pasma przez reader).
### Changed — frontend (`wysiwyg-editor.ts`)
- `getContent()` (ścieżka ZAPISU) scala strony **czystą konkatenacją** (bez `<div class="page-break">`). Akapity są block-atomic → konkatenacja odtwarza treść; jawne page-breaki użytkownika przeżywają jako div w treści.
### Verified (realny plik orginał_GOOD.docx → reader→writer → AFTER)
| metryka | GOOD | BAD | AFTER |
|---|---|---|---|
| pgMar top / bottom | 567 / 851 | 1281 / 1231 | **567 / 850** |
| pgMar header / footer | 6 / 340 | 720 / 720 | **6 / 339** |
| tblW (6 tabel) | auto | pct 5000 | **auto** |
| suma tcMar góra+dół | 0 | 8160 tw | **0** |
| twarde page-breaki | 1 | 3 | **0** |
### Tests
- Backend **366 pass / 0 fail / 6 skip** (+`RoundTripLayoutFidelityTests` 4: marginesy nie-inflowane, header/footer distance, tcMar=0, tblW auto; +regeneracja 2 golden table snapshotów na `padding:0px 7px`). GUI **168** (+2 `getContent` bez page-break).
- Weryfikacja na realnym pliku przez tymczasowy test lokalny (usunięty, niecommitowany) → `zapisany_AFTER.docx` (artefakt do ręcznego porównania w Word).
### Notes / ograniczenia (patrz R-15..R-18)
- AFTER ma 0 twardych breaków (vs GOOD 1) — reader→writer nie przywraca pojedynczego oryginalnego page-breaka (osobny gap writera). Po naprawie marginesów/tabel Word i tak paginuje ~4 strony naturalnie.
- Tabela dzielona przez edytor między strony (`_splitTableForPagination`) zapisuje się jako 2 sąsiednie `<table>` (pre-existing) — do domknięcia (scalanie przy zapisie).
- Font treści Cambria→Calibri, minimalny `styles.xml`, brak `numbering.xml` — model `DocumentContent` stratny (R-16); nie wpływa na liczbę stron tego dokumentu (`numPr`=0), ale na fidelity.

## 2026-06-03 — Dokumentacja: nowy `DOCX_CONVERSION.md` (reference techniczny konwersji) + cross-linki
### Changed
- Nowy `.ai/DOCX_CONVERSION.md` — zweryfikowany w kodzie reference konwersji DOCX↔HTML: pipeline (2 diagramy Mermaid), model `DocumentContent` (diagram klas + ograniczenie R-10), `OoxmlUnits`, **macierz statusów** wszystkich obszarów (Implemented/Partially/Planned z kierunkiem R/W i ścieżkami w kodzie), roadmapa (R-10 multi-section, Etap 3 computed-style, domknięcia 5/6/7), harness testów, ograniczenia trwałe.
- `INDEX.md` + wiersz dla `DOCX_CONVERSION.md`. Cross-linki z `FIDELITY_REPORT.md` i `FEATURES.md`.
### Verified (przeciw kodowi, bez zmian kodu)
- Multi-section: reader/writer `Body.Elements<SectionProperties>().FirstOrDefault()` → tylko pierwsza sekcja (R-10 = Partially).
- Computed-style: `basedOn`+`rStyle`+theme+docDefaults(font/size) działają; brak centralnego resolvera/numbering/linked folding (Partially; pełny = Planned).
- Rotacja obrazów `a:xfrm/@rot`: brak w obu konwerterach (Planned). VML: `ConvertPictureToHtml` tylko obraz+rozmiar (Partially).
- `tcMar`: czytane → CSS padding (Implemented). `tblCellSpacing`: nieczytane (Planned).
- `w:tabs` na zapisie: emitowane **tylko w stylach Header/Footer** (`AddDocumentStyles`, pozycje 4536/9072 tw); per-akapit/custom nieprzenoszone (Partially).
### Notes
- Zadanie dokumentacyjne — zero zmian kodu; testy bez zmian (backend 358, GUI 166).

## 2026-06-03 — Refaktor DOCX→HTML Etap 8: audyt izolacji CSS dokumentu (bez zmian kodu)
### Changed
- Brak zmian kodu — audyt. Ustalono, że treść dokumentu jest już izolowana: style inline wygrywają z regułami klasowymi przez specyficzność CSS. Faktyczna liczba `!important` w `wysiwyg-editor.scss` to **6** (nie 65 — wcześniejsza liczba była artefaktem zbiorczego grepu wielu wzorców), z czego tylko 1 dotyczy treści (świadoma zamiana Calibri Light→Calibri w nagłówkach). `!important` w `styles.scss`/`document-editor.scss` dotyczą podświetleń UI i paddingów dialogów — nie nadpisują font/koloru/marginesów treści.
### Verified
- `FIDELITY_REPORT.md` skorygowany (§3 wysoka wierność: izolacja CSS; §5: Etap 8 = zweryfikowane bez zmian). Backend **358 pass**, GUI **166 pass** bez zmian (audyt nie ruszał kodu).
### Notes
- Premisa planu „65 !important nadpisuje treść" była błędna (miscount). Ryzykowna chirurgia SCSS niepotrzebna. Ewentualny dalszy krok: defensywny scoped reset + test computed-style (ograniczony w jsdom).

## 2026-06-03 — Refaktor DOCX→HTML Etap 7: tab-stopy nagłówka/stopki (układ lewo⇥środek⇥prawo)
### Changed
- `DocxToHtmlConverter`: akapit z tab-stopem Center lub Right/End (`ParagraphHasAlignmentTab`) renderowany jako `display:flex` (`align-items:baseline;width:100%`). Run zawierający wyłącznie `<w:tab/>` (przy aktywnym `_flexTabs`) → `<span style="flex:1 1 0;">\t</span>` (rosnący spacer, **bezpośrednie dziecko flexa**), więc segmenty rozkładają się na szerokości zamiast zlewać w stałe odstępy. Znak taba zachowany (round-trip nienaruszony). Mieszane runy bez zmian.
### Verified
- `Infrastructure.UnitTests` **80/80 pass** (+1 golden `TabStopLeftCenterRight`: flex + 2 spacery + L/C/R tekst). Pozostałe golden niezmienione. Build OK.
### Notes
- Środek nie zawsze idealnie wycentrowany (zależy od szerokości boków) — udokumentowane jako częściowa wierność (FIDELITY_REPORT §3). Writer nie emituje `w:tabs` → po zapisie tab-stopy znikają i układ wraca do tabów (bez regresji vs. stan sprzed). Pozostaje Etap 7: wiele sekcji (R-10).

## 2026-06-03 — Refaktor DOCX→HTML Etap 6: tekst alternatywny obrazów (alt ↔ wp:docPr/@descr)
### Changed
- Reader `DocxToHtmlConverter.ConvertDrawingToHtml`: odczyt `wp:docPr/@descr` (fallback `@title`) → `<img alt="...">` (tylko gdy niepuste → istniejące obrazy bez alt bez zmian). Escape przez `EscapeHtml`.
- Writer `HtmlToDocxConverter`: `<img alt>` → `wp:docPr/@descr` w obu ścieżkach (inline + anchor) przez helper `BuildImageDocProperties`; `HtmlEntity.DeEntitize` na alt (inaczej encje podwójnie się escape'ują przy ponownym odczycie).
### Verified
- `Infrastructure.UnitTests` **79/79 pass** (+3 `ImageAltTextRoundTripTests`: round-trip, znaki specjalne/escape, brak alt → brak atrybutu). Golden-snapshoty niezmienione. Build OK.
### Notes
- Obrazy już wcześniej round-tripowały floating/border/crop/rozmiar (EMU). Pozostaje w Etapie 6: rotacja (`a:xfrm/@rot`), skalowanie przy crop (clip-path nie skaluje), VML (legacy) border/crop/alt.

## 2026-06-03 — Refaktor DOCX→HTML Etap 4 (full-stack): rozmiar strony + orientacja — pełny round-trip
### Changed
- Backend wiring: `PageSize` przewleczony przez `SaveDocumentCommand`/`SignDocumentCommand`/`DownloadEditedDocumentCommand` (+ handlery → `_converter.Convert(... request.PageSize)`) i 3 kontrolery (`DocumentController.save/sign`, `DocumentStorageController.user-download`). DTO (`SaveDocumentRequest`/`SignDocumentRequest`) już niosły pole z Etapu 4 backend.
- Frontend: `document.model.ts` — interfejs `PageSize` + pole na `DocumentContent`/`SaveDocumentRequest`. `document-editor.ts` — sygnał `documentPageSize` ustawiany przy wczytaniu (`content.pageSize` → sygnał + `pageSettings.orientation`) i dołączany w `buildSaveRequest`. Round-trip pełny: DOCX → reader → `pageSize` → front → zapis → writer `w:pgSz`.
### Verified
- Backend **358 pass / 0 fail** (Domain 57, Application 192, Api 33, Infrastructure 76). GUI **166 pass** (+2 testy `PageSize round-trip` w `document-editor.spec`). Build OK.
### Notes
- R-14 **zamknięte** (pełny full-stack). Brak `pageSize` z frontu → backend fallback A4 (bez regresji).

## 2026-06-03 — Refaktor DOCX→HTML Etap 4 (backend): rozmiar strony + orientacja (round-trip)
### Changed
- Domena: nowy model `PageSize { WidthCm, HeightCm, Orientation }`; dodany do `DocumentContent`, `SaveDocumentRequest`, `SignDocumentRequest` (additive, backward-compatible).
- Reader `DocxToHtmlConverter`: `ExtractPageSize` z modelu `PageSettings` (Etap 2) → `DocumentContent.PageSize` (cm + portrait/landscape).
- Writer `HtmlToDocxConverter`: `IHtmlToDocxConverter.Convert` + `Convert` przyjmują opcjonalny `PageSize? pageSize`; `BuildPageSize` emituje `w:pgSz` (+`Orient` dla landscape) zamiast zahardkodowanego A4. Brak `pageSize` → fallback A4 portrait (bez regresji). Alias `OoxmlPageSize` rozwiązuje kolizję nazw `PageSize` (Domain vs OpenXml).
### Verified
- Cała solucja **358 pass / 0 fail / 6 skip** (Domain 57, Application 192, Api 33, Infrastructure 76 [+3 `PageSizeRoundTripTests`: landscape A4, A5 portrait, null→A4]; Integration 6 skip = wymaga Postgresa). Build OK.
- Zmiana interfejsu (opcjonalny param) wymusiła aktualizację mocków Moq w `DownloadEditedDocumentCommandHandlerTests` (CS0854: expression tree + optional arg) — dodany `It.IsAny<PageSize?>()`.
### Notes
- **Pozostaje (Etap 4 full-stack):** przewlec `PageSize` przez komendy/DTO (Save/Sign/Download) i front (`document.model.ts` + editor/ruler/scss + Vitest), żeby zapis faktycznie round-tripował rozmiar (teraz handlery nie przekazują `pageSize` → writer defaultuje A4; reader już zwraca `PageSize` w `DocumentContent` — front może konsumować). Patrz R-14.

## 2026-06-03 — Refaktor DOCX→HTML Etap 3: rozwiązywanie nazwanych stylów znakowych (w:rStyle)
### Changed
- `DocxToHtmlConverter.ConvertRunToHtml`: run z referencją do stylu znakowego (`w:rStyle`) dostaje teraz CSS tego stylu (z dziedziczeniem `basedOn`, z `_styles[rStyleId]`) **pod** formatowaniem bezpośrednim (direct wygrywa na konflikcie). Wcześniej runy formatowane wyłącznie przez styl znakowy (Hyperlink/Strong/Emphasis/własny) renderowały się **bez formatowania** — realna luka wierności.
### Verified
- `Infrastructure.UnitTests` **73/73 pass** (+1 golden `CharacterStyleRun`: bold+kolor+rozmiar ze stylu „Akcent" przez `rStyle`). Istniejące snapshoty **niezmienione** (brak `rStyle` w nich → zero dryfu). Build solucji OK.
### Notes
- Zakres celowy (slice): pełny computed-style (folding docDefaults do każdego elementu, style numerowania, linked styles, pełny theme) — kolejne kroki Etapu 3. Edge: jawne `bold=false` na runie nie wyłącza pogrubienia ze stylu znakowego (brak `font-weight:normal` z direct) — rzadkie, do domknięcia z computed-style.

## 2026-06-03 — Refaktor DOCX→HTML Etap 2: jawny model pośredni (read-side, strangler) — sekcja/strona
### Changed
- Nowy namespace `Infrastructure/DocxModel/`: `PageSettings` (geometria strony/sekcji jako surowe twipsy + page size + orientacja) i `SectionPropertiesReader.ReadPageSettings(sectPr)` (czysty parser, bez defaultów/konwersji).
- `DocxToHtmlConverter`: `ExtractPageMargins` + wysokość pasma nagłówka/stopki liczone teraz z modelu (`SectionPropertiesReader` + nowy helper `ComputeBandHeightCm`), zamiast bezpośredniego grzebania w `PageMargin`. Zachowanie **1:1** (te same wzory/zaokrąglenia). Page size/orientacja parsowane, ale jeszcze nieujawniane w HTML (fundament Etap 4).
### Verified
- `Infrastructure.UnitTests` **72/72 pass** (+5 `SectionPropertiesReaderTests` na modelu: size/margins/distances, landscape, brak pgMar, null sectPr, ujemny top margin). Golden-snapshoty **niezmienione** → potwierdzenie braku zmiany zachowania. Build solucji OK.
### Notes
- Strangler: stara (string) i nowa (model) ścieżka współistnieją; kolejne obszary migrują w Etap 3–7. ADR-0009.
- Kontrakt `DocumentContent`/API bez zmian.

## 2026-06-03 — Refaktor DOCX→HTML Etap 5: tabele (colgroup + fixed layout + fix vMerge/gridSpan)
### Changed
- `DocxToHtmlConverter.ConvertTableToHtml`: emisja `<colgroup><col style="width:..">` z `tblGrid` (autorytatywne szerokości kolumn) + `table-layout:fixed` gdy `tblLayout=fixed` lub tabela ma jawną szerokość (Dxa/Pct). Gdy fixed+brak szerokości tabeli → szerokość = suma kolumn z grid. Naprawia „tabela nie wygląda jak w Wordzie" (przeglądarka ignorowała geometrię kolumn). Nowe helpery `ReadTableGridColumnsPx`/`BuildColgroupHtml`.
- `DocxToHtmlConverter` (merge): naprawiony `rowspan` przy pionowym scaleniu sąsiadującym z `gridSpan` — `CountRowSpan` dopasowuje komórki wg **pozycji kolumny w gridzie** (z uwzględnieniem `gridSpan`), nie wg indeksu komórki (stary bug gubił rowspan). Skip komórki kontynuacji obsługuje teraz też jawny `vMerge val="continue"` (wcześniej tylko pominięty val). Nowe helpery `GetGridSpan`/`GetCellStartColumn`/`FindCellAtColumn`.
### Verified
- `Infrastructure.UnitTests` **67/67 pass** (+1 nowy `MergedCellsTable` golden; `simple-table` baseline zregenerowany — dodane colgroup/fixed). Snapshoty stabilne.
- Round-trip bezpieczny: writer (`HtmlToDocxConverter`) czyta wiersze przez `.//tr` i odtwarza własny `TableGrid` z liczby komórek → ignoruje `<colgroup>`/`<col>` (brak regresji zapisu).
### Notes
- Strona zapisu wciąż wymusza `TableLayout Autofit` — fidelity zapisu (fixed) odłożone (round-trip risk), do Etapu 4/5 cd.
- Pozostaje w Etap 5: `tcMar`/`tblCellMar` pełne, shading dziedziczony z `tblPr`, cellSpacing, border conflict resolution.

## 2026-06-03 — Refaktor DOCX→HTML Etap 0+1: centralne jednostki `OoxmlUnits` + harness regresji
### Changed
- Nowy `D2ViewerEditor.Infrastructure/Conversion/OoxmlUnits.cs` — jedno źródło prawdy dla konwersji OOXML (twips/EMU/half-points/cm/px) z nazwanymi stałymi (`TwipsPerInch`, `EmuPerInch`, `EmuPerPixel=9525`, `TwipsPerPoint=20`, `HalfPointsPerPoint=2`, `DefaultDpi=96`, `CmPerInch=2.54`).
- `DocxToHtmlConverter` i `HtmlToDocxConverter`: wszystkie rozproszone stałe konwersji (`567`, `1440`, `914400`, `9525`, `/20`, `*0.75`, `/2`, `*2.54`) zastąpione wywołaniami `OoxmlUnits`; usunięte zduplikowane prywatne helpery (`TwipsToPx`/`EmuToPx`/`PxToTwips`/`TwipsToCm`/`cmToTwips`). Lokalny rounding/truncation zachowany → zachowanie liczbowe niezmienione. cm liczone dokładnym `1440/2.54` (zastępuje przybliżenie `567`).
- Nowy harness regresji `Infrastructure.UnitTests/Golden/`: `GoldenDocuments` (deterministyczne DOCX w pamięci), `HtmlSnapshot` (normalizacja base64/`data-image-id`/daty + approve-on-missing), `GoldenSnapshotTests` (5 dokumentów: tekst, runy, spacing/indent, tabela, nagłówek+stopka z logo). Baseline w `Golden/__snapshots__/*.approved.html`.
### Verified
- `dotnet build D2ViewerEditor.sln` OK. `Infrastructure.UnitTests` **66/66 pass** (54 baseline + 7 `OoxmlUnitsTests` + 5 Golden). Snapshoty stabilne przy 2× uruchomieniu.
- Asercje punktowe potwierdzają wzory: sz=32→16pt, 240tw→12pt, 720tw→48px, 480tw→32px, 3000tw→200px, 2000tw→133px, EMU 1270000/317500 round-trip.
### Notes
- Kontrakt `DocumentContent`/`IDocxToHtmlConverter`/API bez zmian — front nietknięty. Round-trip (Header/Footer, ImageFloating/BorderCrop) zielony.
- Plan pełny: `~/.claude/plans/fancy-spinning-wirth.md` (Etap 0–9, strangler). Następny: Etap 2 (jawny model pośredni read-side) — duża zmiana, wymaga decyzji.
- Snapshoty lokują *obecne* zachowanie (regression-guard), nie poprawność wobec Worda; wierność podnoszona w Etap 3–8.

## 2026-05-28 — Import nagłówka/stopki: wybór wg referencji sekcji (default + first-page)
### Changed
- Backend `DocxToHtmlConverter`: `ExtractHeader`/`ExtractFooter` rozwiązują part przez `sectPr`/`HeaderReference`/`FooterReference` typu **Default** zamiast `HeaderParts.FirstOrDefault()` (kolejność partów była niezdefiniowana → mógł trafić pusty even/first). Dodano helpery `ResolveHeaderPart`/`ResolveFooterPart`/`HasTitlePage` oraz `ConvertHeaderPartToHtml`/`ConvertFooterPartToHtml`. Pierwsza strona (`titlePg`) → `DifferentFirstPage`+`FirstPageHtml` (model domeny już miał te pola). Fallback do `FirstOrDefault` gdy sekcja nie ma referencji.
- Frontend `document-editor.ts`: oba miejsca ładowania (`loadFromStorage` + ścieżka `openDocument`) spread'ują cały obiekt `content.header`/`content.footer`, zamiast rekonstruować tylko `{html,height}` → wariant first-page/odd-even nie jest gubiony na granicy TS.
### Verified
- Backend build OK; `dotnet test` filtr Header/Footer + SectionReference: **10/10 pass** (5 nowych w `DocxToHtmlConverterSectionReferenceTests`, 5 istniejących).
- Frontend `tsc --noEmit` OK.
- Diagnoza na realnym `stupki.docx`: 3 nagłówki (even/default/first) + 3 stopki, brak `titlePg`/`evenAndOddHeaders`; default=header2 (logo + „ING Bank … Obciągalski"), default footer=footer2 (8pt, #808080). Stary kod mógł renderować pusty even part.
### Notes
- R-10 częściowo zamknięte (default + first-page). Pozostaje: even/odd (brak pól w modelu backendu) i wiele sekcji.
- `stupki.docx` NIE dodano jako fixture (treść wulgarna) — testy używają syntetycznego DOCX o tej samej strukturze.

## 2026-05-28 — Import nagłówka/stopki: wariant even/odd + uzupełnienie .ai
### Changed
- Domena `HeaderFooterContent`: dodane `DifferentOddEven` + `EvenHtml` (obok `DifferentFirstPage`/`FirstPageHtml`).
- `DocxToHtmlConverter`: helper `HasEvenAndOddHeaders` (czyta `settings.xml/evenAndOddHeaders`); `ExtractHeader/ExtractFooter` czytają referencję **Even** gdy włączone → `EvenHtml`/`DifferentOddEven`. „Default" = strona nieparzysta.
- `.ai/FEATURES.md`: nowa sekcja „Feature: Nagłówek i stopka (import + edycja)" + 2 wiersze w tabeli statusu (instrukcja dla agenta: NIE używać `FirstOrDefault` jako głównej ścieżki, spread'ować cały obiekt header/footer).
### Verified
- Backend build OK; testy Header/Footer + SectionReference: **13/13 pass** (3 nowe even/odd).
### Notes
- **Round-trip save** nadal zapisuje tylko default — first/even na zapisie nie są serializowane (R-10 partial, R-11). Import je odczytuje.

## 2026-05-29 — Word-like obraz: 7 trybów zawijania + obramowanie + przycinanie (a:ln + a:srcRect round-trip)
### Changed
- **Panel obrazu** rozszerzony do pełnego Word-like UX:
  - **Zawijaj tekst** (7 trybów wg menu Worda): Równo&nbsp;z&nbsp;tekstem / Ramka / Góra&nbsp;i&nbsp;dół / Za&nbsp;tekstem / Przed&nbsp;tekstem (5 działa) + Przylegle / Na&nbsp;wskroś (2 disabled — wymagają reflow tekstu wokół kształtu, niedostępne natywnie w HTML/CSS). Grid 2 kolumny.
  - **Obramowanie**: checkbox włącz/wyłącz, color picker, grubość (1–20 px), styl (ciągła / przerywana / kropkowana).
  - **Przytnij**: 4 pola % (lewo/prawo/góra/dół, 0–95%) + przycisk „Resetuj przycięcie". Każde pole emituje `cropChange` z aktualnym vector.
  - Wszystkie kontrolki są stateless — emity przekazują się do public methods na `wysiwyg-editor`.
- **`wysiwyg-editor` — nowe public methods**:
  - `setSelectedImagePositionMode` rozszerzony o `'square'` i `'topBottom'` (oprócz `'inline'`, `'front'`, `'behind'`). Square: `float: left; margin: 0 12px 8px 0;`. TopBottom: `display: block; clear: both; margin: 8px auto;`. Inline / front / behind bez zmian.
  - `setSelectedImageBorder({enabled, color, widthPx, style})` — aplikuje inline CSS `border: Npx style #hex` na `<img>` + persist na `data-border-width/color/style`. Walidacja hex.
  - `setSelectedImageCrop({left, right, top, bottom})` — `clip-path: inset(t% r% b% l%)` na `<img>` + persist na `data-crop-l/r/t/b` (clamp 0–95).
  - `resetSelectedImageCrop()` — zero-out crop.
  - `emitImageSelectionState` rozszerzony o pełen snapshot `border` + `crop` z atrybutów (default-safe gdy brak).
- **`wrapExistingImages`** przywraca po loadzie:
  - tryb `square` / `topBottom` z `data-pos-mode` (`float: left` / `display: block + clear: both`),
  - border z `data-border-*` jako inline CSS,
  - crop z `data-crop-*` jako `clip-path`.
- **Backend `HtmlToDocxConverter.BuildImageDrawing`** (round-trip save):
  - Border: `data-border-width/color/style` na `<img>` → `a:ln` w `pic:spPr` z `a:solidFill srgbClr` + `a:prstDash` (mapowanie `solid`→`Solid`, `dashed`→`Dash`, `dotted`→`Dot`). Width: px×9525 EMU. Walidacja hex — niepoprawny kolor → border pomijany (silent fail-safe).
  - Crop: `data-crop-l/r/t/b` (%) → `a:srcRect` na `pic:blipFill` z l/t/r/b w 1/1000 procenta.
- **Backend `DocxToHtmlConverter.ConvertDrawingToHtml`** (round-trip import): czyta `a:ln` (width, srgb, dash) i `a:srcRect` (l/t/r/b) → emituje `data-border-*` i `data-crop-*` na `<img>`. Plus istniejące `data-pos-mode`/`data-x-emu`/`data-y-emu`/`data-width-emu`/`data-height-emu`.
### Verified
- `dotnet build` (Infrastructure) OK; nowe `ImageBorderCropRoundTripTests` (**7/7 pass**: border solid/dashed/dotted, bad-color fail-safe, crop 4-side round-trip, no-attrs-after-plain-image). `ImageFloatingRoundTripTests` (4/4 pass).
- `tsc --noEmit` OK; `npm test`: **164/164 pass** (17 plików; +4 nowe testy panelu: 7 wrap modes z disabled, emisja 5 supported, border toggle, crop clamp 95, resetCrop).
### Notes / ograniczenia (zaktualizowane)
- **Przylegle / Na wskroś** — HTML/CSS nie ma natywnego mechanizmu reflow tekstu wokół arbitralnego kształtu. Pozostają jako disabled w panelu (UI mówi prawdę o tym, co działa).
- **Drag w obrębie strony** dalej działa tylko dla `front`/`behind` (absolute). Dla `square`/`topBottom` (float/block) drag pozostaje inline-drop-into-DOM (rule 14 — sensowne dla float).
- **Crop**: clamp 0–95% per side (pozostawia widoczny obraz). Mapowanie do OOXML 1/1000 procenta (standardowa skala `a:srcRect`).
- **Border**: tylko jednolity kolor + 3 style. Patterny / gradienty / 3D efekty — out of scope.

## 2026-05-29 — Word-like floating: „Przed tekstem" / „Za tekstem" + drag w obrębie strony + round-trip wp:anchor
### Changed
- **Tryby pozycji obrazu** (panel boczny): zastąpiono info-only sekcję 3 prawdziwymi przyciskami: **„W tekście"** (inline), **„Przed"** (float over text), **„Za"** (float behind text). Aktywny tryb podświetlony. `ImageSelectionState` rozszerzony o `positionMode: 'inline' | 'front' | 'behind'`. Nowy `@Output() positionModeChange` w panelu.
- **`wysiwyg-editor` — floating end-to-end**:
  - Public `setSelectedImagePositionMode(mode)`: dla `inline` usuwa wszystkie ślady floating (data-*, position, z-index); dla `front`/`behind` przycina wrapper przy bieżącej pozycji (bounding rect względem `.page`/edytora), ustawia `position: absolute` + `left/top` + `z-index` (10 vs −1) + zapisuje `data-x-px/y-px` na wrapperze i `data-x-emu/y-emu` na `<img>` (do round-tripu).
  - **Drag-aware mousedown**: gdy wrapper jest w trybie floating, kliknięcie + przesunięcie aktualizuje `left/top` (oddzielna ścieżka od inline-drop-into-DOM); na mouse-up zapisuje pozycję w atrybutach + `onContentChange()` (undo/autozapis).
  - `wrapExistingImages` przywraca floating po imporcie: czyta `data-pos-mode` z `<img>` i woła `applyFloatingPosition` na wrapperze (dzięki czemu obraz wczytany z DOCX z `wp:anchor` natychmiast renderuje się jako absolutny w odpowiedniej warstwie).
- **CSS** (`wysiwyg-editor.scss`):
  - `.editor-image-wrapper[data-pos-mode="front"]` → `position: absolute; z-index: 10;`
  - `.editor-image-wrapper[data-pos-mode="behind"]` → `position: absolute; z-index: -1;`
  - Edytory (`.editor-content`/`.header-editor-content`/`.footer-editor-content`/displaye) dostają `position: relative; z-index: 0` → kontekst stacking dla warstwy „za". `.page` już miało `position: relative`.
- **Backend `HtmlToDocxConverter.BuildImageDrawing`**: gdy `<img>` ma `data-pos-mode="front"|"behind"`, emituje `wp:anchor` z `wp:positionH`/`wp:positionV` (relativeFrom=page, offsety z `data-x-emu/y-emu`), `wp:wrapNone`, `behindDoc` zgodnie z trybem, `simplePos=false`, `allowOverlap=true`. Inline pozostaje domyślną ścieżką (`wp:inline`) — brak regresji dla istniejących dokumentów.
- **Backend `DocxToHtmlConverter.ConvertDrawingToHtml`**: wykrywa `wp:anchor`, czyta `BehindDoc` + `PositionOffset` z H/V, emituje na `<img>` `data-pos-mode="front"|"behind"` + `data-x-emu`/`data-y-emu`. Plus istniejące `data-width-emu`/`data-height-emu` (rozmiar dalej round-tripuje).
### Verified
- `dotnet build` (Infrastructure) OK; **4 nowe testy `ImageFloatingRoundTripTests` pass** (front + behind round-trip pozycji i trybu, inline-no-floating-attrs guard, size obok pozycji).
- `tsc --noEmit` OK; `npm test`: **162/162 pass** (+2 nowe testy panelu: active state aktualnego trybu + emisja `positionModeChange` dla 3 przycisków).
### Notes / pozostałe ograniczenia
- **Wrap modes wciąż out of scope**: `square`/`tight`/`through`/`topAndBottom` (z reflow tekstu) odłożone — wymagają nietrywialnego mechanizmu HTML/CSS (text-wrap niedostępny w przeglądarce). Panel jasno mówi: „Square/tight wrap — w przygotowaniu". Front/behind nie wymagają wrap'u (Word też używa `wrapNone`).
- **Drag floating** działa w obrębie kontenera (najbliższy `.page` / `editor-content`). Przeniesienie obrazu między stronami / sekcjami z anchor-rekalkulacją — roadmap.
- **Backend test image bytes**: round-trip używa 1×1 PNG. Realne dokumenty z floating image z Worda — manual verify.

## 2026-05-29 — Word-like pozycjonowanie obrazów (MVP) — selection-driven side panel
### Changed
- **Nowy komponent `d2-image-properties-panel`** (`components/image-properties-panel/`: TS+HTML+SCSS+spec) — czwarty docked panel w istniejącym slocie (Wyszukiwanie / Tabele / Nagłówek+Stopka / **Obraz**). Stateless: każda zmiana w polu emituje output do parenta, parent woła publiczne metody na `wysiwyg-editor`. Sekcje: Rozmiar (szer./wys. + zachowaj proporcje + „Przywróć proporcje"), Wyrównanie (L/C/R/—; aplikowane do akapitu nadrzędnego), Opływanie (info-only — floating w przygotowaniu), Usuń obraz.
- **`wysiwyg-editor` outputs i public API**:
  - Nowy `@Output() imageSelectionChange = EventEmitter<{ widthPx, heightPx, aspectRatio, alignment } | null>` — emitowany przy `selectImageWrapper`, `clearSelectedImage`, po resize-end i po drop-end.
  - Nowe metody publiczne: `setSelectedImageWidth(px, lockAspect=true)`, `setSelectedImageHeight(px, lockAspect=true)`, `setSelectedImageAlignment('left'|'center'|'right'|null)`, `resetSelectedImageAspect()`, `removeSelectedImage()`. Wszystkie reuse istniejącego pipeline: aktualizacja inline `style.width/height` + `data-width-emu`/`data-height-emu` (parametry konsumowane przez `HtmlToDocxConverter` przy eksporcie), po czym `onContentChange()` → debounced undo-stack + autozapis.
- **`document-editor`**: signal `selectedImage` + `imageLockAspect` (default true) + computed `showImagePanel = selectedImage() !== null && !showFindReplace() && !showTablePanel()`. **Panel obrazu ma pierwszeństwo nad panelem nagłówka/stopki** (HF wpuszcza — `showHeaderFooterPanel` rozszerzony o `&& !showImagePanel()`), żeby użytkownik mógł formatować logo w edytowanym nagłówku. Ruler-h-spacer wpuszcza dodatkowy panel.
- **Roundtrip**: rozmiar już round-tripuje przez istniejące `data-*-emu` na `<img>` (czytane w `DocxToHtmlConverter`/`HtmlToDocxConverter`). Pozycja inline (kolejność w DOM) round-tripuje przez sam HTML. Alignment przez `text-align` na akapicie (zachowane w obie strony — patrz `DocxToHtmlConverterFidelityTests`).
### Verified
- `tsc --noEmit` OK; `npm test`: **160/160 pass** (17 plików; +9 nowych `image-properties-panel.spec.ts` + 1 koordynacja w `document-editor.spec.ts`).
### Notes / ograniczenia MVP
- **Floating positioning (wp:anchor) odłożone do następnej iteracji.** Word-like wrap modes (square / tight / through / top-and-bottom / behind / in-front-of) nie są obsługiwane w tym MVP. Panel pokazuje sekcję „Opływanie tekstu" jako info-only („W tekście (inline). Floating w przygotowaniu.") — UI nie udaje funkcji, której nie ma. Roadmap: rozszerzenie `DocxToHtmlConverter`/`HtmlToDocxConverter` o `wp:anchor`/`wp:positionH`/`wp:positionV`/`wp:wrap*` + signal modelu w editor (`data-pos-mode="floating"`, `data-x-emu`, `data-y-emu`, `data-wrap`).
- **Drag** (przenoszenie obrazu w obrębie strony) i **resize** (uchwyty rogu + boków, corner = aspect-locked) ZAWSZE istniały — nie zmieniam mechaniki, tylko emituję snapshot po każdym z tych zdarzeń, więc panel jest spójny z stanem DOM.
- **Klawiatura**: Delete/Backspace usuwa zaznaczony obraz (istniejące), Escape deselect (istniejące). Strzałkowe przesuwanie do roadmapy.

## 2026-05-29 — Reguła `userDownload`: kontrola pobierania edytowanego pliku
### Changed
- **Reguła domenowa**: nowe opcjonalne pole `userDownload` w metadanych dokumentu (`documents.metadata` JSON). Default `false`; brak / null / non-true ⇒ blokada. Niezależne od `returnUrl`.
- **Application: wspólny parser metadanych** `Features/Documents/Common/ExternalDocumentMetadata` (record + tolerant `Parse(string?)` + `Serialize()` + `IsUserDownloadAllowed`). Zastąpił 4 lokalne kopie `sealed record ExternalMetadata` w handlerach (`GetDocumentMetadata`, `GetDocumentStatus`, `UpdateCallbackUrl`, `FinishAndSendDocument`-touchpoints) — single source of truth dla shape JSON.
- **External API** (`POST /api/v1/document`): `CreateDocumentRequest` dostaje opcjonalne pole `UserDownload: bool?`. Mapowane do `metadata.userDownload`. Przesłanie `null`/`false`/braku → pole w JSON `null` (parser interpretuje jako blokadę).
- **Local upload** (`UploadDocumentCommandHandler`): backend **sam** ustawia `userDownload=true` w nowo tworzonym `Document.Metadata` (rule 12 — anti-tamper; klient nie może wymusić ani zablokować). Bo lokalny upload nie ma `returnUrl` — pobranie jest jedynym sposobem odzyskania edytowanego pliku.
- **Nowy gated endpoint**: `POST /api/documentstorage/{masterId}/user-download` (`DocumentStorageController.DownloadEditedDocument`) — konwertuje aktualny stan edytora HTML → DOCX **tylko** gdy `userDownload == true`. Sentinel error `USER_DOWNLOAD_FORBIDDEN:` mapowany w controllerze na **HTTP 403**. Dawny stateless `POST /api/document/save` pozostaje bez zmian (używany przez inne flow); GUI nie korzysta już z niego dla „Pobierz dokument".
- **DTO statusu/metadanych eksponują flagę** (`DocumentMetadataDto.UserDownload`, `DocumentStatusDto.UserDownload`) — system zewnętrzny widzi, czy użytkownik ma prawo do pobrania.
- **Frontend**: `DocumentMetadataDto.userDownload: boolean` + sygnał `userDownload` + `canUserDownload` computed w `document-editor`. Pozycja menu „Pobierz dokument" gated `@if (canUserDownload())`. `document.service.downloadDocument` zastąpiony przez `documentStorageService.downloadEditedDocument(masterId, request)` → POST na nowy gated endpoint; obsługa 403 z czytelnym komunikatem + synchronizacja lokalnego sygnału.
### Verified
- `dotnet build` obu solucji OK (0 błędów); pełna solucja backendu: **325/325 pass** (Application 192 +20 nowych, Domain 57, Api 33, Infrastructure 43; Integration 6 skipped — DB-bound). 
- `tsc --noEmit` OK; `npm test`: **150/150 pass** (+4 nowych w `document-editor.spec.ts`).
### Notes / ograniczenia
- **Raw download endpoints** (`GET .../{masterId}/download`, `GET .../versions/{vid}/download`) NIE są gated — używane przez edytor do *ładowania* bajtów (nie do user-downloadu). Ten threat-vector (zalogowany użytkownik z bezpośrednim wywołaniem URL) jest poza zakresem tego zadania, do udokumentowania jako follow-up R-NEW. W praktyce: te endpointy nie zwracają „edytowanego" pliku — tylko ostatnio zapisany; do faktycznego pobrania edycji niezbędne jest `POST .../user-download`.
- Stary `documentService.downloadDocument` pozostaje w bazie kodu (martwy dla user-download, używany w innych miejscach jak sign flow) — refactor poza zakresem.

## 2026-05-29 — External API: 3 nowe endpointy (callback URL / unlock / status)
### Changed
- **Domena.** `Document` dostaje dwie nowe metody (mirror `MarkEditing/Sending/...`):
  - `MarkSaved()` — odwraca `Editing → Saved` (przeznaczone wyłącznie dla unlock; stany wysyłki forbidden po stronie handlera).
  - `UpdateMetadata(string?)` — ustawia metadane (JSON); walidacja struktury w warstwie aplikacji.
- **Application.** Trzy nowe jednostki MediatR w `Features/Documents/`:
  - `Commands/UpdateCallbackUrl/{Command,Handler}` — walidacja URL przez `DocumentDelivery.IsValidRecipientUrl` (ten sam, co worker), limit 2048 znaków, zachowuje `classification` w JSON, blokuje stany `Sending`/`Sent`/`DeliveryFailed`.
  - `Commands/UnlockDocument/{Command,Handler}` — `Editing → Saved`; idempotentne dla `Saved` (`Result { Changed=false }`), blokowane dla stanów wysyłki. Logowane (poziom Info, z opcjonalnym `Reason`).
  - `Queries/GetDocumentStatus/{Query,Handler}` — `DocumentStatusDto { masterId, status, isLocked, hasCallbackUrl, activeVersionId?, activeVersionNumber?, activeVersionModifiedAt?, latestDelivery? }`. Pełny `callbackUrl` celowo nie surfacowany (może zawierać token).
- **External API.** `D2ServicesViewerEditor.Api/Controllers/DocumentController` rozszerzony o:
  - `PUT /api/v1/document/{masterId:guid}/callback-url` → 204/400/404/409. URL nigdy nie trafia do logów.
  - `POST /api/v1/document/{masterId:guid}/unlock` → 200 `UnlockDocumentResult { masterId, changed }` / 404 / 409.
  - `GET /api/v1/document/{masterId:guid}/status` → 200 `DocumentStatusDto` / 404.
- **Decyzja kontraktowa** (`.ai/API_CONTRACTS.md`): wszystkie trzy operują na **`masterGuid`** (returnUrl w master metadata; status = atrybut master; brak osobnego user-locka — `Editing` to ten lock). Tabela rozstrzygnięć w API_CONTRACTS.
### Verified
- `dotnet build` obu solucji OK (0 błędów).
- Pełna solucja backendu: Domain 57 + Application 172 (+25 nowych) + Api 33 + Infrastructure 43 = **305/305 pass**; Integration 6 skipped (DB-bound).
### Notes
- **Autoryzacja**: nowe endpointy używają tego samego mechanizmu co istniejący `POST /api/v1/document` (obecnie brak `[Authorize]` w `D2ServicesViewerEditor` — match z istniejącą konwencją; ewentualne wprowadzenie auth schematu pokryje WSZYSTKIE endpointy zewnętrzne jednym przepisem).
- **SSRF**: walidacja URL ogranicza protokoły do http(s) i kapuje długość; brak allowlisty hostów (świadome — match z obecnym podejściem ingestu, ślad w RISKS R-07).
- **Concurrency**: `UpdateCallbackUrl` / `UnlockDocument` modyfikują agregat `Document` przez repo + `SaveChangesAsync` (EF Core change tracker + transakcja na poziomie SaveChanges); pełna optymistyczna kontrola wersji nie istnieje (brak `RowVersion`) — zgodne z resztą domeny.

## 2026-05-29 — Menu „Pomoc" + status autozapisu w stopce
### Changed
- **Nowa zakładka menu „Pomoc"** w `document-editor.html` (ostatnia po „Widok"), z jedną pozycją dropdown **„Zgłoś"** → wywołuje istniejące `openReportEmail()` (bez duplikacji logiki — ta sama metoda co dawny przycisk). Sygnał `showHelpMenu = signal(false)` + `toggleHelpMenu()` zgodne z konwencją innych menu (zamyka pozostałe przez `closeAllMenus()`, którego rozszerzono o nowy sygnał). `openReportEmail()` woła teraz `closeAllMenus()` na początku — pozycja w dropdownie zamyka go po kliknięciu, spójnie z resztą menu.
- **Stary przycisk „Zgłoś" usunięty** z `.header-right`. Toolbar w prawym górnym rogu jest teraz czystszy: badge klasyfikacji, switch „Autozapis", opcjonalny „Zapisz"/„Zakończ".
- **Data/czas ostatniego autozapisu przeniesione do stopki.** Dawniej `.autosave-status` (`min-width:70px`) obok switcha pokazywał „Zapisywanie…" / „Zapisano o HH:MM:SS" / „Błąd auto-zapisu" — usunięte z toolbara (przy switchu zostaje tylko switch + label „Autozapis", zgodne z task #2). Nowy `.status-autosave` w `.editor-footer > .status-right`: „Autozapis: HH:MM:SS" / „Autozapis: zapisywanie…" / „Autozapis: błąd zapisu" / fallback „Autozapis: włączony"; ukryty gdy autozapis off lub brak `versionId`. Bez nowych źródeł danych — te same sygnały (`autoSaveEnabled`/`autoSaveStatus`/`lastAutoSaveAt`).
- SCSS: nowy `.status-autosave` (`#555`, `is-saving #1a73e8`, `is-error #cc0000 bold`) — subtelny, dopasowany do reszty status-baru.
- `document-editor.spec.ts` +4 testy: `toggleHelpMenu` otwiera/zamyka i koliduje z `showViewMenu`, `closeAllMenus` zamyka `showHelpMenu`, `openReportEmail()` zamyka dropdown.
### Verified
- `tsc --noEmit` OK; `npm test` → **146/146 pass**.
### Notes
- Logika autozapisu i akcji `openReportEmail` bez zmian (rule 6/7) — przeniesienie czysto prezentacyjne.

## 2026-05-29 — Panel nagłówka/stopki: porządek wizualny + NG8107
### Changed
- **Usunięty primary CTA „Zamknij nagłówek i stopkę"** z `d2-header-footer-panel`. Był wizualnie redundantny z X w nagłówku panelu i z ESC (oba wciąż delegują do `editor.stopEditingHeaderFooter()`). Dock jest spójny z `Wyszukiwanie` i panelem tabeli — X jako jedyne zamknięcie z panelu. Spec testu zaktualizowany: zamiast Primary asercja na X.
- **NG8107 (Angular template lint)**: bindingi panelu w `document-editor.html` używały `editor?.method()`, ale `editor` jest `@ViewChild` z `editor!:` (non-null) i panel renderuje się tylko po `editingSection() !== 'body'`, które emituje sam edytor — więc operand zawsze ≠ null. Zmienione 9 wystąpień `editor?.` → `editor.`; semantyka bez zmian, ng serve nie emituje już warningów NG8107.
### Verified
- `tsc --noEmit` OK; `npm test` → **142/142 pass**; `ng serve` bez NG8107.
### Notes
- Pełna lista wyjść z trybu edycji: X w nagłówku panelu, ESC, klik w body editor.

## 2026-05-29 — UX poprawki: autozapis, prowadnice, obrazy nagłówka, panel boczny
### Changed
- **Autozapis (stabilny layout).** `document-editor.html`: `.autosave-status` jest teraz renderowany ZAWSZE (gdy `documentVersionId()` istnieje), a stan idle/disabled emituje pusty content (klasa `is-empty`). `min-width:70px` rezerwuje przestrzeń, więc toggle nie powoduje reflow toolbara — przyciski Zgłoś/Zapisz/Zakończ nie skaczą.
- **Prowadnice marginesów domyślnie ukryte.** `document-editor.ts` `showMarginGuides = signal(false)`, `wysiwyg-editor.ts` `@Input() showMarginGuides = false`. Menu/dialog Widok dalej działają (toggle bez zmian).
- **Obraz nagłówka — brak reskalowania w trybie edycji.** `wysiwyg-editor.ts` `wrapExistingImages` NIE usuwa już inline `width`/`height` z `<img>`; CSS `.editor-image-wrapper img { width: 100%; height: auto }` zmieniony na samo `max-width: 100%` (rule 14 — bez „pozornego" reskalu). `display: inline-block` wrappera kurczy się do obrazu; obraz w edycji ma ten sam rozmiar co w podglądzie.
- **Pasek nagłówka/stopki → panel boczny `d2-header-footer-panel`.** Nowy komponent (TS+HTML+SCSS+spec) w `components/header-footer-panel/`, dokowany w tym samym slocie co `Wyszukiwanie` i `d2-table-properties-panel` (szerokość 300px, ten sam akcent #1a73e8). Sekcje: Widok („Inna pierwsza strona"), Wstaw (Obraz, Numery stron), Ustawienia (Format nagłówka/stopki, Usuń nagłówek/stopkę), Primary CTA „Zamknij nagłówek i stopkę". Stateless — wszystkie akcje delegowane do `wysiwyg-editor` (rule 9 — brak równoległej logiki).
  - Stary pływający `.header-toolbar`/`.footer-toolbar` USUNIĘTY z `wysiwyg-editor.html` (SCSS pozostaje jako martwy, nie używany — możliwe usunięcie w follow-up).
  - Koordynacja single-mode: `showHeaderFooterPanel = computed(() => editingSection() !== 'body' && !showFindReplace() && !showTablePanel())`. Find / Tables mają pierwszeństwo (spec #10); po ich zamknięciu HF panel wraca, jeśli edycja trwa. ESC nadal deleguje do `editor.stopEditingHeaderFooter()` (Phase 2 commit) → zamyka edycję → zamyka panel.
  - Pozioma linijka: spacer `ruler-h-search-spacer` reaguje też na `showHeaderFooterPanel()`.
- Nowy test `header-footer-panel.spec.ts` (6) + `document-editor.spec.ts` +5 (koordynacja Find/Tables/HF, ESC).
### Verified
- `tsc --noEmit` OK; `npm test` → **142/142 pass** (16 plików).
### Notes / ograniczenia
- **Top-margin drag w trybie edycji nagłówka**: usunięcie pływającego paska eliminuje overlay, który zasłaniał obszar interakcji nad/obok nagłówka. Rzeczywiste przeciągnięcie zależy od linijki pionowej (`d2-ruler` mode=vertical) — pełna weryfikacja wymaga manualnego testu w przeglądarce; w razie pozostałego konfliktu warto sprawdzić `z-index` `.page-header.editing` (obecnie 15) vs ruler.
- Nieużywane reguły SCSS `.header-toolbar`/`.footer-toolbar`/`.header-options-btn` itp. zostają jako dead-code (nie wpływają na layout). Czyszczenie — follow-up.

## 2026-05-28 — Faza 2: audyt trybu edycji nagłówka/stopki (vs spec)
### Changed (gap fixes)
- **Routing wariantu na edycji (rule 10: bez pozornej edycji).** `wysiwyg-editor.ts`:
  - `startEditingHeader`/`startEditingFooter` ładuje teraz wariant aktywny na stronie 0 (`_headerFirstPageHtml` gdy `differentFirstPage`, inaczej `_headerHtml`) — wcześniej zawsze ładował default niezależnie od wyświetlanego wariantu.
  - `onHeaderInput`/`onFooterInput` zapisuje do tego samego sygnału co załadowany wariant (nie do `_headerHtml` zawsze) — wcześniej edycja first-page niewidocznie nadpisywała default.
  - `onHeaderBlur`/`onFooterBlur`: analogicznie + wywołują `emitHeaderFooterChanges()` zamiast emitować niekompletne `{html, height}` (gubiło inne warianty u parenta).
  - `_computeHeaderContent`/`_computeFooterContent`: w trybie odd/even strona nieparzysta używa **kanonicznego** `_headerHtml`/`_footerHtml` (= „default" w OOXML), bo backend nie posyła osobnego `oddHtml`. Even strony bez zmian (`_headerEvenHtml`).
- **ESC zamyka tryb edycji nagłówka/stopki** (`document-editor.ts` `onEscapeKeydown`): gdy `editingSection() !== 'body'`, deleguje do `editor.stopEditingHeaderFooter()` i `preventDefault`. Tryb ma pierwszeństwo nad bocznymi panelami (wyszukiwanie, panel tabeli), bo to ostatnio aktywowany kontekst.
- **Przycisk „Zamknij nagłówek i stopkę"** dodany w obu toolbarach nagłówka i stopki (`wysiwyg-editor.html`, `wysiwyg-editor.scss` — `.header-toolbar-close`/`.footer-toolbar-close`, ten sam akcent co istniejący `1a73e8` w toolbarach). Pełni rolę „głównej" akcji wyjścia z trybu (spec funkcjonalny #3).
- Nowy plik testów `wysiwyg-editor.spec.ts` (7 testów): routing default vs first-page (header+footer), pełne emisje variantów, odd=canonical default, dokument bez nagłówka, `stopEditingHeaderFooter`. `document-editor.spec.ts` +4 testy ESC dla header/footer + pierwszeństwo nad side panel + no-op w body.
### Verified
- `npm test` (`ng test` → vitest): **131/131 pass** (15 plików; 11 nowych testów Phase 2).
- `tsc --noEmit`: OK.
### Notes / ograniczenia
- **Edycja even-page nieosiągalna z UI** — template renderuje contenteditable tylko dla strony 0; even-page edytowanie wymagałoby dodatkowego wejścia (nie w tym MVP). Even-page wciąż jest persistowany w round-trip (z importu).
- Toolbar główny (`d2-editor-toolbar`) — formatowanie tekstu — działa na aktywnym `editingSection` przez `getActiveEditable()` (istniejący kod, bez zmian); confirmed via existing wiring `executeCommand` (linia 1335 wysiwyg-editor).
- Undo/redo dla header/footer — `saveToUndoStack` jest wywoływany przez `onHeaderInput`/`onFooterInput` ścieżkę (`emitContent` debouncer); istniejące, bez zmian.

## 2026-05-28 — Faza 1b: pomiary wierności importu nag/stopki (stupki.docx)
### Changed
- Nowy plik `DocxToHtmlConverterFidelityTests.cs` (9 testów) — struktura odzwierciedla rzeczywisty `stupki.docx` (A4, pgMar 1417 twips, header/footer=708, obraz extent 1272540×327354 EMU, stopka L/C/R):
  - geometria: `Header.Height`/`Footer.Height` ≈ 1.25 cm; `Margins` ≈ 2.5 cm;
  - font family: dziedziczony z `docDefaults` (kontener `.header-footer-content` niesie `font-family`);
  - obraz: EMU → 133×34 px (`EmuToPx = emu/914400*96`), aspect zachowany w ±1 px (test 4:1 stays 4:1), `data-width-emu`/`data-height-emu` zachowane do round-tripu;
  - alignment akapitów w stopce: kolejność L/C/R zachowana, `text-align:center/right` emitowane.
- Diagnoza realnego `stupki.docx` potwierdza brak gapów na powyższych ścieżkach (T1-T12 z `<test_requirements>`; T3-T6 już objęte `DocxToHtmlConverterHeaderFooterTests`).
### Verified
- Pełny `D2ViewerEditor.Infrastructure.UnitTests`: **43/43 pass** (9 nowych fidelity).
### Notes
- `stupki.docx` nie używa tabel layoutowych ani anchor-positioned images — pokrycie tych przypadków pozostaje syntetyczne / przyszłe.
- Pozostaje R-11 (round-trip rozmiaru z `docDefaults` przez wrapper `.header-footer-content`) — bez sygnału z `stupki.docx`.

## 2026-05-28 — Round-trip first/even header/footer (zapis)
### Changed
- `HtmlToDocxConverter.AddHeaderAndFooter` zrefaktorowany: wydzielone `WriteHeaderPart`/`WriteFooterPart(html, type)` tworzą osobne `HeaderPart`/`FooterPart` per wariant (Default/First/Even) i wstawiają `HeaderReference`/`FooterReference` z odpowiednim `Type` (helpery `AddHeaderReference(id, type)`/`AddFooterReference(id, type)` zamiast hard-coded Default).
- Dodane `EnsureTitlePage` (wstawia `TitlePage` do `sectPr`, gdy emitowany wariant First) i `EnsureEvenAndOddHeaders` (wstawia `EvenAndOddHeaders` do `settings.xml`, gdy emitowany wariant Even). Helper `GetOrCreateSectionProps`.
- Nowy plik testów `HeaderFooterRoundTripTests.cs` (6 testów): write → read potwierdza, że `DifferentFirstPage`/`FirstPageHtml`/`DifferentOddEven`/`EvenHtml` przeżywają round-trip (header i footer, każdy wariant oddzielnie, plus „all combined").
### Verified
- Backend build OK; pełny `D2ViewerEditor.Infrastructure.UnitTests`: **34/34 pass** (6 nowych round-trip).
### Notes
- Zamyka brakujący kawałek: edycja wariantu first/even w UI jest teraz trwała (reguła 10).
- Pozostaje OPEN: **wiele sekcji** (`sectPr` per-section) — `R-10` zaktualizowany.

## 2026-05-27 — Fix layoutu: banner środowiska przycinał dolny pasek edytora

### Changed
- `styles.scss`: layout powłoki (`d2-root` flex column 100vh + `d2-root > d2-global-banners` `flex:0 0 auto` + routowane strony `d2-document-editor`/`d2-dashboard`/`d2-pdf-maintenance`/`d2-pdf-viewer`/`d2-admin-shell` `flex:1 1 0; min-height:0; overflow:hidden`) przeniesiony do **globalnych** (nieenkapsulowanych) styli. **Przyczyna błędu:** reguły flex dla stron były w stylach komponentu `App` (encapsulation Emulated), ale strony są wstawiane przez `<router-outlet>` jako rodzeństwo i NIE dziedziczą atrybutów `_ngcontent` App → selektory `d2-document-editor{flex:1}` nie działały. Edytor zostawał przy `:host{height:100%}` = pełne 100vh; zawsze widoczny banner środowiska (~22px) spychał dolny pasek (zoom / liczba stron / wersja) poza ekran, gdzie `overflow:hidden` go ucinał.
- `app.ts`: usunięto martwe selektory routowanych komponentów ze styli komponentu (zostaje tylko `:host`), z komentarzem wskazującym globalny layout.

### Verified
- `npx ng test` — 101/101 passed. `npx ng build` — OK.
- Manualnie: z widocznym bannerem środowiska edytor mieści się w oknie, dolny pasek (zoom/strony/wersja) w całości widoczny; działa dla 4 kombinacji bannerów.

## 2026-05-27 — Nazewnictwo admina, dok tabela↔wyszukiwanie, ukrycie pozycji „Wstaw", „Akapit" w toolbarze

### Changed
- **Admin — nazewnictwo** (`Wysyłki` → `Pliki do wysłania`): `admin-shell.html` (nav), `admin-deliveries.html` (tytuł strony + tekst ładowania), `admin-deliveries.ts` (komunikat błędu). Nie ruszano nazw technicznych (trasa `deliveries`, klasy, DTO, statusy — np. status „Wysyłanie" zostaje, bo to stan akcji, nie nazwa obszaru).
- **Dok boczny — przełączanie tabela ↔ wyszukiwanie** (`document-editor.ts` `syncTablePanel`): naprawiono konflikt — gdy karetka wraca do tabeli przy otwartym wyszukiwaniu, dok przełącza się na formatowanie tabeli (zamyka wyszukiwanie przez `closeFindReplace()`). Wcześniej warunek `!showFindReplace()` blokował pokazanie panelu tabeli (ADR-0006 „search ma pierwszeństwo" — świadomie nadpisane wg decyzji zadania). Dok ma jeden aktywny tryb; `tablePanelManuallyClosed` (×) nadal respektowane; ESC/X bez zmian.
- **Menu „Wstaw" — ukryte pozycje** (`document-editor.html`): QR Code / Kod kreskowy, Linia pozioma, Podział strony owinięte w `@if (false)` (ukryte, logika `openBarcodeDialog`/`insertHorizontalLine`/`insertPageBreak` zostaje). Separatory uporządkowane — brak pustych grup/podwójnych separatorów.
- **Toolbar — przycisk „Akapit"** (`editor-toolbar`): nowy `@Output() openParagraph` + przycisk w grupie „Listy i wcięcia" (ta sama ikona co menu „Narzędzia"), podpięty w `document-editor.html` do istniejącego `openParagraphDialog()` — bez duplikacji dialogu/logiki. `aria-label="Akapit"`. Menu „Narzędzia" bez zmian.

### Verified
- `npx ng test` — **101/101 passed (12 plików)**; nowe/rozszerzone: `admin-shell.spec.ts` (3), `editor-toolbar.spec.ts` (+1 „Akapit" = 7), `document-editor.spec.ts` (+5 przełączanie doku = 12).
- `npx ng build` — OK (tylko istniejące ostrzeżenia budżetu).

### Notes
- **QR/barcode** dostępne też przez `(insertBarcode)` toolbara, ale toolbar NIE renderuje widocznego przycisku QR — więc ukrycie w menu „Wstaw" wystarcza. `insertHorizontalLine`/`insertPageBreak` były tylko w menu „Wstaw". **Do potwierdzenia:** czy ukryć też dialog kodu kreskowego z innych ścieżek (obecnie brak innej widocznej).
- **Scenariusz manualny menu „Wstaw"**: otwórz „Wstaw" → brak QR/Linia pozioma/Podział strony; widoczne: Obraz, Tabela, (separator), Nagłówek/Stopka… — bez pustych grup.
- **Scenariusz manualny dok**: klik w tabelę → panel formatowania; lupa/Ctrl+F → wyszukiwanie; ponowny klik w tabelę → panel wraca do formatowania (wyszukiwanie zamknięte); ESC i × zamykają dok.

## 2026-05-27 — Globalny banner środowiska + naprawa layoutu paska braku API

### Changed
- **Nowy kontener bannerów** `d2-global-banners` (`components/global-banners/`) — renderowany w `App` w normalnym flow flex (zamiast bezpośredniego `<d2-offline-banner>`). Stała kolejność: banner środowiska na górze, pasek offline/API pod nim. Bannery są w flow → `App` (`:host` flex column, 100vh) rezerwuje ich wysokość, edytor (`flex:1`) nigdy nie jest przykrywany — stabilne dla 4 kombinacji (brak / tylko env / tylko API / oba).
- **Nowy banner środowiska** `d2-environment-banner` (`components/environment-banner/`) — prezentacyjny; źródło środowiska = `BuildInfoService.environment()` (jedyne wiarygodne: health-check API z fallbackiem na front-config — **uwaga:** brak `fileReplacements` w `angular.json`, więc `environment.ts` jest zawsze `PRD`, dlatego nie używamy go wprost). Mapowanie przez czystą funkcję `core/utils/environment-banner.util.ts` (`resolveEnvironmentBanner`): Local→niebieski, DEV→zielony, UAT/TST→jasnożółty, PRE→purpurowy, PROD/PRD→ukryty, nieznane→neutralny fallback `Wersja {nazwa}`. `role="status"`, kontrast AA.
- **Naprawa paska braku API** (`offline-banner.scss`): usunięto arbitralny `z-index: 10000` (pasek jest w flow — wysoki z-index tylko ryzykował przebicie przez modale edytora o z-index 1000–2000); dodano `width:100%` i komentarz o roli `position: relative` (kotwica przycisku ×). Logika wykrywania (`ConnectionStatusService`: online/offline + health-check 30 s + interceptor) **bez zmian** — poprawka czysto prezentacyjno-layoutowa.
- `App` (`app.ts`): import `GlobalBannersComponent`, selektor `d2-global-banners` w `:host` (`flex-shrink:0`).
- Naprawiono nieaktualny scaffold `app.spec.ts` (oczekiwał „Hello, frontend"): teraz `provideRouter([])` + stub `BuildInfoService`/`ConnectionStatusService`, asercja montażu `d2-global-banners`.

### Verified
- `npx ng test` — **92/92 passed (11 plików)**; nowe specy: `environment-banner.util.spec.ts` (10), `environment-banner.spec.ts` (10: pełna macierz env + reaktywność + role), `global-banners.spec.ts` (4: kolejność, widoczność offline, niezależność, prod). Suite w pełni zielony (naprawiony `app.spec.ts`).
- `npx ng build` — OK (tylko istniejące ostrzeżenia budżetu).

### Notes
- **Scenariusz manualny**: na DEV/UAT/itd. (gdy API health zwraca env) pojawia się kolorowy pasek na górze, ponad paskiem offline; oba w flow, edytor/toolbar/panel boczny nieprzykryte. Gdy API niedostępne — pasek offline pod bannerem env; gdy API wróci — znika. Na PROD (PRD) banner env ukryty.
- Banner reaguje na health-check: do pierwszej odpowiedzi env=PRD (front fallback) → ukryty (brak migotania błędną etykietą), po odpowiedzi API pokazuje właściwe środowisko.

## 2026-05-27 — ESC zamyka boczny panel + weryfikacja przycisku „Cofnij" (undo)

### Changed
- Frontend `document-editor.ts`: nowy `@HostListener('document:keydown.escape')` (`onEscapeKeydown`) zamyka aktywny boczny panel (Wyszukiwanie / właściwości+stylizacja tabeli) tą samą logiką co przycisk × (`closeFindReplace` / `closeTablePanel`), przez wspólną metodę `closeActiveSidePanel()`. Listener jest na `document` (panel może nie mieć focusu — karetka w edytorze). Pierwszeństwo: jeśli ESC obsłużył już szczegółowy handler (`event.defaultPrevented`, np. deselekcja obrazu w `wysiwyg-editor`) albo otwarty jest dialog/menu kontekstowe (`isAnyDialogOpen()`/`showContextMenu()`) — panel nie jest ruszany; gdy żaden panel nie jest otwarty — brak skutków. Focus wraca do edytora tylko, gdy był wewnątrz panelu.
- Brak zmian mechanizmu undo: przeanalizowano i potwierdzono poprawność. Przycisk „Cofnij" w `editor-toolbar` (`(click)="executeCommand('undo')"`, `[disabled]="!editorState?.canUndo"`) używa tej samej metody `WysiwygEditorComponent.undo()` co skrót Ctrl+Z (`handleKeyboard`). Operacje tabeli (wiersze/kolumny/obramowania/scal/kolor) trafiają do historii przez `notifyEditorChange()` → syntetyczny `input` → `onContentChange` → `saveToUndoStack()`.

### Verified
- `npx ng test` — nowe specy: `editor-toolbar.spec.ts` (6) + `document-editor.spec.ts` (7) = 13/13 passed; cała reszta bez regresji (jedyne czerwone to znany, niezależny scaffold `app.spec.ts` — brak `_HttpClient`).
- `npx ng build` — OK (tylko istniejące ostrzeżenia budżetu SCSS/initial).
- Undo na poziomie contenteditable/`execCommand` nie jest testowalny w jsdom — scenariusz manualny niżej.

### Notes
- **Scenariusz manualny undo** (uruchom GUI, otwórz dokument do edycji): (1) wpisz tekst → przycisk „Cofnij" się aktywuje → klik cofa wpis; (2) zaznacz tekst, kliknij **B** (bold) → „Cofnij" cofa pogrubienie; (3) klik w tabelę → panel tabeli → dodaj wiersz / zmień obramowanie → „Cofnij" cofa operację tabeli; (4) Ctrl+Z daje ten sam efekt co przycisk; (5) „Ponów"/Ctrl+Y przywraca; (6) brak błędów w konsoli; (7) gdy stos pusty — „Cofnij" jest wyszarzony.
- **Scenariusz manualny ESC**: otwórz panel Wyszukiwania → ESC zamyka; klik w tabelę (panel właściwości/stylizacji) → ESC zamyka; ESC przy zamkniętym panelu → nic; zaznacz obraz w edytorze przy otwartym panelu → ESC najpierw odznacza obraz (panel zostaje), kolejny ESC zamyka panel; otwarty dialog (np. wstaw tabelę) + ESC → panel w tle zostaje.
- Redundancja w `wysiwyg-editor`: każdy edytor strony ma jednocześnie `(input)="onPageInput()"` (debounce 500 ms) ORAZ `addEventListener('input', …)` → `onContentChange()` (snapshot natychmiastowy). Undo działa, ale jest granularne (per znak). Nie ruszano — poza zakresem, ryzyko regresji undo dla tabel (zależnych od ścieżki `onContentChange`).

## 2026-05-27 — Fix odwzorowania nagłówka/stopki DOCX→HTML (za duży tekst, kolor, fallback)

### Changed
- Backend `DocxToHtmlConverter`: nagłówek/stopka są teraz owijane w `<div class="header-footer-content" style="font-family:..;font-size:..pt">` z domyślnym krojem/rozmiarem z `w:docDefaults/rPrDefault` — analogicznie do body `.document-content`. Wydzielono wspólny helper `BuildDefaultContainerCss()` (body + header/footer). **Przyczyna „za dużego tekstu":** runy nagłówka/stopki bez własnego `w:sz`/`w:rFonts` nie miały żadnego kontenera z domyślnym rozmiarem (body go miało), więc dziedziczyły domyślny rozmiar EDYTORA, nie dokumentu.
- Frontend `wysiwyg-editor.scss`: `.header-display/.footer-display/.header-editor-content/.footer-editor-content` dostały domyślny `font-family` (corporate var), `font-size:11pt`, `line-height:1.15`, `color:#000` — fallback dla dokumentów bez kontenera z backendu oraz dla regionu edycji. Usunięto `color:#202124` (Word „auto" = czarny). Kontener z backendu (inline font-size z docDefaults) ma wyższą specyficzność i nadpisuje, gdy DOCX definiuje rozmiar; inline color/size runów zawsze wygrywa.
- Testy backendu: `DocxToHtmlConverterHeaderFooterTests` (5) — kontener z docDefault 10pt w nagłówku i stopce, zachowanie jawnego koloru (`#1F3864`/`#C00000`) i rozmiaru (8pt = 16 half-points), brak docDefaults → kontener bez wymuszonego rozmiaru.

### Verified
- `dotnet build` Infrastructure — OK (0 błędów). `dotnet test --filter DocxToHtmlConverterHeaderFooter` — 5/5 passed.
- `npm run build` (front) — OK.

### Notes
- Jednostki: `w:sz` jest w half-points → `pt = sz/2` (potwierdzone w `GetRunStyleClean`/`ConvertRunPropertiesToCss` i teraz w kontenerze docDefaults).
- **Niezałatwione (udokumentowane):** R-10 — bierzemy tylko `HeaderParts.FirstOrDefault()`, bez default/first-page/even-odd i bez wielu sekcji; R-11 — round-trip rozmiaru nagłówka zależy od `docDefaults` (kontener spłaszczany przy zapisie, fallback frontu 11pt zapewnia spójność wizualną).

## 2026-05-27 — Bugfix: panel obramowań gubił aktywną tabelę + nieczyszczone zaznaczenie

### Changed
- GUI: `core/utils/table-context.util.ts` — nowa czysta funkcja `resolveTableContext(anchorNode, editorEl)` → `outside-editor` / `in-table` / `outside-table`. Wydzielona z `detectTableContext`, testowalna.
- GUI: `document-editor.ts` `detectTableContext()` — **nie czyści już aktywnej tabeli, gdy selekcja jest poza edytorem** (interakcja z panelem/toolbar = zachowaj ostatni kontekst). Czyszczenie tylko gdy karetka jest realnie w treści poza tabelą (`outside-table`) — wtedy dodatkowo `clearCellSelection()` (usuwa klasę `table-cell-selected`).
- GUI: podpięto `(selectionChange)="onEditorSelectionChange()"` w `document-editor.html` (wcześniej `selectionChange` z edytora był nieobsłużony) — kliknięcie innego akapitu aktualizuje teraz kontekst tabeli i zamyka/aktualizuje panel.
- GUI: testy `table-context.util.spec.ts` (outside-editor/in-table/outside-table/null) + panel: „border interactions never emit close".

### Verified
- `npm run build` — OK. `npx ng test --watch=false` — 55 passed (2 stare scaffoldowe `app.spec.ts` padają niezależnie).

### Notes
- **Przyczyna #1** (panel pokazywał „Wybierz tabelę" po kliknięciu obramowania): `applyTableBorderScope` → `notifyEditorChange()` → `input` → `updateState` → `stateChange` → `detectTableContext`, które przy selekcji poza edytorem (focus w panelu) zerowało `activeTable`. Naprawa: zachowanie kontekstu przy `outside-editor` (model „last known table context").
- **Przyczyna #2** (tabela nadal wyglądała na zaznaczoną po kliknięciu akapitu): `selectionChange` edytora nie był podpięty, więc sam ruch karetki nie aktualizował kontekstu. Naprawa: podpięcie `selectionChange` + czyszczenie wizualnego zaznaczenia przy `outside-table`.
- `onPanelMouseDown` (preventDefault na nie-inputach) pozostaje jako komplementarne zabezpieczenie zachowujące karetkę; naprawa działa też, gdyby focus jednak uciekł.

## 2026-05-27 — Auto-wykrywanie zakresu obramowania (ukryty ręczny wybór celu)

### Changed
- GUI: usunięto ręczny przełącznik celu („Cała tabela / Komórka / Zaznaczenie") z panelu obramowań (`TableBorderTarget`, `TABLE_BORDER_TARGETS`, sekcja „Zakres", `setTableBorderTarget`). Zakres jest teraz **wnioskowany z bieżącego zaznaczenia**.
- GUI: `core/utils/table-style.util.ts` — nowa czysta funkcja `classifyBorderTarget(table, cells)` → `BorderTargetInfo { kind, rows, cols }`: jedna komórka → `cell`; pełna szerokość → `row`; pełna wysokość → `column`; pełna szer.+wys. → `table`; reszta → `range`; pusty zbiór → `none`.
- GUI: `document-editor.ts` — `resolveAutoTargetCells` (zaznaczone komórki → ten zbiór; brak → aktywna komórka; dalej cała tabela) + `borderTargetInfo` (computed, reaktywny na `selectedCells`/`activeTableCell`). Klik ikony „gdzie narysować" stosuje linię do auto-celu (`applyBorderToCells` liczy krawędzie względem prostokąta zbioru).
- GUI: panel pokazuje tylko **podpis** „Zastosowanie: aktywna komórka / zaznaczone komórki (R×C) / cały wiersz / cała kolumna / cała tabela" (read-only, `aria-live`), bez kontrolki wyboru. UX: użytkownik wybiera rodzaj/grubość/kolor linii i klika miejsce — zakres dobiera się sam.
- GUI: testy — `classifyBorderTarget` (cell/row/column/table/range/none) + panel (podpis zamiast selektora; brak `.table-panel-segmented`).

### Verified
- `npm run build` — OK. `npx ng test --watch=false` — 49 passed (2 stare scaffoldowe `app.spec.ts` padają niezależnie).

### Notes
- Domyślny zakres przy samej karetce (brak zaznaczenia) = **aktywna komórka** (jak w MS Word). Wykrycie „cały wiersz/kolumna" działa, gdy użytkownik zaznaczy te komórki (drag custom cell-selection); edytor nie ma osobnego gestu „klik nagłówka wiersza/kolumny" — to świadome ograniczenie, bezpieczny fallback do zaznaczonego zbioru.
- Tabele z mocno scalonymi komórkami mogą dać przybliżoną *etykietę* zakresu; samo rysowanie obramowania działa zawsze na przekazanym zbiorze komórek.

## 2026-05-27 — Szczegółowy edytor obramowań tabeli (zamiast galerii presetów)

### Changed
- GUI: **usunięto galerię gotowych stylów tabel** (`TABLE_STYLE_PRESETS`, sekcje „Style tabeli"/„Opcje stylu", `applyTablePreset`, markery `data-table-style*`, `readTableStyleState`). Zakładka „Style" → **„Obramowania"**.
- GUI: `models/table-style.model.ts` przebudowany pod obramowania: `TableBorderLineStyle` (solid/dashed/dotted/double/none), rozszerzony `TableBorderScope` (all/none/outer/inner/inner-horizontal/inner-vertical/top/bottom/left/right), `TableBorderSettings` z polem `style`, `TableBorderTarget` (table/cell/selection) + listy prezentacyjne (`TABLE_BORDER_LINE_STYLES/_WIDTHS/_COLORS/_SCOPES/_TARGETS`).
- GUI: `core/utils/table-style.util.ts` — `applyBorderToCells(cells, scope, border)` liczy krawędzie względem prostokąta opisanego na **dowolnym zbiorze komórek** (tabela / komórka / zaznaczenie); `applyBorderScope` deleguje dla całej tabeli; `restoreDefaultTableBorders` przywraca siatkę 1px. Linie kodowane jako `Npx <style> <color>` w inline `border*` — trwałość wg ADR-0007.
- GUI: `table-properties-panel` — zakładka „Obramowania" z sekcjami: **Zakres** (cel: cała tabela/komórka/zaznaczenie), **Rodzaj linii** (chipy z podglądem), **Grubość linii** (Cienka/Standardowa/Średnia/Gruba), **Kolor linii** (paleta + color picker + reset), **Gdzie narysować** (10 neutralnych ikon SVG: wszystkie/brak/zewn./wewn./poziome/pionowe/góra/dół/lewa/prawa), **Reset** (domyślne obramowanie / usuń obramowania). Podświetlany stan aktywny (rodzaj/grubość/kolor/ostatni zakres/cel).
- GUI: `document-editor.ts` — sygnały `tableBorderColor/Width/Style/Target` + `lastBorderScope`; metody `applyTableBorderScope` (z rozwiązaniem celu i fallbackiem zaznaczenie→komórka→tabela), `clearTableBorders`, `restoreDefaultTableBorders`, `resetTableBorderSettings`, settery pióra. Usunięto metody presetów.
- GUI: testy przepisane — `table-style.util.spec.ts` (zakresy, cel komórka/zaznaczenie, rodzaj/grubość linii, restore default) i `table-properties-panel.spec.ts` (zakładka Obramowania, ikony zakresu, zmiana rodzaju/grubości/koloru/celu, reset).

### Verified
- `npm run build` — OK (tylko istniejące ostrzeżenia budżetu).
- `npx ng test --watch=false` — 42 passed. Padają tylko 2 stare scaffoldowe `app.spec.ts` (niezwiązane).

### Notes
- Obramowanie działa na **trzech poziomach**: cała tabela, pojedyncza komórka, zaznaczony fragment (custom cell-selection). „Zaznaczenie" bez zaznaczonych komórek → bezpieczny fallback do aktywnej komórki.
- Zakresy są **addytywne** (dokładają krawędzie, jak przyciski w Word); „Brak" czyści, „Domyślne obramowanie" przywraca siatkę 1px. Zachowanie treści gwarantowane (modyfikujemy tylko `border*`).
- Galeria presetów świadomie wycofana w tej iteracji (priorytet: precyzyjna edycja linii). Patrz ADR-0007 (zaktualizowane).

## 2026-05-27 — Stylizacja tabel (gotowe style, obramowania, opcje, reset) + zakładki w panelu

### Changed
- GUI: nowy model `models/table-style.model.ts` — `TableStylePresetId`, `TableStyleOptions` (headerRow/bandedRows/firstColumn/lastColumn), `TableBorderScope`, `TableBorderSettings`, `TableStylePreset` + `TABLE_STYLE_PRESETS` (6 neutralnych stylów: Prosty, Siatka, Nagłówek, Naprzemienny, Biznesowy, Minimalny).
- GUI: nowy util `core/utils/table-style.util.ts` — czyste, testowalne funkcje DOM: `applyTablePreset` (idempotentne przeliczenie wyglądu), `applyBorderScope` (all/none/outer/inner), `resetTableStyle` (przywraca domyślny wygląd, zachowuje treść), `readTableStyleState` (odczyt presetu/opcji z markerów `data-*`). Styl utrwalany jako **style inline** na `<table>/<tr>/<td>` — spójnie z istniejącymi akcjami tabeli, przeżywa zapis (HTML→DOCX) i ponowne otwarcie. Patrz ADR-0007.
- GUI: `table-properties-panel` rozbudowany o **dwie zakładki** — „Układ" (akcje strukturalne: wiersze/kolumny, komórki, rozmiar, wygląd, usuń) i „Style" (gotowe style z miniaturami, opcje stylu, obramowania z kolorem/grubością, reset). Panel pozostaje czysto prezentacyjny (`@Input` stanu, `@Output` per akcja). Dodano wewnętrzny pionowy scroll body (`min-height:0` + `overflow-y:auto`, `:host` stretch).
- GUI: `document-editor.ts` — sygnały `activeTablePresetId`/`tableStyleOptions`/`tableBorderColor`/`tableBorderWidth`; metody `applyTableStylePreset`/`toggleTableStyleOption`/`applyTableBorderScope`/`setTableBorderColor`/`setTableBorderWidth`/`resetTableStyle` (glue: util na `activeTable()` + `notifyEditorChange()` → auto-save). `detectTableContext()` czyta stan stylu z aktywnej tabeli (`readActiveTableStyle`).
- GUI: testy `core/utils/table-style.util.spec.ts` (11 przypadków: borders none/outer, preset header/banded, zachowanie treści, recompute toggle, first/last column, reset, persist+read state, defaults) oraz rozszerzone `table-properties-panel.spec.ts` (zakładki, presety, opcje, obramowania, reset).

### Verified
- `npm run build` — OK (tylko istniejące ostrzeżenia budżetu).
- `npx ng test --watch=false` — 41 passed (panel + util + badge). Padają tylko 2 stare scaffoldowe `app.spec.ts` (niezwiązane).

### Notes
- Trwałość: style inline przeżywają round-trip HTML w edytorze i zapis do DOCX. Markery `data-table-style*` (zapamiętany preset/opcje dla panelu) mogą zostać usunięte przez konwersję DOCX — wtedy **wygląd pozostaje** (inline), a panel po ponownym otwarciu pokazuje stan domyślny presetu. Wierność DOCX↔HTML zależy od konwertera serwerowego (to samo ryzyko co istniejące cieniowanie/obramowania).
- Bez zmian backendu/API/modelu dokumentu (funkcja czysto frontendowa).
- MVP w pełni: gotowe style, wiersz nagłówka, wiersze naprzemienne, obramowania (kolor/grubość/zakres), reset. Pełne właściwości DOCX (styl linii dashed/double, padding per komórka, wyrównanie pionowe w UI) — etap późniejszy.

## 2026-05-27 — Konfiguracja tabeli przeniesiona do bocznego panelu

### Changed
- GUI: nowy komponent prezentacyjny `components/table-properties-panel` (`d2-table-properties-panel`) — boczny panel „Ustawienia tabeli" dokowany po lewej, spójny z panelem „Wyszukiwanie" (`.search-panel`). Czysto prezentacyjny: wejścia `hasActiveTable`/`gridLinesVisible`/`shadingColors`, wyjścia per akcja; cała logika operująca na `activeTable()`/`activeTableCell()` pozostaje w `document-editor`.
- GUI: `document-editor.html` — usunięto poziomy pasek `.table-toolbar` (pojawiający się przy `isInTable()`) oraz pływający `.shading-dropdown`. Wstawiono `<d2-table-properties-panel>` w doku obok panelu wyszukiwania; rozpórka poziomej linijki reaguje teraz na `showFindReplace() || showTablePanel()`.
- GUI: `document-editor.ts` — sygnał `showTablePanel` + flaga `tablePanelManuallyClosed`; `detectTableContext()` woła `syncTablePanel()` (auto-otwarcie w tabeli, zamknięcie po wyjściu, respekt ręcznego zamknięcia). `openFindReplace()` przejmuje dok (chowa panel tabeli), `closeFindReplace()` przywraca panel tabeli jeśli karetka nadal w tabeli. `tableDeleteTable()` zamyka panel. Panel wykluczony z czyszczenia zaznaczenia komórek w `onCellMouseDown` (scalanie działa). Wszystkie metody `table*()`, `setCellColor`, `clearCellColor` re-użyte bez zmian logiki.
- GUI (test infra): cel `test` w `angular.json` uzupełniony o `buildTarget`/`tsConfig`/`runner: vitest`/`setupFiles` + nowy `src/test-setup.ts` (Zone.js + zone.js/testing). Wcześniej cel testu nie miał `buildTarget`, więc `ng test` w ogóle się nie uruchamiał.
- GUI: testy `components/table-properties-panel/table-properties-panel.spec.ts` (stan pusty, render sekcji, emisja wszystkich akcji, kolor cieniowania, etykieta linii siatki, `preventDefault` mousedown).

### Verified
- `npm run build` — OK (tylko istniejące ostrzeżenia budżetu bundle/SCSS).
- `npx ng test --watch=false` — panel 6/6 i `document-classification-badge` 4/4 passed (27 passed łącznie). Padają jedynie 2 przestarzałe testy scaffoldowe `app.spec.ts` (asercja „Hello, frontend" + brak providera HttpClient) — niezwiązane z tą zmianą, ujawnione przez naprawę uruchamiania testów.

### Notes
- Decyzja UX dot. konfliktu z wyszukiwaniem: jeden dok po lewej; wyszukiwanie (jawnie wywołane) ma pierwszeństwo, panel tabeli (kontekstowy) nie wypiera go i wraca po zamknięciu wyszukiwania, jeśli karetka jest nadal w tabeli. Patrz ADR-0006.
- Stan pusty panelu („Kliknij tabelę…") jest zabezpieczeniem — przy normalnym flow panel zamyka się po opuszczeniu tabeli, więc rzadko widoczny.
- `app.spec.ts` to stary scaffold (`Hello, frontend`) niepasujący do realnego `App` — do przepisania osobno; nie ruszane w tej zmianie.

## 2026-05-27 — „Wklej tylko tekst" (Paste Text Only) w edytorze WYSIWYG

### Changed
- GUI: nowy util `core/utils/paste-text.util.ts` — czyste, testowalne funkcje `normalizeWhitespace`, `htmlToText` (paragrafy/nagłówki → nowe linie, listy z markerem `•`/`1.`, komórki tabel rozdzielane tabulatorem, `<a>` → tekst bez URL, pomijanie `script`/`style`) oraz `resolvePlainText` (preferuje `text/plain`, fallback z HTML).
- GUI: `wysiwyg-editor.ts` — skrót `Ctrl/Cmd+Shift+V` oznacza najbliższe wklejenie jako „tylko tekst" (okno czasowe 1 s, by nieużyty skrót nie wpłynął na kolejne zwykłe Ctrl+V). `handlePaste` przepuszcza wklejenie przez util; ścieżka „tylko tekst" wstawia przez istniejące `insertText` (zachowuje natywne undo i emituje `onContentChange` → auto-save/`contentChange`). Zwykłe Ctrl+V bez zmian (HTML po `sanitizeHtml`).
- GUI: testy `core/utils/paste-text.util.spec.ts` (17 przypadków: whitespace, nbsp, znaki zero-width, taby/TSV, listy, tabele, linki, script/style).

### Verified
- `npm run build` — OK (tylko istniejące ostrzeżenia budżetu bundle/SCSS).
- `npx vitest run --environment jsdom src/app/core/utils/paste-text.util.spec.ts` — 17/17 passed.

### Notes
- Sanitizacja HTML przy zwykłym wklejaniu nadal opiera się na regexie (`sanitizeHtml`) — patrz R-09 (rekomendacja DOMPurify, wymaga decyzji o zależności).
- Toolbar/menu kontekstowe „Wklej tylko tekst" (poza zdarzeniem paste, przez `navigator.clipboard.readText()`) — nie wdrożone, kolejny inkrement.

## 2026-05-26 — Microsoft Entra ID: uwierzytelnianie + role (docelowo)

### Changed
- **Backend (Internal API):** JwtBearer (`Microsoft.AspNetCore.Authentication.JwtBearer`), sekcja `AzureAd` (Authority/Audience/CorporateKeyClaim/role), `RoleClaimType="roles"`. Policy `RequireAppEmployee`/`RequireAppAdmin`. `[Authorize]` na `BaseApiController` (cała aplikacja za logowaniem; Health anonimowy). Endpointy admina (`GET /`, `deliveries`, `deliveries/{id}/retry`) → `RequireAppAdmin`.
- **Tożsamość:** `ICurrentUserProvider` + `IsAdmin`; `ClaimsCurrentUserProvider` (CorporateKey z konfigurowalnego claimu access tokena, admin z roli). `HttpHeaderCurrentUserProvider` jako DEV. `SystemCurrentUserProvider` w External API.
- **Dostęp do dokumentu:** `DocumentAccessGuard` — **APP_Admin omija `allowedCorporateKeys`** (ADR-0011).
- **Frontend:** MSAL (`@azure/msal-angular`/`-browser`) — `msal.config`, providery w `app.config` (MsalInterceptor + init), `MsalGuard` na całej aplikacji, `appAdminGuard` (UX) na `/admin`, obsługa redirect w `App`. `CurrentUserService` czyta rolę/`ck` z konta MSAL (UX). Usunięto `corporateKeyInterceptor`. Sekcja `auth` w `environment(.development).ts`.

### Verified
- Backend: Domain (63) + Application (155) zielone wcześniej; dodano test bypassu admina. Build pełnej solucji blokowany środowiskowo (DLL zajęte przez API/Rider) — weryfikacja przez projekty testowe.
- **Niezweryfikowane (kroki deployowe):** `dotnet restore` JwtBearer; **frontend wymaga `npm install`** MSAL (do tego czasu „Cannot find module '@azure/msal-angular'" jest oczekiwane); `app.spec.ts` będzie wymagał providerów MSAL w TestBed.

### Notes
- ADR-0011: App Roles (nie groups), CorporateKey z access tokena (claim-only), logowanie dla całej aplikacji, admin omija `allowedCorporateKeys`. Wartości Entra to placeholdery per środowisko (nie sekrety).

## 2026-05-26 — Kontrola dostępu do dokumentu (allowedCorporateKeys) — v1

### Changed
- Domena: `DocumentAccessPolicy` (czysta reguła: brak/null/pusta lista → publiczny dla posiadacza linku; niepusta → tylko pasujący `CorporateKey`; porównanie Trim + Ordinal-IgnoreCase).
- Application: `ICurrentUserProvider` (seam tożsamości), `IDocumentAccessGuard`/`DocumentAccessGuard` (parsuje `allowedCorporateKeys` z `documents.metadata`, stosuje politykę). Rejestracja w DI.
- Api: `HttpHeaderCurrentUserProvider` (v1: czyta nagłówek `X-Corporate-Key`; seam pod Entra ID) + `AddHttpContextAccessor`.
- `Result`/`Result<T>`: nowy stan `Forbidden`/`IsForbidden`.
- Egzekwowanie (403) w handlerach zwracających treść/metadane: `GetDocumentMetadata` (brama GUI), `GetDocument`, `GetDocumentBaseContent`, `GetDocumentVersionContent`. Kontroler mapuje `Forbidden → 403`.
- GUI: `CurrentUserService` (seam, opcjonalny `localStorage('corporateKey')`), `corporateKeyInterceptor` (dodaje `X-Corporate-Key`), `documentAccessGuard` (CanActivate na `/editor` i `/viewer` — blokuje wejście przed inicjalizacją edytora), wspólny `DocumentAccessDeniedComponent` + trasa `/access-denied`. `http-error.interceptor` nie pokazuje toasta dla 403 (obsługuje guard/widok).

### Verified
- `dotnet build D2ViewerEditor.Application.UnitTests` — 0 błędów (testy referencują Domain+Application → potwierdza kompilację warstw). Pozostałe MSB3021/3026 w solucji to blokady DLL przez działające API/Rider, nie błędy kodu.
- Nowe testy: `DocumentAccessPolicyTests` (Domain), `DocumentAccessGuardTests` (Application). Zaktualizowano konstruktory w `GetDocument*`/`GetDocumentMetadata*` testach (nowa zależność guarda).

### Notes
- v1 nie ma realnej tożsamości — `CorporateKey` z nagłówka; dokumenty bez `allowedCorporateKeys` pozostają publiczne (brak regresji). Wszystkie odmowy → 403 (401 zarezerwowany na fazę Entra ID). Backend = źródło prawdy; front tylko blokuje wejście i pokazuje widok.

## 2026-05-25 — Panel admina „Wysyłki" + uruchomienie migracji DB

### Changed
- GUI: nowy widok `/admin/deliveries` (`admin-deliveries` component) — lista zadań wysyłki ze statusem, liczbą prób, datami (utworzono/ostatnia/następna/deadline), `locked_by` i ostatnim błędem; przycisk „Ponów" dla `DeadLettered`/`FailedPermanently`. Filtr per kolumna + dropdown statusu (steruje zapytaniem do API) + paginacja klienta, wizualnie spójny z `/admin/files`. Statusy tłumaczone na PL w warstwie prezentacji (enum bez zmian). Link w `admin-shell` nav. Trasa-dziecko w `app.routes.ts`.
- GUI: `document-storage.service.ts` — `getDeliveriesByStatus(status, skip, take)` + `retryDelivery(id)` + interfejsy `DeliveryListItem`, `RequeueDeliveryResult`.
- Application: `DeliveryListItemDto` rozszerzony o `LockedUntil` i `LockedBy` (monitoring claimu workera); mapowanie w `GetDeliveriesByStatusQueryHandler`.
- DB: wykonano skrypty `infra/sql/001`–`007` na lokalnym PostgreSQL (Podman `d2viewereditor_postgres`). 001–006 już istniały (idempotentne), `007` utworzył tabelę `document_deliveries`.

### Verified
- `dotnet build D2ViewerEditor.sln` — 0 błędów (ostrzeżenia tylko MSB3026 — zablokowane DLL przez działające API/Rider).
- Weryfikacja schematu: `document_deliveries` z 20 kolumnami, 4 indeksami (w tym partial `ux_..._active_per_document`, `ix_..._due`, `ix_..._stuck`), CHECK statusu, FK do `documents`.

### Notes
- Zamyka rekomendację „panel admina dla `DeadLettered`" z `TASK_HANDOFF`. Filtry tekstowe działają po stronie klienta na pobranej stronie (jak `admin-files`).

## 2026-05-25 — Funkcja „Zakończ i wyślij" (asynchroniczna wysyłka na returnUrl)

### Changed
- Domena: nowy agregat `DocumentDelivery` (+ enum `DeliveryStatus`: Pending/Sending/RetryScheduled/Sent/FailedPermanently/DeadLettered), interfejsy `IDocumentDeliveryRepository`, `IDeliverySender`, `IBackoffStrategy`. `IDocumentStorageService.UploadRawAsync` (snapshot pod dowolną nazwą).
- Application: `FinishAndSendDocumentCommand` (+handler+validator), `GetDeliveryStatusQuery`, `RequeueDeliveryCommand`, `GetDeliveriesByStatusQuery`. Idempotencja wielokrotnego kliknięcia (aktywne zadanie per dokument).
- Infrastructure: `DocumentDeliveryRepository` (claim `FOR UPDATE SKIP LOCKED` + lease), `HttpDeliverySender` (Idempotency-Key), `ExponentialJitterBackoff`, `DeliveryAttemptRunner`, `DocumentDeliveryWorker` (BackgroundService, bounded concurrency, reclaim po crashu). Rejestracja w DI + `DeliveryWorkerOptions` (sekcja `DeliveryWorker`).
- DB: `infra/sql/007_add_document_deliveries.sql` (tabela + indeksy + unique partial idempotencji + check status). Mapowanie EF `DocumentDeliveryConfiguration` + `DbSet`.
- API (`DocumentStorageController`): `POST {masterId}/versions/{versionId}/finish` (202), `GET deliveries/{id}`, `GET deliveries?status=`, `POST deliveries/{id}/retry`.
- GUI: `finishDocument()` (utrwala stan + wysyłka + polling co 4 s do stanu końcowego), serwis `finishAndSend`/`getDeliveryStatus`, sygnały `isFinishing/deliveryStatus/deliveryId`.

### Verified
- `dotnet build D2ViewerEditor.sln` — OK (0 błędów).
- `dotnet test` — nowe testy (DocumentDelivery, handler FinishAndSend, validator, backoff) zielone. Jeden WCZEŚNIEJSZY błąd niezwiązany: `SaveDocumentVersionCommandHandlerTests.Handle_ValidCommand_ShouldAddNewVersion` (oczekuje `UpdateAsync`, handler go nie woła po przejściu na płaski zapis) — nie ruszany.
- GUI: TypeScript kompiluje się; `npm run build` blokuje WCZEŚNIEJSZY budżet `document-editor.scss` (35.7 kB > 32 kB), niezwiązany z tą zmianą.

### Notes
- Snapshot finalny zamrażany w GCS jako `deliveries/{deliveryId}` (nie wysyłamy później zmodyfikowanej v2). At-least-once + Idempotency-Key. Worker odporny na restart/multi-instancję; `DeliveryWorker:Enabled=false` pozwala wydzielić wysyłkę na dedykowany host.

## 2026-05-25 — Pionowa linijka per strona + lupa otwiera panel

### Changed
- Pionowa linijka: zamiast jednej „zamrożonej" renderowana jest OSOBNA linijka na każdą stronę (`pageList`), z geometrią kartek (segment = wysokość strony*zoom, odstęp = separator 8px*zoom). Pasek scrolluje 1:1, więc linijka restartuje się na granicy stron. Segment ma `overflow:hidden` (ucina ~1px nadmiaru axisPx 1123 vs strona 1122).
- Lupa w toolbarze otwiera panel „Wyszukiwanie" (nowy output `openSearch`), usunięto stary pasek wyszukiwania toolbaru (`showSearchBar`).

### Verified
- `ng build` OK.

### Notes
- `onEditorScroll` używa przybliżonego `PAGE_GAP=40` do wskaźnika „Strona X z Y" (realny separator to 8px) — nietykane, dotyczy tylko zaokrąglenia numeru strony, nie linijki.

## 2026-05-25 — Fix 400 /save (EF) + panel „Wyszukiwanie" zamiast dialogu

### Changed
- BACKEND FIX: `DocumentVersionConfiguration` — `Id` ma `ValueGeneratedNever()`. Bez tego EF (konwencja ValueGeneratedOnAdd dla Guid) traktował nową wersję z ręcznie ustawionym kluczem dodaną do śledzonego dokumentu jako `Modified` → UPDATE w 0 wierszy → `DbUpdateConcurrencyException` → POST `/{master}/save` zwracał 400. To blokowało tworzenie wersji edytowalnej przy ręcznym otwarciu z dysku/dashboardu. (Zmiana tylko w modelu EF, bez migracji schematu.) **Wymaga restartu API.**
- Dialog Znajdź/zamień → lewy panel `.search-panel` „Wyszukiwanie" (jak panel Nawigacja w Word): tylko wyszukiwanie (bez zamiany), bez zakładek Nagłówki/Strony — tylko sekcja „Wyniki". Lista wyników z fragmentami kontekstu; klik = skok do trafienia (`goToResult`→`editor.goToMatch`). Nowe API edytora: `getSearchSnippets()`, `goToMatch(index)`.
- EDYTUJ → pozycja „Znajdź" (Ctrl+F) w obu trybach.

### Verified
- Frontend `ng build` OK; backend `dotnet build` Infrastructure OK. Przyczynę 400 potwierdzono w logach API (DbUpdateConcurrencyException, UPDATE nowej wersji zamiast INSERT).

## 2026-05-25 — Wyszukiwanie po wszystkich stronach + naprawa dialogu Znajdź

### Changed
- `WysiwygEditorComponent.searchText` przeszukuje teraz WSZYSTKIE strony (`pageEditorRefs`) w kolejności dokumentu, nie tylko aktywną; agreguje trafienia do jednej listy `searchHighlights` (findNext/findPrevious nawigują między stronami, `scrollIntoView` przewija). Dotyczy też wyszukiwarki z toolbaru.
- `emitContentChange` po zamianie agreguje treść ze wszystkich stron (`getContent()`), więc zamiana na nieaktywnej stronie jest utrwalana.
- Dialog Znajdź (`DocumentEditorComponent`) przepięty z ułomnego `window.find()` na API edytora: wyszukiwanie na żywo (`onFindInput`), licznik „x z y", przyciski „Poprzednie/Następne", czyszczenie podświetleń przy zamknięciu (`closeFindReplace`).

### Verified
- `ng build --configuration development` OK.

## 2026-05-25 — Prowadnice marginesów + wyszukiwanie w read-only

### Changed
- „Linie marginesów" (martwy przełącznik — styl bez markupu) zaimplementowane: `wysiwyg-editor.html` renderuje `.margin-guides` (4 przerywane linie) per strona z `pageMargins()`, sterowane `[showMarginGuides]`.
- Read-only: cały `d2-editor-toolbar` ukryty (znika lupa). Wyszukiwanie przez EDYTUJ → „Znajdź" (zamiast „Znajdź i zamień") oraz skrót Ctrl+F (`onGlobalKeydown` w `DocumentEditorComponent`; Ctrl+H tylko gdy edycja dozwolona).
- Dialog `showFindReplace`: w trybie zablokowanym tylko pole „Znajdź" + „Znajdź następny" (ukryte „Zamień na"/„Zamień"/„Zamień wszystko"); nagłówek „Znajdź".

### Verified
- `ng build --configuration development` OK.

### Notes
- `showMarginGuides` domyślnie `true` → prowadnice widoczne domyślnie (można wyłączyć w WIDOK/FORMATUJ). Renderu nie potwierdzałem wizualnie.

## 2026-05-25 — Ukrywanie funkcji edycyjnych w trybie read-only / „zajęty"

### Changed
- `DocumentEditorComponent`: nowy `lockedByOther` (hook na backendowy `DocumentStatus`=Editing, domyślnie false) i computed `editingDisabled = readOnly() || lockedByOther()`.
- `EditorToolbarComponent`: nowy `@Input() readOnly` — przy true ukrywa undo/redo i wszystkie grupy edycyjne (styl, czcionka, format, wyrównanie, listy, wstawianie); zostaje wyszukiwarka.
- Menu edytora: ukryte WSTAW/FORMATUJ/NARZĘDZIA w całości; w PLIK ukryte Zapisz + Ustawienia strony; w EDYTUJ ukryte cofnij/ponów/wytnij/wklej/usuń (zostają Kopiuj, Zaznacz wszystko, Znajdź i zamień, Właściwości, Podpisy). WIDOK bez zmian. Toolbar tabeli ukryty (`!editingDisabled()`). Plakietka rozróżnia „ktoś inny edytuje".

### Verified
- `ng build --configuration development` OK.

### Notes
- Blokada „ktoś inny edytuje" wymaga podpięcia statusu dokumentu do edytora (`lockedByOther.set(...)`); obecnie edytor zna tylko read-only z braku `versionId`. Skróty klawiszowe nie są blokowane — chroni je `readOnly` na contenteditable.

## 2026-05-25 — Linijka: linia prowadząca + obrazowanie nagłówka/stopki

### Changed
- `RulerComponent` emituje `dragGuideChange` (active + axis + offset px @100% od krawędzi strony). Linia prowadząca (jak w MS Word) renderowana w `.paper-container` edytora, bo wewnątrz `.ruler-h-bar` (22px, `overflow:hidden`) była przycinana.
- `WysiwygEditorComponent` emituje `editingSectionChange` ('header'|'footer'|'body').
- `DocumentEditorComponent`: `verticalRulerMargins` przełącza obrazowanie pionowej linijki na pasmo nagłówka/stopki podczas ich edycji.
- Pasmo nagłówka/stopki ma `min-height` i rośnie z treścią (obraz), więc położenie na linijce pochodzi z POMIARU DOM: `WysiwygEditorComponent.sectionGeometryChange` (cm od góry strony, `ResizeObserver` re-emituje przy zmianie wysokości; skala liczona z szerokości strony — odporna na wzrost pasma).
- Pionowa linijka w trybie nagłówka/stopki: przeciągnięcie uchwytu zmienia WYSOKOŚĆ pasma (`setHeaderHeight`/`setFooterHeight`), nie marginesy strony. W trybie `body` działa jak wcześniej.

### Verified
- `ng build --configuration development` OK (bez błędów).

### Notes
- Marginesy góra/dół: ekstrakcja w `DocxToHtmlConverter` poprawna (1/567 ≈ 2.54/1440). Treść startuje na `marginTop` (header band + padding). Rozbieżność z Wordem dotyczy raczej POZYCJI pasma nagłówka (renderowane przy `top:0`, nie na „header from edge"=pgMar.Header), nie wysokości marginesu treści. Do potwierdzenia na konkretnym pliku DOCX.

## 2026-05-23 — Dostosowanie `.ai/` do realnego projektu

### Changed
- Przyjęto strukturę `.ai/` wg `template/` (INDEX, PROJECT_CONTEXT, TECH_STACK, ARCHITECTURE, DOMAIN, FEATURES, DATABASE, API_CONTRACTS, SECURITY, TESTING_QUALITY, DEVOPS_DEPLOYMENT, DECISIONS, RISKS_ASSUMPTIONS, GLOSSARY, CURRENT_STATE, TASK_HANDOFF, CHANGELOG, PRODUCT_GOALS + pliki-wytyczne i `templates/`).
- Wypełniono faktami z repo; oznaczono niepewności (port External API, brak compose/CI).
- Dodano `CLAUDE.md` i `AGENTS.md` w root.
- Usunięto stare `.ai/{CONTEXT,API,BACKEND,FRONTEND}.md` (treść przeniesiona do nowej struktury).

### Verified
- Pliki zapisane; fakty zebrane z `*.csproj`, `package.json`, `docker/*`, `launchSettings`, `appsettings`, `infra/sql/`.

### Notes
- Pliki ogólne (BACKEND_DOTNET, FRONTEND_ANGULAR, CODING_STANDARDS, AI_AGENT_WORKFLOW, PROMPTS, README, templates) pozostawiono zgodne ze wzorcem.

## 2026-05-23 — Płaski zapis wersji + auto-save + ujednolicony „Zapisz"

### Changed
- Backend: `Document.UpdateVersion`, `DocumentVersion.UpdateContent`/`ModifiedAt`, `UpdateDocumentVersionCommand` + `PUT /api/documentstorage/{masterId}/versions/{versionId}`, `GetDocumentMetadataQuery` + `GET .../{masterId}/metadata`, skrypt `infra/sql/005_add_version_modified_at.sql` + mapowanie EF.
- Frontend: mechanizm auto-save (timer + `environment.autoSave`), switch „AutoSave", ujednolicony `saveDocument()` (API), `downloadDocument()` (pobranie lokalne), usunięto `saveDocumentAs()`.

### Verified
- `dotnet build D2ViewerEditor.sln` — OK (0 błędów).
- `npm run build` (GUI) — OK.

### Notes
- Tryb podglądu (Krok 2) wciąż ładuje aktywną wersję zamiast v1 — do dokończenia.
- `finishDocument()` (zwrot na returnUrl) — TODO.
