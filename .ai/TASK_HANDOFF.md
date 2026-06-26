# Task Handoff

> Bezpieczne przekazanie pracy kolejnej sesji/agentowi.

## Ostatnia aktualizacja (2026-06-26)

- corpKey ustalany wyłącznie po stronie API z tokenu — usunięty transport z GUI.
   - Powód: GUI czytało `corpKey` ze statycznego snapshotu `getActiveAccount().idTokenClaims` (często null) → `SaveDocumentVersionCommand.CorporateKey` przychodził null. Token wysyłany do API (ID token) i tak niesie `corpKey`.
   - API: usunięty parametr `CorporateKey` z komend `SaveDocumentVersion`/`UpdateDocumentVersion`/`FinishAndSendDocument` i z controller DTO `SaveDocumentVersionRequest`; handlery czytają `_currentUser.CorporateKey` (brak → `Result.Failure`).
   - GUI: usunięty sygnał `corporateKey` + odczyt claimu w `document-editor.ts`; zapisy wysyłają `{ content }`. Listy admina (`corporateKey` z `DocumentDelivery`) bez zmian.
   - Weryfikacja: backend build 0 błędów; `Application.UnitTests` 295/295; `Api.UnitTests` 113/113; GUI build OK. Decyzja: ADR-0021.

## Ostatnia aktualizacja (2026-06-26)

- Naprawa zapisu tożsamości edytującego (corpKey był NULL) + przycisk „Kopiuj link" w `admin-files`.
   - Przyczyna NULL: (1) Entra — `AzureAdOptions.CorporateKeyClaim` domyślnie `"ck"`, a token niesie claim `corpKey`; (2) dev bypass — `HttpHeaderCurrentUserProvider` czytał nagłówek `X-Corporate-Key`, którego GUI nie wysyła.
   - Kierunek (decyzja użytkownika): **frontend wysyła `corpKey`** (z `idTokenClaims`), backend używa go z fallbackiem do tokenu i twardym błędem, gdy nie da się ustalić użytkownika.
   - API: `AzureAdOptions.CorporateKeyClaim` `"ck"`→`"corpKey"`; `HttpHeaderCurrentUserProvider` fallback `DEV-LOCAL`; handlery `SaveDocumentVersion`/`UpdateDocumentVersion` wstrzykują `ICurrentUserProvider`, wszystkie trzy (`Save`/`Update`/`FinishAndSend`) liczą `corporateKey = request.CorporateKey ?? _currentUser.CorporateKey` i zwracają `Result.Failure`, gdy brak.
   - GUI: bez zmian w wysyłce corpKey (już wysyła). `admin-files` — przycisk „Kopiuj link" (kopiuje link `/editor?masterId&versionId` do schowka, baner `notice`), widoczny dopiero w szczegółach pozycji po rozwinięciu.
   - Weryfikacja: backend build 0 błędów; `Application.UnitTests` 295/295; `Api.UnitTests` 113/113; GUI build OK; `ng test` 268/269 (pre-existing fail `spec-layout-shell.spec`).

## Ostatnia aktualizacja (2026-06-24)

- corpKey z tokenu Entra ID przekazywany przy zapisach z edytora; kolumna „Kto modyfikował" w obu listach admina.
   - GUI: `document-editor.ts` czyta claim `corpKey` i wysyła `corporateKey` w body (`save`/`update`/`finish`).
   - API: DTO/komendy +`CorporateKey` (opcjonalne, fallback do claimu); `Document.LastModifiedBy` + migracja `infra/sql/011_add_document_last_modified_by.sql` (uruchomić na bazach).
   - Listy: `DocumentListItemDto.LastModifiedBy`, `DeliveryListItemDto.CorporateKey`; kolumny + filtry w `admin-files` i `admin-deliveries`.
- Weryfikacja: backend build 0 błędów; `D2Api.Api.UnitTests` 113/113; `Application.UnitTests` 295/295; GUI build OK; `ng test` 268/269 (jeden pre-existing fail `spec-layout-shell.spec`, niezwiązany).
- TODO przy deployu: zastosować migrację SQL `011` na środowiskach (brak EF migrations w tym repo — raw SQL w `infra/sql/`).

## Ostatnia aktualizacja (2026-06-21)

## Ostatnia aktualizacja (2026-06-22)

- Wdrożono centralne security policies dla uploadu i callback URL:
   - `IReturnUrlValidator` + `ReturnUrlValidator` (kontrola: schemat, znaki, protocol-relative, user-info, loopback/private IP, allowlista hostów, normalizacja URL),
   - `IFileUploadSecurityService` + `FileUploadSecurityService` (spójność extension↔MIME↔signature, inspekcja DOCX ZIP: zip-slip, limity, wymagane part-y OOXML, blokada VBA),
   - `IFileScanner` (abstrakcja AV) + `NoOpFileScanner` (domyślny adapter).
- Podpięto egzekwowanie polityk w handlerach:
   - uploady: `UploadDocument`, `UploadImage`, `IngestExternalDocument`,
   - callback/recipient: `UpdateCallbackUrl`, `UpdateDeliveryRecipientUrl`, `FinishAndSendDocument`.
- Hosty (`D2Api`, `D2Services`) bindują sekcje konfiguracyjne:
   - `Security:Upload`,
   - `Security:ReturnUrl`.
- Testy po zmianach:
   - Application.UnitTests **295/295**,
   - D2Api.Api.UnitTests **110/110**,
   - D2Services.Api.UnitTests **39/39**.
- Dodane nowe testy security:
   - `ReturnUrlValidatorTests`,
   - `FileUploadSecurityServiceTests`.

- D2Services observability domknięte do standardu:
   - pipeline zawiera `UseRequestObservability()` + `UseExceptionHandlingMiddleware()`,
   - `X-Correlation-ID` propagowany i logowany (scope + LogContext),
   - `ProblemDetails` zawiera `correlationId`, a błędy walidacji serializują `errors`,
   - formatter JSON rozszerzony o pola `level/service/environment/traceId/spanId` i właściwości scope/eventu.
- Testy D2Services unit: **28/28 pass** (`RequestObservabilityMiddlewareTests`, rozszerzone formatter/middleware/extensions tests).

- Zwiększono liczbę testów jednostkowych w obu backendach:
   - **D2ApiViewerEditor**: +3 testy middleware (`RequestObservabilityMiddleware`, `ExceptionHandlingMiddleware`),
   - **D2ServicesViewerEditor**: nowy projekt `D2ServicesViewerEditor.Api.UnitTests`, rozszerzony do 22 testów (`DocumentController`, `ExceptionHandlingMiddleware`, `GcpJsonSerilogFormatter`, `HealthController`, `MiddlewareExtensions`).
- `D2ServicesViewerEditor.sln` zawiera teraz projekt testowy `D2ServicesViewerEditor.Api.UnitTests`.
- Weryfikacja wykonana:
   - `dotnet test D2ApiViewerEditor/D2ViewerEditor.Api.UnitTests/D2ViewerEditor.Api.UnitTests.csproj` → 76/76 pass,
   - `dotnet test D2ServicesViewerEditor/D2ServicesViewerEditor.Api.UnitTests/D2ServicesViewerEditor.Api.UnitTests.csproj` → 22/22 pass.

## Aktualne zadanie

Dokończyć przepływ Krok 1–3 dla D2 ViewerEditor: ingest (gotowy), płaski zapis + auto-save (gotowy), tryb podglądu (Krok 2, w toku), funkcja „Zakończ" (planowana).

## Kontekst

System przyjmuje dokumenty od aplikacji zewnętrznej (External API), edytuje/ogląda w GUI, zapisuje wersjonowanie w DB+GCS przez Internal API. Oryginał (v1) nietykalny; edycja nadpisuje v2 w miejscu.

## Pliki, które trzeba znać

| Plik / obszar | Dlaczego |
|---|---|
| `.ai/PROJECT_CONTEXT.md`, `.ai/DOMAIN.md` | model i reguły (v1 immutable, płaski zapis) |
| `.ai/API_CONTRACTS.md` | katalog endpointów obu API |
| `D2ApiViewerEditor/.../Features/Documents/Commands/UpdateDocumentVersion/` | nadpisanie wersji |
| `D2ApiViewerEditor/.../Features/Documents/Commands/IngestExternalDocument/` | ingest |
| `D2GuiViewerEditor/src/app/components/document-editor/document-editor.ts` | edytor, auto-save, zapis |
| `D2GuiViewerEditor/src/app/services/document-storage.service.ts` | klient storage API |
| `infra/sql/` | schemat (ręczny SQL) |

## Ostatnie wykonane kroki

1. Dodano płaski zapis (`UpdateVersion` + PUT endpoint + `ModifiedAt` + SQL 005).
2. Dodano endpoint metadanych.
3. GUI: auto-save + switch, ujednolicony „Zapisz" / „Pobierz dokument".
4. Przepisano `.ai/` wg wzorca z `template/`.

## Następne sugerowane kroki

1. Podpiąć produkcyjny silnik AV pod `IFileScanner` (ICAP/ClamAV/API) oraz spiąć alerting dla `MalwareDetected`/`MalwareScanUnavailable`.
2. Wymusić i udokumentować politykę host allowlist (`Security:ReturnUrl:AllowedHosts`) na środowiskach prod/UAT.
3. Krok 2 (podgląd): w GUI dla `?masterId=` ładować v1 przez `GET .../{masterId}/download` zamiast aktywnej wersji; routing PDFViewer vs DocxEditor po `mimeType` z `GET .../{masterId}/metadata` lub `/{masterId}`.
2. (Zrobione 2026-05-25) `finishDocument()` — async wysyłka na returnUrl: kolejka `document_deliveries` + worker `DocumentDeliveryWorker` + snapshot GCS `deliveries/{id}` + retry/backoff 24h. Endpointy `finish`/`deliveries`. Patrz ADR-0005.
   - (Zrobione 2026-05-25) Panel admina `/admin/deliveries` (lista + filtry + retry + `locked_by`).
   - Pozostało: dodać DB integration testy claimu (`FOR UPDATE SKIP LOCKED` / reclaim po crashu) na realnym Postgresie (R-06); rozważyć ochronę SSRF dla `RecipientUrl` (R-07).
3. (Zamknięte) Port External API = 15112 wg `appsettings.json` `Urls`; Swagger `/swagger`.
4. Dodać testy: domena `UpdateVersion` (v1 immutable), handler `UpdateDocumentVersionCommand`, serwis GUI auto-save.

## Niedokończone zmiany

| Obszar | Co | Ryzyko |
|---|---|---|
| Krok 2 podgląd | ładuje v2 zamiast v1 | użytkownik podglądu widzi edytowalną, nie oryginał |
| Testy integ. claimu wysyłki | brak (R-06) | regresje w `FOR UPDATE SKIP LOCKED` niewykryte jednostkowo |
| Ochrona SSRF `RecipientUrl` | brak (R-07) | worker robi żądania na adres z danych zewnętrznych |

## Komendy weryfikacji

```bash
dotnet build D2ApiViewerEditor/D2ViewerEditor.sln
cd D2GuiViewerEditor && npm run build
```

Wynik ostatnio: oba OK (0 błędów).

## Nie rób bez zgody

- Nie generuj migracji EF (schemat = `infra/sql/`).
- Nie zmieniaj kontraktów API ani portów bez ustaleń.
- Nie ruszaj reguły v1-immutable / płaskiego zapisu.
- Nie aktualizuj major .NET/Angular. Nie commituj sekretów.
