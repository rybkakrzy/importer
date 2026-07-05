# Architecture

## Cel architektury

Architektura ma pozwolić rozwijać system przez ludzi i agentów AI bez chaosu — jasne granice, przewidywalny kierunek zależności, testowalna logika.

## Obraz systemu (realny)

```text
Aplikacja zewnętrzna (źródłowa)
  | multipart POST + metadane (returnUrl, classification)
  v
D2ServicesViewerEditor  (External API)
  | wywołuje IngestExternalDocumentCommand (warstwa Application D2ViewerEditor)
  v
DB (PostgreSQL: metadane) + GCS (pliki binarne)
  ^
  | REST (JSON / multipart / file)
D2ApiViewerEditor  (Internal API)
  ^
  | HTTP
D2GuiViewerEditor  (Angular SPA) — PDFViewer / DocxEditor
```

Aplikacja zewnętrzna otwiera GUI po URL (`?masterId=...[&versionId=...]`); GUI rozmawia wyłącznie z internal API.

## Backend — realny podział (`D2ApiViewerEditor/`)

```text
D2ViewerEditor.sln
├── D2ViewerEditor.Domain          — encje (Document, DocumentVersion), interfejsy, modele, Result<T>
├── D2ViewerEditor.Application     — CQRS handlers (Features/...), walidatory, behaviours (Logging, Validation)
├── D2ViewerEditor.Infrastructure  — EF Core (DbContext, repo, configurations), GCS, konwertery DOCX↔HTML, podpisy, barcode, worker wysyłki (Services/Delivery)
├── D2ViewerEditor.Api             — kontrolery, Program.cs, DI
├── *.UnitTests (Api/Application/Domain/Infrastructure) + D2ViewerEditor.Benchmarks
└── D2ViewerEditor.Api.IntegrationTests (WebApplicationFactory — testy autoryzacji 401/403)
    + D2ViewerEditor.Infrastructure.IntegrationTests
```

`D2ServicesViewerEditor.Api` to osobny, cienki host (External API) — kontroler ingestu, który deleguje do warstwy Application z `D2ApiViewerEditor` (komenda `IngestExternalDocumentCommand`). Ma własny projekt testowy `D2ServicesViewerEditor.Api.UnitTests`. Osobno w repo: `D2ExampleExternalApp` — przykładowy integrator (upload + callback) do testów end-to-end, poza produktem.

### Granice warstw

- **Api** — routing, DTO request/response, statusy HTTP, mapowanie. Bez ciężkiej logiki.
- **Application** — przypadki użycia (komendy/zapytania), orkiestracja, walidacja aplikacyjna, dostęp do repo/storage przez abstrakcje.
- **Domain** — model i reguły (np. `Document.AddVersion`/`UpdateVersion`/`RestoreVersion`), inwarianty. Zero zależności infrastrukturalnych.
- **Infrastructure** — EF Core, GCS, konwertery, podpisy.

Kierunek: `Api → Application → Domain`, `Api → Infrastructure`, `Infrastructure → Domain/Application abstractions`.

### Przetwarzanie w tle (worker wysyłki)

Funkcja „Zakończ i wyślij" wprowadza asynchroniczne przetwarzanie:

- `DocumentDeliveryWorker : BackgroundService` (`Infrastructure/Services/Delivery`) — hostowany w procesie **Internal API** (`D2ApiViewerEditor.Api`), rejestrowany przez `AddHostedService` w `Infrastructure/DependencyInjection`. Włączany flagą `DeliveryWorker:Enabled` (pozwala wydzielić wysyłkę na dedykowaną instancję bez zmian kodu).
- Worker cyklicznie claimuje paczki zadań z tabeli `document_deliveries` (`SELECT ... FOR UPDATE SKIP LOCKED` + lease `locked_until`) i przetwarza je równolegle (`Parallel.ForEachAsync`, limit `MaxConcurrency`), każde zadanie w osobnym DI scope (`DeliveryAttemptRunner`).
- Wysyłka HTTP: typed `HttpClient` (`HttpDeliverySender`, `IDeliverySender`). Retry: `ExponentialJitterBackoff` (`IBackoffStrategy`). Pliki: niezmienny snapshot w GCS (`IDocumentStorageService.UploadRawAsync`/`DownloadAsync`).
- Odporność: wiele instancji (SKIP LOCKED), restart (lease — zawieszone `Sending` przejmowane po wygaśnięciu), brak utraty zadań (stan w PostgreSQL). Szczegóły decyzji: `DECISIONS.md` ADR-0005.

## Frontend — realny podział (`D2GuiViewerEditor/src/app/`)

```text
src/app/
├── app.ts / app.routes.ts / app.config.ts   — root, routing, providery
├── components/   — document-editor (główny), wysiwyg-editor, editor-toolbar, barcode-dialog,
│                   file-upload, offline-banner, ruler, table-properties-panel,
│                   header-footer-panel, image-properties-panel, global-banners,
│                   environment-banner, document-classification-badge
├── pages/        — dashboard, document-viewer, pdf-viewer (lazy), pdf-maintenance,
│                   access-forbidden (lazy, trasa /brak-uprawnien),
│                   admin/{admin-shell, admin-files, admin-deliveries, admin-dashboard} (lazy)
├── guards/       — auth.guard (MsalGuard/bypass), resource.guard (obszary z GET /identity/resources),
│                   home-redirect.guard (root: Operator→dashboard, Admin→/admin),
│                   document-access.guard (per-dokument 403)
├── services/     — document.service (stateless file API), document-storage.service (CRUD+GCS),
│                   barcode.service, admin.service, file-upload
├── core/         — auth (msal.config, api-token-primer), config (runtime-config),
│                   services (api-config, build-info, connection-status, document-navigation,
│                   notification, resource-access), interceptors (api-token, http-error),
│                   error-handling (global handler), utils (paste-text, return-url,
│                   table-context, table-style, environment-banner)
└── models/document.model.ts
```

Wzorzec: komponenty trzymają stan w Signals; HTTP idzie przez serwisy; `api-config.service` centralizuje `apiUrl`. Routing bramkują guardy (backend = źródło prawdy przez `GET /api/identity/resources`; guardy to UX).

## Granice odpowiedzialności

| Element | Powinien | Nie powinien |
|---|---|---|
| Angular component | prezentacja, lokalny stan UI (Signals) | bezpośredni `HttpClient`, logika domenowa |
| Angular service | komunikacja z API, orkiestracja | renderowanie UI |
| API controller | HTTP, mapping, statusy | reguły biznesowe |
| Application handler | orkiestracja use-case | szczegóły UI/HTTP |
| Domain (Document) | reguły wersji, inwarianty | dostęp do DB/GCS/HTTP |
| Infrastructure | EF/GCS/konwersje | decyzje biznesowe |

## Zakazane skróty

- Logika biznesowa w kontrolerach.
- Bezpośrednie `HttpClient` w komponentach (jest warstwa serwisów).
- Mieszanie DTO API z encją domenową.
- Generowanie migracji EF (schemat = ręczny SQL w `infra/sql/`).
- Refaktor całej struktury przy okazji jednej funkcji.
