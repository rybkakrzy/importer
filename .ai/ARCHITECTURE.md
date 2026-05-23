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
├── D2ViewerEditor.Infrastructure  — EF Core (DbContext, repo, configurations), GCS, konwertery DOCX↔HTML, podpisy, barcode
├── D2ViewerEditor.Api             — kontrolery, Program.cs, DI
└── *.UnitTests (Api/Application/Domain/Infrastructure) + D2ViewerEditor.Benchmarks
```

`D2ServicesViewerEditor.Api` to osobny, cienki host (External API) — kontroler ingestu, który deleguje do warstwy Application z `D2ApiViewerEditor` (komenda `IngestExternalDocumentCommand`).

### Granice warstw

- **Api** — routing, DTO request/response, statusy HTTP, mapowanie. Bez ciężkiej logiki.
- **Application** — przypadki użycia (komendy/zapytania), orkiestracja, walidacja aplikacyjna, dostęp do repo/storage przez abstrakcje.
- **Domain** — model i reguły (np. `Document.AddVersion`/`UpdateVersion`/`RestoreVersion`), inwarianty. Zero zależności infrastrukturalnych.
- **Infrastructure** — EF Core, GCS, konwertery, podpisy.

Kierunek: `Api → Application → Domain`, `Api → Infrastructure`, `Infrastructure → Domain/Application abstractions`.

## Frontend — realny podział (`D2GuiViewerEditor/src/app/`)

```text
src/app/
├── app.ts / app.routes.ts / app.config.ts   — root, routing, providery
├── components/   — document-editor (główny, ~3.3k linii), wysiwyg-editor, editor-toolbar,
│                   barcode-dialog, file-upload, offline-banner, ruler
├── pages/        — dashboard, document-editor (wrapper), pdf-viewer (lazy), pdf-maintenance,
│                   admin/{admin-shell, admin-files, admin-dashboard} (lazy)
├── services/     — document.service (stateless file API), document-storage.service (CRUD+GCS),
│                   barcode.service, admin.service, file-upload
├── core/         — services (api-config, build-info, connection-status, document-navigation,
│                   notification), interceptors (http-error), error-handling (global handler)
└── models/document.model.ts
```

Wzorzec: komponenty trzymają stan w Signals; HTTP idzie przez serwisy; `api-config.service` centralizuje `apiUrl`.

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
