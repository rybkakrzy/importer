# Tech Stack - D2ViewerEditor (Detailed)

## 1. Frontend (D2GuiViewerEditor)

### Core
- `@angular/*` `^20.0.0`
- `typescript` `~5.8.2`
- `rxjs` `~7.8.0`
- `zone.js` `^0.15.1`

### Supporting
- `pdfjs-dist` `^5.5.207` (narzędzia renderowania związane z PDF)
- `vitest` `^3.1.1`
- `jsdom` `^27.1.0`

### Uwagi architektoniczne
- Architektura standalone components
- Lokalny stan oparty o signals w komponentach edytora
- Podział environment: development vs production

## 2. Backend (D2ApiViewerEditor)

### API Layer
- Target framework: `net8.0`
- `Swashbuckle.AspNetCore` `6.5.0`

### Application Layer
- `MediatR` `12.4.1`
- `FluentValidation` `11.11.0`
- `FluentValidation.DependencyInjectionExtensions` `11.11.0`

### Infrastructure Layer
- `DocumentFormat.OpenXml` `3.0.2`
- `HtmlAgilityPack` `1.11.61`
- `Microsoft.EntityFrameworkCore` `8.0.12`
- `Npgsql.EntityFrameworkCore.PostgreSQL` `8.0.11`
- `Google.Cloud.Storage.V1` `4.10.0`
- `ZXing.Net` `0.16.9`
- `SkiaSharp` `3.116.1`
- `SkiaSharp.NativeAssets.Linux` `3.116.1`

## 3. Data and Persistence

### Database
- PostgreSQL 16
- Główne tabele: `documents`, `document_versions`
- Partial i unique indexes dla pobierania aktywnej wersji

### Object Storage
- DEV: fake-gcs-server (`http://localhost:4443`)
- Non-DEV: endpoint kompatybilny z Google Cloud Storage API

## 4. Test Stack

### .NET
- NUnit
- FluentAssertions
- NSubstitute

### Python E2E
- pytest
- pytest-bdd
- playwright

### Performance
- Locust z plikami scenariuszy:
	- `health_locust.py`
	- `barcode_locust.py`
	- `document_locust.py`
	- `all_endpoints_locust.py`

## 5. DevOps and Local Infra

- Podman engine
- podman-compose orchestration
- Kontener frontend Nginx (`docker/gui.dockerfile`)
- Kontener API (`docker/api.dockerfile`)

## 6. Powierzchnia API (aktualne kontrolery)

### Document API
- `POST /api/document/open`
- `POST /api/document/save`
- `GET /api/document/new`
- `POST /api/document/export-pdf` (aktualnie zwraca 501)
- `POST /api/document/upload-image`
- `GET /api/document/templates`
- `GET /api/document/templates/{templateId}`
- `POST /api/document/sign`
- `POST /api/document/verify-signatures`

### Document Storage API
- `POST /api/documentstorage/upload`
- `POST /api/documentstorage/{masterId}/save`
- `GET /api/documentstorage/{masterId}`
- `GET /api/documentstorage/{masterId}/versions`
- `GET /api/documentstorage/{masterId}/download`
- `GET /api/documentstorage/{masterId}/versions/{versionId}/download`
- `POST /api/documentstorage/{masterId}/restore/{versionId}`

### Barcode API
- `GET /api/barcode/types`
- `POST /api/barcode/generate`
- `POST /api/barcode/generate-image`

### Health API
- `GET /api/health`

