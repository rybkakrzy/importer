# API Contracts

## Cel pliku

Zasady projektowania/zmiany API oraz katalog realnych endpointów obu backendów.

## Styl API

REST over HTTP/JSON (+ multipart dla plików). Dwa backendy:
- **Internal API** `D2ApiViewerEditor` — base `http://localhost:5190/api` (DEV). Swagger UI: `http://localhost:5190/swagger`. Konsument: GUI.
- **External API** `D2ServicesViewerEditor` — base `http://localhost:15112/api/v1` (DEV; `Urls` w `appsettings.json`). Swagger UI: `http://localhost:15112/swagger` (Swagger włączony w dev/local). Konsument: aplikacje źródłowe.

## Zasada kompatybilności

Nie zmieniaj istniejącego kontraktu w sposób breaking bez świadomej decyzji. Breaking = zmiana nazwy/typu/znaczenia pola, usunięcie pola, zmiana statusów HTTP, formatu błędu, wymagań walidacyjnych, paginacji.

## Format błędu

Obecny styl projektu: `{ "error": "komunikat" }` (sprawdź `BaseApiController` przy zmianach). Trzymaj się go dla spójności.

## Wersjonowanie

- Internal API: bez wersji w ścieżce (`/api/...`).
- External API: wersjonowanie w URL (`/api/v1/...`).

---

# Katalog endpointów — Internal API (`/api/documentstorage`, `/api/document`, `/api/barcode`)

## DocumentStorageController — `/api/documentstorage` (CRUD: DB + GCS)

| Metoda | Ścieżka | Opis | Odpowiedź |
|---|---|---|---|
| GET | `/?skip=0&take=200` | Lista dokumentów (admin) | `DocumentListItemDto[]` |
| POST | `/upload` | Upload nowego dokumentu (pierwszy zapis) | `UploadDocumentResult { masterId, versionId }` |
| POST | `/{masterId}/save` | Nowa wersja (`AddVersion`) | `SaveDocumentVersionResult { versionId }` |
| **PUT** | `/{masterId}/versions/{versionId}` | **Nadpisanie wersji w miejscu** (auto-save / „Zapisz"). v1 → 400 | `UpdateDocumentVersionResult { versionId, versionNumber, sizeInBytes, modifiedAt }` |
| GET | `/{masterId}` | Dokument + aktywna wersja (z contentem) | `DocumentDto` |
| **GET** | `/{masterId}/metadata` | Metadane (returnUrl, classification) | `{ masterId, mimeType, returnUrl?, classification? }` |
| GET | `/{masterId}/versions` | Historia wersji (bez contentu) | `DocumentVersionDto[]` |
| GET | `/{masterId}/download` | Bajty wersji bazowej (v1, sort po VersionNumber) | plik |
| GET | `/{masterId}/versions/{versionId}/download` | Bajty konkretnej wersji | plik |
| POST | `/{masterId}/restore/{versionId}` | Przywróć wersję jako aktywną | `{ message, versionId }` |
| **POST** | `/{masterId}/versions/{versionId}/finish` | **„Zakończ i wyślij"**: utrwala stan edytora, zamraża snapshot, tworzy zadanie wysyłki. Idempotentne (zwraca istniejące aktywne zadanie). Brak/zły returnUrl → 400 | **202 Accepted** `{ deliveryId, status, statusUrl }` |
| GET | `/deliveries/{deliveryId}` | Status zadania wysyłki (polling z GUI) | `DeliveryStatusDto { deliveryId, documentId, status, attemptCount, lastAttemptAt?, nextAttemptAt?, lastError?, updatedAt }` |
| GET | `/deliveries?status=&skip=&take=` | Lista zadań w danym statusie (monitoring/admin; domyślnie `DeadLettered`) | `DeliveryListItemDto[]` |
| POST | `/deliveries/{deliveryId}/retry` | Ręczne ponowienie nieudanego zadania (`DeadLettered`/`FailedPermanently`) | `RequeueDeliveryResult { deliveryId, status }` |

Request DTO zapisu: `{ content: byte[]/base64, createdBy?: string }` (ten sam DTO `SaveDocumentVersionRequest` używany też przez `finish`).

## DocumentController — `/api/document` (operacje bezstanowe, bez DB/GCS)

| Metoda | Ścieżka | Opis |
|---|---|---|
| POST | `/open` | DOCX (multipart) → `DocumentContent` (HTML) |
| POST | `/save` | HTML → DOCX (download) |
| GET | `/new` | Pusty `DocumentContent` |
| POST | `/upload-image` | Obraz → Base64 (`ImageUploadResponse`) |
| GET | `/templates` / `/templates/{id}` | Szablony |
| POST | `/sign` | HTML → DOCX → podpisany DOCX |
| POST | `/verify-signatures` | DOCX → `DigitalSignatureInfo[]` |
| POST | `/export-pdf` | **501 Not Implemented** (placeholder) |

## BarcodeController — `/api/barcode`

| Metoda | Ścieżka | Opis |
|---|---|---|
| POST | `/generate` | `{ content, barcodeType, width, height, showText }` → `{ imageBase64, mimeType }` |
| GET | `/types` | `string[]` typów (QR_CODE, CODE_128, EAN_13, ...) |

## HealthController — `/api/health`

Standardowe health checki.

---

# Katalog endpointów — External API (`/api/v1/document`)

## DocumentController (`D2ServicesViewerEditor`)

| Metoda | Ścieżka | Opis |
|---|---|---|
| POST | `/api/v1/document` | Ingest DOCX/PDF (multipart). Pola: `File`, `ReturnUrl` (wymagany dla DOCX), `Classification` (C1..C4, obligatoryjna). Nagłówek opcjonalny `X-Created-By`. → 201 `CreateDocumentResponse { masterId, versionId? }` (versionId tylko DOCX) |
| GET | `/api/v1/document/{documentId}` | Placeholder (read flow niezaimplementowany) |

Metadane trafiają do `documents.metadata` jako `{ "returnUrl": "...", "classification": "C2" }`.

## Checklist przed zmianą API

- Czy frontend / aplikacja zewnętrzna są zaktualizowane?
- Czy testy przechodzą? Czy stary klient się nie psuje?
- Czy statusy HTTP i format błędu są spójne ze stylem?
- Czy zmiana jest w `CHANGELOG.md`?
