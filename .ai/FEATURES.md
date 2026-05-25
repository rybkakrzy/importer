# Features

## Cel pliku

Funkcje systemu z perspektywy produktu i implementacji. Aktualizuj przy zmianie funkcji.

## Status funkcji

| Funkcja | Status | Backend | Frontend | API |
|---|---|---|---|---|
| Ingest dokumentu od aplikacji zewnętrznej (Krok 1) | Implemented | `IngestExternalDocumentCommand` | n/d | `POST /api/v1/document` (Services) |
| Tryb podglądu DOCX/PDF (Krok 2) | In Progress | endpointy gotowe | routing PDF/DOCX częściowy | `GET .../{id}/download`, `GET .../{id}` |
| Tryb edycji DOCX (Krok 3) | Implemented | gotowe | `document-editor` | `GET .../versions/{vid}/download` |
| Auto-save (nadpisanie v2 w miejscu) | Implemented | `UpdateDocumentVersionCommand` | timer + switch AutoSave | `PUT .../versions/{vid}` |
| Ujednolicony „Zapisz" + „Pobierz dokument" | Implemented | — | `saveDocument()`/`downloadDocument()` | `PUT`/`POST save` |
| Metadane (returnUrl, classification) | Implemented | `GetDocumentMetadataQuery` | — | `GET .../{id}/metadata` |
| Wersjonowanie + przywracanie | Implemented | `RestoreDocumentVersionCommand` | admin/historia | `POST .../restore/{vid}` |
| Podpisy cyfrowe (Custom XML Part) | Implemented | `DigitalSignatureService` | dialog podpisu | `POST /api/document/sign`, `/verify-signatures` |
| Kody kreskowe / QR | Implemented | `BarcodeGeneratorService` | `barcode-dialog` | `POST /api/barcode/generate` |
| Szablony dokumentów | Implemented | queries templates | menu szablonów | `GET /api/document/templates` |
| PDF viewer | Implemented | n/d (statyczny plik) | `pdf-viewer` (lazy, pdfjs) | `GET .../download` |
| Eksport PDF | Deprecated/Placeholder | `501 Not Implemented` | — | `POST /api/document/export-pdf` |
| Zakończ i wyślij (async zwrot na returnUrl) | Implemented | `FinishAndSendDocumentCommand` + worker `DocumentDeliveryWorker` | `finishDocument()` + polling statusu | `POST .../versions/{vid}/finish` (202), `GET/POST .../deliveries/...` |

## Statusy

`Unknown` · `Planned` · `In Progress` · `Blocked` · `Implemented` · `Verified` · `Deprecated`

## Feature: Ingest zewnętrzny (Krok 1)

### Cel
Przyjąć dokument od aplikacji źródłowej, zapisać oryginał i (dla DOCX) utworzyć kopię edytowalną.

### Backend
- Endpoint: `POST /api/v1/document` (`D2ServicesViewerEditor`, multipart).
- Use case: `IngestExternalDocumentCommand` / `...Handler` (warstwa Application `D2ApiViewerEditor`).
- Walidator: `IngestExternalDocumentCommandValidator` (typ MIME, rozmiar ≤100 MB, JSON metadanych, classification).
- Reguły: BR-003, BR-004, BR-006, BR-007 (patrz `DOMAIN.md`).

### Otwarte pytania
- Czy wersja edytowalna ma być realną konwersją/normalizacją DOCX, czy dosłowną kopią bajtów? Obecnie kopia bajtów.

## Feature: Auto-save / ujednolicony zapis (Krok 3)

### Cel
Zapisywać zmiany w trybie edycji bez mnożenia wersji.

### Backend
- `PUT /api/documentstorage/{masterId}/versions/{versionId}` → `UpdateDocumentVersionCommand` → `Document.UpdateVersion` (nadpisanie GCS pod tym samym `versionId`). Reguły BR-002, BR-005.

### Frontend
- `components/document-editor/document-editor.ts`: timer `rxjs timer(intervalMs)`, sygnały `autoSaveEnabled/autoSaveStatus/lastAutoSaveAt`, switch „AutoSave" w nagłówku (widoczny w trybie edycji).
- `saveDocument()` = zapis przez API (PUT gdy versionId, inaczej POST). `downloadDocument()` = pobranie pliku lokalnie (dawne „Zapisz").
- Konfiguracja: `environment.autoSave { enabled, intervalSeconds: 30 }`.

## Feature: Zakończ i wyślij (Krok 4)

### Cel
Zakończyć pracę nad dokumentem, utrwalić aktualny stan edytora i asynchronicznie wysłać finalny plik na `ReturnUrl` z metadanych — bez blokowania requestu HTTP, odpornie na restart i wiele instancji.

### Backend
- `POST /api/documentstorage/{masterId}/versions/{versionId}/finish` → `FinishAndSendDocumentCommand` / `...Handler`:
  1. nadpisuje wersję edytowalną treścią z edytora (jak auto-save),
  2. zamraża niezmienny snapshot w GCS (`deliveries/{deliveryId}`, SHA-256),
  3. atomowo (jeden `SaveChanges`) ustawia `Document.Status=Sending` i tworzy zadanie `DocumentDelivery`.
- Idempotencja: jeśli istnieje aktywne zadanie dla dokumentu, zwraca je (nie tworzy nowego); unique partial index chroni przed wyścigiem. Reguły BR-010..BR-013.
- Walidator: `FinishAndSendDocumentCommandValidator` (MasterId/VersionId niepuste, Content niepusty ≤ 100 MB, CreatedBy ≤ 255).
- Worker `DocumentDeliveryWorker` (BackgroundService): claim `FOR UPDATE SKIP LOCKED` + lease, równoległość z limitem, `HttpDeliverySender` (POST z `Idempotency-Key`), retry `ExponentialJitterBackoff` (cap 15 min) do 24 h → `DeadLettered`; błąd non-retryable → `FailedPermanently`. Po sukcesie `Document.Status=Sent`, po porażce `DeliveryFailed`.
- Status/monitoring: `GET .../deliveries/{deliveryId}`, `GET .../deliveries?status=`, ręczne ponowienie `POST .../deliveries/{deliveryId}/retry`.

### Frontend
- `components/document-editor/document-editor.ts`: `finishDocument()` serializuje DOCX, woła `finishAndSend`, ustawia sygnały `isFinishing/deliveryStatus/deliveryId` i odpytuje `getDeliveryStatus` co 4 s (`timer` + `takeWhile`) do stanu końcowego. Wysyłka jest kontynuowana po stronie serwera nawet po zamknięciu strony.
- Serwis `document-storage.service.ts`: `finishAndSend(masterId, versionId, { content })`, `getDeliveryStatus(deliveryId)`.

### Uwaga dla agenta
Brak (na dzień aktualizacji) testów integracyjnych claimu na realnym PostgreSQL (SKIP LOCKED / reclaim) — logikę pokrywają testy jednostkowe domeny, handlera, backoffu i walidatora. Patrz `RISKS_ASSUMPTIONS.md`.

## Feature: Tryb podglądu (Krok 2) — In Progress

### Uwaga dla agenta
Endpointy istnieją (`/download` zwraca v1), ale GUI w trybie `?masterId=` ładuje aktualną wersję (po ingeście DOCX = v2), a nie v1. Aby tryb podglądu pokazywał oryginał, trzeba przełączyć ładowanie na `GET .../{masterId}/download`. Patrz `RISKS_ASSUMPTIONS.md` (R-02).
