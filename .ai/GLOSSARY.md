# Glossary

## Cel

Terminy techniczne i domenowe używane w projekcie.

## Terminy domenowe

| Termin | Znaczenie | Uwagi |
|---|---|---|
| Document / master | Agregat-root dokumentu; `Id` = guid_master | używany we wszystkich trasach API |
| DocumentVersion | Wersja dokumentu z plikiem w GCS; `Id` = guid_wersji | tylko jedna aktywna |
| Wersja oryginalna (v1) | Oryginał przysłany przez aplikację zewnętrzną | nietykalna (BR-002) |
| Wersja edytowalna (v2) | Kopia dla DOCX, edytowana w GUI | nadpisywana w miejscu przez auto-save |
| MasterId / VersionId | GUID-y zwracane po ingeście | aplikacja zewnętrzna składa z nich URL do GUI |
| Classification | C1/C2/C3/C4 — klasyfikacja dokumentu | obligatoryjna przy ingeście |
| ReturnUrl | URL zwrotu pliku po „Zakończ" | wymagany dla DOCX |
| Ingest | Przyjęcie dokumentu od aplikacji zewnętrznej | `D2ServicesViewerEditor` |
| Custom XML Part | Miejsce przechowywania podpisu w DOCX | namespace `http://schemas.D2ViewerEditor.app/digitalsignatures` |
| Zakończ i wyślij | Finalizacja dokumentu + asynchroniczna wysyłka na ReturnUrl | `FinishAndSendDocumentCommand` |
| DocumentStatus | Cykl życia dokumentu (Saved/Editing/Sending/DeliveryFailed/Sent) | kolumna `documents.status` |
| DocumentDelivery | Zadanie wysyłki w kolejce `document_deliveries` | aggregate-root |
| DeliveryStatus | Status zadania wysyłki | Pending/Sending/RetryScheduled/Sent/FailedPermanently/DeadLettered |
| Snapshot (delivery) | Niezmienny plik wysyłki w GCS | obiekt `deliveries/{deliveryId}` |
| DeadLettered | Zadanie po wyczerpaniu okna retry (24 h) | stan końcowy, nieudany |

## Terminy techniczne

| Termin | Znaczenie | Uwagi |
|---|---|---|
| Internal API | `D2ApiViewerEditor` | obsługuje GUI |
| External / Services API | `D2ServicesViewerEditor` | obsługuje aplikacje źródłowe |
| GUI | `D2GuiViewerEditor` (Angular 20) | PDFViewer / DocxEditor |
| GCS | Google Cloud Storage (fake-gcs-server w DEV) | obiekt `documents/{versionId}` |
| CQRS / MediatR | komendy zmieniają stan, query czytają | behaviours: Logging, Validation |
| Result<T> | wynik operacji (`IsSuccess`, `IsNotFound`, `Error`, `Value`) | `Domain/Common/Result.cs` |
| DTO | obiekt transferu danych | nie utożsamiać z encją |
| Signal | mechanizm stanu Angular 20 | główny sposób stanu w GUI |
| Interceptor | przechwytywanie HTTP (Angular) | `http-error.interceptor` |
| BackgroundService | hostowany worker .NET | `DocumentDeliveryWorker` (wysyłka w tle) |
| SKIP LOCKED | bezpieczny claim zadań przez wiele instancji | `SELECT ... FOR UPDATE SKIP LOCKED` |
| Lease | techniczna dzierżawa zadania przez workera | kolumny `locked_until`/`locked_by` (NIE status) |
| Idempotency-Key | nagłówek POST = `deliveryId` | dedup po stronie odbiorcy (at-least-once) |
| Backoff + jitter | strategia ponawiania | `ExponentialJitterBackoff` (cap 15 min) |

## Zasada nazewnictwa

Nowy ważny termin dodaj tutaj lub w `DOMAIN.md`.
