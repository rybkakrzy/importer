# Domain

## Cel pliku

Język domenowy, reguły biznesowe i model pojęciowy. Czytaj przed zmianą logiki biznesowej.

## Język domenowy

| Termin | Znaczenie | Status |
|---|---|---|
| Document (master) | Agregat-root; jeden dokument logiczny. `Id` = `guid_master`, używany w trasach API | Active |
| DocumentVersion | Konkretna wersja dokumentu z referencją do pliku w GCS. `Id` = `guid_wersji` | Active |
| Wersja oryginalna (v1) | Pierwsza wersja = oryginał przysłany przez aplikację zewnętrzną. **Nietykalna** | Active |
| Wersja edytowalna (v2) | Kopia oryginału tworzona dla DOCX; na niej pracuje edytor + auto-save | Active |
| MasterId / VersionId | GUID-y zwracane do aplikacji zewnętrznej po ingeście | Active |
| Classification | Klasyfikacja dokumentu: `C1` | `C2` | `C3` | `C4` (obligatoryjna przy ingeście) | Active |
| ReturnUrl | URL, na który aplikacja zewnętrzna oczekuje zwrotu pliku po „Zakończ" | Active |
| Metadata | JSON od aplikacji zewnętrznej (`{ returnUrl, classification, allowedCorporateKeys? }`) w kolumnie `documents.metadata` | Active |
| CorporateKey | Identyfikator korporacyjny użytkownika; v1 z nagłówka `X-Corporate-Key` (seam pod Entra ID) | Active |
| allowedCorporateKeys | Opcjonalna lista `CorporateKey` uprawnionych do podglądu dokumentu (w metadanych) | Active |
| DocumentStatus | Cykl życia dokumentu: `Saved` → `Editing` → `Sending` → `Sent` / `DeliveryFailed` | Active |
| DocumentDelivery | Zadanie wysyłki finalnego pliku na `ReturnUrl` (kolejka „Zakończ i wyślij") | Active |
| DeliveryStatus | Status zadania wysyłki: `Pending` / `Sending` / `RetryScheduled` / `Sent` / `FailedPermanently` / `DeadLettered` | Active |
| Snapshot (delivery) | Niezmienny obiekt GCS `deliveries/{deliveryId}` zamrożony w chwili „Zakończ" (chroni przed wysłaniem później zmienionej v2) | Active |

## Encje domenowe

**`Document`** (`D2ViewerEditor.Domain/Entities/Document.cs`) — agregat-root.
- `Id` (guid_master), `Name`, `MimeType`, `CreatedAt`, `CreatedBy`, `IsDeleted` (soft delete), `Metadata` (string? JSON), `Status` (`DocumentStatus`, domyślnie `Saved`), `Versions`.
- `AddVersion(storagePath, sizeInBytes, createdBy)` — tworzy nową wersję i dezaktywuje wszystkie poprzednie.
- `UpdateVersion(versionId, sizeInBytes)` — nadpisuje istniejącą wersję w miejscu (auto-save); aktualizuje rozmiar + `ModifiedAt`, zachowuje Id/VersionNumber/StoragePath/CreatedAt.
- `RestoreVersion(versionId)` — dezaktywuje wszystkie, aktywuje wskazaną.
- `GetActiveVersion()` — może zwrócić null.
- `MarkEditing()/MarkSending()/MarkSent()/MarkDeliveryFailed()` — przejścia `Status` (wołane przez handlery: auto-save → `Editing`; „Zakończ" → `Sending`; worker → `Sent`/`DeliveryFailed`).
- `Delete()` — soft delete.

**`DocumentVersion`** (`Domain/Entities/DocumentVersion.cs`) — nie agregat-root.
- `Id` (guid_wersji), `DocumentId`, `StoragePath` (`documents/{versionId}`), `SizeInBytes`, `VersionNumber`, `CreatedAt`, `CreatedBy`, `IsActive`, `ModifiedAt` (DateTime?).
- `internal Activate()/Deactivate()/UpdateContent(sizeInBytes)` — mutowane wyłącznie przez agregat `Document`.

**`DocumentDelivery`** (`Domain/Entities/DocumentDelivery.cs`) — aggregate-root kolejki wysyłki „Zakończ i wyślij".
- `Create(...)` — fabryka; waliduje `recipientUrl` (absolutny http(s)), ustawia `Pending`, `DeadlineAt = teraz + okno (24 h)`.
- Pola: `Id`, `DocumentId`, `SourceVersionId`, `SnapshotObjectName`/`SnapshotSizeBytes`/`SnapshotSha256`, `RecipientUrl`, `Status` (`DeliveryStatus`), `AttemptCount`, znaczniki czasu (`Created/Updated/FirstAttempt/LastAttempt/NextAttempt/Deadline`), lease (`LockedUntil`/`LockedBy`), `LastError`, `CorrelationId`, `CreatedBy`.
- `MarkSent()` / `MarkPermanentFailure(error)` / `ScheduleRetryOrDeadLetter(error, backoff)` (retry jeśli `next ≤ DeadlineAt`, inaczej `DeadLettered`) / `Requeue(window)` (ręczne wznowienie zadania w stanie końcowym, oprócz `Sent`).
- `IsTerminal` = `Sent` ∨ `FailedPermanently` ∨ `DeadLettered`. `IsValidRecipientUrl(url)` — statyczna walidacja URL.
- Logika wysyłki HTTP/GCS jest w infrastrukturze (`IDeliverySender`, `IDocumentStorageService`); domena decyduje tylko o przejściach stanu.

## Reguły biznesowe

| ID | Reguła | Status |
|---|---|---|
| BR-001 | Tylko jedna wersja jest aktywna naraz; `AddVersion` dezaktywuje poprzednie | Active |
| BR-002 | Wersja oryginalna (v1, `VersionNumber == 1`) jest nietykalna — `UpdateVersion` na v1 rzuca wyjątkiem | Active |
| BR-003 | Ingest DOCX → master + v1 (oryginał) + v2 (kopia edytowalna); zwraca `{ MasterId, VersionId }` | Active |
| BR-004 | Ingest PDF → master + tylko v1; zwraca `{ MasterId, null }` (brak edytowalnego duplikatu) | Active |
| BR-005 | Auto-save i ręczny „Zapisz" nadpisują v2 w miejscu (ten sam obiekt GCS) — nie tworzą v3/v4/... | Active |
| BR-006 | Klasyfikacja `C1..C4` jest obligatoryjna przy ingeście; zła wartość → 400 | Active |
| BR-007 | `ReturnUrl` wymagany dla DOCX, opcjonalny dla PDF | Active |
| BR-008 | Backend jest źródłem prawdy dla walidacji; frontend waliduje tylko UX | Active |
| BR-009 | Podpis cyfrowy to Custom XML Part (RSA-SHA256), nie standardowe OOXML; hash liczony z `MainDocumentPart` | Active |
| BR-010 | „Zakończ i wyślij" wymaga poprawnego `ReturnUrl` (absolutny http(s)) w metadanych — inaczej operacja odrzucona | Active |
| BR-011 | Dla jednego dokumentu może istnieć tylko jedno aktywne (nieterminalne) zadanie wysyłki — wymuszone unique partial index; wielokrotne kliknięcie zwraca istniejące zadanie (idempotencja) | Active |
| BR-012 | Wysyłany jest niezmienny snapshot zamrożony w chwili „Zakończ" (`deliveries/{deliveryId}`), nie bieżąca v2 — auto-save po zakończeniu nie zmienia wysyłanego pliku | Active |
| BR-013 | Wysyłka jest at-least-once z retry (exponential backoff + jitter, cap 15 min) do twardego limitu 24 h; po przekroczeniu zadanie → `DeadLettered`; błąd non-retryable → `FailedPermanently` | Active |
| BR-014 | Dostęp do podglądu: gdy metadane mają niepustą `allowedCorporateKeys`, dokument widzi tylko użytkownik z pasującym `CorporateKey` (Trim + ignore-case); brak/null/pusta lista = dokument dla każdego **zalogowanego**. Egzekwowane w backendzie (`DocumentAccessPolicy`/`IDocumentAccessGuard`) → 403; brak tożsamości → 401. `CorporateKey` z claimu access tokena Entra (konfigurowalny). **`APP_Admin` omija `allowedCorporateKeys`** (pełny wgląd). Patrz ADR-0011 | Active |
| BR-015 | Cała aplikacja wymaga uwierzytelnienia (Entra ID). Moduł admina (`APP_Admin`) chroniony backendowo (policy `RequireAppAdmin`); guardy Angulara to tylko UX | Active |

## Zasady dla agenta

- Nie zmieniaj nazewnictwa domenowego bez aktualizacji tego pliku.
- Nie upraszczaj reguł wersji (v1 immutable, płaski zapis v2) — to rdzeń produktu.
- Niejasną regułę zapisz w `RISKS_ASSUMPTIONS.md`.
