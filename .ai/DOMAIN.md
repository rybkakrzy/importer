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
| Metadata | JSON od aplikacji zewnętrznej (`{ returnUrl, classification }`) w kolumnie `documents.metadata` | Active |

## Encje domenowe

**`Document`** (`D2ViewerEditor.Domain/Entities/Document.cs`) — agregat-root.
- `Id` (guid_master), `Name`, `MimeType`, `CreatedAt`, `CreatedBy`, `IsDeleted` (soft delete), `Metadata` (string? JSON), `Versions`.
- `AddVersion(storagePath, sizeInBytes, createdBy)` — tworzy nową wersję i dezaktywuje wszystkie poprzednie.
- `UpdateVersion(versionId, sizeInBytes)` — nadpisuje istniejącą wersję w miejscu (auto-save); aktualizuje rozmiar + `ModifiedAt`, zachowuje Id/VersionNumber/StoragePath/CreatedAt.
- `RestoreVersion(versionId)` — dezaktywuje wszystkie, aktywuje wskazaną.
- `GetActiveVersion()` — może zwrócić null.
- `Delete()` — soft delete.

**`DocumentVersion`** (`Domain/Entities/DocumentVersion.cs`) — nie agregat-root.
- `Id` (guid_wersji), `DocumentId`, `StoragePath` (`documents/{versionId}`), `SizeInBytes`, `VersionNumber`, `CreatedAt`, `CreatedBy`, `IsActive`, `ModifiedAt` (DateTime?).
- `internal Activate()/Deactivate()/UpdateContent(sizeInBytes)` — mutowane wyłącznie przez agregat `Document`.

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

## Zasady dla agenta

- Nie zmieniaj nazewnictwa domenowego bez aktualizacji tego pliku.
- Nie upraszczaj reguł wersji (v1 immutable, płaski zapis v2) — to rdzeń produktu.
- Niejasną regułę zapisz w `RISKS_ASSUMPTIONS.md`.
