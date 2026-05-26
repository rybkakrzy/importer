# Current State

> Plik krótki i aktualny. Czytaj na początku każdej sesji.

## Ostatnia aktualizacja

2026-05-25 — panel admina `/admin/deliveries` (lista wysyłek + filtry + retry + `locked_by`, statusy PL) + uruchomienie migracji `infra/sql/001..007` na lokalnym PostgreSQL (Podman). Wcześniej: funkcja „Zakończ i wyślij" (kolejka `document_deliveries` + worker) + audyt/synchronizacja `.ai/` z kodem.

## Aktualny cel prac

Domknięcie przepływu Krok 1–4 (ingest → podgląd → edycja → zakończ i wyślij). Krok 4 zaimplementowany; Krok 2 (podgląd v1) wciąż w toku.

## Aktualny etap

- Branch: `feature/baz-astral-obciaga-mb`
- Główne obszary kodu:
  - `D2ApiViewerEditor/` (internal API + Application/Domain/Infrastructure)
  - `D2ServicesViewerEditor/` (external ingest API)
  - `D2GuiViewerEditor/` (Angular)
  - `infra/sql/` (schemat)

## Co zostało zrobione (ta i poprzednie sesje)

- Ingest zewnętrzny: `IngestExternalDocumentCommand` + endpoint `POST /api/v1/document` (DOCX → v1+v2, PDF → v1).
- Płaski zapis: `Document.UpdateVersion` + `UpdateDocumentVersionCommand` + `PUT /api/documentstorage/{masterId}/versions/{versionId}` (nadpisanie GCS pod tym samym versionId).
- `DocumentVersion.ModifiedAt` + skrypt `infra/sql/005_add_version_modified_at.sql` + mapowanie EF.
- Endpoint metadanych: `GetDocumentMetadataQuery` + `GET .../{masterId}/metadata`.
- GUI: mechanizm auto-save (timer, `environment.autoSave { enabled, intervalSeconds: 30 }`), switch „AutoSave", sygnały `documentVersionId/autoSaveEnabled/autoSaveStatus/lastAutoSaveAt`.
- GUI: ujednolicony „Zapisz" (API, PUT/POST) + „Pobierz dokument" (dawne „Zapisz" = download lokalny). Usunięto `saveDocumentAs()`.
- Dokumentacja `.ai/` przepisana wg wzorca z `template/`.
- „Zakończ i wyślij": `FinishAndSendDocumentCommand` + tabela `document_deliveries` (SQL 007) + worker `DocumentDeliveryWorker` (claim `FOR UPDATE SKIP LOCKED`, retry/backoff do 24h, snapshot GCS) + endpointy `finish`/`deliveries`/`retry` + GUI `finishDocument()` z pollingiem. Patrz `DECISIONS.md` ADR-0005.
- `documents.status` (`DocumentStatus`) + skrypt SQL 006 (dodane przed tą sesją, teraz udokumentowane).
- Panel admina `/admin/deliveries` (`admin-deliveries`): lista zadań wysyłki (filtr per kolumna + dropdown statusu + paginacja, spójny z `/admin/files`), `locked_by`/`locked_until`, przycisk „Ponów" (`POST deliveries/{id}/retry`), statusy tłumaczone na PL. `DeliveryListItemDto` rozszerzony o `LockedUntil`/`LockedBy`.
- Migracje `infra/sql/001..007` wykonane na lokalnym PostgreSQL (Podman `d2viewereditor_postgres`); `007` utworzył `document_deliveries`.

## Jak uruchomić projekt lokalnie

Patrz `DEVOPS_DEPLOYMENT.md`. Skrót:

```bash
cd D2ApiViewerEditor && dotnet run --project D2ViewerEditor.Api      # 5190
cd D2ServicesViewerEditor && dotnet run --project D2ServicesViewerEditor.Api  # 15112 (swagger: /swagger)
cd D2GuiViewerEditor && npm install && npm start                     # 4200
# Wymaga: PostgreSQL (skrypty infra/sql/ 001..005) + GCS/fake-gcs-server (bucket d2viewereditor-documents)
```

## Jak zweryfikować stan

```bash
dotnet build D2ApiViewerEditor/D2ViewerEditor.sln   # ostatnio: OK (0 błędów)
cd D2GuiViewerEditor && npm run build               # ostatnio: OK
```

## Znane problemy / niedokończone

| Problem | Wpływ | Status |
|---|---|---|
| Tryb podglądu (Krok 2) ładuje aktywną wersję (po DOCX = v2), nie v1 oryginał | Medium | Open — przełączyć GUI na `/download` |
| `finishDocument()` — zaimplementowane: async wysyłka na returnUrl (kolejka `document_deliveries` + worker) | — | Zamknięte |
| `SaveDocumentVersionCommandHandlerTests` oczekuje `UpdateAsync` (płaski zapis go nie woła) | Low | Zamknięte — test asercją `DidNotReceive().UpdateAsync` + `Received(1).SaveChangesAsync` |
| Budżet SCSS `document-editor.scss` przekraczał limit | Low | Zamknięte — `angular.json` `anyComponentStyle` podniesiony (error 48 kB / warn 24 kB) |
| Ręczny „Zapisz" przy włączonym auto-save jest częściowo redundantny | Low | Akceptowalne (wymuszenie zapisu) |
| Port External API potwierdzony: 15112 (`appsettings.json` → `Urls`) | — | Zamknięte |

## Ostatni bezpieczny punkt

Build backendu i GUI przechodzą po zmianach auto-save / „Zapisz/Pobierz".
