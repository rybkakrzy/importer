# Current State

> Plik krótki i aktualny. Czytaj na początku każdej sesji.

## Ostatnia aktualizacja

2026-05-23 — dostosowanie `.ai/` do realnego projektu D2 ViewerEditor + opis funkcji ingest/auto-save.

## Aktualny cel prac

Domknięcie przepływu Krok 1–3 (ingest → podgląd → edycja) oraz płaski model zapisu wersji edytowalnej z auto-save.

## Aktualny etap

- Branch: `feature/baz-astral-sie-kuta-sb-har`
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
| `finishDocument()` w GUI to TODO (brak zwrotu pliku na returnUrl) | Medium | Open |
| Ręczny „Zapisz" przy włączonym auto-save jest częściowo redundantny | Low | Akceptowalne (wymuszenie zapisu) |
| Port External API potwierdzony: 15112 (`appsettings.json` → `Urls`) | — | Zamknięte |

## Ostatni bezpieczny punkt

Build backendu i GUI przechodzą po zmianach auto-save / „Zapisz/Pobierz".
