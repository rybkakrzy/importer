# Task Handoff

> Bezpieczne przekazanie pracy kolejnej sesji/agentowi.

## Aktualne zadanie

Dokończyć przepływ Krok 1–3 dla D2 ViewerEditor: ingest (gotowy), płaski zapis + auto-save (gotowy), tryb podglądu (Krok 2, w toku), funkcja „Zakończ" (planowana).

## Kontekst

System przyjmuje dokumenty od aplikacji zewnętrznej (External API), edytuje/ogląda w GUI, zapisuje wersjonowanie w DB+GCS przez Internal API. Oryginał (v1) nietykalny; edycja nadpisuje v2 w miejscu.

## Pliki, które trzeba znać

| Plik / obszar | Dlaczego |
|---|---|
| `.ai/PROJECT_CONTEXT.md`, `.ai/DOMAIN.md` | model i reguły (v1 immutable, płaski zapis) |
| `.ai/API_CONTRACTS.md` | katalog endpointów obu API |
| `D2ApiViewerEditor/.../Features/Documents/Commands/UpdateDocumentVersion/` | nadpisanie wersji |
| `D2ApiViewerEditor/.../Features/Documents/Commands/IngestExternalDocument/` | ingest |
| `D2GuiViewerEditor/src/app/components/document-editor/document-editor.ts` | edytor, auto-save, zapis |
| `D2GuiViewerEditor/src/app/services/document-storage.service.ts` | klient storage API |
| `infra/sql/` | schemat (ręczny SQL) |

## Ostatnie wykonane kroki

1. Dodano płaski zapis (`UpdateVersion` + PUT endpoint + `ModifiedAt` + SQL 005).
2. Dodano endpoint metadanych.
3. GUI: auto-save + switch, ujednolicony „Zapisz" / „Pobierz dokument".
4. Przepisano `.ai/` wg wzorca z `template/`.

## Następne sugerowane kroki

1. Krok 2 (podgląd): w GUI dla `?masterId=` ładować v1 przez `GET .../{masterId}/download` zamiast aktywnej wersji; routing PDFViewer vs DocxEditor po `mimeType` z `GET .../{masterId}/metadata` lub `/{masterId}`.
2. Zaimplementować `finishDocument()` — pobranie `returnUrl` z metadanych i odesłanie pliku.
3. (Zamknięte) Port External API = 15112 wg `appsettings.json` `Urls`; Swagger `/swagger`.
4. Dodać testy: domena `UpdateVersion` (v1 immutable), handler `UpdateDocumentVersionCommand`, serwis GUI auto-save.

## Niedokończone zmiany

| Obszar | Co | Ryzyko |
|---|---|---|
| Krok 2 podgląd | ładuje v2 zamiast v1 | użytkownik podglądu widzi edytowalną, nie oryginał |
| `finishDocument()` | TODO | brak zwrotu pliku do aplikacji źródłowej |
| Port External API | rozbieżność repo vs ustalenia | błędne URL integracji |

## Komendy weryfikacji

```bash
dotnet build D2ApiViewerEditor/D2ViewerEditor.sln
cd D2GuiViewerEditor && npm run build
```

Wynik ostatnio: oba OK (0 błędów).

## Nie rób bez zgody

- Nie generuj migracji EF (schemat = `infra/sql/`).
- Nie zmieniaj kontraktów API ani portów bez ustaleń.
- Nie ruszaj reguły v1-immutable / płaskiego zapisu.
- Nie aktualizuj major .NET/Angular. Nie commituj sekretów.
