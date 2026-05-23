# Testing and Quality

## Cel

Testy mają chronić zachowanie systemu, nie tylko podnosić coverage.

## Stack testowy (zweryfikowany)

- Backend: **NUnit 4.2.2** + FluentAssertions 8.8.0 + Moq 4.20.72 / NSubstitute 5.3.0. Projekty: `D2ViewerEditor.{Api,Application,Domain,Infrastructure}.UnitTests`. Benchmarki: `D2ViewerEditor.Benchmarks` (BenchmarkDotNet).
- Frontend: **Vitest 3.1.1** (`ng test`).

## Komendy

```bash
# Backend (z katalogu D2ApiViewerEditor)
dotnet build D2ViewerEditor.sln
dotnet test  D2ViewerEditor.sln

# Frontend (z katalogu D2GuiViewerEditor)
npm install
npm run build      # ng build — szybka weryfikacja kompilacji TS/templatek
npm test           # vitest
```

> `npm run lint` NIE istnieje w `package.json` — nie wywołuj. Build (`ng build`) jest podstawową bramką dla frontu.

## Hierarchia weryfikacji

1. Kompilacja / type-check (`dotnet build`, `ng build`).
2. Unit testy logiki (NUnit / Vitest).
3. Testy integracyjne API/bazy — jeśli istnieją.
4. Manualny smoke test, jeśli brak testów.

## Co testować

Backend: reguły domenowe (wersje, v1 immutable), walidatory, handlery komend/zapytań, mapowanie DTO, statusy HTTP, konwersje DOCX↔HTML, podpisy. Istnieją już testy m.in. dla `SaveDocumentVersion`, `UploadDocument`, `GetDocument`, `GetDocumentBaseContent`, `GetDocumentVersionContent`, kontrolera storage.

Frontend: serwisy API (`document.service`, `document-storage.service`), logika `document-editor` (zapis/auto-save), guardy/routing, obsługa loading/error/empty.

## Minimalna macierz weryfikacji

| Typ zmiany | Minimalna weryfikacja |
|---|---|
| C# bez API | `dotnet build` + powiązane testy NUnit |
| Zmiana API | `dotnet build` + test handlera/endpointu + sprawdzenie GUI |
| Angular component | `ng build` + test komponentu jeśli istnieje |
| Serwis API w Angularze | `ng build` + test serwisu |
| Zmiana modelu danych | skrypt SQL `infra/sql/` + mapowanie EF + weryfikacja ręczna |
| Auto-save / wersje | testy domeny (`UpdateVersion`) + build obu projektów |

## Jakość kodu (przed zakończeniem)

- brak dead code / nieużywanych importów,
- brak przypadkowych `console.log` (poza istniejącą celową diagnostyką),
- brak twardych sekretów,
- brak formatowania niezwiązanego z zadaniem,
- brak wyciszania błędów typowania bez powodu.

## Raport końcowy agenta

```md
## Changed
- ...
## Verified
- `command` — result
## Risks
- ...
## Next
- ...
```
