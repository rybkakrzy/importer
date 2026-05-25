# Project Context

## Nazwa projektu

D2 ViewerEditor (repo: `ImportParametryzacji/importer`).

## Typ projektu

Fullstack web application — webowy edytor/viewer dokumentów DOCX i PDF z wersjonowanym storage, przeznaczony do integracji z aplikacjami zewnętrznymi.

Kontekst technologiczny (zweryfikowany w repo):

- backend w .NET 8 / ASP.NET Core (Clean Architecture + CQRS/MediatR),
- frontend w Angular 20 (standalone components + Signals),
- komunikacja przez REST API (JSON + multipart),
- pliki binarne w Google Cloud Storage (fake-gcs-server w DEV),
- metadane w PostgreSQL (EF Core 8 + Npgsql),
- konteneryzacja przez Docker (osobne obrazy api/gui),
- praca wspierana przez agentów AI (ten katalog `.ai/`).

## Krótki opis

System przyjmuje dokument (DOCX lub PDF) od aplikacji zewnętrznej wraz z metadanymi (URL zwrotny + klasyfikacja C1..C4), zapisuje go i — dla DOCX — tworzy wersję edytowalną będącą kopią oryginału. Użytkownik otwiera dokument w GUI: PDF w trybie podglądu, DOCX w edytorze WYSIWYG. Zmiany w trybie edycji są nadpisywane w miejscu (auto-save) na wersji edytowalnej; oryginał pozostaje nietknięty. System wspiera też podpisy cyfrowe (własny format X.509 w Custom XML Part), generowanie kodów kreskowych/QR i bibliotekę szablonów.

## Trzy projekty w repo

| Projekt | Rola | Port DEV (launchSettings) | Docker |
|---|---|---|---|
| `D2ApiViewerEditor` | **Internal API** — obsługuje GUI (CRUD + konwersje + podpisy) | 5190 / 7190 | `docker/api.dockerfile`, EXPOSE 8080 |
| `D2ServicesViewerEditor` | **External integration API** — przyjmuje dokumenty od aplikacji źródłowych | 15112 (`appsettings.json` → `Urls`); Swagger `/swagger` | `D2ServicesViewerEditor/Dockerfile`, EXPOSE 80/443 |
| `D2GuiViewerEditor` | **Angular SPA** — edytor/viewer | 4200 (`ng serve`) | `docker/gui.dockerfile`, nginx, EXPOSE 80 |

**Kto z kim rozmawia:** aplikacja zewnętrzna → `D2ServicesViewerEditor` (ingest + otwarcie URL do GUI). GUI → `D2ApiViewerEditor`. Aplikacja zewnętrzna nie rozmawia z internal API — sama składa URL do GUI z `MasterId`/`VersionId`.

## Główni użytkownicy

- użytkownik końcowy — edytuje/ogląda dokument w GUI,
- aplikacja źródłowa (integrator) — wysyła dokumenty przez Services API i otwiera GUI,
- administrator — zarządza dokumentami (sekcja `/admin` w GUI),
- developer / maintainer.

## Najważniejsze scenariusze

1. **Ingest** (Krok 1): aplikacja zewnętrzna wysyła plik + metadane → Services API zapisuje (DOCX: oryginał v1 + edytowalna v2; PDF: tylko v1) → zwraca `{ MasterId, VersionId? }`.
2. **Tryb podglądu** (Krok 2): GUI z `?masterId=` ładuje treść (PDFViewer dla PDF, DocxEditor dla DOCX) w trybie read-only.
3. **Tryb edycji** (Krok 3): GUI z `?masterId=&versionId=` ładuje wersję edytowalną; użytkownik edytuje, auto-save nadpisuje v2 w miejscu.
4. **Zakończ i wyślij** (Krok 4): użytkownik kończy pracę → backend zamraża snapshot finalnego pliku i tworzy zadanie wysyłki (`document_deliveries`); worker w tle wysyła plik na `ReturnUrl` z metadanych (asynchronicznie, z retry do 24 h). GUI odpytuje status wysyłki.

## Cele techniczne

- Utrzymywalna Clean Architecture, czytelne granice (Domain → Application → Infrastructure → Api).
- Stabilne kontrakty REST między GUI a internal API oraz między aplikacją zewnętrzną a Services API.
- Płaski model wersji: edycja nadpisuje wersję edytowalną, nie mnoży wersji w GCS.
- Oryginał (v1) nietykalny.
- Testy dla logiki krytycznej (NUnit backend, Vitest frontend).
- Brak sekretów w repo (`appsettings.*.secrets.json` poza repo).

## Główne ograniczenia

- Nie zmieniać publicznych kontraktów API bez świadomej decyzji.
- Nie przepisywać architektury przy małych taskach.
- Schemat bazy jest zarządzany ręcznym SQL (`infra/sql/`), nie EF Migrations — nie generować migracji EF.
- Nie mieszać logiki domenowej z prezentacją.
- Nie usuwać testów, żeby build przeszedł.

## Definicja sukcesu

- Nowe funkcje dodaje się bez dużego długu technicznego.
- Backend i frontend mają jasny kontrakt.
- Agent AI kontynuuje pracę bez odkrywania repo od zera.
- Zmiany są małe, odwracalne i łatwe do review.
