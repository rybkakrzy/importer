# Importer Parametryzacji

System do importu i analizy plików ZIP zawierających pliki parametryzacji.

## Struktura projektu

- **backend** - ASP.NET 8 Web API
- **frontend** - Angular (najnowsza wersja)

## Funkcjonalność

### Backend (ASP.NET 8 Web API)
- Endpoint: `POST /api/FileUpload/upload`
- Przyjmuje plik ZIP
- Analizuje zawartość archiwum
- Zwraca informacje o plikach znajdujących się w archiwum (docx, json, certyfikaty)

### Frontend (Angular)
- Komponent do wyboru pliku ZIP
- Wysyłanie pliku do API
- Wyświetlanie szczegółowej informacji o zawartości archiwum

## Uruchomienie

### Backend

```powershell
cd backend
dotnet run
```

API będzie dostępne pod adresem: `http://localhost:5190`

### Frontend

```powershell
cd frontend
npm start
```

Aplikacja będzie dostępna pod adresem: `http://localhost:4200`

## Oczekiwana struktura pliku ZIP

W archiwum ZIP powinny znajdować się:
- Pliki DOCX (dokumenty Word)
- Plik JSON z konfiguracją (struktura zostanie zdefiniowana później)
- Certyfikat (.cer, .crt, .pem, .pfx)

## Technologie

### Backend
- ASP.NET 8 Web API
- System.IO.Compression dla obsługi ZIP

### Frontend
- Angular (standalone components)
- HttpClient dla komunikacji z API
- Reactive Forms

## TODO
- Zdefiniować dokładną strukturę pliku JSON
- Dodać walidację zawartości JSON
- Dodać obsługę certyfikatów




Plan na 5 dni (pon.–pt.) – „Minimum Viable Handover”

Zakładam tydzień roboczy 02–06.03.2026.

Dzień 1 — End-to-end overview + uruchamianie

Zakres

Cel/zakres automatyzacji (smoke/regression/e2e) + granice odpowiedzialności

Jak uruchomić: lokalnie i w CI (komendy, parametry, markery/tagi)

Wymagania środowiskowe (env vars, test accounts, dependencies)

Deliverables

1–2 str. “Test Automation Overview”

“Quick Start” (komendy + prerequisites)

Lista krytycznych zależności (secrets, konta, endpointy)

Dzień 2 — Architektura frameworka w Playwright/Python

Zakres

Struktura repo, konwencje, wzorce (fixtures, page objects/helpers)

Standardy pisania testów + selektory i stabilność

Strategie danych testowych (setup/teardown, izolacja)

Deliverables

“How to add a new test” (szablon + checklist)

Lista konwencji i “do/don’t” (krótko, punktowo)

Dzień 3 — CI/CD, raporty, re-run, debug

Zakres

Pipeline: kroki, sharding/parallel, retry policy, timeouts

Raportowanie: gdzie są wyniki, jak je interpretować

Artefakty diagnostyczne: trace/video/screenshoty, jak odtwarzać lokalnie

Deliverables

Runbook: “Test failure triage in CI”

Linki do pipeline’ów + opis najważniejszych jobów i parametrów

Dzień 4 — Stabilność (flaky), utrzymanie, dług techniczny

Zakres

Top flaky testy + root causes

Plan stabilizacji: waits vs assertions, mockowanie vs real backend, dane testowe

Optymalizacja czasu wykonania suite

Deliverables

“Flaky tests backlog” (priorytet, przyczyna, rekomendowany fix, owner)

1 szybki fix/PR (nawet mały) jako przykład wzorca

Dzień 5 — Operating model po odejściu + zamknięcie handover

Zakres

Ownership model: kto reaguje na awarie testów, SLA, eskalacje

Definition of Done: kiedy dodajemy E2E, kiedy nie

Risk review + lista otwartych tematów

Deliverables

RACI/ownership (nawet mini) dla test automation

“Handover checklist completed” + lista open items

Potwierdzenie przekazania dostępów / sposobu rotacji sekretów (bez ujawniania wartości)

Handover Log – struktura notatki po każdym spotkaniu (copy/paste)

Date / Session topic:

Scope covered:

Key takeaways: (3–5 punktów)

Links: (repo paths, pipeline, dashboards, docs)

Risks / constraints:

Action items:

 Task – Owner – Due date

Priorytetyzacja „co musi wyjść” przy małej ilości czasu

Jeśli macie tylko 30 min dziennie i wszystko się sypie, to w tej kolejności:

Jak uruchomić i zdebugować w CI (bez tego jesteś ślepy)

Gdzie są sekrety/konta i jak nie zablokować pipeline’u

Konwencje frameworka + jak dodać test

Flaky i triage process (żeby nie wyłączać testów w panice)

Ownership + eskalacje

Jeśli powiesz mi:

czy macie Confluence/Jira czy raczej Teams/SharePoint,

jaki CI (GitHub Actions / GitLab / Jenkins / Azure DevOps),
to dopasuję Ci dokładnie deliverables (np. pod konkretne sekcje Confluence + naming w Jira) i zaproponuję gotowy szablon stron/runbooków pod Wasze standardy.
