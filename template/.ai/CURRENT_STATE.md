# Current State

> Ten plik ma być krótki i aktualny. Agent ma go czytać na początku każdej sesji.

## Ostatnia aktualizacja

2026-05-22 — initial prefilled AI project memory

## Aktualny cel prac

Dodać do repozytorium katalog `.ai/` jako trwałą pamięć projektu dla agentów AI oraz dostosować go do realnego projektu.

## Aktualny etap

- Status: `Prepared template`
- Branch: `TODO: wpisz aktualny branch`
- Powiązane zadanie: `TODO: issue/link/opis`
- Główne obszary kodu:
  - `.ai/`
  - `CLAUDE.md`
  - `AGENTS.md`

## Co zostało zrobione

- Przygotowano strukturę `.ai/`.
- Dodano instrukcje startowe dla Claude w `CLAUDE.md`.
- Dodano neutralne instrukcje dla innych agentów w `AGENTS.md`.
- Wstępnie uzupełniono pliki pod projekt .NET + Angular.
- Dodano zasady pracy, architektury, jakości, bezpieczeństwa i handoffu.

## Co trzeba dostosować po wrzuceniu do repo

- Nazwa i opis projektu w `PROJECT_CONTEXT.md`.
- Faktyczny stack i wersje w `TECH_STACK.md`.
- Realna architektura i ścieżki katalogów w `ARCHITECTURE.md`.
- Realne funkcje w `FEATURES.md`.
- Terminy domenowe w `DOMAIN.md` i `GLOSSARY.md`.
- Komendy uruchomieniowe i testowe w tym pliku.
- Pipeline/deployment w `DEVOPS_DEPLOYMENT.md`.
- Baza danych w `DATABASE.md`.

## Jak uruchomić projekt lokalnie

Do potwierdzenia w repozytorium.

Typowy wariant:

```bash
# Backend
dotnet restore
dotnet build
dotnet run

# Frontend
npm install
npm start

# Docker
docker compose up -d
```

## Jak zweryfikować aktualny stan

Do potwierdzenia w repozytorium.

Typowy wariant:

```bash
dotnet test
npm run build
npm test
npm run lint
```

## Znane problemy

| Problem | Wpływ | Status | Uwagi |
|---|---|---|---|
| Pliki są wstępnie uzupełnione, ale nie znają realnej domeny | Medium | Open | Trzeba dostosować po dodaniu do repo |
| Komendy build/test mogą nie pasować do projektu | Medium | Open | Sprawdzić `package.json`, `.csproj`, CI |
| Struktura architektury może różnić się od rzeczywistej | Medium | Open | Agent ma najpierw sprawdzić repo |

## Ostatni bezpieczny punkt

Initial `.ai/` memory template before customization.
