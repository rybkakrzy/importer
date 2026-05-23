# Task Handoff

> Ten plik służy do bezpiecznego przekazania pracy kolejnemu agentowi lub przyszłej sesji.

## Aktualne zadanie

Dostosować wstępnie przygotowany katalog `.ai/` do realnego projektu .NET + Angular.

## Kontekst

Katalog `.ai/` ma być pamięcią operacyjną projektu dla agentów AI. Pliki są już wstępnie uzupełnione pod typowy projekt fullstack, ale wymagają dopasowania do rzeczywistego repozytorium.

## Pliki, które trzeba znać

| Plik / katalog | Dlaczego jest ważny |
|---|---|
| `CLAUDE.md` | Punkt wejścia dla Claude |
| `AGENTS.md` | Punkt wejścia dla innych agentów |
| `.ai/INDEX.md` | Mapa całej pamięci projektu |
| `.ai/CURRENT_STATE.md` | Aktualny stan prac |
| `.ai/TECH_STACK.md` | Stack i źródła prawdy |
| `.ai/ARCHITECTURE.md` | Domyślne granice architektury |
| `.ai/BACKEND_DOTNET.md` | Zasady backendu |
| `.ai/FRONTEND_ANGULAR.md` | Zasady frontendu |

## Ostatnie wykonane kroki

1. Utworzono strukturę `.ai/`.
2. Dodano instrukcje startowe dla Claude i innych agentów.
3. Wstępnie uzupełniono kontekst pod projekt .NET + Angular.
4. Dodano szablony tasków, ADR, handoffu i specyfikacji funkcji.

## Następne sugerowane kroki

1. Skopiować katalog `.ai/`, `CLAUDE.md` i `AGENTS.md` do repozytorium.
2. Uzupełnić `PROJECT_CONTEXT.md` realnym opisem projektu.
3. Zweryfikować stack i wpisać fakty do `TECH_STACK.md`.
4. Dopasować komendy build/test/run w `CURRENT_STATE.md`.
5. Spisać realne funkcje w `FEATURES.md`.
6. Spisać realne terminy domenowe w `DOMAIN.md` i `GLOSSARY.md`.
7. Uruchomić build/test, żeby potwierdzić komendy.
8. Zaktualizować `CHANGELOG.md`.

## Niedokończone zmiany

| Obszar | Co jest niedokończone | Ryzyko |
|---|---|---|
| Domenowy opis projektu | Brak realnej domeny | Agent może działać zbyt ogólnie |
| Stack technologiczny | Brak potwierdzonych wersji | Agent może użyć złych API frameworka |
| Komendy developerskie | Brak potwierdzenia | Agent może odpalać złe komendy |
| Deployment | Brak realnego opisu środowisk | Ryzyko błędnych założeń przy DevOps |

## Komendy użyte do weryfikacji

```bash
# Na etapie template'u brak komend projektowych.
# Po wrzuceniu do repo uruchom:
dotnet build
dotnet test
npm run build
npm test
```

## Wynik weryfikacji

Template utworzony. Weryfikacja projektu wymaga uruchomienia komend w realnym repozytorium.

## Nie rób bez zgody

- Nie zmieniaj architektury całego projektu.
- Nie aktualizuj major version .NET lub Angulara.
- Nie zmieniaj provider'a auth.
- Nie migruj bazy danych.
- Nie dodawaj nowej biblioteki UI.
- Nie wykonuj destrukcyjnych zmian w danych.
