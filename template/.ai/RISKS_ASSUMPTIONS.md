# Risks and Assumptions

## Cel pliku

Ten plik zawiera niepewności, ryzyka i założenia. Agent powinien go aktualizować, gdy nie może czegoś potwierdzić w repozytorium lub dokumentacji.

## Aktywne założenia

| ID | Założenie | Wpływ | Jak potwierdzić | Status |
|---|---|---|---|---|
| A-001 | Projekt jest aplikacją fullstack .NET + Angular | Medium | Sprawdzić `.csproj`, `package.json`, strukturę repo | Open |
| A-002 | API jest REST/JSON | Medium | Sprawdzić endpointy i frontend services | Open |
| A-003 | Docker jest lub ma być podstawą local/deploy workflow | Low | Sprawdzić `Dockerfile`, `docker-compose.yml`, CI | Open |
| A-004 | GCP jest możliwym targetem deploymentu | Low | Sprawdzić IaC, README, pipeline | Open |

## Aktywne ryzyka

| ID | Ryzyko | Prawdopodobieństwo | Wpływ | Mitigacja | Status |
|---|---|---|---|---|---|
| R-001 | Wstępnie uzupełnione pliki mogą nie pasować w 100% do realnego repo | High | Medium | Dostosować po pierwszym skanie repo | Open |
| R-002 | Agent może użyć złych komend build/test | Medium | Medium | Zweryfikować `package.json`, `.csproj`, CI | Open |
| R-003 | Agent może potraktować przykładowe funkcje jako istniejące | Medium | Medium | Oznaczono status `Unknown`; trzeba potwierdzić w repo | Open |
| R-004 | Nieaktualizowana pamięć `.ai/` może szkodzić zamiast pomagać | Medium | High | Aktualizować `CURRENT_STATE.md` i `TASK_HANDOFF.md` po pracy | Open |

## Zamknięte założenia i ryzyka

| ID | Opis | Wynik | Data |
|---|---|---|---|
| None | Brak zamkniętych wpisów | N/A | 2026-05-22 |

## Kiedy dodawać wpis

Dodaj wpis, gdy:

- agent zgaduje wersję technologii,
- nie wiadomo, czy kontrakt API jest używany przez klienta,
- brakuje testów dla zmienianego obszaru,
- zmiana może być breaking change,
- decyzja zależy od środowiska produkcyjnego,
- istnieje ryzyko migracji danych,
- wymaganie biznesowe jest niepełne,
- nie wiadomo, czy dany moduł jest jeszcze używany.
