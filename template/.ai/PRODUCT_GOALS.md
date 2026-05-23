# Product Goals

## Cel produktu

`TODO: wpisz konkretny cel biznesowy produktu`

Domyślny kierunek:

> Produkt ma usprawniać konkretny proces biznesowy przez aplikację webową, która pozwala użytkownikom wykonywać operacje szybciej, bezpieczniej i z mniejszą liczbą błędów niż proces ręczny.

## Problem, który rozwiązujemy

Do doprecyzowania:

- rozproszone dane,
- ręczna obsługa procesu,
- brak automatyzacji,
- brak spójnego UI,
- brak audytu operacji,
- trudna integracja z innymi systemami,
- powtarzalne błędy użytkowników.

## Najważniejsze wartości dla użytkownika

- Szybkie wykonanie najważniejszych operacji.
- Przejrzysty stan systemu.
- Stabilność działania.
- Bezpieczne przetwarzanie danych.
- Dobre komunikaty błędów.
- Przewidywalny interfejs.
- Brak utraty danych przy błędach walidacji.
- Możliwość dalszego rozwoju systemu.

## Cele krótkoterminowe

| Cel | Status | Uwagi |
|---|---|---|
| Uporządkować pamięć projektu dla agentów AI | In Progress | Katalog `.ai/` i pliki startowe |
| Spisać faktyczny stack technologiczny | Planned | Uzupełnić po dodaniu do repo |
| Spisać aktualny stan funkcji | Planned | Uzupełnić `FEATURES.md` |
| Spisać komendy build/test/run | Planned | Uzupełnić `CURRENT_STATE.md` i `DEVOPS_DEPLOYMENT.md` |

## Cele długoterminowe

| Cel | Status | Uwagi |
|---|---|---|
| Utrzymać spójność architektury backendu i frontendu | Active | Każdy agent powinien stosować istniejące wzorce |
| Minimalizować regresje przez testy | Active | Najpierw testy dla logiki krytycznej |
| Ułatwić pracę AI i ludziom | Active | Aktualizować handoff po istotnych zmianach |
| Utrzymać bezpieczny deployment | Planned | GCP / Docker / CI do uzupełnienia |

## Poza zakresem bez jawnej decyzji

Agent nie powinien implementować tych rzeczy sam z siebie:

- pełnego redesignu UI,
- zmiany frameworka frontendowego,
- zmiany architektury backendu,
- migracji do innej bazy danych,
- wdrożenia nowego provider'a auth,
- integracji płatności,
- automatycznego usuwania danych produkcyjnych,
- breaking changes w API,
- aktualizacji major version Angulara lub .NET,
- przebudowy pipeline CI/CD.

## Priorytety przy konflikcie

1. Poprawność domenowa.
2. Bezpieczeństwo.
3. Stabilność i kompatybilność.
4. Testowalność.
5. Czytelność.
6. Wydajność.
7. Estetyka implementacji.

Wydajność jest ważna, ale nie powinna prowadzić do skomplikowanego kodu bez pomiarów albo konkretnego problemu.
