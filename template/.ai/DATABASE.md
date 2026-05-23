# Database

## Cel pliku

Ten plik opisuje zasady pracy z bazą danych, migracjami i zapytaniami.

## Status

Do uzupełnienia po sprawdzeniu repozytorium.

Agent powinien sprawdzić:

- connection stringi,
- modele encji,
- migracje,
- seedy,
- indeksy,
- skrypty inicjalizacyjne,
- docker-compose,
- dokumentację środowisk.

Nie wolno zakładać typu bazy bez sprawdzenia repozytorium.

## Baza danych

| Obszar | Wartość |
|---|---|
| Engine | `TODO: PostgreSQL / SQL Server / MySQL / SQLite / inne` |
| ORM | `TODO: EF Core / Dapper / inne` |
| Migracje | `TODO` |
| Seeding | `TODO` |
| Lokalny setup | `TODO` |
| Test database | `TODO` |

## Domyślny model pracy

Jeżeli projekt używa EF Core:

- encje nie są DTO API,
- migracje są jawne,
- odczyty list powinny mieć projekcję do DTO,
- operacje modyfikujące powinny być transakcyjne, gdy zmieniają kilka agregatów/tabel,
- query performance trzeba sprawdzać dla list i dashboardów.

## Zasady migracji

- Nie generuj migracji bez rzeczywistej zmiany modelu danych.
- Nie usuwaj kolumn ani tabel bez analizy danych.
- Zmiany destrukcyjne wymagają osobnej decyzji.
- Migracje powinny być małe i zrozumiałe.
- Jeżeli migracja przenosi dane, opisz ryzyko i sposób rollbacku.
- Nie mieszaj migracji technicznej z refaktoryzacją domeny.

## Zasady zapytań

- Unikaj N+1.
- Ograniczaj zakres danych pobieranych z bazy.
- Dla list stosuj paginację, jeżeli dane mogą rosnąć.
- Dodawaj indeksy świadomie, pod konkretne zapytania.
- Nie optymalizuj przedwcześnie bez danych lub problemu.
- Nie pobieraj encji z wieloma relacjami tylko po to, żeby zmapować kilka pól.

## Dane testowe

`TODO: opisz skąd pochodzą dane testowe i jak je resetować`

Możliwe warianty:

- seed przy starcie aplikacji,
- skrypty SQL,
- migracje testowe,
- Docker volume reset,
- Testcontainers,
- osobna baza developerska.

## Checklist przed zmianą modelu danych

- Czy zmiana jest wymagana przez funkcję?
- Czy migracja jest bezpieczna dla istniejących danych?
- Czy istnieją testy lub ręczna procedura weryfikacji?
- Czy backend i frontend rozumieją nowy model?
- Czy dokumentacja domeny została zaktualizowana?
- Czy trzeba dodać indeks?
- Czy trzeba uwzględnić rollback?
