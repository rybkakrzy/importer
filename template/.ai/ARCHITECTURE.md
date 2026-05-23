# Architecture

## Cel architektury

Architektura ma pozwolić rozwijać system przez ludzi i agentów AI bez chaosu. Kod powinien mieć jasne granice odpowiedzialności, przewidywalny kierunek zależności i testowalną logikę.

## Domyślny obraz systemu

```text
Angular app
  |
  | HTTP / REST / JSON
  v
ASP.NET Core API
  |
  v
Application layer
  |
  v
Domain model
  |
  v
Infrastructure: database, external APIs, files, queues
```

Jeżeli repozytorium ma inną architekturę, agent ma dostosować się do repo, a nie do tego diagramu.

## Backend — preferowany podział

```text
src/
  Backend/
    Api/
    Application/
    Domain/
    Infrastructure/
    Contracts/
tests/
  Backend.UnitTests/
  Backend.IntegrationTests/
```

### Api

Odpowiada za:

- routing,
- endpointy HTTP,
- request/response DTO,
- autoryzację endpointów,
- mapowanie wejścia/wyjścia,
- middleware,
- filtry,
- konfigurację Swagger/OpenAPI, jeżeli jest używana.

Nie powinna zawierać ciężkiej logiki biznesowej.

### Application

Odpowiada za:

- przypadki użycia,
- komendy i zapytania,
- orkiestrację operacji,
- transakcje,
- walidację aplikacyjną,
- komunikację z repozytoriami przez abstrakcje,
- integrację logiki domenowej z infrastrukturą.

### Domain

Odpowiada za:

- model domenowy,
- reguły biznesowe,
- inwarianty,
- value objects,
- zdarzenia domenowe, jeżeli projekt ich używa.

Warstwa domenowa nie powinna zależeć od infrastruktury.

### Infrastructure

Odpowiada za:

- bazę danych,
- implementacje repozytoriów,
- integracje zewnętrzne,
- kolejki,
- storage,
- providerów czasu, plików, maili, itp.

## Frontend — preferowany podział

```text
src/app/
  core/
  shared/
  features/
    feature-name/
      data-access/
      ui/
      pages/
      models/
```

### Core

Kod globalny i singletony:

- interceptory,
- guards,
- globalne serwisy,
- konfiguracja auth,
- konfiguracja API,
- error handling,
- globalna konfiguracja aplikacji.

### Shared

Elementy wielokrotnego użytku bez zależności domenowych:

- komponenty UI,
- pipe'y,
- dyrektywy,
- helpery,
- typy techniczne.

### Features

Kod konkretnej funkcji:

- strony,
- komponenty,
- serwisy,
- modele,
- integracja z API,
- lokalny stan funkcji.

## Kierunek zależności

Backend:

```text
Api -> Application -> Domain
Api -> Infrastructure
Infrastructure -> Application abstractions / Domain
```

Frontend:

```text
features -> shared
features -> core
shared -> no feature dependencies
core -> no feature dependencies unless project already does it intentionally
```

## Granice odpowiedzialności

| Element | Powinien robić | Nie powinien robić |
|---|---|---|
| Angular component | prezentacja, interakcje, lokalny stan UI | bezpośrednie zapytania HTTP, logika domenowa |
| Angular service | komunikacja z API, state orchestration | renderowanie UI |
| API endpoint | HTTP, auth, mapping, statusy | ciężka logika biznesowa |
| Application service/use case | orkiestracja biznesowa | szczegóły UI, szczegóły HTTP |
| Domain object | reguły biznesowe | dostęp do bazy, HTTP, pliki |
| Infrastructure | szczegóły techniczne | decyzje biznesowe |

## Zakazane skróty

- Nie dodawaj logiki biznesowej bezpośrednio w kontrolerach.
- Nie dodawaj zapytań HTTP bezpośrednio w komponentach, jeżeli projekt używa serwisów.
- Nie mieszaj DTO API z modelem domenowym.
- Nie twórz globalnego serwisu dla funkcji lokalnej.
- Nie dodawaj cyklicznych zależności.
- Nie twórz dużej abstrakcji tylko dla jednej implementacji.
- Nie zmieniaj struktury całego projektu przy okazji jednej funkcji.
