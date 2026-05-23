# API Contracts

## Cel pliku

Ten plik opisuje zasady projektowania i zmiany API między backendem i frontendem.

## Domyślny styl API

Zakładamy REST API over HTTP/JSON, dopóki repozytorium nie pokaże inaczej.

Typowe ścieżki:

```text
GET    /api/resources
GET    /api/resources/{id}
POST   /api/resources
PUT    /api/resources/{id}
PATCH  /api/resources/{id}
DELETE /api/resources/{id}
```

Dostosuj do istniejącego stylu projektu.

## Zasada kompatybilności

Nie zmieniaj istniejącego kontraktu API w sposób breaking change bez świadomej decyzji.

Breaking changes obejmują:

- zmianę nazwy pola,
- zmianę typu pola,
- usunięcie pola,
- zmianę znaczenia pola,
- zmianę statusów HTTP,
- zmianę formatu błędu,
- zmianę wymagań walidacyjnych,
- zmianę paginacji/sortowania/filtrowania.

## Request / Response DTO

- DTO powinny być jawne i nie powinny ujawniać encji bazodanowych.
- Nazwy pól powinny być spójne z istniejącym API.
- Nie zwracaj wewnętrznych identyfikatorów lub danych technicznych, jeżeli nie są potrzebne klientowi.
- Nullable powinno oznaczać realną możliwość braku wartości.
- Unikaj pól o niejasnym znaczeniu, np. `status2`, `flag`, `data`.
- Response DTO może być inne niż request DTO.

## Błędy API

Jeżeli projekt nie ma własnego formatu, preferowany model:

```json
{
  "code": "VALIDATION_ERROR",
  "message": "The request contains invalid data.",
  "details": [
    {
      "field": "email",
      "message": "Email is required."
    }
  ]
}
```

## Paginacja

Jeżeli lista może rosnąć, powinna mieć paginację.

Przykładowy model:

```http
GET /api/resources?page=1&pageSize=20
```

Przykładowa odpowiedź:

```json
{
  "items": [],
  "page": 1,
  "pageSize": 20,
  "totalCount": 0
}
```

Dostosuj do stylu projektu.

## Sortowanie i filtrowanie

- Parametry powinny być jawne.
- Nie przyjmuj dowolnych nazw kolumn bez whitelisty.
- Waliduj zakresy dat, rozmiary stron i pola sortowania.

## Wersjonowanie

Status: `TODO: sprawdzić w repo`

Opcje:

- brak wersjonowania,
- wersjonowanie w URL, np. `/api/v1/...`,
- wersjonowanie nagłówkiem,
- osobne kontrakty dla klientów.

## Checklist przed zmianą API

- Czy frontend jest zaktualizowany?
- Czy testy kontraktu lub integracyjne przechodzą?
- Czy stary klient nie zostanie zepsuty?
- Czy dokumentacja funkcji została zaktualizowana?
- Czy zmiana jest opisana w `CHANGELOG.md`?
- Czy statusy HTTP są zgodne z dotychczasowym stylem?
