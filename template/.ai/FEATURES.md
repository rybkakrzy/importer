# Features

## Cel pliku

Ten plik opisuje funkcje systemu z perspektywy produktu i implementacji. Agent powinien go aktualizować, gdy powstaje nowa funkcja albo zmienia się istniejąca.

## Aktualny status funkcji

Uzupełnij po dodaniu katalogu do repo.

| Funkcja | Status | Backend | Frontend | API | Uwagi |
|---|---|---|---|---|---|
| Authentication | Unknown | Unknown | Unknown | Unknown | Sprawdzić w repo |
| User management | Unknown | Unknown | Unknown | Unknown | Sprawdzić w repo |
| Main business workflow | Unknown | Unknown | Unknown | Unknown | Dostosować do domeny |
| Administration | Unknown | Unknown | Unknown | Unknown | Dostosować do domeny |
| Reporting / dashboard | Unknown | Unknown | Unknown | Unknown | Dostosować do domeny |

## Statusy

- `Unknown` — niezweryfikowane w repo.
- `Planned` — zaplanowane, bez implementacji.
- `In Progress` — praca rozpoczęta.
- `Blocked` — praca zablokowana.
- `Implemented` — implementacja gotowa.
- `Verified` — gotowe i zweryfikowane testami / buildem / review.
- `Deprecated` — funkcja wycofywana lub zastąpiona.

## Szablon opisu funkcji

```md
## Feature: <nazwa>

### Cel

<po co istnieje funkcja>

### Użytkownicy

- <typ użytkownika>

### Główne scenariusze

1. <scenariusz>
2. <scenariusz>

### Reguły biznesowe

- <reguła>

### Backend

- Endpointy:
  - `GET /api/...`
  - `POST /api/...`
- Use case:
  - `<nazwa>`
- Encje / modele:
  - `<nazwa>`

### Frontend

- Route:
  - `/...`
- Komponenty:
  - `<nazwa>`
- Serwisy:
  - `<nazwa>`

### API

- Request DTO:
  - `<nazwa>`
- Response DTO:
  - `<nazwa>`
- Error cases:
  - `<kod/status>`

### Testy

- Unit:
  - <zakres>
- Integration:
  - <zakres>
- E2E:
  - <zakres>

### Otwarte pytania

- <pytanie>
```

## Feature: Authentication

### Cel

Zapewnić kontrolowany dostęp do aplikacji.

### Status

`Unknown` — agent musi sprawdzić, czy projekt ma już auth.

### Typowe elementy do sprawdzenia

Backend:

- konfiguracja authentication/authorization,
- JWT/cookies/OIDC,
- policies/roles/claims,
- endpointy login/logout/refresh, jeżeli są lokalne,
- ochrona endpointów.

Frontend:

- login page,
- auth guard,
- interceptor tokenów,
- obsługa 401/403,
- przechowywanie sesji.

### Reguła

Frontend może ukrywać elementy UI, ale backend zawsze musi egzekwować uprawnienia.

## Feature: Main Business Workflow

### Cel

`TODO: opisz główny proces biznesowy`

### Status

`Unknown`

### Uwagi dla agenta

To jest najważniejszy obszar domenowy. Przed zmianą agent musi przeczytać:

- `DOMAIN.md`,
- `API_CONTRACTS.md`,
- pliki implementacji funkcji,
- testy, jeżeli istnieją.
