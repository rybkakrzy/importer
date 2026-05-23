# Backend .NET Guidelines

## Zasada podstawowa

Najpierw sprawdź istniejący styl backendu. Jeżeli projekt ma własne konwencje, stosuj je. Poniższe zasady są domyślne.

## C#

- Używaj nullable reference types, jeżeli są włączone.
- Preferuj `async/await` dla operacji I/O.
- Metody asynchroniczne powinny kończyć się sufiksem `Async`, jeżeli taki styl jest stosowany w projekcie.
- Przekazuj `CancellationToken` przez warstwy aplikacji, zwłaszcza dla operacji I/O.
- Preferuj niemutowalne modele tam, gdzie to praktyczne.
- Unikaj `dynamic`.
- Unikaj reflection bez mocnego uzasadnienia.
- Nie używaj statycznego globalnego stanu dla logiki biznesowej.
- Nie wyciszaj nullable warnings operatorem `!` bez powodu.

## ASP.NET Core

Endpoint/kontroler powinien:

- przyjąć request,
- uruchomić walidację,
- wywołać use case / service,
- zmapować wynik na HTTP response.

Nie powinien zawierać ciężkiej logiki biznesowej.

## HTTP status codes

Domyślnie:

- `200 OK` — odczyt albo operacja ze zwracanym body,
- `201 Created` — utworzenie zasobu,
- `204 No Content` — skuteczna operacja bez body,
- `400 Bad Request` — technicznie błędny request,
- `401 Unauthorized` — brak uwierzytelnienia,
- `403 Forbidden` — brak uprawnień,
- `404 Not Found` — brak zasobu,
- `409 Conflict` — konflikt stanu,
- `422 Unprocessable Entity` — walidacja domenowa, jeżeli projekt tak robi.

## Application Layer

Use case powinien:

- mieć jedną odpowiedzialność,
- używać jasnego request/command/query modelu,
- zawierać orkiestrację, nie szczegóły infrastruktury,
- być testowalny bez prawdziwej bazy danych,
- obsługiwać transakcję, jeżeli wymaga tego spójność,
- przyjmować `CancellationToken` przy operacjach I/O.

## Domain Layer

- Reguły biznesowe trzymaj możliwie blisko modelu domenowego.
- Nie dopuszczaj do tworzenia niepoprawnych obiektów domenowych.
- Preferuj value objects dla wartości z regułami, np. email, money, date range.
- Nie dodawaj zależności infrastrukturalnych do domeny.
- Nie przekazuj encji ORM jako kontraktu API.

## Entity Framework Core

Jeżeli projekt używa EF Core:

- Nie generuj migracji bez rzeczywistej zmiany modelu danych.
- Nie zmieniaj schematu bazy przy okazji zmian UI.
- Sprawdzaj wpływ migracji na istniejące dane.
- Uważaj na lazy loading i problem N+1.
- Dla odczytów bez modyfikacji używaj `AsNoTracking()`, jeżeli pasuje do stylu projektu.
- Nie zwracaj encji EF bezpośrednio przez API.
- Dla skomplikowanych zapytań preferuj jawne projekcje do DTO.
- Nie wykonuj zapytań w pętli, jeżeli da się pobrać dane jednym zapytaniem.

## Dependency Injection

- Rejestruj zależności w miejscu zgodnym z projektem.
- Nie używaj Service Locator.
- Nie wstrzykuj zbyt wielu zależności do jednej klasy.
- Jeżeli klasa wymaga wielu zależności, sprawdź, czy nie ma zbyt szerokiej odpowiedzialności.
- Nie rejestruj przypadkowo serwisów stateful jako singleton.

## Validation

Preferowany kierunek:

- request validation blisko API/application layer,
- domain invariants w domenie,
- bezpieczeństwo i uprawnienia na backendzie,
- czytelne błędy dla frontendu.

## Logging

- Logi powinny pomagać w diagnostyce.
- Nie loguj haseł, tokenów, danych osobowych ani sekretów.
- Preferuj structured logging.
- Nie używaj `Console.WriteLine` w kodzie aplikacyjnym.
- Loguj kontekst techniczny, ale nie ujawniaj danych wrażliwych.

## Przykładowy styl komentarzy

```csharp
// Keep this check before database access to avoid leaking resource existence.
if (!currentUser.CanAccess(resourceId))
{
    return Result.Forbidden();
}
```

Komentarz wyjaśnia powód, a nie przepisuje kod.
