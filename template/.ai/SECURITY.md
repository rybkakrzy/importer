# Security

## Zasada podstawowa

Bezpieczeństwo ma pierwszeństwo przed wygodą implementacji.

## Sekrety

Agent nie może:

- drukować sekretów w odpowiedzi,
- commitować `.env`,
- commitować kluczy prywatnych,
- commitować certyfikatów,
- przenosić sekretów do dokumentacji,
- wklejać tokenów do testów,
- logować tokenów, haseł lub danych wrażliwych.

Jeżeli sekret pojawi się przypadkowo w pliku, agent powinien zgłosić problem i nie powielać wartości.

## Auth i autoryzacja

- Rozróżniaj authentication i authorization.
- Backend musi egzekwować uprawnienia niezależnie od frontendu.
- Nie opieraj bezpieczeństwa na ukrywaniu przycisków w UI.
- Dla zasobów użytkownika sprawdzaj ownership lub uprawnienie dostępu.
- Nie ujawniaj istnienia zasobów użytkownikom bez uprawnień, jeżeli może to być problemem.
- Testuj przynajmniej przypadek dostępu dozwolonego i zabronionego przy zmianach auth.

## Walidacja

- Waliduj dane na backendzie.
- Frontendowa walidacja jest dla UX, nie dla bezpieczeństwa.
- Nie ufaj danym z klienta.
- Nie ufaj ID przekazanemu w requestcie bez sprawdzenia dostępu.
- Waliduj pliki, rozmiary, typy MIME i rozszerzenia, jeżeli projekt obsługuje upload.
- Ograniczaj rozmiary requestów, paginacji i uploadów.

## Dane osobowe

Status: `TODO: opisać realny zakres danych osobowych`

Zasady:

- minimalizuj zakres danych,
- nie loguj danych osobowych bez potrzeby,
- maskuj dane w logach,
- nie pokazuj danych innym użytkownikom bez uprawnień,
- nie kopiuj danych produkcyjnych do przykładów,
- nie używaj realnych danych użytkowników w testach.

## API

- Używaj poprawnych statusów `401` i `403`.
- Nie ujawniaj stack trace klientowi.
- Nie zwracaj szczegółów infrastruktury w błędach.
- Ograniczaj paginację i rozmiary requestów.
- Rozważ rate limiting dla endpointów publicznych.
- Nie dodawaj debugowych endpointów na produkcji.

## Frontend

- Nie przechowuj sekretów w kodzie frontendu.
- Nie traktuj zmiennych środowiskowych frontendu jako tajnych.
- Uważaj na XSS przy renderowaniu HTML.
- Preferuj bezpieczne mechanizmy Angulara zamiast ręcznego manipulowania DOM.
- Nie zapisuj tokenów w niebezpieczny sposób bez sprawdzenia istniejącej strategii projektu.

## Logowanie

Nie loguj:

- haseł,
- tokenów,
- refresh tokenów,
- pełnych danych osobowych,
- numerów dokumentów,
- sekretów integracji,
- surowych payloadów z danymi wrażliwymi.

## Checklist security przy zmianie

- Czy backend sprawdza uprawnienia?
- Czy request jest walidowany?
- Czy błędy nie ujawniają szczegółów?
- Czy logi nie zawierają sekretów?
- Czy frontend nie przechowuje sekretów?
- Czy dane użytkownika nie wyciekają między kontami/tenantami?
