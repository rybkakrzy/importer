# Testing and Quality

## Cel

Testy mają chronić zachowanie systemu, a nie tylko zwiększać coverage.

## Hierarchia weryfikacji

Agent powinien dobrać weryfikację do zmiany:

1. Kompilacja / type-check.
2. Unit tests dla logiki.
3. Integration tests dla API, bazy, integracji.
4. Component tests dla UI.
5. E2E dla krytycznych ścieżek.
6. Manual smoke test, jeżeli automatyczne testy nie istnieją.

## Backend

Typowe komendy do potwierdzenia w repo:

```bash
dotnet restore
dotnet build
dotnet test
```

Dla backendu testuj:

- reguły domenowe,
- walidację,
- mapping DTO,
- use case'y,
- autoryzację,
- statusy HTTP,
- operacje na bazie danych,
- transakcje,
- przypadki błędów.

## Frontend

Typowe komendy do potwierdzenia w repo:

```bash
npm install
npm run build
npm test
npm run lint
```

Dla frontendu testuj:

- komponenty z logiką,
- serwisy API,
- mappery,
- walidatory formularzy,
- routing i guardy,
- krytyczne flow użytkownika,
- obsługę loading/error/empty state.

## Jakość kodu

Przed zakończeniem pracy agent powinien sprawdzić:

- brak oczywistych dead code,
- brak nieużywanych importów,
- brak przypadkowych `console.log`,
- brak twardo wpisanych sekretów,
- brak debugowych endpointów,
- brak dużych niespójnych refaktoryzacji,
- brak formatowania niezwiązanego z zadaniem,
- brak wyciszenia błędów typowania bez powodu.

## Gdy testy nie istnieją

Jeżeli projekt nie ma testów dla danego obszaru:

- nie udawaj, że zmiana jest w pełni zweryfikowana,
- opisz, co zostało sprawdzone ręcznie lub przez build,
- zaproponuj minimalny test zabezpieczający zmianę,
- nie twórz ogromnego frameworka testowego przy okazji małej poprawki.

## Minimalna macierz weryfikacji

| Typ zmiany | Minimalna weryfikacja |
|---|---|
| Zmiana C# bez API | `dotnet build`, powiązane testy |
| Zmiana API | `dotnet build`, test endpointu/use case, sprawdzenie frontendu |
| Zmiana Angular komponentu | `npm run build`, test komponentu jeżeli istnieje |
| Zmiana serwisu API w Angularze | `npm run build`, test serwisu/mappingu |
| Zmiana modelu danych | migracja, test integracyjny albo ręczna procedura |
| Zmiana Docker/CI | build obrazu albo walidacja pipeline |
| Zmiana auth/security | test pozytywny i negatywny dostępu |

## Raport końcowy agenta

Raport powinien zawierać:

```md
## Changed

- ...

## Verified

- `command` — result

## Risks

- ...

## Next

- ...
```
