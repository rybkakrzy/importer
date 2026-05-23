# Frontend Angular Guidelines

## Zasada podstawowa

Najpierw sprawdź istniejący styl frontendu:

- `package.json`,
- `angular.json`,
- strukturę `src/app`,
- wzorzec routingu,
- sposób obsługi formularzy,
- sposób komunikacji z API,
- sposób zarządzania stanem.

Nie migruj projektu do nowego stylu Angulara bez wyraźnego zadania.

## TypeScript

- Używaj silnego typowania.
- Unikaj `any`. Jeżeli jest konieczne, ogranicz zakres i dodaj komentarz wyjaśniający powód.
- Preferuj jawne typy modeli API.
- Nie mieszaj typów DTO z typami formularzy, jeżeli mają różną strukturę.
- Unikaj dużych plików z wieloma niepowiązanymi typami.
- Nie używaj non-null assertion `!` jako sposobu na ukrycie problemu.
- Preferuj typy związkowe dla stanów zamiast magic stringów.

## Angular components

Komponent powinien:

- mieć jedną odpowiedzialność,
- nie zawierać ciężkiej logiki biznesowej,
- delegować komunikację z API do serwisu,
- mieć czytelny template,
- unikać złożonych wyrażeń w HTML,
- mieć uporządkowane wejścia i wyjścia,
- obsługiwać loading/error/empty state.

## Standalone components

Jeżeli projekt używa standalone components, kontynuuj ten styl.

Jeżeli projekt używa NgModules, nie migruj go przypadkowo na standalone components przy okazji zwykłej funkcji.

## Signals, RxJS i stan

- Stosuj wzorzec obecny w projekcie.
- Nie mieszaj bez powodu kilku stylów zarządzania stanem w jednej funkcji.
- Dla operacji asynchronicznych zachowaj czytelny lifecycle subskrypcji.
- Unikaj ręcznych subskrypcji w komponentach, jeżeli można użyć `async` pipe lub istniejącego mechanizmu projektu.
- Nie twórz globalnego store dla lokalnego stanu komponentu.
- Nie trzymaj stanu domenowego w przypadkowych singletonach.
- Czyść subskrypcje zgodnie ze stylem projektu.

## Formularze

- Preferuj Reactive Forms, jeżeli projekt ich używa.
- Walidacja powinna być czytelna i spójna z backendem.
- Komunikaty błędów powinny być konkretne i przyjazne użytkownikowi.
- Nie duplikuj skomplikowanych reguł domenowych w frontendzie, jeżeli źródłem prawdy jest backend.
- Frontend może walidować UX-owo, ale backend musi walidować bezpieczeństwo i reguły biznesowe.

## API services

- Komunikację HTTP trzymaj w serwisach.
- Nie wywołuj `HttpClient` bezpośrednio z komponentów, jeżeli projekt ma warstwę data-access.
- Typuj requesty i response'y.
- Obsługuj błędy zgodnie z istniejącym wzorcem.
- Nie buduj stringów URL w wielu miejscach, jeżeli projekt ma centralną konfigurację endpointów.
- Nie ignoruj błędów HTTP bez świadomej obsługi.

## UI

- Zachowaj spójność z istniejącą biblioteką komponentów.
- Nie dodawaj nowej biblioteki UI bez decyzji.
- Dbaj o podstawową dostępność:
  - labelki formularzy,
  - sensowne teksty przycisków,
  - obsługę loading/error/empty state,
  - poprawną semantykę HTML,
  - focus management dla dialogów i formularzy.

## Testy frontendu

Dla ważnej logiki dodawaj testy:

- komponentów,
- serwisów,
- mapperów,
- guardów,
- walidatorów,
- krytycznych flow E2E, jeżeli projekt ma E2E.

## Przykładowy styl komentarzy

```ts
// Backend returns dates in UTC; keep conversion in one place to avoid timezone bugs.
const createdAt = this.dateTimeMapper.fromUtc(response.createdAt);
```

Komentarz wyjaśnia decyzję, nie oczywistą operację.
