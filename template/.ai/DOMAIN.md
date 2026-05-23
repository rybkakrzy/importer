# Domain

## Cel pliku

Ten plik opisuje język domenowy, reguły biznesowe i model pojęciowy. Agent powinien go czytać przed zmianą logiki biznesowej.

## Status

Domena nie jest jeszcze uzupełniona. Po wrzuceniu do repo trzeba przepisać tutaj realne pojęcia projektu.

## Domyślne założenie

Projekt prawdopodobnie posiada:

- użytkowników,
- role lub uprawnienia,
- główne zasoby biznesowe,
- operacje tworzenia/edycji/usuwania,
- statusy procesu,
- walidację danych,
- historię lub audyt zmian.

Agent nie może implementować reguł domenowych na podstawie tych ogólnych założeń bez sprawdzenia repo i wymagań.

## Język domenowy

| Termin | Znaczenie | Status |
|---|---|---|
| User | Osoba lub konto korzystające z systemu | Do potwierdzenia |
| Role | Zestaw uprawnień lub kontekst działania użytkownika | Do potwierdzenia |
| Resource | Główny zasób biznesowy systemu | Do zastąpienia nazwą domenową |
| Status | Stan procesu lub zasobu | Do potwierdzenia |
| Audit | Informacja o tym, kto i kiedy wykonał operację | Do potwierdzenia |

## Encje domenowe

| Encja | Odpowiedzialność | Najważniejsze reguły | Status |
|---|---|---|---|
| User | Reprezentuje konto/użytkownika | Dostęp kontrolowany przez auth | Do potwierdzenia |
| MainEntity | Główny obiekt biznesowy | Reguły zależą od domeny | Do zastąpienia |
| AuditEntry | Ślad wykonanej operacji | Nie powinien zawierać sekretów | Do potwierdzenia |

## Typy wartości

Potencjalne value objects do rozważenia, jeżeli występują w domenie:

| Typ | Znaczenie | Walidacja |
|---|---|---|
| EmailAddress | Adres email | Format, długość, normalizacja |
| Money | Kwota pieniężna | Waluta, precyzja, zakres |
| DateRange | Zakres dat | Start <= end |
| ExternalId | Identyfikator z systemu zewnętrznego | Format, długość |
| Slug | Techniczny identyfikator w URL | Znaki, unikalność |

Nie dodawaj value objectów mechanicznie. Dodawaj je tylko, gdy upraszczają reguły i chronią inwarianty.

## Reguły biznesowe

| ID | Reguła | Źródło | Status |
|---|---|---|---|
| BR-001 | Backend jest źródłem prawdy dla walidacji i uprawnień | Standard projektu | Active |
| BR-002 | Frontend może walidować UX-owo, ale nie zastępuje backendu | Standard projektu | Active |
| BR-003 | Operacje modyfikujące dane powinny być autoryzowane | Standard projektu | Active |
| BR-004 | Błędy walidacji powinny być możliwe do pokazania użytkownikowi | Standard projektu | Active |

## Zasady dla agenta

- Nie zmieniaj nazewnictwa domenowego bez aktualizacji tego pliku.
- Nie upraszczaj reguł biznesowych tylko dlatego, że kod wyglądałby prościej.
- Jeżeli reguła jest niejasna, zapisz założenie w `RISKS_ASSUMPTIONS.md`.
- Jeżeli implementujesz nową regułę, dodaj ją do tabeli.
