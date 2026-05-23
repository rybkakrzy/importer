# Coding Standards

## Zasady ogólne

- Kod ma być czytelny, prosty i możliwy do testowania.
- Preferuj małe klasy, małe funkcje i jasne nazwy.
- Nie twórz abstrakcji bez realnej potrzeby.
- Nie ukrywaj złożoności w helperach o niejasnej odpowiedzialności.
- Nie mieszaj zmian funkcjonalnych z masowym formatowaniem.
- Nie poprawiaj przypadkowo całego pliku, jeżeli zadanie dotyczy jednej funkcji.
- Komentarze w kodzie powinny być po angielsku.
- Komentarze mają tłumaczyć „dlaczego”, a nie „co” oczywistego robi kod.
- Nie dodawaj komentarzy typu `summary` przed każdą metodą.
- Nie dodawaj TODO zamiast implementacji, chyba że to świadomie uzgodniony etap.

## Nazewnictwo

- Nazwy powinny opisywać intencję biznesową lub techniczną.
- Unikaj skrótów, chyba że są powszechnie przyjęte w projekcie.
- Unikaj nazw typu `Manager`, `Helper`, `Processor`, jeżeli można nazwać odpowiedzialność precyzyjniej.
- Metody wykonujące operację powinny zwykle zaczynać się od czasownika.
- Klasy domenowe powinny używać języka domenowego z `DOMAIN.md`.
- DTO powinny jasno mówić, czy są requestem, responsem, commandem albo query.

## Projektowanie metod

Metoda powinna:

- robić jedną rzecz,
- mieć jasny poziom abstrakcji,
- nie przekraczać rozsądnej długości,
- unikać głębokiego zagnieżdżenia,
- nie przyjmować wielu luźno powiązanych parametrów.

Jeżeli metoda wymaga wielu parametrów, rozważ:

- obiekt komendy,
- DTO requestu,
- value object,
- obiekt konfiguracji,
- rozbicie odpowiedzialności.

## Obsługa błędów

- Nie połykaj wyjątków.
- Nie loguj danych wrażliwych.
- Komunikaty błędów powinny być użyteczne, ale nie powinny ujawniać sekretów ani szczegółów infrastruktury.
- Walidacja wejścia powinna być jawna i testowalna.
- Nie używaj wyjątków do zwykłego sterowania przepływem, jeżeli projekt ma lepszy wzorzec.
- Nie zamieniaj wszystkich błędów na generyczne `Exception`.

## Refaktoryzacja

Refaktoryzuj tylko wtedy, gdy:

- jest to wymagane przez zadanie,
- zmniejsza ryzyko zmiany,
- usuwa realny dług techniczny w dotykanym obszarze,
- jest pokryte testami lub łatwe do ręcznej weryfikacji.

Nie wykonuj dużych refaktoryzacji przy okazji małych poprawek.

## Minimalny standard przed zakończeniem pracy

Agent powinien sprawdzić:

- czy kod się kompiluje,
- czy testy związane ze zmianą przechodzą,
- czy nie dodano niepotrzebnej zależności,
- czy nie naruszono kontraktu API,
- czy nie dodano sekretów,
- czy nie zostawiono debugowego kodu,
- czy zaktualizowano `.ai/` tam, gdzie było to potrzebne.
