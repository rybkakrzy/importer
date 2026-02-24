# language: pl
Właściwość: Edytor dokumentów — podstawowe operacje
  Jako użytkownik D2 Tools
  Chcę móc tworzyć i edytować dokumenty
  Aby przygotowywać dokumenty Word online

  Tło:
    Zakładając że otwieram aplikację D2

  # ────────────────────────────────────────────
  # Smoke
  # ────────────────────────────────────────────

  @ui @smoke
  Scenariusz: Aplikacja ładuje się poprawnie
    Wtedy widzę nagłówek aplikacji z napisem "Doc2"
    I widzę pasek narzędzi edytora
    I widzę obszar edycji dokumentu

  @ui @smoke
  Scenariusz: Wpisanie tekstu w edytorze
    Kiedy wpisuję tekst "Witaj Świecie" w edytorze
    Wtedy edytor zawiera tekst "Witaj Świecie"

  # ────────────────────────────────────────────
  # Nowy dokument
  # ────────────────────────────────────────────

  @ui @regression
  Scenariusz: Utworzenie nowego dokumentu
    Kiedy wpisuję tekst "Istniejący tekst" w edytorze
    I wybieram z menu "Plik" opcję "Nowy dokument"
    Wtedy edytor jest pusty

  # ────────────────────────────────────────────
  # Szablony
  # ────────────────────────────────────────────

  @ui @regression
  Szablon scenariusza: Załadowanie szablonu dokumentu
    Kiedy otwieram dialog szablonów
    I wybieram szablon "<nazwa_szablonu>"
    Wtedy edytor zawiera tekst "<oczekiwany_tekst>"

    Przykłady:
      | nazwa_szablonu | oczekiwany_tekst    |
      | List           | Szanowni Państwo    |
      | Raport         | Raport              |
      | CV             | Imię                |

  # ────────────────────────────────────────────
  # Undo / Redo
  # ────────────────────────────────────────────

  @ui @regression
  Scenariusz: Cofanie i ponawianie zmian
    Kiedy wpisuję tekst "Pierwszy" w edytorze
    I wykonuję cofnij
    Wtedy edytor jest pusty
    Kiedy wykonuję ponów
    Wtedy edytor zawiera tekst "Pierwszy"
