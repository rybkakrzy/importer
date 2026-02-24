# language: pl
Właściwość: Dialog kodu kreskowego
  Jako użytkownik D2 Tools
  Chcę generować kody kreskowe i QR
  Aby wstawiać je do dokumentów

  Tło:
    Zakładając że otwieram aplikację D2
    I otwieram dialog kodu kreskowego

  # ────────────────────────────────────────────
  # Smoke
  # ────────────────────────────────────────────

  @ui @smoke
  Scenariusz: Dialog kodu kreskowego otwiera się poprawnie
    Wtedy dialog kodu kreskowego jest widoczny
    I przycisk "Wstaw" jest nieaktywny

  # ────────────────────────────────────────────
  # Generowanie kodów
  # ────────────────────────────────────────────

  @ui @regression
  Szablon scenariusza: Generowanie kodu kreskowego
    Kiedy wybieram typ kodu "<typ>"
    I wpisuję zawartość "<zawartosc>"
    Wtedy podgląd kodu kreskowego jest widoczny
    I przycisk "Wstaw" jest aktywny

    Przykłady:
      | typ        | zawartosc           |
      | QRCode     | https://example.com |
      | Code128    | ABC-12345           |
      | EAN13      | 5901234123457       |

  @ui @regression
  Scenariusz: Wstawianie kodu QR do dokumentu
    Kiedy wybieram typ kodu "QRCode"
    I wpisuję zawartość "Test QR"
    I klikam "Wstaw"
    Wtedy dialog kodu kreskowego jest zamknięty
    I edytor zawiera obraz kodu kreskowego

  @ui @regression
  Scenariusz: Anulowanie dialogu kodu kreskowego
    Kiedy wpisuję zawartość "test"
    I klikam "Anuluj"
    Wtedy dialog kodu kreskowego jest zamknięty

  # ────────────────────────────────────────────
  # Parametry wymiarów
  # ────────────────────────────────────────────

  @ui @regression
  Szablon scenariusza: Zmiana wymiarów kodu kreskowego
    Kiedy wpisuję zawartość "test"
    I ustawiam szerokość na <szerokosc>
    I ustawiam wysokość na <wysokosc>
    Wtedy podgląd kodu kreskowego jest widoczny

    Przykłady:
      | szerokosc | wysokosc |
      | 100       | 100      |
      | 300       | 300      |
      | 500       | 500      |
