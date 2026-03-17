# language: pl
Właściwość: Podpisywanie dokumentu cyfrowym certyfikatem
  Jako użytkownik D2 Tools
  Chcę podpisywać dokumenty cyfrowym certyfikatem
  Aby zapewnić autentyczność i integralność dokumentu

  Tło:
    Zakładając że otwieram aplikację D2
    I wpisuję tekst "Dokument do podpisania" w edytorze

  # ────────────────────────────────────────────
  # Dialog — otwarcie i zamknięcie
  # ────────────────────────────────────────────

  @ui @smoke
  Scenariusz: Dialog podpisywania otwiera się z menu Narzędzia
    Kiedy otwieram dialog podpisywania
    Wtedy dialog podpisywania jest widoczny
    I przycisk "Podpisz" jest nieaktywny

  @ui @regression
  Scenariusz: Zamknięcie dialogu podpisywania przez Anuluj
    Kiedy otwieram dialog podpisywania
    I klikam "Anuluj" w dialogu podpisywania
    Wtedy dialog podpisywania jest zamknięty

  # ────────────────────────────────────────────
  # Walidacja formularza
  # ────────────────────────────────────────────

  @ui @regression
  Scenariusz: Przycisk Podpisz wymaga imienia i nazwiska podpisującego
    Kiedy otwieram dialog podpisywania
    I wpisuję imię podpisującego "Jan Kowalski"
    Wtedy przycisk "Podpisz" jest nadal nieaktywny bez certyfikatu

  @ui @regression
  Scenariusz: Walidacja adresu email podpisującego
    Kiedy otwieram dialog podpisywania
    I wpisuję imię podpisującego "Jan Kowalski"
    I wpisuję email podpisującego "nieprawidlowy-email"
    Wtedy pole email wskazuje błąd walidacji
