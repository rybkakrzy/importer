# language: pl
Właściwość: Wyszukiwanie i zamiana tekstu
  Jako użytkownik D2 Tools
  Chcę wyszukiwać i zamieniać tekst w dokumencie
  Aby szybko edytować treść

  Tło:
    Zakładając że otwieram aplikację D2
    I wpisuję tekst "Ala ma kota. Kota ma też Ola." w edytorze

  # ────────────────────────────────────────────
  # Wyszukiwanie
  # ────────────────────────────────────────────

  @ui @smoke
  Scenariusz: Otwarcie paska wyszukiwania
    Kiedy otwieram pasek wyszukiwania
    Wtedy pasek wyszukiwania jest widoczny

  @ui @regression
  Scenariusz: Wyszukanie tekstu w dokumencie
    Kiedy otwieram pasek wyszukiwania
    I wyszukuję tekst "kota"
    Wtedy liczba wyników wyszukiwania jest większa niż 0

  @ui @regression
  Scenariusz: Wyszukanie tekstu bez wyników
    Kiedy otwieram pasek wyszukiwania
    I wyszukuję tekst "xyz_brak_tekstu"
    Wtedy liczba wyników wyszukiwania wynosi 0
