# language: pl
Właściwość: Pasek narzędzi — formatowanie tekstu
  Jako użytkownik D2 Tools
  Chcę formatować tekst za pomocą paska narzędzi
  Aby nadawać dokumentowi odpowiedni wygląd

  Tło:
    Zakładając że otwieram aplikację D2
    I wpisuję tekst "Tekst testowy" w edytorze
    I zaznaczam cały tekst w edytorze

  # ────────────────────────────────────────────
  # Formatowanie inline
  # ────────────────────────────────────────────

  @ui @smoke
  Szablon scenariusza: Zastosowanie formatowania inline
    Kiedy klikam przycisk "<przycisk>" na pasku narzędzi
    Wtedy tekst jest sformatowany jako "<format>"

    Przykłady:
      | przycisk      | format        |
      | Pogrubienie   | bold          |
      | Kursywa       | italic        |
      | Podkreślenie  | underline     |

  # ────────────────────────────────────────────
  # Style blokowe
  # ────────────────────────────────────────────

  @ui @regression
  Szablon scenariusza: Zmiana stylu blokowego
    Kiedy wybieram styl "<styl>" z dropdown
    Wtedy aktualny styl to "<styl>"

    Przykłady:
      | styl        |
      | Normalny    |
      | Nagłówek 1  |
      | Nagłówek 2  |
      | Nagłówek 3  |
      | Tytuł       |
      | Podtytuł    |

  # ────────────────────────────────────────────
  # Czcionki
  # ────────────────────────────────────────────

  @ui @regression
  Szablon scenariusza: Zmiana rozmiaru czcionki
    Kiedy ustawiam rozmiar czcionki na <rozmiar>
    Wtedy rozmiar czcionki wynosi "<rozmiar>"

    Przykłady:
      | rozmiar |
      | 8       |
      | 12      |
      | 24      |
      | 48      |
      | 72      |

  @ui @regression
  Szablon scenariusza: Zmiana rodziny czcionki
    Kiedy wybieram czcionkę "<czcionka>"
    Wtedy wybrana czcionka to "<czcionka>"

    Przykłady:
      | czcionka          |
      | Arial             |
      | Times New Roman   |
      | Courier New       |

  # ────────────────────────────────────────────
  # Wyrównanie
  # ────────────────────────────────────────────

  @ui @regression
  Szablon scenariusza: Zmiana wyrównania tekstu
    Kiedy klikam wyrównanie "<wyrownanie>"
    Wtedy tekst jest wyrównany "<wyrownanie>"

    Przykłady:
      | wyrownanie   |
      | Do lewej     |
      | Wyśrodkuj    |
      | Do prawej    |
      | Wyjustuj     |

  # ────────────────────────────────────────────
  # Listy
  # ────────────────────────────────────────────

  @ui @regression
  Scenariusz: Wstawienie listy punktowanej
    Kiedy klikam przycisk "Lista punktowana" na pasku narzędzi
    Wtedy tekst jest na liście punktowanej

  @ui @regression
  Scenariusz: Wstawienie listy numerowanej
    Kiedy klikam przycisk "Lista numerowana" na pasku narzędzi
    Wtedy tekst jest na liście numerowanej
