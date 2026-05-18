# Firmowa czcionka

Aby wgrać firmową czcionkę do edytora:

1. Skopiuj plik czcionki (np. `CorporateSans-Regular.woff2`) do tego katalogu.
   - Zalecane formaty: `.woff2` (najlepsza kompresja) lub `.woff`.
   - Możesz dorzucić więcej wariantów (bold, italic) — każdy w osobnym pliku.

2. Otwórz [`_corporate-font.scss`](./_corporate-font.scss) i:
   - Ustaw `$corporate-font-family` na nazwę kroju (np. `'CorporateSans'`).
   - Dla każdego pliku odkomentuj / dodaj blok `@font-face` wskazujący na właściwy plik.

3. Przebuduj front (`npm start` / `ng build`). Edytor automatycznie zacznie używać
   tej czcionki jako domyślnej dla każdego elementu, który nie ma własnej deklaracji
   `font-family`.

> **Strona API**: równolegle ustaw nazwę kroju w `appsettings.json`:
> ```json
> "DocumentDefaults": {
>   "FontFamily": "CorporateSans",
>   "HeadingFontFamily": "CorporateSans",
>   "FontSizePt": 11
> }
> ```
> Dzięki temu zapis HTML → DOCX (i odwrotnie) zachowa firmowy krój nawet dla
> akapitów / runów, które nie mają jawnie zadeklarowanej czcionki.

> **Uwaga**: nazwa kroju w `$corporate-font-family` (frontend) i `FontFamily`
> (backend) **muszą być identyczne** — inaczej DOCX otwarty w Wordzie pokaże
> inny krój niż edytor w przeglądarce.
