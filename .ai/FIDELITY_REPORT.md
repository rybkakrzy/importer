# DOCX → HTML Fidelity Report

> Raport zgodności odwzorowania DOCX w edytorze webowym względem Microsoft Word.
> Stan: 2026-06-03, po refaktoryzacji konwersji (Etapy 0–6 + 9 z planu 10-etapowego).
> Plan: `~/.claude/plans/fancy-spinning-wirth.md`. Decyzje: `DECISIONS.md` ADR-0008/0009.
> Reference techniczny konwersji (pipeline, model, macierz statusów, diagramy): `DOCX_CONVERSION.md`.

## 1. Cel i podejście

Refaktoryzacja konwersji `DocxToHtmlConverter` (2950 l.) / `HtmlToDocxConverter` (2525 l.)
pod maksymalną, **mierzalną** wierność z Wordem — bez psucia round-tripu, autozapisu
i kontraktów API. Realizacja strangler-pattern w bezpiecznych, otestowanych krojach,
każdy osłonięty harnessem regresji (golden DOCX + snapshoty HTML).

## 2. Co zrobiono (Etapy)

| Etap | Zakres | Status |
|---|---|---|
| 0 | Harness regresji: golden DOCX in-memory + snapshoty HTML (normalizacja base64/id/dat) | ✅ |
| 1 | `OoxmlUnits` — centralne jednostki (twips/EMU/half-points/cm/px), koniec magicznych stałych (reguła 14) | ✅ |
| 2 | Jawny model pośredni (read-side, strangler): `PageSettings` + `SectionPropertiesReader`; migracja sekcji/strony | ✅ (sekcja/strona) |
| 3 | Rozwiązywanie nazwanych stylów znakowych `w:rStyle` (dziedziczenie; direct wygrywa) | ◐ slice |
| 4 | Rozmiar + orientacja strony — **pełny round-trip full-stack** (Domain→Api→Angular→Vitest) | ✅ |
| 5 | Tabele: `<colgroup>` z `tblGrid` + `table-layout:fixed`; fix `rowspan` przy `vMerge`+`gridSpan` | ◐ slice |
| 6 | Obrazy: tekst alternatywny `alt` ↔ `wp:docPr/@descr` (na bazie wcześniejszego round-tripu floating/border/crop/EMU) | ◐ slice |
| 9 | Ten raport zgodności | ✅ |

## 3. Macierz wierności

### Wysoka wierność (odwzorowane dobrze, otestowane)
- **Jednostki**: twips/EMU/half-points/cm/px — jedno źródło prawdy, DPI 96, testy referencyjne.
- **Marginesy strony** (cm), **rozmiar + orientacja strony** (A4/A5/landscape) — pełny round-trip.
- **Typografia bezpośrednia**: bold/italic/underline/strike, font-family (z theme), font-size (half-points), kolor (+ theme/auto), highlight, letter-spacing, sub/superscript.
- **Akapit**: alignment (L/C/R/justify), wcięcia (left/right/first-line/hanging), spacing before/after, line-height (exact/atLeast/auto).
- **Style**: akapitowe z `basedOn` (łańcuch) + **znakowe `w:rStyle`** (Etap 3); docDefaults (font/rozmiar) na kontenerze.
- **Tabele**: szerokości kolumn z `tblGrid` (colgroup + fixed layout), borders (per-komórka + domyślne), shading, cell margins, scalenia `gridSpan`/`vMerge` (w tym sąsiadujące — Etap 5), wyrównanie, wysokość wiersza.
- **Obrazy**: rozmiar (EMU, proporcje), inline + floating (`wp:anchor` front/behind + offsety), border (`a:ln`), crop (`a:srcRect`), **alt** (Etap 6), obrazy w nagłówku/stopce.
- **Nagłówki/stopki**: warianty default / first-page (`titlePg`) / even-odd (`evenAndOddHeaders`) — round-trip; geometria pasma (margines − dystans); font/rozmiar/kolor z docDefaults; układ lewo⇥środek⇥prawo przez tab-stopy (Etap 7, flex).
- **Izolacja CSS dokumentu (Etap 8 — zweryfikowane)**: style inline dokumentu wygrywają z regułami klasowymi przez specyficzność CSS. Audyt `!important`: w `wysiwyg-editor.scss` jest ich **6** (nie 65 — wcześniejsza liczba to artefakt zbiorczego grepu), z czego tylko 1 dotyczy treści (świadoma zamiana Calibri Light→Calibri w nagłówkach, by uniknąć zbyt cienkiego kroju); reszta to interakcja edytora (user-select/cursor przy resize/drag). `!important` w `styles.scss`/`document-editor.scss` dotyczą podświetleń UI (zaznaczenie find, zaznaczenie komórki) i paddingów dialogów — **nie** nadpisują font-size/koloru/marginesów treści. Wniosek: izolacja zdrowa, brak ryzykownej chirurgii SCSS.

### Częściowa wierność (działa z ograniczeniami)
- **Crop obrazu**: `clip-path: inset()` pokazuje wykadrowany obszar, ale **nie skaluje** go do pierwotnego boxa jak Word (box zachowuje rozmiar extentu). Wizualnie zbliżone, nie 1:1.
- **Style znakowe + jawne wyłączenie**: `bold=false` na runie nie nadpisze pogrubienia ze stylu znakowego (brak `font-weight:normal` z direct) — rzadkie.
- **Computed-style**: Etap 3 to slice (rStyle); pełny computed-style (folding docDefaults do każdego elementu, style numerowania, linked styles, pełny theme) — niedokończony. Obecnie polega częściowo na kaskadzie CSS.
- **Listy**: numbering.xml, punktatory (w tym picture bullets), zagnieżdżenia — działają; pełna wierność restart/custom format — częściowa.

### Niewspierane / niemożliwe w HTML/CSS (z przyczyną techniczną)
- **Zawijanie tekstu wokół obrazu „przylegle/na wskroś" (tight/through)**: wymaga reflow tekstu wokół kształtu — niedostępne natywnie w HTML/CSS (tylko `float` ≈ square). Panel edytora ma te tryby disabled.
- **Dokładna paginacja Worda** (podział na strony, `keep-with-next`, wdowy/sieroty): silnik layoutu przeglądarki ≠ silnik Worda; numery stron `X z Y` zależne od layoutu.
- **Kerning / hinting czcionek / dokładne metryki**: renderowanie przeglądarki różni się od Worda; przy braku zainstalowanej czcionki — fallback.
- **Wiele sekcji z różnymi nagłówkami/stopkami i geometrią** (R-10): model `DocumentContent` niesie pojedynczy header/footer + marginesy pierwszej sekcji.
- **Tab-stopy w nagłówku/stopce** (układ „lewo⇥środek⇥prawo" w jednej linii przez `w:tabs`): tab renderowany jako stała przerwa, nie pozycjonowany tab-stop.

## 4. Pokrycie testami (mierzalność)

- **Harness regresji** (`Infrastructure.UnitTests/Golden/`): golden DOCX in-memory (`GoldenDocuments`) + snapshoty HTML (`HtmlSnapshot`, approve-on-missing, normalizacja niestabilnych fragmentów) + asercje punktowe jednostek/geometrii.
- **Testy modelu**: `SectionPropertiesReaderTests` (Etap 2).
- **Round-tripy**: `PageSizeRoundTripTests`, `ImageAltTextRoundTripTests`, `ImageBorderCropRoundTripTests`, `ImageFloatingRoundTripTests`, `HeaderFooterRoundTripTests`, `*FidelityTests`.
- **Jednostki**: `OoxmlUnitsTests` (wartości referencyjne 1 cal = 1440 tw = 914400 EMU = 96 px = 72 pt = 2.54 cm).
- **Front**: Vitest round-trip `pageSize` w `document-editor.spec`.
- **Wyniki**: backend **358 pass / 0 fail / 6 skip**; GUI **166 pass**. Golden-snapshoty Etapów 0–2 niezmienione → dowód braku regresji w refaktorze jednostek/modelu.

Uruchomienie:
```
dotnet test D2ApiViewerEditor/D2ViewerEditor.sln
cd D2GuiViewerEditor && npx ng test --watch=false
```

## 5. Rekomendowane kolejne kroki (priorytet)

1. **Etap 7 — nagłówki/stopki, tab-stopy i wiele sekcji** (R-10): model wielosekcyjny + odwzorowanie `w:tabs` (center/right tab) dla układu L/C/R w jednej linii.
2. **Etap 3 cd. — pełny computed-style resolver**: docDefaults folding, style numerowania, linked styles, pełny theme; eliminacja sklejania stringów CSS + regex-dedup.
3. **Etap 8 — izolacja CSS (zweryfikowane, bez zmian)**: audyt wykazał, że treść jest już izolowana (inline > klasy; brak `!important` nadpisujących treść). Ewentualny dalszy krok: defensywny scoped reset na kontenerze treści + test computed-style (ograniczony w jsdom).
4. **Etap 5 cd. — tabele**: pełne `tcMar`/`tblCellMar`, cellSpacing, rozwiązywanie konfliktów obramowań; `table-layout:fixed` także po stronie zapisu.
5. **Etap 6 cd. — obrazy**: rotacja (`a:xfrm/@rot`), skalowanie przy crop, VML (legacy) border/crop/alt.
6. **Visual regression (Playwright)**: screenshot edytora dla golden DOCX + baseline z tolerancją (po izolacji CSS).

## 6. Ograniczenia i ryzyka (skrót)

- Snapshoty Golden lokują **obecne** zachowanie (regression-guard), nie poprawność pikselową wobec Worda — po świadomej poprawie wierności baseline aktualizować i przeglądać diff (R-12).
- 100% zgodności z Wordem w HTML/CSS jest nieosiągalne (paginacja, tight-wrap, metryki czcionek) — cel to maksymalna **mierzalna** wierność, nie 1:1.
- Pełna lista: `RISKS_ASSUMPTIONS.md` (R-10..R-14).
