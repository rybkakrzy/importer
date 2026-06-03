# DOCX ↔ HTML Conversion — reference techniczny

> Źródło prawdy o konwersji dokumentów w D2 ViewerEditor. Każdy fakt zweryfikowany w kodzie
> (ścieżki podane). Statusy oznaczone jawnie. Dokument komplementarny do:
> - `FEATURES.md` (sekcja nagłówek/stopka — opis funkcjonalny),
> - `FIDELITY_REPORT.md` (macierz wierności z Wordem),
> - `RISKS_ASSUMPTIONS.md` (R-10..R-14),
> - `DECISIONS.md` (ADR-0008 jednostki/harness, ADR-0009 model pośredni).
>
> Ostatnia weryfikacja względem kodu: 2026-06-03.

## 1. Przegląd

Konwersja jest dwukierunkowa i w pełni własna (parser/generator OOXML na bazie
`DocumentFormat.OpenXml` 3.0.2 — bez zewnętrznego renderera):

| Kierunek | Klasa | Plik | Kontrakt |
|---|---|---|---|
| DOCX → HTML (import/render) | `DocxToHtmlConverter` | `D2ViewerEditor.Infrastructure/Services/DocxToHtmlConverter.cs` | `IDocxToHtmlConverter.Convert(Stream) : DocumentContent` |
| HTML → DOCX (zapis/eksport) | `HtmlToDocxConverter` | `D2ViewerEditor.Infrastructure/Services/HtmlToDocxConverter.cs` | `IHtmlToDocxConverter.Convert(string html, …, PageMargins?, PageSize?) : byte[]` |

Interfejsy: `D2ViewerEditor.Domain/Interfaces/IDocxToHtmlConverter.cs`,
`IHtmlToDocxConverter.cs`. Warstwy pomocnicze:

- **Jednostki**: `D2ViewerEditor.Infrastructure/Conversion/OoxmlUnits.cs` — jedno źródło prawdy
  dla twips/EMU/half-points/cm/px (DPI 96). Patrz ADR-0008.
- **Model pośredni (read-side, częściowy)**: `D2ViewerEditor.Infrastructure/DocxModel/`
  (`PageSettings`, `SectionPropertiesReader`) — geometria sekcji/strony. Patrz ADR-0009.

> Architektura jest hybrydą: parsowanie sekcji/strony idzie przez jawny model
> (`PageSettings`), reszta (akapit/run/tabela/obraz) wciąż konwertuje **XML → string HTML**
> jednym przebiegiem metodami `Convert*ToHtml`. Pełny model pośredni jest wprowadzany
> strangler-pattern (ADR-0009) — nie jest kompletny.

### Pipeline DOCX → HTML

```mermaid
flowchart TD
    A[Stream DOCX] --> B[WordprocessingDocument.Open]
    B --> C[LoadThemeFonts / LoadNumberingPictureBullets]
    B --> D[LoadDocumentStyles + LoadDocDefaults]
    D --> E["_styles: styleId -> CSS (basedOn chain, regex-dedup)"]
    B --> F[ConvertBodyToHtml]
    F --> G[ConvertParagraphToHtml]
    G --> H[ConvertRunToHtml + rStyle]
    F --> I[ConvertTableToHtml + colgroup]
    H --> J[ConvertDrawingToHtml / ConvertPictureToHtml]
    B --> K[ExtractHeader / ExtractFooter]
    B --> L["ExtractPageMargins / ExtractPageSize (PageSettings)"]
    G --> M[DocumentContent.Html]
    K --> N[DocumentContent.Header/Footer]
    L --> O[DocumentContent.Margins/PageSize]
    M --> P[DocumentContent]
    N --> P
    O --> P
```

### Pipeline HTML → DOCX

```mermaid
flowchart TD
    A["HTML (editor.innerHTML)"] --> B[HtmlAgilityPack LoadHtml]
    B --> C[AddDocumentStyles: Normal/Heading/Header/Footer/ListParagraph]
    B --> D[ConvertHtmlToBody]
    D --> E["paragrafy/runy -> w:p/w:r (InnerText, style -> rPr/pPr)"]
    D --> F["table -> w:tbl (TableLayout Autofit; grid z liczby komorek)"]
    D --> G["img[data-*] -> w:drawing (inline/anchor, a:ln, a:srcRect, docPr descr)"]
    B --> H[AddHeaderAndFooter: Default + opcj. First/Even party]
    B --> I["AddPageSettings: w:pgSz (PageSize) + w:pgMar (PageMargins)"]
    E --> J[WordprocessingDocument]
    F --> J
    G --> J
    H --> J
    I --> J
    J --> K[byte array DOCX]
```

## 2. Model `DocumentContent`

Definicja: `D2ViewerEditor.Domain/Models/DocumentModels.cs`. Odpowiednik TS:
`D2GuiViewerEditor/src/app/models/document.model.ts`.

```mermaid
classDiagram
    class DocumentContent {
        string Html
        DocumentMetadata Metadata
        List~DocumentImage~ Images
        List~DocumentStyle~ Styles
        HeaderFooterContent Header
        HeaderFooterContent Footer
        PageMargins Margins
        PageSize PageSize
    }
    class HeaderFooterContent {
        string Html
        double Height
        bool DifferentFirstPage
        string FirstPageHtml
        bool DifferentOddEven
        string EvenHtml
    }
    class PageMargins {
        double Top
        double Bottom
        double Left
        double Right
    }
    class PageSize {
        double WidthCm
        double HeightCm
        string Orientation
    }
    class DocumentImage {
        string Id
        string ContentType
        string Base64Data
    }
    DocumentContent --> HeaderFooterContent : Header/Footer
    DocumentContent --> PageMargins
    DocumentContent --> PageSize
    DocumentContent --> DocumentImage : Images
```

**Kluczowe ograniczenie modelu (R-10):** `DocumentContent` reprezentuje dokument jako
**pojedynczą sekcję** — jeden `Header`, jeden `Footer`, jedne `Margins`, jeden `PageSize`,
jeden ciąg `Html`. Warianty first-page / even-odd są wspierane **w obrębie tej jednej
sekcji** (pola `DifferentFirstPage`/`FirstPageHtml`/`DifferentOddEven`/`EvenHtml`), ale
**wiele sekcji z różnymi nagłówkami/stopkami/marginesami/rozmiarami strony nie jest
reprezentowanych**. Reader i writer biorą `Body.Elements<SectionProperties>().FirstOrDefault()`.

Docelowe rozszerzenie modelu (Planned / not implemented yet) powinno objąć co najmniej:
kolekcję sekcji, `SectionProperties` per sekcja, osobne referencje header/footer
(default/first/even) per sekcja, `PageSettings` per sekcja oraz powiązanie bloków treści
z sekcją.

## 3. Jednostki (`OoxmlUnits`) — Implemented

`D2ViewerEditor.Infrastructure/Conversion/OoxmlUnits.cs`. Jedno źródło prawdy; nazwane
stałe (`TwipsPerInch=1440`, `EmuPerInch=914400`, `EmuPerPixel=9525`, `TwipsPerPoint=20`,
`HalfPointsPerPoint=2`, `DefaultDpi=96`, `CmPerInch=2.54`). Oba konwertery używają go
zamiast rozproszonych stałych (reguła 14). Testy: `Conversion/OoxmlUnitsTests.cs`
(1 cal = 1440 tw = 914400 EMU = 96 px = 72 pt = 2.54 cm). Patrz ADR-0008.

## 4. Macierz statusów (zweryfikowane w kodzie)

Statusy: `Implemented` / `Partially implemented` / `Planned / not implemented yet`.
Kierunek: R = DOCX→HTML (read), W = HTML→DOCX (write).

| Obszar | Status | Kier. | Gdzie w kodzie | Uwagi / ograniczenia |
|---|---|---|---|---|
| Jednostki (twips/EMU/half-pt/cm/px) | Implemented | R+W | `OoxmlUnits` | DPI 96; testy referencyjne |
| Marginesy strony | Implemented | R+W | `ExtractPageMargins`; `AddPageSettings` | cm, round-trip |
| Rozmiar strony + orientacja | Implemented | R+W | `ExtractPageSize`; `BuildPageSize` | full-stack (Etap 4); fallback A4 |
| Sekcje — wiele sekcji | Partially implemented | R+W | `…FirstOrDefault()` | tylko pierwsza sekcja (R-10) |
| Nagłówek/stopka default | Implemented | R+W | `ExtractHeader/Footer`; `AddHeaderAndFooter` | wybór wg `sectPr` referencji |
| First-page / even-odd header/footer | Implemented | R+W | `HasTitlePage`/`HasEvenAndOddHeaders` | w obrębie 1 sekcji; edycja even-page z UI niedostępna |
| Styl akapitu (`basedOn` chain) | Implemented | R | `ConvertStyleToCssWithInheritance` | sklejanie stringów CSS + `DeduplicateCss` (regex) |
| Styl znakowy (`w:rStyle`) | Implemented | R | `ConvertRunToHtml` (`_styles[rStyleId]`) | Etap 3 slice; direct wygrywa |
| docDefaults | Partially implemented | R | `LoadDocDefaults` | tylko domyślny font + rozmiar; nie pełny rPr/pPr folding |
| Linked styles | Partially implemented | R | `_styles` per styl | brak jawnego łączenia paragraph↔character |
| Numbering (listy) | Partially implemented | R | `LoadNumberingPictureBullets`, list rendering | listy renderowane; numbering **styles** nie foldowane w computed-style |
| Theme fonts / colors | Implemented | R | `GetFontName`, `ResolveThemeColor` | asciiTheme/minor/major; auto color |
| Computed-style (jeden obiekt) | Planned / not implemented yet | R | — | brak `ComputedParagraph/RunStyle`; kaskada CSS |
| Typografia bezpośrednia | Implemented | R+W | `GetRunStyleClean`/`ConvertRunPropertiesToCss` | bold/italic/underline/strike/color/size/highlight/sub-sup/letter-spacing |
| Akapit: align/indent/spacing/line | Implemented | R+W | `ConvertParagraphPropertiesToCss`; writer pPr | exact/atLeast/auto line spacing |
| Tabela: szerokości kolumn | Implemented | R | `ReadTableGridColumnsPx`/`BuildColgroupHtml` | `<colgroup>` + `table-layout:fixed` (Etap 5) |
| Tabela: borders / shading | Implemented | R | `GetTableCellStyleDetailed`, `GetCellBorderCss` | `border-collapse:collapse` |
| Tabela: `tcMar` / `tblCellMar` | Implemented | R+W | `GetTableCellStyleDetailed`, `TableCellMarginDefault`; writer `tblCellMar` 0/108 | domyślny padding = Word default (top/bottom **0**, left/right 108 tw); wcześniej `4px 8px` pompowało wiersze |
| Tabela: `tblCellSpacing` | Planned / not implemented yet | R | — | nieczytane; `border-collapse:collapse` i tak wyklucza spacing |
| Tabela: `gridSpan`/`vMerge` | Implemented | R | `AppendTableCellHtml`, `CountRowSpan` | fix dopasowania wg kolumny gridu (Etap 5) |
| Tabela: szerokość (zapis) | Implemented | W | writer table-width path | brak/auto width → `tblW auto` (nie `pct 5000`); `%`→pct, `px`→dxa |
| Tabela (zapis) — pozostałe | Partially implemented | W | writer `ConvertHtmlToBody` table path | zawsze `TableLayout Autofit`; grid odtwarzany z liczby komórek (nie z colgroup); split-table → 2 `<table>` (R-17) |
| Obraz: rozmiar/EMU/proporcje | Implemented | R+W | `ConvertDrawingToHtml`; writer image path | `data-*-emu` round-trip |
| Obraz: inline + floating (`wp:anchor`) | Implemented | R+W | anchor read/write | front/behind + offsety |
| Obraz: border (`a:ln`) | Implemented | R+W | `data-border-*` ↔ `a:ln` | solid/dashed/dotted |
| Obraz: crop (`a:srcRect`) | Partially implemented | R+W | `data-crop-*` ↔ `a:srcRect` | display `clip-path: inset()` — **nie skaluje** do boxa jak Word |
| Obraz: alt text | Implemented | R+W | `wp:docPr/@descr` ↔ `<img alt>` | `DeEntitize` po stronie zapisu (Etap 6) |
| Obraz: rotacja (`a:xfrm/@rot`) | Planned / not implemented yet | R+W | — | brak odczytu/zapisu; bounding box bez obrotu |
| Obraz: VML / legacy shapes | Partially implemented | R | `ConvertPictureToHtml` (VML ImageData) | tylko obraz + rozmiar; brak border/crop/alt/rot dla VML |
| Tab-stopy: render L⇥C⇥R | Implemented | R | `ParagraphHasAlignmentTab`, `_flexTabs` | akapit z center/right tab → `display:flex` (Etap 7) |
| Tab-stopy: zapis `w:tabs` | Partially implemented | W | `AddDocumentStyles` (style Header/Footer) | center/right tab-stopy emitowane **tylko w stylach Header/Footer**; per-akapit/własne pozycje nie zachowywane |
| Hyperlinki | Implemented | R | `ConvertHyperlinkToHtml` | |
| Pola: data | Partially implemented | R | `ConvertSimpleFieldToHtml`, complex fields | `DateTime.Now` (dynamiczne, nie wartość z DOCX) |
| Pola: numery stron | Partially implemented | R+W | field handling; front „Numery stron" | placeholder/edycja; brak realnej paginacji |

## 5. Roadmapa konwersji (wymagane obszary)

### R-10 — wiele sekcji z różnymi nagłówkami/stopkami

- **Status:** `Partially implemented`. Pojedyncza sekcja w pełni (z first/even); wiele sekcji
  **nie** jest reprezentowanych w `DocumentContent`.
- **Kod:** reader `ExtractHeader/ExtractFooter/ExtractPageMargins/ExtractPageSize` i writer
  `AddPageSettings/GetOrCreateSectionProps` używają `Body.Elements<SectionProperties>().FirstOrDefault()`.
- **Ograniczenie:** dokument z wieloma `sectPr` (różne marginesy/rozmiary/orientacja/nagłówki
  per sekcja) pokaże tylko geometrię i nagłówki/stopki pierwszej sekcji.
- **Kierunek docelowy (Planned / not implemented yet):** rozszerzyć `DocumentContent` o
  kolekcję sekcji (`SectionProperties` + `PageSettings` + referencje header/footer
  default/first/even per sekcja + powiązanie bloków treści z sekcją). Rozszerzyć
  `SectionPropertiesReader` na wszystkie `sectPr`. Dotyczy **R i W**.
- **Testy zabezpieczające:** golden DOCX z 2+ sekcjami (różne marginesy/orientacja, różne
  nagłówki) + asercje per sekcja; round-trip wielu `sectPr`.

### Etap 3 — pełny computed-style

- **Status:** `Partially implemented`. Działa: `basedOn` (akapit), `w:rStyle` (znakowy),
  theme fonts/colors, docDefaults (font/rozmiar). Brak: jeden centralny resolver zwracający
  `ComputedParagraphStyle`/`ComputedRunStyle`; folding pełnych docDefaults do każdego
  elementu; numbering styles; linked styles jako pary.
- **Kod:** `ConvertStyleToCssWithInheritance` (+`DeduplicateCss` regex), `LoadDocDefaults`,
  `ResolveThemeColor`, `GetFontName`, `ConvertRunToHtml` (rStyle). Brak typu computed-style.
- **Ograniczenie:** rozwiązywanie stylów to **sklejanie stringów CSS + dedup regexem** i
  poleganie na kaskadzie CSS — kruche dla złożonych łańcuchów i numbering/theme.
- **Kierunek docelowy (Planned / not implemented yet):** centralny `StyleResolver`
  (docDefaults → styl akapitu `basedOn` → linked character style → numbering style →
  direct formatting → theme/auto color) zwracający jawny computed-style; renderer emituje CSS
  z computed-style (eliminacja regex-dedup). Dotyczy głównie **R**.
- **Testy zabezpieczające:** dziedziczenie (`basedOn` łańcuch), direct nadpisujący bazę,
  theme color/font, auto color, dokument z numbering + theme (regresja).

### Domknięcia (technical debt)

| # | Temat | Status | Kier. | Kod | Ograniczenie / next step |
|---|---|---|---|---|---|
| 6 | Rotacja obrazów (`a:xfrm/@rot`) | Planned / not implemented yet | R+W | — | brak odczytu/zapisu rotacji; next: czytać `rot` (1/60000°) → `transform:rotate()` + `data-rot`, writer odwrotnie; uwzględnić bounding box |
| 6 | VML / legacy obrazy | Partially implemented | R | `ConvertPictureToHtml` | tylko obraz+rozmiar; brak border/crop/alt/rot; obrazy VML w nagłówkach/stopkach jak zwykłe |
| 5 | `tcMar` (marginesy komórek) | Implemented | R | `GetTableCellStyleDetailed` | mapowane na `padding`; OK |
| 5 | `tblCellSpacing` | Planned / not implemented yet | R | — | nieczytane; wymaga `border-collapse:separate; border-spacing` (kolizja z obecnym `collapse`) |
| 7 | `w:tabs` na zapisie | Partially implemented | W | `AddDocumentStyles` | tab-stopy center/right tylko w **stylach Header/Footer** (pozycje 4536/9072 tw); per-akapit/custom nieprzenoszone; znak taba zachowany jako tekst |

## 6. Harness regresji i testy

Lokalizacja: `D2ViewerEditor.Infrastructure.UnitTests/`.

- **Golden/** — deterministyczne DOCX in-memory (`GoldenDocuments`) + snapshoty HTML
  (`HtmlSnapshot`, approve-on-missing, normalizacja base64/`data-image-id`/dat) +
  `GoldenSnapshotTests` z asercjami punktowymi (jednostki, geometria, tabele, tab-stopy,
  styl znakowy). Patrz ADR-0008.
- **Conversion/OoxmlUnitsTests** — wartości referencyjne jednostek.
- **DocxModel/SectionPropertiesReaderTests** — model sekcji/strony.
- **Services/** round-tripy: `PageSizeRoundTripTests`, `ImageAltTextRoundTripTests`,
  `ImageBorderCropRoundTripTests`, `ImageFloatingRoundTripTests`, `HeaderFooterRoundTripTests`,
  `DocxToHtmlConverterFidelityTests`, `DocxToHtmlConverterHeaderFooterTests`,
  `DocxToHtmlConverterSectionReferenceTests`.

Uruchomienie:
```bash
dotnet test D2ApiViewerEditor/D2ViewerEditor.sln
```
Stan na 2026-06-03: backend **358 pass / 0 fail / 6 skip**, GUI **166 pass**.

> Uwaga: snapshoty Golden lokują **obecne** wyjście HTML (regression-guard), nie poprawność
> pikselową wobec Worda (R-12). Po świadomej poprawie wierności baseline aktualizować
> (`delete *.approved.html` + re-run) i przeglądać diff.

## 6a. Round-trip layout — regresja (orginał_GOOD vs zapisany_BAD)

Analiza dwóch realnych plików (kontrakt sprzedaży, 4 strony) ujawniła rozjazd round-tripu
(4→7 stron, powiększone tabele). Naprawione **źródłowo** (2026-06-03):

| Root cause | Fix | Plik |
|---|---|---|
| Marginesy napompowane `Math.Max(top, headerH+720)` + hardkod Header/Footer=720 | marginesy jak autorskie; `Header/Footer = clamp(margin−band, 0, 720)` | `HtmlToDocxConverter.AddPageSettings` |
| Każda komórka +4px góra/dół (`tcMar` 60 tw) | domyślny padding = Word default 0/108 | `DocxToHtmlConverter.ConvertTableToHtml` |
| Tabele rozciągane do 100% (`tblW pct 5000`) | brak/auto width → `tblW auto` | `HtmlToDocxConverter` (table-width) |
| Tabelowy `tblCellMar` 40/80 | Word default 0/108 | `HtmlToDocxConverter` (TableCellMarginDefault) |
| +2 twarde page-breaki z auto-paginacji edytora | `getContent()` scala strony bez `page-break` div | `wysiwyg-editor.ts` |

Weryfikacja (reader→writer na realnym pliku): top 567 / bottom 850 / header 6 / footer 339 /
`tblW auto` / `tcMar` góra+dół 0 / 0 twardych breaków — **AFTER ≈ GOOD** (BAD miało
1281/1231/720/720, pct 5000, 8160 tw, 3 breaki).

Testy regresji: `Infrastructure.UnitTests/Services/RoundTripLayoutFidelityTests` (marginesy
nie-inflowane, header/footer distance, `tcMar`=0, `tblW` auto na syntetycznych DOCX
odtwarzających strukturę GOOD — bez commitowania realnego pliku) + Vitest
`wysiwyg-editor.spec` (getContent bez `page-break`).

Pozostałe ograniczenia round-tripu: R-15 (1 oryginalny break nieodtworzony), R-16 (model
stratny: `styles.xml`/`numbering.xml`/theme/font Cambria→Calibri), R-17 (split-table → 2
tabele), R-18 (zapiekane `trHeight`).

## 6b. Pass-through pakietu + ręczna weryfikacja (R-15/R-16/R-17)

**Pass-through (R-16, częściowy):** `HtmlToDocxConverter.ConvertPreservingPackage(html, Stream? original, …)`
generuje DOCX jak `Convert`, a następnie zachowuje z oryginalnego pakietu `styles.xml` (pełny
zestaw, w tym style tabel), `theme` i `fontTable` (FeedData). Body/sekcja/nagłówki-stopki/
obrazy/numbering pochodzą z konwersji HTML. Wpięte w `DownloadEditedDocumentCommandHandler`
(ładuje bazową wersję v1 przez `IDocumentStorageService.DownloadAsync`; fallback do `Convert`
gdy brak wersji / nie-DOCX). Status szczegółowy: R-19/R-20/R-21 w `RISKS_ASSUMPTIONS.md`.

| Temat | Status | Plik |
|---|---|---|
| Page break round-trip | `Implemented` | `HtmlToDocxConverter.IsPageBreakNode`/`CreateRunsFromNode` |
| Pass-through styles/theme/fontTable | `Partially implemented` | `HtmlToDocxConverter.ConvertPreservingPackage` (Download path) |
| Pass-through numbering.xml | `Planned / not implemented yet` | — (R-19, konflikt numId) |
| Pass-through w autosave | `Planned / not implemented yet` | — (R-20, ścieżka bezstanowa) |
| Scalanie split-table na zapisie | `Implemented` | `wysiwyg-editor._mergeSplitTables` |

**Generacja `zapisany_AFTER.docx`:** lokalnie, reader → `ConvertPreservingPackage(html, oryginał)`
(test był tymczasowy, niecommitowany). Wynik (reader→writer): styles **164** (=GOOD), theme
**Cambria**, page-breaki **1** (=GOOD).

**Checklista ręcznej weryfikacji w MS Word:**
1. Otwórz `orginał_GOOD.docx` — potwierdź ~4 strony (wzorzec).
2. Otwórz `zapisany_AFTER.docx` — sprawdź:
   - liczba stron ≈ 4 (nie 7),
   - tabele nie są rozciągnięte na 100% / nie urosły w pionie,
   - szerokości kolumn ≈ oryginał,
   - marginesy strony ≈ oryginał (góra wąska, nie 2,3 cm),
   - **font treści szeryfowy (Cambria)**, nie Calibri,
   - manualny page break w tym samym miejscu co w oryginale,
   - nagłówek/stopka na miejscu, nie wypchnięte.
3. (Kontrast) `zapisany_BAD.docx` — 7 stron, Calibri, napompowane tabele.
4. Zapisz odchylenia w raporcie (co nadal różni się od oryginału).

## 7. Ograniczenia trwałe (HTML/CSS)

Nieosiągalne 1:1 w przeglądarce (udokumentowane w `FIDELITY_REPORT.md` §3):
zawijanie tekstu wokół obrazu „przylegle/na wskroś" (tight/through), dokładna paginacja i
podział stron Worda, kerning/metryki czcionek, dynamiczne pola zależne od layoutu
(`PAGE`/`NUMPAGES`). Cel: maksymalna **mierzalna** wierność, nie 100%.
