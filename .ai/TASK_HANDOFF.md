# Task Handoff

> Bezpieczne przekazanie pracy kolejnej sesji/agentowi.

## Ostatnia aktualizacja (2026-07-09)

- **Edytor: ENTER na dole strony nie rozciąga już kartki i nie duplikuje treści (ADR-0032).**
   - Dwie przyczyny (zdiagnozowane w REALNEJ przeglądarce przez headless Chrome + CDP — jsdom/stubowane testy tego nie łapały): (1) desync — Angular nie nadpisuje `[innerHTML]` strony o niezmienionej wartości SafeHtml (cache w `getPageContentSafe`), więc strona z żywymi edycjami użytkownika nie jest resetowana do rozkładu paginacji (rośnie w pion + duplikuje nadmiar na kolejne strony i do `getContent`/zapisu); (2) CSS — `.page{min-height}` + `.editor-content{flex:1}` pozwalały flex-kontenerowi rosnąć do max-content, więc `overflow:hidden` nie przycinał.
   - Fix: `_syncPageEditorDom(newPageContents)` po repaginacji wymusza `innerHTML` edytorów = obliczony rozkład (tylko rozjechane strony, guard `!==`, re-wrap obrazów); `.page` ma definitywną `height` (`pageHeightPx`, było `pageMinHeightPx`/`min-height`).
   - Zoom był OK od początku (czysto wizualny `transform:scale` na `.editor-wrapper`) — rozmiar strony/model niezależne od zoomu jak w Wordzie.
   - Weryfikacja: CDP — 1122 px (A4) przez cały burst 45×ENTER, po ustabilizowaniu 3 strony A4, `getContent` 34 unikalne akapity bez duplikatów/utraty, pełna strona 50 akapitów bez przycięcia; `pagination-overflow.spec` 16/16 (+2), `wysiwyg-editor` 89/89, `ng build` 0 błędów.
   - **Do rozważenia dalej:** repaginacja nadal opiera się na klonowaniu bloków + redystrybucji `innerHTML` (kruche przy jednoczesnym pisaniu i re-renderze) — docelowo warto oddzielić warstwę layoutu/paginacji od contenteditable; pojedynczy blok wyższy niż strona jest teraz przycinany (świadome ograniczenie).
   - **Narzędzie (scratchpad, nietrwałe):** driver CDP `cdp-drive.mjs`/`probe*.mjs` odtwarza scenariusz ENTER na realnym layoucie (wzorzec ze skilla `verify`: auth off w `config.json`, `ng serve --port 4299`, headless Chrome `--remote-debugging-port`). Przydatny do regresji rozciągania/duplikacji, których jsdom nie wykryje.

## Ostatnia aktualizacja (2026-07-08)

- **Znaki `w:sym` (Symbol/Wingdings) widoczne i round-tripujące.**
   - Objaw: strzałki i inne znaki „Wstaw → Symbol" niewidoczne w edytorze i gubione przy zapisie. Root cause: reader (`DocxToHtmlConverter.ConvertRunChildToHtml`, gałąź `SymbolChar`) emitował surowy PUA `&#xF0E0;` (tofu bez fontu symbolicznego) i IGNOROWAŁ `sym.Font`.
   - Fix: `ConvertSymbolCharToHtml` + `MapSymbolCharToUnicode` — znane glify (strzałki/operatory Symbol, checkboxy/haczyk Wingdings) → prawdziwy Unicode (tekst, widoczny wszędzie, round-trip jako `w:t`); nieznane → znak (niski bajt) w spanie z ORYGINALNYM fontem glifowym (`font-family`), który writer odtwarza jako `w:rFonts` → glif przeżywa zapis i jest poprawny w Wordzie. Zwykłe strzałki Unicode w `w:t` działały już wcześniej.
   - Testy: `SymbolCharImportTests` 6/6, Infrastructure **376/376**, golden bez zmian.
   - **Do rozważenia dalej:** rozszerzyć `MapSymbolCharToUnicode` o więcej pewnych glifów (np. Symbol → greka, dokładne strzałki Wingdings) i/lub pełny round-trip `w:sym`→`w:sym` (dziś mapowane wracają jako tekst, nieznane jako `w:t`+font — bez utraty danych, ale bez tożsamości elementu).

- **Zapis przez pass-through (domknięcie ADR-0031).**
   - `POST /document/save` używa `ConvertPreservingPackage`, gdy request niesie `masterId` (opcjonalne pole na `SaveDocumentCommand`/`SaveDocumentRequest` + TS). Handler ma teraz opcjonalne `IDocumentRepository`/`IDocumentStorageService`/`ILogger` (nullable): jest masterId + DOCX + wersja bazowa → wczytuje oryginał z GCS i zachowuje `styles.xml`/`theme`/`fontTable`/`numbering`; brak / błąd / uszkodzony oryginał → **best-effort fallback** do `Convert` (zapis nigdy się nie wywala).
   - GUI: `buildSaveRequest()` dokłada `masterId: documentMasterId() ?? undefined` — autosave/Zapisz/Pobierz/Zakończ dla ISTNIEJĄCEGO dokumentu idą pass-through; nowy dokument (brak masterId) → regeneracja.
   - Dowód: harness `passthrough` na qutable — regeneracja 14 stylów / brak definicji `Tabela-Siatka`; pass-through 38 stylów / definicja obecna + theme + fontTable. Testy: `SaveDocumentCommandHandlerTests` 13/13 (+5), Application 311/311, Api 130/130; GUI `document-editor.spec` +2 (371/372, fail = pre-existing layout-shell); build sln + `ng build` OK.
   - **Do rozważenia dalej:** Sign/`FinishAndSend` nadal regeneruje pakiet (nie dostał masterId) — jeśli PODPISYWANE dokumenty tracą style tabel, dołożyć tam analogiczny pass-through (`SignDocumentCommand` ma dostęp do dokumentu). `numbering.xml` oryginału jest przenoszony, ale listy tworzone w edytorze nadal generują własną numerację (R-19 — pełne remapowanie numId poza zakresem).

- **Domyślne odstępy/interlinia dokumentu + tabele nie psują się po zapisie (ADR-0031).**
   - Kontrakt: reader emituje na `.document-content` font + `data-default-before/after-tw`/`data-default-line`/`data-default-line-rule` + `line-height`; writer (`CaptureDocumentDefaults`) odtwarza z nich docDefaults/Normal regenerowanego pakietu; GUI (`_captureDocumentDefaults`/`_wrapWithDocumentContainer`) rozwija wrapper przy imporcie i owija treść z powrotem w `getContent()` — bez tego writer w produkcyjnym zapisie nie widzi domyślnych wartości dokumentu.
   - Akapity w KOMÓRKACH tabel niosą inline rozwiązany spacing (docDefaults + `w:pPr` łańcucha stylu tabeli, `TableStyleContext.ParagraphDefaultCss`) — eksport odporny na brak definicji stylu tabeli w regenerowanym pakiecie (wiersze nie puchną w Wordzie). Akapity body inline'u NIE dostają (interlinia z kontenera).
   - Tabele: `<col data-w-tw>` = dokładne twips tblGrid (writer preferuje nad px dla gridCol/tcW; suma spanowanych kolumn dla scaleń; % nietykane; `syncTableColgroup` usuwa atrybut przy zmianie szerokości); tabela z `data-tbl-style` bez CSS `border:` → writer NIE emituje tblBorders (val=none nadpisywało styl); `data-no-borders="1"` → jawne none jak dotąd.
   - Schemat OOXML: `NormalizeParagraphPropertiesOrder` (po `ApplyParagraphStyle`/`Extras` i po dołożeniu pStyle/tabs w `ConvertParagraphElement`), `NormalizeTableCellPropertiesOrder` (przed `cell.Append(cellProps)`), tcBorders top→left→bottom→right. Walidator (Office2013) na eksporcie qutable: 0 błędów.
   - Testy: `DocumentDefaultsAndTableSpacingFidelityTests` 11/11 (Infrastructure 370/370; goldeny simple/merged-cells-table zregenerowane — diff tylko `data-w-tw`); GUI +4 specy (369/370, fail = pre-existing layout-shell); `ng build` + build sln OK.
   - **ZROBIONE (patrz wpis pass-through wyżej):** Save/autosave przełączone na `ConvertPreservingPackage` przez opcjonalny `masterId` w `/save` — definicje stylów tabel/motywu wracają dla istniejących dokumentów. Pozostałe drobne: `tcMar` dryf 108→105 tw (kosmetyka), wrapper `<span>` wokół `&nbsp;` pustych akapitów po 1. zapisie (idempotentne) — oba nieszkodliwe i dotyczą już tylko ścieżki regeneracji (nowy dokument).

## Ostatnia aktualizacja (2026-07-07)

- **Kotwiczenie elementów pływających jak w Wordzie — pełny round-trip textboxów + znacznik kotwicy (ADR-0030).**
   - Model kotwicy = pozycja w DOM (bez sztucznych ID): obraz pływający WEWNĄTRZ akapitu-kotwicy; `div.docx-textbox` hoistowany przez reader bezpośrednio PRZED akapit-kotwicę (div w `<p>` był re-parentowany przez parser przeglądarki i rozcinał akapit); writer przypina drawing textboxa do NASTĘPNEGO akapitu (`BufferTextBoxDrawing`/`AttachPendingTextBoxes` w `HtmlToDocxConverter`; flush awaryjny na końcu body/header/footer; `<li>` i zagnieżdżenia obsłużone). Round-trip idempotentny — test 2 cykli bez wzrostu liczby akapitów.
   - Writer wreszcie odtwarza pola tekstowe: `BuildTextBoxDrawing` → `wps:wsp`+`w:txbxContent` (wcześniej generyczny div = spłaszczenie i utrata ramki przy 1. autosave). Kontrakt data-*: `data-pos-mode/x-emu/y-emu/width-emu/height-emu/border-*/wrap` (osie ADR-0029). `data-wrap` przywraca `wrapSquare`/`wrapTopBottom` także dla obrazów (koniec degradacji do WrapNone).
   - Obramowanie textboxa rozdzielone: DOKUMENTOWE (`a:ln` kształtu) = inline `border` + `data-border-*` (round-trip); EDYCYJNE = outline w SCSS (hover/focus-within dashed, tb-selected/tb-dragging solid) — stan nieaktywny bez ramki, zero wpływu na box-model.
   - GUI: znacznik kotwicy (`.anchor-badge`, SVG anchor, `aria-label`, pointer-events:none) overlay w `.page` poza contenteditable; pokazywany dla zaznaczonego elementu PŁYWAJĄCEGO przy jego akapicie-kotwicy, repozycjonowany rAF-em po edycji/repaginacji, znika po odznaczeniu/usunięciu; zoom/scroll bez przeliczeń. Czyste funkcje w `core/utils/floating-anchor.util.ts`. Drag textboxa z pasa 8px krawędzi (wnętrze = edycja tekstu), delty dzielone przez skalę zoomu (ten sam fix nałożony na drag pływających obrazów — wcześniej uciekały kursorowi przy zoomie ≠ 100%). Kotwica przy dragu BEZ zmian (tylko offsety — reguła przewidywalna).
   - Testy: `TextBoxAnchorRoundTripTests` 12/12 (Infrastructure 359/359, sln build 0 błędów); GUI `floating-anchor.util.spec` 15 + `wysiwyg-editor.anchor.spec` 9 (pełne 365/366, fail tylko pre-existing `layout-shell`; `ng build` OK).
   - **Do rozważenia dalej:** `mc:AlternateContent` z fallbackiem VML dla wps:wsp (Word <2010); rendering wrap=square jako realne obtekanie w edytorze (dziś pozycja absolutna); re-anchoring przy przeciągnięciu na inny akapit/stronę (dziś kotwica stała); tło kształtu textboxa (`a:solidFill` spPr) w round-tripie.

- **Stopka/nagłówek: tekst obok numeru strony nie znika — SdtRun w maszynie pól złożonych.**
   - Pierwsza utrata była w readerze: `ConvertComplexFieldParagraphContent` pomijał `SdtRun` (formanty inline). Objaw „widać tylko numer strony" = tytuł/klauzula w content controls obok pola PAGE; galeria Worda „Strona X z Y" (pole WEWNĄTRZ SdtRun) znikała w całości.
   - Reader: `ComplexFieldState` + `AppendComplexFieldContent`/`AppendComplexFieldRun`/`AppendComplexFieldSdtRun` — wspólny stan pola przez granice formantów, wrapper `span.sdt-inline` jak w `ConvertSdtRunToHtml`, kolejność zachowana. Writer: `BuildSdtRunFromHtml` odtwarza pola PAGE/NUMPAGES z field-spanów (wcześniej literalny „{page}" po 1. autosave).
   - Testy: `HeaderFooterFieldNeighborContentTests` 6/6 (w tym pełny round-trip), Infrastructure 347/347, GUI `wysiwyg-editor.spec` 32/32 (+1). Front bez zmian kodu.
   - **Znane ograniczenie (świadome):** akapit z polem złożonym nie używa POZYCYJNYCH tabów (`usePositionedTabs` wyklucza `hasComplexField` — runy pola są stanowe); treść kompletna, układ L/C/R przybliżony flexem/inline. Ewentualne odwzorowanie pozycyjne = osobne zadanie (wymaga segmentacji świadomej stanu pola).

- **Kotwiczone obiekty (wp:anchor) — pozycja jak w Wordzie (ADR-0029).**
   - Reader rozwiązuje kotwicę do współrzędnych edytora przez `ResolveAnchorPosition`/`ResolveAxis`: X = lewa krawędź strony (page-relative), Y = góra obszaru treści (page-relative − górny margines). Uwzględnia `relativeFrom` (page/margin/column/leftMargin/rightMargin/insideMargin/outsideMargin/…) i `wp:align` (right/center/left/inside/outside; rozmiar z `wp:extent`). Geometria 1. sekcji w polach instancji (`LoadPageGeometry`, brak `w:pgMar` → 1440 twips). Wpięte w `ConvertDrawingToHtml` (obrazy → `data-x-emu`/`data-y-emu`) i `BuildTextBoxLayoutCss`/`RenderVectorShapeAsHtml` (pola tekstowe/kształty → inline `position:absolute`).
   - Writer: pionowa kotwica `RelativeFrom=Margin` (poziomo `Page`) → round-trip idempotentny (reader dodaje i odejmuje ten sam górny margines) i poprawny render w Wordzie. Front NIE zmieniany (restore czyta gotowe współrzędne).
   - Testy: nowy `DocxAnchorPositionFidelityTests` 8/8; zaktualizowane `Doc2ImportFidelityTests.AnchoredTextBox…` i `ImageFloatingRoundTripTests`; Infrastructure 339/339, build solucji 0 błędów.
   - **Do rozważenia dalej:** dokładniejszy Y (pomiar realnego pasma nagłówka zamiast założenia „≈ górny margines"), `wp:align` pionowy center/bottom, rozróżnienie inside/outside dla stron parzystych, `wrapSquare`/`wrapTight` (opływanie tekstem) zamiast obecnego front/behind.

- **SVG „puste białe logo" w nagłówku (Doc2/ING) — sanitizer naprawiony u źródła.**
   - `GraphicConversionService.SanitizeSvg`: (1) wewnętrzny `<use href="#id">` ZOSTAJE (wcześniej wycinany bezwarunkowo — logo z `<defs>`+`<use>` = pusty biały obraz o poprawnych wymiarach), zewnętrzny/data:/bez href usuwany (`HasInternalFragmentHref`); (2) UTF-8 BOM trimowany przed parsowaniem (wcześniej cały SVG odrzucany); (3) `SafeParse`: `DtdProcessing.Prohibit`→`Ignore` (DOCTYPE Illustratora przechodzi; XXE nadal null — encje niezdefiniowane rzucają).
   - Testy: `GraphicConversionSecurityTests` +4 i zaktualizowany `SanitizeSvg_KeepsSafeDataHrefAndFragment` (poprzednio pilnował usuwania use), `Doc2ImportFidelityTests` +2 E2E (defs/use w data-URI, BOM). SVG-testy 25/25; Infrastructure 325/328 w czystym worktree (3 faile = WIP hMerge/tab-in-cell z sesji tabelowej, niezwiązane).
   - **Triage gdy logo nadal puste po wdrożeniu:** logi backendu „SVG part pominięty…" / „Media part bez rastra web…"; w DOM edytora `data-legacy-graphic="blank"` na `<img>` ⇒ metafile EMF+ (poza etapem 1 ADR-0027 — patrz kandydaci etapu 2: parsowanie EMF+ z GDICOMMENT). Najlepiej pozyskać źródłowy DOCX i przepuścić przez `tools/docx-diagnostics/inspect-docx`.
   - Ograniczenie bez zmian: writer dropuje SVG na eksporcie HTML→DOCX (utrata na 1. autosave; roadmapa `asvg:svgBlip`/rasteryzacja SVG→PNG).

- **Tabele Doc2/ING — druga runda (screenshot edytor vs Word):**
   - `w:hMerge` (legacy scalanie poziome) obsłużony w readerze: `BuildRowRenderPlan` + `GetHMerge` w `DocxToHtmlConverter` (restart pochłania continue jako colspan; continue pomijane jak kontynuacje vMerge; pominięty val = continue; sierota renderowana normalnie). Fix „deficytu kolumn" z 2026-07-06 zostaje — dotyczył KRÓTKICH wierszy, ten dokument używał hMerge.
   - Taby POZYCYJNE wyłączone wewnątrz komórek tabel (`usePositionedTabs && !paragraph.Ancestors<TableCell>().Any()`) — absolutne segmenty wyjeżdżały poza wąską komórkę (nałożone nagłówki) i wyłączały text-align komórki. W komórce: inline spacer/flex jak dawniej; `data-tab-stops` dalej round-tripuje.
   - Testy: `Doc2ImportFidelityTests` 24/24 (+4). Infrastructure 321/322.
   - **Wyjaśnione:** fail `SanitizeSvg_KeepsSafeDataHrefAndFragment` pochodził z RÓWNOLEGŁEJ sesji SVG w tym samym drzewie (produkcja już zachowywała `<use>`, test chwilowo stary) — domknięte wpisem „SVG puste białe logo" wyżej, test zaktualizowany, komplet zielony.
   - Jeśli tabela ING nadal odbiega: kandydaci = warunkowe formatowanie TEKSTU ze stylu tabeli (jc/bold nagłówka ze stylu — znane ograniczenie ADR-0026) oraz reguła „tab skacze do NASTĘPNEGO stopu za bieżącą pozycją x" (k-ty tab → k-ty stop).

- **Własny tłumacz wektorowy EMF/WMF → SVG, etap 1 (ADR-0027).**
   - Nowy `MetafileVectorTranslator` (Infrastructure, internal, pure-managed — zero zależności) tłumaczy podzbiór rekordów GDI na SVG; wpięty jako strategia `vector-translate` w `GraphicConversionService.ConvertMetafile` po `dib-rasterize`, przed blankiem. SVG tylko podgląd — eksport zawsze niesie oryginalny metafile (data-original-src / pass-through), `LoadImageFromPart` nie podmienia bajtów na SVG.
   - Pokrycie etapu 1: pióra/pędzle/stock/dash, linie, rect/ellipse/roundrect, poly*/polypolygon (16/32-bit, fill-rule), ścieżki, world transform, StretchDIBits→`<image>`; WMF: slotowa tabela obiektów + odwrócone parametry. Limity anty-DoS; wyjątek → blank.
   - Testy: `MetafileVectorTranslationTests` 10/10; Infrastructure 285/285 (stare blank-testy bez zmian); Application 306/306.
   - **Etap 2 (kandydaci):** ExtTextOut (tekst z przybliżeniem metryk), clipping (SelectClipPath → clipPath), Arc/Pie/Chord, PatternBrush, parsowanie rekordów EMF+ z GDICOMMENT; ewentualnie rasteryzacja SVG→PNG dla spójności miniatur. Rozważyć realne pliki EMF od użytkowników jako fixtures.

- **Import obrazów z DOCX — naprawa 6 potwierdzonych bugów w istniejącym mechanizmie (bez nowych zależności, bez nowego systemu konwersji).**
   - Reader (`DocxToHtmlConverter`): klucz obrazów per część pakietu (`ImageCacheKey(part, rId)` — kolizje rId main vs header/footer podmieniały obrazy), nowa gałąź `ConvertAlternateContentToHtml` (mc:Choice→mc:Fallback; wcześniej całe `mc:AlternateContent` szło w pusty string), zero-extent → wymiary intrinsic (koniec `width:0px`), `WebGraphicForLegacy` obejmuje `Unknown` (EMZ/WMZ/nieznane nie trafiają do `src` jako nierenderowalny data URL), `r:link`-only pomijany z logiem, opcjonalny `ILogger` (diagnostyka etapu: część/relId/typ/rozmiar/status/powód).
   - `GraphicConversionService`: dekompresja GZIP (EMZ/WMZ) przed detekcją, bounded do `MaxInputBytes`.
   - Writer (`HtmlToDocxConverter.BuildImageDrawing`): prawdziwy content type partu (Tiff/Icon mapowane; nieznane `image/*` przez `AddImagePart(contentType)`; wcześniej wszystko nieznane = „Jpeg" → obrazy psuły się po pierwszym autosave).
   - Testy: nowy `ImageImportRegressionTests` 10/10; Infrastructure 275/275, Application 306/306, build solucji 0 błędów, golden nietknięte.
   - **Do zrobienia w następnych sesjach (świadomie poza zakresem):** rendering obrazów w textboxach/grupach wieloobrazowych (obecnie pierwszy blip), SVG-blip extension (`asvg:svgBlip` — dziś renderowany PNG-fallback Worda), potwierdzenie testów w kontenerze Linux w CI (zmiany pure-managed, ryzyko niskie), ewentualny placeholder wizualny dla `r:link`.

- **Pokrycie testami zmian ADR-0025/0026 (+16 testów, tylko testy — zero zmian w kodzie produkcyjnym).**
   - Application (+4): `SaveDocumentCommandHandlerTests` i `DownloadEditedDocumentCommandHandlerTests` pilnują, że `SectionHeadersFooters`/`PageSize` dochodzą do `IHtmlToDocxConverter` nietknięte (obie ścieżki: `Convert` i `ConvertPreservingPackage`); regresją byłoby ciche spłaszczenie nagłówków sekcyjnych przy autosave.
   - Infrastructure (+12): `TabStopFidelityTests` (leader `dot`, `w:val=clear` usuwa stop ze STYLU akapitowego, direct nadpisuje pozycję stylu, bar-tab pomijany), `TableStyleFidelityTests` (lastRow, firstColumn, legacy maska hex `tblLook@w:val`, `noHBand` wyłącza pasy, pasy pionowe band1Vert), `MultiSectionFidelityTests` writer (wpis footer-only na sectPr swojej sekcji, wpis z indeksem poza zakresem/0 ignorowany bez wyjątku — scenariusz skasowanego markera sekcji).
   - Weryfikacja: Application.UnitTests 306/306, Infrastructure.UnitTests 265/265.

- **Tabele: rozwiązywanie stylu tabeli + wierność Word (ADR-0026).**
   - Reader: `ResolveTableStyleContext`/`ComputeConditionalRegions` (tblStyle→basedOn per-side, tblLook, tblStylePr firstRow/lastRow/kolumny/pasy), pozycyjne krawędzie outer vs insideH/V, theme fill/tint/shade, wzory pctNN, `sz/6` px, `height:`+data-row-* na tr, data-tbl-header/cant-split, tblCellSpacing, `data-tbl-style`/`data-tbl-look`.
   - Writer: `w:tblStyle`+`w:tblLook` z data-*, trHeight z twips+hRule, tblHeader/cantSplit/tblCellSpacing/tblInd, px×6, dziesiętne %; fix duplikacji wierszy zagnieżdżonych tabel.
   - GUI: koniec zapiekania wysokości wierszy (R-18), resize wiersza czyści data-row-*, `syncTableColgroup` (nowy `core/utils/table-grid.util`) po resize i insert/delete kolumny.
   - Weryfikacja: `TableStyleFidelityTests` 21/21, Infrastructure 253/253, solucja 0 fail; GUI 317/318 (pre-existing fail), `ng build` OK; Golden simple-table/merged-cells-table zregenerowane.
   - Ograniczenia/kierunek: warunkowe formatowanie TEKSTU ze stylu (bold nagłówka) nieprzenoszone; tblHeader/cantSplit tylko round-trip (paginacja edytora ich nie egzekwuje); przekątne tl2br/tr2bl brak. Szczegóły: ADR-0026 + macierz w `DOCX_CONVERSION.md`.

- **Edytor: fix paginacji — ENTER przy dolnej krawędzi tworzy nową stronę (kartka nie rośnie).**
   - Root cause (GUI, `wysiwyg-editor.ts`): pomiar bloków w `_repaginateNow` odbywał się w gołym `<div>` poza kontekstem stylów `.editor-content` (pusty `<p>` = 0 px; brak scope'owanych marginesów) i przez `rect.height` bez marginesów → paginacja nie widziała przepełnienia → `.page` (min-height, overflow:visible) rosła w pion.
   - Fix: `_createBlockMeasurer` (measurer z klasą `editor-content`) + `_measureBlockRunHeights` (delty pozycji z sentinelem = marginesy + kolaps; 1 layout-flush na przebieg) + max-wait 600 ms w `_schedulePaginate` (przytrzymany Enter nie głodzi repaginacji). Kursor/listenery nowych stron — istniejące mechanizmy, bez zmian.
   - Weryfikacja: nowy `wysiwyg-editor.pagination-overflow.spec.ts` 9/9; pełne GUI 312/313 (fail tylko pre-existing `spec-layout-shell`); `ng build` OK. Backend nietknięty. Szczegóły: CHANGELOG + `EDITOR_KEYBOARD.md` (sekcja „Pomiar przepełnienia strony").

- **Listy wielopoziomowe: tożsamość logiczna + liczniki Worda + round-trip formatów.**
   - Reader (`DocxToHtmlConverter`): liczniki numeracji per (abstractNumId, ilvl) — wspólny abstrakt kontynuuje, `startOverride` restartuje przy 1. użyciu instancji, poziomy głębsze restartują po powrocie na płytszy (`lvlRestart=0` honorowane), `numStyleLink` rozwiązywany; grupowanie ol/ul po **numId** (koniec sklejania niezależnych list o tym samym wyglądzie); `<ol start>` z liczników (kontynuacja po przerwaniu akapitem); kontener niesie `data-num-id/-abstract-num-id/-ilvl/-num-fmt/-start/-lvl-text/-bullet-font`.
   - Writer (`HtmlToDocxConverter`): `_numIdByHtmlList` — fragmenty z tym samym `data-num-id` współdzielą jedną `NumberingInstance` (Word kontynuuje numerację); `CreateAbstractNumbering` odtwarza `numFmt/lvlText/start/font` z data-* (koniec hardkodowanej drabinki niszczącej numerację przy autosave).
   - Weryfikacja: nowy `ListNumberingFidelityTests` **12/12** (w tym pełny round-trip); pozostałe testy list/golden bez regresji.
   - **UWAGA — stan drzewa:** równolegle trwa refaktor stylów tabel (`ResolveTableStyleContext`/`TableRenderContext` w readerze); snapshoty `SimpleTable`/`MergedCellsTable` failują na JEGO diffach (celowe atrybuty stylów tabel) — baseline'y zregenerować po domknięciu tamtej pracy (`delete *.approved.html` + re-run). Snapshot `tab-stops-lcr` już zregenerowany (niesie `data-tab-stops` z tamtej sesji).
   - Ograniczenia list (HTML): sufiks `lvlText` ("1)") i numeracja multilevel "1.1.2" nierenderowane wizualnie w edytorze (wracają do DOCX bez strat); nowe listy z edytora (bez data-*) → domyślna drabinka formatów.

- **Wierność DOCX↔HTML: wiele sekcji + tabele + interlinia (ADR-0023).**
   - Reader: `PageSize`/`Margins`/nagłówek/stopka z PIERWSZEJ sekcji (`GetSectionPropertiesInDocumentOrder` w `DocxToHtmlConverter`); paragraph-level `pPr/sectPr` → `div.page-break` + niewidoczny `div.docx-section-break` z geometrią następnej sekcji w `data-*`.
   - Writer: markery → paragraph-level `w:sectPr` (body-level = ostatnia sekcja); refs nagłówka/stopki + `titlePg` na pierwszym sectPr; `page-break` przed markerem nie emituje `w:br`; tabele: `colgroup`→`tblGrid`, `table-layout:fixed`→`Fixed`, kontynuacje `vMerge` pod rowspan; interlinia `atLeast` przez `--w-line-rule:atLeast`.
   - GUI: `wysiwyg-editor.scss` ukrywa `.docx-section-break`; test w `wysiwyg-editor.spec` pilnuje przeżywalności markera przez split/getContent.
   - Weryfikacja: Infrastructure 210/210 (MultiSection 12, TableWrite 6, LineSpacing +3), cała solucja 0 fail, GUI 288/289 (fail tylko pre-existing `spec-layout-shell`), `ng build` OK.
   - **(ZROBIONE w tej sesji) rendering per sekcja w edytorze:** `wysiwyg-editor` ma input `pageSize` + `pageGeometries` per strona (sekcja 1 z inputów, kolejne z data-* markera; marker otwierający stronę → geometria od tej strony, continuous → od następnej); szablon binduje wymiary/orientację/paddingi/prowadnice per strona; `_repaginateNow` używa wysokości/szerokości bieżącej sekcji + pomiar wsadowy bloków (1 reflow na przebieg); guard `_flattenTopBlocks` nie wchodzi w markery. Weryfikacja: specy wysiwyg 51/51 (+6), pełne GUI 294/295 (pre-existing fail), `ng build` OK, solucja backend 730/730.
   - **(ZROBIONE 2026-07-05, ADR-0025) nagłówki/stopki per sekcja:** `SectionHeaderFooter` w `DocumentContent.SectionHeadersFooters`/`SaveDocumentRequest` (pole opcjonalne, Domain+TS); reader wpis per sekcja ≥ 1 z własnymi refs; writer kładzie wpis na sectPr swojej sekcji (`_emittedSectionProps`); handlery Save/DownloadEdited; GUI: `pageSectionIndexes`, wybór nagłówka per strona z dziedziczeniem, edycja pasma na klikniętej stronie (`editingHfPageIndex`) z routingiem do właściciela, output `sectionHeadersFootersChange` → save. Testy: MultiSection +4 backend, wysiwyg +6 GUI.
   - **(ZROBIONE 2026-07-05, ADR-0025) dynamiczne pasmo nagłówka/stopki:** pasmo zaczyna się headerDistance od krawędzi (margin-top), min-height = margines − dystans, treść wyższa spycha body; `_repaginateNow` mierzy realne pasma z DOM (`_measureBandHeightsPx`). `PageGeometry` +`headerDistanceCm/footerDistanceCm`.
   - **(ZROBIONE 2026-07-05, ADR-0025) tab-stopy per akapit:** `data-tab-stops` round-trip (reader efektywne w:tabs styl+direct; writer `ParseTabStops`); nagłówek/stopka renderuje segmenty pozycyjnie (`span.docx-tab-seg`); literalny `	` → element `w:tab`.
   - **Następne kroki (opcjonalne):** Sign z sekcyjnymi nagłówkami; leader (kropki) tab-stopów w edytorze; pełna reguła przypisania tabów („następny stop za bieżącym x").
   - Uwaga: marker jest elementem contenteditable — użytkownik może go skasować edycją przy granicy sekcji (degradacja do jednej sekcji, jak przed zmianą).

- corpKey ustalany wyłącznie po stronie API z tokenu — usunięty transport z GUI.
   - Powód: GUI czytało `corpKey` ze statycznego snapshotu `getActiveAccount().idTokenClaims` (często null) → `SaveDocumentVersionCommand.CorporateKey` przychodził null. Token wysyłany do API (ID token) i tak niesie `corpKey`.
   - API: usunięty parametr `CorporateKey` z komend `SaveDocumentVersion`/`UpdateDocumentVersion`/`FinishAndSendDocument` i z controller DTO `SaveDocumentVersionRequest`; handlery czytają `_currentUser.CorporateKey` (brak → `Result.Failure`).
   - GUI: usunięty sygnał `corporateKey` + odczyt claimu w `document-editor.ts`; zapisy wysyłają `{ content }`. Listy admina (`corporateKey` z `DocumentDelivery`) bez zmian.
   - Weryfikacja: backend build 0 błędów; `Application.UnitTests` 295/295; `Api.UnitTests` 113/113; GUI build OK. Decyzja: ADR-0021.

## Ostatnia aktualizacja (2026-06-26)

- Naprawa zapisu tożsamości edytującego (corpKey był NULL) + przycisk „Kopiuj link" w `admin-files`.
   - Przyczyna NULL: (1) Entra — `AzureAdOptions.CorporateKeyClaim` domyślnie `"ck"`, a token niesie claim `corpKey`; (2) dev bypass — `HttpHeaderCurrentUserProvider` czytał nagłówek `X-Corporate-Key`, którego GUI nie wysyła.
   - Kierunek (decyzja użytkownika): **frontend wysyła `corpKey`** (z `idTokenClaims`), backend używa go z fallbackiem do tokenu i twardym błędem, gdy nie da się ustalić użytkownika.
   - API: `AzureAdOptions.CorporateKeyClaim` `"ck"`→`"corpKey"`; `HttpHeaderCurrentUserProvider` fallback `DEV-LOCAL`; handlery `SaveDocumentVersion`/`UpdateDocumentVersion` wstrzykują `ICurrentUserProvider`, wszystkie trzy (`Save`/`Update`/`FinishAndSend`) liczą `corporateKey = request.CorporateKey ?? _currentUser.CorporateKey` i zwracają `Result.Failure`, gdy brak.
   - GUI: bez zmian w wysyłce corpKey (już wysyła). `admin-files` — przycisk „Kopiuj link" (kopiuje link `/editor?masterId&versionId` do schowka, baner `notice`), widoczny dopiero w szczegółach pozycji po rozwinięciu.
   - Weryfikacja: backend build 0 błędów; `Application.UnitTests` 295/295; `Api.UnitTests` 113/113; GUI build OK; `ng test` 268/269 (pre-existing fail `spec-layout-shell.spec`).

## Ostatnia aktualizacja (2026-06-24)

- corpKey z tokenu Entra ID przekazywany przy zapisach z edytora; kolumna „Kto modyfikował" w obu listach admina.
   - GUI: `document-editor.ts` czyta claim `corpKey` i wysyła `corporateKey` w body (`save`/`update`/`finish`).
   - API: DTO/komendy +`CorporateKey` (opcjonalne, fallback do claimu); `Document.LastModifiedBy` + migracja `infra/sql/011_add_document_last_modified_by.sql` (uruchomić na bazach).
   - Listy: `DocumentListItemDto.LastModifiedBy`, `DeliveryListItemDto.CorporateKey`; kolumny + filtry w `admin-files` i `admin-deliveries`.
- Weryfikacja: backend build 0 błędów; `D2Api.Api.UnitTests` 113/113; `Application.UnitTests` 295/295; GUI build OK; `ng test` 268/269 (jeden pre-existing fail `spec-layout-shell.spec`, niezwiązany).
- TODO przy deployu: zastosować migrację SQL `011` na środowiskach (brak EF migrations w tym repo — raw SQL w `infra/sql/`).

## Ostatnia aktualizacja (2026-06-21)

## Ostatnia aktualizacja (2026-06-22)

- Wdrożono centralne security policies dla uploadu i callback URL:
   - `IReturnUrlValidator` + `ReturnUrlValidator` (kontrola: schemat, znaki, protocol-relative, user-info, loopback/private IP, allowlista hostów, normalizacja URL),
   - `IFileUploadSecurityService` + `FileUploadSecurityService` (spójność extension↔MIME↔signature, inspekcja DOCX ZIP: zip-slip, limity, wymagane part-y OOXML, blokada VBA),
   - `IFileScanner` (abstrakcja AV) + `NoOpFileScanner` (domyślny adapter).
- Podpięto egzekwowanie polityk w handlerach:
   - uploady: `UploadDocument`, `UploadImage`, `IngestExternalDocument`,
   - callback/recipient: `UpdateCallbackUrl`, `UpdateDeliveryRecipientUrl`, `FinishAndSendDocument`.
- Hosty (`D2Api`, `D2Services`) bindują sekcje konfiguracyjne:
   - `Security:Upload`,
   - `Security:ReturnUrl`.
- Testy po zmianach:
   - Application.UnitTests **295/295**,
   - D2Api.Api.UnitTests **110/110**,
   - D2Services.Api.UnitTests **39/39**.
- Dodane nowe testy security:
   - `ReturnUrlValidatorTests`,
   - `FileUploadSecurityServiceTests`.

- D2Services observability domknięte do standardu:
   - pipeline zawiera `UseRequestObservability()` + `UseExceptionHandlingMiddleware()`,
   - `X-Correlation-ID` propagowany i logowany (scope + LogContext),
   - `ProblemDetails` zawiera `correlationId`, a błędy walidacji serializują `errors`,
   - formatter JSON rozszerzony o pola `level/service/environment/traceId/spanId` i właściwości scope/eventu.
- Testy D2Services unit: **28/28 pass** (`RequestObservabilityMiddlewareTests`, rozszerzone formatter/middleware/extensions tests).

- Zwiększono liczbę testów jednostkowych w obu backendach:
   - **D2ApiViewerEditor**: +3 testy middleware (`RequestObservabilityMiddleware`, `ExceptionHandlingMiddleware`),
   - **D2ServicesViewerEditor**: nowy projekt `D2ServicesViewerEditor.Api.UnitTests`, rozszerzony do 22 testów (`DocumentController`, `ExceptionHandlingMiddleware`, `GcpJsonSerilogFormatter`, `HealthController`, `MiddlewareExtensions`).
- `D2ServicesViewerEditor.sln` zawiera teraz projekt testowy `D2ServicesViewerEditor.Api.UnitTests`.
- Weryfikacja wykonana:
   - `dotnet test D2ApiViewerEditor/D2ViewerEditor.Api.UnitTests/D2ViewerEditor.Api.UnitTests.csproj` → 76/76 pass,
   - `dotnet test D2ServicesViewerEditor/D2ServicesViewerEditor.Api.UnitTests/D2ServicesViewerEditor.Api.UnitTests.csproj` → 22/22 pass.

## Aktualne zadanie

Przepływ Krok 1–4 jest **zaimplementowany**: ingest (Krok 1), tryb podglądu ładujący v1 (Krok 2), edycja + płaski zapis/auto-save (Krok 3), „Zakończ i wyślij" z synchroniczną 1. próbą + workerem w tle (Krok 4). Bieżąca praca koncentruje się na wierności konwersji DOCX↔HTML (ADR-0023: wiele sekcji, tabele, interlinia) i jakości edytora.

## Kontekst

System przyjmuje dokumenty od aplikacji zewnętrznej (External API), edytuje/ogląda w GUI, zapisuje wersjonowanie w DB+GCS przez Internal API. Oryginał (v1) nietykalny; edycja nadpisuje v2 w miejscu.

## Pliki, które trzeba znać

| Plik / obszar | Dlaczego |
|---|---|
| `.ai/PROJECT_CONTEXT.md`, `.ai/DOMAIN.md` | model i reguły (v1 immutable, płaski zapis) |
| `.ai/API_CONTRACTS.md` | katalog endpointów obu API |
| `D2ApiViewerEditor/.../Features/Documents/Commands/UpdateDocumentVersion/` | nadpisanie wersji |
| `D2ApiViewerEditor/.../Features/Documents/Commands/IngestExternalDocument/` | ingest |
| `D2GuiViewerEditor/src/app/components/document-editor/document-editor.ts` | edytor, auto-save, zapis |
| `D2GuiViewerEditor/src/app/services/document-storage.service.ts` | klient storage API |
| `infra/sql/` | schemat (ręczny SQL) |

## Ostatnie wykonane kroki

1. Dodano płaski zapis (`UpdateVersion` + PUT endpoint + `ModifiedAt` + SQL 005).
2. Dodano endpoint metadanych.
3. GUI: auto-save + switch, ujednolicony „Zapisz" / „Pobierz dokument".
4. Przepisano `.ai/` wg wzorca z `template/`.

## Następne sugerowane kroki

1. Podpiąć produkcyjny silnik AV pod `IFileScanner` (ICAP/ClamAV/API) oraz spiąć alerting dla `MalwareDetected`/`MalwareScanUnavailable` (R-28).
2. Wymusić i udokumentować politykę host allowlist (`Security:ReturnUrl:AllowedHosts`) na środowiskach prod/UAT (R-07 Partial).
3. Dodać DB integration testy claimu wysyłki (`FOR UPDATE SKIP LOCKED` / reclaim po crashu) na realnym Postgresie (R-06).
4. (Opcjonalnie, ADR-0023 follow-up) Nagłówki/stopki per sekcja — wymaga rozszerzenia modelu `DocumentContent` i UI.
5. Sanitizacja paste w edytorze: rozważyć DOMPurify zamiast regexów (R-09).

## Niedokończone zmiany

| Obszar | Co | Ryzyko |
|---|---|---|
| Testy integ. claimu wysyłki | brak (R-06) | regresje w `FOR UPDATE SKIP LOCKED` niewykryte jednostkowo |
| SSRF `RecipientUrl` | `ReturnUrlValidator` wdrożony (ADR-0024), ale bez weryfikacji DNS-rebind i bez obowiązkowej allowlisty na prod (R-07 Partial) | żądania serwerowe na adres z danych zewnętrznych |
| Skaner AV uploadu | `NoOpFileScanner` (R-28) | plik złośliwy przechodzi bez realnego skanu |
| Nagłówki/stopki per sekcja | jeden komplet w modelu (ADR-0023 ograniczenie) | dokumenty z różnymi nagłówkami sekcji spłaszczają je do jednego kompletu |

## Komendy weryfikacji

```bash
dotnet build D2ApiViewerEditor/D2ViewerEditor.sln
cd D2GuiViewerEditor && npm run build
```

Wynik ostatnio: oba OK (0 błędów).

## Nie rób bez zgody

- Nie generuj migracji EF (schemat = `infra/sql/`).
- Nie zmieniaj kontraktów API ani portów bez ustaleń.
- Nie ruszaj reguły v1-immutable / płaskiego zapisu.
- Nie aktualizuj major .NET/Angular. Nie commituj sekretów.
