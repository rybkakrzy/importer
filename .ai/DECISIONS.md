# Technical Decisions

Lekki rejestr decyzji architektonicznych i technicznych.

## Format wpisu

```md
## ADR-XXXX: <tytuł>
- Date: YYYY-MM-DD
- Status: Proposed / Accepted / Superseded / Rejected
### Context
### Decision
### Consequences
### Alternatives considered
```

## ADR-0037: Semantyka wariantów nagłówka/stopki (titlePg / evenAndOddHeaders) — flagi per pasmo, null = dziedzicz, "" = jawnie puste, wariant per pierwsza strona sekcji
- Date: 2026-07-13
- Status: Accepted

### Context
Zgłoszenie: DOC2 Viewer ignorował „Inną pierwszą stronę" — str. 1 renderowała nagłówek default,
a zapis gubił `w:titlePg`. Trzy współdziałające przyczyny: (1) GUI trzymało JEDNĄ flagę
`differentFirstPage` wspólną dla nagłówka i stopki, nadpisywaną przez setter wykonany później
(kontrakt `HeaderFooterContent` raportuje flagę per pasmo — nagłówek może mieć wariant first,
stopka nie); (2) reader traktował pusty/brakujący part first przy aktywnym titlePg jako „brak
wariantu" → default wyciekał na stronę 1, choć Word pokazuje puste pasmo; (3) writer pisał
warianty first/even TYLKO wewnątrz gałęzi niepustego default i tylko przy niepustym first —
dokument first-only (Qutalo/ING) tracił nagłówek przy pierwszym zapisie, a `EnsureTitlePage`
appendował `w:titlePg` przed późniejszym dopisaniem pgSz/pgMar (błąd sekwencji CT_SectPr
w każdym eksporcie z titlePg). Dodatkowo wybór wariantu first opierał się na `pageIndex === 0`
(pierwsza strona CAŁEGO dokumentu), a wpisy sekcyjne ignorowały własne warianty first/even.

### Decision
1. **Flagi per pasmo w GUI** (`_headerDifferentFirstPage`/`_footerDifferentFirstPage` + odd/even);
   publiczne `differentFirstPage()` = OR obu (checkbox), toggle/dialog ustawiają OBA pasma
   (w DOCX titlePg jest właściwością sekcji wspólną dla nagłówka i stopki).
2. **Kontrakt wartości wariantów** w `HeaderFooterContent`: `null` = wariant niezdefiniowany
   (sekcja dziedziczy go z wcześniejszej sekcji / bazy — odpowiednik „Połącz z poprzednim");
   `""` = wariant zdefiniowany i celowo pusty (puste pasmo, NIE fallback do default).
   Baza (sekcja 0) przy titlePg zawsze niesie `FirstPageHtml != null` (brak poprzednika →
   puste pasmo); wpisy sekcyjne ≥ 1 mogą nieść null (dziedziczenie rozwiązuje GUI).
3. **Jeden resolver wariantu** w edytorze (`_resolveHfVariant`) używany przez rendering,
   ładowanie edycji i routing zapisu (rule 10): first na pierwszej stronie KAŻDEJ sekcji
   (z `pageSectionIndexes`), potem even (strony parzyste globalnie), potem default;
   właściciel treści = najbliższy wpis definiujący wariant, inaczej sygnały bazowe.
4. **Writer**: default/first/even emitowane niezależnie; `w:titlePg` zawsze przy
   `DifferentFirstPage` (pusty first = part z pustym akapitem — CT_HdrFtr wymaga bloku);
   pgSz/pgMar wstawiane PRZED titlePg (`AppendBeforeTitlePage`), titlePg przed
   textDirection/bidi/rtlGutter/docGrid (`EnsureTitlePage` pozycjonuje wg CT_SectPr).

### Consequences
- Strona 1 (i pierwsza strona każdej sekcji) respektuje titlePg także dla pustych pasm;
  zapis→otwarcie nie gubi `w:titlePg`/referencji first; eksport z titlePg przechodzi
  walidator OOXML (wcześniej błąd sekwencji).
- Świadome przybliżenia: titlePg sekcji BEZ własnych referencji (brak wpisu w modelu)
  przybliżany flagą najbliższego wcześniejszego wpisu (titlePg w OOXML nie jest dziedziczone,
  ale Word kopiuje je przy tworzeniu sekcji — rozjazd tylko gdy użytkownik ręcznie wyłączył
  titlePg w sekcji dziedziczącej wszystkie party); parzystość even/odd liczona globalnie
  po indeksie strony (bez `w:pgNumType/@start` i restartów numeracji).
- Dokumenty z `w:evenAndOddHeaders` bez referencji even pokazują PUSTE strony parzyste
  (jak Word) — wcześniej pokazywały default (niezgodnie z Wordem).

### Alternatives considered
- Wspólna flaga + „OR" setterów (nie da się wyłączyć titlePg per pasmo, dalej gubi stan).
- Pełne per-sekcyjne `TitlePg` w modelu `DocumentContent` (osobna lista właściwości sekcji)
  — odrzucone jako zmiana kontraktu API nieproporcjonalna do zysku; do rewizji przy
  ewentualnym renderingu `pgNumType`.

## ADR-0036: Kompletna obsługa list DOCX — wariant A: semantyka specyfikacji na istniejącym transporcie `data-*` (bez kanonicznego modelu w kontraktach API)
- Date: 2026-07-12
- Status: Accepted

### Context
Przyjęta specyfikacja „kompletnej obsługi list" wymaga m.in. rozdziału definicja/instancja/
przypisanie (abstractNum/num/numPr), eksportu restartów przez `w:lvlOverride/w:startOverride`,
punktatorów graficznych (`w:numPicBullet`), silnika liczenia etykiet wspólnego dla podglądu
i eksportu oraz kanonicznego modelu JSON niezależnego od HTML. Tymczasem CAŁA aplikacja
(przypisy ADR-0032, sekcje ADR-0023, kotwice ADR-0030, docDefaults ADR-0031) używa HTML
z kontraktem `data-*` jako transportu Domain↔GUI; pełny model kanoniczny w kontraktach API
= przebudowa `DocumentContent`/`SaveDocumentRequest` i modelu stanu edytora (zakazany „broad
refactor"). Writer dodatkowo NIE emitował w ogóle `startOverride`/`lvlOverride` ani
`numPicBullet` — restart numeracji i punktator graficzny ginęły przy pierwszym autosave.

### Decision
1. **Wariant A**: semantyka specyfikacji realizowana na obecnym transporcie — kontrakt `data-*`
   na kontenerach `ol/ul` jest formalną serializacją modelu list (definicja poziomu + tożsamość
   instancji + przypisanie), bez zmian kontraktów API. Wariant B (kanoniczny model JSON
   w odpowiedzi importu / żądaniu zapisu) odrzucony na teraz — do rewizji, gdyby powstał
   frontowy silnik etykiet wymagający pełnych definicji poza DOM.
2. Kontrakt `data-*` list rozszerzony o: `data-start-override` (w:startOverride instancji,
   emitowany ODDZIELNIE od `data-start` = w:start definicji), `data-suffix` (w:suff ≠ tab),
   `data-is-legal` (w:isLgl), `data-lvl-restart` (surowa wartość w:lvlRestart, jednobazowa),
   `data-pic-bullet` (poziom z w:lvlPicBulletId), `data-ind-left/hanging/first-line-tw`
   (wcięcia definicji poziomu w twips).
3. Writer: `w:abstractNum` współdzielony po `data-abstract-num-id` (`_abstractIdByHtmlAbstract`),
   instancja per `data-num-id` jak dotąd; restart = `w:lvlOverride/w:startOverride` na instancji
   (FR-EXPORT-004), NIE kopia definicji z przepisanym w:start. Wcięcia poziomu z `data-ind-*-tw`
   zamiast hardkodowanej drabinki 720×(lvl+1).
4. Punktator graficzny (FR-EXPORT-006): marker `<span class="list-marker"><img data:URI/></span>`
   → `TryCreatePictureBullet` tworzy ImagePart na CZĘŚCI NUMERACJI + `w:numPicBullet`
   (wariant VML `w:pict/v:shape/v:imagedata` — ten sam, który zapisuje Word) + `w:lvlPicBulletId`;
   deduplikacja po data URI; obraz NIE jest już bake'owany jako inline run (dublował się);
   nieosadzalny src/SVG → dotychczasowy fallback (numFmt=none + obraz inline w treści).
5. Kolejność dzieci `w:lvl` naprawiona do sekwencji CT_Lvl (start, numFmt, lvlRestart, isLgl,
   suff, lvlText, lvlPicBulletId, lvlJc, pPr, rPr) — wcześniej rPr szedł przed lvlJc (poza
   schematem). `w:numPicBullet` przed `w:abstractNum` przed `w:num`.
6. Stan konwertera resetowany per `Convert` (`_numberingId`, `_numberingPart`, mapy list) —
   deterministyczny wynik dla tego samego wejścia (FR-EXPORT-002) i brak przecieku części
   między konwersjami.
7. Pełne nadpisanie wyglądu poziomu na instancji (`w:lvlOverride` z własnym `w:lvl`, nie tylko
   startOverride) round-tripuje przez `data-lvl-override="1"`: reader znaczy poziomy, których
   efektywna definicja pochodzi z lvlOverride instancji; writer buduje dla nich `w:lvl`
   WEWNĄTRZ `w:lvlOverride` i NIE wkłada ich definicji do współdzielonego abstraktu
   (dwie instancje wspólnego abstraktu mogą wyglądać różnie — bez tego zlewały się w jeden wygląd).
8. Fragment listy zaczynający się na głębszym poziomie (top-level `ol` z `data-ilvl=N`, np.
   kontynuacja poziomu 1 po zwykłym akapicie) eksportuje się z właściwym `w:ilvl` —
   `ResolveListLevel` honoruje `data-ilvl` w `ConvertListElement` i wszystkich skanach poziomów
   (wcześniej spłaszczany do ilvl=0: złe wcięcie i format poziomu 0).
9. Współdzielony abstrakt jest UZUPEŁNIANY (`UpgradeSharedAbstractLevels`): fragment tworzący
   abstrakt definiuje tylko poziomy, których używa (reszta = drabinka domyślna); późniejszy
   fragment współdzielący abstrakt dosyła definicje brakujących poziomów. Poziomy już zbudowane
   z jawnych data-* nie są podmieniane (pierwsza definicja wygrywa). Bezpieczne, bo fragment
   tworzący nie miał elementów na upgradowanym poziomie.

### Consequences
- „Rozpocznij od nowa" z Worda przeżywa zapis (wspólny abstrakt + nowa instancja z override);
  restart nie zmienia numeracji wcześniejszych elementów po round-tripie.
- Punktatory graficzne przechodzą pełny cykl import→zapis→import bez degradacji do kropki.
- Niestandardowe wcięcia list, suffix, isLgl i lvlRestart nie giną przy pierwszym autosave.
- Eksport listy waliduje się czysto (`OpenXmlValidator` Office2013 = 0 błędów — test).
- Kontynuacja na głębszym poziomie, mieszanie wyglądów instancji wspólnego abstraktu
  i doposażanie poziomów abstraktu przez późniejsze fragmenty — pokryte testami round-trip.
- Silnik etykiet w GUI ZROBIONY (runda 3): `core/utils/list-label.util.ts` (TS, lustrzany do
  liczników readera) + `data-list-label` renderowane przez CSS `::before` poza edytowalnym
  tekstem; atrybuty prezentacyjne zdejmowane przy serializacji. Kontynuacja fragmentów
  przelicza się także po edycji (statyczny `<ol start>` tego nie umiał).
- NADAL poza zakresem (kolejne etapy planu): komendy edytora (Tab/Shift+Tab, kontynuuj/
  restart/ustaw wartość; nowe listy z edytora bez data-* = drabinka domyślna), wcięcia
  znacznika z `data-ind-*-tw` w podglądzie, styl znacznika (rPr poziomu) w kontrakcie,
  pełny łańcuch `basedOn`/`styleLink` w resolverze, `lvlRestart=N` w licznikach READERA
  (silnik TS już honoruje), diagnostyka kodów LIST_*.

### Alternatives considered
Wariant B (pełny kanoniczny model JSON w kontraktach API, ListDefinitionDto/ListInstanceDto) —
zgodny w 100% ze specyfikacją, odrzucony na tym etapie: zmiana kontraktów API i modelu stanu
edytora, wielotygodniowy refactor sprzeczny z zasadą małych zmian; wariant A pokrywa ~90%
wymagań bez ruszania architektury. Emisja `w:numPicBullet` w wariancie DrawingML — odrzucona:
Word sam zapisuje wariant VML, walidator i starsze wersje Worda przyjmują go bez zastrzeżeń.

## ADR-0035: Placeholder pustych bloków edytora — `<br>` w komórkach tabel, `&nbsp;` tylko w akapitach, czyszczenie na beforeinput
- Date: 2026-07-11
- Status: Accepted

### Context
Puste komórki nowych tabel dostawały placeholder `&nbsp;` (żeby kursor miał się gdzie ustawić).
Twarda spacja zostawała przed wpisanym tekstem (UAT: `&nbsp;test` w DOM), przesuwała zawartość
względem MS Word i trafiała jako U+00A0 do zapisanego DOCX. Ten sam problem mają bloki z importu
(reader emituje `&nbsp;` dla pustych akapitów, w tym w `<td><p>&nbsp;</p></td>`).

### Decision
1. Placeholder NOWYCH komórek tabel = `<br>` (insertTable + wstaw wiersz/kolumnę, split/merge).
   Goły `<br>` jako dziecko `td` eksportuje się do pustego akapitu (`ConvertHtmlNode` case "br").
2. Placeholder pustych AKAPITÓW (separator po tabeli, import) zostaje `&nbsp;` — `<p><br></p>`
   eksportowałby się do akapitu z `w:br` (dodatkowa pusta linia w Wordzie).
3. `onEditorBeforeInput` (beforeinput na stronach, nagłówku i stopce): gdy blok
   (p/h1-6/li/td/th) zawiera wyłącznie U+00A0, placeholder jest zaznaczany tuż przed wstawieniem
   tekstu, więc pisanie go zastępuje. Pokrywa import i dokumenty zapisane przed poprawką.

### Consequences
- Tekst w komórce zaczyna się od lewej krawędzi jak w Wordzie; DOCX bez zbędnych U+00A0.
- Dokumenty zapisane przed poprawką czyszczą się dopiero przy edycji danego bloku.
- Kod sprawdzający „pustość" komórki musi akceptować oba placeholdery (robi to merge w
  `document-editor.tableMergeCells`).

### Alternatives considered
- Strip wiodących U+00A0 przy zapisie/imporcie — odrzucone: nie odróżnimy placeholdera od
  celowej twardej spacji użytkownika.
- `<br>` także w akapitach — odrzucone: zmienia eksport (`w:br`), regresja pionowego rytmu.

## ADR-0034: Ochrona przed edycją z DOCX (documentProtection/writeProtection) = tryb tylko-do-odczytu w edytorze
- Date: 2026-07-11
- Status: Accepted

### Context
Dokument Word z „Ogranicz edycję" (settings.xml: `w:documentProtection w:edit="readOnly"
w:enforcement="1"`) albo z hasłem zapisu / zaleceniem tylko-do-odczytu (`w:writeProtection`)
otwierał się w DOC2 Editor jako w pełni edytowalny — konwerter w ogóle nie czytał tych ustawień.
Zgłoszenie: niezgodność zachowania względem dokumentu źródłowego.

### Decision
1. Reader wykrywa ochronę i zwraca `DocumentContent.IsReadOnlyProtected`; GUI otwiera taki
   dokument w istniejącym trybie read-only (badge + toast + blokada zapisu/auto-save).
2. KAŻDY wymuszony tryb `w:edit` ≠ none (readOnly, comments, trackedChanges, forms) = pełna
   blokada edycji — edytor nie umie egzekwować trybów częściowych, więc bezpieczniej blokować
   całość niż pozwolić na edycję ponad uprawnienia.
3. `w:writeProtection` z hasłem = zawsze tylko-do-odczytu (haseł nie weryfikujemy, brak dialogu
   „podaj hasło zapisu"); `w:recommended` traktowane jak ochrona (strona bezpieczna).
4. Egzekwowanie na poziomie GUI; API zapisu nie sprawdza ochrony (patrz RISKS_ASSUMPTIONS).

### Consequences
- Zachowanie zgodne z Wordem dla enforced readOnly; surowsze dla trybów częściowych i samego
  `w:recommended` (Word pozwala wybrać edycję — my nie; użytkownik dostaje jasny komunikat).
- Ochrona NIE jest round-tripowana do eksportu (dokument chroniony i tak nie jest zapisywany).

### Alternatives considered
- Tylko komunikat ostrzegawczy bez blokady — odrzucone: nadal łamie ograniczenie źródła.
- Weryfikacja hasła writeProtection (legacy hash Worda) — odrzucone: niestandardowy algorytm,
  marginalny zysk.
- Egzekwowanie częściowych trybów (comments/forms) w edytorze — poza zasięgiem obecnego WYSIWYG.

## ADR-0033: Wypełnienie kształtu custom-geometry — rozwiązywanie `a:schemeClr`/`a:fillRef` z motywu; brak czarnego fallbacku
- Date: 2026-07-11
- Status: Accepted

### Context
Kształt DrawingML z własną geometrią (`a:custGeom`) bez obrazu/tekstu (np. logo/wordmark Qutalo w stopce)
renderował się jako **czarny wypełniony blob** („czarny knefel"). `GetShapeFillHex` czytał wyłącznie
`spPr/a:solidFill/a:srgbClr` (jawny hex); gdy fill był kolorem motywu (`a:schemeClr val="accent1"`) lub
zdefiniowany przez referencję stylu (`wps:style/a:fillRef`), zwracał null, a `BuildCustomGeometrySvg`
wpadał w `fill="#000000"`. `ResolveThemeColor` obsługuje tylko enum `w:themeColor` (wordprocessing),
nie DrawingML-owy `a:schemeClr`.

### Decision
1. `GetShapeFillHex` rozwiązuje fill w kolejności: (a) `spPr/a:solidFill` → `a:srgbClr` albo `a:schemeClr`;
   (b) fallback do `wps:style/a:fillRef` → `a:srgbClr`/`a:schemeClr`.
2. Nowy `ResolveDrawingSchemeColor` mapuje DrawingML `a:schemeClr` (dk1/lt1/dk2/lt2/tx1/bg1/tx2/bg2/
   accent1..6/hlink/folHlink; domyślny clrMap) na hex ze schematu kolorów theme1.xml.
3. Fallback nierozwiązywalnego wypełnienia w SVG: `currentColor` (dziedziczy kolor tekstu otoczenia),
   NIE solidny `#000000`.

### Consequences
- Brandowe logo z fillem motywowym pokolorowane poprawnie (np. pomarańcz Qutalo z `accent1`).
- Nierozwiązywalny fill degraduje do koloru tekstu zamiast czarnego bloba zasłaniającego grafikę.
- Tylko podgląd (reader); writer nie odtwarza tych kształtów do DOCX — oryginał v1 nietykalny.

### Alternatives considered
- `fill="none"`+stroke (sam kontur) — odrzucone: dla kształtów powierzchniowych gubi masę wizualną.
- Neutralny szary placeholder — odrzucone: `currentColor` lepiej trafia w kolor brandu w stopkach.

## ADR-0032: Przypisy dolne — treść w bocznym modelu (`Footnotes`), odwołania jako `<sup data-footnote-id>` w HTML; tożsamość stabilna, numeracja i `w:id` OOXML liczone osobno
- Date: 2026-07-10
- Status: Accepted

### Context
Aplikacja nie obsługiwała przypisów dolnych Worda (reader je cicho dropował, writer nie tworzył
`footnotes.xml`). Wewnętrzny „model" dokumentu to ciąg HTML (`DocumentContent.Html`) + boczne pola
(nagłówki/stopki/marginesy/sekcje). Przypis musi rozdzielać: miejsce odwołania w treści, treść
przypisu, stabilną tożsamość, numer widoczny i numeryczny `w:id` OOXML — i nie duplikować treści
przy każdym odwołaniu.

### Decision
1. **Model**: nowy `Footnote { Id, Html }`; `DocumentContent.Footnotes` / `SaveDocumentRequest.Footnotes`
   (backend + TS). Treść przypisu = jedno źródło prawdy; odwołania w treści niosą tylko
   `data-footnote-id`. To jest spójne z istniejącym wzorcem „boczne pola obok Html" (jak
   `SectionHeadersFooters`), a NIE osadzanie treści przypisu w body HTML.
2. **Tożsamość vs numer vs `w:id`**: `Id` (np. `fn-1`) stabilny w obrębie importu; numer widoczny
   liczony przy renderowaniu z kolejności PIERWSZYCH odwołań (reader — przy emisji `<sup>`; GUI —
   `syncFootnotesWithBody` po edycji); numeryczny `w:id` OOXML przydzielany dopiero na eksporcie
   (1..N wg kolejności listy). ID nie round-tripują 1:1 — writer je przemapowuje deterministycznie,
   reader odtwarza `fn-<ooxmlId>`. Separatory techniczne mają zarezerwowane id -1 / 0.
3. **Import/eksport przez ISTNIEJĄCE konwertery**: treść przypisu (akapity/runy/formatowanie/tabele)
   idzie przez `ConvertParagraphToHtml`/`ConvertTableToHtml` (reader) i `ConvertHtmlToBody` (writer)
   — bez osobnego, konkurencyjnego konwertera treści.
4. **GUI**: odwołania renderują się w treści strony jako `<sup class="footnote-ref">`; treść
   przypisów w dedykowanym panelu POZA contenteditable body (jedno źródło prawdy, edytowalne
   osobno). Renumeracja/przycięcie osieroconych przy każdej operacji edycyjnej.

### Consequences
Pełny pionowy przepływ DOCX → import → model/API → render+edycja Angular → eksport → poprawny DOCX
→ ponowny import, potwierdzony testami (backend `FootnoteFidelityTests` 18/18 z walidacją OOXML;
GUI 9 speców TestBed). Dokument bez przypisów nie dostaje `footnotes.xml`. Brak infrastruktury E2E
(tylko Vitest) → headless przepływ pokryty TestBed + backend round-trip zamiast Playwright
(dodanie E2E = nieproporcjonalne). Panel przypisów jest podglądowo-edycyjny dla całego dokumentu
(bez przypisania do konkretnej strony — świadome uproszczenie, jak brak wiarygodnego modelu stron
dla przypisów).

### Alternatives considered
(a) Osadzić treść przypisów w body HTML (ukryta sekcja) — odrzucone: łamie „jedno źródło prawdy",
miesza treść przypisu z treścią dokumentu, utrudnia serializację i walidację odwołań.
(b) Użyć numeru widocznego jako tożsamości — odrzucone wprost przez wymagania (niestabilne przy
edycji/reorderze). (c) Zachować identyczne `w:id` OOXML w round-tripie — zbędne: liczy się
semantyka; deterministyczne przemapowanie jest prostsze i odporne na edycję.

## ADR-0031: Domyślne wartości dokumentu (docDefaults) round-tripują przez kontener `.document-content`; spacing stylu tabeli zapiekany inline w komórkach

- Date: 2026-07-08
- Status: Accepted

### Context
Zapis z edytora (autosave / „Zakończ") idzie przez pełną regenerację pakietu (`HtmlToDocxConverter.Convert`), która budowała styles.xml z hardkodowanymi wartościami: docDefaults 11pt / `after=160` / `line=259` i bez definicji stylów tabel. Dokument 12pt z `line=278` i tabelą „Tabela – Siatka" po PIERWSZYM zapisie: tekst malał do 11pt, interlinia się zmieniała, a akapity w komórkach — pozbawione `w:pPr` stylu tabeli (`after=0/line=240`), którego definicja nie istniała w nowym pakiecie — dostawały pełne odstępy docDefaults i wiersze tabel puchły ~2× w Wordzie („tabele po zapisie totalnie się psują"). Dodatkowo GUI rozwijało wrapper `.document-content` przy paginacji i nie odtwarzało go w `getContent()`, więc writer w produkcyjnym zapisie nigdy nie widział nawet fontu dokumentu. Siatka kolumn dryfowała o kilka twips na każdym zapisie (px→twips), a writer emitował `tblBorders val=none` nadpisując obramowania stylu.

### Decision
1. **Kontrakt kontenera**: reader emituje na `.document-content` domyślne wartości dokumentu — inline `font-family`/`font-size`/`line-height` + `data-default-before/after-tw`, `data-default-line`, `data-default-line-rule` (docDefaults/pPrDefault nadpisane per właściwość przez domyślny styl akapitowy `w:default="1"`, jak w Wordzie). Writer (`CaptureDocumentDefaults`) odtwarza z nich docDefaults + Normal; bez kontenera zostają dotychczasowe fallbacki konfiguracyjne.
2. **GUI domyka pętlę**: `_captureDocumentDefaults` przechwytuje wszystkie atrybuty wrappera i rozwija go przy imporcie; `getContent()` owija scaloną treść z powrotem. Interlinia/odstęp akapitu stosowane wizualnie przez `line-height` na `.editor-content` i CSS var `--doc-par-margin` (fallback 10px).
3. **Spacing akapitów w komórkach tabel = rozwiązany INLINE** (docDefaults + łańcuch `w:pPr` stylu tabeli po basedOn) — świadome formatowanie bezpośrednie: dzięki temu eksport jest odporny na brak definicji stylu tabeli w regenerowanym pakiecie. Akapity BODY inline'u nie dostają (interlinia z kontenera; `w:spacing` wraca z docDefaults writera).
4. **Tabele**: `data-w-tw` na `<col>` niesie dokładne twips `w:tblGrid` (writer preferuje je nad px dla gridCol i tcW; ręczny resize kolumny w edytorze usuwa atrybut); tabela z `data-tbl-style` bez CSS `border:` nie dostaje `tblBorders` (styl rządzi liniami; jawny `data-no-borders="1"` = zamierzone none).
5. **Zgodność ze schematem**: dzieci `pPr`/`tcPr` sortowane wg CT_PPrBase/CT_TcPr (`NormalizeParagraphPropertiesOrder`/`NormalizeTableCellPropertiesOrder`), `tcBorders` w kolejności top→left→bottom→right; finalny CSS akapitu deduplikowany (regexy writera biorą pierwsze wystąpienie właściwości).

### Consequences
Round-trip qutable (12pt/278 + Tabela–Siatka + gridSpan/vAlign/jc/trHeight): eksport zachowuje rozmiar/interlinię/odstępy, komórki niosą `w:spacing after=0 line=240`, tblGrid `3020/3021/3021` bez dryfu, walidator OOXML (Office2013) 0 błędów (było 52). Edytor renderuje odstępy między akapitami i interlinię z dokumentu zamiast sztywnych 10px. Pozostałości: `data-no-borders` pojawia się od 2. otwarcia regenerowanego pakietu (brak definicji stylu tabeli — obramowania i tak zapieczone per komórka); definicje stylów tabel wracają w pełni tylko przez `ConvertPreservingPackage` (dziś Download) — docelowo Save powinien dostać pass-through (wymaga masterId w `POST /document/save`, zmiana kontraktu do ustalenia).

### Alternatives considered
(a) Inline'owanie docDefaults spacing na KAŻDYM akapicie body — odrzucone: masywny churn HTML/goldenów i zamiana formatowania stylowego na bezpośrednie w całym dokumencie; kontener załatwia to samo bez churnu. (b) Emisja definicji stylu tabeli do regenerowanego styles.xml — odrzucone: rekonstrukcja stylu z resolved CSS jest stratna i dubluje pass-through; inline spacing komórek pokrywa jedyną realną stratę. (c) Przeniesienie Save na ConvertPreservingPackage — właściwe docelowo, ale to zmiana kontraktu API (masterId w request) — zostawione jako następny krok.

## ADR-0030: Model kotwicy elementów pływających = pozycja w DOM; pola tekstowe round-tripują jako wps:wsp

- Date: 2026-07-07
- Status: Accepted

### Context
Pola tekstowe (`div.docx-textbox`) nie miały żadnej obsługi w writerze — spłaszczane do akapitów przy 1. autosave (utrata ramki/pozycji/kotwicy). Relacja „obiekt ↔ akapit-kotwicy" nie była nigdzie jawnie modelowana, a blokowy `div` emitowany WEWNĄTRZ `<p>` był re-parentowany przez parser przeglądarki (wypadał z akapitu i rozcinał go). Typ zawijania (wrapSquare/topAndBottom) degradował do WrapNone. Edytor nie pokazywał kotwicy i nie pozwalał przesuwać textboxów; drag pływających obrazów ignorował zoom.

### Decision
- Kotwica = **relacja pozycyjna w DOM**, bez sztucznych identyfikatorów akapitów: pływający obraz leży WEWNĄTRZ akapitu-kotwicy (span w `<p>` — legalne), pole tekstowe jest hoistowane bezpośrednio PRZED akapit-kotwicę (reader), a writer przypina jego `wp:anchor` do NASTĘPNEGO akapitu (`BufferTextBoxDrawing`/`AttachPendingTextBoxes`). Round-trip jest idempotentny (test: 2 cykle bez wzrostu liczby akapitów) i odporny na parser HTML.
- Metadane kotwicy jawnie w `data-*` (`data-pos-mode/x-emu/y-emu/width-emu/height-emu/wrap`, jednostki EMU, osie ADR-0029) — wspólny kontrakt obrazów i textboxów; writer odtwarza `wps:wsp`+`w:txbxContent` (pływające → `wp:anchor` H=page/V=margin, inline → `wp:inline`), `data-wrap` → `wrapSquare`/`wrapTopBottom`.
- Obramowanie EDYCYJNE textboxa przeniesione z treści (`border:1px solid #ccc` u readera) do SCSS edytora jako `outline` na hover/zaznaczenie/fokus/drag; obramowanie DOKUMENTOWE (`a:ln` kształtu) jest częścią treści (inline `border` + `data-border-*` → round-trip do `a:ln`).
- Znacznik kotwicy w edytorze = overlay w `.page` (poza contenteditable), pozycjonowany w px układu strony (zoom/scroll bez przeliczeń), czysto informacyjny (pointer-events:none, aria-label). Drag zachowuje kotwicę (zmieniają się tylko offsety) — reguła przewidywalna, bez skoków.

### Consequences
- Pola tekstowe (np. adres Qutalo w stopce) przeżywają autosave z pozycją, rozmiarem, ramką i kotwicą; Word odzyskuje opływanie tekstem dla wrap=square/topAndBottom.
- Ograniczenia: `wps:wsp` bez `mc:AlternateContent`+VML fallback (Word < 2010); wrapTight/Through ≈ wrapSquare (HTML nie niesie wrapPolygon); usunięcie akapitu-kotwicy usuwa obraz w nim zawarty (jak Word), textbox-brat dopina się do kolejnego akapitu; inline textbox w akapicie z tekstem po jednym cyklu ląduje we własnym akapicie przed kotwicą (wizualnie bez zmian — parser i tak go wyjmował).

### Alternatives considered
- Stabilne ID akapitów (`w14:paraId` → `data-para-id`) — odrzucone: wymaga generowania/utrzymania ID w obu konwerterach i edytorze (duplikacja przy kopiuj/wklej, dryf przy split/merge akapitów); relacja pozycyjna daje tę samą semantykę bez nowego stanu.
- Emisja textboxa w `<p>` z naprawą po stronie edytora — odrzucone: parser przeglądarki już zdążył rozciąć akapit zanim JS się uruchomi.
- Serializacja pozycji wyłącznie w inline CSS (stan sprzed zmiany) — odrzucone: sanityzacja/parser gubi kontekst, brak jednostek dokumentowych (EMU) i trybu zawijania.

## ADR-0029: Kotwiczone obiekty (wp:anchor) pozycjonowane jak w Wordzie — relativeFrom + wp:align

- Date: 2026-07-07
- Status: Accepted

### Context
Obiekty pływające (logo, pola tekstowe z adresem) w dokumentach Qutalo renderowały się w edytorze w innym miejscu niż w MS Word — zwykle o cały margines za bardzo w lewo / za wysoko, albo wyrównane do lewej zamiast do prawej. Przyczyna: reader (`ConvertDrawingToHtml`, `BuildTextBoxLayoutCss`) brał surowy `wp:posOffset`, ignorując `relativeFrom` (page/margin/column/…) oraz `wp:align` (right/center/…). Offset kotwiczony do marginesu/kolumny startuje od obszaru treści, nie od krawędzi strony; a wyrównanie bez jawnego offsetu (klasyczny letterhead: logo do prawego marginesu) zeruje się do lewego górnego rogu. Dodatkowo writer eksportował pionową kotwicę jako `RelativeFrom=Page`, mimo że edytor traktuje Y względem góry obszaru treści.

### Decision
- Reader rozwiązuje kotwicę do współrzędnych **układu edytora**: X = lewa krawędź strony (page-relative wprost), Y = góra obszaru treści (od page-relative odjęty górny margines, bo pasmo body zaczyna się pod nagłówkiem). `ResolveAxis` wybiera bazę + rozpiętość wg `relativeFrom`, po czym stosuje jawny offset ALBO wyrównanie `wp:align` w obrębie tej rozpiętości. Rozmiar obiektu do `align` brany z `wp:extent`. Geometria pierwszej sekcji trzymana w polach instancji (`LoadPageGeometry`); brak marginesów → domyślne 1440 twips (jak w Wordzie).
- Writer emituje pionową kotwicę jako `RelativeFrom=Margin` (góra tekstu), poziomą jako `Page`. Dzięki temu round-trip jest idempotentny (reader dodaje i odejmuje ten sam górny margines), a wyeksportowany plik renderuje się w Wordzie tam, gdzie pokazuje go edytor.

### Consequences
- Logo i pola tekstowe kotwiczone do marginesu/kolumny oraz wyrównane do prawej trafiają na właściwe miejsce; brak zmian we froncie (restore czyta gotowe `data-x-emu`/`data-y-emu`). Poprawiony też eksport (dawniej obiekt o górny margines za wysoko w Wordzie).
- Ograniczenia/przybliżenia: pionowe `wp:align` center/bottom słabo określone przy rosnącym obszarze treści (traktowane jak offset/top); `inside`/`outside` bez rozróżnienia stron parzystych/nieparzystych (edytor tego nie modeluje); Y względem góry treści to przybliżenie (pasmo nagłówka ≈ górny margines).

### Alternatives considered
- Liczenie pozycji w froncie z geometrii DOM (`pageMarginPx`) — odrzucone: reader zna dokładne twipsy i rozmiar `wp:extent`, więc rozwiązanie po stronie serwera jest prostsze i wspólne dla obrazów oraz pól tekstowych (te ostatnie nie przechodzą przez JS pływających obrazów). Pozostawienie writer V=Page — odrzucone (łamie idempotencję i zawyża pozycję w Wordzie).

## ADR-0028: Wspólne źródło fontów, moduł jednostek, semantyczny marker podziału strony, zachowanie wartości pól

- Date: 2026-07-06
- Status: Accepted

### Context
Audyt import/edycja/eksport DOCX (prompt „Principal SWE") wskazał konkretne wady edytora i konwersji: (1) pole wyboru fontu w głównym toolbarze białło na focusie (natywny `<select>` + `selectedIndex=-1`); (2) toolbar główny i kontekstowy trzymały DWIE różne listy fontów, bez fontu firmowego; (3) podział strony renderowany jako widoczna kreskowana grafika z etykietą; (4) magiczna stała px/cm rozjechana (linijka 37.795 vs strona 37.8); (5) reader gubił wartość cache pól złożonych ≠ PAGE/NUMPAGES i nadpisywał DATE datą serwera. Brakowało też narzędzi diagnostycznych do porównań round-trip.

### Decision
- **`FontProviderService` (providedIn root)** jako jedyne źródło listy fontów i normalizacji nazw, z fontem firmowym z `--corporate-font-family`. Oba toolbary czytają tę samą `displayNames`.
- **Combobox fontu** (`<input list>`+`<datalist>`) zamiast natywnego selecta: wpisywanie/wyszukiwanie, brak znikania na focusie, commit dopiero po zatwierdzeniu, stan mieszany (`EditorState.fontMixed`).
- **`core/utils/units.util.ts`** — jedyne miejsce konwersji jednostek (`CSS_PX_PER_CM=37.8`). Zoom pozostaje czysto wizualnym `transform:scale` (nie dotyka modelu/eksportu).
- **Podział strony = semantyczny niedrukowalny marker** (`div.page-break`, `contenteditable=false`), grafika tylko opcjonalnie w trybie „znaki formatowania". Eksport bez zmian (już `w:br type=page`).
- **Reader zachowuje wartości pól** (KR-05): tylko PAGE/NUMPAGES dostają dynamiczny placeholder, reszta (w tym DATE) renderuje wartość cache jako tekst; koniec `DateTime.Now` dla DATE (KR-08).
- **`tools/docx-diagnostics/`** — zero-dependency Node ESM `inspect-docx`/`compare-docx` (własny czytnik ZIP na `zlib`), zamiast dodawania zależności JS DOCX.

### Consequences
- Font firmowy dostępny w obu toolbarach; jedna normalizacja. Combobox spełnia kryteria item 6. Linijka i strona nie dryfują. Podział strony nie jest już grafiką ani w edytorze, ani w DOCX. Wartości pól (TOC/REF/MERGEFIELD/DATE) nie znikają przy pierwszym autosave.
- Ograniczenie: round-trip KODU pola (nie wartości) nadal poza zakresem — wartość wraca jako statyczny tekst („unlink field"). Skan „mieszanego" fontu nie działa w jsdom (degraduje do false). Narzędzia diagnostyczne oparte na regex (wystarczające dla diagnostyki, nie pełny parser OOXML).

### Alternatives considered
- Native `<select>` z hackami na blanking — odrzucone (nie da typeahead/wyszukiwania). Dodanie biblioteki DOCX/JS-zip do narzędzi — odrzucone (repo trzyma się zero-dependency). Precyzyjne 37.7953 px/cm wszędzie — odrzucone (łamie istniejące testy na 37.8; korzyść 0,01%).

## ADR-0001: `.ai/` jako pamięć projektu dla agentów AI

- Date: 2026-05-23
- Status: Accepted

### Context
Projekt rozwijany z pomocą agentów AI; potrzebny trwały, jawny kontekst między sesjami.

### Decision
`.ai/` jako pamięć projektu (struktura wg `template/`), plus `CLAUDE.md` i `AGENTS.md` w root.

### Consequences
Stały kontekst i handoff; wymaga dyscypliny aktualizacji.

## ADR-0021: corpKey ustalany wyłącznie z tokenu po stronie API (bez transportu z GUI) — 2026-06-26

- Date: 2026-06-26
- Status: Accepted (zastępuje transport corpKey z 2026-06-24)

### Context
Edytor odczytywał claim `corpKey` ze statycznego snapshotu `getActiveAccount().idTokenClaims` (raz, w konstruktorze) i wysyłał go w body zapisów. Snapshot bywał pusty (claim pojawia się dopiero na cicho odświeżonym tokenie używanym przez `api-token.interceptor.ts`), więc `SaveDocumentVersionCommand.CorporateKey` przychodził null. Token faktycznie wysyłany do API (ID token) niesie `corpKey`, więc backend i tak czytał poprawną wartość przez fallback `_currentUser.CorporateKey`.

### Decision
Backend jest jedynym źródłem tożsamości edytującego: handlery `SaveDocumentVersion`/`UpdateDocumentVersion`/`FinishAndSendDocument` czytają `_currentUser.CorporateKey` (zweryfikowany token; brak → `Result.Failure`). Usunięto pole `CorporateKey` z komend i z controller DTO `SaveDocumentVersionRequest` oraz cały odczyt/wysyłkę `corpKey` w GUI.

### Consequences
Niezawodne i spójne pozyskanie corpKey (zawsze z walidowanego tokenu), mniej kruchego kodu i brak mylącego null w komendzie. Wyświetlanie `corporateKey` na listach admina (z `DocumentDelivery`) pozostaje bez zmian.

### Alternatives considered
Naprawa odczytu w GUI (czytać claim z `acquireTokenSilent()` zamiast snapshotu) — odrzucone: duplikuje logikę, która i tak jest dostępna serwerowo. Pozostawienie fallbacku `request ?? _currentUser` — odrzucone: utrzymuje martwe, mylące pole.

## ADR-0024: Centralne polityki bezpieczeństwa uploadu i callback URL

> Uwaga: pierwotnie zapisane omyłkowo jako drugi „ADR-0020" (numer zajęty przez decyzję o grafikach z tej samej daty). Przenumerowane na ADR-0024 podczas audytu dokumentacji 2026-07-05; odwołania do „ADR-0020" w innych plikach `.ai` dotyczą grafik i pozostają poprawne.

- Date: 2026-06-22
- Status: Accepted

### Context
Walidacja uploadu i `returnUrl` była rozproszona i częściowo oparta na prostych regułach (np. sam schemat http/https), co zwiększało ryzyko SSRF/open-redirect bypassów i przyjęcia złośliwych plików (mismatch MIME/signature, zip abuse).

### Decision
Wprowadzono centralne serwisy policyjne w warstwie Application:
- `IReturnUrlValidator` (`ReturnUrlValidator`) — jedna reguła dla callback/recipient URL (normalizacja + kontrola schematu, znaków, hosta, loopback/private IP, allowlisty).
- `IFileUploadSecurityService` (`FileUploadSecurityService`) — jedna reguła dla uploadów dokumentów/obrazów (extension↔MIME↔magic bytes, DOCX archive hardening).
- `IFileScanner` — abstrakcja integracji AV, używana przez upload policy; domyślnie `NoOpFileScanner` (wymienialny przez DI).
Polityki egzekwowane w handlerach upload/callback (`UploadDocument`, `UploadImage`, `IngestExternalDocument`, `UpdateCallbackUrl`, `UpdateDeliveryRecipientUrl`, `FinishAndSendDocument`).

### Consequences
Spójne zachowanie security w całym backendzie i mniejsze ryzyko regresji przy kolejnych endpointach.
Koszt: dodatkowe zależności konstruktorów handlerów i nowy kontrakt integracyjny dla produkcyjnego skanera AV.

### Alternatives considered
Pozostawienie walidacji per-controller/per-handler (odrzucone: dryf reguł).
Walidacja wyłącznie na API edge (odrzucone: pomija wywołania wewnętrzne i utrudnia testy jednostkowe logiki domenowej).

## ADR-0002: Metadane aplikacji zewnętrznej jako JSON w jednej kolumnie

- Date: (wcześniejsza praca; potwierdzone 2026-05-23)
- Status: Accepted

### Context
Ingest niesie metadane (returnUrl, classification C1..C4), które mogą się rozszerzać.

### Decision
Trzymać je jako JSON w `documents.metadata` (TEXT), bez osobnych kolumn/indeksów. Deserializacja w `GetDocumentMetadataQuery`.

### Consequences
Elastyczność i brak migracji przy nowych polach; brak indeksowania po classification (akceptowalne dziś).

### Alternatives considered
Osobne kolumny `return_url`, `classification` (odrzucone — sztywne, więcej migracji).

## ADR-0003: Płaski model zapisu wersji edytowalnej (nadpisanie w miejscu)

- Date: 2026-05-23
- Status: Accepted

### Context
Auto-save co 30 s tworzyłby v3/v4/v5... przy użyciu `AddVersion`, zaśmiecając GCS.

### Decision
Auto-save i ręczny „Zapisz" nadpisują wersję edytowalną (v2) w miejscu: `Document.UpdateVersion` + `UpdateDocumentVersionCommand` + `PUT .../versions/{versionId}`; re-upload pod tym samym `versionId` zastępuje obiekt GCS (`documents/{versionId}`). Wersja oryginalna (v1) jest nietykalna.

### Consequences
Brak przyrostu obiektów w GCS; brak rollbacku poszczególnych auto-save (akceptowane — oryginał zachowany w v1). `ModifiedAt` śledzi ostatni zapis.

### Alternatives considered
Nowy obiekt + aktualizacja `storage_path` (umożliwia rollback ostatniego zapisu) — odrzucone ze względu na zaśmiecanie storage.

## ADR-0004: „Zapisz" = zapis API, „Pobierz dokument" = pobranie lokalne

- Date: 2026-05-23
- Status: Accepted

### Context
Wcześniej „Zapisz" w GUI tylko pobierał plik (download), nie utrwalał w bazie.

### Decision
„Zapisz" przez `persistDocument()` (PUT przy versionId, inaczej POST). Dawne zachowanie przeniesione do „Pobierz dokument" (`downloadDocument()`). Usunięto `saveDocumentAs()`.

### Consequences
Spójna jedna ścieżka zapisu (ręczny + auto-save). Ręczny „Zapisz" przy włączonym auto-save jest częściowo redundantny (wymuszenie natychmiastowego zapisu).

## ADR-0005: „Zakończ i wyślij" jako kolejka w PostgreSQL + worker (bez brokera)

- Date: 2026-05-25
- Status: Accepted

### Context
Wysyłka finalnego dokumentu na returnUrl musi być asynchroniczna, trwała, odporna na restart i wiele instancji, z retry do 24 h. Nie wolno blokować requestu HTTP ani wysłać później zmodyfikowanej wersji.

### Decision
Tabela `document_deliveries` + claim `FOR UPDATE SKIP LOCKED` z lease (`locked_until`), worker jako `BackgroundService` z bounded concurrency. `FinishAndSend` (handler aplikacyjny) atomowo (jeden SaveChanges) zmienia status dokumentu i tworzy zadanie; finalny plik to niezmienny snapshot GCS `deliveries/{id}` (SHA-256). Retry = exponential backoff + full jitter (cap 15 min), twardy deadline 24 h → DeadLettered. At-least-once + Idempotency-Key. Idempotencja kliknięcia = unique partial index (jedno aktywne zadanie per dokument).

### Consequences
Zero nowej infrastruktury; transakcyjna spójność doc+job bez outboxa. Koszt: polling DB (akceptowalny). `DeliveryWorker:Enabled` pozwala wydzielić wysyłkę na dedykowany host bez zmian kodu.

### Alternatives considered
Broker (Pub/Sub/RabbitMQ) — odrzucony: dokłada drugi system transakcyjny i problem outboxa bez korzyści przy tej skali. Transactional Outbox — zbędny (jedna baza, jedna transakcja). Stałe retry co 15 min — odrzucone na rzecz backoff+jitter (szybsza reakcja na błędy przejściowe, brak thundering herd).

## ADR-0006: Konfiguracja tabeli w bocznym panelu (jeden dok, wyszukiwanie ma pierwszeństwo)

- Date: 2026-05-27
- Status: Accepted (zmienione 2026-05-27: zasada „wyszukiwanie ma pierwszeństwo" odwrócona)

> Aktualizacja 2026-05-27: pierwotna reguła („wyszukiwanie jawnie wywołane nie jest wypierane przez kontekstowy panel tabeli") powodowała konflikt — po powrocie karetki do tabeli dok pozostawał w trybie wyszukiwania i nie pokazywał formatowania tabeli. Nowa reguła: **ponowne zaznaczenie tabeli wypiera otwarte wyszukiwanie** (zamyka je przez `closeFindReplace()` i pokazuje formatowanie tabeli). Dok ma nadal jeden aktywny tryb na raz; `tablePanelManuallyClosed` (×) i ESC bez zmian. Decyzja na żądanie (UX: kontekst użytkownika = tabela). Implementacja w `syncTablePanel()`. Odrzucony wcześniej wariant „auto-przełączanie z wyszukiwania na panel tabeli przy kliknięciu tabeli" jest teraz przyjęty.

### Context
Konfiguracja tabeli była w poziomym pasku `.table-toolbar` pojawiającym się na żądanie nad dokumentem (przy `isInTable()`). Wymóg: przenieść ją do bocznego panelu spójnego z panelem „Wyszukiwanie" (`.search-panel`), dostępnego po kliknięciu istniejącej tabeli. Edytor ma tylko jeden dokowany obszar po lewej, współdzielony dziś przez wyszukiwanie.

### Decision
Nowy komponent prezentacyjny `d2-table-properties-panel` renderowany w tym samym doku co `.search-panel`. Logika tabeli (`table*()`, `setCellColor`, `clearCellColor`, `showTableGridLines`) zostaje w `DocumentEditorComponent` — panel tylko emituje zdarzenia (brak duplikacji, jedno źródło prawdy o aktywnej tabeli przez `activeTable()`/`activeTableCell()`). Sterowanie dokiem: `showTablePanel` ustawiane przez `syncTablePanel()` w `detectTableContext()` (auto-otwarcie po wejściu w tabelę, zamknięcie po wyjściu, respekt ręcznego zamknięcia `tablePanelManuallyClosed`). Konflikt z wyszukiwaniem: jeden dok, wyszukiwanie (jawnie wywołane przez Ctrl+F/lupę) ma pierwszeństwo i nie jest wypierane przez kontekstowy panel tabeli; po zamknięciu wyszukiwania panel tabeli wraca, jeśli karetka jest nadal w tabeli.

### Consequences
Spójny UX z wyszukiwaniem, mniej zagracony obszar nad dokumentem, czytelne grupy ustawień + wydzielona strefa „Usuń tabelę". Brak zmian backendu/API/modelu (funkcja czysto frontendowa). Zachowane wszystkie dotychczasowe akcje tabeli. Koszt: panel i wyszukiwanie nie mogą być widoczne jednocześnie (świadomy kompromis — jeden dok). Pływający `.shading-dropdown` zastąpiony paletą wbudowaną w panel.

### Alternatives considered
Wspólny generyczny mechanizm side-paneli — odrzucony jako zbyt duży refactor względem zakresu (oba panele to proste bloki `@if`). Pozostawienie części akcji w toolbarze jako szybkich akcji — odrzucone, by nie utrzymywać dwóch konkurencyjnych UI; całość trafia do panelu. Auto-przełączanie z wyszukiwania na panel tabeli przy kliknięciu tabeli — odrzucone (wyrywałoby użytkownikowi jawnie otwarte wyszukiwanie).

## ADR-0007: Stylizacja/obramowania tabel jako style inline (bez klas CSS, bez zmian backendu)

- Date: 2026-05-27
- Status: Accepted (zaktualizowane: galeria presetów wycofana na rzecz szczegółowego edytora obramowań)

> Aktualizacja 2026-05-27: galeria gotowych stylów (`TABLE_STYLE_PRESETS`) i opcje header/banded zostały usunięte. Zasada utrwalania (style inline) pozostaje w mocy i dotyczy teraz precyzyjnego edytora obramowań: rodzaj linii (solid/dashed/dotted/double/none), grubość, kolor, miejsce (all/none/outer/inner/inner-horizontal/inner-vertical/top/bottom/left/right) oraz **cel** (cała tabela / komórka / zaznaczony fragment). `applyBorderToCells` liczy krawędzie względem prostokąta opisanego na zbiorze komórek, więc np. „zewnętrzne" rysuje obrys zaznaczenia. Markery `data-*` nie są już potrzebne (brak stanu presetu do odtworzenia).
>
> Aktualizacja 2026-05-27 (auto-zakres): ręczny wybór celu usunięty. Zbiór komórek jest wnioskowany z zaznaczenia (`resolveAutoTargetCells`), a `classifyBorderTarget` zwraca etykietę (komórka/wiersz/kolumna/tabela/zakres) do podpisu w panelu. Domyślnie (sama karetka) → aktywna komórka. Brak gestu „klik nagłówka wiersza/kolumny" w edytorze — wykrycie wiersza/kolumny opiera się na zaznaczeniu komórek (drag).

### Context
Rozbudowa panelu tabeli o stylizację w duchu MS Word (gotowe style, obramowania, wiersz nagłówka, wiersze naprzemienne, reset). Dokument jest utrwalany jako `editor.innerHTML` → serializowany do DOCX i nadpisywany w GCS; po otwarciu HTML wraca do edytora. Styl musi przetrwać zapis/autozapis i ponowne otwarcie, bez utraty treści.

### Decision
Styl tabeli stosowany jako **style inline** na `<table>/<tr>/<td>` (background-color, border, font-weight, color, padding) — te same właściwości i mechanizm, których używają istniejące akcje tabeli (cieniowanie/obramowania/autofit). Świadomie NIE używamy klas CSS. Logika to czyste funkcje w `core/utils/table-style.util.ts` (testowalne w jsdom), preset = deklaratywny model w `models/table-style.model.ts`. `applyTablePreset` przelicza wygląd idempotentnie (czyść → nałóż), więc przełączanie opcji nie kumuluje artefaktów. Zapamiętany preset/opcje (dla stanu panelu) trzymane w atrybutach `data-table-style*` na elemencie tabeli. UI: zakładka „Style" w istniejącym panelu; akcje przez `@Output` → `DocumentEditorComponent` woła util na `activeTable()` + `notifyEditorChange()` (auto-save bez dodatkowej integracji).

### Consequences
Styl przeżywa round-trip HTML w edytorze, zapis do DOCX i eksport poza aplikację (inline jest przenośne; klasy CSS nie przetrwałyby konwersji DOCX). Zero zmian backendu/API/modelu dokumentu. Treść nietknięta (operujemy tylko na `style`, nie na `innerHTML`). Ograniczenia: markery `data-*` mogą zniknąć po konwersji DOCX → po ponownym otwarciu wygląd pozostaje, ale panel pokazuje stan domyślny; wierność DOCX↔HTML zależy od konwertera serwerowego (to samo ryzyko co dotychczasowe cieniowanie). Zaawansowane właściwości (styl linii dashed/double, padding per komórka, wyrównanie pionowe) — etap późniejszy.

### Alternatives considered
Klasy CSS + arkusz stylów tabel — odrzucone: nie przetrwałyby konwersji DOCX ani eksportu, łamią przenośność. Osobny model stylów w backendzie/JSON obok dokumentu — odrzucone: dubluje źródło prawdy, wymaga zmian API i synchronizacji z treścią; inline w samym dokumencie jest prostsze i samowystarczalne. Natywne style tabel DOCX (`w:tblStyle`) — poza zakresem MVP (wymaga rozbudowy konwertera serwerowego), oznaczone jako etap docelowy.

## ADR-0008: Centralna konwersja jednostek OOXML (`OoxmlUnits`) + harness regresji snapshotów HTML

- Date: 2026-06-03
- Status: Accepted

### Context
Refaktoryzacja DOCX→HTML pod wierność z Wordem (plan 10-etapowy). Audyt wykazał: (a) konwersje jednostek (twips/EMU/half-points/cm/px) rozproszone z magicznymi stałymi w obu konwerterach (reguła 14), z dwiema różnymi bazami dla tej samej wielkości (`567` dla cm vs `1440` dla px/pt — `567` to zaokrąglenie `1440/2.54 = 566.929`); (b) brak siatki bezpieczeństwa wykrywającej regresję wierności przed dużymi zmianami pipeline (Etap 2+). Konwertery to ~2950 + ~2525 linii; nie da się ich bezpiecznie refaktoryzować bez testów blokujących wyjście.

### Decision
Etap 1: `D2ViewerEditor.Infrastructure/Conversion/OoxmlUnits.cs` — jedno źródło prawdy: nazwane stałe (`TwipsPerInch=1440`, `EmuPerInch=914400`, `EmuPerPixel=9525`, `TwipsPerPoint=20`, `HalfPointsPerPoint=2`, `DefaultDpi=96`, `CmPerInch=2.54`) + metody konwersji. Oba konwertery przekierowane na `OoxmlUnits`; usunięte zduplikowane prywatne helpery (`TwipsToPx`/`EmuToPx`/`PxToTwips`/`TwipsToCm`/`cmToTwips=567`). Lokalny rounding/truncation pozostaje przy wywołaniach (zachowanie liczbowe niezmienione), centralizujemy tylko stałe i wzory. Bazowe DPI = 96 (definicja CSS px). cm liczone teraz dokładnym `1440/2.54` (zastępuje przybliżenie `567`).

Etap 0: harness regresji w `Infrastructure.UnitTests/Golden/` — deterministyczne buildy DOCX w pamięci (`GoldenDocuments`), helper snapshotów (`HtmlSnapshot`: normalizacja base64 → `[base64]`, `data-image-id` → `[id]`, pól daty → `[date]`; approve-on-missing → baseline w `__snapshots__/*.approved.html`, commitowany). Snapshot blokuje *aktualne* wyjście HTML (regresja, nie poprawność pikselowa) przed zmianami Etap 2+.

### Consequences
Reguła 14 spełniona: zero rozproszonych stałych konwersji; jeden punkt zmiany przy korekcie wzoru/DPI. Zmiany w pipeline mają teraz blokadę regresji (snapshoty + asercje punktowe font-size/indent/szerokości/EMU). Build + `Infrastructure.UnitTests` **66/66** (54 baseline + 7 `OoxmlUnits` + 5 Golden). Brak zmian kontraktu `DocumentContent`/`IDocxToHtmlConverter`/API. Ograniczenie: snapshoty lokują obecne zachowanie (część niezgodne z Wordem) — to świadomie regression-guard, nie golden-correctness; baseline aktualizuje się po świadomej zmianie (`delete *.approved.html` + re-run).

### Alternatives considered
Konwersje per-call (status quo) — odrzucone: magiczne stałe, dryf baz (567 vs 1440), nietestowalne. Zewnętrzna biblioteka jednostek — odrzucone: trywialna matematyka, niepotrzebna zależność. Pełna visual regression (Playwright) jako Etap 0 — odłożone do Etapu 8 (wymaga renderu w przeglądarce); snapshot HTML jest tańszy i wystarczający dla warstwy konwersji. Golden-correctness baselines weryfikowane wzrokowo wobec Worda — przedwczesne (najpierw blokujemy regresję, wierność podnosimy w Etap 3–8).

## ADR-0009: Jawny model pośredni DOCX wprowadzany strangler-pattern (Etap 2+)

- Date: 2026-06-03
- Status: Accepted (in progress)

### Context
`DocxToHtmlConverter` (2950 l.) konwertuje XML→string HTML w jednym przebiegu — brak jawnego modelu (Document/Section/Paragraph/Run/Table), więc testować można tylko finalny string, a parsowanie miesza się z renderowaniem. Plan docelowy: `DOCX → parser → model pośredni → normalizacja → renderer`. Przepisanie całości naraz jest zbyt ryzykowne (round-trip + autozapis chronione testami).

### Decision
Model pośredni wprowadzany **strangler-pattern**: namespace `Infrastructure/DocxModel/`, migracja obszar po obszarze, każdy pod osłoną golden-snapshotów (Etap 0). Pierwszy element: `PageSettings` (geometria strony/sekcji jako surowe twipsy) + `SectionPropertiesReader` (czysty parser `sectPr`→model). `DocxToHtmlConverter` (margines + wysokość pasma nagłówka/stopki) przeliczany teraz z modelu (`ComputeBandHeightCm`), zachowanie 1:1 (snapshoty niezmienione). Page size + orientacja są już parsowane (nieujawniane jeszcze w HTML) — fundament pod Etap 4. Publiczny kontrakt `DocumentContent`/`IDocxToHtmlConverter` bez zmian; każdy etap to mały commit z testami modelu.

### Consequences
Parsowanie oddzielone od renderowania dla sekcji/strony; testy na modelu (`SectionPropertiesReaderTests`) zamiast tylko na stringu. Reszta konwertera (akapit/run/tabela/obraz) migruje w kolejnych etapach (3=style resolver, 4=layout/page, 5=tabele, 6=obrazy, 7=header/footer). Ryzyko: tymczasowo współistnieją stara (string) i nowa (model) ścieżka — akceptowalne przy strangler; golden-snapshoty wykrywają niezamierzone zmiany. Infrastructure **72/72** po migracji.

### Alternatives considered
Big-bang rewrite parsera+renderera — odrzucone: zbyt ryzykowne dla round-tripu/autozapisu, niereview'owalne. Model w warstwie Domain — odrzucone: to szczegół infrastruktury konwersji (DPI/rendering), nie reguła domenowa; trzymamy w `Infrastructure`.

## ADR-0010: Otwieranie DOCX z hasłem (NPOI) + brak in-process konwersji binarnego .doc

### Status
Przyjęte (2026-06-09).

### Context
Wymóg: (1) otwierać DOCX zabezpieczone hasłem, (2) przyjmować .doc i konwertować do DOCX dla edytora.
Oba formaty to kontenery CFB (compound file). Reguły projektu: bez LibreOffice/soffice, bez
System.Drawing, Linux/GCP-safe, minimalne zależności.

### Decision
Wprowadzony `IDocumentInputNormalizer` (Domain) + `DocumentInputNormalizer` (Infrastructure) jako
jeden punkt normalizacji wejścia przed parserem DOCX→HTML (wpięty w `OpenDocumentQueryHandler`):
detekcja po magic-bytes (ZIP/CFB), **dekrypcja DOCX hasłem przez NPOI** (POIFS/Crypt — pure-managed,
zweryfikowane że buduje i działa), detekcja binarnego .doc (stream `WordDocument`).

**Binarny .doc NIE jest konwertowany in-process** — zwraca kontrolowany status `UnsupportedLegacyDoc`
(→ 400 z instrukcją „zapisz jako .docx"). Powód: brak dostępnego, czysto zarządzanego konwertera
.doc→.docx — NPOI 2.8.0 nie zawiera HWPF; pakiet NuGet `b2xtranslator` 1.0.2 zawiera tylko
OpenXmlLib/StructuredStorage, bez `DocFileFormat`/`WordprocessingMLMapping` (parsera .doc); LibreOffice
zakazane. Obsługiwany jest natomiast mislabeled .doc będący w istocie DOCX (ZIP) — pass-through.

Dodano `NPOI 2.8.0` do Infrastructure; SkiaSharp podbity 3.116.1→3.119.2 (NU1605, tranzytywnie z NPOI).

### Consequences
Hasło: działa end-to-end (`/open` + `password`, sentinele PASSWORD_REQUIRED/WRONG_PASSWORD → 422;
GUI prompt+retry). Binarny .doc: kontrolowany komunikat zamiast cichego błędu parsera albo udawania.
Pełna konwersja binarnego .doc wymaga decyzji infrastrukturalnej (sidecar konwersji / usługa) —
poza zakresem reguły „bez ciężkich zależności". `IDocumentInputNormalizer` jest gotowym szwem do
podłączenia takiej konwersji w przyszłości.

### Alternatives considered
Ręczna dekrypcja CFB + Agile/Standard (własny kod) — odrzucone: ~400 linii kryptografii bez fixtury
testowej, wysokie ryzyko cichych błędów; NPOI jest sprawdzone. b2xtranslator do .doc — odrzucone:
NuGet nie zawiera parsera. LibreOffice headless — odrzucone regułą projektu (kontener/GCP).
Normalizacja na ścieżce ingest/storage (dashboard) — odłożone: osobny flow (persist), follow-up.

## ADR-0011: Uwierzytelnianie i autoryzacja przez Microsoft Entra ID

- Date: 2026-05-26
- Status: Accepted; **częściowo zastąpione przez ADR-0012** (2026-06-10) — biblioteka (JwtBearer→Microsoft.Identity.Web) oraz model ról (App Roles → **mapowanie grup→role**, App Roles zachowane jako współistniejące). Reszta (CorporateKey z claimu, admin omija `allowedCorporateKeys`, backend = źródło prawdy, MSAL na froncie) bez zmian.

### Context
Aplikacja nie miała uwierzytelniania; kontrola dostępu do dokumentu (v1) opierała się o `allowedCorporateKeys` + nagłówek `X-Corporate-Key` (seam). Cel: pełne uwierzytelnianie użytkowników + dwie role aplikacyjne (`APP_Pracownik`, `APP_Admin`) z zabezpieczeniem modułu admina.

### Decision
- **App Roles** (nie group claims) — `APP_Pracownik`/`APP_Admin` jako appRoles → claim `roles` (eliminuje group overage, stabilne, czytelne).
- **CorporateKey z access tokena** — konfigurowalny claim (`AzureAd:CorporateKeyClaim`, domyślnie `ck`); brak/pusty → null (dokumenty ograniczone → 403). Bez fallbacku DB w tej wersji.
- **Logowanie dla całej aplikacji** — `[Authorize]` na `BaseApiController`; „publiczny po linku" = każdy zalogowany (authenticated-public). Health pozostaje anonimowy.
- **APP_Admin omija `allowedCorporateKeys`** — decyzja biznesowa: rola admina daje pełny wgląd w treść (bypass w `DocumentAccessGuard`).
- Backend: JwtBearer (`Microsoft.AspNetCore.Authentication.JwtBearer`), `RoleClaimType="roles"`, policy `RequireAppEmployee`/`RequireAppAdmin`, `ClaimsCurrentUserProvider`. Endpointy admina (`GET /`, `deliveries`, `deliveries/{id}/retry`) → `RequireAppAdmin` (401/403). Backend = źródło prawdy.
- Frontend: MSAL (`@azure/msal-angular`/`-browser`), `MsalGuard` na całej aplikacji, `appAdminGuard` (UX) na `/admin`, MsalInterceptor (Bearer), `documentAccessGuard` zostaje. Front nie decyduje o bezpieczeństwie.

### Consequences
Spójna tożsamość; admin chroniony backendowo (wcześniej tylko trasa Angulara). Istniejące linki integratorów wymagają logowania (zmiana zachowania — świadoma). Wymaga konfiguracji Entra (app registration, appRoles, optional claim CorporateKey, redirect URI per env) i `npm install` MSAL. External API pozostaje app-to-app (osobny temat).

### Alternatives considered
Group claims — odrzucone (overage, GUID-y). Fallback DB/Graph dla CorporateKey — odłożony (claim-only na start). Logowanie tylko dla dokumentów ograniczonych — odrzucone na rzecz spójności (cała aplikacja za logowaniem). Admin bez bypassu — rozważone; biznes wybrał pełny wgląd admina.

## ADR-0012: Pełny wzorzec Qutas/D2WebCore dla Entra ID (Identity.Web + grupy→role + Graph + Secret Manager + Keycloak)

- Date: 2026-06-10
- Status: Accepted; **Keycloak dual-auth USUNIĘTY 2026-06-13** (patrz CURRENT_STATE). Aktywacja per-środowisko wymaga realnych wartości Entra/GCP.

### Context
Po wdrożeniu lekkiej integracji Entra (ADR-0011: goły JwtBearer + App Roles) padła decyzja, by ViewerEditor przyjął **pełny wzorzec referencyjny Qutas/D2WebCore** (analiza: `analiza_implementacji_entra_id_pelna.md`). Wybór użytkownika: pełny wzorzec, **mapowanie grup→role**, praca na bieżącym drzewie. Wzorzec Qutas to serwerowy web-app+API; ViewerEditor to SPA+API — adoptujemy części pasujące do tego kształtu (bez serwerowego OIDC/cookie).

### Decision
- **Biblioteka:** `Microsoft.Identity.Web` 3.12.0 (`AddMicrosoftIdentityWebApi`) zamiast gołego `JwtBearer`. JwtBearer NIE jest już pinowany (Identity.Web dostarcza go tranzytywnie per-TFM — pin 8.0.12 dawał NU1605 vs wymóg 9.x na konsumentach net9.0).
- **Mapowanie grup→role (Qutas):** `RolesOptions` (sekcja `Roles`: `GroupPrefix` + `Roles[]{RoleName,GroupNames}`) + `ClaimsTransformer : IClaimsTransformation` — claim `groups` → role aplikacyjne (`APP_Pracownik`/`APP_Admin`). **App Roles zachowane** (claim `roles` z tokena przeżywa) → grupy i App Roles **współistnieją**. Polityki `RequireAppEmployee`/`RequireAppAdmin` bez zmian (wymagają tych samych nazw ról).
- **Microsoft Graph** (v5, **5.103.0** jak Qutas): `IGraphUserService`/`GraphUserService` app-only (`ClientSecretCredential` + `.default`), `GET /api/identity/users?query=` (RequireAppAdmin). Aktywny tylko gdy jest ClientSecret; inaczej `DisabledGraphUserService` (no-op) → lokalnie bez sekretu działa. **Nie** użyto `Microsoft.Identity.Web.MicrosoftGraph` (to Graph v4).
- **GCP Secret Manager** (`Google.Cloud.SecretManager.V1` 2.6.0): `EntraSecretLoader` wstrzykuje `AzureAd:ClientSecret` z `SMC01{ENV}2_APP00404_entra_secret` na starcie. Guard `Enabled` + try/catch → nigdy nie wywala startu (lokalnie wyłączone).
- ~~Keycloak (legacy) dual-auth~~ — **USUNIĘTE 2026-06-13** (`AuthSchemes`/`KeycloakOptions`/sekcja `Keycloak` skasowane; wyłącznie Entra).
- **Frontend runtime config (Qutas):** `assets/configs/config.json` ładowany w `main.ts` **przed** bootstrapem → `RUNTIME_AUTH_CONFIG` (token z root-factory = `environment.auth` jako fallback). Fabryki MSAL + `appAdminGuard` + `CurrentUserService` czytają z runtime-configu. Jeden build na wszystkie środowiska.
- **Wszystkie wartości środowiskowe to placeholdery** (tenant, clientId, grupy `*`, sekrety, GCP project, Keycloak) — nigdy realnych sekretów w repo.

### Consequences
Architektura zbieżna z Qutas w częściach pasujących do SPA+API. Mapowanie grup→role odwraca decyzję „App Roles only" z ADR-0011, ale je zachowuje (hybryda). Build backendu OK; testy: Api.UnitTests **50** (`ClaimsTransformerTests` 6, `AuthSchemesTests` 5, `IdentityControllerTests` 4), GUI **240** (`runtime-config.spec` 4), AOT build OK. **Aktywacja wymaga** realnych wartości w `appsettings.{ENV}.json`/`config.json` + (dla Graph) sekretu z GCP + (dla dual-auth) realnego Keycloaka. Graph używa app-only — wymaga uprawnień aplikacyjnych (`User.Read.All`) + admin consent. **Niezweryfikowane runtime** (brak tenanta/GCP/Keycloak lokalnie): start aplikacji z realnym Identity.Web, walidacja tokenów Entra, faktyczne mapowanie grup z realnego tokena, pobranie sekretu z GCP, selekcja schematu Keycloak.

**Hardening (analiza problemów):** (1) `RoleClaimType="roles"` ustawiony przez **`PostConfigure`** (rejestrowany po `AddMicrosoftIdentityWebApi`) — gwarantuje, że wygrywa nad konfiguracją Identity.Web, inaczej `RequireRole` mogłoby szukać `ClaimTypes.Role`. (2) `AuthSchemes.SelectByIssuer` w try/catch — niepoprawny token → fallback Entra (→401). (3) `RUNTIME_AUTH_CONFIG` ma root-factory default (=`environment.auth`) → injection zawsze się rozwiązuje (testy + bezpiecznik); `main.ts` nadpisuje wartością runtime; fetch `config.json` z fallbackiem przy 404/niepoprawnym JSON.

**Znane ograniczenia:** (a) **Group overage** — user w > ~200 grupach: Entra pomija claim `groups` (emituje `_claim_names`/`_claim_sources` → Graph). Transformer NIE rozwiązuje overage przez Graph → tacy użytkownicy nie dostają ról z grup; ratują App Roles (współistnieją) lub przyszłe rozwiązanie overage. (b) `groups` domyślnie niesie **object-id** (GUID), nie nazwy — konfiguracja Entra (optional claim) musi emitować nazwy `KUTAS_200_*` albo w `RolesOptions.GroupNames` trzeba wpisać GUID-y. (c) start z PUSTYM `AzureAd:ClientId` (bazowy `appsettings.json` PRD) może rzucić walidacją Identity.Web przy pierwszym żądaniu — uruchamiać z env DEV (placeholdery niepuste) lub realną konfiguracją.

### Alternatives considered
Serwerowy OIDC (`AddMicrosoftIdentityWebApp` + cookie) jak w Qutas — **pominięty**: ViewerEditor to SPA (interaktywny login robi MSAL w przeglądarce), backend pozostaje czystym resource-serverem. `Microsoft.Identity.Web.MicrosoftGraph` (Graph v4) — odrzucone na rzecz Graph v5 (zgodność z Qutas 5.103.0 + nowocześniejsze API). Delegated Graph zamiast app-only — odrzucone (lookup userów to funkcja admina, app-only prostsze).

---

## ADR-0013: Realna rasteryzacja EMF/WMF (SkiaSharp) i konwersja binarnego .doc (pure-managed) — 2026-06-11

### Context
Dwa zgłoszenia odrzucały dotychczasowe „honest fallbacky" wprowadzone w ADR-0010 i przy konwersji
grafik (placeholder „EMF — podgląd w Word"): (P4) wymóg realnego podglądu grafiki EMF w przeglądarce,
(P2) realna obsługa starszych plików `.doc`. Reguły bez zmian: bez LibreOffice/GDI/System.Drawing,
Linux/GCP-safe, minimalne zależności. Dostępne darmowe libki: **SkiaSharp 3.119.2** (już w repo,
NativeAssets.Linux) i **NPOI 2.8.0** (POIFS/CFB). Brak HWPF w NPOI potwierdzony.

### Decision
**EMF/WMF → PNG (preview):** `GraphicConversionService` wydobywa osadzony **DIB** z rekordów metafile
(`EMR_STRETCHDIBITS`/`SETDIBITSTODEVICE` — offBmiSrc@48; `BITBLT`/`STRETCHBLT`/`ALPHABLEND` — @84;
+ konserwatywny skan `BITMAPINFOHEADER`), opakowuje w plik BMP i dekoduje przez **SkiaSharp**
(`SKBitmap.Decode` → PNG). To pokrywa najczęstszy realny przypadek (EMF/WMF opakowujący bitmapę).
Placeholder SVG zostaje TYLKO dla czysto wektorowego metafile bez rastra. **Eksport:** reader niesie
oryginalny metafile w `data-original-src` dla KAŻDEGO EMF/WMF (nie tylko placeholdera), a writer
(`ResolveImageSrc`) preferuje oryginał → DOCX dostaje wektorowy EMF/WMF; PNG to wyłącznie podgląd.

**Binarny .doc → .docx:** nowy `LegacyDocBinaryConverter` (pure-managed) parsuje FIB + piece table
(CLX/PlcPcd; PCD.fc compressed=CP1252/Latin1 vs 16-bit Unicode), wyciąga tekst + podział akapitów i
buduje DOCX przez OpenXML. Wpięty w `DocumentInputNormalizer` (gałąź `WordDocument`) z fallbackiem na
`UnsupportedLegacyDoc` przy niespójnej strukturze.

### Consequences
EMF/WMF z rastrem renderują się w edytorze (koniec placeholdera w dominującym przypadku) bez utraty
wierności eksportu (wektor zostaje). `.doc` otwiera się z odzyskanym tekstem zamiast twardego
odrzucenia. **Świadome ograniczenia:** (1) czysto wektorowe EMF wciąż placeholder (brak pure-managed
interpretera wektora); (2) `.doc` odzyskuje tylko tekst+akapity, nie formatowanie/tabele/obrazy.
Zero nowych zależności (SkiaSharp/NPOI już były). Roadmapa pełnej wierności: sidecar LibreOffice.

### Alternatives considered
System.Drawing (Windows-only) — odrzucone (PlatformNotSupported na Linux, sprzeczne z regułą).
Magick.NET z natywnym delegatem WMF/EMF — odrzucone (brak delegatów w obrazie kontenera).
Płatny Aspose/Spire — odrzucone (licencja). Sidecar LibreOffice — roadmapa (infrastruktura).

---

## ADR-0014: Strukturalne logi JSON dla Google Cloud Logging (severity) — 2026-06-11

### Context
W GCP Logs Explorer wyjątki (logowane `LogError`/`LogCritical`/Serilog `Error`/`Fatal`) pojawiały się
jako severity **INFO**. Przyczyna nie była w kodzie (poziomy logów były poprawne), lecz w sinku:
domyślny formatter konsoli (.NET `Microsoft.Extensions.Logging` w internal API; Serilog `WriteTo.Console()`
w external API) pisze **zwykły tekst na stdout**, a Cloud Logging nadaje każdemu wpisowi stdout severity
DEFAULT/INFO, bo nie ma pola `severity`.

### Decision
Oba hosty emitują na stdout **JSON w jednej linii z polem `severity`** (mapowanym z poziomu logu na
LogSeverity GCP), gdy nie działają w Development (lub gdy `Logging:UseGcpFormat=true`); lokalnie zostaje
czytelny tekst. Wyjątek dołączany do `message` (Error Reporting grupuje po stack trace).
- Internal API: `GcpJsonConsoleFormatter : ConsoleFormatter` + `AddGcpStructuredLogging()` (Program.cs).
- External API: `GcpJsonSerilogFormatter : ITextFormatter` podany do `WriteTo.Console(...)` (Program.cs).
Zero nowych zależności (formattery z frameworka / już obecnego Seriloga).

### Consequences
`LogError`/`LogCritical` widać w Logs Explorer jako ERROR/CRITICAL; severity filtrowalne; stack trace w
treści wpisu. Plik-sink Seriloga (lokalny dev) bez zmian. Mapowanie pokryte testami (`GcpJsonConsoleFormatterTests`).

### Alternatives considered
`Serilog.Sinks.GoogleCloudLogging` / `Google.Cloud.Logging` (push do API) — odrzucone: wymaga creds/SDK,
a na Cloud Run/GKE idiomatyczny jest structured stdout. Logowanie błędów na stderr — odrzucone: gubi
rozróżnienie WARNING/ERROR/CRITICAL (stderr = ERROR ryczałtem).

### Known follow-up
`LoggingBehaviour<TRequest,TResponse>` loguje `{@Request}` na Information — serializuje pełny payload
(np. base64 treści dokumentu): hałas + potencjalne dane wrażliwe. Rekomendacja: logować tylko nazwę
żądania / wybrane pola. Poza zakresem tej zmiany (dotyczy severity).

---

## ADR-0015: Front MSAL standalone (MsalRedirectComponent) + dwurolowy model „Administrator"/„Przeglądający" — 2026-06-16

- Date: 2026-06-16
- Status: Accepted; **model ról zastąpiony przez ADR-0016 (2026-06-17)** — frontowe guardy per-nazwa-roli (`Administrator`/`Przeglądający`, `documentRoleGuard`/`appAdminGuard`, `CurrentUserService`, `adminRole`/`viewerRole` w config) **usunięte** na rzecz autoryzacji resource-based. Reszta ADR-0015 (MSAL standalone, `MsalRedirectComponent`, config.json jako źródło auth, rename tokenu, `apiScopesFor`/`.default`, `navigateToLoginRequestUrl`) — **bez zmian, aktualna**.

### Context
Doprecyzowanie integracji Entra na froncie (Angular 20 **standalone**, nie NgModule jak referencja Qutas/D2AngularNew). Cele: poprawny redirect handling, jeden build na środowiska bez trzymania wartości auth w `environment`, oraz docelowy model RBAC z dwiema rolami: **moduł admina tylko dla „Administrator"**, **podgląd/edycja dokumentów dla „Przeglądający"** (admin = nadzbiór).

### Decision
- **Redirect handling:** `MsalRedirectComponent` (`<app-redirect>` w `index.html`) bootstrapowany jako drugi komponent przez `appRef.bootstrap(...)` w `main.ts` — kanoniczny wzorzec MSAL dla aplikacji standalone, **bez AppModule**. `App` przestaje wołać `handleRedirectObservable()` (robi to redirect-component); aktywne konto ustawiane po `MsalBroadcastService.inProgress$ === None`.
- **Źródło wartości auth = `assets/configs/config.json`** (plik w `src/assets/configs/`, nie `public/`). Usunięto blok `auth` z `environment.ts`/`environment.development.ts`. `DEFAULT_AUTH_CONFIG` w `runtime-config.ts` to **neutralne defaulty strukturalne** (fallback/testy), nie wartości środowiskowe.
- **Token DI przemianowany** `RUNTIME_AUTH_CONFIG` → **`MSAL_CUSTOM_CONFIG`** (parytet nazewnictwa z analizą Qutas).
- **Scope:** helper `apiScopesFor(auth)` — preferuje jawne `apiScopes` z config.json, fallback `{clientId}/.default`; używany spójnie w guardzie i interceptorze. `navigateToLoginRequestUrl: true` ustawiony jawnie (domyślny MSAL).
- **Dwurolowy model (front, UX):** `AppAuthConfig.adminRole` (domyślnie **„Administrator"**) + nowe `viewerRole` (**„Przeglądający"**). Logika scentralizowana w `CurrentUserService`: `isAdmin()`, `isViewer()`, `canAccessDocuments()` (= viewer **lub** admin — admin nadzbiór; dev-bypass `enabled=false` przepuszcza). Nowy `documentRoleGuard` na `/editor` i `/viewer` (przed `documentAccessGuard`); `appAdminGuard` zrefaktoryzowany do `CurrentUserService.isAdmin()`. Dashboard bez zmian (gating na trasie). Backend = źródło prawdy.

### Consequences
Front spójny z MSAL standalone; jeden build na środowiska; auth wyłącznie w config.json (deploy podmienia plik). **Rozjazd nazw ról z backendem:** ADR-0011/0012 używały `APP_Pracownik`/`APP_Admin`; front używa teraz `Przeglądający`/`Administrator`. Trzeba albo (a) skonfigurować Entra appRoles/`RolesOptions` na te nazwy, albo (b) nadpisać `adminRole`/`viewerRole` w config.json wartościami zgodnymi z backendem — inaczej guardy odmówią dostępu mimo poprawnego tokenu. GUI testy: 264 (dodano `current-user.service.spec` — 5; `app.spec`/`runtime-config.spec` zaktualizowane). Niezweryfikowane runtime (brak realnego tenanta lokalnie): faktyczny przepływ redirect + obecność claimu `roles` z realnymi nazwami.

### Alternatives considered
Konwersja na NgModule (`platformBrowserDynamic().bootstrapModule(AppModule)`) jak dosłownie w analizie — **odrzucone** (broad refactor, brak korzyści; standalone już realizuje runtime-config). Ścisły rozdział ról (dokumenty wymagają *dokładnie* „Przeglądający", admin bez niej zablokowany) — odrzucone na rzecz nadzbioru (admin nie zablokuje się sam). Własny `jwt.interceptor.ts` (jak legacy) — odrzucone: `MsalInterceptor` realizuje to samo (Bearer + silent refresh + cache) bez ręcznego trzymania tokenu w localStorage.

---

## ADR-0016: Autoryzacja resource-based (backend `/identity/resources` + `ResourcesProvider`, front `resourceGuard`) — 2026-06-17

- Date: 2026-06-17
- Status: Accepted; **zastępuje frontowy model ról z ADR-0015** (per-nazwa-roli). Zgodne z wzorcem D2WebCore (`ResourcesProvider` + `AuthGuard.hasAccessToResource`). Rozwiązuje rozjazd nazw ról (R-25): front nie zna nazw ról.

### Context
Inne aplikacje ekosystemu (D2WebCore/D2AngularNew) autoryzują **po zasobach**, nie po nazwach ról na froncie: backend mapuje role→zasoby i wystawia listę dozwolonych zasobów, a frontowy `AuthGuard` bierze nazwę trasy i pyta backend „czy mam dostęp". ADR-0015 zrobił gating po nazwach ról czytanych z claimu na froncie (`Administrator`/`Przeglądający`) — niezgodne z backendem (`APP_Pracownik`/`APP_Admin`, R-25) i sprzęgające front z nazwami ról.

### Decision
- **Backend (D2ApiViewerEditor):** `ResourcesProvider.GetForUser(ClaimsPrincipal)` mapuje role (`User.IsInRole` z `AzureAdOptions.EmployeeRole`/`AdminRole`) → zasoby: employee/admin → `["editor","viewer"]`, admin dodatkowo → `"admin"` (nadzbiór). Endpoint `GET /api/identity/resources` (`[Authorize(RequireAppEmployee)]`) zwraca `string[]`. W `IdentityController` polityka admina przeniesiona z poziomu klasy na akcję `users` (żeby `resources` było dostępne dla employee). DI: `AddScoped<ResourcesProvider>()` poza gałęzią dev/prod (w dev `IOptions<AzureAdOptions>` = defaulty `APP_Pracownik`/`APP_Admin`). `DevAuthHandler` daje obie role → dev widzi wszystko.
- **Front (D2GuiViewerEditor):** `ResourceAccessService` (cache udanej odpowiedzi per sesja, dev-bypass `enabled=false` → `of(true)` bez wołania API) + `resourceGuard` (nazwa zasobu = `route.routeConfig.path`; fail-closed → `/access-denied` przy odmowie/błędzie). `resourceGuard` zastąpił `documentRoleGuard` i `appAdminGuard` na `/editor`,`/viewer`,`/admin`. Token dokleja `MsalInterceptor` (URL pod `apiUrl`).
- **Usunięte (front):** `documentRoleGuard`, `appAdminGuard`, `CurrentUserService` (+spec), pola `adminRole`/`viewerRole` z `AppAuthConfig`/`DEFAULT_AUTH_CONFIG`/`config.json` — front nie zna już nazw ról.
- **Config (parytet ekosystemu):** `AzureAdOptions` rozszerzony o `Scopes: string[]` i `Proxy: { Url }` (+ w `appsettings.json`/`appsettings.DEV.json`). **Konsumpcja proxy wdrożona 2026-06-17** (`EntraBackchannel` → JwtBearer `BackchannelHttpHandler` + `HttpClient.DefaultProxy`, bypass GCS; patrz CHANGELOG). **Pozostaje follow-up:** proxy dla Graph/Azure.Identity (własny pipeline) oraz downstream `Scopes`.

### Consequences
Front odsprzężony od nazw ról (znika R-25 po stronie frontu); backend = jedyne źródło mapy rola→zasób. Dodanie nowego zasobu = zmiana w `ResourcesProvider` + nazwa trasy. Testy: API **64** (IdentityController 7, w tym 3 resources), GUI **262** (dodano `resource-access.service.spec` 3; usunięto `current-user.service.spec` 5). Build API 0 błędów, `tsc` GUI czysto. **Wymaga**, by realne Entra appRoles/`RolesOptions` emitowały `APP_Pracownik`/`APP_Admin` (nazwy z `AzureAdOptions`) — inaczej `IsInRole` zwróci false i lista zasobów będzie pusta. **Niezweryfikowane runtime:** realny token z claimem `roles`, wywołanie `/resources` z frontu z Bearer.

### Alternatives considered
Zostawić role-claim na froncie (ADR-0015) z poprawą nazw do `APP_Admin`/`APP_Pracownik` — odrzucone: utrzymuje sprzężenie frontu z nazwami ról i duplikuje wiedzę o autoryzacji. Resource = nazwa semantyczna (`documents`/`admin`) z mapą tras na froncie — odrzucone na rzecz `route.path` = nazwa zasobu (prościej, wiernie wzorcowi). `return of(true)` na błędzie (jak legacy `AuthGuard`) — odrzucone: fail-open to dziura; u nas fail-closed (`/access-denied`).

---

## ADR-0017: Finalne nazewnictwo ról aplikacyjnych — `Administrator` / `Operator` — 2026-06-17

- Date: 2026-06-17
- Status: Accepted; **zastępuje nazwy z ADR-0011/0012** (`APP_Admin`/`APP_Pracownik`). Wcześniejsze ADR-y odnoszą się do starych nazw historycznie.

### Context
Dotychczasowe role: `APP_Admin` (admin) i `APP_Pracownik` (standardowy użytkownik). Decyzja biznesowa: ostateczne nazewnictwo to **`Administrator`** i **`Operator`**. Po przejściu na autoryzację resource-based (ADR-0016) nazwy ról żyją wyłącznie w backendzie — GUI nie zna nazw ról.

### Decision
- Globalny rename wartości ról w backendzie: `APP_Admin`→`Administrator`, `APP_Pracownik`→`Operator`. Objęte: `AzureAdOptions.AdminRole` (default), sekcja `Roles` w `appsettings.json`/`appsettings.DEV.json` (`RoleName`), `DevAuthHandler` (claimy `roles`), polityki (przez opcje, bez literałów), testy (`ClaimsTransformerTests`, `IdentityControllerTests`) + komentarze.
- **Rename identyfikatorów** (spójność z „Operator"): właściwość `AzureAdOptions.EmployeeRole`→**`OperatorRole`** (więc też klucz configu `AzureAd:OperatorRole`), stała polityki `RequireAppEmployee`→**`RequireAppOperator`** (nazwa + wartość; konsumowana tylko przez stałą, więc bezpieczne), zmienna/komentarze (`isEmployee`→`isOperator`, nazwy testów `*_employee_*`→`*_operator_*`). `RequireAppAdmin`/`AdminRole` bez zmian („Admin" = Administrator).
- **GUI bez zmian kodu** — autoryzacja resource-based (ADR-0016), front nie zna nazw ról; potwierdzone grepem (zero literałów ról w `D2GuiViewerEditor/src`).
- Semantyka bez zmian: `RequireAppOperator` = `Operator` lub `Administrator`; `RequireAppAdmin` = `Administrator`; admin nadzbiór (omija `allowedCorporateKeys`). `ResourcesProvider`: Operator/Administrator→`editor,viewer`; Administrator→`+admin`.

### Consequences
Jedna spójna nazwa w całym kodzie/konfigu/testach. **Wymaga** (R-26), by realne Entra appRoles/`RolesOptions` emitowały `Administrator`/`Operator` (albo nadpisać `AzureAd:EmployeeRole`/`AdminRole` per środowisko) — inaczej `IsInRole` = false → pusta lista zasobów. Testy: API **64/64**, build 0 błędów. Dokumenty „żywe" (SECURITY/DOMAIN/FEATURES/API_CONTRACTS/CURRENT_STATE) zaktualizowane; historyczne ADR-0011/0012/0015/0016 zachowują stare nazwy. **Uwaga niezwiązana:** `D2ViewerEditor.Application.UnitTests` ma wcześniej istniejący błąd kompilacji (`GetDocumentVersionContentQueryHandlerTests` — brak arg. `accessGuard`), niezależny od tej zmiany.

### Alternatives considered
Zostawić `APP_*` i tylko zmapować w Entra — odrzucone: rozjazd kodu z biznesowym nazewnictwem. Trzymać nazwy też na froncie — bezprzedmiotowe po ADR-0016 (front nie zna ról).

---

## ADR-0018: Observability ELK — structured JSON na stdout + correlation middleware — 2026-06-17

- Date: 2026-06-17
- Status: Accepted; rozszerza ADR-0014 (GCP JSON severity) o pełny standard ELK. Bez nowej biblioteki.

### Context
Logi miały `severity` (GCP) ale formatter **ignorował scope'y**, brak correlationId per HTTP request, brak access-logu i pól service/environment/trace. Cel: spójne, korelowalne, filtrowalne logi dla ELK (Elasticsearch/Logstash/Kibana) na produkcji.

### Decision
- **Format:** rozszerzony `GcpJsonConsoleFormatter` — jedna linia JSON na stdout, dual ELK+GCP: `timestamp/severity/level/message/category/service/environment/traceId/spanId/exceptionType` + flatten **scope'ów** (`ForEachScope`) i **argumentów szablonu** (klucze zarezerwowane chronione). `service`=ApplicationName, `environment`=EnvironmentName przez `StructuredLogFormatterOptions`.
- **Korelacja:** `RequestObservabilityMiddleware` (outermost, przed exception) — `X-Correlation-ID` z nagłówka lub generowany (`Activity.TraceId`/GUID), scope `{correlationId, requestId}` na cały request (więc wyjątki też go niosą), nagłówek w odpowiedzi, `correlationId` w ProblemDetails. Jeden **access-log** na request (Info/Warn/Error wg statusu) z method/path/statusCode/elapsedMs/userId.
- **Bezpieczeństwo:** nie logujemy treści plików/tokenów/sekretów/PII; `userId`=subject/oid (nieosobowy); `corporateKey` w wysyłce tylko jako flaga obecności.
- **Stack:** wyłącznie wbudowany `ILogger` + scope'y + ConsoleFormatter (zero nowych zależności — bez Serilog/OTel/Elastic APM).
- **Dokumentacja:** `.ai/OBSERVABILITY.md` (pola, korelacja, poziomy, KQL Kibana, zasady nie-logowania).

### Consequences
Logi indeksowalne i korelowalne w Kibanie; błędy z kontekstem (correlationId, trace, exceptionType, stack). Format niesie nadal `severity` → GCP Cloud Logging bez zmian. Testy: `GcpJsonConsoleFormatterTests` (+service/env/level, +scopes), `RequestObservabilityMiddlewareTests` (correlation header). Lokalnie (Development) zostaje czytelny tekst — JSON włączany poza Dev / `Logging:UseGcpFormat=true`.

### Alternatives considered
Serilog + Elastic.Serilog.Sinks / Elastic APM — odrzucone: nowa zależność, a wbudowany formatter+scope realizują structured JSON na stdout (idiomatyczne dla zbieraczy ELK). Osobny correlation-id provider/DI — odrzucone na rzecz `HttpContext.Items` + scope (prościej, bez stanu współdzielonego).

---

## ADR-0019: Hybryda WebApi + WebApp (serwerowy OIDC `AddMicrosoftIdentityWebApp`) — 2026-06-19

- Date: 2026-06-19
- Status: Accepted; **odwraca „resource-server-only"** z ADR-0011/0012 na wyraźną decyzję właściciela (parytet z D2WebCore `ConfigureAuthentication`).

### Context
ADR-0011/0012 świadomie pominęły serwerowy OIDC (`AddMicrosoftIdentityWebApp`), bo backend jest resource-serverem dla SPA (login interaktywny robi MSAL w przeglądarce). Właściciel zdecydował o przyjęciu pełnego wzorca ekosystemu z **obydwoma** schematami oraz jawnym `IS_LOCAL_DEV`.

### Decision
- **Wiring w `Security/ConfigureAuthentication.cs`** (`AddEntraIdAuthentication`). **WebApi** (JWT bearer) pozostaje **schematem domyślnym** — API SPA bez zmian. **WebApp** dodany jako scheme **`MyAzureAdScheme`**: `AddMicrosoftIdentityWebApp` (code flow, `SignInScheme`=cookie, `NonceCookie/CorrelationCookie SecurePolicy=Always`, `ResponseType=Code`, scope `offline_access`/`email`, `ValidateIssuerSigningKey`) + `EnableTokenAcquisitionToCallDownstreamApi()` + `AddInMemoryTokenCaches()`.
- **Proxy:** jawny `IS_LOCAL_DEV` (env var) — lokalnie `UseProxy=false`; inaczej `WebProxy` z `AzureAd:Proxy:Url` (bypass `storage.googleapis.com`) jako `BackchannelHttpHandler` (oba schematy) + `HttpClient.DefaultProxy`. `EntraBackchannel.CreateProxy` buduje `WebProxy`.
- `ShowPII`/`IncludeErrorDetails` gated `!IsProduction()`. `RoleClaimType="roles"` nadal przez `PostConfigure` (wygrywa).
- **Bez nowej zależności** — `AddMicrosoftIdentityWebApp`/OIDC handler tranzytywnie z `Microsoft.Identity.Web` 3.12.0.

### Consequences
Backend potrafi serwerowy interaktywny login (cookie/OIDC) obok walidacji tokenów API. Build OK; Api.UnitTests 71. **Wymaga realnej konfiguracji** (ClientId/TenantId/ClientSecret z GCP) i — dla realnego flow — endpointu logowania/redirect URI w app registration. WebApp to scheme **niedomyślny**: uruchamia się tylko na jawny challenge `MyAzureAdScheme`, więc dla obecnych endpointów API (JWT) nic się nie zmienia. **Runtime niezweryfikowane** (brak tenanta/secretu lokalnie) — R-27.

### Alternatives considered
Zostawić tylko WebApi (ADR-0011/0012) — odrzucone decyzją właściciela. Zastąpić WebApi przez WebApp — odrzucone: SPA potrzebuje walidacji JWT dla wywołań API; hybryda zachowuje oba.

## ADR-0020: Brak widocznego placeholdera dla nierenderowalnych grafik — przezroczysty blank + pass-through — 2026-06-22

- Date: 2026-06-22
- Status: Accepted; **zastępuje** wcześniejsze zachowanie placeholdera SVG (szare tło + „… — podgląd w Word") z GRAPHICS_CONVERSION.md.

### Context
Konwerter grafik wstawiał dla EMF/WMF bez osadzonego rastra (oraz nieobsługiwanych formatów) widoczny szary placeholder z tekstem. To wprowadza w błąd (udaje treść/komunikat błędu w edytorze). Wymóg: przeglądarka nie może pokazywać żadnego fałszywego placeholdera; realna treść albo nic widocznego + raport wewnętrzny. Stack jest pure-managed (Linux/GCP), bez GDI/System.Drawing/LibreOffice → brak rasteryzera wektora EMF/WMF (ADR/GRAPHICS_CONVERSION §4).

### Decision
- Łańcuch strategii produkuje **wyłącznie formaty renderowalne**: osadzony raster (PNG/JPEG) → DIB→PNG (SkiaSharp) → dekoder rastra (SkiaSharp dla TIFF/Unknown). 
- Gdy żadna realna strategia nie zwróci rastra: **przezroczysty, pusty SVG** o wymiarach z layoutu (`IsBlankFallback=true`) — zero widocznej treści (brak rect/fill/stroke/text). **Nigdy** szare tło ani tekst. Niepowodzenie raportowane w `GraphicConversionDiagnostics` (`Status`, `AttemptedStrategies`, `FailureReason`).
- Oryginalny metafile **pass-through** do DOCX (`data-original-src`) — Word renderuje prawdziwy wektor. Atrybut podglądu `data-legacy-graphic="blank"`.
- Deduplikacja po SHA-256 treści (cache w obrębie instancji, bounded).

### Consequences
Edytor pokazuje prawdziwy raster gdy się da, a dla czystego wektora — niewidoczny element zachowujący układ (stabilność layoutu) zamiast mylącego placeholdera; eksport do Word zachowuje pełną wierność. Pełny podgląd wektora EMF/WMF w przeglądarce wymaga rasteryzera out-of-process (sidecar) — roadmapa. Testy: Infrastructure.UnitTests 188, GraphicConversion 35 (w tym regresje no-placeholder).

### Alternatives considered
(a) `Web=null` i pominięcie elementu — odrzucone: psuje układ i grozi wyemitowaniem surowego `data:image/x-emf` w `src` (przeglądarka nie renderuje → broken image). (b) Rasteryzacja wektora przez GDI/System.Drawing — odrzucone (Windows-only, crash na Linux/GCP). (c) LibreOffice headless — odrzucone (proces zewnętrzny, niestabilny w kontenerze).

## ADR-0022: Rola `Operator` egzekwowana na backendzie edytora (klasowy `RequireAppOperator`) — 2026-07-03

- Date: 2026-07-03
- Status: Accepted; wzmacnia ADR-0011/ADR-0016/ADR-0017 (model resource-based był tylko UX-gate na froncie; backend edytora nie egzekwował roli).

### Context
Audyt bezpieczeństwa wykazał **broken access control (OWASP A01)**: `DocumentController` (open/save/sign/upload-image/new/templates/verify-signatures) oraz większość `DocumentStorageController` (upload, `{id}/save`, `PUT versions/{id}`, restore, finish&send, download) dziedziczyły wyłącznie `[Authorize]` z `BaseApiController` — czyli **dowolny zalogowany użytkownik, także z pustą listą ról**. `RequireAppOperator` była nałożona tylko na `GET /api/identity/resources` (lista zasobów dla front-guarda) i endpointy admina. W efekcie użytkownik bez roli `Operator`/`Administrator` mógł przez bezpośrednie API (curl/Postman/DevTools) wykonać pełny cykl życia dokumentu. Dodatkowo pulpit (`''`) był chroniony tylko `authGuard` (bez `resourceGuard`) i pokazywał akcje edytora — stąd objaw „user bez roli widział edytor".

### Decision
- Klasowy `[Authorize(Policy = AuthorizationPolicies.RequireAppOperator)]` na `DocumentController` i `DocumentStorageController`. Endpointy admina zachowują metodowy `RequireAppAdmin` — atrybuty łączą się **AND**, więc wymagają Administratora (który i tak spełnia `RequireAppOperator`).
- Front (defense-in-depth, nie zabezpieczenie): pulpit bramkuje akcje sygnałem `canUseEditor` (`ResourceAccessService.hasAccessToResource('editor')`) + twardy `ensureEditorAccess()` w `newDocument()`/`openFile()`. `resourceGuard` rozróżnia `loading`/transient (401/0/5xx → retry ×3, 300 ms) od definitywnego `403` (deny → `/brak-uprawnien`), by niegotowy MSAL po deep-linku/reloadzie nie był mylony z brakiem uprawnień.
- Egzekwowanie backendowe pod testem regresji: `D2ViewerEditor.Api.IntegrationTests` (WebApplicationFactory<Program> + `TestAuthHandler`, schemat testowy jako domyślny; host hermetyczny bez DB/GCS/workera, dummy ClientId/TenantId aby przeszła walidacja Microsoft.Identity.Web).

### Consequences
Backend jest źródłem prawdy dla dostępu do edytora; ukrycie przycisków to już tylko UX. Użytkownik bez roli aplikacyjnej dostaje **403** na endpointach dokumentu (i **401** bez tokenu), niezależnie od frontu. Weryfikacja: integracyjne **9/9**, GUI `resource.guard`/`dashboard`/`resource-access` **9/9**, build API 0 błędów. Do potwierdzenia poza repo (Entra): przypisania App Roles `Operator`/`Administrator`, „Assignment required=Yes", redirect URI typu SPA.

### Alternatives considered
(a) Globalna `FallbackPolicy = RequireAppOperator` w `Program.cs` — odrzucone: objęłaby też `HealthController` (anonimowy) i wymagałaby jawnych wyjątków; zmiana szersza niż problem. (b) `[Authorize(Roles="Operator,Administrator")]` na metodach — odrzucone: duplikacja i ryzyko pominięcia nowej akcji; polityka klasowa domyka wszystkie akcje kontrolera. (c) Tylko front-guard/ukrycie przycisków — odrzucone: nie jest autoryzacją (obejście przez bezpośrednie API).

## ADR-0023: Wiele sekcji DOCX przez markery HTML (`div.docx-section-break`), pierwsza sekcja jako bazowa — 2026-07-05

- Date: 2026-07-05
- Status: Accepted; realizuje kierunek R-10 z DOCX_CONVERSION.md bez zmiany kontraktów API (alternatywa dla rozszerzania `DocumentContent` o kolekcję sekcji).

### Context
`DocumentContent` reprezentuje dokument jako jedną sekcję. Reader brał `Body.Elements<SectionProperties>().FirstOrDefault()` — a to w OOXML sectPr **ostatniej** sekcji (wcześniejsze sekcje kończą się paragrafem z `pPr/sectPr`). Skutki: (1) dokument pionowy z poziomym aneksem otwierał się w geometrii i z nagłówkami aneksu; (2) przerwy sekcji nie łamały strony w edytorze; (3) writer emitował wyłącznie body-level sectPr, więc autosave **spłaszczał dokument wielosekcyjny do jednej sekcji** — trwała utrata orientacji/marginesów sekcji.

### Decision
- Semantyka pól `DocumentContent.PageSize/Margins/Header/Footer` = **pierwsza** sekcja (to widzi użytkownik na stronie 1); nagłówek/stopka: pierwsza sekcja z referencją Default w kolejności dokumentu (fallback do `HeaderParts.FirstOrDefault()` bez zmian).
- Dane pozostałych sekcji jadą **w polu `Html`**: paragraph-level `pPr/sectPr` → `div.page-break` (gdy przerwa łamie stronę) + niewidoczny `div.docx-section-break` z geometrią sekcji NASTĘPNEJ w `data-*` (break-type, rozmiar/orientacja, marginesy, dystanse header/footer; cm, InvariantCulture). Marker jest osobnym elementem, bo splitter stron GUI kanonizuje divy `page-break` i zgubiłby data-*.
- Writer: marker → `w:p/pPr/sectPr` z geometrią sekcji ZAMYKANEJ (tak koduje OOXML; `w:type` przerwy należy do sekcji następnej); body-level sectPr = ostatnia sekcja; `div.page-break` bezpośrednio przed markerem nie emituje `w:br type=page` (sectPr nextPage sam łamie stronę); referencje header/footer + `titlePg` na PIERWSZYM sectPr (Word dziedziczy je na kolejne sekcje bez własnych referencji).
- GUI traktuje marker jako niewidoczny nośnik danych (`display:none` w `wysiwyg-editor.scss`); przechodzi przez split/getContent bez zmian w logice paginacji.

### Consequences
Round-trip dokumentów wielosekcyjnych zachowuje sekcje, orientacje, rozmiary i marginesy (testy `MultiSectionFidelityTests` 12/12); kontrakty REST i model TS bez zmian. Ograniczenia: edytor renderuje wszystkie strony w geometrii pierwszej sekcji (per-page rendering = przyszła zmiana paginacji `wysiwyg-editor` czytająca markery); markery mogą zostać usunięte przez użytkownika edytującego wokół granicy sekcji (degradacja = powrót do jednej sekcji, jak dotąd); sectPr w elementach listy nie emituje markera (rzadkie).

### Alternatives considered
(a) Kolekcja `Sections` w `DocumentContent` + DTO + model TS + rendering per sekcja — odrzucone na ten krok: zmiana kontraktu API i dużej powierzchni frontu; markery dają zachowanie danych od razu i są kompatybilne z przyszłym modelem. (b) Klasa `page-break` na markerze sekcji — odrzucone: splitter GUI zastępuje takie divy czystym markerem i gubi data-*. (c) Pass-through całego body z oryginału — odrzucone: body niesie zmiany użytkownika.

## ADR-0025: Nagłówki/stopki per sekcja przez `SectionHeadersFooters` + dynamiczne pasmo + tab-stopy pozycyjne — 2026-07-05

> Uwaga: pierwotnie zapisane jako drugi „ADR-0024" (numer zajęty przez centralne polityki security). Przenumerowane na ADR-0025 2026-07-05.

- Date: 2026-07-05
- Status: Accepted; domyka R-10 (warstwa nagłówków/stopek wielosekcyjnych) i „częściowe" pozycje tab-stopów z DOCX_CONVERSION; rozszerza ADR-0023.

### Context
Po ADR-0023 geometria sekcji była zachowywana, ale model niósł JEDEN komplet nagłówków/stopek — dokument z innym nagłówkiem w każdej sekcji pokazywał i zapisywał tylko pierwszy zestaw. Pasmo nagłówka rysowało się od samej krawędzi strony (Word zaczyna je na wysokości w:pgMar header) i miało stałą wysokość — treść wyższa niż pasmo nachodziła na body, a paginacja jej nie widziała. Tab-stopy (klasyczny układ „lewa⇥środek⇥prawa") były przybliżane flexem 50%/100% szerokości, a przy zapisie pozycje per akapit ginęły (tylko sztywne 4536/9072 w stylach Header/Footer); literalny 	 w tekście trafiał do w:t, którego Word nie renderuje.

### Decision
- **Model (kontrakt, pole opcjonalne):** `SectionHeaderFooter { SectionIndex, Header, Footer }`; `DocumentContent.SectionHeadersFooters` i `SaveDocumentRequest.SectionHeadersFooters` (Domain + model TS). Wpis tylko dla sekcji ≥ 1 z WŁASNYMI referencjami; sekcja bez wpisu dziedziczy z poprzedniej (semantyka Worda, rozwiązywana we froncie po max indeksie ≤ sekcji strony). Sekcja 0 zostaje w `Header`/`Footer` (kompatybilność wstecz).
- **Writer:** wpis sekcyjny → HeaderPart/FooterPart z referencją na sectPr SWOJEJ sekcji (`_emittedSectionProps[s]`; ostatnia sekcja = body-level). Handlery Save/DownloadEdited przenoszą pole; Sign świadomie nie.
- **Dynamiczne pasmo (GUI):** `.page-header` z `margin-top = headerDistance` i `min-height = margines − dystans`; treść wyższa spycha body (flex); `_repaginateNow` mierzy realne pasma z DOM i odejmuje je od dostępnej wysokości strony. Dystanse per sekcja z data-* markera (`PageGeometry.headerDistanceCm/footerDistanceCm`); sekcja 1 = odwrotność wzoru readera.
- **Tab-stopy:** reader emituje `data-tab-stops="pos:align[:leader]"` (efektywne stopy: łańcuch stylów + direct, semantyka clear) i w nagłówku/stopce renderuje segmenty pozycyjnie (`span.docx-tab-seg`; center=translateX(-50%), right=translateX(-100%) NA pozycji stopu); writer odtwarza `w:tabs` per akapit i emituje `w:tab` z segmentów oraz z literalnych 	 w tekście.

### Consequences
Dokumenty wielosekcyjne pokazują i round-tripują właściwe nagłówki/stopki per sekcja (testy `MultiSectionFidelityTests` +4, GUI +6); nagłówki „lewa⇥środek⇥prawa" lądują na pozycjach Worda i wracają do DOCX bez utraty pozycji (`TabStopFidelityTests` 7); wysoki nagłówek odbiera miejsce treści zamiast rozjeżdżać strony. Ograniczenia: przypisanie k-ty tab → k-ty stop (bez pełnej reguły „następny stop za bieżącym x"); leader nierysowany w edytorze; akapity ze złożonymi polami zostają na flexie; edycja pasma sekcji wymaga kliknięcia na stronie tej sekcji.

### Alternatives considered
(a) Nagłówki sekcji w data-* markera sekcji — odrzucone: HTML nagłówka w atrybucie jest kruchy i nieedytowalny. (b) Grid CSS dla tab-stopów — odrzucone: center-NA-pozycji nie mapuje się na wyrównanie w komórce gridu. (c) Mierzenie szerokości tekstu w backendzie dla tabów — odrzucone: zależne od fontu przeglądarki.

## ADR-0026: Style tabel rozwiązywane do inline CSS + referencja w data-* (kontrolowane przybliżenie) — 2026-07-05

- Date: 2026-07-05
- Status: Accepted

### Context
Reader ignorował `w:tblStyle` — a większość tabel Worda („Tabela – Siatka", style z akcentami)
ma obramowania/cieniowanie/marginesy w STYLU tabeli (styles.xml), nie w bezpośrednim tblPr.
Tabele renderowały się w edytorze bez linii i teł, marker `data-no-borders` utrwalał stratę na
eksporcie, a zewnętrzne `tblBorders` szły na wewnętrzne krawędzie komórek (brak rozróżnienia
outer vs `insideH`/`insideV`). Dodatkowo: `themeFill`/wzory `pctNN` ginęły, `min-height` na `tr`
nie działa w CSS (atLeast traciło wysokość), grubość linii sz/8+min 1px rosła przy każdym
zapisie, a GUI zapiekało zmierzone wysokości WSZYSTKICH wierszy do DOCX (R-18).

### Decision
1. Reader rozwiązuje pełny łańcuch stylu tabeli (`tblStyle` → `basedOn`, per-side merge) +
   formaty warunkowe `tblStylePr` (firstRow/lastRow/firstCol/lastCol/pasy wg flag `tblLook`,
   banding pomija wiersz nagłówkowy jak Word) i **zapisuje rozwiązane wartości w inline CSS**
   komórek (pozycyjnie: krawędź zewnętrzna vs insideH/V). Kolory motywu (`themeFill`/`themeColor`
   + tint/shade) i wzory `pctNN` są rozwiązywane do hex (blend).
2. Referencja stylu jedzie w `data-tbl-style` + `data-tbl-look`; writer re-emituje
   `w:tblStyle`/`w:tblLook`. Po round-tripie formatowanie pochodzące ze stylu staje się
   formatowaniem BEZPOŚREDNIM (wygląd identyczny w Wordzie; styl dalej podpięty).
3. Wysokości wierszy: `height:` na tr (min-height nie działa) + `data-row-height-tw`/
   `data-row-hrule` (writer preferuje twips/regułę z data-*); GUI nie zapieka mierzonych
   wysokości (tylko jawne inline z importu/ręcznego resize; resize czyści data-*).
4. Grubość linii: sz/6 px przy odczycie i px×6 przy zapisie (symetria, koniec pogrubiania).
5. `tblHeader`/`cantSplit`/`tblCellSpacing`/`tblInd` — round-trip (data-* / CSS), bez
   egzekwowania w paginacji edytora.
6. GUI synchronizuje `<colgroup>` (źródło `w:tblGrid`) po resize i wstawieniu/usunięciu kolumny
   (`table-grid.util.syncTableColgroup`).

### Consequences
Tabele stylowane wyglądają w edytorze jak w Wordzie i nie tracą wyglądu na eksporcie; dokument
nie puchnie od sztucznych trHeight. Koszt: po edycji w naszym edytorze zmiana definicji stylu
w Wordzie nie zmieni już wyglądu tabeli (formatowanie bezpośrednie wygrywa) — świadomy trade-off.
NIE obsługujemy: warunkowego formatowania TEKSTU ze stylu (bold nagłówka), regionów narożnych,
przekątnych `tl2br`/`tr2bl`, `fitText`/`hideMark`, powtarzania wiersza nagłówkowego w paginacji
edytora.

### Alternatives considered
Emisja klas CSS + arkusza stylów per dokument (odrzucone: kontrakt edytora opiera się na inline
styles — przeżywa split/merge stron i sanitizację). Pełne zachowanie rozdziału styl/direct przy
eksporcie (odrzucone: wymagałoby śledzenia pochodzenia każdej właściwości w HTML; niewspółmierna
złożoność do zysku).

## ADR-0027: Własny pure-managed tłumacz wektorowy EMF/WMF → SVG (etap 1, tylko podgląd) — 2026-07-05

- Date: 2026-07-05
- Status: Accepted

### Context
Czysto wektorowe EMF/WMF (bez osadzonego rastra) renderowały się w edytorze jako przezroczysty
blank (ADR-0020) — dokument działał, ale użytkownik nie widział grafiki do czasu otwarcia w Word.
Rasteryzacja przez LibreOffice/GDI/System.Drawing jest zakazana (Windows-only / proces zewnętrzny),
Magick.NET deleguje metafile do natywnych delegatów nieobecnych w kontenerze, a roadmapowy sidecar
to osobna infrastruktura. Brak gotowej, permisywnej, pure-managed biblioteki renderującej EMF/WMF.

### Decision
1. Nowy `MetafileVectorTranslator` (Infrastructure, internal): własny parser rekordów binarnych
   MS-EMF / MS-WMF tłumaczący BEZPOŚREDNIO na SVG podzbiór GDI etapu 1: pióra/pędzle (kolor,
   grubość, dash, stock objects), MoveTo/LineTo, Rectangle/Ellipse/RoundRect,
   Polygon/Polyline/PolyBezier (warianty 16/32-bit), PolyPolygon (fill-rule),
   ścieżki (BeginPath/CloseFigure/Fill/Stroke/StrokeAndFill), SetWorldTransform/
   ModifyWorldTransform (macierz 2×3, rotacja → polygon), SetPolyFillMode,
   StretchDIBits → `<image>` (DIB przez istniejące SliceDib/DibToPng + Skia).
   WMF: tabela obiektów slotowa (fonty/palety/regiony zajmują sloty), placeable bbox /
   SETWINDOWORG/EXT jako viewBox, kolejność parametrów odwrócona (spec).
2. Wpięty jako strategia `vector-translate` w łańcuch `ConvertMetafile`:
   `embedded-raster` → `dib-rasterize` → **`vector-translate`** → blank. Wynik przechodzi przez
   `SanitizeSvg`; Status=Converted, Fidelity=Lossy; rekordy spoza podzbioru liczone i raportowane
   w Warnings/LostProperties (rekordy czysto stanowe pomijane bez liczenia).
3. Tłumaczenie jest WYŁĄCZNIE podglądem: oryginalny metafile nadal jedzie do DOCX
   (data-original-src / pass-through) — writer nigdy nie zapisuje SVG jako blip.
   `LoadImageFromPart` nie podmienia bajtów partu na SVG (podgląd powstaje w renderze, z cache).
4. Limity niezaufanego wejścia: 200k rekordów, 100k punktów/poly, 20k elementów SVG, 2 MB
   wyjścia; każdy wyjątek → null → łańcuch przechodzi na blank.

### Consequences
Wektorowe EMF/WMF (ramki, schematy, loga, pieczątki, proste cliparty) są wreszcie WIDOCZNE
w edytorze — z poprawnymi kolorami, grubościami linii i transformacjami; brak treści możliwej
do przetłumaczenia degraduje się do dotychczasowego blanku (zero regresji — stare testy blank
przechodzą bez zmian). Poza zakresem etapu 1 (świadomie): tekst (ExtTextOut — wymaga metryk
fontów), clipping, ROP-y rastrowe, pędzle wzorkowe, rekordy EMF+ (GDI+; pliki dual niosą
fallback EMF, który tłumaczymy). Testy: `MetafileVectorTranslationTests` 10/10 (EMF rect/polygon/
line/path/transform/WMF/integracja DOCX z eksportem oryginału), Infrastructure 285/285.

### Alternatives considered
(a) Sidecar rasteryzujący w osobnym kontenerze — odrzucone na teraz: infrastruktura + latencja;
pozostaje opcją dla pełnej wierności (EMF+/tekst). (b) Magick.NET / natywne delegaty — odrzucone:
niedeterministyczne w kontenerze (ADR-0020 §4). (c) Rasteryzacja SVG→PNG po stronie serwera —
zbędna: przeglądarka renderuje SVG natywnie, wektor skaluje się lepiej.
