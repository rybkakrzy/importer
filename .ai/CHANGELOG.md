# AI-Assisted Changelog

Istotne zmiany dla kontynuacji pracy (nie zastępuje changeloga produktu).

## Format

```md
## YYYY-MM-DD — <tytuł>
### Changed
### Verified
### Notes
```

## 2026-07-12 — Listy DOCX etap 1 wariantu A (ADR-0036): eksport restartów, punktatorów graficznych i rzadkich właściwości poziomu

### Changed
- `DocxToHtmlConverter`: kontrakt `data-*` list rozszerzony o `data-start-override` (w:startOverride
  instancji — emitowany ODDZIELNIE od `data-start`, które teraz niesie zawsze w:start DEFINICJI),
  `data-suffix` (w:suff space/nothing), `data-is-legal` (w:isLgl), `data-lvl-restart` (surowa
  jednobazowa wartość w:lvlRestart), `data-pic-bullet` (poziom z lvlPicBulletId),
  `data-ind-left/hanging/first-line-tw` (wcięcia definicji poziomu w twips).
- `HtmlToDocxConverter`: w:abstractNum współdzielony po `data-abstract-num-id`; restart =
  `w:lvlOverride/w:startOverride` na nowej instancji (wcześniej writer w ogóle ich nie emitował —
  „Rozpocznij od nowa" ginęło przy pierwszym autosave); `TryCreatePictureBullet` odtwarza
  `w:numPicBullet` (ImagePart na części numeracji + VML v:shape/v:imagedata + w:lvlPicBulletId,
  deduplikacja po data URI) zamiast bake'ować obraz w treść runu; suffix/isLgl/lvlRestart/wcięcia
  z data-*; kolejność dzieci w:lvl zgodna z sekwencją CT_Lvl (rPr na końcu — było przed lvlJc);
  stan numeracji (`_numberingId`, `_numberingPart`, mapy) resetowany per Convert.

- Runda 2 (weryfikacja kontynuacji/wcięć/rodzaju numeracji): `ResolveListLevel` — writer honoruje
  `data-ilvl` (fragment kontynuacji na głębszym poziomie nie jest spłaszczany do ilvl=0);
  pełne `w:lvlOverride/w:lvl` na instancji przez `data-lvl-override` (reader znaczy poziomy
  z definicją z lvlOverride instancji — wygląd per instancja nie zlewa się we wspólnym
  abstrakcie); `UpgradeSharedAbstractLevels` (późniejszy fragment dosyła definicje poziomów,
  których fragment tworzący abstrakt nie używał). Wydzielone `BuildAbstractLevel`.

- GUI (prezentacja): nowy `core/utils/list-label.util.ts` — silnik etykiet TS lustrzany do
  liczników readera (szablony %N z formatem poziomu odwołania, isLegal, suffix, lvlRestart=0/N,
  startOverride, kontynuacja fragmentów przez strony); `wysiwyg-editor.refreshListLabels()`
  (hooki: content-input/setContent/persist/undo/redo/pageEditorRefs.changes) nadaje
  `data-list-label` na li, render CSS `li[data-list-label]::before` (etykieta poza edytowalnym
  tekstem, nie kopiuje się); `stripListLabelAttributes` w serializacji — zapis bez atrybutów
  prezentacyjnych.

### Verified
- `ListNumberingFidelityTests` 21/21 (+9: start-override osobno od start, lvlOverride na wspólnym
  abstrakcie, pełny round-trip restartu, suffix/isLgl/lvlRestart/wcięcia, punktator graficzny
  z bajtami 1:1 i bez inline Drawing, walidator OOXML Office2013 = 0 błędów, fragment ilvl=1,
  upgrade poziomów wspólnego abstraktu, pełny lvlOverride zostaje na instancji).
- Infrastructure.UnitTests 421/421 (0 fail, goldeny bez zmian), build solucji 0 błędów.
- GUI: `list-label.util.spec` 12 + `wysiwyg-editor.list-labels.spec` 5, pełne Vitest 415/415,
  `ng build` OK; wizualna weryfikacja headless Chrome+CDP (skill verify) na żywym edytorze:
  1. / 2. / 2.a) / 2.b) / 3. (kontynuacja po akapicie) / 1. (startOverride), zapis czysty.
- Runda 4 (katalog wymagań użytkownika): fix readera — WIELOZNAKOWY lvlText punktatora
  ("TODO:", "Pkt", "§ 1", "o czym mowa") renderuje się w całości (wcześniej tylko pierwszy
  znak przez MapBulletChar; skróty disc/circle/square ograniczone do jednoznakowego lvlText).
  Testy: `MultilevelSchemes_FormatsAndTemplates_SurviveRoundTrip` (1.1.1. / 1.1.1) / A.1.a. /
  I.A.1. / § 1.1.1), `UnicodeBulletCatalog_LvlTextSurvivesRoundTrip` (~70 znaków: kropki/koła,
  kwadraty, trójkąty/strzałki, myślniki, gwiazdki, romby, znaczniki wyboru),
  `TextMarkerBullets_RenderFullTextAndRoundTrip`; TS: schematy + numeratory z literałami
  ((%1) / [%1] / Krok %1 / Pkt %1). Infrastructure 424/424, GUI Vitest 417/417.
- Runda 5 (ENTER w listach + eksport po edycjach — weryfikacja EMPIRYCZNA na żywym edytorze,
  headless Chrome + realne zdarzenia klawiszy): (a) Enter na końcu/w środku li → nowy element
  tej samej listy, etykiety (w tym DALSZE fragmenty) przeliczają się po debounce; (b) 2×Enter
  (wyjście z listy) → Chrome dzieli <ol> KOPIUJĄC kontrakt data-* na drugi fragment, akapit
  ląduje jako <div>; (c) nowy test `Writer_EditorHtmlAfterEnterAndListExit_KeepsOneLogicalList`
  na DOSŁOWNYM HTML przechwyconym z edytora: fragmenty scalone do JEDNEJ instancji, div →
  zwykły akapit, kolejność treści, walidator 0 błędów, reimport z kontynuacją start=3/4;
  (d) fix: `ensureBulletMarkers` (list-label.util, wołane w refreshListLabels) — li utworzony
  Enterem w liście z marker spanem (własny symbol/„TODO:"/obraz) nie dziedziczył znacznika;
  uzupełniany klonem z rodzeństwa (add-only, bez ryzyka dla kotwicy kursora); potwierdzone
  na żywo (TODO:/✔ — dokładnie 1 marker per li, zapis bez atrybutów prezentacyjnych).
  Infrastructure 425/425, GUI Vitest 419/419.
- Runda 6 (naprawy z audytu, R-30 częściowo): (1) **Listy w komórkach tabel przestały ginąć** —
  `AppendTableCellHtml` grupuje akapity listowe przez `ConvertConsecutiveListItems` (jak body/
  SdtBlock); li w komórce dostaje też inline spacing docDefaults/stylu tabeli (ADR-0031) przez
  `_tableParagraphDefaultCss` + `DeduplicateCss`. Kontynuacja numeracji działa przez granicę
  komórek (wspólny numId). (2) **Nieznane w:numFmt nie degradują się do decimal** (pkt 22.10):
  `NumFmtToken` niesie surowy token (`Val.InnerText`) — ordinal/cardinalText/ordinalText/
  chicago/językowe; `TryMapNumFmt` odtwarza token 1:1 (`new NumberFormatValues(token)`,
  guard regex na kształt). (3) Przy okazji pre-existing bug writera: kolejność dzieci
  `w:tblBorders` była top→bottom→left→right — schemat CT_TblBorders wymaga top→left→bottom→
  right (błąd walidacji OOXML przy KAŻDEJ tabeli bez data-tbl-style; wykryty, bo nowy test
  jako pierwszy walidował eksport tabeli). Testy: `ListsInTableCells_GroupIntoOlAndSurviveRoundTrip`,
  `UnknownNumFmt_RoundTripsRawToken…`; Infrastructure 427/427, build sln 0 błędów.

### Notes
- Decyzja architektoniczna (wariant A vs B) i pełna lista konsekwencji: ADR-0036.
- GUI i kontrakty API nietknięte; następne etapy w TASK_HANDOFF (silnik etykiet, komendy edytora,
  lvlRestart=N w licznikach, łańcuch basedOn/styleLink, diagnostyka LIST_*).

## 2026-07-11 — Fix: przyciski wyrównania/list nie pokazują stanu aktywnego (toolbar główny, mini-toolbar, menu)

### Changed
- `TextFormatting` (+`alignment?`/`bulletList?`/`numberedList?`): stan formatowania akapitu wreszcie
  jest częścią kontraktu `EditorState` — wcześniej model niósł tylko bold/italic/underline/strike/sub/sup,
  więc toolbar fizycznie nie miał czego podświetlić.
- `wysiwyg-editor.updateFormattingState`: wyrównanie liczone z computed `text-align` bloku pod karetką
  (pokrywa inline z `execCommand justify*` ORAZ wartości z importu DOCX; `start`/brak → `left` jak
  w Wordzie), listy z przodka `li` (UL→bullet, OL→numbered). Odpala się na selectionchange — stan
  jedzie za karetką bez dodatkowych zdarzeń.
- `editor-toolbar`: `isAlignActive()` (dokładnie jeden aktywny, domyślnie left) + `[class.active]`
  na 4 przyciskach wyrównania i 2 list; `isActive()` utwardzone (`=== true`, bo model ma teraz
  pole nie-boolean).
- `document-editor`: `alignmentActive()`; mini-toolbar `[class.mt-btn-active]` na wyrównaniu
  i listach; menu „Formatuj→Wyrównanie" i kontekstowe podmenu wyrównania dostają klasę
  `dropdown-item-checked`/`context-menu-item-checked` (nowe style — paleta jak `.toolbar-btn.active`).

### Verified
- Nowy `wysiwyg-editor.formatting-state.spec.ts` 7/7 (left domyślne, center/right/justify, ul/ol,
  brak listy, toolbar: jeden aktywny + null-state). Pełne GUI **396/396** (0 fail), `ng build` OK.

### Notes
- Word-parity: „do lewej" aktywne także bez jawnego text-align (default akapitu).

## 2026-07-11 — Fix (UAT): twarda spacja na początku każdej komórki tabeli

### Changed
- `wysiwyg-editor.insertTable` + operacje tabelowe w `document-editor` (wstaw wiersz/kolumnę,
  podziel/scal komórki): placeholder pustej komórki zmieniony z `&nbsp;` na `<br>` —
  `&nbsp;` zostawał przed wpisanym tekstem (przesunięcie względem MS Word) i trafiał jako
  U+00A0 do zapisanego DOCX. Goły `<br>` w `td` eksportuje się do pustego akapitu
  (`HtmlToDocxConverter.ConvertHtmlNode` case "br" → `new Paragraph()`).
- Nowy handler `onEditorBeforeInput` (bindowany na `beforeinput` stron + nagłówka/stopki):
  gdy blok (p/h1-6/li/td/th) zawiera wyłącznie U+00A0, zaznacza placeholder tuż przed
  wstawieniem tekstu, więc pisanie go zastępuje. Obejmuje komórki/akapity z importu DOCX
  (`<td><p>&nbsp;</p></td>`) i dokumenty zapisane przed poprawką.
- Separatory-akapity po tabeli zostają z `&nbsp;` celowo — `<p><br></p>` eksportowałby się
  do akapitu z `w:br` (dodatkowa linia w Wordzie); czyszczenie przy pisaniu robi beforeinput.

### Verified
- Nowy spec `wysiwyg-editor.table-cell-placeholder.spec.ts` (3 testy: placeholder `<br>`,
  zaznaczenie nbsp przez beforeinput, brak ingerencji przy prawdziwej treści).
- Pełny frontend: 389/389 testów, 42 pliki (`npx ng test --watch=false`).

### Notes
- W `wysiwyg-editor.ts` jest literalny bajt NUL w kluczu cache przypisów
  (`${fn.id}\0${fn.html}`) — przez to ripgrep traktuje plik jako binarny; grep po tym
  pliku robić przez Select-String.

## 2026-07-11 — Fix: dokument Word „tylko do odczytu" był edytowalny w DOC2 Editor

### Changed
- `DocxToHtmlConverter`: nowa detekcja `HasEnforcedEditProtection` (settings.xml) — wymuszone
  `w:documentProtection` (enforcement=1, edit≠none; tryby częściowe comments/forms/trackedChanges
  też liczone jako ochrona) oraz `w:writeProtection` (w:recommended lub hasło zapisu w:hash/
  w:hashValue). Wynik w nowym polu `DocumentContent.IsReadOnlyProtected` (Domain + model TS).
- GUI `document-editor`: nowy sygnał `documentEditProtected` ustawiany w `_applyLoadedContent`
  z flagi konwersji (samoresetujący przy kolejnym dokumencie); `editingDisabled` rozszerzone
  o ten stan; badge „Tylko do odczytu — dokument chroniony przed edycją"; toast informacyjny
  przy otwarciu; guard `saveDocument()` zmieniony z `readOnly()` na `editingDisabled()`;
  dodatkowy guard w ticku auto-save; przełącznik Autozapisu ukrywany dla chronionych dokumentów.
- GUI: `d2-wysiwyg-editor [readOnly]` podpięty pod `editingDisabled()` zamiast surowego
  `readOnly()` — domyka to też lukę `lockedByOther` (treść była edytowalna mimo ukrytego toolbara).
- `layout-shell.spec.ts`: dodane stuby `ResourceAccessService`/`MsalService` (pre-existing fail
  NG0201 po tym, jak Dashboard zaczął wstrzykiwać gating zasobów — niezwiązany z tą zmianą).

### Verified
- Nowe `DocumentProtectionImportTests` 8/8 (enforced readOnly/comments, brak enforcement,
  edit=none, writeProtection recommended/hash, puste settings, brak części settings).
- Front: 3 nowe testy w `document-editor.spec.ts` (blokada edycji + toast, blokada zapisu,
  reset flagi na kolejnym dokumencie). Cała suita **386/386**, backend Infrastructure zielony.

### Notes
- Egzekwowanie WYŁĄCZNIE po stronie GUI — patrz RISKS: endpoint zapisu nadal przyjmie PUT,
  jeśli klient zignoruje flagę. Haseł ochrony nie weryfikujemy (dokument z hasłem zapisu
  jest zawsze tylko-do-odczytu w edytorze).

## 2026-07-11 — Fix: znaki specjalne z fontów symbolicznych (strzałka →) renderowane jako kwadrat/tofu

### Changed
- `DocxToHtmlConverter`: `w:sym` przestaje być emitowany jako goła encja PUA (`&#xF0E0;` bez fontu =
  tofu w przeglądarce). Nowy `ConvertSymbolCharToHtml` + `TryMapSymbolicChar` + tabele
  `SymbolFontMap` (greka, operatory, strzałki — kodowanie Adobe) i `WingdingsFontMap` (strzałki,
  checkboxy, kształty): kod (po zdjęciu przesunięcia PUA U+F000..U+F0FF) mapowany na odpowiednik
  Unicode — renderuje się wszędzie i round-tripuje jako zwykły tekst. Kod spoza tabel → encja
  w spanie z `font-family` fontu symbolicznego (fonty są na Windows; writer odtwarza rFonts z CSS).
  `w:sym` ze zwykłym fontem i normalnym code-pointem → znak wprost.
- `MapSymbolicTextRun`: te same mapowania dla ZWYKŁEGO `w:t` — Word zapisuje symbole także jako
  znaki PUA / znaki bajtowe w runie z fontem symbolicznym (autokorekta „-->" = literalne `è`
  w foncie Wingdings). PUA bez fontu symbolicznego nietykane (nie zgadujemy glifu).
  `NormalizeSymbolFontName` dopasowuje nazwę DOKŁADNIE (symbol/wingdings/wingdings 2/3/webdings) —
  „Segoe UI Symbol" to normalny font Unicode i NIE podlega mapowaniu po młodszym bajcie.

### Verified
- Nowe `SymbolCharFidelityTests` 10/10 (w:sym Symbol/Wingdings/PUA i bez przesunięcia, zwykły font,
  fallback span+font dla nieznanego kodu, w:t z PUA i znakiem bajtowym, ochrona Segoe UI Symbol,
  czysty Unicode bez zmian). Infrastructure **412/412**, build solucji 0 błędów, goldeny nietknięte.

### Notes
- `MapBulletChar` (punktatory list) celowo NIE ruszany — ma własną semantykę fallbacku (•).
- Eksport: zmapowany znak wraca jako zwykły tekst w `w:t` (Word renderuje poprawnie); fallback
  span+PUA wraca jako run z rFonts fontu symbolicznego.

## 2026-07-11 — Fix: kursor skacze na początek dokumentu podczas pisania (rebind stron przy każdej repaginacji)

### Changed
- `wysiwyg-editor.ts` `_repaginateNow`: decyzja o rebindzie `[innerHTML]` stron porównuje nowy rozkład
  bloków z **żywym DOM** (serializacja per strona tym samym mechanizmem co `newPageContents`), a nie
  z sygnałem `pageContents` — sygnał jest celowo przestarzały między repaginacjami (`onPageInput` go
  nie aktualizuje), więc stary check `identical` był przy pisaniu zawsze false i KAŻDA repaginacja
  (max-wait 600 ms) wymieniała DOM wszystkich stron: selekcja ginęła, kursor na ułamek sekundy spadał
  na początek contenteditable, a znaki wpisane w oknie rebind→restore lądowały w złym miejscu lub
  ginęły ze starym DOM. Dodatkowy warunek: liczba stron zgodna z sygnałem (sygnał steruje `@for`).
  Pominięcie `set()` jest bezpieczne — `getContent()` czyta żywy DOM.
- Restore karetki po realnym rebindzie: `afterNextRender(..., { injector })` zamiast `setTimeout(0)`
  — odtworzenie selekcji zaraz po renderze zwęża okno, w którym klawisz trafia w zresetowaną karetkę
  (przypadek prawdziwego przelania treści na inną stronę).

### Verified
- `wysiwyg-editor.pagination-overflow.spec.ts` +4 (18/18): pisanie bez przelania = zero rebindów,
  przelanie nadal rebinduje, opróżniona strona odzyskuje syntetyczny `<p></p>`, rozjazd liczby stron
  sygnał↔DOM wymusza rebind (scalanie w górę). Pełna suita GUI 382/383 (fail = pre-existing layout-shell).

### Notes
- Sygnał `pageContents` pozostaje przestarzały do czasu realnej zmiany rozkladu — to istniejący kontrakt
  (DOM jest źródłem prawdy; serializacja w `getContent`). Skip nie dotyka `pageGeometries`/`pageSectionIndexes`.

## 2026-07-11 — Fix: custom-geometry logo (custGeom) renderowane jako czarny blob zamiast grafiki

### Changed
- `DocxToHtmlConverter.GetShapeFillHex`: rozwiązuje fill kształtu z `a:solidFill` (`a:srgbClr` **lub**
  `a:schemeClr`) oraz z referencji stylu `wps:style/a:fillRef`. Nowy `ResolveDrawingSchemeColor` mapuje
  DrawingML `a:schemeClr` (dk1/lt1/dk2/lt2/tx1/bg1/tx2/bg2/accent1..6/hlink/folHlink) na hex z theme1.xml.
- `BuildCustomGeometrySvg`: fallback nierozwiązywalnego wypełnienia z `#000000` → `currentColor`
  (dziedziczy kolor tekstu otoczenia) — koniec „czarnego knefla" zasłaniającego logo. Patrz ADR-0033.

### Verified
- `Doc2ImportFidelityTests`: 30/30 (3 nowe: schemeClr→motyw, fillRef→motyw, brak fill→brak `#000000`).

### Notes
- Tylko podgląd (reader); writer nie odtwarza tych kształtów do DOCX. Dokument źródłowy (Qutalo
  „Pamiętniczek_V3") niedostępny — diagnoza z wyrenderowanego HTML; naprawa pokrywa wszystkie
  prawdopodobne źródła fill niezależnie od wariantu.

## 2026-07-11 — Fix: obraz w pozycjonowanym segmencie tab-stopu renderuje się jako pionowy pasek

### Changed
- `wysiwyg-editor.scss`: nowa reguła `.docx-tab-seg .editor-image-wrapper, .docx-tab-seg img { max-width: none !important }`.
  `.docx-tab-seg` jest `position:absolute` (szerokość shrink-to-fit); inline `max-width:100%` na `<img>`
  (importer) i na wrapperze (`wrapExistingImages`) tworzy cykl rozmiaru — wkład obrazka do intrinsic width
  kontenera liczy się jako 0, kontener zapada się i obraz (np. logo w nagłówku wyrównane tabem do prawej)
  renderował się jako pasek ~0px szerokości × inline-height. `!important` konieczny, bo max-width:100% jest inline.
  Bezpieczne: obrazy z importera zawsze niosą jawne `width`/`height` w px.

### Verified
- `npx sass` kompiluje plik bez błędów; zmiana czysto CSS, brak wpływu na round-trip (writer czyta width/height/EMU, nie max-width).

### Notes
- `data-width-emu=575945` (≈1.52 cm ≈ 60px) to rozmiar zadeklarowany w DOCX (wp:extent), nie intrinsic PNG (165×165) — to poprawne.

## 2026-07-10 — Fix: tabela traci obramowania po zapisie (writer nie parsował rozbitych border-* / rgb)

### Changed
- `HtmlToDocxConverter.ApplyCellBorders`: rozpoznaje teraz obramowanie komórki zapisane jako OSOBNE
  właściwości `border-width` + `border-style` + `border-color` (tak przeglądarka serializuje jednolite
  obramowanie `<td>` przy getContent/outerHTML) oraz kolory `rgb()`/`rgba()` (przeglądarka normalizuje
  hex→rgb przy edycji). Wcześniej parser akceptował tylko formę skróconą `border-top: .. #hex`, więc po
  PIERWSZYM zapisie w edytorze `w:tcBorders` w ogóle nie powstawało → komórki traciły wszystkie linie i
  tabela „rozpadała się" wizualnie (wszystkie krawędzie `none`, tabela znaczona `data-no-borders`).
- Nowe helpery: `TryParseBorderShorthand` (per-strona/`border:`, hex+rgb), `GetCssDeclarationValue`,
  `NormalizeCssColorToken` (#rgb/#rrggbb/rgb()/rgba() → 6-hex). `border-style:none` w formie rozbitej nie
  emituje obramowań (styl tabeli decyduje). Ścieżka table-level i `data-no-borders` NIE zmieniane.

### Verified
- Harness round-trip (HTML→DOCX→HTML) tabeli „kredyty" ze zrzutu: przed = 14 komórek `border-*:none`;
  po = 14 komórek `border-top:0.7px solid #000000` (linie zachowane). Wariant per-side hex bez regresji.
- `TableStyleFidelityTests` 29/29 (+3), pełne `Infrastructure.UnitTests` 391/391, 0 fail.

### Notes
- `data-no-borders="1"` na TABELI to osobna, wcześniejsza kosmetyka (ADR-0031, tblStyle bez definicji w
  regenerowanym pakiecie) — nieszkodliwa dla gridlines, bo komórki niosą własne linie. Poza zakresem fixa.

## 2026-07-10 — Kompletna obsługa przypisów dolnych DOCX (ADR-0032, +27 testów)

### Changed
- **Model domenowy**: nowy `Footnote { Id, Html }` (`DocumentModels.cs` + `document.model.ts`); `DocumentContent.Footnotes` i `SaveDocumentRequest.Footnotes` (backend+TS). `Id` = STABILNA tożsamość (`fn-<ooxmlId>`), niezależna od numeru widocznego; numer liczony z kolejności pierwszych odwołań; numeryczny `w:id` OOXML przydzielany deterministycznie dopiero na eksporcie. Treść przypisu = jedno źródło prawdy (odwołania niosą tylko `data-footnote-id`).
- **Reader (`DocxToHtmlConverter`)**: `w:footnoteReference` w treści → `<sup class="footnote-ref" data-footnote-id aria-label>N</sup>` (numer wg pierwszego wystąpienia, wspólny dla powtórzonych odwołań); `FootnoteReferenceMark` pomijany; `ExtractFootnotes` czyta `FootnotesPart`, POMIJA separator/continuationSeparator/continuationNotice, konwertuje treść istniejącymi konwerterami akapitu/tabeli (formatowanie/wieloakapitowość/Unicode), buduje `Footnotes` w kolejności odwołań; jeden wadliwy przypis nie przerywa importu (try/catch + `ILogger`); odwołanie bez treści = zachowane z pustą treścią + diagnostyka. Brak odwołań → `Footnotes = null`.
- **Writer (`HtmlToDocxConverter`)**: `AssignFootnoteOoxmlIds` (htmlId→1..N); `CreateFootnoteReferenceRun` (`<sup>`→`w:footnoteReference`; odwołanie do nieistniejącego przypisu POMIJANE); `AddFootnotes` tworzy `word/footnotes.xml` (relacja+content type przez `AddNewPart<FootnotesPart>`) z separatorami technicznymi (id=-1/0) i treścią przypisów (pierwszy akapit dostaje `w:footnoteRef`; treść przez `ConvertHtmlToBody`, relacje obrazów zakresowane do części przypisów). Brak przypisów → brak części. `Convert`/`ConvertPreservingPackage` + `IHtmlToDocxConverter` dostały opcjonalny param `footnotes`.
- **API**: `footnotes` przewleczone przez `SaveDocumentCommand`/handler (`/api/document/save`) i `DownloadEditedDocumentCommand`/handler (`/user-download`) + oba kontrolery.
- **GUI (`wysiwyg-editor`)**: input `footnotes` + output `footnotesChange`; panel treści POZA contenteditable (`.footnotes-panel`, numer wg kolejności, edytowalny, przycisk usuń, aria-label); odwołania w treści strony jako `<sup class="footnote-ref">` (SCSS superscript). Metody `commitFootnoteContent`/`addFootnoteAtCursor`/`removeFootnote`/`syncFootnotesWithBody` (renumeracja wg DOM + przycięcie osieroconych, wpięte w debounce persist). `document-editor`: sygnał `footnotes` z importu → `[footnotes]`/`(footnotesChange)` + `buildSaveRequest`.

### Verified
- Backend `FootnoteFidelityTests` **18/18** (import, eksport z inspekcją ZIP/XML, walidacja OOXML Office2013 = 0 błędów, round-trip semantyczny, add/edit/delete/reorder+renumeracja, regresja bez przypisów), fixture `FootnoteTestDocuments` na czystym OpenXML SDK (niezależny od eksportera). Solucja: Domain 79 / Application 306 / Api 130 / Api.Integration 9 / Infrastructure **388** — 0 fail.
- GUI `wysiwyg-editor.footnotes.spec` **7/7** + `document-editor.footnotes.spec` **2/2**; pełne GUI **378/379** (jedyny fail = pre-existing `spec-layout-shell`). `ng build` + `dotnet build` OK.

### Notes
- Repo bez infrastruktury E2E (tylko Vitest) — headless pełen przepływ pokryty przez TestBed (render+edycja+serializacja) + backend round-trip/OOXML zamiast Playwright (nieproporcjonalne). ADR-0032.
- ID OOXML nie są zachowywane 1:1 — writer przemapowuje deterministycznie (1..N), reader odtwarza `fn-<ooxmlId>`; semantyka (liczba/treść/kolejność/powiązania) zachowana.

## 2026-07-08 — Domyślne odstępy/interlinia dokumentu i tabele przestają się psuć po zapisie (ADR-0031, +15 testów)

### Changed
- **Writer (`HtmlToDocxConverter`) — root cause „tabele po zapisie totalnie się psują"**: `Convert` regenerował styles.xml z HARDKODOWANYMI wartościami (docDefaults 11pt / after=160 / line=259, brak definicji stylów tabel), więc pierwszy zapis dokumentu 12pt z `w:line=278` i tabelą „Tabela – Siatka": zmniejszał cały tekst 12→11pt, zmieniał interlinię, a akapity w komórkach traciły pPr stylu tabeli (after=0/line=240) i dostawały pełne odstępy docDefaults → **każdy wiersz tabeli puchł ~2× w Wordzie**. Nowe: `CaptureDocumentDefaults` czyta z wrappera `.document-content` inline font-family/font-size i `data-default-before/after-tw`, `data-default-line`, `data-default-line-rule`; `AddDocumentStyles` odtwarza z nich docDefaults + Normal (fallback = dotychczasowe wartości konfiguracyjne, zero regresji bez wrappera).
- **Writer — tabele**: (1) tabela z `data-tbl-style` i bez CSS `border:` na `<table>` NIE dostaje już `tblBorders val=none` (bezpośrednie none NADPISYWAŁO obramowania stylu w Wordzie i po ponownym otwarciu utrwalało się jako `data-no-borders`); jawny `data-no-borders="1"` nadal emituje none. (2) `w:tblGrid` i `w:tcW` z DOKŁADNYCH twips `data-w-tw` na `<col>` (suma spanowanych kolumn dla scaleń; szerokości % nietykane) — koniec dryfu siatki o kilka twips przy każdym zapisie (3020→3015→…). (3) Kolejność dzieci wg schematu OOXML: `NormalizeParagraphPropertiesOrder` (spacing/ind/tabs/pStyle PRZED jc — naruszenie ujawniało się, gdy akapit miał jednocześnie wyrównanie i spacing, np. wyśrodkowana komórka), `NormalizeTableCellPropertiesOrder` (tcW przed gridSpan, tcBorders przed tcMar/vAlign), `ApplyCellBorders` emituje top→left→bottom→right. Walidator OOXML (Office2013) na eksporcie qutable: **0 błędów** (było 52).
- **Reader (`DocxToHtmlConverter`)**: (1) czyta `docDefaults/pPrDefault/w:spacing` + nadpisanie z DOMYŚLNEGO stylu akapitowego (`w:default="1"` — typowy Word trzyma spacing w Normal) i emituje na kontenerze `.document-content` atrybuty `data-default-*` oraz `line-height` (interlinia dziedziczy na akapity bez własnej); (2) **pPr stylu tabeli stosowane do akapitów komórek** — `TableStyleContext.ParagraphDefaultCss` (docDefaults + łańcuch `w:style/w:pPr` po basedOn) idzie INLINE na `<p>` w komórkach (`margin-bottom:0pt;line-height:1;` dla Tabela–Siatka), zagnieżdżone tabele przez save/restore — dzięki temu eksport niesie jawny `w:spacing` i wiersze nie puchną nawet w pakiecie bez definicji stylu (częściowe domknięcie luki ADR-0026); (3) `<col>` niesie `data-w-tw` z `w:tblGrid`; (4) finalny CSS akapitu przechodzi `DeduplicateCss` — duplikat właściwości (styl vs direct pPr) był brany przez regexy writera z PIERWSZEGO wystąpienia, więc nadpisanie stylu ginęło na eksporcie.
- **GUI (`wysiwyg-editor`, `table-grid.util`)**: (1) `_captureDocumentDefaults` przechwytuje WSZYSTKIE atrybuty wrappera i rozwija go od razu (strony nigdy nie niosą kontenera), a `getContent()` owija scaloną treść z powrotem — wcześniej wrapper ginął przy paginacji i writer NIGDY nie widział domyślnych wartości dokumentu w produkcyjnym zapisie; (2) nowe sygnały `documentDefaultLineHeight` (binding `line-height` na `.editor-content`) i `documentDefaultParagraphSpacing` (CSS var `--doc-par-margin`; SCSS: `p { margin: 0 0 var(--doc-par-margin, 10px) 0 }`) — edytor renderuje odstępy/interlinię jak Word zamiast sztywnych 10px; (3) `syncTableColgroup` przy zmianie szerokości kolumny usuwa `data-w-tw` (stary atrybut przypiąłby geometrię sprzed resize'u).

### Verified
- Nowe `DocumentDefaultsAndTableSpacingFidelityTests` **11/11** (docDefaults round-trip, pPr stylu tabeli inline, tblBorders-omission, dokładny tblGrid/tcW, 2 cykle bez dryfu, walidator 0 błędów, kolejność pPr); Infrastructure **370/370** (goldeny `simple-table`/`merged-cells-table` zregenerowane — diff tylko `data-w-tw`); build solucji 0 błędów.
- GUI: +4 specy (re-wrap kontenera z atrybutami, capture interlinii/odstępu, higiena `data-w-tw`); pełne **369/370** (jedyny fail = pre-existing `spec-layout-shell`); `ng build` OK.
- Harness round-trip na realnym `qutable.docx` (docDefaults 12pt/278 + Tabela–Siatka + gridSpan 2/3 + vAlign/jc center + trHeight): eksport niesie `sz=24`/`after=160 line=278`, komórki `w:spacing after=0 line=240`, tblGrid `3020/3021/3021` bez dryfu, tblPr bez tblBorders, walidator OOXML 0 błędów.

### Notes
- `data-no-borders="1"` pojawia się w 2. przejściu (regenerowany pakiet nie ma definicji stylu tabeli, więc resolved borders są puste) — bez skutków wizualnych (obramowania zapieczone per komórka), znika przy pass-through. Docelowe domknięcie = ConvertPreservingPackage w ścieżce Save (dziś tylko Download) — wymaga podania masterId do `POST /document/save` (zmiana kontraktu, do ustalenia).
- Pozostały dryf kosmetyczny: `tcMar` 108→105 tw (7px), `&nbsp;` pustych akapitów zyskuje wrapper `<span>` po 1. zapisie (idempotentne od 2. przejścia).

## 2026-07-07 — Kotwiczenie jak w Wordzie: pola tekstowe przeżywają zapis (wps:wsp), jawny model kotwicy w HTML, znacznik kotwicy w edytorze, obramowanie edycyjne textboxa (+36 testów)

### Changed
- **Writer (`HtmlToDocxConverter`) — NAJWIĘKSZA utrata danych domknięta**: `div.docx-textbox` wpadał w generyczną gałąź `div` i był spłaszczany do zwykłych akapitów — przy PIERWSZYM autosave ginęła ramka, pozycja, rozmiar i kotwica pola tekstowego (adres Qutalo w stopce itd.). Nowe: `BuildTextBoxDrawing` odtwarza `wps:wsp` + `w:txbxContent` (format DrawingML, Word 2010+; treść przez istniejące ConvertParagraph/Heading/Table/List), pływające → `wp:anchor` w konwencji edytora (X=posOffset od lewej krawędzi strony/`page`, Y=od górnego marginesu/`margin` — te same osie co obrazy, ADR-0029), inline → `wp:inline`; `data-border-*` → `a:ln`. Model kotwicy: **textbox kotwiczy do NASTĘPNEGO akapitu** — `BufferTextBoxDrawing`/`AttachPendingTextBoxes` przypina run z drawingiem do najbliższego następnego akapitu (`ConvertParagraphElement`/`ConvertHeadingElement`), awaryjny flush na końcu body/header/footer; bufor izolowany dla zagnieżdżonych pól; obsłużony też textbox inline w `<li>`/akapicie (`AppendInlineContent`).
- **Reader (`DocxToHtmlConverter`)**: (1) `RenderTextBoxContent` emituje jawne metadane kotwicy `data-pos-mode/x-emu/y-emu/width-emu/height-emu` (kontrakt wspólny z obrazami); (2) **hoisting**: div pola tekstowego jest emitowany bezpośrednio PRZED akapitem-kotwicą (`HoistTextBox` + wstrzyknięcie w `ConvertParagraphToHtml`), bo blokowy `div` w `<p>` jest re-parentowany przez parser przeglądarki — wypadał z akapitu i ROZCINAŁ go przy pierwszym renderze (wyjątek: `<li>` — div w li jest legalny, zostaje w środku); `ConvertAlternateContentToHtml` traktuje zbuforowany textbox jako skonsumowaną gałąź (bez dublowania Choice/Fallback); (3) koniec zapiekania obramowania EDYCYJNEGO `border:1px solid #ccc` w treść — realne obramowanie DOKUMENTOWE kształtu (`a:ln` z solidFill, poza txbxContent) idzie jako inline `border:` + `data-border-*`; (4) `data-wrap` (square/tight/through/topAndBottom) na obrazach i textboxach — writer odtwarza `wrapSquare`/`wrapTopBottom` (tight/through ≈ square, brak wrapPolygon w HTML) zamiast degradować wszystko do `WrapNone` (Word przestawał opływać obiekt po 1. autosave).
- **GUI (`wysiwyg-editor` + nowy `core/utils/floating-anchor.util.ts`)**: (1) **znacznik kotwicy jak w Wordzie** — zaznaczenie elementu PŁYWAJĄCEGO (obraz front/behind, textbox) pokazuje ikonę kotwicy (SVG Material „anchor") przy akapicie-kotwicy; overlay w `.page` POZA contenteditable (nie serializuje się), `pointer-events:none` + `role=img`/`aria-label` (nie kradnie kliknięć/fokusu), pozycja w px układu strony → zoom (transform) i scroll przesuwają go razem z treścią bez przeliczeń; repozycjonowanie koalescowane rAF po `onContentChange`/repaginacji (element usunięty → znacznik znika); kotwica akapitu: obraz = blokowy przodek, textbox = następny brat-akapit (czyste funkcje `findAnchorParagraph`/`computeAnchorBadgePosition`); (2) **interakcje textboxa** — klik zaznacza (`tb-selected`), drag pływającego pola TYLKO z pasa 8px krawędzi (wnętrze zostaje dla edycji tekstu, kursor `move` przez klasę `tb-edge` z mousemove), delty dzielone przez skalę zoomu, drag-end zapisuje `left/top` + `data-x/y-emu` (kotwica pozostaje przy tym samym akapicie — tylko offsety, bez skoków); (3) **fix zoomu w drag pływających OBRAZÓW** — delty `clientX/Y` dzielone przez `_pageVisualScale` (wcześniej obraz uciekał kursorowi przy zoomie ≠ 100%); (4) serializacja usuwa klasy stanu (`tb-selected/tb-dragging/tb-edge`) i defensywnie `.anchor-badge`; sprzątanie rAF/badge w `ngOnDestroy`.
- **SCSS**: obramowanie EDYCYJNE textboxa = `outline` (zero wpływu na box-model) widoczne przy hover/`:focus-within` (dashed) i zaznaczeniu/drag (solid) — stan nieaktywny bez ramki; realne obramowanie dokumentowe (inline `border` z readera) nieruszane; style `.anchor-badge`.

### Verified
- Backend: nowy `TextBoxAnchorRoundTripTests` **12/12** (metadane kotwicy, hoisting przed akapit, brak #ccc, realne a:ln → border, writer: anchor w NASTĘPNYM akapicie + offsety + inline, pełny round-trip DOCX→HTML→DOCX→HTML bez dublowania treści i **bez wzrostu liczby akapitów po 2 cyklach**, border round-trip, data-wrap read+write). Infrastructure **359/359**; build `D2ViewerEditor.sln` 0 błędów.
- GUI: nowe `floating-anchor.util.spec` **15/15** + `wysiwyg-editor.anchor.spec` **9/9** (badge dla obrazu front/inline/behind, chowanie, przełączanie zaznaczenia, usunięcie elementu, serializacja bez klas stanu i bez badge, drag aktualizuje EMU i nie zmienia kotwicy); pełne GUI **365/366** (jedyny fail = pre-existing `layout-shell`, failuje też w izolacji); `ng build` OK.

### Notes
- Model kotwicy = pozycja w DOM (bez sztucznych ID): obraz w akapicie-kotwicy, textbox przed nim. Usunięcie akapitu-kotwicy usuwa też element pływający będący w jego wnętrzu (obrazy — jak w Wordzie); textbox jako brat pozostaje i dopina się do kolejnego akapitu.
- Ograniczenia: `wps:wsp` bez opakowania `mc:AlternateContent` z fallbackiem VML (Word < 2010 nie pokaże pola; nowoczesny Word czyta wprost); wrapTight/Through wracają jako wrapSquare; edytor nadal renderuje wrap „square" jako float (przybliżenie); pozycja Y względem góry obszaru treści = przybliżenie ADR-0029.

## 2026-07-07 — Stopka/nagłówek: treść obok numeru strony nie znika — obsługa SdtRun w ścieżce pól złożonych (+7 testów)

### Changed
- `DocxToHtmlConverter.ConvertComplexFieldParagraphContent`: akapit zawierający `w:fldChar` (numer strony) przechodził przez ścieżkę pól złożonych, która iterowała tylko `Run`/`Hyperlink`/`SimpleField` — **`SdtRun` (formant inline) był dropowany w całości**. To dokładnie przypadek stopek Worda: tytuł dokumentu/klauzula obok numeru strony to content controls, a galeria „Strona X z Y" wkłada CAŁE pole PAGE do SdtRun (wtedy znikał nawet numer). Przepisane na maszynę stanów `ComplexFieldState` współdzieloną między poziomem akapitu a zawartością formantów (`AppendComplexFieldContent` + `AppendComplexFieldRun` + `AppendComplexFieldSdtRun`): rekurencja w `SdtContentRun`, wrapper `span.sdt-inline`+`data-sdt-*` spójny z `ConvertSdtRunToHtml`, kolejność elementów zachowana, PAGE/NUMPAGES → `{page}`/`{pages}`, pozostałe pola → wartość zbuforowana (KR-05).
- `HtmlToDocxConverter.BuildSdtRunFromHtml`: `span.field-page`/`page-number`/`field-numpages` wewnątrz `span.sdt-inline` wraca jako pole `fldSimple` PAGE/NUMPAGES (wcześniej literalny tekst „{page}" — utrata pola przy 1. autosave); warunek pustej zawartości uwzględnia `SimpleField`.
- Front: bez zmian kodu — `_computeFooterContent` podmienia `{page}/{pages}` na całym HTML (obejmuje sdt-inline), render przez `bypassSecurityTrustHtml`, CSS stopki nie ukrywa spanów.

### Verified
- Nowy `HeaderFooterFieldNeighborContentTests` 6/6: tekst SDT przed/po polu z asercją kolejności; galeria PAGE+NUMPAGES wewnątrz SDT (placeholdery dynamiczne, tekst zachowany); pole DATE w SDT trzyma wartość zbuforowaną; plain-runs wokół pól (pin poprawnego zachowania); writer eksportuje pola z sdt-inline bez „{page}"; pełny round-trip DOCX→HTML→DOCX→HTML.
- Infrastructure.UnitTests **347/347**; `dotnet build` solucji 0 błędów; GUI `wysiwyg-editor.spec` **32/32** (+1: treść i formanty obok numeru strony trafiają do stopki strony w poprawnej kolejności).
- Diagnoza potwierdzona harnessem przed/po (scratchpad, 8 wariantów stopki): przed fixem wariant „SDT obok pola" gubił oba teksty, wariant „pole w SDT" gubił wszystko; warianty taby-direct/taby-ze-stylu/tabela/SdtBlock/fldSimple/hyperlink były poprawne już wcześniej i pozostały bez zmian.

### Notes
- Taby w akapitach z polami złożonymi renderują się inline/flex (nie pozycyjnie) — `usePositionedTabs` wyklucza `hasComplexField`; treść jest kompletna, przybliżone jest tylko rozmieszczenie. Ewentualne pozycyjne taby w tej ścieżce = osobne zadanie.
- `.doc` (NPOI) poza zakresem tej poprawki — wada dotyczyła ścieżki OOXML.

## 2026-07-07 — Edytor: ENTER na dole strony nie rozciąga już wizualnie kartki — natychmiastowa repaginacja + przewinięcie kursora (+5 testów)

### Changed
- `wysiwyg-editor.ts`: ENTER nie był obsługiwany specjalnie — po `insertParagraph` przeglądarki repaginacja szła przez **debounce** `_schedulePaginate` (250 ms, max 600), więc przez ~250 ms strona była wizualnie „rozciągnięta" (CSS: `.page{min-height;overflow:visible}` + `.editor-content{flex:1}` z domyślnym `min-height:auto` flex-itemu rośnie z treścią), a nowa strona pojawiała się z opóźnieniem. Dodatkowo `_restoreGlobalCaret` nie wołał `scrollIntoView`, więc przelana treść lądowała poza widokiem — użytkownik musiał ręcznie scrollować, by zobaczyć drugą stronę.
- Fix A: `handleKeyboard` przy `Enter` (bez ctrl/meta/alt) woła `_flushPaginateSoon()` — koalescujący `requestAnimationFrame`, który po mutacji DOM wymusza repaginację NATYCHMIAST (`_flushPaginateNow` kasuje oczekujący debounce i woła `_repaginateNow`). Bez `preventDefault` (domyślny insertParagraph zostaje). Zwykłe pisanie nadal debounce'owane. Przytrzymany ENTER = 1 repaginacja/klatkę (guard `_paginateRafHandle`, sprzątany w `ngOnDestroy`).
- Fix B: `_restoreGlobalCaret` po ustawieniu karetki woła `target.scrollIntoView({block:'nearest',inline:'nearest'})` — nowa strona wchodzi w widok bez ręcznego scrolla; `nearest` nie rusza widoku, gdy kursor jest już widoczny (brak skoków przy pisaniu w środku strony).

### Verified
- `wysiwyg-editor.pagination-overflow.spec` **14/14** (+5: flush kasuje debounce, ENTER→flush, Ctrl+Enter→brak flushu, koalescencja rAF, scrollIntoView); `enter-caret` 3/3, `page-break` 2/2, `wysiwyg-editor.spec` 31/31 — brak regresji. `ng test --include` (builder Angulara).

### Notes
- Zmiana wyłącznie we froncie: writer HTML→DOCX, kontrakty API i model nietknięte. Rozciąganie kartki to wciąż stan przejściowy między keystroke a repaginacją — dla zwykłego pisania (nie-ENTER) nadal istnieje w oknie debounce'a, ale ENTER (główny objaw zgłoszony przez użytkownika) jest teraz natychmiastowy.

## 2026-07-07 — Wektorowe EMF (logo w stopce) przestały wychodzić jako niewidoczny blank — mapowanie window/viewport w tłumaczu metafile (+1 test)

### Changed
- `MetafileVectorTranslator` (strategia `vector-translate` w `GraphicConversionService`): tłumacz EMF ignorował mapowanie **page→device** (`SETWINDOWEXTEX/ORGEX`, `SETVIEWPORTEXTEX/ORGEX` były na liście „silent"), a `viewBox` SVG brał z `rclBounds` nagłówka (jednostki URZĄDZENIA). Współrzędne rysowania są LOGICZNE, więc dla metafile z niejednostkowym mapowaniem window/viewport (typowe logo Office, np. stopka wyciągu Qutalo = 2× klasyczny EMF ze ścieżek `BEGINPATH`/`POLYBEZIERTO16`/`LINETO`/`FILLPATH`) ścieżki lądowały **poza viewBox** → SVG poprawny, ale wizualnie pusty → logo widoczne w Wordzie, niewidoczne w edytorze.
- Fix: `GdiState.Apply` po world-transformie stosuje teraz mapowanie window→viewport (`(p−winOrg)·(vpExt/winExt)+vpOrg`), aktywne tylko gdy metafile faktycznie zdefiniował OBA zakresy (inaczej tożsamość — zero zmian dla metafile bez tych rekordów). Dodane obsłużone rekordy EMR 9/10/11/12. Safety-net w `BuildSvg`: jeśli mimo mapowania treść nie przecina `rclBounds` (np. nieobsłużony tryb metryczny), `viewBox` bierze bounding box realnie narysowanej treści — grafika nie wyjdzie pusta.

### Verified
- Nowy `MetafileVectorTranslationTests.Emf_WindowViewportMapping_ScalesLogicalCoordsIntoDeviceBounds` (prostokąt logiczny 500..1500 przy skali 0.1 → device 50..150 w viewBox `0 0 200 100`, `IsBlankFallback == false`).
- `MetafileVectorTranslationTests` + `GraphicConversionServiceTests` + `ImageImportRegressionTests` + integracja: 51/51. Pełny `Infrastructure.UnitTests`: **341/341**.

### Notes
- Oryginalny EMF nadal jedzie do DOCX bez zmian (pass-through / `data-original-src`) — zmiana dotyczy WYŁĄCZNIE podglądu w edytorze; Word/eksport nietknięte.
- Follow-up (ten sam dzień): logo renderowało się **CZARNE** (ścieżki widoczne, ale kolor prawdopodobnie w rekordzie, który pomijaliśmy). Dodano rendering **`EMR_ALPHABLEND` (typ 114)** → `<image>` z osadzonego DIB (Skia `DibToPng`), `SrcConstantAlpha`→`opacity` — logo Office często trzyma KOLOROWY raster w tym rekordzie, a ścieżki są czarną maską. Diagnostyka: `MetafileSvg.FillColors`/`HasEmbeddedImage`, w logu `fills=[...] embeddedImage=yes/no`; gdy wszystkie wypełnienia `#000000` i brak rastra → warning z markerem **`PODEJRZANE`** (eskalowany do WARNING w `WebGraphicForLegacy`, widoczny bez Debug).
- Poza zakresem (świadomie): tryby mapowania metryczne (MM_LOMETRIC..HITWIPS) nieskalowane wprost (łapie je safety-net content-bbox); grupa `wpg:wgp` = pierwszy blip (istniejące ograniczenie).

## 2026-07-07 — Kształty DrawingML z własną/preset geometrią (a:custGeom, ellipse, roundRect) renderowane w edytorze — grafika z oryginału przestała znikać (+1 test)

### Changed
- Reader (`DocxToHtmlConverter.RenderVectorShapeAsHtml`): kształt `wps:wsp` bez obrazu i bez pola tekstowego z **`a:custGeom`** (dowolna ścieżka wektorowa — np. wordmark „Qutalo", ikona ostrzeżenia „!") był dotąd dropowany w całości (metoda zwracała `""`), więc grafika z dokumentu oryginalnego **nie rysowała się w edytorze**. Teraz ścieżka jest tłumaczona na inline `<svg><path>` (`BuildCustomGeometrySvg`: komendy moveTo/lnTo/cubicBezTo/quadBezTo/close, współrzędne literalne w przestrzeni `a:path w/h`, `viewBox`+`preserveAspectRatio=none` = dokładne dopasowanie do `wp:extent`).
- Rozszerzono blok preset-geometrii o **elipsę** (`border-radius:50%`) i **zaokrąglony prostokąt** (`border-radius:12%`) obok istniejącego prostokąta/linii.
- Kolor wypełnienia brany precyzyjnie z properties kształtu (`spPr/a:solidFill`) przez nowy `GetShapeFillHex` — nie z pierwszego `a:solidFill` w poddrzewie (mógł należeć do obrysu `a:ln` lub ukrytej linii w `extLst`). Obrys emitowany tylko gdy `a:ln` ma realne wypełnienie (`noFill` → brak ramki, jak w Wordzie).

### Verified
- Nowy `Doc2ImportFidelityTests.CustomGeometryShape_WithoutImage_RendersAsInlineSvgPath`; pełne `Infrastructure.UnitTests` **340/340**.
- Harness na realnym szablonie Qutalo (`szablon 1`): wordmark „Qutalo" (`custGeom` navy `000066`) renderuje się jako `<svg>` (wcześniej: 0 kształtów w body); podgląd potwierdzony wizualnie (headless Chrome). Round-trip HTML→DOCX (`HtmlToDocxConverter.Convert`) nie rzuca na inline `<svg>`.

### Notes
- PODGLĄD-only: writer (jak przy istniejących liniach/prostokątach) NIE odtwarza tych kształtów z powrotem do DOCX — na 1. autosave kształt znika z v2 (oryginał v1 nietykalny; do rozważenia round-trip przez `data-original-*`). Wypełnienie motywowe (`a:schemeClr`) bez mapowania → fallback czarny (kształt nadal widoczny). Wypełnienie regułą nonzero (zgodnie ze spec DrawingML) — litery z zamkniętymi „oczkami" mogłyby wymagać evenodd.

## 2026-07-07 — Kotwiczone obiekty (wp:anchor): pozycja liczona jak w Wordzie (relativeFrom + wp:align) — logo/pole tekstowe trafiają na właściwe miejsce (+8 testów)

### Changed
- Reader (`DocxToHtmlConverter`): kotwice `wp:anchor` rozwiązywane przez nowy `ResolveAnchorPosition`/`ResolveAxis` z uwzględnieniem **`relativeFrom`** (page/margin/column/leftMargin/…) oraz **`wp:align`** (right/center/left/inside/outside) — wcześniej brany był surowy `wp:posOffset` z pominięciem obu, więc obiekty kotwiczone do marginesu/kolumny lub wyrównane do prawej lądowały o cały margines za bardzo w lewo / za wysoko. Dotyczy obu ścieżek: obrazów (`ConvertDrawingToHtml` → `data-x-emu`/`data-y-emu`) i pól tekstowych/kształtów (`BuildTextBoxLayoutCss`, `RenderVectorShapeAsHtml` → inline `position:absolute`). Origin edytora: X = lewa krawędź strony, Y = góra obszaru treści (odjęty górny margines).
- Geometria pierwszej sekcji (rozmiar strony + marginesy w twipach) zapamiętywana w polach instancji (`LoadPageGeometry`) — potrzebna do przeliczenia align/relativeFrom na piksele.
- Writer (`HtmlToDocxConverter`): pionowa kotwica eksportowana jako `RelativeFrom=Margin` (góra obszaru treści) zamiast `Page` — spójne z pionowym originem edytora i poprawne w Wordzie (dotąd obiekt renderował się o górny margines za wysoko). Poziom zostaje `Page`.
- `OoxmlUnits.TwipsToEmu` (+`EmuPerTwip=635`).

### Verified
- Nowy `DocxAnchorPositionFidelityTests` 8/8 (page/margin/column offset; align right/center; page-align-right; vertical content-relative; page-Y odejmuje margines; brak offsetu/align → baza). Zaktualizowane: `Doc2ImportFidelityTests.AnchoredTextBox…` (page-Y 2″ → 1″ poniżej treści), `ImageFloatingRoundTripTests` (idempotentny round-trip po zmianie writer V→Margin). Infrastructure **339/339**, solucja build 0 błędów.

### Notes
- Round-trip idempotentny: writer H=Page/V=Margin ↔ reader dodaje i odejmuje ten sam górny margines. Frontend bez zmian (restore czyta gotowe współrzędne edytora). Przybliżenia: pionowe `wp:align` center/bottom rzadkie i słabo określone przy rosnącym obszarze treści (traktowane jak offset/top); inside/outside bez rozróżnienia stron parzystych. Decyzja: **ADR-0029**.

## 2026-07-07 — „Inne na pierwszej stronie" (titlePg): stopka/nagłówek pierwszej strony nie wyciekają na kolejne strony (+2 testy)
### Changed
- `DocxToHtmlConverter.ExtractHeader`/`ExtractFooter`: fallback `mainPart.HeaderParts/FooterParts.FirstOrDefault()` używany TYLKO gdy sekcja nie deklaruje ŻADNEJ referencji nagłówka/stopki. Gdy sekcja ma `titlePg` + referencję `first` (i BRAK referencji `default`), domyślny nagłówek/stopka jest CELOWO pusty (Word nic nie pokazuje na zwykłych stronach) — wcześniej fallback wciągał część pierwszej strony jako domyślną i renderował ją na WSZYSTKICH stronach (np. adres Qutalo w stopce widoczny na str. 2, choć w Wordzie tylko na str. 1).
- Metoda zwraca teraz `HeaderFooterContent` także gdy domyślny wariant jest pusty, ale istnieje `FirstPageHtml`/`EvenHtml` (`Html = string.Empty`, `DifferentFirstPage = true`); `null` tylko gdy nie ma ani domyślnego, ani first/even.
- Nowe pomocnicze `SectionDeclaresAnyHeaderReference`/`SectionDeclaresAnyFooterReference`.
- Front bez zmian — `wysiwyg-editor` już rozwiązuje `differentFirstPage` (str. 0 = `firstPageHtml`, reszta = `html`).
### Verified
- `DocxToHtmlConverterSectionReferenceTests` 10/10 (+2: `Header_FirstPageOnly...`/`Footer_FirstPageOnly...`), Infrastructure build 0 błędów.
- Uwaga: 3 pre-existing faile `ImageFloatingRoundTripTests` pochodzą z RÓWNOLEGŁEGO WIP kotwic (`ResolveAnchorPosition`) w drzewie roboczym — nie z tej zmiany (potwierdzone: przechodzą na czystym HEAD konwertera).
### Notes
- Round-trip (writer) poza zakresem tej poprawki — dotyczy importu/renderu. Wielosekcyjne `ExtractHeaderOwnedBySection`/`ExtractFooterOwnedBySection` nie mają tego fallbacku (już rozwiązują tylko własny `default`), więc bez zmian.

## 2026-07-07 — SVG „puste białe logo": sanitizer wycinał wewnętrzne `<use>` + odrzucał BOM/DOCTYPE (+6 testów)
### Changed
- **`GraphicConversionService.SanitizeSvg`** — trzy przyczyny, dla których legalne logo SVG (przypadek: nagłówek Doc2/Qutalo) traciło treść lub cały plik:
  1. `<use>` był na liście `killTags` i wycinany BEZWARUNKOWO — logo zbudowane z `<defs>`+`<use href="#id">` (typowy eksport korporacyjny) renderowało się jako pusty biały obraz o poprawnych wymiarach (defs są niewidoczne). Teraz `<use>` z wewnętrznym odnośnikiem `#id` (href/xlink:href) ZOSTAJE (wskazuje już-sanityzowaną treść tego samego dokumentu); `<use>` z zewnętrznym URL/data:/bez href — usuwany w całości (`HasInternalFragmentHref`).
  2. Bajty z prefiksem **UTF-8 BOM** (częste w plikach z Windows) — U+FEFF na początku stringa wywala parser XML → poprawny SVG odrzucany → obraz znikał; `TrimStart('﻿','​',…)` przed parsowaniem.
  3. **DOCTYPE** (standard w SVG z eksportu Illustratora) — `DtdProcessing.Prohibit` rzucał na sam DOCTYPE → odrzucenie; zmiana na `DtdProcessing.Ignore` w `SafeParse`: DTD pomijany bez przetwarzania, XXE/billion-laughs nadal niemożliwe (encje pozostają niezdefiniowane → parser rzuca → null, potwierdzone testem).
### Verified
- `GraphicConversionSecurityTests` (w tym nowe: defs+use zachowane, use zewnętrzny usuwany, BOM przyjęty, DOCTYPE Illustratora przyjęty, XXE nadal null) + `Doc2ImportFidelityTests` (nowe: logo defs/use w data-URI z `<use>`, SVG z BOM osadzony) — testy SVG **25/25**; pełne Infrastructure **325/328** w czystym worktree (3 faile = testy hMerge/tab-in-cell równoległej sesji wymagające jej WIP `LoadPageGeometry`, niezwiązane).
### Notes
- Objaw u użytkownika (biały pusty obraz logo w nagłówku edytora przy poprawnym widoku w Wordzie) pasuje 1:1 do przyczyny (1); (2)/(3) dawały twardsze zniknięcie obrazu. Jeżeli po wdrożeniu logo nadal puste — sprawdzić w logach backendu wpisy „SVG part pominięty…"/„Media part bez rastra web…" oraz atrybut `data-legacy-graphic="blank"` w DOM (wtedy to EMF+/metafile poza etapem 1 tłumacza ADR-0027).
- Ograniczenie bez zmian: writer nadal dropuje obraz SVG przy eksporcie HTML→DOCX (utrata na 1. autosave; roadmapa `asvg:svgBlip`/rasteryzacja).

## 2026-07-07 — Formant blokowy w nagłówku/stopce: treść znikała z podglądu (+1 test)
### Changed
- **Reader (`DocxToHtmlConverter.ConvertHeaderFooterToHtml`)** — pętla po elementach nagłówka/stopki obsługiwała tylko `Paragraph`/`Table`, więc formant blokowy (`SdtBlock`) osadzony BEZPOŚREDNIO w stopce (np. klauzula prawna `removeif_nondigitalversion`) był po cichu pomijany → cała treść formantu znikała z podglądu edytora (Word ją pokazywał). Dodana gałąź `SdtBlock` → `ConvertSdtBlockToHtml(sdt, document, part)`.
- **`ConvertSdtBlockToHtml`** — przyjmuje teraz opcjonalny `OpenXmlPart? sourcePart` i przekazuje go do konwersji dzieci (Paragraph/Table/zagnieżdżony SdtBlock), żeby obrazy w formancie osadzonym w nagłówku/stopce rozwiązywały rId względem właściwej części pakietu (rId unikalne per część). Ścieżki body/komórki tabeli bez zmian (sourcePart=null jak dotąd).
### Verified
- `SdtContentControlRoundTripTests` **5/5** (nowy `BlockContentControl_InFooter_ContentIsNotDropped`); `dotnet build` Infrastructure 0 błędów (build równoległy potrafi zgłaszać fałszywe CS0103 — `-m:1` czysty).
### Notes
- Formant inline w akapicie stopki (`SdtRun`) już działał — obsługuje go `ConvertParagraphToHtml`. Luką był wyłącznie poziom bloku w nagłówku/stopce.
- To render treści; interaktywność formantu (klikalny checkbox/dropdown) nadal osobne zadanie frontu. QR-kod w nagłówku = osobny wątek importu obrazów.

## 2026-07-07 — Tabele (Doc2/Qutalo follow-up): legacy w:hMerge + taby w komórkach bez pozycjonowania absolutnego (+4 testy)
### Changed
- **Reader (`DocxToHtmlConverter`) — legacy scalanie poziome `w:hMerge`** — dotąd obsługiwany był wyłącznie `w:gridSpan`; komórki scalone mechanizmem legacy (`hMerge` restart + continue) renderowały się jako OSOBNE `<td>` (treść ściśnięta w pierwszej wąskiej kolumnie, reszta puste widma — wiersz „Umowa wieloproduktowa…"). Nowy `BuildRowRenderPlan`: restart pochłania `gridSpan` kolejnych continue jako `colspan`, continue nie emitują `<td>` (analogicznie do kontynuacji vMerge); pominięty `w:val` = continue (ECMA-376); continue-sierota bez restartu renderuje się normalnie (bez utraty treści). Round-trip: writer i tak emituje `gridSpan` z colspan — reprezentacja równoważna dla Worda.
- **Reader — taby w komórkach tabel bez segmentów absolutnych** — rozszerzenie renderingu pozycyjnego tabów na body (2026-07-06) objęło też akapity WEWNĄTRZ komórek: segment `position:absolute;left:{stop}px` jest kotwiczony do akapitu, a pozycje stopów opisują geometrię strony → w wąskiej komórce segment wyjeżdżał poza komórkę i malował się po sąsiedniej kolumnie (nałożone nagłówki „Waluta"/„Termin spłaty kredytu"), a `text-align` komórki przestawał działać (rozjechane wyśrodkowanie nagłówków i wartości). `usePositionedTabs` wyklucza teraz akapity z przodkiem `TableCell` — w komórce tab renderuje się inline/flex jak przed regresją; `data-tab-stops` nadal niesie stopy (eksport bez zmian).
### Verified
- `Doc2ImportFidelityTests` **24/24** (+4: hMerge restart+2×continue → jeden `<td colspan="3">` i brak widm; pominięty val=continue; continue-sierota bez utraty treści; tab w komórce → brak `docx-tab-seg`/`position:absolute`, `data-tab-stops` zachowane). Infrastructure **321/322**, build 0 błędów.
### Notes
- Jedyny fail pełnego przebiegu: `SanitizeSvg_KeepsSafeDataHrefAndFragment` (`GraphicConversionSecurityTests`) — PRE-EXISTING na HEAD (sanitizer SVG zachowuje `<use href="#g">`, test oczekuje usunięcia; obszar nietknięty tą zmianą). Do wyjaśnienia osobno.
- Pionowe centrowanie komórek i `td>p{margin:0}` naprawione już 2026-07-06; ta zmiana przywraca działanie POZIOMEGO wyrównania w komórkach z tabami. Warunkowe formatowanie TEKSTU ze stylu tabeli (np. centrowanie nagłówka ze stylu) pozostaje znanym ograniczeniem (ADR-0026).

## 2026-07-06 — Audyt import/edycja/eksport DOCX: fixy edytora (font/page-break/geometria) + zachowanie wartości pól + narzędzia diagnostyczne (ADR-0028)
### Changed
- **Podział strony (edytor)** — `insertPageBreak` wstawia semantyczny, NIEDRUKOWALNY marker `<div class="page-break" contenteditable="false"></div>` bez inline-grafiki; SCSS: domyślnie zero wysokości/bez linii/etykiety, subtelna podpowiedź tylko w trybie „znaki formatowania" (`.editor-content.show-formatting-marks`, sygnał `showFormattingMarks`). Eksport bez zmian (writer i tak mapował `div.page-break` → `w:br type=page`, nigdy drawing/picture/shape) — potwierdzone.
- **Wybór fontu (główny toolbar)** — natywny `<select>` (zerował `selectedIndex=-1` na mousedown → pole białło) zastąpiony **comboboxem** `<input list>`+`<datalist>`: pokazuje efektywny font, pozwala wpisać i wyszukać, NIE znika na focusie, commit dopiero na Enter/change/blur, Escape przywraca, stan **mieszany** (puste) dla zaznaczenia wielu fontów (`EditorState.fontMixed` + skan `computeFontMixed`). Usunięte debug `console.log` (updateBlockFormatFromState, updateFormattingState, applyDocumentStyle).
- **Wspólne źródło fontów (item 7)** — nowy `FontProviderService` (providedIn root): jedna lista dla toolbara głównego i kontekstowego, z **fontem firmowym** czytanym z `--corporate-font-family`, wspólna normalizacja nazw. `editor-toolbar` i `document-editor` (mini-toolbar) czytają tę samą `displayNames`.
- **Jednostki/geometria (item 1)** — nowy `core/utils/units.util.ts` (jedyne źródło `CSS_PX_PER_CM=37.8` + `cmToPx/pxToCm/twipsToPx/ptToPx`); linijka używała **37.795**, strona **37.8** → dryf usunięty. Zoom potwierdzony jako czysto wizualny (`transform:scale`), nie zmienia modelu ani eksportu.
- **Reader: wartości pól (KR-05/KR-08, item 3)** — złożone pola inne niż PAGE/NUMPAGES nie gubią już wartości cache (`if (fieldSeparated && fieldValueHandled) continue;`); DATE (simple i complex) zachowuje datę z dokumentu zamiast `DateTime.Now`.
- **Narzędzia diagnostyczne** — nowe `tools/docx-diagnostics/` (zero-dependency Node ESM): `inspect-docx` (struktura + pola + relacje + elementy z ryzykiem utraty, JSON+konsola, exit 1/2) i `compare-docx` (diff strukturalny źródło↔eksport, klasyfikacja equal/…/lost, exit 1 na regresje). Własny czytnik ZIP (zlib) + writer w fixtures.
### Verified
- GUI: nowe specy `wysiwyg-editor.page-break` 2, `editor-toolbar.font` 6, `font-provider.service` 4, `units.util` 5 — zielone; pełne GUI **334/335** (jedyny fail pre-existing `spec-layout-shell`); `ng build` OK.
- Backend: nowe `FieldValuePreservationTests` 3; Infrastructure **318/318**; `dotnet build` całej solucji 0 błędów.
- Narzędzia: `tools/docx-diagnostics` self-test **8/8**; `inspect-docx`/`compare-docx` uruchomione end-to-end na syntetycznych fixture'ach (exit codes potwierdzone).
### Notes
- Font firmowy testowany deterministycznie przez podklasę providera (CSS var w jsdom). „Mieszany" font: skan zaznaczenia działa w przeglądarce; w jsdom degraduje do „nie-mieszany".
- Round-trip KODU pola (nie tylko wartości) dla TOC/REF nadal poza zakresem (wartość zachowana jako tekst = „unlink field" jak w Word) — patrz `AUDIT_WORD_COMPATIBILITY.md` KR-05.

## 2026-07-06 — Formanty (SDT / Content Controls): zachowanie TYPU i właściwości przez round-trip (ST-04, +4 testy)
### Changed
- **Reader (`DocxToHtmlConverter`)** — `BuildSdtDataAttrs`: obok `data-sdt-tag`/`data-sdt-alias` niesie PEŁNE `w:sdtPr` (base64 OuterXml) w `data-sdt-props` dla `sdt-block` i `sdt-inline`. Wcześniej czytane były tylko tag/alias → typ formantu ginął.
- **Writer (`HtmlToDocxConverter.BuildSdtProperties`)** — odtwarza `SdtProperties` 1:1 z `data-sdt-props` (`new SdtProperties(xml)`), usuwając tylko `w:id` (Word nadaje nowe — brak kolizji). Fallback do tag/alias, gdy brak/uszkodzone. Skutek: dropDownList (opcje), date (format), comboBox, checkbox (w14), text/richText/picture, lock, placeholder, databinding **przeżywają eksport/autosave** zamiast degradować do generycznego formantu.
### Verified
- Nowe `SdtContentControlRoundTripTests` **4/4** (dropDownList: typ+tag+opcje; date: format; inline text: typ+tag+treść; duplikat id usuwany). Infrastructure **315/315**; `dotnet build` 0 błędów.
### Notes
- To zachowanie DANYCH (round-trip), NIE interaktywny UI — w edytorze formant nadal renderuje się jako edytowalna treść (checkbox nieklikalny, dropdown bez listy). Interaktywność = osobne zadanie frontu.
- `SdtCell` (formant na poziomie komórki) nadal rozpakowywany do zwykłej `TableCell` — treść zostaje, opakowanie formantu na komórce nie round-tripuje (do zrobienia).

## 2026-07-06 — Numeracja stron: eksport pola (nie literał „{page}") + poprawny licznik „z N" per strona (+4 testy)
### Changed
- **Eksport (`HtmlToDocxConverter`)** — pola PAGE/NUMPAGES w TREŚCI (body) wychodziły jako **literalny tekst „{page}"/„{pages}"** (placeholder readera), bo ścieżka body (`AppendInlineContent`) nie rozpoznawała `span.field-page`/`page-number`/`field-numpages` (obsługiwała je tylko ścieżka nagłówka/stopki). Dodano rozpoznawanie tych klas w `AppendInlineContent` → `w:fldSimple` PAGE/NUMPAGES. Dodatkowo `BuildFieldRun` nie używa już `{page}`/`{pages}` jako **zbuforowanej wartości** pola (Word pokazuje cache do przeliczenia) — placeholder/empty/nie-liczba → „1".
- **GUI (`wysiwyg-editor.ts`)** — „numeracja tylko na 1 stronie / 1 z 1": `headerContents`/`footerContents` iterowały **martwy sygnał `pages`** (inicjowany długością 1, NIGDY nie ustawiany), więc nagłówek/stopka liczyły się tylko dla strony 0, a strony 2+ dostawały surowy fallback z nierozwiniętymi `{page}`/`{pages}`. Repointowano na `pageContents()` (rzeczywista lista stron); `{pages}` liczy `pageContents().length`; usunięto martwy sygnał `pages`.
### Verified
- Nowe `PageNumberFieldExportTests` **4/4** (body PAGE+NUMPAGES jako pola, page-number span → PAGE, cache ≠ placeholder, stopka nadal OK); Infrastructure **311/311**; `dotnet build` 0 błędów; GUI `ng test` (patrz niżej).
### Notes
- Wartość pola to nadal placeholder-liczba „1" w cache — realny numer wylicza Word/edytor przy renderze/przeliczeniu (brak twardej paginacji w modelu = ograniczenie stałe).

## 2026-07-06 — Wierność importu Word („Doc2"/Qutalo): tab-stopy body, scalone wiersze nieregularne, v-align, SVG (+16 testów)
### Changed
- `Infrastructure/Services/DocxToHtmlConverter.cs`:
  - Tab-stopy body renderowane POZYCYJNIE (`usePositionedTabs` już nie wymaga `HeaderPart`/`FooterPart`) — lewy/prawy/wiele stopów na realnych pozycjach `w:tabs` zamiast flexa 50/100%; flex tylko jako fallback bez rozwiązywalnych pozycji.
  - Tabele — wiersze krótkie: `ConvertTableToHtml` liczy `deficit = GridColumnCount − Σ gridSpan` i przekazuje `extraColspan` do OSTATNIEJ komórki (`AppendTableCellHtml`), która pochłania brakujące kolumny; nowy helper `FlattenRowCells` (rozpakowuje SdtCell). Jawny gridSpan/vMerge/pełne wiersze bez zmian.
  - Wyrównanie pionowe komórki emitowane RAZ z rozwiązaną wartością (`top`/`middle`/`bottom`) — koniec zduplikowanego `vertical-align`.
  - SVG: `NormalizeImageContentType` (`img/svg+xml`/`image/svg`→`image/svg+xml`), `IsSvgContentType`; części SVG sanityzowane przez `IGraphicConversionService.SanitizeSvg` (limit `MaxSvgBytes`=2 MB) — niepoprawne/za duże odrzucane zamiast osadzania surowego.
  - **Pola tekstowe (KR-06)**: `RenderTextBoxContent` renderuje `w:txbxContent` (nowoczesne `wps:txbx` ORAZ legacy VML `v:textbox`) jako blok `div.docx-textbox` zamiast cichego dropu; wpięte w `ConvertDrawingToHtml` (kształt bez blipa), `ConvertPictureToHtml` (VML bez obrazu) i `ConvertAlternateContentToHtml` (fallback po nieudanym obrazie; Choice/Fallback niosą tę samą treść → renderowana raz). Treść pola tekstowego przestaje ginąć na 1. autosave. **Pozycja jak w MS Word:** `BuildTextBoxLayoutCss` — kotwica `wp:anchor` → `position:absolute` z offsetów (EMU→px, inline-CSS reader-side, działa też w podglądzie stopki/nagłówka bez JS edytora); inline → blok w przepływie; rozmiar z `wp:extent`.
  - **Kształty wektorowe bez obrazu (linie/prostokąty)**: `RenderVectorShapeAsHtml` — preset `line`/`straightConnector1` → widoczna pozioma linia (`div.docx-line`, grubość/kolor z `a:ln`), prostokąt z `a:solidFill` → kolorowy blok (`div.docx-rect`); wcześniej dropowane (brak blipa) → niewidoczne w stopce (separatory/ramki). Kotwiczone → pozycja absolutna.
- `D2GuiViewerEditor/.../wysiwyg-editor.scss`: `.editor-content table td>p/th>p { margin:0 }` (Word: komórki mają 0 odstępu; blankietowy `p{margin:0 0 10px}` psuł centrowanie i pompował tabele; inline marginesy nadal wygrywają).
### Verified
- `Doc2ImportFidelityTests` **20/20** (m.in. tab lewy/prawy/wiele, v-align+dedup, krótkie wiersze, spacing+`<br/>`, SVG valid/normalize/sanitize/reject, pole tekstowe wps:txbx w AlternateContent raz + VML standalone + kotwica→position:absolute z offsetów, linia straightConnector1→docx-line kolor/rozmiar).
- Infrastructure **305/305** (golden `tab-stops-lcr` zregenerowany), Application **306/306**, `dotnet build` solucji 0 błędów, GUI `ng test` **317/318** (jedyny fail pre-existing `spec-layout-shell`), `ng build` OK.
### Notes
- Writer HTML→DOCX nadal pomija obraz SVG (goły `a:blip` = uszkodzony DOCX) → SVG ginie na 1. autosave; docelowo rasteryzacja lub `asvg:svgBlip` z fallbackiem PNG.
- Pole tekstowe/kształt: `div.docx-textbox`/`docx-line`/`docx-rect` na eksporcie staną się zwykłymi akapitami/pustymi divami (treść ZACHOWANA, kształt/pozycja nie round-tripują jako `wps:wsp` — przybliżenie); pełny round-trip kształtu = osobne zadanie writer'a.
- Pozycja kotwicy: `relativeFrom` (page/margin/column/paragraph) nie jest w pełni rozróżniane — offset stosowany bezpośrednio (dobre dla page/margin; column/paragraph = przybliżenie).
- **Logo/linie w stopce niewidoczne** = wektorowe EMF (`image2..image6.emf`; 530 KB = logo). Histogram rekordów `image2.emf` (EMF+ = False): treść to ścieżki z wariantów „To" (`POLYLINETO16`/`POLYBEZIERTO16`) + `MOVETOEX`/`BEGINPATH`/`FILLPATH`. **Root cause w `MetafileVectorTranslator`:** gdy `MOVETOEX` wypada PRZED `BEGINPATH` (typowy wzorzec GDI+ FillPath), wiodące „M" nie było zapisywane, więc ścieżka SVG zaczynała się od „L"/„C" → niepoprawna → nic się nie renderowało (całe logo niewidoczne). Fix: `EnsurePathStart` domyka wiodące moveto z bieżącej pozycji w `LINETO`/`POLYLINETO`/`POLYBEZIERTO` (gałąź ścieżki). Teraz wektorowe logo RENDERUJE się cross-platform (pure-managed, bez zależności). Dodatkowo widoczny placeholder dla metafile, które NADAL są blank: `wysiwyg-editor.scss` `img[data-legacy-graphic='blank']` (ramka + wzór; `<img>`/round-trip EMF nietknięte). Ograniczenia: `SELECTCLIPPATH`/`SETCLIPPATH` (clipping) i tekst `ExtTextOut` nadal nieobsługiwane (etap 2). **Wymaga PONOWNEGO IMPORTU oryginału** — reader (DOCX→HTML) uruchamia translator przy imporcie; zapisane v2 (HTML) trzeba odświeżyć.
- Issue „Formatny" — nie występuje w kodzie/`.ai`/diagramach; brak jednoznacznej interpretacji, nie implementowano (patrz raport końcowy), dane nie są usuwane.

## 2026-07-05 — Własny tłumacz wektorowy EMF/WMF → SVG, etap 1 (ADR-0027, +10 testów)
### Changed
- Nowy `Infrastructure/Services/MetafileVectorTranslator.cs` (internal, pure-managed): parser rekordów binarnych MS-EMF/MS-WMF → SVG. Podzbiór etapu 1: pióra/pędzle (CreatePen/ExtCreatePen/CreateBrushIndirect, stock objects, dash style, COLORREF), MoveTo/LineTo, Rectangle/Ellipse/RoundRect, Polygon/Polyline/PolyBezier(To) 16/32-bit, PolyPolygon (fill-rule z SetPolyFillMode), ścieżki (BeginPath/CloseFigure/Fill/Stroke/StrokeAndFillPath), transformacje świata (SetWorldTransform/ModifyWorldTransform — rotacja rect → polygon), StretchDIBits → `<image>` (istniejące SliceDib/DibToPng + Skia). WMF: slotowa tabela obiektów (fonty/palety/regiony zajmują sloty!), placeable bbox / SETWINDOWORG/EXT jako viewBox, odwrócona kolejność parametrów. Limity: 200k rekordów, 100k pkt/poly, 20k elementów, 2 MB; wyjątek → null.
- `GraphicConversionService.ConvertMetafile`: nowa strategia **`vector-translate`** między `dib-rasterize` a blankiem; wynik przez `SanitizeSvg`; Status=Converted/Fidelity=Lossy; pominięte rekordy liczone w Warnings + `LostProperties`; `SliceDib`/`DibToPng` udostępnione internal.
- Reader: `LoadImageFromPart` loguje „media part bez rastra web" tylko dla Fallback/Unsupported/Rejected (udana translacja SVG to nie problem); SVG nie podmienia bajtów w `_images` — oryginalny metafile zostaje do `data-original-src` (eksport = prawdziwy EMF/WMF, SVG tylko podgląd).
### Verified
- Nowe `MetafileVectorTranslationTests` **10/10**: EMF rect (kolory pióra/pędzla, grubość), POLYGON16, linia, ścieżka (M/L/Z, fill none), SetWorldTransform (translacja), header-only → nadal blank, nieznane rekordy → blank bez wyjątku, WMF rect+polygon, integracja DOCX (SVG w src, bez `data-legacy-graphic="blank"`, eksport = part `image/x-emf`, nie SVG).
- Infrastructure.UnitTests **285/285** (stare testy blank-fallback bez zmian — tłumacz nie fabrykuje treści), Application **306/306**, build solucji 0 błędów.
### Notes
- Zero nowych zależności; pure-managed → identyczne zachowanie Windows/Linux/GCP. Poza etapem 1: tekst (ExtTextOut), clipping, ROP, pędzle wzorkowe, rekordy EMF+ (dual niesie fallback EMF). Decyzja i szczegóły: **ADR-0027** + `GRAPHICS_CONVERSION.md`.

## 2026-07-05 — Import obrazów z DOCX: 6 potwierdzonych bugów naprawionych w istniejącym mechanizmie (+10 testów regresyjnych)
### Changed
- **Reader (`DocxToHtmlConverter`)**:
  - **Kolizja rId między częściami pakietu** — `_images` był kluczowany samym rId, a rId-y są unikalne tylko w obrębie części (main/header/footer zaczynają od rId1); obraz nagłówka o kolidującym rId renderował obraz z body. Nowy klucz `ImageCacheKey(part, rId)` = `{part.Uri}|{rId}`; lookup i lazy-load w `ConvertDrawingToHtml`/`ConvertPictureToHtml` rozwiązują relacje względem `sourcePart ?? MainDocumentPart`.
  - **`mc:AlternateContent` był dropowany w całości** (default w switchu → pusty string) — obrazy zakotwiczone z efektami/grupy/kanwy znikały bez śladu. Nowa gałąź `ConvertAlternateContentToHtml`: Choice-e w kolejności dokumentu, potem Fallback (w:pict VML); pierwsza gałąź dająca HTML wygrywa.
  - **Zerowy `wp:extent` (cx/cy=0)** dawał `width:0px` — obraz w DOM, ale niewidoczny. Przy ≤0 wymiary intrinsic z nagłówka pliku (probe przez `ConvertForEditor`, wynik cache'owany po hashu).
  - **`WebGraphicForLegacy` obejmuje teraz `GraphicKind.Unknown`** — EMZ/WMZ i formaty rozpoznawalne tylko przez Skia nie trafiają już do `src` jako nierenderowalny `data:{ct}` (ikona złamanego obrazka); w najgorszym razie przezroczysty blank + `data-original-src`.
  - **Obraz linkowany (`r:link` bez `r:embed`)** — kontrolowane pominięcie z logiem zamiast cichego dropa; lazy-load loguje nieudane relacje zamiast `catch {}`.
  - Opcjonalny `ILogger<DocxToHtmlConverter>` (NullLogger domyślnie): część DOCX, relId, deklarowany typ, wykryty format, rozmiar, status, strategie, powód błędu — bez danych binarnych.
- **`GraphicConversionService`**: dekompresja **GZIP (EMZ/WMZ)** przed detekcją (`gzip-decompress` w `AttemptedStrategies`), bounded do `MaxInputBytes` (ochrona przed decompression bomb); `IsNonBrowserNativeContentType` w readerze obejmuje `emz`/`wmz`.
- **Writer (`HtmlToDocxConverter.BuildImageDrawing`)**: koniec fałszowania content type partu — wcześniej wszystko spoza krótkiej listy szło jako `ImagePartType.Jpeg` (TIFF/ICO/WEBP przestawały się renderować po pierwszym autosave). Teraz: mapowania `image/tiff→Tiff`, `image/x-icon→Icon`, nieznane `image/*` (np. `image/webp`, `image/x-emz`) → `AddImagePart(contentType)` z zachowanym typem; nie-obrazowe MIME → historyczny fallback Jpeg.
### Verified
- Nowe `ImageImportRegressionTests` **10/10** (kolizja rId body vs header, Choice z drawingiem, Fallback z VML, EMZ z osadzonym PNG → PNG w src, EMZ czysty wektor → blank, gunzip w serwisie, zero-extent → intrinsic 20x20, writer TIFF/WEBP content type, r:link bez wyjątku).
- Infrastructure.UnitTests **275/275**, Application.UnitTests **306/306**, `dotnet build D2ViewerEditor.sln` 0 błędów. Golden snapshoty bez zmian (poprawki nie zmieniają HTML dla wcześniej działających obrazów).
### Notes
- Zero nowych zależności (GZipStream = BCL, pure-managed; zgodne z zakazem LibreOffice/Office/COM — Linux/GCP safe). Testy na Linuxie: do potwierdzenia w CI/kontenerze (lokalny Docker niedostępny w tej sesji); zmiany nie dotykają niczego platform-specific.
- Znane, świadome ograniczenia (bez zmian): czysty wektor EMF/WMF → przezroczysty blank (pass-through do DOCX), obrazy linkowane `r:link` nie są pobierane (SSRF), textboxy/oMath/wykresy w AlternateContent nadal nierenderowane (audyt), grupa z wieloma obrazami renderuje pierwszy blip.

## 2026-07-05 — Zwiększenie pokrycia testami: propagacja SectionHeadersFooters + brzegi ADR-0025/0026 (+16 testów, tylko testy)
### Changed
- **Application.UnitTests (+4)**: `SaveDocumentCommandHandlerTests` — `PageSize`+`SectionHeadersFooters` przekazywane do `Convert` nietknięte (ta sama referencja) oraz null-default; `DownloadEditedDocumentCommandHandlerTests` — lista sekcji dochodzi do OBU ścieżek (`Convert` bez wersji bazowej i `ConvertPreservingPackage` z wersją bazową).
- **Infrastructure.UnitTests (+12)**: `TabStopFidelityTests` +4 (leader `dot` serializowany w `data-tab-stops`, `w:val=clear` usuwa stop odziedziczony ze STYLU akapitowego, direct nadpisuje stop stylu na tej samej pozycji, bar-tab pomijany); `TableStyleFidelityTests` +5 (lastRow tylko na ostatnim wierszu, firstColumn tylko w kolumnie 0, legacy maska hex `tblLook@w:val` bez atrybutów boolowskich, `noHBand` wyłącza pasy mimo definicji band1Horz w stylu, pasy PIONOWE band1Vert alternują kolumny); `MultiSectionFidelityTests` +3 writer (wpis footer-only → `FooterReference` wyłącznie na sectPr sekcji 1, wpis z indeksem poza zakresem ignorowany bez wyjątku i bez osieroconych partów — scenariusz „użytkownik skasował marker sekcji", wpis dla sekcji 0 ignorowany — sekcja 0 = pola bazowe).
### Verified
- `dotnet test` Application.UnitTests **306/306** (było 302), Infrastructure.UnitTests **265/265** (było 253) — 0 niepowodzeń, zero zmian w kodzie produkcyjnym.
### Notes
- Luki pokryte wg diffu bieżącej pracy (ADR-0025/0026): kontrakt handlerów na nowy parametr był nietestowany (regresja = ciche spłaszczenie nagłówków sekcyjnych przy autosave), podobnie odporność writera na wpis sekcyjny wskazujący nieistniejącą sekcję.
- Nadal niepokryte (świadomie): warunkowe formatowanie TEKSTU ze stylu tabeli, regiony narożne i przekątne (nieobsłużone w kodzie — patrz ADR-0026), leader nierysowany w edytorze.

## 2026-07-05 — Audyt zgodności edytora z MS Word (tylko dokumentacja, zero zmian w kodzie)
### Changed
- Nowy `.ai/AUDIT_WORD_COMPATIBILITY.md`: kompletny audyt pipeline DOCX→HTML→edycja→DOCX (~60 problemów, każdy z kategorią, reprodukcją, przyczyną, warstwą, ryzykiem i kierunkiem naprawy; podział potwierdzone/potencjalne + 6-etapowa kolejność napraw). Wpis w `INDEX.md`.
### Verified
- Fakty zweryfikowane w kodzie (`DocxToHtmlConverter`, `HtmlToDocxConverter`, `wysiwyg-editor.ts`, handlery Save/Sign/Finish/Download) — kluczowe: cichy drop nieznanych elementów w readerze + pełna regeneracja pakietu w autosave/„Zakończ i wyślij" (pass-through tylko w Download); `w:ins` (wstawiony tekst) znika; pola ≠ PAGE/NUMPAGES/DATE tracą kod i wartość (TOC!); txbx/oMath/wykresy/kształty drop; Sign bez `SectionHeadersFooters`.
### Notes
- Bez zmian w kodzie i testach. Najpilniejsze wg audytu: Etap 1 „stop utracie danych" (w:ins, wartości pól, txbx, zakładki, vanish, keep*, Sign, marker sekcji contenteditable=false, ostrzeżenie o nieobsługiwanych elementach).

## 2026-07-05 — Tabele: styl tabeli (tblStyle/tblLook/tblStylePr), pozycyjne krawędzie, wysokości wierszy, round-trip (ADR-0026)
### Changed
- **Reader (`DocxToHtmlConverter`)**: rozwiązywanie stylu tabeli (`ResolveTableStyleContext`: łańcuch basedOn per-side, flagi tblLook, formaty warunkowe firstRow/lastRow/firstCol/lastCol/pasy z pominięciem nagłówka), pozycyjne krawędzie komórek (outer vs insideH/insideV), `themeFill`/`themeColor`+tint/shade, wzory `pctNN` (blend), `sz/6` px, `height:`+`data-row-height-tw`/`data-row-hrule` na tr, `data-tbl-header`/`data-cant-split`, `tblCellSpacing`→`border-collapse:separate`+`border-spacing`+data-*, `%` z ułamkiem, cell width auto/nil bez `width:0px`, `data-tbl-style`/`data-tbl-look` na `<table>`.
- **Writer (`HtmlToDocxConverter`)**: `w:tblStyle` (pierwsze dziecko tblPr) + `w:tblLook` z data-*, trHeight z twips+hRule (data-* preferowane nad px), `w:tblHeader`/`w:cantSplit`, `w:tblCellSpacing`, `w:tblInd` z margin-left px, px×6 dla wszystkich borderów (symetria z readerem), dziesiętne szerokości %/px, fix zagnieżdżonych tabel (`.//tr` duplikowało wiersze).
- **GUI**: `_serializeSingleEditor` nie zapieka już mierzonych wysokości wierszy (R-18 — eksportowane tylko jawne inline: import/ręczny resize); resize wiersza czyści `data-row-height-tw`/`data-row-hrule`; nowy `core/utils/table-grid.util.syncTableColgroup` wołany po resize kolumn/tabeli (wysiwyg-editor) i po wstawieniu/usunięciu kolumny (document-editor) — `w:tblGrid` podąża za edycją.
### Verified
- Nowe `TableStyleFidelityTests` 21/21 (styl/basedOn/tblLook/banding/theme/pattern/priorytety/wysokości/spacing/nested/round-trip); `dotnet test D2ViewerEditor.sln` — 0 niepowodzeń (Infrastructure 253/253); Golden `simple-table`/`merged-cells-table`/`tab-stops-lcr` zregenerowane.
- GUI: `table-grid.util.spec` 5/5; pełne `ng test` 317/318 (fail wyłącznie pre-existing `spec-layout-shell`); `ng build` OK (tylko pre-existing budżet scss).
### Notes
- Kontrolowane przybliżenie: formatowanie ze stylu wraca do DOCX jako bezpośrednie (wygląd w Wordzie identyczny; referencja stylu zachowana w `w:tblStyle`). Nieobsłużone (udokumentowane w ADR-0026): warunkowe formatowanie tekstu ze stylu, regiony narożne, przekątne tl2br/tr2bl, fitText/hideMark, powtarzanie wiersza nagłówkowego w paginacji edytora.

## 2026-07-05 — Nagłówki/stopki: dynamiczne pasmo, tab-stopy pozycyjne (w:tabs per akapit), nagłówki/stopki PER SEKCJA (ADR-0025)
### Changed
- **Dynamiczne pasmo nagłówka/stopki (GUI, geometria Worda)**: `.page-header` zaczyna się teraz `headerDistance` (w:pgMar header) od krawędzi strony (`margin-top`), ma `min-height` = pasmo (margines − dystans) i **rośnie z treścią spychając body** — wcześniej pasmo rysowało się od samej krawędzi (za wysoko), a treść dostawała sztywny padding liczony ze stałego pasma. `PageGeometry` +`headerDistanceCm`/`footerDistanceCm` (z `data-header/footer-distance-cm` markera sekcji; sekcja 1 = odwrotność wzoru readera: dystans = margines − pasmo). `_repaginateNow` mierzy REALNE wysokości pasm z DOM (`_measureBandHeightsPx`, osobno strona 1 / reszta) i odejmuje je od dostępnej wysokości — nagłówek z obrazkiem/kilkoma liniami odbiera miejsce treści zamiast rozjeżdżać strony.
- **Tab-stopy per akapit**: reader czyta efektywne `w:tabs` (łańcuch stylów + direct pPr, semantyka `clear`) i ZAWSZE emituje `data-tab-stops="pos:align[:leader]"` (twips). W nagłówku/stopce akapit z tabulatorami renderuje się **pozycyjnie**: segmenty na pozycjach stopów (`span.docx-tab-seg`, center = `translateX(-50%)` NA pozycji, right = `translateX(-100%)`), zamiast przybliżenia flex 50%/100% (flex zostaje dla body). Writer: `data-tab-stops` → `w:tabs` w pPr (pozycje/wyrównania/leadery wracają per akapit, nie tylko sztywne 4536/9072 w stylach Header/Footer); `span.docx-tab-seg` → `w:tab` + treść; **literalny `\t` w tekście → element `w:tab`** (wcześniej trafiał do `w:t`, którego Word nie renderuje). Refaktor: `BuildRunWrapper`/`ConvertRunChildToHtml` wydzielone z `ConvertRunToHtml` (segmentacja domyka i reotwiera formatowanie runu wokół segmentów).
- **Nagłówki/stopki per sekcja (zmiana kontraktu)**: nowy typ `SectionHeaderFooter { SectionIndex, Header, Footer }`; `DocumentContent.SectionHeadersFooters` + `SaveDocumentRequest.SectionHeadersFooters` (Domain + model TS). Reader: wpis dla każdej sekcji ≥ 1 z WŁASNYMI referencjami (default/first/even + geometria pasma tej sekcji); sekcje bez wpisu dziedziczą (jak Word). Writer: `Convert`/`ConvertPreservingPackage` +param; wpisy lądują na sectPr SWOJEJ sekcji (`_emittedSectionProps[s]` / body-level dla ostatniej); baza (sekcja 0) bez zmian na pierwszym sectPr. Handlery: `SaveDocument`, `DownloadEditedDocument` (+ kontrolery). GUI: `pageSectionIndexes` (strona→sekcja, split + repaginacja), `_computeHeader/FooterContent` wybiera wpis sekcyjny (dziedziczenie po max indeksie ≤ sekcji strony), **edycja pasma na klikniętej stronie** (`editingHfPageIndex`, template `$index === editingHfPageIndex()`), routing zapisu do właściciela (`_applyEditedHeader/FooterHtml` — wpis sekcji / wariant first-page / default; ujednolicone też w `onContentChange`/blur), output `sectionHeadersFootersChange` → `document-editor.sectionHeadersFooters` → `buildSaveRequest`.
### Verified
- Backend: Infrastructure **232/232** (nowe: `TabStopFidelityTests` 7, `MultiSectionFidelityTests` +4 sekcyjne nagłówki — w tym pełne round-tripy DOCX→HTML→DOCX); pełna solucja Domain 79 / Application 302 / Api 130 / Integration 9 — 0 fail. Snapshoty golden tabel odświeżone (zmiana z równoległej pracy nad stylami tabel — dokładne szerokości ramek 0.7px; przejrzane).
- GUI: `ng test` **312/313** (+9 nowych: dynamiczne pasmo 3, sekcyjne nagłówki 6; fail tylko pre-existing `spec-layout-shell` NG0201). `npm run build` OK (warningi budżetów pre-existing). `tsc` czysto.
### Notes
- Kontrakt: `sectionHeadersFooters` to pole OPCJONALNE (dokumenty jednosekcyjne — bez zmian w payloadach). Ścieżka Sign nie przenosi sekcyjnych nagłówków (świadomie odłożone).
- Ograniczenia: przypisanie k-ty tab → k-ty stop (bez pełnej semantyki „następny stop za bieżącą pozycją x"); leader (kropki) nie jest rysowany w edytorze (wraca do DOCX); akapity ze złożonymi polami (PAGE w fldChar) zostają na flexie.

## 2026-07-05 — Edytor: fix paginacji — ENTER przy dolnej krawędzi tworzy nową stronę zamiast rozciągać kartkę
### Changed
- **`wysiwyg-editor.ts` — pomiar bloków w `_repaginateNow`** (root cause):
  - bloki były mierzone w gołym `<div>` doczepionym do `document.body`, POZA kontekstem stylów `.editor-content` — pusty `<p>` mierzył **0 px** (reguła `.editor-content p:empty::before {content:'\00a0'}` nie działała w measurerze), znikały też scope'owane marginesy akapitów/nagłówków/tabel;
  - dodatkowo `getBoundingClientRect().height` **nie zawiera marginesów** (−10 px na akapit, −30 px na nagłówek/tabelę) → suma pomiarów zaniżona → `currentHeight + h > availableHeight` nie odpalał → wszystkie bloki zostawały na stronie, a `.page` (tylko `min-height`, `overflow:visible`) rosła w pion.
  - Fix: `_createBlockMeasurer` — measurer z klasą `editor-content` (style globalne, `ViewEncapsulation.None`) + `_measureBlockRunHeights` — wysokość KONSUMOWANA liczona deltami pozycji kolejnych bloków (+ sentinel 0 px): zawiera marginesy i ich realny kolaps, margin-top 1. bloku doliczany; nadal 1 append + 1 layout-flush na przebieg (bez regresji wydajności). Fragmenty tabel (`_splitTableForPagination`) mierzą teraz w kontekście stylów (paddingi `td` liczone poprawnie).
- **`_schedulePaginate` — max-wait 600 ms**: czysty debounce 250 ms resetował się na każdym `input`, więc PRZYTRZYMANY Enter (auto-repeat ~30 ms) odsuwał repaginację w nieskończoność; teraz repaginacja odpala najpóźniej 600 ms od pierwszego zaplanowania.
- Kursor i tworzenie kolejnej strony NIE wymagały zmian: kotwica `{block, offset}` (`_saveGlobalCaret`/`_restoreGlobalCaret`) przenosi karetkę na nową stronę, `pageEditorRefs.changes` podpina listenery nowym stronom — problemem był wyłącznie pomiar przepełnienia.
### Verified
- Nowy `wysiwyg-editor.pagination-overflow.spec.ts` **9/9**: measurer w kontekście `.editor-content`, pomiar z marginesami (delty pozycji), 5 przepełnionych bloków → 5 stron, pusty akapit po ENTER przenoszony na nową stronę, kolejność bez utraty/duplikacji, scalanie w górę po usunięciu bloków, max-wait przy ciągłym strumieniu input, zwykły debounce 250 ms.
- Specy wysiwyg **69/69**; pełne GUI **312/313** (jedyny fail — pre-existing `spec-layout-shell`); `ng build` OK (warningi budżetów pre-existing). Backend nietknięty.
### Notes
- `.page` celowo zostaje na `min-height` (nie sztywny `height`+clip): między wejściem a repaginacją (≤600 ms) treść nie może zniknąć spod kursora; po repaginacji strona wraca do formatu.
- Ograniczenia bez zmian: blok wyższy niż strona (poza tabelą) nie jest dzielony (block-atomic); bloki out-of-flow (absolute/float) na najwyższym poziomie mierzą 0 konsumpcji (jak w realnym układzie).

## 2026-07-05 — Listy wielopoziomowe: logiczna tożsamość + liczniki Worda + round-trip formatów (numPr/abstractNum)
### Changed
- **Reader (`DocxToHtmlConverter`)** — semantyka numeracji Worda odtworzona jawnie:
  - liczniki per **(abstractNumId, ilvl)** (`_listCounters`): różne `w:num` na wspólny abstrakt KONTYNUUJĄ numerację; `w:startOverride` restartuje poziom przy pierwszym użyciu instancji (`ApplyStartOverridesOnFirstUse`); element płytszy restartuje poziomy głębsze (honorowane `w:lvlRestart=0`); `w:numStyleLink` rozwiązywany (`ResolveAbstractNumId`).
  - `<ol start>` = FAKTYCZNY numer pierwszego elementu z liczników (kontynuacja po przerwaniu akapitem / współdzielony abstrakt), nie sama definicja `w:start`.
  - grupowanie `ConvertConsecutiveListItems`: tożsamość listy = **numId** (koniec sklejania niezależnych list o identycznym wyglądzie; usunięty warunek „ten sam format ⇒ ta sama lista").
  - kontener `ol/ul` niesie **data-num-id / data-abstract-num-id / data-ilvl / data-num-fmt / data-start / data-lvl-text / data-bullet-font** (wspólny resolver `FindLevelDefinition`: lvlOverride/w:lvl instancji → abstrakt).
- **Writer (`HtmlToDocxConverter`)**:
  - `_numIdByHtmlList`: fragmenty listy z tym samym `data-num-id` (lista przerwana akapitem) współdzielą JEDNĄ `NumberingInstance` → Word kontynuuje numerację; różne `data-num-id` → osobne instancje.
  - `CreateAbstractNumbering` przyjmuje specyfikacje poziomów z data-* (`ScanListLevelSpecs`/`HtmlListLevelSpec`): dokładny `w:numFmt` (decimal/…/upperRoman/bullet/none), `w:lvlText` (np. "%1)"), `w:start`, font punktatora — zamiast hardkodowanej drabinki formatów, która przy każdym autosave niszczyła oryginalną numerację.
### Verified
- Nowy `ListNumberingFidelityTests` **12/12**: zagnieżdżenie z kontynuacją głównego poziomu, kontynuacja po zwykłym akapicie (start=3), niezależne listy NIE sklejane, `w:start=5`, wspólny abstrakt = kontynuacja, `startOverride` = restart, restart głębszego poziomu po powrocie na poziom główny, formaty mieszane per poziom (upperRoman "%1)" / bullet / lowerLetter), writer: wspólny/rozdzielny numId, round-trip formatu do abstractNum, pełny round-trip DOCX→HTML→DOCX→HTML (struktura + kontynuacja + tożsamość zachowane).
- Infrastructure 225/226 w trakcie zmian (jedyny fail: snapshot `tab-stops-lcr` — `data-tab-stops` z RÓWNOLEGŁEJ sesji, baseline do regeneracji; niezwiązane z listami). Pełny przebieg po ustabilizowaniu drzewa — patrz kolejny wpis/TASK_HANDOFF.
### Notes
- Ograniczenia (HTML/przeglądarka): sufiks markera z `lvlText` ("1)" vs "1.") nie jest renderowany wizualnie w edytorze (CSS `list-style-type`); numeracja złożona multilevel ("1.1.2") nierenderowana; oba wracają do DOCX bez strat przez data-*. Listy NOWE z edytora (bez data-*) dostają domyślną drabinkę formatów jak dotąd. Egzotyczne `w:lvlRestart>0` uproszczone do zachowania domyślnego.

## 2026-07-05 — Audyt zgodności dokumentacji `.ai` z kodem (tylko docs, zero zmian w kodzie)
### Changed
- **TECH_STACK**: auth = Microsoft.Identity.Web 3.12.0 (nie pinowany JwtBearer 8.0.12), SkiaSharp 3.119.2, + NPOI/System.Security.Cryptography.Xml/Serilog, pełna lista projektów testowych, + `D2ExampleExternalApp`.
- **PROJECT_CONTEXT/ARCHITECTURE**: + `D2ExampleExternalApp` i projekty testowe (Api.IntegrationTests, Services.Api.UnitTests); zaktualizowana realna struktura frontendu (guards/, nowe komponenty/pages/core-utils).
- **API_CONTRACTS**: external `GET /status` zwraca okrojone `{masterId,status}` (nie pełny DTO); `Classification` opcjonalna; + `POST /api/barcode/generate-image`; + `GET /api/identity/resources`; autoryzacja klasowa `RequireAppOperator` (ADR-0022).
- **DATABASE**: skrypty 008–011; statusy `Queued`/`SendAborted`/`Cancelled`; kolumny `last_modified_by`/`corporate_key`.
- **SECURITY**: guardy frontu (bez `appAdminGuard`), claim `corpKey` case-insensitive (nie `ck`), centralne polityki uploadu + SSRF `ReturnUrlValidator` (zamiast „rekomendacji"), wysyłka multipart, trasa `/brak-uprawnien`.
- **DOMAIN/GLOSSARY**: pełne enumy statusów, metody inline-delivery/cancel/`LastModifiedBy`, klasyfikacja opcjonalna (BR-006), CorporateKey z tokenu.
- **FEATURES/PRODUCT_GOALS**: Krok 2 = Implemented (podgląd ładuje v1 — R-02 zamknięte); „Zakończ" = synchroniczna 1. próba (200, nie 202) + abort/continue; + wiersze RBAC obszarów, security upload, userDownload/showSaveState, external callback-url/unlock/status, .doc, grafiki, multi-section, LastModifiedBy.
- **TESTING_QUALITY**: + projekty integracyjne i Services.UnitTests.
- **OBSERVABILITY**: obejmuje oba hosty; pola ECS+GCP (obiekt `service{}` — migracja 2026-06-28), redaction, `StructuredLogFormatterOptions`.
- **DOCX_CONVERSION**: sekcja 2 i roadmapa R-10 dostosowane do ADR-0023 (koniec `FirstOrDefault`, markery sekcji; Open: nagłówki per sekcja).
- **RISKS_ASSUMPTIONS**: R-02→Closed, A-05→Closed, A-03 001..011, A-06 multipart potwierdzone, duplikat R-25→R-29.
- **DECISIONS**: drugi „ADR-0020" (centralne polityki security) przenumerowany na **ADR-0024**.
- **TASK_HANDOFF**: „Aktualne zadanie"/„Następne kroki"/„Niedokończone" odświeżone (Krok 1–4 done; otwarte R-06/R-07/R-28/R-09 + nagłówki per sekcja).
- **Diagramy/Confluence** (`sequence-document-lifecycle`, `external-api-browser-open-diagram`, `confluence-*`): erraty + poprawki: auth istnieje (Entra + role + allowedCorporateKeys), finish 200 sync, multipart potwierdzony, R-02 zamknięte.
- **INDEX**: + 4 brakujące pliki (diagramy, confluence).
### Verified
- Fakty porównane z kodem: `*.csproj` (wersje pakietów/TFM), `package.json`, kontrolery obu API (routingi), `app.routes.ts` + `guards/`, `infra/sql/001..011`, `HttpDeliverySender` (multipart), `ClaimsCurrentUserProvider`/`AzureAdOptions` (corpKey), `document-editor.loadFromStorage` (v1 w podglądzie), enumy `DocumentStatus`/`DeliveryStatus`, `GcpJsonConsoleFormatter`/`StructuredLogFormatterOptions`, `Document`/`DocumentDelivery` (metody domenowe).
### Notes
- Bez zmian w kodzie/konfiguracji. Logi historyczne (`CURRENT_STATE`/`CHANGELOG`/stare wpisy ADR) pozostawione bez retro-edycji — opisują stan na swoją datę.

## 2026-07-05 — Edytor: rendering stron per sekcja (domknięcie ADR-0023) + realny PageSize + szybsza repaginacja
### Changed
- **`wysiwyg-editor` — geometria per strona**: nowy input `pageSize` (z `documentPageSize()` w `document-editor.html`) + `PageGeometry` i sygnał `pageGeometries` (per strona). `baseGeometry` = pageSize/pageMargins/pageOrientation (orientacja z toolbara wygrywa — wymiary obracane); kolejne sekcje czytane z `data-*` markera `div.docx-section-break` (`_parseSectionGeometry`, brakujące atrybuty dziedziczą). Marker OTWIERAJĄCY stronę zmienia geometrię od tej strony; marker w środku strony (continuous) — od następnej. Szablon binduje per strona: `width/min-height` (px), klasę `landscape`, paddingi treści i prowadnice marginesów. **Koniec renderowania wszystkich stron w sztywnej geometrii A4 pierwszej sekcji** — dokument pionowy z poziomym aneksem renderuje i łamie strony we właściwych wymiarach; dokumenty A5/Letter dostają realną wysokość łamania (wcześniej hardkody 1122/816 px i „21 vs 29.7").
- **`_repaginateNow`**: dostępna wysokość i szerokość pomiaru z geometrii bieżącej sekcji (skala szerokości z DOM strony 1 → proporcjonalnie dla sekcji); **pomiar wsadowy** ciągłych przebiegów zwykłych bloków (`measureRun`: jeden append + jeden layout-flush na przebieg zamiast reflow per blok — istotne przy dużych dokumentach; tabele mierzone jak dotąd).
- **Guard**: `_flattenTopBlocks` nie rekursuje do wnętrza markerów `page-break`/`docx-section-break` (strona z samym markerem gubiła go przy repaginacji); `emitSectionGeometry` i `getContentAreaHeight` liczą z geometrii strony (wcześniej stałe A4).
### Verified
- GUI: specy `wysiwyg-editor` **51/51** (nowy blok „geometria stron per sekcja": 6 testów — baseGeometry/pageSize/orientacja, marker otwierający stronę, continuous od następnej strony, dziedziczenie data-*, guard flatten); pełne `ng test` **294/295** (jedyny fail pre-existing `spec-layout-shell`); `ng build` OK (warning budżetu SCSS pre-existing).
- Backend bez zmian w tym kroku; pełna solucja **730/730** (Domain 79, Application 302, Api 130, Integration 9, Infrastructure 210).
### Notes
- Nagłówki/stopki per sekcja nadal wspólne (model `DocumentContent` niesie jeden komplet — rendering treści R-10 domknięty, warianty header/footer per sekcja = osobny krok). Edycja przy granicy sekcji może skasować marker (degradacja do jednej sekcji — bez zmian względem ADR-0023).

## 2026-07-05 — DOCX↔HTML: wiele sekcji (R-10 read+write), pierwsza sekcja jako bazowa, tabele (tblGrid/fixed/vMerge), interlinia atLeast
### Changed
- **Reader (`DocxToHtmlConverter`)**: nowy `GetSectionPropertiesInDocumentOrder` — `Body.Elements<SectionProperties>().FirstOrDefault()` zwracało sectPr OSTATNIEJ sekcji (body-level), więc dokument pionowy z poziomym aneksem otwierał się cały poziomo, z nagłówkami aneksu. Teraz `PageSize`/`Margins`/wysokości pasm liczone z PIERWSZEJ sekcji; nagłówek/stopka: pierwsza sekcja z referencją Default (kolejność dokumentu), fallback bez zmian.
- **Reader**: paragraph-level `pPr/sectPr` (koniec sekcji) emituje parę: `div.page-break` (dla przerw nextPage/oddPage/evenPage) + niewidoczny marker `div.docx-section-break` z geometrią NASTĘPNEJ sekcji w `data-*` (`data-break-type`, `data-page-width/height-cm`, `data-orientation`, `data-margin-*-cm`, `data-header/footer-distance-cm`).
- **Writer (`HtmlToDocxConverter`)**: marker odtwarza `w:p/pPr/sectPr` (geometria sekcji zamykanej; `w:type` z markera otwierającego; dystanse header/footer wprost z data-*), body-level sectPr dostaje geometrię OSTATNIEJ sekcji; `page-break` bezpośrednio przed markerem NIE staje się `w:br type=page` (sectPr sam łamie stronę); referencje nagłówka/stopki + `titlePg` idą do PIERWSZEGO sectPr (dziedziczenie na kolejne sekcje). Koniec spłaszczania dokumentu wielosekcyjnego do jednej sekcji przy autosave.
- **Writer tabele**: `w:tblGrid` budowany z `<colgroup>` (szerokości px→twips; wcześniej pusty grid z samej liczby komórek — szerokości kolumn ginęły przy każdym zapisie); `table-layout:fixed` → `TableLayout Fixed` (wcześniej zawsze Autofit); komórki kontynuacji `w:vMerge` (bez val) wstawiane pod rowspan z pozycjonowaniem po kolumnie gridu (wcześniej brak → komórki przesuwały się w lewo, tabela uszkodzona w Wordzie).
- **Interlinia**: reader oznacza `lineRule=atLeast` markerem CSS `--w-line-rule:atLeast;` przy `line-height:Xpt`; writer odtwarza `AtLeast` (wcześniej każde pt → `Exact`, co przycina w Wordzie tekst wyższy niż linia).
- **GUI**: `wysiwyg-editor.scss` ukrywa `.docx-section-break` (`display:none`) — marker to nośnik danych, przeżywa split stron (nie ma klasy `page-break`, więc splitter go nie kanonizuje) i wraca w `getContent()`.
### Verified
- Backend: `Infrastructure.UnitTests` **210/210** (nowe: `MultiSectionFidelityTests` 12, `TableWriteFidelityTests` 6, `LineSpacingMappingTests` +3); pełna solucja: Domain 79, Application 302, Api 130, IntegrationTests 9 — 0 fail; snapshoty Golden bez zmian (brak regresji single-section).
- GUI: `ng test` **288/289** (+1 nowy test przeżywalności markera sekcji; jedyny fail `spec-layout-shell.spec` — pre-existing, NG0201 HttpClient w teście dashboardu, niezwiązany). `npm run build` OK (warningi budżetów pre-existing).
### Notes
- Kontrakty API bez zmian — sekcje jadą w polu `Html` jako markery; `DocumentContent.PageSize/Margins/Header/Footer` = pierwsza sekcja.
- Ograniczenie: edytor nadal RENDERUJE wszystkie strony w geometrii pierwszej sekcji (per-page orientacja to zmiana paginacji w `wysiwyg-editor`); dane sekcji są już jednak zachowywane w round-tripie. Decyzja: **ADR-0023**.

## 2026-07-03 — SECURITY: egzekwowanie roli `Operator` na backendzie edytora (broken access control)
### Changed
- **Backend (krytyczne):** `DocumentController` i `DocumentStorageController` dostały klasowy `[Authorize(Policy = RequireAppOperator)]`. Wcześniej dziedziczyły tylko `[Authorize]` (dowolny zalogowany), więc użytkownik **bez roli aplikacyjnej** mógł przez bezpośrednie wywołanie API tworzyć/zapisywać/nadpisywać/przywracać/podpisywać i „Zakończyć i wysłać" dokumenty. Endpointy admina zachowują `RequireAppAdmin` (AND z polityką klasy → wymagany Administrator).
- **Frontend (defense-in-depth + UX):** pulpit (`dashboard`) bramkuje akcje „Nowy dokument"/„Otwórz plik" sygnałem `canUseEditor` (`ResourceAccessService.hasAccessToResource('editor')`); `newDocument()`/`openFile()` mają twardy `ensureEditorAccess()`. `resourceGuard` rozróżnia stan `loading`/transient (401/0/5xx → retry ×3, 300 ms) od definitywnego `403` (deny → `/brak-uprawnien`) — bez traktowania niegotowego MSAL jako „brak uprawnień".
### Verified
- Nowy projekt `D2ViewerEditor.Api.IntegrationTests` (WebApplicationFactory<Program> + `TestAuthHandler`): **9/9** — 401 bez tokenu (editor/save/admin), 403 bez roli i z obcą rolą, 200 dla `Operator`/`Administrator`, 403 `Operator` na endpoint admina.
- GUI: `resource.guard.spec` 3, `dashboard.spec` 3 (gating), `resource-access.service.spec` 3 — pass. Build API 0 błędów; build GUI OK.
### Notes
- Do potwierdzenia w Entra: przypisania App Roles `Operator`/`Administrator`, „Assignment required=Yes", redirect URI (SPA). Decyzja: **ADR-0022**. `main`/interceptor/MSAL bez zmian (kolejność inicjalizacji poza zakresem tej poprawki).

## 2026-06-28 — Fix: okno podwójnej wysyłki w „Zakończ" (Pending-due przed Sending)
### Changed
- `FinishAndSendDocumentCommandHandler`: nowy rekord wysyłki nie jest już commitowany jako `Pending` z `next_attempt_at=now()` przed próbą inline. `CreateQueuedDeliveryAsync` robi tylko `AddAsync` (tracked, bez SaveChanges); pierwszy commit następuje w `AttemptInlineDeliveryAsync` po `BeginInlineAttempt()` — INSERT od razu jako `Sending`. Rekord nigdy nie istnieje w bazie jako claimowalny `Pending`, więc worker tła (`ClaimDueBatchAsync`: `Pending/RetryScheduled AND next_attempt_at<=now`) nie przejmie go i nie wyśle drugi raz. Usunięto martwe `document.MarkQueued()`.
### Verified
- `Application.UnitTests` 302/302 (FinishAndSend 11, w tym nowy `Handle_NewDelivery_IsPersistedAsSending_NeverClaimablePending`).
### Notes
- Kompromis: znika okno podwójnej wysyłki; w zamian crash między uploadem snapshotu a jedynym commitem = osierocony snapshot w GCS + brak rekordu (okno sub-ms; duplikat zewn. wysyłki gorszy niż leak snapshotu). Inline-`Sending` ma `locked_until=NULL` → worker nie reclaimuje w trakcie wolnego `SendAsync`. At-least-once (timeout-on-success + „Kontynuuj"/worker) nadal możliwe — chroni `Idempotency-Key` (odbiorca dedupuje). Powtarzalny dubel u użytkownika prawdopodobnie we wdrożonym `HttpDeliverySender` (OGate/Keycloak, `CreateClient("okapi")`) — poza repo.

## 2026-06-28 — GcpJsonConsoleFormatter: bogate logi GCP + ECS, structured exceptions, redaction
### Changed
- `GcpJsonConsoleFormatter` przepisany na małe funkcje. Dodane: pola GCP (`logging.googleapis.com/trace|spanId|trace_sampled|sourceLocation`, `httpRequest`, `labels`), pola ECS (`@timestamp`, `log.*`, `service.*`, `trace.*`, `span.*`, `transaction.*`, `event.*`, `error.{type,message,stack_trace,inner}`, `http.*`, `url.*`, `user.id`), klasyfikacja błędu (`event.reason`: database/dependency/validation/authorization/cancellation/code), korelacja (`HttpContext.Items`/header/scope/Activity baggage) jako `correlation_id`+`labels.correlation_id`, maskowanie pól wrażliwych ("[REDACTED]", case/separator-insensitive) i query.
- `StructuredLogFormatterOptions` rozbudowane: ProjectId, ServiceName, ServiceVersion, EnvironmentName, ServiceInstanceId, Include{Scopes,EventId,SourceLocation,HttpRequest,ElasticCommonSchemaFields,GoogleCloudFields}, RedactedPropertyNames.
- `IHttpContextAccessor` (opcjonalny) wstrzykiwany do formattera; działa też bez HTTP (worker). `LoggingExtensions` wypełnia ProjectId (`GOOGLE_CLOUD_PROJECT`), ServiceVersion, instance id (`K_REVISION`/`HOSTNAME`).
### Verified
- `Api.UnitTests` 129/129 (formatter 21, w tym GCP trace, ECS, redaction, http/no-http, special chars, system-field-not-overwritten).
### Notes
- BREAKING (opisane): przy ECS ON (default) `service`/`environment` przeniesione do obiektu `service{name,version,environment}` — zmigrowany 1 test. Konflikt GCP/ECS dla `service` rozwiązany hybrydowo: płaskie `severity`+nested `log.level`, płaskie `traceId`+`trace.id`. `event.reason` to celowe odejście od oficjalnego ECS (low-cardinality triage 500).

## 2026-06-27 — Wysyłka: „w toku" ≠ błąd, zamykanie okna, anulowanie w trakcie
### Changed
- Backend: `DocumentDelivery.CancelByUser()` (anulowanie z edytora obejmuje też próbę INLINE w toku — Sending bez lease; worker-Sending z lease i stany końcowe nadal rzucają). `AbortSendCommandHandler` używa `CancelByUser()`. `FinishAndSendDocumentCommandHandler` łapie `OperationCanceledException` jako anulowanie (propaguje, NIE zamienia na `Result.Failure`/`DeliveryFailed`).
- Frontend (`document-editor`): `finishDocument()` rozróżnia trzy stany — `delivered` (koniec), `status==='Sending'` (stan przejściowy, **okno informacyjne, nie błąd**), realny błąd (modal problemu). Nowe `cancelSend()` („Przerwij wysyłkę" w trakcie — anuluje request w locie + `abortSend`, stan „Cancelled", bez błędu) i `closeSendingModal()` („Zamknij" — leci dalej w tle, UI odpięte przez guard `sendDetachedFromUi`). Modal wysyłki dostał komunikat + przyciski Przerwij/Zamknij; styl `.finish-dialog-text`.
### Verified
- Backend: Domain `DocumentDelivery` 27/27, Application `FinishAndSend`+`AbortSend` 14/14.
- Frontend: `document-editor.spec` 71/71 (+3: status Sending bez błędu, cancelSend, closeSendingModal+guard).
### Notes
- Projekt nie ma i18n — teksty po polsku, spójnie z resztą.
- Ryzyko/follow-up: przy abort-during-inline-Sending request inline i `abortSend` biegną równolegle; inline po `OperationCanceled` nie zapisuje już stanu, a `abortSend` (CancelByUser) ustawia Cancelled — brak konfliktu zapisu, ale to do potwierdzenia na realnym Postgresie (R-06).

## Entries

## 2026-06-26 — corpKey wyłącznie z tokenu po stronie API (usunięcie transportu z GUI)
### Changed
- **API**: komendy `SaveDocumentVersionCommand`/`UpdateDocumentVersionCommand`/`FinishAndSendDocumentCommand` — usunięty parametr `CorporateKey`. Handlery ustalają `corporateKey = _currentUser.CorporateKey` (zweryfikowany token; brak → `Result.Failure`). Controller DTO `SaveDocumentVersionRequest` — usunięte pole `CorporateKey`; akcje `save`/`update`/`finish` nie przekazują już tej wartości.
- **GUI**: usunięty sygnał `corporateKey` i odczyt claimu `corpKey` z `idTokenClaims` w `document-editor.ts`; `saveDocumentVersion`/`updateDocumentVersion`/`finishAndSend` wysyłają już tylko `{ content }`. Interfejs `SaveDocumentVersionRequest` (GUI) — usunięte `corporateKey?`.
### Verified
- backend build 0 błędów; `Application.UnitTests` 295/295; `Api.UnitTests` 113/113; GUI `npm run build` OK.
### Notes
- Powód: GUI czytało `corpKey` ze statycznego snapshotu `getActiveAccount().idTokenClaims` (często null), przez co `SaveDocumentVersionCommand.CorporateKey` przychodził null. Token wysyłany do API (ID token, `api-token.interceptor.ts`) i tak niesie `corpKey`, więc `_currentUser.CorporateKey` jest niezawodnym, jedynym źródłem. Wyświetlanie `corporateKey` na listach admina (z `DocumentDelivery`) bez zmian. Decyzja: ADR-0021 (zastępuje transport z 2026-06-24).

## 2026-06-26 — Naprawa zapisu edytującego (corpKey) + „Kopiuj link" w admin-files
### Changed
- API `AzureAdOptions.CorporateKeyClaim`: domyślnie `"ck"` → `"corpKey"` (zgodność z claimem w tokenie Entra).
- API `HttpHeaderCurrentUserProvider`: fallback `DEV-LOCAL`, gdy brak nagłówka `X-Corporate-Key` (dev bypass nigdy nie zwraca null).
- API handlery `SaveDocumentVersion`/`UpdateDocumentVersion`: wstrzyknięcie `ICurrentUserProvider`; wszystkie trzy ścieżki zapisu (`Save`/`Update`/`FinishAndSend`) ustalają `corporateKey = request.CorporateKey ?? _currentUser.CorporateKey` i zwracają `Result.Failure`, gdy nie da się ustalić użytkownika.
- GUI `admin-files`: przycisk „Kopiuj link" w szczegółach pozycji (kopiuje `/editor?masterId&versionId` do schowka, baner `notice`). Brak przycisku na samej liście.
### Verified
- backend build 0 błędów; `Application.UnitTests` 295/295; `Api.UnitTests` 113/113; GUI build OK; `ng test` 268/269 (pre-existing `spec-layout-shell.spec`).
### Notes
- Decyzja: corpKey wysyła frontend (z `idTokenClaims`); backend traktuje token jako fallback. Poprawka claimu sprawia, że fallback działa też w prod.


## 2026-06-24 — corpKey z tokenu Entra ID + kolumna „Kto modyfikował" w administracji
### Changed
- **GUI → API**: edytor odczytuje claim `corpKey` z `idTokenClaims` aktywnego konta MSAL (`document-editor.ts`, sygnał `corporateKey`) i przekazuje go w body przy każdym zapisie z edytora: `saveDocumentVersion`, `updateDocumentVersion` (auto-save + nowy dokument) oraz `finishAndSend`. Interfejs `SaveDocumentVersionRequest` rozszerzony o opcjonalne `corporateKey`.
- **API**: `SaveDocumentVersionRequest` (controller DTO) +`CorporateKey` (opcjonalne); komendy `SaveDocumentVersion`/`UpdateDocumentVersion`/`FinishAndSendDocument` +`CorporateKey` (opcjonalne, fallback do claimu `_currentUser.CorporateKey`).
- **Domena**: `Document.LastModifiedBy` (+`SetLastModifiedBy`) — ustawiane przez handlery zapisu/aktualizacji/finish. Kolumna `documents.last_modified_by` (migracja `infra/sql/011_add_document_last_modified_by.sql`), mapowanie EF w `DocumentConfiguration`.
- **Listy admina**: `GetDocumentsQuery.DocumentListItemDto` +`LastModifiedBy`; `GetDeliveriesByStatusQuery.DeliveryListItemDto` +`CorporateKey` (reużycie istniejącego `DocumentDelivery.CorporateKey`). Dodano kolumnę „Kto modyfikował" (z filtrem) w `admin-files` (pokazuje `lastModifiedBy`) i `admin-deliveries` (pokazuje `corporateKey`).
### Verified
- `dotnet build D2ViewerEditor.sln` → 0 błędów. `D2Api.Api.UnitTests` **113/113**, `Application.UnitTests` **295/295**.
- `npm run build` (GUI) → OK. `ng test --watch=false` → **268/269** (jedyny fail: `spec-layout-shell.spec` — pre-existing, `NG0201 MsalService` w `DashboardComponent`, potwierdzony także bez zmian z tej sesji; niezwiązany).
### Notes
- Transport corporateKey = request body (decyzja agenta). Serwer rozwiązuje `request.CorporateKey ?? _currentUser.CorporateKey`, więc returnUrl callback nadal otrzymuje identyfikator nadawcy gdy GUI nie poda claimu.
- `SetLastModifiedBy` aktualizuje tylko jedną kolumnę na śledzonej encji (jak `MarkEditing`), omijając znany problem pełnego UPDATE z `created_at` Kind=Unspecified.

## 2026-06-22 — Grafiki: usunięcie widocznego placeholdera + cache + TIFF/ICO/WEBP
### Changed
- `GraphicConversionService` przepisany na czytelny łańcuch strategii (`Execute` → per-format) z raportowaniem strukturalnym. **Usunięto widoczny placeholder** (szare tło + tekst „… — podgląd w Word"): metafile EMF/WMF bez osadzonego rastra zwraca teraz **przezroczysty, pusty SVG** (`IsBlankFallback`), zachowujący wymiary z layoutu — zero udawanej treści. Oryginalny part nadal jedzie do DOCX (pass-through), więc Word renderuje wektor.
- Dodano **deduplikację po hashu treści** (SHA-256 + parametry → `ConcurrentDictionary`, bounded 1024) — identyczne assety konwertowane raz.
- Rozszerzono detekcję i obsługę: **TIFF** (→ PNG przez SkiaSharp, inaczej blank), **WEBP/ICO** (web-native passthrough), final raster-rescue dla nieznanych media partów przez SkiaSharp.
- `GraphicConversionModels`: `GraphicKind` +Webp/Ico/Tiff; `WebGraphicRepresentation.IsPlaceholder`→`IsBlankFallback`; `GraphicSource.SourcePath`; diagnostyka +`SourcePath/CacheKey/AttemptedStrategies/FailureReason`.
- `DocxToHtmlConverter`: `WebGraphicForLegacy` obejmuje też TIFF; atrybut `data-legacy-graphic="placeholder"`→`"blank"`; `IsMetafileContentType`→`IsNonBrowserNativeContentType` (+TIFF). Round-trip `data-original-src` bez zmian.
### Verified
- `dotnet test Infrastructure.UnitTests --filter GraphicConversion` → **35/35** (było 25). Pełny Infrastructure.UnitTests → **188/188**. `dotnet build D2ViewerEditor.sln` → 0 błędów.
- Nowe testy: detekcja TIFF/WEBP/ICO, no-placeholder (skan SVG/HTML: brak `<text`/`fill`/„requires conversion"/„podgląd w Word"), dedup po hashu, cache-key, TIFF→blank, integracyjne DOCX WMF / mixed native+EMF / duplicated EMF.
### Notes
- **Ograniczenie**: czysto wektorowy EMF/WMF (bez osadzonego rastra) nadal nie jest rasteryzowany w przeglądarce (brak pure-managed rasteryzera wektora bez GDI/LibreOffice — patrz ADR-0020) → przezroczysty blank w podglądzie + pełny render w Word z zachowanego oryginału. Rasteryzacja wektora = sidecar (roadmapa).

## 2026-06-22 — Security hardening: upload pipeline + returnUrl validation
### Changed
- Dodano centralne serwisy bezpieczeństwa w `D2ViewerEditor.Application/Common/Security`:
  - `IReturnUrlValidator` + `ReturnUrlValidator` + `ReturnUrlSecurityOptions` (normalizacja i walidacja callback URL: kontrola znaków, schematu, hosta, loopback/private IP, allowlista hostów),
  - `IFileUploadSecurityService` + `FileUploadSecurityService` + `UploadSecurityOptions` (extension↔MIME↔signature, inspekcja DOCX ZIP: zip-slip, limity entry/uncompressed/compression ratio, wymagane part-y OOXML, blokada makr VBA),
  - `IFileScanner` (`FileScanRequest/Result`) + `NoOpFileScanner` jako domyślna implementacja abstrakcji skanera.
- Reguły podpięte do krytycznych flow:
  - uploady: `UploadDocumentCommandHandler`, `UploadImageCommandHandler`, `IngestExternalDocumentCommandHandler`,
  - returnUrl/callback: `UpdateCallbackUrlCommandHandler`, `UpdateDeliveryRecipientUrlCommandHandler`, `FinishAndSendDocumentCommandHandler`.
- Hosty (`D2Api`/`D2Services`) bindują konfigurację: `Security:Upload` i `Security:ReturnUrl`.
- Zaktualizowano testy handlerów pod nowe zależności i dodano nowe testy security:
  - `ReturnUrlValidatorTests`,
  - `FileUploadSecurityServiceTests`.
### Verified
- `dotnet test D2ApiViewerEditor/D2ViewerEditor.Application.UnitTests/D2ViewerEditor.Application.UnitTests.csproj` → **295/295** pass.
- `dotnet test D2ApiViewerEditor/D2ViewerEditor.Api.UnitTests/D2ViewerEditor.Api.UnitTests.csproj` → **110/110** pass.
- `dotnet test D2ServicesViewerEditor/D2ServicesViewerEditor.Api.UnitTests/D2ServicesViewerEditor.Api.UnitTests.csproj` → **39/39** pass.
### Notes
- Zastosowano fail-fast na upload i callback URL jeszcze przed utrwaleniem danych.
- `NoOpFileScanner` jest punktem integracji; produkcyjny silnik AV należy podmienić przez DI bez zmiany kontraktu handlerów.

## 2026-06-22 — Pokrycie testami jednostkowymi ≥60% (wszystkie assembly)
### Changed
- `DocumentStorageControllerTests` (+~40 testów): pełne pokrycie endpointów (UpdateDocumentVersion, GetDocumentMetadata 403/404, user-download 200/403/404/400, finish, abort-send, continue-delivery, delivery status/list/retry/cancel/recipient-url).
- D2Services `DocumentControllerTests` (+~10): UpdateCallbackUrl (NoContent/404/Conflict/BadRequest), UnlockDocument (Ok/404/Conflict), CreateDocument (DOCX z flagami, unsupported mime, ingest fail), rekordy request.
- Nowe `DeliveryAttemptRunnerTests` (Sent/PermanentError/Retry/no-op/storage-throw), `ConverterRoundTripCoverageTests` (bogaty HTML→DOCX→HTML), `DeliveryOptionsTests`.
### Verified
- Coverage (coverlet + ReportGenerator, merge 5 projektów): line **64.3% → 72.1%**. Per-assembly: Domain 92.2%, Application 94.2%, Infrastructure 57.3%→**67.3%**, D2ViewerEditor.Api 49.4%→**61.5%**, D2ServicesViewerEditor.Api 53.9%→**60.2%** — wszystkie ≥60%.
- Testy zielone: Domain 77, Application 282, Infrastructure 178, D2Api.UnitTests 110, D2Services.UnitTests 39.
### Notes
- Pozostałe luki to świadomie infra/seam/startup (GCS client, EF DbContext/repozytoria, `Program.cs`, `DocumentDeliveryWorker`) — pokrywane testami integracyjnymi, nie jednostkowymi.

## 2026-06-22 — D2Services observability: middleware + correlation id + JSON payload parity
### Changed
- `D2ServicesViewerEditor.Api/Program.cs`:
  - dodano `UseRequestObservability()` i `UseExceptionHandlingMiddleware()` do pipeline,
  - rozszerzono `UseSerilogRequestLogging` o `EnrichDiagnosticContext` (`httpMethod`, `httpPath`, `statusCode`, `requestId`, `correlationId`).
- Dodano `D2ServicesViewerEditor.Api/Middleware/RequestObservabilityMiddleware.cs`:
  - nagłówek `X-Correlation-ID` (echo/generowanie),
  - `HttpContext.Items[CorrelationId]`,
  - scope (`correlationId`, `requestId`) i Serilog `LogContext` dla pełnej korelacji.
- `D2ServicesViewerEditor.Api/Middleware/ExceptionHandlingMiddleware.cs`:
  - `correlationId` w `ProblemDetails.Extensions`,
  - serializacja `ProblemDetails` po typie runtime (zachowanie `errors` dla `ValidationProblemDetails`).
- `D2ServicesViewerEditor.Api/Logging/GcpJsonSerilogFormatter.cs`:
  - ujednolicony payload do standardu D2Api: `severity`, `level`, `timestamp`, `message`, `category`, `service`, `environment`, `traceId`, `spanId`, `exceptionType`, plus spłaszczone właściwości eventu/scope.
### Verified
- `dotnet test D2ServicesViewerEditor/D2ServicesViewerEditor.Api.UnitTests/D2ServicesViewerEditor.Api.UnitTests.csproj` → **28/28** pass.
### Notes
- Zmiana usuwa wcześniejsze luki: brak globalnego middleware wyjątków i brak korelacji requestów w D2Services.

## 2026-06-21 — D2Services: testy HealthController + MiddlewareExtensions
### Changed
- Dodano `HealthControllerTests` w `D2ServicesViewerEditor.Api.UnitTests`:
  - `Get_ReturnsOk_WithExpectedPayloadShape`,
  - `GetDetailed_ReturnsOk_WithDependenciesSection`.
- Dodano `MiddlewareExtensionsTests`:
  - `UseExceptionHandlingMiddleware_ReturnsSameBuilderInstance`.
### Verified
- `dotnet test D2ServicesViewerEditor/D2ServicesViewerEditor.Api.UnitTests/D2ServicesViewerEditor.Api.UnitTests.csproj` → **22/22** pass.
### Notes
- Rozszerzenie domyka podstawowe klasy API po wcześniejszym dodaniu testów dla kontrolera dokumentów, middleware wyjątków i formattera logów.

## 2026-06-21 — Zwiększenie liczby testów jednostkowych (D2Api + D2Services)
### Changed
- **D2ServicesViewerEditor:** utworzono nowy projekt `D2ServicesViewerEditor.Api.UnitTests` (dodany do `D2ServicesViewerEditor.sln`) z testami:
  - `DocumentControllerTests` (walidacja wejścia, wymagany `ReturnUrl` dla DOCX, rozpoznanie MIME po rozszerzeniu, mapowanie odpowiedzi statusu),
  - `ExceptionHandlingMiddlewareTests` (mapowania wyjątków 400/401/500 + pass-through),
  - `GcpJsonSerilogFormatterTests` (mapowanie `LogEventLevel`→`severity`, `category` z `SourceContext`, serializacja wyjątków).
- **D2ApiViewerEditor:** rozszerzono `RequestObservabilityMiddlewareTests` o przypadki brzegowe `X-Correlation-ID` (whitespace, zbyt długi >128); rozszerzono `ExceptionHandlingMiddlewareTests` o propagację `correlationId` do `ProblemDetails.Extensions`.
### Verified
- `dotnet test D2ServicesViewerEditor/D2ServicesViewerEditor.Api.UnitTests/D2ServicesViewerEditor.Api.UnitTests.csproj` → **19/19** pass.
- `dotnet test D2ApiViewerEditor/D2ViewerEditor.Api.UnitTests/D2ViewerEditor.Api.UnitTests.csproj` → **76/76** pass.
### Notes
- W projekcie testowym D2Services wersja pakietu `Serilog` ustawiona na `4.0.0`, aby uniknąć NU1605 (downgrade względem `Serilog.Sinks.Console`).

## 2026-06-21 — „Zakończ" sync + statusy, „Pobierz oryginał", ukrycie „Ostatnia modyfikacja"
### Changed
- **DocumentStatus** +`Queued` („Zlecono do wysyłki"), +`SendAborted` („UzytkownikPrzerwałWysyłkę"); `Document.MarkQueued/MarkSendAborted`.
- **DocumentDelivery** +`BeginInlineAttempt()` (Pending/RetryScheduled→Sending, inline, bez lease) +`HoldAfterFailedInlineAttempt()` (RetryScheduled zaparkowane na DeadlineAt — worker nie przejmie).
- **FinishAndSendDocumentCommandHandler** — synchroniczna pierwsza próba przez `IDeliverySender`; wynik `{DeliveryId,Status,DocumentStatus,Delivered,Error?}`. Endpoint `finish` → 200 (było 202).
- Nowe komendy/endpointy: `POST /{masterId}/abort-send` (Przerwij → delivery `Cancelled` + doc `SendAborted`), `POST /{masterId}/continue-delivery` (Kontynuuj w tle → `Requeue` + doc `Queued`).
- **GUI** `document-editor`: nowa pozycja „Plik → Pobierz oryginał dokumentu" (`canDownloadOriginal = loadedFromDisk || userDownload`, zwraca v1); flow „Zakończ" = „Trwa wysyłanie dokumentu ..." → sukces zamyka kartę / błąd: modal „Wystąpiły problemy..." + Przerwij/Kontynuuj; „Ostatnia modyfikacja" w stopce gated `showSaveState()`.
- `infra/sql/010_extend_document_status.sql` (COMMENT only — kolumna bez CHECK).
### Verified
- `dotnet build` D2Api + D2Services: 0 błędów. Domain 77, Application 282 zielone. GUI Vitest 267 zielone.
- `ng build` AOT kompiluje; błąd tylko z pre-existing budżetu bundla (nietknięte pliki).
### Notes
- Decyzje potwierdzone z użytkownikiem: synchroniczna pierwsza próba (nowy endpoint) + angielski PascalCase enuma (etykiety PL tylko w UI).
- „Pobierz oryginał" z aplikacji zewn. używa `GET {masterId}/download` (view-gated, jak podgląd Krok 2) — oryginał był już pobieralny dla każdego z dostępem do podglądu; `userDownload` bramkuje tylko pobranie EDYTOWANEJ kopii.

## 2026-06-19 — `showSaveState` (ingest) — ukrywanie UI zapisu w edytorze
### Changed
- **D2Services** `DocumentController.CreateDocument`: nowe opcjonalne pole `ShowSaveState: bool?` w `CreateDocumentRequest`; zapisywane do `documents.metadata` jako `showSaveState` **tylko gdy jawne `false`** (inverse-default — brak/`true` → `null` → widoczne).
- **Application**: `ExternalDocumentMetadata` +`ShowSaveState` (pozycyjny param z domyślnym `null`, nie łamie istniejących wywołań) + computed `IsSaveStateVisible => ShowSaveState != false`. `DocumentMetadataDto` +`ShowSaveState`; `GetDocumentMetadataQueryHandler` mapuje `IsSaveStateVisible`.
- **GUI** `document-editor`: sygnał `showSaveState` (domyślnie `true`), ustawiany z `meta.showSaveState !== false`. Szablon ukrywa switch Autozapis, status autozapisu w stopce i przycisk „Zapisz" gdy `false`. `DocumentMetadataDto` (service) +`showSaveState`.
### Verified
- `dotnet build` D2Api + D2Services: 0 błędów. Testy Application: 31/31 (metadata). GUI Vitest: 263/263 (w tym 2 nowe testy `showSaveState`).
### Notes
- Decyzja: `showSaveState=false` ukrywa **tylko UI** — autozapis dalej działa w tle (brak utraty danych), a „Zakończ i wyślij" pozostaje. Reguła prezentacyjna, nie zabezpieczenie.

## 2026-06-19 — Kestrel: globalny limit body 150 MB (duże dokumenty) + AddServerHeader=false
### Changed
- `Program.cs` `ConfigureKestrel`: `MaxRequestBodySize` konfigurowalny (`Kestrel:MaxRequestBodySizeBytes`, domyślnie 150 MB), `MinRequestBodyDataRate`/`MinResponseDataRate=null`, `AddServerHeader=false`. Powód: endpointy `DocumentStorageController` (upload/save/finish, base64) nie miały limitu → domyślny ~30 MB Kestrel groził **413** dla dużych dokumentów. Per-action `[RequestSizeLimit]` (DocumentController) dalej obowiązuje gdzie ostrzejszy.
### Verified
- `dotnet build` 0 błędów.
### Changed (cd.)
- **Wymuszenie `https`** (wzorzec D2WebCore, bezwarunkowo): middleware `context.Request.Scheme = "https"` jako pierwszy w pipeline — poprawne URL-e / OIDC redirect URI za TLS-terminującym proxy (potrzebne dla OIDC z ADR-0019).
### Notes
- Z przeglądu wzorca D2WebCore `Program.cs` **pominięto**: NLog (mamy structured `ILogger` — ADR-0014/0018), `InitializeCulture` (konwertery już używają `InvariantCulture`), `CodePagesEncodingProvider` (kod używa wbudowanego `Encoding.Latin1`, nie code-page; wymagałby nowego pakietu), inline GCP secret load (mamy `EntraSecretLoader`), `UseWindowsService`/`Startup` (kontenery + minimal hosting).

## 2026-06-19 — Parytet z D2WebCore: ClaimsTransformer (szersze typy claimów) + proxy z credentialami
### Changed
- `ClaimsTransformer`: dopasowuje identyfikatory grup/ról niezależnie od typu claimu (`groups`/`roles`/`ClaimTypes.Role`/`*identity/claims/role*`) — jak wzorcowy `ClaimsTransformer`. **Output nadal `roles`** (nasz `RoleClaimType`), nie `ClaimTypes.Role` (inaczej `IsInRole` by nie działał). Tolerancja null `GroupNames`.
- `ProxyOptions` (= `BusinessProxy`): +`Username`/`Password`; `EntraBackchannel` używa `NetworkCredential` gdy username podany, inaczej `UseDefaultCredentials`. Dodane do appsettings (puste; Password z secret store).
- **Nazwy klas bez zmian** (`AzureAdOptions`/`ClaimsTransformer`/`RolesOptions`) — ugruntowane; pominięto martwe pola wzorca (`UserNameClaimType`/`RefreshThresholdMinutes`).
### Verified
- `dotnet build` 0 błędów; Api.UnitTests **73/73** (+`CreateProxy_uses_explicit_credentials`, +`Maps_group_identifier_carried_under_role_claim_type`).

## 2026-06-19 — Entra: jawny `IS_LOCAL_DEV` + serwerowy OIDC `AddMicrosoftIdentityWebApp` (hybryda WebApi+WebApp)
### Changed
- `AddEntraIdAuthentication` (ConfigureAuthentication): **jawny** `IS_LOCAL_DEV` (env var) steruje proxy — lokalnie `UseProxy=false`, inaczej `WebProxy` z `AzureAd:Proxy:Url` (bypass GCS) jako `BackchannelHttpHandler` (WebApi + WebApp) i `HttpClient.DefaultProxy`.
- **Dodano serwerowy OIDC** `AddMicrosoftIdentityWebApp` (scheme `MyAzureAdScheme`): code flow + cookie (`SignInScheme`, `NonceCookie/CorrelationCookie SecurePolicy=Always`, scope `offline_access`/`email`, `ResponseType=Code`) + `EnableTokenAcquisitionToCallDownstreamApi()` + `AddInMemoryTokenCaches()`. WebApi (JWT) pozostaje **domyślnym** schematem dla API SPA. Na wyraźną prośbę — **odwraca** „resource-server-only" z ADR-0011/0012. Patrz ADR-0019.
- Bez nowej zależności (`AddMicrosoftIdentityWebApp` tranzytywnie z `Microsoft.Identity.Web` 3.12.0). `Microsoft.AspNetCore.Authentication.OpenIdConnect` ściągany tranzytywnie.
### Verified
- `dotnet build` 0 błędów; Api.UnitTests **71/71**. **Runtime NIEzweryfikowane** (brak tenanta/ClientSecret lokalnie): faktyczny flow OIDC, cookie, token acquisition — patrz R-27.

## 2026-06-19 — Refaktor: wiring Entra do `ConfigureAuthentication` (parytet z D2WebCore)
### Changed
- Wydzielono auth z `Program.cs` do `Security/ConfigureAuthentication.cs`: `AddDevBypassAuthentication()` + `AddEntraIdAuthentication(configuration, environment)` (bind AzureAd, ShowPII gated, proxy backchannel, Identity.Web JWT, RoleClaimType PostConfigure, grupy→role, polityki, ClaimsCurrentUserProvider, Graph). `Program.cs` = `if (devAuthBypass) AddDevBypass… else AddEntraId…`. Zachowanie 1:1.
- Dodano (z wzorca, gated) `JwtBearerOptions.IncludeErrorDetails = !IsProduction()` — szczegóły 401 w dev/TST, off w PRD.
- ~~Świadomie NIE dodano `AddMicrosoftIdentityWebApp`~~ — **zmienione w nowszym wpisie (powyżej)**: dodane na wyraźną prośbę (ADR-0019).
### Verified
- `dotnet build` 0 błędów; Api.UnitTests **71/71**.

## 2026-06-17 — 3 zgłoszenia: payload wysyłki (master/version/corporateKey), widok „Brak uprawnień", observability ELK
### Changed
- **Z1 (wysyłka):** `HttpDeliverySender` (multipart) dokłada obok `file` pola `masterId`/`versionId`/`corporateKey` + log techniczny (bez treści, corporateKey jako flaga). `DeliveryDispatch` +`MasterId`/`VersionId`/`CorporateKey`; `DeliveryAttemptRunner` wypełnia z encji. `DocumentDelivery` +`CorporateKey` (kolumna `corporate_key`, SQL `009`, EF map), `FinishAndSendDocumentCommandHandler` czyta `ICurrentUserProvider.CorporateKey`. Backward compatible (pole `file` bez zmian).
- **Z2 (frontend):** nowy generyczny widok `AccessForbiddenComponent` „Brak uprawnień" (`/brak-uprawnien`); `resourceGuard`+`documentAccessGuard`+interceptor (403) kierują tu; `/access-denied` → redirect. Usunięto `document-access-denied`. 403 nie mylone z 401/404/500.
- **Z3 (ELK):** `GcpJsonConsoleFormatter` wzbogacony (scope'y→pola, `service`/`environment`/`traceId`/`spanId`/`level`); nowy `RequestObservabilityMiddleware` (correlationId z `X-Correlation-ID` + access-log method/path/status/elapsed/userId, scope na cały request); `correlationId` w ProblemDetails. Dokumentacja `.ai/OBSERVABILITY.md` (pola + KQL Kibana + zasady nie-logowania danych wrażliwych).
### Verified
- Backend: `dotnet build` 0 błędów; unit testy **577** (Api 71, Application 267, Infrastructure 166, Domain 73) — w tym naprawiony wcześniej istniejący `GetDocumentVersionContentQueryHandlerTests`. GUI: `tsc` czysto; `ng test` **261**.
### Notes
- corporateKey = claim `ck` użytkownika kończącego (async wysyłka → utrwalony na delivery); brak → puste pole (bez błędu).
- Integration testy delivery na realnym Postgres niezweryfikowane lokalnie (jak dotąd, R-06).

## 2026-06-17 — Konsumpcja proxy Entra: `AzureAd:Proxy:Url` → JwtBearer backchannel (D2WebCore pattern)
### Changed
- Nowy `EntraBackchannel.CreateProxy(AzureAdOptions)` — buduje `WebProxy` z `AzureAd:Proxy:Url` (bypass `storage.googleapis.com`, `UseDefaultCredentials=true`) lub `null` gdy brak URL. `Program.cs` (gałąź Entra): gdy proxy ustawione → `HttpClient.DefaultProxy = proxy` oraz `JwtBearerOptions.BackchannelHttpHandler` (pobieranie OpenID metadata/JWKS zza korpo-proxy). Brak URL / lokalnie → bez zmian (direct).
### Verified
- `dotnet build` 0 błędów; `dotnet test` Api.UnitTests **67/67** (dodano `EntraBackchannelTests` 3).
### Notes
- `IdentityModelEventSource.ShowPII` z wzorca dodany, ale **gated `!IsProduction()`** (nigdy w PRD — wyciek PII to HIGH w analizie; w PRD off). Gating na `!IsProduction()` a nie `IsDevelopment()`, bo env nazywa się „DEV"/itp., nie „Development".
- **Graph/Azure.Identity** (`ClientSecretCredential` w `GraphUserService`) NIE jest objęty `HttpClient.DefaultProxy` — Azure.Identity używa własnego pipeline'u (proxy via env `HTTPS_PROXY` lub `TokenCredentialOptions.Transport`). Follow-up, jeśli Graph ma działać zza proxy.

## 2026-06-17 — Finalne nazewnictwo ról: `Administrator` / `Operator` (rename z `APP_Admin`/`APP_Pracownik`)
### Changed
- Backend: globalny rename wartości `APP_Admin`→`Administrator`, `APP_Pracownik`→`Operator` w `AzureAdOptions` (default `AdminRole`), `appsettings.json`/`appsettings.DEV.json` (sekcja `Roles`→`RoleName`), `DevAuthHandler` (claimy), testach (`ClaimsTransformerTests`, `IdentityControllerTests`) + komentarzach.
- Rename **identyfikatorów**: właściwość `AzureAdOptions.EmployeeRole`→`OperatorRole` (klucz `AzureAd:OperatorRole`), stała polityki `RequireAppEmployee`→`RequireAppOperator` (nazwa+wartość), `isEmployee`→`isOperator`, nazwy testów. `RequireAppAdmin`/`AdminRole` bez zmian. Semantyka `ResourcesProvider`/polityk bez zmian.
- **GUI:** brak zmian kodu — autoryzacja resource-based (ADR-0016), front nie zna nazw ról (potwierdzone grepem).
- Docs „żywe" (SECURITY/DOMAIN/FEATURES/API_CONTRACTS/CURRENT_STATE) zaktualizowane na nowe nazwy; historyczne ADR-y zachowane.
### Verified
- API: `dotnet build` 0 błędów; `dotnet test` Api.UnitTests **64/64**.
### Notes
- Patrz ADR-0017. R-26: realne Entra appRoles/`RolesOptions` muszą emitować `Administrator`/`Operator` (albo nadpisać `AzureAd:OperatorRole`/`AdminRole`).
- Niezwiązane: `D2ViewerEditor.Application.UnitTests` ma **wcześniej istniejący** błąd kompilacji (`GetDocumentVersionContentQueryHandlerTests` — brak arg. `accessGuard`) — nie z tej zmiany.

## 2026-06-17 — Autoryzacja resource-based (backend /identity/resources + ResourcesProvider, front resourceGuard)
### Changed
- **Backend (D2ApiViewerEditor):** nowy `ResourcesProvider` (rola→zasoby: employee/admin→editor,viewer; admin→+admin) + endpoint `GET /api/identity/resources` (`RequireAppEmployee`). Polityka admina w `IdentityController` przeniesiona z klasy na akcję `users`. DI `AddScoped<ResourcesProvider>()`. `AzureAdOptions` + `Scopes[]`/`Proxy{Url}` (+ appsettings.json/DEV) — powierzchnia konfiguracji (konsumpcja = follow-up).
- **Front (D2GuiViewerEditor):** `ResourceAccessService` (cache + dev-bypass) + `resourceGuard` (gating po `route.path` względem listy z backendu, fail-closed→/access-denied). `resourceGuard` zastąpił `documentRoleGuard`/`appAdminGuard` na /editor,/viewer,/admin. **Usunięto:** `documentRoleGuard`, `appAdminGuard`, `CurrentUserService` (+spec), pola `adminRole`/`viewerRole` z `AppAuthConfig`/`config.json` — front nie zna już nazw ról.
### Verified
- API: `dotnet build` 0 błędów; `dotnet test` Api.UnitTests **64/64** (IdentityController 7). GUI: `tsc` czysto; `ng test` **262/262** (dodano `resource-access.service.spec` 3; usunięto `current-user.service.spec` 5).
### Notes
- Zamyka R-25 (rozjazd nazw ról front↔backend — front już nie zna ról). Otwiera **R-26**: backend `IsInRole` wymaga, by Entra emitowało `APP_Pracownik`/`APP_Admin` (albo nadpisać `AzureAd:EmployeeRole`/`AdminRole`). Patrz ADR-0016.
- `Scopes`/`Proxy` to na razie tylko config — konsumpcja (proxy na Graph/token, downstream scopes) niewdrożona.

## 2026-06-16 — Front Entra: MSAL standalone (redirect-component), config.json jako źródło auth, model ról Administrator/Przeglądający
### Changed
- `main.ts`/`index.html`: `MsalRedirectComponent` (`<app-redirect>`) bootstrapowany przez `appRef.bootstrap(...)` (standalone, bez AppModule); `App` nie woła już `handleRedirectObservable()` (aktywne konto po `inProgress$===None`).
- Wartości auth usunięte z `environment.*` → jedyne źródło to `src/assets/configs/config.json` (przeniesiony z `public/`). `DEFAULT_AUTH_CONFIG` = neutralne defaulty strukturalne. Token DI `RUNTIME_AUTH_CONFIG` → `MSAL_CUSTOM_CONFIG`.
- `msal.config.ts`: helper `apiScopesFor()` (jawne `apiScopes` lub fallback `{clientId}/.default`), jawne `navigateToLoginRequestUrl: true`.
- RBAC (UX): `AppAuthConfig.viewerRole` (domyślnie „Przeglądający"), `adminRole` domyślnie „Administrator". `CurrentUserService` scentralizowany (`isAdmin`/`isViewer`/`canAccessDocuments`, admin=nadzbiór, dev-bypass). Nowy `documentRoleGuard` na `/editor` i `/viewer`; `appAdminGuard` przez `CurrentUserService`.
- Lokalny `config.json` = dev-bypass (`auth.enabled=false`) — aplikacja startuje bez realnej rejestracji Entra.
### Verified
- `tsc --noEmit -p tsconfig.app.json` — czysto. `ng test --watch=false` — **264/264** (dodano `current-user.service.spec` — 5; zaktualizowano `app.spec`, `runtime-config.spec`).
### Notes
- **Rozjazd nazw ról z backendem** (ADR-0011/0012: `APP_Pracownik`/`APP_Admin`): dosynchronizować Entra appRoles/`RolesOptions` lub nadpisać `adminRole`/`viewerRole` w config.json. Patrz ADR-0015.
- Niezweryfikowane runtime (brak tenanta lokalnie): realny redirect + claim `roles`.

## 2026-06-13 — Fix: ProblemDetails serializowany przez typ runtime (errors w body)
### Changed
- `ExceptionHandlingMiddleware` serializował `problemDetails` przez **statyczny typ bazowy** `ProblemDetails` → `ValidationProblemDetails.Errors` ginęło w body (System.Text.Json honoruje typ statyczny). Fix: `JsonSerializer.Serialize(problemDetails, problemDetails.GetType(), options)` — serializacja po typie runtime, więc słownik `errors` (zgrupowany po `PropertyName`) trafia do odpowiedzi 400.
- Test `InvokeAsync_ValidationException_EmitsGroupedErrorsInBody` asercjonuje teraz realną obecność `errors.Name`/`errors.Email` w body (poprzedni test „realnego zachowania bez errors" zastąpiony zachowaniem docelowym).
### Verified
- `dotnet test D2ViewerEditor.Api.UnitTests` — **51** pass (+1). Reszta solucji bez zmian.
### Notes
- Kontrakt API wzbogacony (dodane pole `errors` w 400 dla błędów walidacji) — zgodne z RFC 7807 `ValidationProblemDetails`; klienci dotychczas i tak nie dostawali `errors`, więc brak regresji.

## 2026-06-13 — Pokrycie testami: pipeline behaviours, metadata, exception middleware
### Changed
- **Application.UnitTests** (+19): `ExternalDocumentMetadataTests` (parse tolerancyjny null/empty/malformed/json-null, mapowanie pól, case-insensitive, `IsUserDownloadAllowed` tylko dla jawnego `true`, round-trip serialize→parse, camelCase keys); `ValidationBehaviourTests` (brak walidatorów→next, wszystkie pass→next, fail→`ValidationException`+short-circuit, agregacja błędów z wielu walidatorów); `LoggingBehaviourTests` (przekazanie odpowiedzi, log na wejściu+wyjściu, brak połykania wyjątku).
- **Api.UnitTests** (+7): `ExceptionHandlingMiddlewareTests` — pass-through bez wyjątku, mapowanie `ValidationException`→400 problem+json, `ArgumentException`→400 z detail, `KeyNotFoundException`→404, `OperationCanceledException`→400, nieoczekiwany→500 **bez przeciekania** szczegółów, camelCase ProblemDetails.
### Verified
- `dotnet test D2ViewerEditor.sln` — Domain 67, Api **50** (+7), Application **258** (+19), Infrastructure 164, Integration 6 skip. 0 failures.
### Notes
- Odkryta latentna obserwacja: middleware serializuje `ValidationProblemDetails` przez statyczny typ `ProblemDetails`, więc słownik `errors` NIE trafia do body (test asercjonuje realne zachowanie: 400 + tytuł „Błąd walidacji", bez `errors`). Nie zmieniano produkcyjnego zachowania w ramach dodawania testów.
- Nadal nietestowane (świadomie, wymaga infra/seam): `DeliveryAttemptRunner`, `DocumentDeliveryWorker`, `OoxmlAgileDecryptor`, repo/GCS.

## 2026-06-11 — „Otwórz w edytorze": z osobnej zakładki na przycisk w wierszu wysyłki
### Changed
- **Usunięto** stronę `admin-open-document` (+route `/admin/open`, +pozycję sidebaru w `admin-shell`, +spec).
- **Dodano** w `admin-deliveries` przycisk „Otwórz" per wiersz → `openInEditor(item)`: `router.navigate(['/editor'], { queryParams })` z `masterId=documentId`, `versionId=sourceVersionId` (bez versionId → podgląd). `Router` wstrzyknięty do komponentu.
### Verified
- `admin-deliveries.spec` **13** (+2 nawigacja z/bez sourceVersionId); `ng build` OK.

## 2026-06-11 — Wysyłka na returnUrl jako multipart/form-data
### Changed
- `HttpDeliverySender.SendAsync`: zamiast `ByteArrayContent` (`application/octet-stream`) wysyła `MultipartFormDataContent` — plik w polu **`file`**, filename `document.docx`, part Content-Type = DOCX MIME. Nagłówki `Idempotency-Key`/`X-Content-SHA256` i klasyfikacja statusów bez zmian.
### Verified
- `Infrastructure.UnitTests` +3 `HttpDeliverySenderTests` (multipart + name=file/filename + nagłówki; 422→Permanent; 503→Retryable). Build OK.
### Notes
- **Zmiana kontraktu wysyłki**: odbiorca na `returnUrl` musi odczytać plik z pola form-data `file` (np. `IFormFile file`), nie z surowego body. Nazwa pola jest stała (`file`) — gdyby odbiorca wymagał innej, do sparametryzowania w `HttpDeliverySender`.

## 2026-06-11 — Observability: strukturalne logi JSON dla GCP (severity)
### Changed
- **Internal API** (`D2ViewerEditor.Api`): nowy `Logging/GcpJsonConsoleFormatter` (ConsoleFormatter → JSON z `severity`, `message`+stack trace, `category`, `eventId`, `exceptionType`); `Extensions/LoggingExtensions.AddGcpStructuredLogging()` wpięte w `Program.cs`. Aktywne poza Development lub gdy `Logging:UseGcpFormat=true`.
- **External API** (`D2ServicesViewerEditor.Api`): nowy `Logging/GcpJsonSerilogFormatter : ITextFormatter`; `Program.cs` `WriteTo.Console(new GcpJsonSerilogFormatter())` poza Development (inaczej zwykły tekst). File-sink bez zmian.
### Verified
- `dotnet build` obu hostów OK; `Api.UnitTests` **43** (+8 `GcpJsonConsoleFormatterTests`: mapowanie poziomów, wyjątek=ERROR+stack w message, jedna linia JSON).
- Manualnie do potwierdzenia w GCP: po deployu wyjątki w Logs Explorer mają severity ERROR/CRITICAL i są filtrowalne; Error Reporting grupuje po stack trace.
### Notes
- Root cause: stdout plain-text → Cloud Logging nadaje INFO; rozwiązanie = pole `severity` w JSON. ADR-0014.
- Follow-up: `LoggingBehaviour` loguje `{@Request}` (pełny payload — hałas/dane wrażliwe), do okrojenia osobno.

## 2026-06-11 — Fix: „Zamknij" w modalu wysyłki zostawiał edytowalny dokument
### Changed
- `document-editor`: nowy sygnał `workFinished`; `closeFinishModalAndExit()` zatrzymuje auto-save i ustawia stan końcowy zamiast tylko zamykać modal. Nowy blokujący `work-finished-overlay` (z-index 3000) z przyciskiem „Zamknij kartę" (`closeFinishedTab()`). `window.close()` pozostaje best-effort, ale gdy zawiedzie, użytkownik widzi ekran końcowy i nie wraca do edycji.
### Verified
- `ng build` OK; `document-editor.spec` **62** (+1: po „Zamknij" `workFinished=true`, auto-save off, modal zamknięty).
### Notes
- Root cause: `window.close()` działa tylko dla kart otwartych przez `window.open` — dla zwykłych/otwartych przez link przeglądarka odmawia. Dlatego potrzebny jawny stan końcowy w UI.

## 2026-06-11 — „Pliki do wysłania": edycja adresu odbiorcy (returnUrl)
### Changed
- **Domena:** `DocumentDelivery.UpdateRecipientUrl(url)` — walidacja absolutnego http(s) + blokada dla `Sent`/`Sending`.
- **Application/API:** `UpdateDeliveryRecipientUrlCommand`+handler (przycina URL, mapuje Invalid/Argument → Failure, brak → NotFound); endpoint `PUT /api/documentstorage/deliveries/{id}/recipient-url` (body `{ recipientUrl }`) + `UpdateDeliveryRecipientUrlRequest`.
- **GUI:** `admin-deliveries` — przycisk „Edytuj" (gdy `canEdit`: ≠ Sent/Sending) → modal (`edit-overlay`/`edit-dialog`) z polem URL, walidacją i Zapisz/Anuluj; sygnały `editingId`/`editUrlValue`/`editError`/`savingEdit`; autoodświeżanie pauzowane przy otwartym modalu. `updateDeliveryRecipientUrl()` + `UpdateDeliveryRecipientUrlResult` w `document-storage.service`. `dl-btn-primary` (SCSS).
### Verified
- Backend: `dotnet build` sln OK; Domain `DocumentDelivery` 21 (+5 UpdateRecipientUrl), Application `UpdateDeliveryRecipientUrlCommandHandler` 3.
- GUI: `ng build` OK; `admin-deliveries.spec` 11 (+4 edycja).
### Notes
- Bez migracji SQL (zmiana wartości kolumny `recipient_url`, nie schematu). Po zmianie adresu zadanie zwykle wymaga „Wznów".

## 2026-06-11 — „Pliki do wysłania": autoodświeżanie 3 s + akcje Anuluj/Wznów
### Changed
- **Domena:** `DeliveryStatus.Cancelled` (nowy stan końcowy); `DocumentDelivery.IsTerminal` obejmuje Cancelled; `Cancel()` (Pending/RetryScheduled→Cancelled, czyści lease, ustawia LastError „Anulowano ręcznie"); `Requeue()` rozszerzony o RetryScheduled (wyślij teraz) i Cancelled (poprzednio tylko DeadLettered/FailedPermanently). Worker bez zmian — claim selektuje Pending/RetryScheduled/stuck-Sending, więc Cancelled jest wykluczony.
- **Application/API:** `CancelDeliveryCommand`+handler (anulowanie + `Document.MarkSaved`); endpoint `POST /api/documentstorage/deliveries/{id}/cancel`. „Wznów" reużywa istniejącego `/retry` (Requeue).
- **SQL:** `infra/sql/008_add_delivery_cancelled_status.sql` — DROP+ADD `ck_document_deliveries_status` z wartością `Cancelled` (indeksy częściowe due/active bez zmian).
- **GUI:** `admin-deliveries` — autoodświeżanie co 3 s (`interval`, ciche `fetch(silent)`, przełącznik, `OnDestroy`); przyciski „Wznów" (`canResume`) i „Anuluj" (`canCancel`, `dl-btn-danger`); `cancelDelivery()` w `document-storage.service`; `DeliveryStatus` +`Cancelled`; statusLabel „Anulowano"/statusClass `status-cancelled`; status w dropdownie filtra.
### Verified
- Backend: `dotnet build` sln OK; Domain `DocumentDelivery` 16 (+6 Cancel/Requeue), Application `CancelDeliveryCommandHandler` 3.
- GUI: `ng build` OK; `ng test` **244** (+7 `admin-deliveries.spec`).
### Notes
- **Wymaga uruchomienia migracji `008` na każdym środowisku** (CHECK constraint). Bez niej zapis statusu `Cancelled` odrzuci baza.
- Anulowanie zablokowane dla `Sending` (lease workera) — uniknięcie wyścigu z trwającą próbą.

## 2026-06-11 — 5 zgłoszeń: dialog hasła, .doc, GUI „wysyłanie", EMF→PNG, VersionId w mailu
### Changed
- **(P1) Dialog hasła** — `document-editor.html`: usunięto `(click)=cancelPasswordDialog()` z overlay i `(keydown.escape)` z pola (dialog zamykają tylko `Anuluj`/`Otwórz`; dodano `role=dialog`/`aria-modal`). `document-editor.ts` `cancelPasswordDialog()`: bez `returnUrl` → `router.navigate(['/'])`, z `returnUrl` → `tryCloseBrowserTab()` — koniec pustego dokumentu po anulowaniu.
- **(P3) Modal „Trwa wysyłanie pliku…"** — zamiast ikony papierowego samolotu używa `.loading-spinner` (klasa `.finish-dialog-spinner`, +SCSS `margin-bottom:16px`) — spójny z ekranem „Przetwarzanie…". Logika/countdown bez zmian.
- **(P5) Mail „Zgłoś"** — `openReportEmail()` podstawia `Version ID` = `documentVersionId()` (było `documentMetadata().version` = zwykle `—`); fallback `—` tylko gdy wersji brak; `Master ID` bez zmian.
- **(P4) EMF/WMF → PNG** — `GraphicConversionService`: `TryRasterizeMetafileToPng` + `TryExtractEmfDib` (rekordy STRETCHDIBITS/SETDIBITSTODEVICE/BITBLT/STRETCHBLT/ALPHABLEND) + `TryFindDibGeneric` + `WrapDibInBmpFile` + `DecodeToPng` (SkiaSharp). `DocxToHtmlConverter`: `data-original-src` niesiony dla każdego metafile (drawing + VML), nie tylko placeholdera → eksport wektorowy zachowany.
- **(P2) Binarny .doc → .docx** — nowy `LegacyDocBinaryConverter` (FIB + piece table → tekst+akapity → DOCX/OpenXML); wpięty w `DocumentInputNormalizer` z fallbackiem `UnsupportedLegacyDoc`.
### Verified
- Backend: `dotnet build` sln OK (0 błędów); `Infrastructure.UnitTests` 155 (+1 EMF→PNG `Emf_WithEmbeddedDib_IsRasterizedToPng`, +4 `LegacyDocBinaryConverterTests`); `Api.UnitTests` DocumentController 13.
- GUI: `ng build` OK; `ng test` **233** (+4 w `document-editor.spec`: VersionId obecny/fallback, cancel bez/z returnUrl).
### Notes
- Ograniczenia świadome: czysto wektorowe EMF (bez rastra) wciąż placeholder; `.doc` odzyskuje tekst+akapity, nie formatowanie/tabele/obrazy. Patrz ADR-0013. Pełna wierność = sidecar LibreOffice (roadmapa).
## 2026-06-10 — Pełny wzorzec Qutas/D2WebCore dla Entra ID (Identity.Web + grupy→role + Graph + Secret Manager + Keycloak + runtime config)
### Changed (ADR-0012)
- **Backend — biblioteka:** `JwtBearer` (goły) → **`Microsoft.Identity.Web` 3.12.0** (`AddMicrosoftIdentityWebApi`). Pin `JwtBearer 8.0.12` **usunięty** (Identity.Web dostarcza per-TFM; pin dawał NU1605 na net9.0). Dodano `AzureAd:ClientId`/`ClientSecret`.
- **Backend — grupy→role (Qutas):** `RolesOptions` (sekcja `Roles`: `GroupPrefix` + `Roles[]{RoleName,GroupNames}`) + `ClaimsTransformer : IClaimsTransformation` mapuje claim `groups` → role `APP_Pracownik`/`APP_Admin`. **App Roles zachowane** (współistnieją). Polityki bez zmian. Test `ClaimsTransformerTests` **6/6**.
- **Backend — Microsoft Graph v5 (5.103.0):** `IGraphUserService`/`GraphUserService` (app-only `ClientSecretCredential`) + `GET /api/identity/users?query=` (`IdentityController`, RequireAppAdmin). Aktywny tylko z ClientSecret; inaczej `DisabledGraphUserService`. (NIE `Identity.Web.MicrosoftGraph` = Graph v4.)
- **Backend — GCP Secret Manager** (`Google.Cloud.SecretManager.V1` 2.6.0): `EntraSecretLoader` wstrzykuje `AzureAd:ClientSecret` z `SMC01{ENV}2_APP00404_entra_secret`; guard `Enabled` + try/catch (nigdy nie wywala startu).
- **Backend — Keycloak dual-auth:** gdy `Keycloak:Enabled` — drugi schemat JwtBearer + policy scheme `EntraOrKeycloak` wybierający po issuerze tokena (`AuthSchemes.SelectByIssuer`). Wyłączony → tylko Entra.
- **Frontend — runtime config (Qutas):** `public/assets/configs/config.json` ładowany w `main.ts` przed bootstrapem → `RUNTIME_AUTH_CONFIG` (root-factory fallback = `environment.auth`). Fabryki MSAL (`msal.config`), `appAdminGuard`, `CurrentUserService` czytają z runtime-configu. Jeden build na wszystkie środowiska.
- **Konfiguracja:** `appsettings.json`/`appsettings.DEV.json` rozszerzone o `AzureAd.ClientId/ClientSecret`, `Roles`, `Keycloak`, `GCPSecretManager` — wszystko **placeholdery** (zero realnych sekretów/tenanta).
### Verified
- `dotnet build` API OK (Identity.Web/Graph 5.103.0/Azure.Identity/SecretManager restore). `Api.UnitTests` **41** (+6 `ClaimsTransformerTests`). GUI `tsc` OK, `ng test` **236**, `ng build` (AOT) OK; `config.json` shippowany do `dist/.../assets/configs/`.
- **Niezweryfikowane runtime** (brak tenanta/GCP/Keycloak lokalnie): walidacja tokenów Entra, mapowanie grup z realnego tokena, pobranie sekretu z GCP, selekcja schematu Keycloak. Aktywacja wymaga realnych wartości per środowisko + app consent dla Graph (`User.Read.All`).
### Notes
- Adaptacja Qutas do kształtu SPA+API: **bez** serwerowego OIDC/cookie (`AddMicrosoftIdentityWebApp`) — interaktywny login robi MSAL w przeglądarce, backend = resource server.
- Zmiana wykonana na bieżącym drzewie (niezakończony merge `feature/azure`); `.claude/settings.json` nadal do rozwiązania przez użytkownika.

## 2026-06-09 — „Lista plików" (admin-files): statusy dokumentów po polsku
### Changed
- `admin-files.ts` `statusLabel` — etykiety EN→PL: `Saved`→„Zapisany", `Editing`→„W edycji", `Sending`→„Wysyłanie do odbiorcy", `DeliveryFailed`→„Odbiorca nie odpowiada", `Sent`→„Wysłany". Wartości enuma `DocumentStatus` z API (klucze) bez zmian; `statusClass` (kolory) bez zmian. Filtr statusu jest free-text i matchuje po `statusLabel`, więc działa na polskich etykietach.
### Verified
- `tsc --noEmit` OK. Brak testów odwołujących się do starych etykiet EN.

## 2026-06-09 — Panel administracji (GUI): domyślnie firmowa czcionka
### Changed
- `admin-shell.scss` `:host` — `font-family: var(--corporate-font-family, 'Calibri', 'Segoe UI', Arial, sans-serif)`. Cała zawartość panelu admina (sidebar + strony routowane `admin-files`/`admin-deliveries`) dziedziczy firmowy krój. `--corporate-font-family` to istniejący punkt konfiguracji z `assets/fonts/_corporate-font.scss` (gdy firmowy font nieskonfigurowany → fallback Calibri/Segoe UI; gdy skonfigurowany — automatycznie się podstawi, jak w edytorze).
- `styles.scss` — globalna reguła `d2-admin-shell input, select, button, textarea { font-family: inherit }`. Natywne kontrolki formularzy mają własny font systemowy i nie dziedziczą; reguła musi być globalna (nie component-scoped), bo dotyczy kontrolek w stronach routowanych przez `<router-outlet>`.
### Verified
- `ng build` (AOT) OK; `admin-shell.spec` 3/3. Bez zmian logiki/TS.

## 2026-06-09 — Panel „Pliki do wysłania": rozwijane szczegóły wysyłki po kliknięciu wiersza
### Changed
- **Backend:** `DeliveryListItemDto` rozszerzony o `SourceVersionId` (Guid) i `RecipientUrl` (string) — mapowane w `GetDeliveriesByStatusQueryHandler` z `DocumentDelivery.SourceVersionId`/`RecipientUrl` (oba już istniały na encji; `recipientUrl` = returnUrl, nie jest sekretem — i tak wystawiany przez `…/metadata`).
- **Frontend model:** `DeliveryListItem` +`sourceVersionId`, +`recipientUrl`.
- **`admin-deliveries`:** sygnał `expandedId` + `toggleExpand(id)`/`isExpanded(id)` (jeden wiersz rozwinięty naraz). Klik w `.doc-row` rozwija/zwija; przycisk „Ponów" ma `event.stopPropagation()`, więc nie koliduje. Wiersz `.detail-row` (colspan=11) renderuje grid: **Adres odbiorcy**, **Master ID** (=documentId), **Version ID** (=sourceVersionId), **Status**, **Utworzono**, **Ostatnia próba**, **Komunikat błędu**. Caret ▸ obraca się 90° po rozwinięciu.
### Verified
- Backend `Application.UnitTests` **205** (`GetDeliveriesByStatusQueryHandlerTests` = 8; +1 test mapowania `RecipientUrl`/`SourceVersionId`/`DocumentId`). `dotnet build` OK.
- Frontend `tsc` OK; `ng build` (AOT) OK — `admin-deliveries` kompiluje template z rozwijanym wierszem.

## 2026-06-09 — Panel „Pliki do wysłania": domyślnie wszystkie statusy (zamiast tylko DeadLettered)
### Changed
- **Przyczyna pustego panelu:** domyślny filtr statusu = `DeadLettered`, a świeżo zakolejkowane wysyłki mają `Pending`/`Sending`/`RetryScheduled`/`Sent`. Endpoint `GET …/deliveries` wymagał konkretnego statusu, więc panel startował pusty.
- **Backend:** `GET …/documentstorage/deliveries` — parametr `status` opcjonalny: pusty / `null` / `"all"` (case-insensitive) ⇒ wszystkie statusy. `IDocumentDeliveryRepository.GetAllAsync(skip, take)` (+impl w `DocumentDeliveryRepository`, `OrderByDescending(CreatedAt)` jak `GetByStatusAsync`). `GetDeliveriesByStatusQuery.Status` → `string?` (default `null`); handler rozgałęzia all / konkretny status / nieznany (Failure). Kontroler: domyślny `status` `"DeadLettered"` → `""`.
- **Frontend:** `DocumentStorageService.getDeliveriesByStatus` → **`getDeliveries(status: DeliveryStatus | null = null, …)`** — pomija param `status` w żądaniu, gdy `null`. `admin-deliveries`: `selectedStatus` typu `DeliveryStatus | 'all'`, domyślnie **`'all'`**; w dropdownie nowa opcja **„Wszystkie"** (przekazuje `null` do serwisu).
### Verified
- Backend `D2ViewerEditor.Application.UnitTests` **204** (+7 nowy `GetDeliveriesByStatusQueryHandlerTests`: empty/whitespace/null/all/ALL → `GetAllAsync`; konkretny → `GetByStatusAsync`; nieznany → Failure bez zapytań). `dotnet build` solucji OK (0 błędów).
- Frontend `tsc --noEmit` OK; `ng build` (AOT) OK — `admin-deliveries` kompiluje się z nową opcją.
### Notes
- To zmiana tylko domyślnego widoku + dodanie opcji „wszystkie". Jeśli `document_deliveries` jest pusta (nikt nie dokończył „Zakończ"), panel nadal pokaże „Brak wyników" — to poprawne.

## 2026-06-09 — UX „Zakończ": modal „Trwa wysyłanie pliku" + odliczanie 30 s + best-effort zamknięcie karty
### Changed
- **Frontend `document-editor.ts`** (`finishDocument` przebudowany): po zapisie (`saveDocument`) i zleceniu wysyłki (`finishAndSend`) od razu otwiera modal „Trwa wysyłanie pliku do Twojej aplikacji" z przyciskiem **„Zamknij (NN)"** odliczającym 30→0. Nowe sygnały `showFinishModal`, `finishCountdown`; metody `openFinishModal`, `onFinishCountdownTick` (wydzielony tyk — deterministycznie testowalny bez fake-timerów), `closeFinishModalAndExit` (wspólna akcja dla końca odliczania i kliknięcia „Zamknij"), `tryCloseBrowserTab` (`window.close()` w try/catch — best-effort, bez crashu gdy przeglądarka odmówi). Odliczanie przez `timer(1000,1000)` w `finishCountdownSub`.
- **Usunięto blokujący `pollDeliveryStatus`** (timer + `getDeliveryStatus` + toasty „Dokument został wysłany"/„nie powiodła się"): nie czekamy już na stan końcowy dostarczenia — robi to backendowy `DocumentDeliveryWorker` (retry/backoff do 24h). Usunięty też nieużywany import `takeWhile` oraz pola `deliveryPollSub`/`DELIVERY_POLL_MS`.
- **Błąd natychmiastowego zakolejkowania** (`finishAndSend` rzuca): toast **„Nie udało się natychmiast wysłać pliku. Ponowimy próbę wysłania w tle."** — modal i odliczanie zostają (brak blokady użytkownika).
- **Ochrona przed wielokrotnym „Zakończ"**: guard `if (isFinishing()) return;` + `[disabled]="isFinishing()"` na przycisku; backendowy `finish` jest dodatkowo **idempotentny** (re-klik = to samo zadanie).
- **`template`**: nowy modal `@if (showFinishModal())` (overlay BEZ zamykania kliknięciem w tło — tylko przycisk „Zamknij", by nie dało się przypadkowo zamknąć/ponowić). **`scss`**: `.finish-dialog-overlay` (z-index/blur jak leave-dialog) + `.finish-close-btn` (stała `min-width`, by licznik nie skakał).
- **Higiena**: subskrypcja `finishSendSub` anulowana w `ngOnDestroy` (brak zapisów do sygnałów po zniszczeniu komponentu → koniec „document is not defined" przy teardownie testów).
### Verified
- GUI pełny zestaw **229/229** (`document-editor.spec.ts` **57**, +8 nowych: modal się pokazuje, countdown 30→29→28, dojście do 0 zamyka+`window.close`, klik „Zamknij" robi to samo, błąd→toast retry-w-tle + modal otwarty, wielokrotny klik = 1 flow, wyjątek `window.close` nie crashuje, readOnly bez modala). `tsc --noEmit` OK.
- **Backend bez zmian** — wykorzystany istniejący endpoint `finish` + worker/retry.
### Notes
- Eksperymentalny runner (Vitest unit-test builder) **nie wspiera `fakeAsync`** → odliczanie testowane przez bezpośrednie wołanie `onFinishCountdownTick(n)`; łańcuch async domykany `vi.waitFor`.
- `window.close()` zadziała niezawodnie tylko dla kart otwartych skryptem — przy zwykłym wejściu użytkownika karta może nie zostać zamknięta; to świadomie zaakceptowany, nie-krytyczny stan (modal zamknięty, edytor zostaje).

## 2026-06-09 — Podbicie podatnej zależności System.Security.Cryptography.Xml 8.0.2 → 10.0.6 (PRISMA)
### Changed
- `D2ViewerEditor.Infrastructure.csproj`: dodany **jawny** `PackageReference` na `System.Security.Cryptography.Xml` **10.0.6**. Pakiet nie był nigdzie referencjonowany bezpośrednio — wchodził **tranzytywnie** (przez NPOI/OpenXml) w wersji 8.0.2, flagowanej przez PRISMA. Jawny pin podbija rozwiązaną wersję (wyższa bezpośrednia wygrywa nad tranzytywną).
- Pin wpisany **tylko w Infrastructure**: oba hosty (`D2ViewerEditor.Api`, `D2ServicesViewerEditor.Api`) konsumują Infrastructure przez `ProjectReference`, więc 10.0.6 propaguje się do wszystkich konsumentów.
### Verified
- `dotnet list package --include-transitive` we **wszystkich trzech** projektach: `System.Security.Cryptography.Xml` rozwiązany na **10.0.6** (zamiast 8.0.2).
- `dotnet build -c Release` (Infrastructure + oba API): „Kompilacja powiodła się" — **0 błędów, brak ostrzeżeń NU1605** (downgrade). net8.0 konsumuje pakiet 10.0.x bez problemu (projekt już wcześniej mieszał `Microsoft.Extensions.Configuration.Abstractions 10.0.4` w net8.0).
### Notes
- Zmiana wyłącznie w build/restore — zero zmian kodu. Warto przepuścić pełny `dotnet test` (NUnit) dla potwierdzenia braku regresji w ścieżce podpisów (XML signing).

## 2026-06-09 — ENTER w polu rozmiaru czcionki kasował zaznaczony tekst
### Changed
- **Root cause:** `onFontSizeInputEnter` wołał `input.blur()` synchronicznie w trakcie obsługi ENTER. `setFontSize` przywraca wtedy fokus+zaznaczenie do edytora JESZCZE w trakcie tego zdarzenia, więc domyślna akcja Enter (nowa linia) trafia w przywrócone zaznaczenie i je kasuje. Klik poza pole (blur myszką) nie miał problemu — brak zdarzenia Enter.
- **Fix:** `onFontSizeInputEnter` robi `event.preventDefault()` (zabija domyślny Enter) + `setTimeout(() => input.blur(), 0)` — aplikacja przez blur dzieje się PO zakończeniu zdarzenia Enter, rozdzielona od niego. Ścieżka blur (klik) bez zmian.
### Verified
- GUI **221** (+2: ENTER preventDefault + odroczony blur; blur emituje fontSizeChange od razu).

## 2026-06-09 — Dialog hasła wyzwalany TEŻ przy ładowaniu z bazy (nie tylko „Plik→Otwórz")
### Changed
- **Root cause zgłoszenia „nie widzę gdzie podać hasło":** dialog pojawiał się tylko w ścieżce dyskowej (`loadDocument`), a realny przepływ (dashboard upload → editor `loadFromStorage`, odświeżenie, link z aplikacji zewn.) pobierał bajty i wołał `/open`, ale jego error-handler robił tylko `showError` — kod `PASSWORD_REQUIRED`/`WRONG_PASSWORD` ignorowany.
- Wspólny `_convertAndLoad(file, fileName, password?, announce?)` + `_applyLoadedContent` — używany przez OBIE ścieżki. `loadFromStorage` pobiera bajty → buduje `File` → `_convertAndLoad` (zamiast inline openDocument). Dialog hasła ma teraz callback ponawiający (`_passwordRetry`), więc działa niezależnie od źródła (dysk vs baza).
### Verified
- GUI `tsc` OK + **219** (+3: PASSWORD_REQUIRED z /open otwiera dialog; zatwierdzenie ponawia konwersję z hasłem; WRONG_PASSWORD pokazuje błąd).
### Notes
- Dashboard zapisuje zaszyfrowane bajty surowo (v1/v2), edytor odszyfrowuje przy otwarciu (prompt) — po edycji+autosave v2 staje się odszyfrowany. Read-only base (v1) zaprosi o hasło ponownie. Akceptowalne (v1 immutable = oryginał zaszyfrowany).

## 2026-06-09 — Stylizowany dialog hasła (zamiast window.prompt)
### Changed
- `document-editor`: `window.prompt` zastąpiony modalem `.password-dialog` spójnym z `leave-dialog` (overlay z blur, ikona kłódki, pole `type=password` z focus-ringiem, komunikat błędu, Anuluj/Otwórz). Sygnały `showPasswordDialog`/`passwordDialogValue`/`passwordDialogError` + `openPasswordDialog`/`confirmPasswordDialog`/`cancelPasswordDialog`. Enter=Otwórz, Esc/klik-tło=Anuluj, autofokus pola; przy błędnym haśle dialog wraca z komunikatem i ponawia `loadDocument(file, pwd)`.
### Verified
- GUI `tsc` OK + **216** (+3: confirm bez hasła→błąd, cancel czyści, confirm zamyka).

## 2026-06-09 — Dwa bugi: dekrypcja hasłem (NPOI nie działa→własna) + input font-size gubił selekcję
### Changed
- **Hasło — root cause i realna naprawa:** `NPOI 2.8.0` ma interfejs dekryptora, ale **NIE implementację Agile** (typ `AgileEncryptionInfoBuilder` nieobecny we wszystkich TFM-ach → `EncryptionInfo(fs)` rzuca `EncryptedDocumentException`), więc poprzedni kod mapował każdy plik (nawet z dobrym hasłem) na „WrongPassword". Nowy `OoxmlAgileDecryptor` — **własna, czysto zarządzana dekrypcja Agile** (MS-OFFCRYPTO §2.3.4.10+: SHA512 spinCount + AES-256-CBC, blockKeys, weryfikacja hasła, dekrypcja pakietu segmentami 4096B). NPOI używane już TYLKO do czytania kontenera CFB (`CreateDocumentInputStream`). `DocumentInputNormalizer` wymaga `EncryptionInfo` **i** `EncryptedPackage`; rozróżnia WrongPassword (niezgodność weryfikatora) od Invalid/Unsupported (Standard/CryptoAPI Office 2007 nieobsługiwany).
- **Font-size input (Qutas-FMT-004, naprawa właściwa):** `saveSelection()` zapisywał też ZWINIĘTE selekcje, a leciał na `blur`/`selectionchange` — czyli dokładnie gdy klik w input toolbara zwija zaznaczenie do karetki „wciąż w edytorze" → realne zaznaczenie nadpisywane pustą karetką → „nie ma na czym". Fix: nowy strażnik `editorHasFocus()` (activeElement w obszarze edytowalnym) — `saveSelection` zapisuje TYLKO gdy edytor ma fokus; przy przejściu do toolbara zostaje zaznaczenie z `mouseup`/`keyup`.
### Verified
- Backend Infrastructure **150** (+2: realny round-trip Agile encrypt→decrypt — poprawne hasło → odszyfrowany ważny DOCX otwierany OpenXml SDK; złe hasło → WrongPassword; brak hasła → PasswordRequired; enkryptor testowy też spec-zgodny). `D2ViewerEditor.sln` build OK. GUI **213** (+1: saveSelection nie gubi zaznaczenia po utracie fokusu).
### Notes
- Aktualizacja ADR-0010: NPOI **nie** dekryptuje (vs wcześniejszy zapis) — dekrypcja jest własna; NPOI zostaje tylko jako czytnik CFB. Standard/CryptoAPI (Office 2007) → status Invalid (rzadkie; ewentualny follow-up).

## 2026-06-09 — Otwieranie: DOCX z hasłem + .doc + opcjonalna klasyfikacja (Services)
### Changed
- **Hasło (Task 1):** nowy `IDocumentInputNormalizer` (Domain) + `DocumentInputNormalizer` (Infrastructure) — detekcja formatu po magic-bytes, dekrypcja DOCX zabezpieczonego hasłem przez **NPOI** (POIFS/Crypt, pure-managed, Linux/GCP-safe). Wpięty w `OpenDocumentQueryHandler` (przed `Convert`). `OpenDocumentQuery` + kontroler `/open` przyjmują `password`; sentinele `PASSWORD_REQUIRED`/`WRONG_PASSWORD` → 422 z `code`. GUI: `document.service.openDocument(file, password?)` + typed `OpenDocumentError`; edytor „Plik → Otwórz" przez `/open` z promptem hasła i retry.
- **.doc (Task 2):** normalizer wykrywa CFB: mislabeled-DOCX (ZIP w .doc) → pass-through (działa); binarny .doc (stream WordDocument) → `UNSUPPORTED_LEGACY_DOC` → 400 z instrukcją konwersji. Pickery (`dashboard`, edytor) dopuszczają `.doc`. **Brak wbudowanej konwersji binarnego .doc** — żaden pure-managed NuGet nie konwertuje (NPOI bez HWPF; b2xtranslator NuGet bez DocFileFormat; LibreOffice zakazany) → kontrolowany komunikat, nie udawanie. Patrz DECISIONS ADR-0010.
- **Klasyfikacja (Task 3):** `D2ServicesViewerEditor` `POST /api/v1/document` — pole `Classification` **opcjonalne** (brak/puste = bez klasyfikacji; podana musi być C1..C4). Metadata serializuje `classification` jako null gdy brak.
- **SkiaSharp** 3.116.1 → **3.119.2** (wyrównanie do tranzytywnej zależności NPOI, NU1605).
### Verified
- Backend: Application **8** (OpenDocument, +4 sentinel/password), Infrastructure **148** (+4 normalizer: ZIP/garbage/encrypted-no-pwd/binary-doc, fixtury CFB przez NPOI), Api **OpenDocument 4** (binary-doc→400+code, password→422). `D2ViewerEditor.sln` build OK. Services API build OK. GUI `tsc` OK + **212** testów.
### Notes
- Ścieżka „utwórz nowy z dysku" w dashboardzie zapisuje surowo (bez normalizera) → .doc/hasło kierujemy do edytora „Otwórz" (ścieżka /open). Normalizacja przy ingest/storage = ewentualny follow-up.
- Prompt hasła = `window.prompt` (funkcjonalne); stylizowany dialog = follow-up.

## 2026-06-08 — Zmiana czcionki (font-family) nie wpływała na nowo wpisywany tekst (FMT-005)
### Changed
- **Root cause:** `WysiwygEditorComponent.setFontFamily` wołał `editor.focus()` **przed** odtworzeniem zapisanej selekcji. Wybór czcionki z natywnego `<select>` w toolbarze zabiera fokus i czyści selekcję contenteditable; `focus()` przywracał karetkę na **początek dokumentu** (`rangeCount > 0`, więc strażnik `rangeCount === 0` nie odpalał `restoreSelection`), więc ZWS-span z `font-family` lądował w złym miejscu. Tekst wpisywany w realnej pozycji karetki dalej dziedziczył domyślną/firmową czcionkę. `setFontSize` miał już to zabezpieczenie — `setFontFamily` nie.
- **Fix:** `setFontFamily` odtwarza selekcję **ZANIM** dotknie fokusu (gdy live-selection nie jest w edytorze, a jest `savedSelection`) — lustrzane do `setFontSize`. Dodatkowo: gdy karetka siedzi już w ZWS-spanie z poprzedniego wyboru, aktualizuje jego `font-family` zamiast zagnieżdżać kolejny pusty span; zapisuje nową pozycję karetki + `updateFormattingState()`; po zmianie selekcji na zaznaczeniu odświeża `savedSelection`.
### Verified
- GUI `wysiwyg-editor.toolbar-actions.spec.ts` **7/7** (+1 nowy: `setFontFamily` odtwarza zapisaną selekcję, gdy fokus opuścił edytor). `tsc --noEmit` OK.
- Manualnie do potwierdzenia w przeglądarce (jsdom nie odwzorowuje contenteditable/Selection/focus): wybór czcionki przy zwiniętej karetce → nowo wpisany tekst dostaje wybraną czcionkę i round-tripuje do `w:rFonts`.
### Notes
- `pendingFontFamily`/`pendingFontSize` nadal tylko zapisywane (nieczytane) — ścieżka „brak selekcji" polega na ZWS-spanie; nie ruszane w tej zmianie.

## 2026-06-08 — EMF/eksport + font-size import + page-break import + selekcja font-size
### Changed
- **Qutas-IMG-009 (EMF psuje DOCX) — root cause + fix:** writer pisał placeholder SVG jako goły `a:blip` (SVG bez rastra = NIEPOPRAWNY OOXML → Word „uszkodzony"). Reader (`DocxToHtmlConverter`) niesie teraz oryginalny metafile w `data-original-src`; writer (`ResolveImageSrc`) preferuje go i zapisuje prawdziwy **EMF part** (`ImagePartType.Emf`, Word renderuje natywnie); `BuildImageDrawing` ma **twardy guard**: nigdy nie emituje gołego SVG blip (return null). Mapowanie x-emf/x-wmf dodane.
- **DWA pre-existing bugi schematu w `styles.xml`** (wykryte OpenXmlValidatorem, psuły KAŻDY zapis): heading `w:rPr` miał złą kolejność (`sz`/`color`/`b`/`i`) → poprawione na `rFonts→b→i→color→sz`; heading `w:pPr` miał `spacing` przed `keepNext` → `keepNext→keepLines→spacing→outlineLvl`.
- **Qutas-IMP-005 (14pt → ~10pt) — root cause frontend:** reader poprawny (4 testy: direct/styl/docDefaults/Normal). Bug: `_flattenTopBlocks` ROZWIJA wrapper `.document-content`, na którym reader trzyma default → ginął. Fix: `_captureDocumentDefaults` czyta font-size/family z wrappera; nowe sygnały `documentDefaultFontSize/Family` zbindowane na `[style.font-size/font-family]` contenteditable strony.
- **Qutas-IMP-008 (3 strony → 1):** reader IGNOROWAŁ `w:pageBreakBefore`. Fix: `HasPageBreakBefore` → emit `<div class="page-break">` przed akapitem (ten sam mechanizm co manualny break; writer round-tripuje).
- **Qutas-FMT-004 (input font-size gubił selekcję) — root cause:** `savedSelection` zapisywany tylko na `blur` (zbyt późno — selekcja już znika przy klik w input). Fix: `onSelectionChange` zapisuje selekcję na bieżąco. Dodatkowo Enter w input nie aplikuje podwójnie (był apply + blur→apply → zagnieżdżone spany).
### Verified
- Backend `dotnet test Infrastructure.UnitTests` → **144 passed** (+EMF round-trip z OpenXmlValidator=0 błędów, +4 FontSizeImport, +2 pageBreakBefore). Frontend `ng test` → **211 passed** (+2 captureDocumentDefaults).
### Notes
- EMF: dla osadzonego rastra w EMF/WMF zapis idzie rastrem (PNG/JPEG) — też poprawny. Pełna rasteryzacja EMF w przeglądarce nadal roadmapa (placeholder), ale ZAPIS jest teraz wierny (oryginalny EMF) i niełamiący.

## 2026-06-08 — Edytor: 4 realne naprawy (paginacja / Backspace / context-menu / PDF diag)
### Changed
- **Qutas-PAG-002 (Enter wydłuża stronę):** `_schedulePaginate` debounce **600 → 250 ms** (komentarz mówił 300 — kod zdryfował). Strona `min-height:1122px; overflow:visible` rosła widocznie przez całe okno debounce zanim treść spłynęła; 250 ms eliminuje rozciąganie poza A4.
- **Qutas-PAG-001 (Backspace na 2. stronie):** nowy `_tryMergeAcrossPageBackwards()` wpięty w `handleKeyboard` (`_tryDeletePageBreakBackwards() || _tryMergeAcrossPageBackwards()`). Wcześniej obsługiwany był TYLKO manualny page-break; strony z AUTO-paginacji nie scalały się (osobne contenteditable → przeglądarka nie łączy). Teraz: pusty blok wiodący → usuń; bloki mergeable (P/DIV/H/LI) → scal treść jak Word; tabela/niekompatybilne → tylko nawigacja karetki (treść nietknięta). Po operacji repaginacja spływa treść w górę.
- **Qutas-UI-006 (context-menu zasłania UI):** `onContextMenu` clamp z `Math.max(8, …)` na obu osiach — wcześniej dolny clamp dawał ujemne `y` na niskim oknie → menu nad viewportem zasłaniało toolbar.
- **Qutas-PDF-010 (PDF „prace techniczne"):** `pdf-viewer.ts` `catch {}` → `catch (err)` z `console.error` + rozróżnieniem awarii workera (cold-start/404/MIME `.mjs`) od uszkodzonego pliku. Przestaje maskować przyczynę.
### Verified
- `npx tsc --noEmit` OK; `npx ng test --watch=false` → **209 passed** (+4 merge w `caret-field.spec.ts`, +3 context-menu w `document-editor.spec.ts`).
### Notes
- Diagnoza Problem 4 (font-size input): apply-path JEST poprawny (`applyFontSizeToSelection` re-selektuje wstawioną treść 1996–2006; `setFontSize` odtwarza `savedSelection`). Residualne ryzyko = stała `Range` po repaginacji — wymaga selekcji path-based (większa zmiana, nie ruszane).

## 2026-06-08 — Akapit: „Ustaw jako domyślne" zapisuje zamiast resetować (Qutas-PAR-007)
### Changed
- `document-editor.html` — przycisk „Ustaw jako domyślne" wołał `resetParagraphDefaults()` (reset do wartości bazowych) → teraz `setParagraphAsDefault()`.
- `document-editor.ts` — usunięto `resetParagraphDefaults()`; dodano `setParagraphAsDefault()` (snapshot bieżących ustawień → `_paragraphDefaults` + `applyParagraphSettings()` na bieżącym akapicie). Nowe pole `_paragraphDefaults` (default per-sesja). `readCurrentParagraphSettings()` przy braku selekcji seeduje formularz z `_paragraphDefaults` zamiast zostawiać stałe wartości.
- Model trwałości: per-sesja edytora (reset po odświeżeniu); bez localStorage (brak wzorca w app) i bez zmian API. Nowe akapity dziedziczą styl po bieżącym bloku (contenteditable) → default propaguje się naturalnie.
### Verified
- `npx ng test --watch=false` → **202 passed** (+3 nowe w `document-editor.spec.ts`: brak resetu do bazowych / seed dialogu z defaultu / kopia niezależna). Reszta pakietu bez regresji.
### Notes
- Pełna diagnostyka 11 zgłoszeń edytora w raporcie sesji; pozostałe 10 były już wdrożone w sesjach 2026-06-03/04 (zweryfikowane w drzewie: metody klawiatury/paste/font-size/page-break/grafiki obecne). Ten wpis dotyczy jedynego niewdrożonego wcześniej zgłoszenia.

## 2026-06-04 — Health-check API: pierwszy strzał natychmiast po starcie + dedup
### Problem
- Komunikat „brak komunikacji z API" pojawiał się zbyt późno. Pierwszy check zależał od **implicit constructor side-effect** `ConnectionStatusService` (lazy `providedIn:'root'`) — nieodporne na regresję; brak strażnika duplikatów (start + zdarzenie `online` + interwał mogły wystrzelić równoległe żądania).
### Changed (frontend)
- `connection-status.service.ts`: prywatne `checkApi()` → **publiczne `checkNow()`** z **strażnikiem in-flight** (`_checkInFlight`) — pojedyncze żądanie w locie, brak duplikatów. Guard `if (!apiConfig.baseUrl) return` (brak fałszywego offline zanim URL znany — tu URL statyczny z `environment`). Konstruktor, interwał (30 s) i handler `online` wołają `checkNow()`.
- `app.config.ts`: `provideAppInitializer` woła **`inject(ConnectionStatusService).checkNow()`** (jawnie, nie polega na side-effekcie konstruktora) — pierwszy check zaraz po inicjalizacji aplikacji, przed renderem komponentów. Non-blocking (nie zwraca/await). Strażnik in-flight scala check z konstruktora i z initializera w **jedno** żądanie.
- Subsequent polling bez zmian (30 s `setInterval`, retry/timeout 5 s). Banner/offline-status bez zmian.
### Verified
- GUI **199** (+2: „flags API unreachable quickly on first check"; „checkNow() nie duplikuje w locie"). Istniejące 5 testów `ConnectionStatusService` zielone (konstrukcja nadal = 1 natychmiastowe żądanie).
- Scenariusz manualny: start przy API down → pierwszy `/health` od razu po starcie → status 0 → banner offline „szybko"; start przy API up → `/health` 200 → brak bannera.

## 2026-06-04 — Konwerter grafik VML/EMF/WMF → web (pure-managed, bez LibreOffice)
Pełna referencja: `.ai/GRAPHICS_CONVERSION.md`.
### Problem
- Reader emitował EMF/WMF media parts jako `data:image/x-emf|x-wmf` → przeglądarka nie renderuje → złamany obraz. VML kształty wektorowe nieobsługiwane.
### Architektura (DDD)
- `IGraphicConversionService` (Domain/Interfaces) + modele `GraphicSource`/`WebGraphicRepresentation`/`GraphicConversionResult`/`GraphicConversionDiagnostics` + enumy (Domain/Models).
- `GraphicConversionService` (Infrastructure) — **pure-managed**, bez LibreOffice/GDI/System.Drawing → Linux/GCP-safe. Detekcja (magic bytes + content-type), wymiary z nagłówków (PNG/JPEG/GIF/BMP/EMF rclFrame/WMF placeable), ekstrakcja osadzonego rastra z EMF/EMF+, placeholder SVG, sanitizer SVG, VML rect/oval/line/roundrect→SVG.
- Hook w `DocxToHtmlConverter.WebGraphicForLegacy` — EMF/WMF (a:blip i v:imagedata) → renderowalny `data:URL` (osadzony raster albo placeholder), web-native bez zmian. Atrybut `data-legacy-graphic="placeholder"`.
### Strategia
- EMF/WMF **nie rasteryzowane** (brak bezpiecznej pure-managed ścieżki na Linux) → placeholder + **pass-through oryginalnego partu** (`ConvertPreservingPackage`) — Word renderuje prawdziwą grafikę, dokument nietknięty. Honest fallback, nie udawanie.
- Bezpieczeństwo: `DtdProcessing.Prohibit`/`XmlResolver=null` (XXE blok), SVG sanitizer (script/on*/external/javascript), limity rozmiaru/czasu/wymiarów + cancellation; wyjątki → Fallback (nie wywraca importu).
### Decyzja: zero nowych zależności graficznych
- LibreOffice zakazane; System.Drawing Windows-only; brak permisywnego pure-managed renderera EMF. Rasteryzacja przez out-of-process sidecar = roadmapa.
### USUNIĘTO poprzednią ścieżkę EMF opartą o soffice + System.Drawing (krytyczne dla GCP)
- Reader miał `TryConvertMetafileToPng` → **LibreOffice `soffice`** (zakazane) + fallback **`System.Drawing`** (Windows-only, `PlatformNotSupported` na Linux) → działało tylko na Windows, **nie na GCP**. Usunięto wszystkie te metody + **pakiet `System.Drawing.Common` z .csproj**; oba wywołania (`LoadImageFromPart`, picture-bullet) przepięto na pure-managed `GraphicConversionService`. `SkiaSharp`(+NativeAssets.Linux) zostaje (GCP-safe, tylko barcody). Weryfikacja: `grep System.Drawing|soffice` w .cs = pusto; build OK; Infrastructure 137/137.
### Verified / tests
- Infrastructure **137** (+21: `GraphicConversionServiceTests` 13, `GraphicConversionSecurityTests` 7, `GraphicConversionIntegrationTests` 1 — DOCX z EMF a:blip → reader emituje svg+xml, **nigdy** image/x-emf). Brak regresji istniejących testów obrazów.
- Benchmark `GraphicConversionBenchmarks` (BenchmarkDotNet `[MemoryDiagnoser]`, param GraphicCount 1/10/100) — `dotnet run -c Release -- --filter *GraphicConversion*`.
### Ograniczenia
- EMF/WMF bez podglądu wektorowego w przeglądarce (placeholder; pełny render w Word). VML: tylko bezpieczny podzbiór kształtów; hook kształtów wektorowych do readera + SVG→DOCX PNG-fallback + cache = roadmapa.

## 2026-06-04 — Edytor: 8 błędów (ENTER, font-size, font-family, paste plain, interlinia, link, .doc, Backspace/strzałki)
Pełna referencja: `.ai/EDITOR_KEYBOARD.md`.
### Issue 1 — ENTER cofał kursor do poprzedniej linii (FIX)
- Root cause: po wpisaniu znaku edytor repaginuje (~600 ms) i przebudowuje DOM stron; karetkę odtwarzał **globalny offset tekstowy**, niejednoznaczny na granicy bloków — nowy **pusty** akapit po ENTER ma 0 znaków → restore lądował na końcu poprzedniego akapitu.
- Fix: kotwica karetki **`{ block, offset }`** (`_saveGlobalCaret`/`_restoreGlobalCaret` + współdzielony `_flattenTopBlocks`, `_placeCaretAtTextOffset`). Pusty blok → karetka na początku bloku. Indeks bloku stabilny między repaginacjami.
### Issue 2 — font-size przez input ukrywał tekst (FIX)
- Root cause: wartość `0`/NaN dawała `0pt`/`NaNpt` → tekst niewidoczny. Toolbar walidował, ale publiczne `setFontSize` nie.
- Fix: `setFontSize` odrzuca `!finite`/`<1`/`>400` (defense-in-depth).
### Issue 3 — font-family „nie działała"/wracała do domyślnej (FIX, root cause backend)
- Root cause: `innerHTML` serializuje nazwy wielowyrazowe jako encję `&quot;`; writer (`ApplyRunStyle`) regex `[^,;]+` ucinał nazwę na `;` **wewnątrz `&quot;`** → `w:rFonts ascii="&quot"` → Word wracał do domyślnego fontu. (Single-word jak Arial działały.)
- Fix: **HTML-decode stylu** przed regexem + odcięcie cudzysłowów i fallbacku generycznego. `'Font',serif` / `"Font"` / `Arial` → czysta nazwa.
### Issue 4 — „Wklej bez formatowania" nie działało (FIX)
- Root cause: klik w pozycję menu zabierał fokus edytorowi → `execCommand('insertText')` bez karetki nic nie wstawiał.
- Fix: `insertText` odtwarza zapisaną selekcję + fokus przed wstawieniem; `pasteWithoutFormatting` dostał `.catch`. Ctrl+Shift+V działał wcześniej.
### Issue 5 — interlinia względem Word (ANALIZA + regresja)
- Ustalenie: mapowanie OOXML→CSS jest **poprawne** (`auto`→`line/240` bezjednostkowe; `exact`/`atLeast`→`pt`; before/after→margin pt; bez podwójnego liczenia). Pinned testami `LineSpacingMappingTests`.
- Ograniczenia nieusuwalne mapowaniem (HTML/CSS): „single" ≠ `line-height:1` (Word dolicza line-gap fontu); kolaps sąsiednich marginesów vs sumowanie before+after w Word. Udokumentowane.
### Issue 6 — wstawianie linku nie działało (FIX)
- Root cause: pole URL dialogu zabierało fokus → selekcja kolabowała → `createLink` bez celu.
- Fix: `insertLink` odtwarza zapisaną selekcję, **normalizuje URL** (`normalizeLinkUrl`), escapuje etykietę; zwinięta karetka → `<a … rel=noopener>`. Writer już emituje `w:hyperlink`+relację.
### Issue 7 — brak obsługi `.doc` (DECYZJA: Wariant B — jawne odrzucenie)
- `.doc` = stary binarny format (OLE/CFBF), nie OOXML; brak konwertera DOC→DOCX w pipeline. Odrzucenie z instrukcją konwersji zamiast udawania obsługi: backend `OpenDocument` (`.doc`→400) + frontend `dashboard.openFile`. Wariant A (LibreOffice headless) opisany jako przyszłość.
### Issue 8 — Backspace nie usuwał stron; strzałki „skakały" (FIX częściowy)
- Backspace na **początku strony** usuwa **manualny page-break** poprzedniej strony (`_tryDeletePageBreakBackwards` + `_isCaretAtEditorStart`/`_removeTrailingPageBreak`) i repaginuje; **nie** rusza wrappera strony.
- Nawigacja ArrowDown/Up między stronami była dodana wcześniej (`_tryMoveCaretAcrossPages`); „skakanie" złagodzone przez block-aware karetkę (Issue 1). Detekcja skrajnej linii (`_isCaretOnEdgeLine`) jest layout-zależna — pozostaje obszar do dostrojenia (manualnie).
### Verified / tests
- Backend: Infrastructure **116** (+3 `FontFamilyWriteTests`, +6 `LineSpacingMappingTests`), Api **34** (+1 `.doc`), Application **193**, Domain **57** — zielone.
- GUI: **197** (+3 `enter-caret`, +6 `toolbar-actions` font-size/link/paste, +3 Backspace page-break w `caret-field`).
- jsdom nie pokrywa contenteditable/execCommand/layout → wstawianie linku/tekstu i repaginacja end-to-end mają **scenariusze manualne** (`EDITOR_KEYBOARD.md` §7). Import/eksport/autozapis/undo/tabele/obrazy nietknięte (te same testy zielone).
### Ograniczenia
- Block-index karetki: rzadki edge gdy tabela **przed** karetką zmienia podział w danym przebiegu. Interlinia: różnice metryk fontu i kolapsu marginesów (jw.). `.doc`: brak realnej konwersji (świadomie). Strzałki: edge-line detection layout-zależna.

## 2026-06-04 — Layout: banner środowiska spychał dashboard w dół (fix shell `100vh`→`100%`)
### Diagnoza
- App shell (`d2-root`, global `styles.scss`) jest flex-column `height:100vh; overflow:hidden`; globalne bannery (`d2-global-banners` = environment + offline) mają `flex:0 0 auto` (intrinsic height), a obszar strony flexuje resztę: `d2-root > d2-dashboard/…-editor/…-pdf-viewer/…-pdf-maintenance/…-admin-shell { flex:1 1 0; min-height:0; overflow:hidden }`. **Banner jest w normalnym flow i to jest poprawne** — host strony dostaje `100vh − banner`.
- **Edytor (wzorzec OK):** `:host{height:100%}` → `.document-editor-container{height:100%}` — wypełnia obszar przydzielony przez shell, banner uwzględniony automatycznie.
- **Bug:** `dashboard.scss .dashboard-wrapper{min-height:100vh}` oraz `pdf-maintenance .maintenance-wrapper{min-height:100vh}` wymuszały **pełną wysokość viewportu** wewnątrz hosta wysokiego na `100vh − banner` → wrapper wystawał o wysokość bannera, host (`overflow-y:auto`) scrollował, treść wizualnie zjeżdżała w dół, stopka ucinana. Klasyczny „dashboard obniżony o wysokość labela".
### Changed (frontend)
- `dashboard.scss`: `.dashboard-wrapper` `min-height:100vh` → **`min-height:100%`** (wypełnia host przydzielony przez shell, nie viewport). Komentarz wyjaśniający.
- `pdf-maintenance.ts` (inline styles): `.maintenance-wrapper` `min-height:100vh` → **`min-height:100%` + `height:100%`** (centrowanie w obszarze hosta, nie viewportu).
- **Bez** magicznych marginesów, bez zmiennej `--banner-height`, bez `!important`, bez globalnego CSS — fix to ujednolicenie do banner-agnostycznego idiomu `100%-of-host`, którego shell już używa dla edytora. `app.scss .app-container{min-height:100vh}` to martwy CSS (brak `styleUrl` w `app.ts`, inline template bez tej klasy) — pozostawiony bez zmian.
### Verified / tests
- Nowy `pages/layout-shell.spec.ts` (2): dashboard i pdf-maintenance — wstrzyknięty scoped CSS wrappera **nie zawiera `100vh`** i ma `min-height:100%` (kontrakt layoutu; jsdom nie robi layoutu, więc asercja na faktycznie zregresowanym CSS, nie na pikselach). GUI **185** (było 183, +2). Edytor nietknięty (te same testy zielone).
### Scenariusze manualne
- Dashboard z bannerem DEV: „Witaj w Qutas" wyśrodkowane, stopka `© 2026 Qutalo` widoczna, brak scrolla/skoku. Bez bannera (PROD): identyczny układ. Edytor (screen A): toolbar/panele/stopka bez zmian. Środowiska Local/DEV/TST/PRE: różny kolor bannera, **ta sama wysokość** layoutu. Małe rozdzielczości: gdy treść > obszar, host scrolluje (nic nie ucięte).
### Ograniczenia
- `min-height:100%` wymaga definite-height hosta — zapewnia go shell (`flex:1 1 0` w `d2-root height:100vh`); ten sam mechanizm, na którym opiera się edytor.

## 2026-06-04 — 3 poprawki: font fallback (serif), nawigacja kursora między stronami, rozmiar numeru strony w stopce
### P1 — font fallback wg rodziny (serif/mono/sans) zamiast twardego `sans-serif`
- Diagnoza: reader emitował każdy font jako `font-family:'X',sans-serif`. Gdy `Times New Roman` niedostępny w przeglądarce, przeglądarka spadała na **bezszeryfowy** default (wygląd jak Calibri), mimo że nazwa była zachowana.
- Changed (reader `DocxToHtmlConverter`): `FontFamilyCss(name)` + `GenericFontFallback(name)` — szeryfowe (times/cambria/georgia/garamond/minion/palatino/book antiqua/„serif") → `serif`, courier/consolas/mono → `monospace`, reszta → `sans-serif`. Zastąpiono 3 miejsca emisji `,sans-serif`. **Bez** hardcode konkretnego fontu, bez dosadzania plików czcionek, bez zmiany fontu dokumentu.
### P2 — ArrowDown/ArrowUp nie przechodził między stronami (osobne contenteditable per strona)
- Changed (frontend `wysiwyg-editor`): w `handleKeyboard` (przed blokiem Ctrl) gdy ArrowDown/ArrowUp bez modyfikatorów i `_tryMoveCaretAcrossPages(dir)` zwróci true → `preventDefault`. Nowe: `_tryMoveCaretAcrossPages` (zwinięta karetka na skrajnej linii → przeniesienie do sąsiedniej strony), `_isCaretOnEdgeLine` (porównanie rectów karetki z górną/dolną linią edytora), `_placeCaretAtEditorEdge` (focus + range na start/end + aktywacja strony). Shift/Ctrl/Alt i zaznaczenia (range) nietknięte (guard `isCollapsed` + brak modyfikatorów).
### P3 — numer strony w stopce ignorował rozmiar fontu stopki
- Diagnoza: pole PAGE było emitowane „goło", numer dziedziczył rozmiar kontenera (np. 10.5pt) zamiast runów stopki (np. 8pt).
- Changed (reader): `FieldSpan(cssClass, placeholder, run)` emituje `<span class=… style="{run rPr CSS}">` dla PAGE/NUMPAGES/DATE (simple + complex field) — numer niesie font-size własnego runu.
- Changed (writer): `BuildFieldRun` buduje poprawny `w:fldSimple` z **wewnętrznym runem** niosącym `rPr` (font-size) — zamiast wcześniejszego (schema-niepoprawnego) `SimpleField` wewnątrz `Run`, który po round-tripie gubił właściwości. `FlushPending` w stopce/nagłówku uznaje też paragraf zawierający `SimpleField` (wcześniej wymagał `Run`/`Hyperlink` → pole było pomijane).
- Changed (frontend): `_pageNumberHtml(editor)` + `_inlineFieldFontSize(editor)` — wstawiany `.page-number` dostaje font-size pierwszego inline-sized spanu stopki (fallback computed). CSS `.field-page/.field-numpages/.field-date/.page-number { font-size:inherit; … }`.
### Verified / tests
- Backend: Infrastructure **107** (+4: `DefaultStyleFontTests` serif/sans fallback; `PageFieldFontTests` reader emituje font-size pola + writer round-trip `w:fldSimple` z rPr=16). Application **193**, Domain **57** — zielone. (Api.UnitTests pominięte: lock DLL od działającej instancji API/debuggera — nie dotyczy zmienionego kodu Infrastructure.)
- GUI: **183** (+7: `wysiwyg-editor.caret-field.spec` — ArrowDown/Up przechodzi między stronami, range/Shift nie rusza karetki, brak ruchu poza ostatnią/przy jednej stronie, `_pageNumberHtml`/`_inlineFieldFontSize` dziedziczą font-size).
### Ograniczenia
- `_isCaretOnEdgeLine` jest layout-zależne (rect karetki) — nieodtwarzalne w jsdom (brak `Range.getClientRects`); w testach stubowane, logika nawigacji/granic testowana realnie.
- Font fallback poprawia rendering, ale nie podstawia brakującego kroju — przy braku Times New Roman tekst renderuje się szeryfowym fontem systemowym (świadomie, bez licencjonowania krojów).

## 2026-06-03 — Font treści (Times New Roman→Calibri) + page break „od nowej strony" (display)
### Diagnoza — font
- Oryginał `orginał_GOOD`: domyślny styl akapitu **`Normalny` (`w:default="1"`)** ma `<w:rFonts w:ascii="Times New Roman"/>` (sz=21 → 10.5pt). To **nadpisuje** docDefaults (`asciiTheme=minorHAnsi` → theme minor = Cambria). Runy body mają tylko `w:cs="Times New Roman"` (complex-script), bez `ascii`.
- Bug: reader `LoadDocDefaults` czytał **wyłącznie docDefaults** (→ Cambria), ignorował font domyślnego stylu akapitu → kontener `.document-content` dostawał Cambria, a edytor i tak spadał na własny default (Calibri). Stąd Calibri na ekranie.
### Diagnoza — page break
- To manualny `<w:br w:type="page"/>` we własnym akapicie (po tabeli z podpisami, przed „PROTOKÓŁ…"). Round-trip był już naprawiony (R-15: writer → `w:br`; reader → top-level `<div class="page-break">`). **Pozostał problem DISPLAY**: edytor `_splitHtmlIntoPages` **konsumował** marker przy podziale, więc `_repaginateNow` (re-paginacja wg wysokości) go nie widziała i scalała treść → „PROTOKÓŁ" pod podpisami.
### Changed — backend (reader)
- `DocxToHtmlConverter.ApplyDefaultParagraphStyleFont`: po wczytaniu stylów czyta font domyślnego stylu akapitu (`w:default="1"`, typ paragraph) i ustawia go jako `_defaultFontFamily`/rozmiar kontenera (override docDefaults). Z dokumentu, nie hardcode. **Weryfikacja na realnym pliku**: kontener `font-family:'Times New Roman',sans-serif;font-size:10.5pt`.
### Changed — frontend (display)
- `wysiwyg-editor._splitHtmlIntoPages`: **zachowuje** marker `<div class="page-break">` (doklejony na końcu każdej strony poza ostatnią) → `_repaginateNow` honoruje podział (`_isPageBreakBlock` wymusza nową stronę), marker przeżywa zapis (getContent → writer → `w:br`). `setContent` join plain (bez dublowania); usunięty martwy `_joinPagesWithBreaks`.
### Verified / tests
- Backend **386 pass** (+4 `DefaultStyleFontTests`: default-style override, docDefaults fallback, direct docDefaults, pass-through zachowuje font). GUI **176** (+1 `_splitHtmlIntoPages` zachowuje marker + round-trip). Golden snapshoty bez zmian (syntetyki nie mają default-style fontu). Solucja zielona.
### Ograniczenia
- Font: jeśli `Times New Roman` niedostępny w przeglądarce → fallback `,sans-serif` (nazwa zachowana w CSS; rendering zależny od systemu). Toolbar edytora może pokazywać własny default fontu (kosmetyka) — renderowany tekst używa fontu kontenera.
- Save font: pełna wierność na zapisie zależy od pass-through (R-16); ścieżka autosave bezstanowa (R-20) wciąż regeneruje style.
- Page break: honorowane są **manualne** breaki; naturalna paginacja Worda nadal liczona wg wysokości (nie 1:1).

## 2026-06-03 — Benchmarki + fidelity-checks (perf/pamięć/regresja wierności DOCX)
### Added — backend (`D2ViewerEditor.Benchmarks`, BenchmarkDotNet 0.14)
- `BenchmarkAssets` — syntetyczne DOCX in-memory (simple/tables/large-table/images/header-footer/page-breaks/multi-style) + opcjonalny realny plik regresyjny przez env `D2_BENCH_ASSETS` (domyślnie repo; używany tylko gdy istnieje). Brak commitowania wrażliwych plików.
- `DocxImportBenchmarks`/`DocxExportBenchmarks`/`DocxRoundTripBenchmarks`/`TableBenchmarks` — `[MemoryDiagnoser]`, baseline per grupa, params Rows=100/400 (`[ShortRunJob]` na large-table). Mierzą reader/writer/`ConvertPreservingPackage`/round-trip.
- `DocxPackageReport` (analizator pakietu ZIP/regex) + `FidelityComparison` (PASS/WARN/FAIL, pass-through-aware) + `FidelityReportRunner` (markdown, report-only; `--fidelity [--fail-on-regression]` w `Program.cs`).
### Added — fidelity gate (NUnit, CI)
- `DocxFidelityRegressionTests` (4): R-16 styles/theme zachowane (pass-through), R-15 page breaki, R-17 tabele niemnożone, R-18 brak sztucznych `trHeight`. Parser-niezależny (ZIP/regex).
### Added — frontend perf harness (Vitest, report-only)
- `wysiwyg-editor.perf.spec.ts` (4): timing `getContent`/merge split-table/`setContent` (Performance API, lenient ceiling 4000 ms). Ograniczenie: jsdom nie mierzy layoutu/paginacji (potrzebny Playwright).
### Verified
- Backend **382 pass** (+4 fidelity gate). GUI **175** (+4 perf). Benchmarki uruchamiają się (dry-job: import ~85 ms cold). Fidelity report: syntetyki **PASS**, realny `orginał_GOOD` (pass-through) **WARN** (164 style + Cambria + 6 tabel + 1 page break zachowane; tylko `numbering.xml` nieprzeniesiony — R-19).
### Docs
- Nowy `.ai/BENCHMARKS.md` (jak uruchomić perf/fidelity/frontend, progi, baseline, CI report-only→fail-on-regression, ograniczenia) + `INDEX`.

## 2026-06-03 — Page break „od nowej strony": edytor nie honorował manualnego podziału (display)
### Problem
Manualny page break z DOCX (np. przed „PROTOKÓŁ WYDANIA POJAZDU") nie zaczynał nowej strony w edytorze — treść lądowała tuż pod poprzednią. Round-trip zapisu działał (R-15), ale **paginacja widoku** ignorowała break.
### Root cause
Reader emitował break-only akapit (`<w:p><w:r><w:br type=page/></w:r></w:p>`) jako **zagnieżdżony** `<p><span><div class="page-break"></div></span></p>`. To: (1) psuło `_splitHtmlIntoPages` (regex dzielił string w środku `<p>`), (2) `_repaginateNow` paginował tylko wg wysokości — ignorował break.
### Changed
- **Reader** `DocxToHtmlConverter`: nowy `IsPageBreakOnlyParagraph` — akapit zawierający wyłącznie `w:br type=page` (bez tekstu/grafiki) emitowany jako **top-level** `<div class="page-break"></div>` (zamiast zagnieżdżony). Writer nadal mapuje go na `w:br type=page` (R-15).
- **Frontend** `wysiwyg-editor.ts`: `_repaginateNow` wymusza **nową stronę** na bloku page-break (`_isPageBreakBlock` — top-level lub zagnieżdżony bez tekstu); marker zostaje w treści → przeżywa zapis.
### Verified / tests
- Backend **378 pass** (+1 reader: break-only akapit → top-level block + round-trip=1). GUI **171** (+1 `_isPageBreakBlock`). Golden snapshoty bez zmian (brak page-breaków). Solucja zielona.
### Uwaga
- Wymaga **ponownego wczytania dokumentu** w edytorze, by zobaczyć efekt (paginacja liczona przy load/edycji).

## 2026-06-03 — R-15/R-16/R-17: page break round-trip, pass-through pakietu DOCX, scalanie split-table
### R-15 — manualny page break round-trip (writer)
- `HtmlToDocxConverter`: nowy `IsPageBreakNode` (class `page-break` / `data-docx-break="page"` / CSS `page-break-before`/`break-before:page`). `CreateRunsFromNode` konwertuje **zagnieżdżony** `<div class="page-break">` (reader emituje go wewnątrz akapitu) na `w:br type=page`; blok-level używa tej samej detekcji. Wcześniej break ginął (AFTER=0). Bez duplikacji; dokument bez breaków nie dostaje żadnego. Testy: `PageBreakRoundTripTests` (5).
### R-16 — pass-through oryginalnego pakietu (writer + handler)
- `IHtmlToDocxConverter.ConvertPreservingPackage(html, Stream? original, ...)`: generuje DOCX jak `Convert`, po czym **zachowuje z oryginału** `styles.xml` (pełny zestaw, w tym style tabel), `theme` i `fontTable` (FeedData do istniejących partów). Body/sekcja/nagłówki-stopki/obrazy/numbering pochodzą z konwersji HTML. `null`/pusty oryginał → zachowuje się jak `Convert` (fallback). Best-effort: błąd oryginału → fallback (nie psuje zapisu).
- Wiring: `DownloadEditedDocumentCommandHandler` ładuje **bazową wersję (v1)** przez `IDocumentStorageService.DownloadAsync` i używa pass-through dla DOCX; fallback gdy brak wersji/nie-DOCX. Testy: `PassThroughPackageTests` (5) + handler test pass-through (1).
### R-17 — scalanie fragmentów split-table (frontend)
- `wysiwyg-editor.ts`: `_splitTableForPagination` taguje fragmenty jednej logicznej tabeli `data-split-table-id` + klonuje `colgroup`; `getContent()` woła `_mergeSplitTables` (scala sąsiednie fragmenty o tym samym id, zachowuje wiersze/kolumny, sprząta marker). Niezależne sąsiednie tabele NIE są scalane. Testy: Vitest (2).
### Verified (realny plik: reader → ConvertPreservingPackage(html, oryginał GOOD) → AFTER)
| metryka | GOOD | BAD | AFTER |
|---|---|---|---|
| styles.xml liczba stylów | 164 | 16 | **164** (part `styles2.xml`, relacja+content-type OK — Word czyta po relacji) |
| theme minor font | Cambria | Calibri | **Cambria** |
| twarde page-breaki | 1 | 3 | **1** |
| marginesy / tabele | wzorzec | rozjechane | ≈ GOOD (z poprz. fixów) |
### Tests
- Backend **377 pass / 0 fail / 6 skip** (Infrastructure +10: 5 page-break + 5 pass-through; Application +1). GUI **170** (+2 split-table). Pełna solucja zielona.
### Ograniczenia (R-19..R-21)
- **R-16 niepełny pass-through:** zachowane tylko `styles.xml`/`theme`/`fontTable`; `numbering.xml` NIE jest przenoszony (uniknięcie konfliktu numId z generowanym body) — dla dokumentów z listami custom-bullety mogą się różnić. Part stylów zapisywany jako `styles2.xml` (artefakt SDK; relacja poprawna, Word czyta). **Wiring tylko w DownloadEditedDocument** — ścieżka autosave (`/api/document/save`, bezstanowa) nadal generuje od zera; pełne wpięcie wymaga master-aware endpointu konwersji (next).
- **R-16 model body:** body nadal z konwersji HTML (inline style), nie referencje oryginalnych nazwanych stylów; font dziedziczy z zachowanego docDefaults (Cambria), ale akapity nie wracają do oryginalnych `w:pStyle`.

## 2026-06-03 — Round-trip fidelity: naprawa rozjazdu 4→7 stron i powiększonych tabel (orginał_GOOD vs zapisany_BAD)
### Root causes (z analizy XML dwóch plików)
- **Marginesy napompowane**: writer `AddPageSettings` `Math.Max(topTwips, headerHeightTwips+720)` + hardkod `Header/Footer=720` → top 567→1281, bottom 851→1231 (mniej treści/stronę).
- **Tabele wyższe**: reader `ConvertTableToHtml` domyślny `padding:4px 8px` (4px=60 tw góra/dół) do każdej komórki → BAD +8160 tw (14.4 cm) `tcMar`. Word default = 0 góra/dół.
- **Tabele szersze**: writer domyślnie `tblW pct 5000` (100%) gdy brak/auto width → rozciągnięcie tabel content-sized do pełnej szerokości. + hardkod `tblCellMar 40/80`.
- **+2 twarde page-breaki (4→7)**: edytor `getContent()` materializował auto-paginację wg wysokości (`_repaginateNow`) jako `<div class="page-break">` między każdą stroną → twarde `<w:br type=page>` w DOCX.
### Changed — backend (`HtmlToDocxConverter`/`DocxToHtmlConverter`)
- Reader: domyślny padding komórek = Word default (top/bottom **0**, left/right **108 tw**) zamiast `4px 8px`; fallbacki `tcMar` → 0/108.
- Writer: `tblW` domyślnie **auto** (`{Width="0",Type=Auto}`) zamiast `pct 5000`; tabelowy `tblCellMar` → **0/108/0/108** (Word default) zamiast 40/80.
- Writer `AddPageSettings`: marginesy zapisywane **jak autorskie** (bez `Math.Max(...+720)`); `Header/Footer` distance **rekonstruowane** = `clamp(margin − bandHeight, 0, 720)` (odwrotność wyliczenia pasma przez reader).
### Changed — frontend (`wysiwyg-editor.ts`)
- `getContent()` (ścieżka ZAPISU) scala strony **czystą konkatenacją** (bez `<div class="page-break">`). Akapity są block-atomic → konkatenacja odtwarza treść; jawne page-breaki użytkownika przeżywają jako div w treści.
### Verified (realny plik orginał_GOOD.docx → reader→writer → AFTER)
| metryka | GOOD | BAD | AFTER |
|---|---|---|---|
| pgMar top / bottom | 567 / 851 | 1281 / 1231 | **567 / 850** |
| pgMar header / footer | 6 / 340 | 720 / 720 | **6 / 339** |
| tblW (6 tabel) | auto | pct 5000 | **auto** |
| suma tcMar góra+dół | 0 | 8160 tw | **0** |
| twarde page-breaki | 1 | 3 | **0** |
### Tests
- Backend **366 pass / 0 fail / 6 skip** (+`RoundTripLayoutFidelityTests` 4: marginesy nie-inflowane, header/footer distance, tcMar=0, tblW auto; +regeneracja 2 golden table snapshotów na `padding:0px 7px`). GUI **168** (+2 `getContent` bez page-break).
- Weryfikacja na realnym pliku przez tymczasowy test lokalny (usunięty, niecommitowany) → `zapisany_AFTER.docx` (artefakt do ręcznego porównania w Word).
### Notes / ograniczenia (patrz R-15..R-18)
- AFTER ma 0 twardych breaków (vs GOOD 1) — reader→writer nie przywraca pojedynczego oryginalnego page-breaka (osobny gap writera). Po naprawie marginesów/tabel Word i tak paginuje ~4 strony naturalnie.
- Tabela dzielona przez edytor między strony (`_splitTableForPagination`) zapisuje się jako 2 sąsiednie `<table>` (pre-existing) — do domknięcia (scalanie przy zapisie).
- Font treści Cambria→Calibri, minimalny `styles.xml`, brak `numbering.xml` — model `DocumentContent` stratny (R-16); nie wpływa na liczbę stron tego dokumentu (`numPr`=0), ale na fidelity.

## 2026-06-03 — Dokumentacja: nowy `DOCX_CONVERSION.md` (reference techniczny konwersji) + cross-linki
### Changed
- Nowy `.ai/DOCX_CONVERSION.md` — zweryfikowany w kodzie reference konwersji DOCX↔HTML: pipeline (2 diagramy Mermaid), model `DocumentContent` (diagram klas + ograniczenie R-10), `OoxmlUnits`, **macierz statusów** wszystkich obszarów (Implemented/Partially/Planned z kierunkiem R/W i ścieżkami w kodzie), roadmapa (R-10 multi-section, Etap 3 computed-style, domknięcia 5/6/7), harness testów, ograniczenia trwałe.
- `INDEX.md` + wiersz dla `DOCX_CONVERSION.md`. Cross-linki z `FIDELITY_REPORT.md` i `FEATURES.md`.
### Verified (przeciw kodowi, bez zmian kodu)
- Multi-section: reader/writer `Body.Elements<SectionProperties>().FirstOrDefault()` → tylko pierwsza sekcja (R-10 = Partially).
- Computed-style: `basedOn`+`rStyle`+theme+docDefaults(font/size) działają; brak centralnego resolvera/numbering/linked folding (Partially; pełny = Planned).
- Rotacja obrazów `a:xfrm/@rot`: brak w obu konwerterach (Planned). VML: `ConvertPictureToHtml` tylko obraz+rozmiar (Partially).
- `tcMar`: czytane → CSS padding (Implemented). `tblCellSpacing`: nieczytane (Planned).
- `w:tabs` na zapisie: emitowane **tylko w stylach Header/Footer** (`AddDocumentStyles`, pozycje 4536/9072 tw); per-akapit/custom nieprzenoszone (Partially).
### Notes
- Zadanie dokumentacyjne — zero zmian kodu; testy bez zmian (backend 358, GUI 166).

## 2026-06-03 — Refaktor DOCX→HTML Etap 8: audyt izolacji CSS dokumentu (bez zmian kodu)
### Changed
- Brak zmian kodu — audyt. Ustalono, że treść dokumentu jest już izolowana: style inline wygrywają z regułami klasowymi przez specyficzność CSS. Faktyczna liczba `!important` w `wysiwyg-editor.scss` to **6** (nie 65 — wcześniejsza liczba była artefaktem zbiorczego grepu wielu wzorców), z czego tylko 1 dotyczy treści (świadoma zamiana Calibri Light→Calibri w nagłówkach). `!important` w `styles.scss`/`document-editor.scss` dotyczą podświetleń UI i paddingów dialogów — nie nadpisują font/koloru/marginesów treści.
### Verified
- `FIDELITY_REPORT.md` skorygowany (§3 wysoka wierność: izolacja CSS; §5: Etap 8 = zweryfikowane bez zmian). Backend **358 pass**, GUI **166 pass** bez zmian (audyt nie ruszał kodu).
### Notes
- Premisa planu „65 !important nadpisuje treść" była błędna (miscount). Ryzykowna chirurgia SCSS niepotrzebna. Ewentualny dalszy krok: defensywny scoped reset + test computed-style (ograniczony w jsdom).

## 2026-06-03 — Refaktor DOCX→HTML Etap 7: tab-stopy nagłówka/stopki (układ lewo⇥środek⇥prawo)
### Changed
- `DocxToHtmlConverter`: akapit z tab-stopem Center lub Right/End (`ParagraphHasAlignmentTab`) renderowany jako `display:flex` (`align-items:baseline;width:100%`). Run zawierający wyłącznie `<w:tab/>` (przy aktywnym `_flexTabs`) → `<span style="flex:1 1 0;">\t</span>` (rosnący spacer, **bezpośrednie dziecko flexa**), więc segmenty rozkładają się na szerokości zamiast zlewać w stałe odstępy. Znak taba zachowany (round-trip nienaruszony). Mieszane runy bez zmian.
### Verified
- `Infrastructure.UnitTests` **80/80 pass** (+1 golden `TabStopLeftCenterRight`: flex + 2 spacery + L/C/R tekst). Pozostałe golden niezmienione. Build OK.
### Notes
- Środek nie zawsze idealnie wycentrowany (zależy od szerokości boków) — udokumentowane jako częściowa wierność (FIDELITY_REPORT §3). Writer nie emituje `w:tabs` → po zapisie tab-stopy znikają i układ wraca do tabów (bez regresji vs. stan sprzed). Pozostaje Etap 7: wiele sekcji (R-10).

## 2026-06-03 — Refaktor DOCX→HTML Etap 6: tekst alternatywny obrazów (alt ↔ wp:docPr/@descr)
### Changed
- Reader `DocxToHtmlConverter.ConvertDrawingToHtml`: odczyt `wp:docPr/@descr` (fallback `@title`) → `<img alt="...">` (tylko gdy niepuste → istniejące obrazy bez alt bez zmian). Escape przez `EscapeHtml`.
- Writer `HtmlToDocxConverter`: `<img alt>` → `wp:docPr/@descr` w obu ścieżkach (inline + anchor) przez helper `BuildImageDocProperties`; `HtmlEntity.DeEntitize` na alt (inaczej encje podwójnie się escape'ują przy ponownym odczycie).
### Verified
- `Infrastructure.UnitTests` **79/79 pass** (+3 `ImageAltTextRoundTripTests`: round-trip, znaki specjalne/escape, brak alt → brak atrybutu). Golden-snapshoty niezmienione. Build OK.
### Notes
- Obrazy już wcześniej round-tripowały floating/border/crop/rozmiar (EMU). Pozostaje w Etapie 6: rotacja (`a:xfrm/@rot`), skalowanie przy crop (clip-path nie skaluje), VML (legacy) border/crop/alt.

## 2026-06-03 — Refaktor DOCX→HTML Etap 4 (full-stack): rozmiar strony + orientacja — pełny round-trip
### Changed
- Backend wiring: `PageSize` przewleczony przez `SaveDocumentCommand`/`SignDocumentCommand`/`DownloadEditedDocumentCommand` (+ handlery → `_converter.Convert(... request.PageSize)`) i 3 kontrolery (`DocumentController.save/sign`, `DocumentStorageController.user-download`). DTO (`SaveDocumentRequest`/`SignDocumentRequest`) już niosły pole z Etapu 4 backend.
- Frontend: `document.model.ts` — interfejs `PageSize` + pole na `DocumentContent`/`SaveDocumentRequest`. `document-editor.ts` — sygnał `documentPageSize` ustawiany przy wczytaniu (`content.pageSize` → sygnał + `pageSettings.orientation`) i dołączany w `buildSaveRequest`. Round-trip pełny: DOCX → reader → `pageSize` → front → zapis → writer `w:pgSz`.
### Verified
- Backend **358 pass / 0 fail** (Domain 57, Application 192, Api 33, Infrastructure 76). GUI **166 pass** (+2 testy `PageSize round-trip` w `document-editor.spec`). Build OK.
### Notes
- R-14 **zamknięte** (pełny full-stack). Brak `pageSize` z frontu → backend fallback A4 (bez regresji).

## 2026-06-03 — Refaktor DOCX→HTML Etap 4 (backend): rozmiar strony + orientacja (round-trip)
### Changed
- Domena: nowy model `PageSize { WidthCm, HeightCm, Orientation }`; dodany do `DocumentContent`, `SaveDocumentRequest`, `SignDocumentRequest` (additive, backward-compatible).
- Reader `DocxToHtmlConverter`: `ExtractPageSize` z modelu `PageSettings` (Etap 2) → `DocumentContent.PageSize` (cm + portrait/landscape).
- Writer `HtmlToDocxConverter`: `IHtmlToDocxConverter.Convert` + `Convert` przyjmują opcjonalny `PageSize? pageSize`; `BuildPageSize` emituje `w:pgSz` (+`Orient` dla landscape) zamiast zahardkodowanego A4. Brak `pageSize` → fallback A4 portrait (bez regresji). Alias `OoxmlPageSize` rozwiązuje kolizję nazw `PageSize` (Domain vs OpenXml).
### Verified
- Cała solucja **358 pass / 0 fail / 6 skip** (Domain 57, Application 192, Api 33, Infrastructure 76 [+3 `PageSizeRoundTripTests`: landscape A4, A5 portrait, null→A4]; Integration 6 skip = wymaga Postgresa). Build OK.
- Zmiana interfejsu (opcjonalny param) wymusiła aktualizację mocków Moq w `DownloadEditedDocumentCommandHandlerTests` (CS0854: expression tree + optional arg) — dodany `It.IsAny<PageSize?>()`.
### Notes
- **Pozostaje (Etap 4 full-stack):** przewlec `PageSize` przez komendy/DTO (Save/Sign/Download) i front (`document.model.ts` + editor/ruler/scss + Vitest), żeby zapis faktycznie round-tripował rozmiar (teraz handlery nie przekazują `pageSize` → writer defaultuje A4; reader już zwraca `PageSize` w `DocumentContent` — front może konsumować). Patrz R-14.

## 2026-06-03 — Refaktor DOCX→HTML Etap 3: rozwiązywanie nazwanych stylów znakowych (w:rStyle)
### Changed
- `DocxToHtmlConverter.ConvertRunToHtml`: run z referencją do stylu znakowego (`w:rStyle`) dostaje teraz CSS tego stylu (z dziedziczeniem `basedOn`, z `_styles[rStyleId]`) **pod** formatowaniem bezpośrednim (direct wygrywa na konflikcie). Wcześniej runy formatowane wyłącznie przez styl znakowy (Hyperlink/Strong/Emphasis/własny) renderowały się **bez formatowania** — realna luka wierności.
### Verified
- `Infrastructure.UnitTests` **73/73 pass** (+1 golden `CharacterStyleRun`: bold+kolor+rozmiar ze stylu „Akcent" przez `rStyle`). Istniejące snapshoty **niezmienione** (brak `rStyle` w nich → zero dryfu). Build solucji OK.
### Notes
- Zakres celowy (slice): pełny computed-style (folding docDefaults do każdego elementu, style numerowania, linked styles, pełny theme) — kolejne kroki Etapu 3. Edge: jawne `bold=false` na runie nie wyłącza pogrubienia ze stylu znakowego (brak `font-weight:normal` z direct) — rzadkie, do domknięcia z computed-style.

## 2026-06-03 — Refaktor DOCX→HTML Etap 2: jawny model pośredni (read-side, strangler) — sekcja/strona
### Changed
- Nowy namespace `Infrastructure/DocxModel/`: `PageSettings` (geometria strony/sekcji jako surowe twipsy + page size + orientacja) i `SectionPropertiesReader.ReadPageSettings(sectPr)` (czysty parser, bez defaultów/konwersji).
- `DocxToHtmlConverter`: `ExtractPageMargins` + wysokość pasma nagłówka/stopki liczone teraz z modelu (`SectionPropertiesReader` + nowy helper `ComputeBandHeightCm`), zamiast bezpośredniego grzebania w `PageMargin`. Zachowanie **1:1** (te same wzory/zaokrąglenia). Page size/orientacja parsowane, ale jeszcze nieujawniane w HTML (fundament Etap 4).
### Verified
- `Infrastructure.UnitTests` **72/72 pass** (+5 `SectionPropertiesReaderTests` na modelu: size/margins/distances, landscape, brak pgMar, null sectPr, ujemny top margin). Golden-snapshoty **niezmienione** → potwierdzenie braku zmiany zachowania. Build solucji OK.
### Notes
- Strangler: stara (string) i nowa (model) ścieżka współistnieją; kolejne obszary migrują w Etap 3–7. ADR-0009.
- Kontrakt `DocumentContent`/API bez zmian.

## 2026-06-03 — Refaktor DOCX→HTML Etap 5: tabele (colgroup + fixed layout + fix vMerge/gridSpan)
### Changed
- `DocxToHtmlConverter.ConvertTableToHtml`: emisja `<colgroup><col style="width:..">` z `tblGrid` (autorytatywne szerokości kolumn) + `table-layout:fixed` gdy `tblLayout=fixed` lub tabela ma jawną szerokość (Dxa/Pct). Gdy fixed+brak szerokości tabeli → szerokość = suma kolumn z grid. Naprawia „tabela nie wygląda jak w Wordzie" (przeglądarka ignorowała geometrię kolumn). Nowe helpery `ReadTableGridColumnsPx`/`BuildColgroupHtml`.
- `DocxToHtmlConverter` (merge): naprawiony `rowspan` przy pionowym scaleniu sąsiadującym z `gridSpan` — `CountRowSpan` dopasowuje komórki wg **pozycji kolumny w gridzie** (z uwzględnieniem `gridSpan`), nie wg indeksu komórki (stary bug gubił rowspan). Skip komórki kontynuacji obsługuje teraz też jawny `vMerge val="continue"` (wcześniej tylko pominięty val). Nowe helpery `GetGridSpan`/`GetCellStartColumn`/`FindCellAtColumn`.
### Verified
- `Infrastructure.UnitTests` **67/67 pass** (+1 nowy `MergedCellsTable` golden; `simple-table` baseline zregenerowany — dodane colgroup/fixed). Snapshoty stabilne.
- Round-trip bezpieczny: writer (`HtmlToDocxConverter`) czyta wiersze przez `.//tr` i odtwarza własny `TableGrid` z liczby komórek → ignoruje `<colgroup>`/`<col>` (brak regresji zapisu).
### Notes
- Strona zapisu wciąż wymusza `TableLayout Autofit` — fidelity zapisu (fixed) odłożone (round-trip risk), do Etapu 4/5 cd.
- Pozostaje w Etap 5: `tcMar`/`tblCellMar` pełne, shading dziedziczony z `tblPr`, cellSpacing, border conflict resolution.

## 2026-06-03 — Refaktor DOCX→HTML Etap 0+1: centralne jednostki `OoxmlUnits` + harness regresji
### Changed
- Nowy `D2ViewerEditor.Infrastructure/Conversion/OoxmlUnits.cs` — jedno źródło prawdy dla konwersji OOXML (twips/EMU/half-points/cm/px) z nazwanymi stałymi (`TwipsPerInch`, `EmuPerInch`, `EmuPerPixel=9525`, `TwipsPerPoint=20`, `HalfPointsPerPoint=2`, `DefaultDpi=96`, `CmPerInch=2.54`).
- `DocxToHtmlConverter` i `HtmlToDocxConverter`: wszystkie rozproszone stałe konwersji (`567`, `1440`, `914400`, `9525`, `/20`, `*0.75`, `/2`, `*2.54`) zastąpione wywołaniami `OoxmlUnits`; usunięte zduplikowane prywatne helpery (`TwipsToPx`/`EmuToPx`/`PxToTwips`/`TwipsToCm`/`cmToTwips`). Lokalny rounding/truncation zachowany → zachowanie liczbowe niezmienione. cm liczone dokładnym `1440/2.54` (zastępuje przybliżenie `567`).
- Nowy harness regresji `Infrastructure.UnitTests/Golden/`: `GoldenDocuments` (deterministyczne DOCX w pamięci), `HtmlSnapshot` (normalizacja base64/`data-image-id`/daty + approve-on-missing), `GoldenSnapshotTests` (5 dokumentów: tekst, runy, spacing/indent, tabela, nagłówek+stopka z logo). Baseline w `Golden/__snapshots__/*.approved.html`.
### Verified
- `dotnet build D2ViewerEditor.sln` OK. `Infrastructure.UnitTests` **66/66 pass** (54 baseline + 7 `OoxmlUnitsTests` + 5 Golden). Snapshoty stabilne przy 2× uruchomieniu.
- Asercje punktowe potwierdzają wzory: sz=32→16pt, 240tw→12pt, 720tw→48px, 480tw→32px, 3000tw→200px, 2000tw→133px, EMU 1270000/317500 round-trip.
### Notes
- Kontrakt `DocumentContent`/`IDocxToHtmlConverter`/API bez zmian — front nietknięty. Round-trip (Header/Footer, ImageFloating/BorderCrop) zielony.
- Plan pełny: `~/.claude/plans/fancy-spinning-wirth.md` (Etap 0–9, strangler). Następny: Etap 2 (jawny model pośredni read-side) — duża zmiana, wymaga decyzji.
- Snapshoty lokują *obecne* zachowanie (regression-guard), nie poprawność wobec Worda; wierność podnoszona w Etap 3–8.

## 2026-05-28 — Import nagłówka/stopki: wybór wg referencji sekcji (default + first-page)
### Changed
- Backend `DocxToHtmlConverter`: `ExtractHeader`/`ExtractFooter` rozwiązują part przez `sectPr`/`HeaderReference`/`FooterReference` typu **Default** zamiast `HeaderParts.FirstOrDefault()` (kolejność partów była niezdefiniowana → mógł trafić pusty even/first). Dodano helpery `ResolveHeaderPart`/`ResolveFooterPart`/`HasTitlePage` oraz `ConvertHeaderPartToHtml`/`ConvertFooterPartToHtml`. Pierwsza strona (`titlePg`) → `DifferentFirstPage`+`FirstPageHtml` (model domeny już miał te pola). Fallback do `FirstOrDefault` gdy sekcja nie ma referencji.
- Frontend `document-editor.ts`: oba miejsca ładowania (`loadFromStorage` + ścieżka `openDocument`) spread'ują cały obiekt `content.header`/`content.footer`, zamiast rekonstruować tylko `{html,height}` → wariant first-page/odd-even nie jest gubiony na granicy TS.
### Verified
- Backend build OK; `dotnet test` filtr Header/Footer + SectionReference: **10/10 pass** (5 nowych w `DocxToHtmlConverterSectionReferenceTests`, 5 istniejących).
- Frontend `tsc --noEmit` OK.
- Diagnoza na realnym `stupki.docx`: 3 nagłówki (even/default/first) + 3 stopki, brak `titlePg`/`evenAndOddHeaders`; default=header2 (logo + „Qutalo Qutalo … Obciągalski"), default footer=footer2 (8pt, #808080). Stary kod mógł renderować pusty even part.
### Notes
- R-10 częściowo zamknięte (default + first-page). Pozostaje: even/odd (brak pól w modelu backendu) i wiele sekcji.
- `stupki.docx` NIE dodano jako fixture (treść wulgarna) — testy używają syntetycznego DOCX o tej samej strukturze.

## 2026-05-28 — Import nagłówka/stopki: wariant even/odd + uzupełnienie .ai
### Changed
- Domena `HeaderFooterContent`: dodane `DifferentOddEven` + `EvenHtml` (obok `DifferentFirstPage`/`FirstPageHtml`).
- `DocxToHtmlConverter`: helper `HasEvenAndOddHeaders` (czyta `settings.xml/evenAndOddHeaders`); `ExtractHeader/ExtractFooter` czytają referencję **Even** gdy włączone → `EvenHtml`/`DifferentOddEven`. „Default" = strona nieparzysta.
- `.ai/FEATURES.md`: nowa sekcja „Feature: Nagłówek i stopka (import + edycja)" + 2 wiersze w tabeli statusu (instrukcja dla agenta: NIE używać `FirstOrDefault` jako głównej ścieżki, spread'ować cały obiekt header/footer).
### Verified
- Backend build OK; testy Header/Footer + SectionReference: **13/13 pass** (3 nowe even/odd).
### Notes
- **Round-trip save** nadal zapisuje tylko default — first/even na zapisie nie są serializowane (R-10 partial, R-11). Import je odczytuje.

## 2026-05-29 — Word-like obraz: 7 trybów zawijania + obramowanie + przycinanie (a:ln + a:srcRect round-trip)
### Changed
- **Panel obrazu** rozszerzony do pełnego Word-like UX:
  - **Zawijaj tekst** (7 trybów wg menu Worda): Równo&nbsp;z&nbsp;tekstem / Ramka / Góra&nbsp;i&nbsp;dół / Za&nbsp;tekstem / Przed&nbsp;tekstem (5 działa) + Przylegle / Na&nbsp;wskroś (2 disabled — wymagają reflow tekstu wokół kształtu, niedostępne natywnie w HTML/CSS). Grid 2 kolumny.
  - **Obramowanie**: checkbox włącz/wyłącz, color picker, grubość (1–20 px), styl (ciągła / przerywana / kropkowana).
  - **Przytnij**: 4 pola % (lewo/prawo/góra/dół, 0–95%) + przycisk „Resetuj przycięcie". Każde pole emituje `cropChange` z aktualnym vector.
  - Wszystkie kontrolki są stateless — emity przekazują się do public methods na `wysiwyg-editor`.
- **`wysiwyg-editor` — nowe public methods**:
  - `setSelectedImagePositionMode` rozszerzony o `'square'` i `'topBottom'` (oprócz `'inline'`, `'front'`, `'behind'`). Square: `float: left; margin: 0 12px 8px 0;`. TopBottom: `display: block; clear: both; margin: 8px auto;`. Inline / front / behind bez zmian.
  - `setSelectedImageBorder({enabled, color, widthPx, style})` — aplikuje inline CSS `border: Npx style #hex` na `<img>` + persist na `data-border-width/color/style`. Walidacja hex.
  - `setSelectedImageCrop({left, right, top, bottom})` — `clip-path: inset(t% r% b% l%)` na `<img>` + persist na `data-crop-l/r/t/b` (clamp 0–95).
  - `resetSelectedImageCrop()` — zero-out crop.
  - `emitImageSelectionState` rozszerzony o pełen snapshot `border` + `crop` z atrybutów (default-safe gdy brak).
- **`wrapExistingImages`** przywraca po loadzie:
  - tryb `square` / `topBottom` z `data-pos-mode` (`float: left` / `display: block + clear: both`),
  - border z `data-border-*` jako inline CSS,
  - crop z `data-crop-*` jako `clip-path`.
- **Backend `HtmlToDocxConverter.BuildImageDrawing`** (round-trip save):
  - Border: `data-border-width/color/style` na `<img>` → `a:ln` w `pic:spPr` z `a:solidFill srgbClr` + `a:prstDash` (mapowanie `solid`→`Solid`, `dashed`→`Dash`, `dotted`→`Dot`). Width: px×9525 EMU. Walidacja hex — niepoprawny kolor → border pomijany (silent fail-safe).
  - Crop: `data-crop-l/r/t/b` (%) → `a:srcRect` na `pic:blipFill` z l/t/r/b w 1/1000 procenta.
- **Backend `DocxToHtmlConverter.ConvertDrawingToHtml`** (round-trip import): czyta `a:ln` (width, srgb, dash) i `a:srcRect` (l/t/r/b) → emituje `data-border-*` i `data-crop-*` na `<img>`. Plus istniejące `data-pos-mode`/`data-x-emu`/`data-y-emu`/`data-width-emu`/`data-height-emu`.
### Verified
- `dotnet build` (Infrastructure) OK; nowe `ImageBorderCropRoundTripTests` (**7/7 pass**: border solid/dashed/dotted, bad-color fail-safe, crop 4-side round-trip, no-attrs-after-plain-image). `ImageFloatingRoundTripTests` (4/4 pass).
- `tsc --noEmit` OK; `npm test`: **164/164 pass** (17 plików; +4 nowe testy panelu: 7 wrap modes z disabled, emisja 5 supported, border toggle, crop clamp 95, resetCrop).
### Notes / ograniczenia (zaktualizowane)
- **Przylegle / Na wskroś** — HTML/CSS nie ma natywnego mechanizmu reflow tekstu wokół arbitralnego kształtu. Pozostają jako disabled w panelu (UI mówi prawdę o tym, co działa).
- **Drag w obrębie strony** dalej działa tylko dla `front`/`behind` (absolute). Dla `square`/`topBottom` (float/block) drag pozostaje inline-drop-into-DOM (rule 14 — sensowne dla float).
- **Crop**: clamp 0–95% per side (pozostawia widoczny obraz). Mapowanie do OOXML 1/1000 procenta (standardowa skala `a:srcRect`).
- **Border**: tylko jednolity kolor + 3 style. Patterny / gradienty / 3D efekty — out of scope.

## 2026-05-29 — Word-like floating: „Przed tekstem" / „Za tekstem" + drag w obrębie strony + round-trip wp:anchor
### Changed
- **Tryby pozycji obrazu** (panel boczny): zastąpiono info-only sekcję 3 prawdziwymi przyciskami: **„W tekście"** (inline), **„Przed"** (float over text), **„Za"** (float behind text). Aktywny tryb podświetlony. `ImageSelectionState` rozszerzony o `positionMode: 'inline' | 'front' | 'behind'`. Nowy `@Output() positionModeChange` w panelu.
- **`wysiwyg-editor` — floating end-to-end**:
  - Public `setSelectedImagePositionMode(mode)`: dla `inline` usuwa wszystkie ślady floating (data-*, position, z-index); dla `front`/`behind` przycina wrapper przy bieżącej pozycji (bounding rect względem `.page`/edytora), ustawia `position: absolute` + `left/top` + `z-index` (10 vs −1) + zapisuje `data-x-px/y-px` na wrapperze i `data-x-emu/y-emu` na `<img>` (do round-tripu).
  - **Drag-aware mousedown**: gdy wrapper jest w trybie floating, kliknięcie + przesunięcie aktualizuje `left/top` (oddzielna ścieżka od inline-drop-into-DOM); na mouse-up zapisuje pozycję w atrybutach + `onContentChange()` (undo/autozapis).
  - `wrapExistingImages` przywraca floating po imporcie: czyta `data-pos-mode` z `<img>` i woła `applyFloatingPosition` na wrapperze (dzięki czemu obraz wczytany z DOCX z `wp:anchor` natychmiast renderuje się jako absolutny w odpowiedniej warstwie).
- **CSS** (`wysiwyg-editor.scss`):
  - `.editor-image-wrapper[data-pos-mode="front"]` → `position: absolute; z-index: 10;`
  - `.editor-image-wrapper[data-pos-mode="behind"]` → `position: absolute; z-index: -1;`
  - Edytory (`.editor-content`/`.header-editor-content`/`.footer-editor-content`/displaye) dostają `position: relative; z-index: 0` → kontekst stacking dla warstwy „za". `.page` już miało `position: relative`.
- **Backend `HtmlToDocxConverter.BuildImageDrawing`**: gdy `<img>` ma `data-pos-mode="front"|"behind"`, emituje `wp:anchor` z `wp:positionH`/`wp:positionV` (relativeFrom=page, offsety z `data-x-emu/y-emu`), `wp:wrapNone`, `behindDoc` zgodnie z trybem, `simplePos=false`, `allowOverlap=true`. Inline pozostaje domyślną ścieżką (`wp:inline`) — brak regresji dla istniejących dokumentów.
- **Backend `DocxToHtmlConverter.ConvertDrawingToHtml`**: wykrywa `wp:anchor`, czyta `BehindDoc` + `PositionOffset` z H/V, emituje na `<img>` `data-pos-mode="front"|"behind"` + `data-x-emu`/`data-y-emu`. Plus istniejące `data-width-emu`/`data-height-emu` (rozmiar dalej round-tripuje).
### Verified
- `dotnet build` (Infrastructure) OK; **4 nowe testy `ImageFloatingRoundTripTests` pass** (front + behind round-trip pozycji i trybu, inline-no-floating-attrs guard, size obok pozycji).
- `tsc --noEmit` OK; `npm test`: **162/162 pass** (+2 nowe testy panelu: active state aktualnego trybu + emisja `positionModeChange` dla 3 przycisków).
### Notes / pozostałe ograniczenia
- **Wrap modes wciąż out of scope**: `square`/`tight`/`through`/`topAndBottom` (z reflow tekstu) odłożone — wymagają nietrywialnego mechanizmu HTML/CSS (text-wrap niedostępny w przeglądarce). Panel jasno mówi: „Square/tight wrap — w przygotowaniu". Front/behind nie wymagają wrap'u (Word też używa `wrapNone`).
- **Drag floating** działa w obrębie kontenera (najbliższy `.page` / `editor-content`). Przeniesienie obrazu między stronami / sekcjami z anchor-rekalkulacją — roadmap.
- **Backend test image bytes**: round-trip używa 1×1 PNG. Realne dokumenty z floating image z Worda — manual verify.

## 2026-05-29 — Word-like pozycjonowanie obrazów (MVP) — selection-driven side panel
### Changed
- **Nowy komponent `d2-image-properties-panel`** (`components/image-properties-panel/`: TS+HTML+SCSS+spec) — czwarty docked panel w istniejącym slocie (Wyszukiwanie / Tabele / Nagłówek+Stopka / **Obraz**). Stateless: każda zmiana w polu emituje output do parenta, parent woła publiczne metody na `wysiwyg-editor`. Sekcje: Rozmiar (szer./wys. + zachowaj proporcje + „Przywróć proporcje"), Wyrównanie (L/C/R/—; aplikowane do akapitu nadrzędnego), Opływanie (info-only — floating w przygotowaniu), Usuń obraz.
- **`wysiwyg-editor` outputs i public API**:
  - Nowy `@Output() imageSelectionChange = EventEmitter<{ widthPx, heightPx, aspectRatio, alignment } | null>` — emitowany przy `selectImageWrapper`, `clearSelectedImage`, po resize-end i po drop-end.
  - Nowe metody publiczne: `setSelectedImageWidth(px, lockAspect=true)`, `setSelectedImageHeight(px, lockAspect=true)`, `setSelectedImageAlignment('left'|'center'|'right'|null)`, `resetSelectedImageAspect()`, `removeSelectedImage()`. Wszystkie reuse istniejącego pipeline: aktualizacja inline `style.width/height` + `data-width-emu`/`data-height-emu` (parametry konsumowane przez `HtmlToDocxConverter` przy eksporcie), po czym `onContentChange()` → debounced undo-stack + autozapis.
- **`document-editor`**: signal `selectedImage` + `imageLockAspect` (default true) + computed `showImagePanel = selectedImage() !== null && !showFindReplace() && !showTablePanel()`. **Panel obrazu ma pierwszeństwo nad panelem nagłówka/stopki** (HF wpuszcza — `showHeaderFooterPanel` rozszerzony o `&& !showImagePanel()`), żeby użytkownik mógł formatować logo w edytowanym nagłówku. Ruler-h-spacer wpuszcza dodatkowy panel.
- **Roundtrip**: rozmiar już round-tripuje przez istniejące `data-*-emu` na `<img>` (czytane w `DocxToHtmlConverter`/`HtmlToDocxConverter`). Pozycja inline (kolejność w DOM) round-tripuje przez sam HTML. Alignment przez `text-align` na akapicie (zachowane w obie strony — patrz `DocxToHtmlConverterFidelityTests`).
### Verified
- `tsc --noEmit` OK; `npm test`: **160/160 pass** (17 plików; +9 nowych `image-properties-panel.spec.ts` + 1 koordynacja w `document-editor.spec.ts`).
### Notes / ograniczenia MVP
- **Floating positioning (wp:anchor) odłożone do następnej iteracji.** Word-like wrap modes (square / tight / through / top-and-bottom / behind / in-front-of) nie są obsługiwane w tym MVP. Panel pokazuje sekcję „Opływanie tekstu" jako info-only („W tekście (inline). Floating w przygotowaniu.") — UI nie udaje funkcji, której nie ma. Roadmap: rozszerzenie `DocxToHtmlConverter`/`HtmlToDocxConverter` o `wp:anchor`/`wp:positionH`/`wp:positionV`/`wp:wrap*` + signal modelu w editor (`data-pos-mode="floating"`, `data-x-emu`, `data-y-emu`, `data-wrap`).
- **Drag** (przenoszenie obrazu w obrębie strony) i **resize** (uchwyty rogu + boków, corner = aspect-locked) ZAWSZE istniały — nie zmieniam mechaniki, tylko emituję snapshot po każdym z tych zdarzeń, więc panel jest spójny z stanem DOM.
- **Klawiatura**: Delete/Backspace usuwa zaznaczony obraz (istniejące), Escape deselect (istniejące). Strzałkowe przesuwanie do roadmapy.

## 2026-05-29 — Reguła `userDownload`: kontrola pobierania edytowanego pliku
### Changed
- **Reguła domenowa**: nowe opcjonalne pole `userDownload` w metadanych dokumentu (`documents.metadata` JSON). Default `false`; brak / null / non-true ⇒ blokada. Niezależne od `returnUrl`.
- **Application: wspólny parser metadanych** `Features/Documents/Common/ExternalDocumentMetadata` (record + tolerant `Parse(string?)` + `Serialize()` + `IsUserDownloadAllowed`). Zastąpił 4 lokalne kopie `sealed record ExternalMetadata` w handlerach (`GetDocumentMetadata`, `GetDocumentStatus`, `UpdateCallbackUrl`, `FinishAndSendDocument`-touchpoints) — single source of truth dla shape JSON.
- **External API** (`POST /api/v1/document`): `CreateDocumentRequest` dostaje opcjonalne pole `UserDownload: bool?`. Mapowane do `metadata.userDownload`. Przesłanie `null`/`false`/braku → pole w JSON `null` (parser interpretuje jako blokadę).
- **Local upload** (`UploadDocumentCommandHandler`): backend **sam** ustawia `userDownload=true` w nowo tworzonym `Document.Metadata` (rule 12 — anti-tamper; klient nie może wymusić ani zablokować). Bo lokalny upload nie ma `returnUrl` — pobranie jest jedynym sposobem odzyskania edytowanego pliku.
- **Nowy gated endpoint**: `POST /api/documentstorage/{masterId}/user-download` (`DocumentStorageController.DownloadEditedDocument`) — konwertuje aktualny stan edytora HTML → DOCX **tylko** gdy `userDownload == true`. Sentinel error `USER_DOWNLOAD_FORBIDDEN:` mapowany w controllerze na **HTTP 403**. Dawny stateless `POST /api/document/save` pozostaje bez zmian (używany przez inne flow); GUI nie korzysta już z niego dla „Pobierz dokument".
- **DTO statusu/metadanych eksponują flagę** (`DocumentMetadataDto.UserDownload`, `DocumentStatusDto.UserDownload`) — system zewnętrzny widzi, czy użytkownik ma prawo do pobrania.
- **Frontend**: `DocumentMetadataDto.userDownload: boolean` + sygnał `userDownload` + `canUserDownload` computed w `document-editor`. Pozycja menu „Pobierz dokument" gated `@if (canUserDownload())`. `document.service.downloadDocument` zastąpiony przez `documentStorageService.downloadEditedDocument(masterId, request)` → POST na nowy gated endpoint; obsługa 403 z czytelnym komunikatem + synchronizacja lokalnego sygnału.
### Verified
- `dotnet build` obu solucji OK (0 błędów); pełna solucja backendu: **325/325 pass** (Application 192 +20 nowych, Domain 57, Api 33, Infrastructure 43; Integration 6 skipped — DB-bound). 
- `tsc --noEmit` OK; `npm test`: **150/150 pass** (+4 nowych w `document-editor.spec.ts`).
### Notes / ograniczenia
- **Raw download endpoints** (`GET .../{masterId}/download`, `GET .../versions/{vid}/download`) NIE są gated — używane przez edytor do *ładowania* bajtów (nie do user-downloadu). Ten threat-vector (zalogowany użytkownik z bezpośrednim wywołaniem URL) jest poza zakresem tego zadania, do udokumentowania jako follow-up R-NEW. W praktyce: te endpointy nie zwracają „edytowanego" pliku — tylko ostatnio zapisany; do faktycznego pobrania edycji niezbędne jest `POST .../user-download`.
- Stary `documentService.downloadDocument` pozostaje w bazie kodu (martwy dla user-download, używany w innych miejscach jak sign flow) — refactor poza zakresem.

## 2026-05-29 — External API: 3 nowe endpointy (callback URL / unlock / status)
### Changed
- **Domena.** `Document` dostaje dwie nowe metody (mirror `MarkEditing/Sending/...`):
  - `MarkSaved()` — odwraca `Editing → Saved` (przeznaczone wyłącznie dla unlock; stany wysyłki forbidden po stronie handlera).
  - `UpdateMetadata(string?)` — ustawia metadane (JSON); walidacja struktury w warstwie aplikacji.
- **Application.** Trzy nowe jednostki MediatR w `Features/Documents/`:
  - `Commands/UpdateCallbackUrl/{Command,Handler}` — walidacja URL przez `DocumentDelivery.IsValidRecipientUrl` (ten sam, co worker), limit 2048 znaków, zachowuje `classification` w JSON, blokuje stany `Sending`/`Sent`/`DeliveryFailed`.
  - `Commands/UnlockDocument/{Command,Handler}` — `Editing → Saved`; idempotentne dla `Saved` (`Result { Changed=false }`), blokowane dla stanów wysyłki. Logowane (poziom Info, z opcjonalnym `Reason`).
  - `Queries/GetDocumentStatus/{Query,Handler}` — `DocumentStatusDto { masterId, status, isLocked, hasCallbackUrl, activeVersionId?, activeVersionNumber?, activeVersionModifiedAt?, latestDelivery? }`. Pełny `callbackUrl` celowo nie surfacowany (może zawierać token).
- **External API.** `D2ServicesViewerEditor.Api/Controllers/DocumentController` rozszerzony o:
  - `PUT /api/v1/document/{masterId:guid}/callback-url` → 204/400/404/409. URL nigdy nie trafia do logów.
  - `POST /api/v1/document/{masterId:guid}/unlock` → 200 `UnlockDocumentResult { masterId, changed }` / 404 / 409.
  - `GET /api/v1/document/{masterId:guid}/status` → 200 `DocumentStatusDto` / 404.
- **Decyzja kontraktowa** (`.ai/API_CONTRACTS.md`): wszystkie trzy operują na **`masterGuid`** (returnUrl w master metadata; status = atrybut master; brak osobnego user-locka — `Editing` to ten lock). Tabela rozstrzygnięć w API_CONTRACTS.
### Verified
- `dotnet build` obu solucji OK (0 błędów).
- Pełna solucja backendu: Domain 57 + Application 172 (+25 nowych) + Api 33 + Infrastructure 43 = **305/305 pass**; Integration 6 skipped (DB-bound).
### Notes
- **Autoryzacja**: nowe endpointy używają tego samego mechanizmu co istniejący `POST /api/v1/document` (obecnie brak `[Authorize]` w `D2ServicesViewerEditor` — match z istniejącą konwencją; ewentualne wprowadzenie auth schematu pokryje WSZYSTKIE endpointy zewnętrzne jednym przepisem).
- **SSRF**: walidacja URL ogranicza protokoły do http(s) i kapuje długość; brak allowlisty hostów (świadome — match z obecnym podejściem ingestu, ślad w RISKS R-07).
- **Concurrency**: `UpdateCallbackUrl` / `UnlockDocument` modyfikują agregat `Document` przez repo + `SaveChangesAsync` (EF Core change tracker + transakcja na poziomie SaveChanges); pełna optymistyczna kontrola wersji nie istnieje (brak `RowVersion`) — zgodne z resztą domeny.

## 2026-05-29 — Menu „Pomoc" + status autozapisu w stopce
### Changed
- **Nowa zakładka menu „Pomoc"** w `document-editor.html` (ostatnia po „Widok"), z jedną pozycją dropdown **„Zgłoś"** → wywołuje istniejące `openReportEmail()` (bez duplikacji logiki — ta sama metoda co dawny przycisk). Sygnał `showHelpMenu = signal(false)` + `toggleHelpMenu()` zgodne z konwencją innych menu (zamyka pozostałe przez `closeAllMenus()`, którego rozszerzono o nowy sygnał). `openReportEmail()` woła teraz `closeAllMenus()` na początku — pozycja w dropdownie zamyka go po kliknięciu, spójnie z resztą menu.
- **Stary przycisk „Zgłoś" usunięty** z `.header-right`. Toolbar w prawym górnym rogu jest teraz czystszy: badge klasyfikacji, switch „Autozapis", opcjonalny „Zapisz"/„Zakończ".
- **Data/czas ostatniego autozapisu przeniesione do stopki.** Dawniej `.autosave-status` (`min-width:70px`) obok switcha pokazywał „Zapisywanie…" / „Zapisano o HH:MM:SS" / „Błąd auto-zapisu" — usunięte z toolbara (przy switchu zostaje tylko switch + label „Autozapis", zgodne z task #2). Nowy `.status-autosave` w `.editor-footer > .status-right`: „Autozapis: HH:MM:SS" / „Autozapis: zapisywanie…" / „Autozapis: błąd zapisu" / fallback „Autozapis: włączony"; ukryty gdy autozapis off lub brak `versionId`. Bez nowych źródeł danych — te same sygnały (`autoSaveEnabled`/`autoSaveStatus`/`lastAutoSaveAt`).
- SCSS: nowy `.status-autosave` (`#555`, `is-saving #1a73e8`, `is-error #cc0000 bold`) — subtelny, dopasowany do reszty status-baru.
- `document-editor.spec.ts` +4 testy: `toggleHelpMenu` otwiera/zamyka i koliduje z `showViewMenu`, `closeAllMenus` zamyka `showHelpMenu`, `openReportEmail()` zamyka dropdown.
### Verified
- `tsc --noEmit` OK; `npm test` → **146/146 pass**.
### Notes
- Logika autozapisu i akcji `openReportEmail` bez zmian (rule 6/7) — przeniesienie czysto prezentacyjne.

## 2026-05-29 — Panel nagłówka/stopki: porządek wizualny + NG8107
### Changed
- **Usunięty primary CTA „Zamknij nagłówek i stopkę"** z `d2-header-footer-panel`. Był wizualnie redundantny z X w nagłówku panelu i z ESC (oba wciąż delegują do `editor.stopEditingHeaderFooter()`). Dock jest spójny z `Wyszukiwanie` i panelem tabeli — X jako jedyne zamknięcie z panelu. Spec testu zaktualizowany: zamiast Primary asercja na X.
- **NG8107 (Angular template lint)**: bindingi panelu w `document-editor.html` używały `editor?.method()`, ale `editor` jest `@ViewChild` z `editor!:` (non-null) i panel renderuje się tylko po `editingSection() !== 'body'`, które emituje sam edytor — więc operand zawsze ≠ null. Zmienione 9 wystąpień `editor?.` → `editor.`; semantyka bez zmian, ng serve nie emituje już warningów NG8107.
### Verified
- `tsc --noEmit` OK; `npm test` → **142/142 pass**; `ng serve` bez NG8107.
### Notes
- Pełna lista wyjść z trybu edycji: X w nagłówku panelu, ESC, klik w body editor.

## 2026-05-29 — UX poprawki: autozapis, prowadnice, obrazy nagłówka, panel boczny
### Changed
- **Autozapis (stabilny layout).** `document-editor.html`: `.autosave-status` jest teraz renderowany ZAWSZE (gdy `documentVersionId()` istnieje), a stan idle/disabled emituje pusty content (klasa `is-empty`). `min-width:70px` rezerwuje przestrzeń, więc toggle nie powoduje reflow toolbara — przyciski Zgłoś/Zapisz/Zakończ nie skaczą.
- **Prowadnice marginesów domyślnie ukryte.** `document-editor.ts` `showMarginGuides = signal(false)`, `wysiwyg-editor.ts` `@Input() showMarginGuides = false`. Menu/dialog Widok dalej działają (toggle bez zmian).
- **Obraz nagłówka — brak reskalowania w trybie edycji.** `wysiwyg-editor.ts` `wrapExistingImages` NIE usuwa już inline `width`/`height` z `<img>`; CSS `.editor-image-wrapper img { width: 100%; height: auto }` zmieniony na samo `max-width: 100%` (rule 14 — bez „pozornego" reskalu). `display: inline-block` wrappera kurczy się do obrazu; obraz w edycji ma ten sam rozmiar co w podglądzie.
- **Pasek nagłówka/stopki → panel boczny `d2-header-footer-panel`.** Nowy komponent (TS+HTML+SCSS+spec) w `components/header-footer-panel/`, dokowany w tym samym slocie co `Wyszukiwanie` i `d2-table-properties-panel` (szerokość 300px, ten sam akcent #1a73e8). Sekcje: Widok („Inna pierwsza strona"), Wstaw (Obraz, Numery stron), Ustawienia (Format nagłówka/stopki, Usuń nagłówek/stopkę), Primary CTA „Zamknij nagłówek i stopkę". Stateless — wszystkie akcje delegowane do `wysiwyg-editor` (rule 9 — brak równoległej logiki).
  - Stary pływający `.header-toolbar`/`.footer-toolbar` USUNIĘTY z `wysiwyg-editor.html` (SCSS pozostaje jako martwy, nie używany — możliwe usunięcie w follow-up).
  - Koordynacja single-mode: `showHeaderFooterPanel = computed(() => editingSection() !== 'body' && !showFindReplace() && !showTablePanel())`. Find / Tables mają pierwszeństwo (spec #10); po ich zamknięciu HF panel wraca, jeśli edycja trwa. ESC nadal deleguje do `editor.stopEditingHeaderFooter()` (Phase 2 commit) → zamyka edycję → zamyka panel.
  - Pozioma linijka: spacer `ruler-h-search-spacer` reaguje też na `showHeaderFooterPanel()`.
- Nowy test `header-footer-panel.spec.ts` (6) + `document-editor.spec.ts` +5 (koordynacja Find/Tables/HF, ESC).
### Verified
- `tsc --noEmit` OK; `npm test` → **142/142 pass** (16 plików).
### Notes / ograniczenia
- **Top-margin drag w trybie edycji nagłówka**: usunięcie pływającego paska eliminuje overlay, który zasłaniał obszar interakcji nad/obok nagłówka. Rzeczywiste przeciągnięcie zależy od linijki pionowej (`d2-ruler` mode=vertical) — pełna weryfikacja wymaga manualnego testu w przeglądarce; w razie pozostałego konfliktu warto sprawdzić `z-index` `.page-header.editing` (obecnie 15) vs ruler.
- Nieużywane reguły SCSS `.header-toolbar`/`.footer-toolbar`/`.header-options-btn` itp. zostają jako dead-code (nie wpływają na layout). Czyszczenie — follow-up.

## 2026-05-28 — Faza 2: audyt trybu edycji nagłówka/stopki (vs spec)
### Changed (gap fixes)
- **Routing wariantu na edycji (rule 10: bez pozornej edycji).** `wysiwyg-editor.ts`:
  - `startEditingHeader`/`startEditingFooter` ładuje teraz wariant aktywny na stronie 0 (`_headerFirstPageHtml` gdy `differentFirstPage`, inaczej `_headerHtml`) — wcześniej zawsze ładował default niezależnie od wyświetlanego wariantu.
  - `onHeaderInput`/`onFooterInput` zapisuje do tego samego sygnału co załadowany wariant (nie do `_headerHtml` zawsze) — wcześniej edycja first-page niewidocznie nadpisywała default.
  - `onHeaderBlur`/`onFooterBlur`: analogicznie + wywołują `emitHeaderFooterChanges()` zamiast emitować niekompletne `{html, height}` (gubiło inne warianty u parenta).
  - `_computeHeaderContent`/`_computeFooterContent`: w trybie odd/even strona nieparzysta używa **kanonicznego** `_headerHtml`/`_footerHtml` (= „default" w OOXML), bo backend nie posyła osobnego `oddHtml`. Even strony bez zmian (`_headerEvenHtml`).
- **ESC zamyka tryb edycji nagłówka/stopki** (`document-editor.ts` `onEscapeKeydown`): gdy `editingSection() !== 'body'`, deleguje do `editor.stopEditingHeaderFooter()` i `preventDefault`. Tryb ma pierwszeństwo nad bocznymi panelami (wyszukiwanie, panel tabeli), bo to ostatnio aktywowany kontekst.
- **Przycisk „Zamknij nagłówek i stopkę"** dodany w obu toolbarach nagłówka i stopki (`wysiwyg-editor.html`, `wysiwyg-editor.scss` — `.header-toolbar-close`/`.footer-toolbar-close`, ten sam akcent co istniejący `1a73e8` w toolbarach). Pełni rolę „głównej" akcji wyjścia z trybu (spec funkcjonalny #3).
- Nowy plik testów `wysiwyg-editor.spec.ts` (7 testów): routing default vs first-page (header+footer), pełne emisje variantów, odd=canonical default, dokument bez nagłówka, `stopEditingHeaderFooter`. `document-editor.spec.ts` +4 testy ESC dla header/footer + pierwszeństwo nad side panel + no-op w body.
### Verified
- `npm test` (`ng test` → vitest): **131/131 pass** (15 plików; 11 nowych testów Phase 2).
- `tsc --noEmit`: OK.
### Notes / ograniczenia
- **Edycja even-page nieosiągalna z UI** — template renderuje contenteditable tylko dla strony 0; even-page edytowanie wymagałoby dodatkowego wejścia (nie w tym MVP). Even-page wciąż jest persistowany w round-trip (z importu).
- Toolbar główny (`d2-editor-toolbar`) — formatowanie tekstu — działa na aktywnym `editingSection` przez `getActiveEditable()` (istniejący kod, bez zmian); confirmed via existing wiring `executeCommand` (linia 1335 wysiwyg-editor).
- Undo/redo dla header/footer — `saveToUndoStack` jest wywoływany przez `onHeaderInput`/`onFooterInput` ścieżkę (`emitContent` debouncer); istniejące, bez zmian.

## 2026-05-28 — Faza 1b: pomiary wierności importu nag/stopki (stupki.docx)
### Changed
- Nowy plik `DocxToHtmlConverterFidelityTests.cs` (9 testów) — struktura odzwierciedla rzeczywisty `stupki.docx` (A4, pgMar 1417 twips, header/footer=708, obraz extent 1272540×327354 EMU, stopka L/C/R):
  - geometria: `Header.Height`/`Footer.Height` ≈ 1.25 cm; `Margins` ≈ 2.5 cm;
  - font family: dziedziczony z `docDefaults` (kontener `.header-footer-content` niesie `font-family`);
  - obraz: EMU → 133×34 px (`EmuToPx = emu/914400*96`), aspect zachowany w ±1 px (test 4:1 stays 4:1), `data-width-emu`/`data-height-emu` zachowane do round-tripu;
  - alignment akapitów w stopce: kolejność L/C/R zachowana, `text-align:center/right` emitowane.
- Diagnoza realnego `stupki.docx` potwierdza brak gapów na powyższych ścieżkach (T1-T12 z `<test_requirements>`; T3-T6 już objęte `DocxToHtmlConverterHeaderFooterTests`).
### Verified
- Pełny `D2ViewerEditor.Infrastructure.UnitTests`: **43/43 pass** (9 nowych fidelity).
### Notes
- `stupki.docx` nie używa tabel layoutowych ani anchor-positioned images — pokrycie tych przypadków pozostaje syntetyczne / przyszłe.
- Pozostaje R-11 (round-trip rozmiaru z `docDefaults` przez wrapper `.header-footer-content`) — bez sygnału z `stupki.docx`.

## 2026-05-28 — Round-trip first/even header/footer (zapis)
### Changed
- `HtmlToDocxConverter.AddHeaderAndFooter` zrefaktorowany: wydzielone `WriteHeaderPart`/`WriteFooterPart(html, type)` tworzą osobne `HeaderPart`/`FooterPart` per wariant (Default/First/Even) i wstawiają `HeaderReference`/`FooterReference` z odpowiednim `Type` (helpery `AddHeaderReference(id, type)`/`AddFooterReference(id, type)` zamiast hard-coded Default).
- Dodane `EnsureTitlePage` (wstawia `TitlePage` do `sectPr`, gdy emitowany wariant First) i `EnsureEvenAndOddHeaders` (wstawia `EvenAndOddHeaders` do `settings.xml`, gdy emitowany wariant Even). Helper `GetOrCreateSectionProps`.
- Nowy plik testów `HeaderFooterRoundTripTests.cs` (6 testów): write → read potwierdza, że `DifferentFirstPage`/`FirstPageHtml`/`DifferentOddEven`/`EvenHtml` przeżywają round-trip (header i footer, każdy wariant oddzielnie, plus „all combined").
### Verified
- Backend build OK; pełny `D2ViewerEditor.Infrastructure.UnitTests`: **34/34 pass** (6 nowych round-trip).
### Notes
- Zamyka brakujący kawałek: edycja wariantu first/even w UI jest teraz trwała (reguła 10).
- Pozostaje OPEN: **wiele sekcji** (`sectPr` per-section) — `R-10` zaktualizowany.

## 2026-05-27 — Fix layoutu: banner środowiska przycinał dolny pasek edytora

### Changed
- `styles.scss`: layout powłoki (`d2-root` flex column 100vh + `d2-root > d2-global-banners` `flex:0 0 auto` + routowane strony `d2-document-editor`/`d2-dashboard`/`d2-pdf-maintenance`/`d2-pdf-viewer`/`d2-admin-shell` `flex:1 1 0; min-height:0; overflow:hidden`) przeniesiony do **globalnych** (nieenkapsulowanych) styli. **Przyczyna błędu:** reguły flex dla stron były w stylach komponentu `App` (encapsulation Emulated), ale strony są wstawiane przez `<router-outlet>` jako rodzeństwo i NIE dziedziczą atrybutów `_ngcontent` App → selektory `d2-document-editor{flex:1}` nie działały. Edytor zostawał przy `:host{height:100%}` = pełne 100vh; zawsze widoczny banner środowiska (~22px) spychał dolny pasek (zoom / liczba stron / wersja) poza ekran, gdzie `overflow:hidden` go ucinał.
- `app.ts`: usunięto martwe selektory routowanych komponentów ze styli komponentu (zostaje tylko `:host`), z komentarzem wskazującym globalny layout.

### Verified
- `npx ng test` — 101/101 passed. `npx ng build` — OK.
- Manualnie: z widocznym bannerem środowiska edytor mieści się w oknie, dolny pasek (zoom/strony/wersja) w całości widoczny; działa dla 4 kombinacji bannerów.

## 2026-05-27 — Nazewnictwo admina, dok tabela↔wyszukiwanie, ukrycie pozycji „Wstaw", „Akapit" w toolbarze

### Changed
- **Admin — nazewnictwo** (`Wysyłki` → `Pliki do wysłania`): `admin-shell.html` (nav), `admin-deliveries.html` (tytuł strony + tekst ładowania), `admin-deliveries.ts` (komunikat błędu). Nie ruszano nazw technicznych (trasa `deliveries`, klasy, DTO, statusy — np. status „Wysyłanie" zostaje, bo to stan akcji, nie nazwa obszaru).
- **Dok boczny — przełączanie tabela ↔ wyszukiwanie** (`document-editor.ts` `syncTablePanel`): naprawiono konflikt — gdy karetka wraca do tabeli przy otwartym wyszukiwaniu, dok przełącza się na formatowanie tabeli (zamyka wyszukiwanie przez `closeFindReplace()`). Wcześniej warunek `!showFindReplace()` blokował pokazanie panelu tabeli (ADR-0006 „search ma pierwszeństwo" — świadomie nadpisane wg decyzji zadania). Dok ma jeden aktywny tryb; `tablePanelManuallyClosed` (×) nadal respektowane; ESC/X bez zmian.
- **Menu „Wstaw" — ukryte pozycje** (`document-editor.html`): QR Code / Kod kreskowy, Linia pozioma, Podział strony owinięte w `@if (false)` (ukryte, logika `openBarcodeDialog`/`insertHorizontalLine`/`insertPageBreak` zostaje). Separatory uporządkowane — brak pustych grup/podwójnych separatorów.
- **Toolbar — przycisk „Akapit"** (`editor-toolbar`): nowy `@Output() openParagraph` + przycisk w grupie „Listy i wcięcia" (ta sama ikona co menu „Narzędzia"), podpięty w `document-editor.html` do istniejącego `openParagraphDialog()` — bez duplikacji dialogu/logiki. `aria-label="Akapit"`. Menu „Narzędzia" bez zmian.

### Verified
- `npx ng test` — **101/101 passed (12 plików)**; nowe/rozszerzone: `admin-shell.spec.ts` (3), `editor-toolbar.spec.ts` (+1 „Akapit" = 7), `document-editor.spec.ts` (+5 przełączanie doku = 12).
- `npx ng build` — OK (tylko istniejące ostrzeżenia budżetu).

### Notes
- **QR/barcode** dostępne też przez `(insertBarcode)` toolbara, ale toolbar NIE renderuje widocznego przycisku QR — więc ukrycie w menu „Wstaw" wystarcza. `insertHorizontalLine`/`insertPageBreak` były tylko w menu „Wstaw". **Do potwierdzenia:** czy ukryć też dialog kodu kreskowego z innych ścieżek (obecnie brak innej widocznej).
- **Scenariusz manualny menu „Wstaw"**: otwórz „Wstaw" → brak QR/Linia pozioma/Podział strony; widoczne: Obraz, Tabela, (separator), Nagłówek/Stopka… — bez pustych grup.
- **Scenariusz manualny dok**: klik w tabelę → panel formatowania; lupa/Ctrl+F → wyszukiwanie; ponowny klik w tabelę → panel wraca do formatowania (wyszukiwanie zamknięte); ESC i × zamykają dok.

## 2026-05-27 — Globalny banner środowiska + naprawa layoutu paska braku API

### Changed
- **Nowy kontener bannerów** `d2-global-banners` (`components/global-banners/`) — renderowany w `App` w normalnym flow flex (zamiast bezpośredniego `<d2-offline-banner>`). Stała kolejność: banner środowiska na górze, pasek offline/API pod nim. Bannery są w flow → `App` (`:host` flex column, 100vh) rezerwuje ich wysokość, edytor (`flex:1`) nigdy nie jest przykrywany — stabilne dla 4 kombinacji (brak / tylko env / tylko API / oba).
- **Nowy banner środowiska** `d2-environment-banner` (`components/environment-banner/`) — prezentacyjny; źródło środowiska = `BuildInfoService.environment()` (jedyne wiarygodne: health-check API z fallbackiem na front-config — **uwaga:** brak `fileReplacements` w `angular.json`, więc `environment.ts` jest zawsze `PRD`, dlatego nie używamy go wprost). Mapowanie przez czystą funkcję `core/utils/environment-banner.util.ts` (`resolveEnvironmentBanner`): Local→niebieski, DEV→zielony, UAT/TST→jasnożółty, PRE→purpurowy, PROD/PRD→ukryty, nieznane→neutralny fallback `Wersja {nazwa}`. `role="status"`, kontrast AA.
- **Naprawa paska braku API** (`offline-banner.scss`): usunięto arbitralny `z-index: 10000` (pasek jest w flow — wysoki z-index tylko ryzykował przebicie przez modale edytora o z-index 1000–2000); dodano `width:100%` i komentarz o roli `position: relative` (kotwica przycisku ×). Logika wykrywania (`ConnectionStatusService`: online/offline + health-check 30 s + interceptor) **bez zmian** — poprawka czysto prezentacyjno-layoutowa.
- `App` (`app.ts`): import `GlobalBannersComponent`, selektor `d2-global-banners` w `:host` (`flex-shrink:0`).
- Naprawiono nieaktualny scaffold `app.spec.ts` (oczekiwał „Hello, frontend"): teraz `provideRouter([])` + stub `BuildInfoService`/`ConnectionStatusService`, asercja montażu `d2-global-banners`.

### Verified
- `npx ng test` — **92/92 passed (11 plików)**; nowe specy: `environment-banner.util.spec.ts` (10), `environment-banner.spec.ts` (10: pełna macierz env + reaktywność + role), `global-banners.spec.ts` (4: kolejność, widoczność offline, niezależność, prod). Suite w pełni zielony (naprawiony `app.spec.ts`).
- `npx ng build` — OK (tylko istniejące ostrzeżenia budżetu).

### Notes
- **Scenariusz manualny**: na DEV/UAT/itd. (gdy API health zwraca env) pojawia się kolorowy pasek na górze, ponad paskiem offline; oba w flow, edytor/toolbar/panel boczny nieprzykryte. Gdy API niedostępne — pasek offline pod bannerem env; gdy API wróci — znika. Na PROD (PRD) banner env ukryty.
- Banner reaguje na health-check: do pierwszej odpowiedzi env=PRD (front fallback) → ukryty (brak migotania błędną etykietą), po odpowiedzi API pokazuje właściwe środowisko.

## 2026-05-27 — ESC zamyka boczny panel + weryfikacja przycisku „Cofnij" (undo)

### Changed
- Frontend `document-editor.ts`: nowy `@HostListener('document:keydown.escape')` (`onEscapeKeydown`) zamyka aktywny boczny panel (Wyszukiwanie / właściwości+stylizacja tabeli) tą samą logiką co przycisk × (`closeFindReplace` / `closeTablePanel`), przez wspólną metodę `closeActiveSidePanel()`. Listener jest na `document` (panel może nie mieć focusu — karetka w edytorze). Pierwszeństwo: jeśli ESC obsłużył już szczegółowy handler (`event.defaultPrevented`, np. deselekcja obrazu w `wysiwyg-editor`) albo otwarty jest dialog/menu kontekstowe (`isAnyDialogOpen()`/`showContextMenu()`) — panel nie jest ruszany; gdy żaden panel nie jest otwarty — brak skutków. Focus wraca do edytora tylko, gdy był wewnątrz panelu.
- Brak zmian mechanizmu undo: przeanalizowano i potwierdzono poprawność. Przycisk „Cofnij" w `editor-toolbar` (`(click)="executeCommand('undo')"`, `[disabled]="!editorState?.canUndo"`) używa tej samej metody `WysiwygEditorComponent.undo()` co skrót Ctrl+Z (`handleKeyboard`). Operacje tabeli (wiersze/kolumny/obramowania/scal/kolor) trafiają do historii przez `notifyEditorChange()` → syntetyczny `input` → `onContentChange` → `saveToUndoStack()`.

### Verified
- `npx ng test` — nowe specy: `editor-toolbar.spec.ts` (6) + `document-editor.spec.ts` (7) = 13/13 passed; cała reszta bez regresji (jedyne czerwone to znany, niezależny scaffold `app.spec.ts` — brak `_HttpClient`).
- `npx ng build` — OK (tylko istniejące ostrzeżenia budżetu SCSS/initial).
- Undo na poziomie contenteditable/`execCommand` nie jest testowalny w jsdom — scenariusz manualny niżej.

### Notes
- **Scenariusz manualny undo** (uruchom GUI, otwórz dokument do edycji): (1) wpisz tekst → przycisk „Cofnij" się aktywuje → klik cofa wpis; (2) zaznacz tekst, kliknij **B** (bold) → „Cofnij" cofa pogrubienie; (3) klik w tabelę → panel tabeli → dodaj wiersz / zmień obramowanie → „Cofnij" cofa operację tabeli; (4) Ctrl+Z daje ten sam efekt co przycisk; (5) „Ponów"/Ctrl+Y przywraca; (6) brak błędów w konsoli; (7) gdy stos pusty — „Cofnij" jest wyszarzony.
- **Scenariusz manualny ESC**: otwórz panel Wyszukiwania → ESC zamyka; klik w tabelę (panel właściwości/stylizacji) → ESC zamyka; ESC przy zamkniętym panelu → nic; zaznacz obraz w edytorze przy otwartym panelu → ESC najpierw odznacza obraz (panel zostaje), kolejny ESC zamyka panel; otwarty dialog (np. wstaw tabelę) + ESC → panel w tle zostaje.
- Redundancja w `wysiwyg-editor`: każdy edytor strony ma jednocześnie `(input)="onPageInput()"` (debounce 500 ms) ORAZ `addEventListener('input', …)` → `onContentChange()` (snapshot natychmiastowy). Undo działa, ale jest granularne (per znak). Nie ruszano — poza zakresem, ryzyko regresji undo dla tabel (zależnych od ścieżki `onContentChange`).

## 2026-05-27 — Fix odwzorowania nagłówka/stopki DOCX→HTML (za duży tekst, kolor, fallback)

### Changed
- Backend `DocxToHtmlConverter`: nagłówek/stopka są teraz owijane w `<div class="header-footer-content" style="font-family:..;font-size:..pt">` z domyślnym krojem/rozmiarem z `w:docDefaults/rPrDefault` — analogicznie do body `.document-content`. Wydzielono wspólny helper `BuildDefaultContainerCss()` (body + header/footer). **Przyczyna „za dużego tekstu":** runy nagłówka/stopki bez własnego `w:sz`/`w:rFonts` nie miały żadnego kontenera z domyślnym rozmiarem (body go miało), więc dziedziczyły domyślny rozmiar EDYTORA, nie dokumentu.
- Frontend `wysiwyg-editor.scss`: `.header-display/.footer-display/.header-editor-content/.footer-editor-content` dostały domyślny `font-family` (corporate var), `font-size:11pt`, `line-height:1.15`, `color:#000` — fallback dla dokumentów bez kontenera z backendu oraz dla regionu edycji. Usunięto `color:#202124` (Word „auto" = czarny). Kontener z backendu (inline font-size z docDefaults) ma wyższą specyficzność i nadpisuje, gdy DOCX definiuje rozmiar; inline color/size runów zawsze wygrywa.
- Testy backendu: `DocxToHtmlConverterHeaderFooterTests` (5) — kontener z docDefault 10pt w nagłówku i stopce, zachowanie jawnego koloru (`#1F3864`/`#C00000`) i rozmiaru (8pt = 16 half-points), brak docDefaults → kontener bez wymuszonego rozmiaru.

### Verified
- `dotnet build` Infrastructure — OK (0 błędów). `dotnet test --filter DocxToHtmlConverterHeaderFooter` — 5/5 passed.
- `npm run build` (front) — OK.

### Notes
- Jednostki: `w:sz` jest w half-points → `pt = sz/2` (potwierdzone w `GetRunStyleClean`/`ConvertRunPropertiesToCss` i teraz w kontenerze docDefaults).
- **Niezałatwione (udokumentowane):** R-10 — bierzemy tylko `HeaderParts.FirstOrDefault()`, bez default/first-page/even-odd i bez wielu sekcji; R-11 — round-trip rozmiaru nagłówka zależy od `docDefaults` (kontener spłaszczany przy zapisie, fallback frontu 11pt zapewnia spójność wizualną).

## 2026-05-27 — Bugfix: panel obramowań gubił aktywną tabelę + nieczyszczone zaznaczenie

### Changed
- GUI: `core/utils/table-context.util.ts` — nowa czysta funkcja `resolveTableContext(anchorNode, editorEl)` → `outside-editor` / `in-table` / `outside-table`. Wydzielona z `detectTableContext`, testowalna.
- GUI: `document-editor.ts` `detectTableContext()` — **nie czyści już aktywnej tabeli, gdy selekcja jest poza edytorem** (interakcja z panelem/toolbar = zachowaj ostatni kontekst). Czyszczenie tylko gdy karetka jest realnie w treści poza tabelą (`outside-table`) — wtedy dodatkowo `clearCellSelection()` (usuwa klasę `table-cell-selected`).
- GUI: podpięto `(selectionChange)="onEditorSelectionChange()"` w `document-editor.html` (wcześniej `selectionChange` z edytora był nieobsłużony) — kliknięcie innego akapitu aktualizuje teraz kontekst tabeli i zamyka/aktualizuje panel.
- GUI: testy `table-context.util.spec.ts` (outside-editor/in-table/outside-table/null) + panel: „border interactions never emit close".

### Verified
- `npm run build` — OK. `npx ng test --watch=false` — 55 passed (2 stare scaffoldowe `app.spec.ts` padają niezależnie).

### Notes
- **Przyczyna #1** (panel pokazywał „Wybierz tabelę" po kliknięciu obramowania): `applyTableBorderScope` → `notifyEditorChange()` → `input` → `updateState` → `stateChange` → `detectTableContext`, które przy selekcji poza edytorem (focus w panelu) zerowało `activeTable`. Naprawa: zachowanie kontekstu przy `outside-editor` (model „last known table context").
- **Przyczyna #2** (tabela nadal wyglądała na zaznaczoną po kliknięciu akapitu): `selectionChange` edytora nie był podpięty, więc sam ruch karetki nie aktualizował kontekstu. Naprawa: podpięcie `selectionChange` + czyszczenie wizualnego zaznaczenia przy `outside-table`.
- `onPanelMouseDown` (preventDefault na nie-inputach) pozostaje jako komplementarne zabezpieczenie zachowujące karetkę; naprawa działa też, gdyby focus jednak uciekł.

## 2026-05-27 — Auto-wykrywanie zakresu obramowania (ukryty ręczny wybór celu)

### Changed
- GUI: usunięto ręczny przełącznik celu („Cała tabela / Komórka / Zaznaczenie") z panelu obramowań (`TableBorderTarget`, `TABLE_BORDER_TARGETS`, sekcja „Zakres", `setTableBorderTarget`). Zakres jest teraz **wnioskowany z bieżącego zaznaczenia**.
- GUI: `core/utils/table-style.util.ts` — nowa czysta funkcja `classifyBorderTarget(table, cells)` → `BorderTargetInfo { kind, rows, cols }`: jedna komórka → `cell`; pełna szerokość → `row`; pełna wysokość → `column`; pełna szer.+wys. → `table`; reszta → `range`; pusty zbiór → `none`.
- GUI: `document-editor.ts` — `resolveAutoTargetCells` (zaznaczone komórki → ten zbiór; brak → aktywna komórka; dalej cała tabela) + `borderTargetInfo` (computed, reaktywny na `selectedCells`/`activeTableCell`). Klik ikony „gdzie narysować" stosuje linię do auto-celu (`applyBorderToCells` liczy krawędzie względem prostokąta zbioru).
- GUI: panel pokazuje tylko **podpis** „Zastosowanie: aktywna komórka / zaznaczone komórki (R×C) / cały wiersz / cała kolumna / cała tabela" (read-only, `aria-live`), bez kontrolki wyboru. UX: użytkownik wybiera rodzaj/grubość/kolor linii i klika miejsce — zakres dobiera się sam.
- GUI: testy — `classifyBorderTarget` (cell/row/column/table/range/none) + panel (podpis zamiast selektora; brak `.table-panel-segmented`).

### Verified
- `npm run build` — OK. `npx ng test --watch=false` — 49 passed (2 stare scaffoldowe `app.spec.ts` padają niezależnie).

### Notes
- Domyślny zakres przy samej karetce (brak zaznaczenia) = **aktywna komórka** (jak w MS Word). Wykrycie „cały wiersz/kolumna" działa, gdy użytkownik zaznaczy te komórki (drag custom cell-selection); edytor nie ma osobnego gestu „klik nagłówka wiersza/kolumny" — to świadome ograniczenie, bezpieczny fallback do zaznaczonego zbioru.
- Tabele z mocno scalonymi komórkami mogą dać przybliżoną *etykietę* zakresu; samo rysowanie obramowania działa zawsze na przekazanym zbiorze komórek.

## 2026-05-27 — Szczegółowy edytor obramowań tabeli (zamiast galerii presetów)

### Changed
- GUI: **usunięto galerię gotowych stylów tabel** (`TABLE_STYLE_PRESETS`, sekcje „Style tabeli"/„Opcje stylu", `applyTablePreset`, markery `data-table-style*`, `readTableStyleState`). Zakładka „Style" → **„Obramowania"**.
- GUI: `models/table-style.model.ts` przebudowany pod obramowania: `TableBorderLineStyle` (solid/dashed/dotted/double/none), rozszerzony `TableBorderScope` (all/none/outer/inner/inner-horizontal/inner-vertical/top/bottom/left/right), `TableBorderSettings` z polem `style`, `TableBorderTarget` (table/cell/selection) + listy prezentacyjne (`TABLE_BORDER_LINE_STYLES/_WIDTHS/_COLORS/_SCOPES/_TARGETS`).
- GUI: `core/utils/table-style.util.ts` — `applyBorderToCells(cells, scope, border)` liczy krawędzie względem prostokąta opisanego na **dowolnym zbiorze komórek** (tabela / komórka / zaznaczenie); `applyBorderScope` deleguje dla całej tabeli; `restoreDefaultTableBorders` przywraca siatkę 1px. Linie kodowane jako `Npx <style> <color>` w inline `border*` — trwałość wg ADR-0007.
- GUI: `table-properties-panel` — zakładka „Obramowania" z sekcjami: **Zakres** (cel: cała tabela/komórka/zaznaczenie), **Rodzaj linii** (chipy z podglądem), **Grubość linii** (Cienka/Standardowa/Średnia/Gruba), **Kolor linii** (paleta + color picker + reset), **Gdzie narysować** (10 neutralnych ikon SVG: wszystkie/brak/zewn./wewn./poziome/pionowe/góra/dół/lewa/prawa), **Reset** (domyślne obramowanie / usuń obramowania). Podświetlany stan aktywny (rodzaj/grubość/kolor/ostatni zakres/cel).
- GUI: `document-editor.ts` — sygnały `tableBorderColor/Width/Style/Target` + `lastBorderScope`; metody `applyTableBorderScope` (z rozwiązaniem celu i fallbackiem zaznaczenie→komórka→tabela), `clearTableBorders`, `restoreDefaultTableBorders`, `resetTableBorderSettings`, settery pióra. Usunięto metody presetów.
- GUI: testy przepisane — `table-style.util.spec.ts` (zakresy, cel komórka/zaznaczenie, rodzaj/grubość linii, restore default) i `table-properties-panel.spec.ts` (zakładka Obramowania, ikony zakresu, zmiana rodzaju/grubości/koloru/celu, reset).

### Verified
- `npm run build` — OK (tylko istniejące ostrzeżenia budżetu).
- `npx ng test --watch=false` — 42 passed. Padają tylko 2 stare scaffoldowe `app.spec.ts` (niezwiązane).

### Notes
- Obramowanie działa na **trzech poziomach**: cała tabela, pojedyncza komórka, zaznaczony fragment (custom cell-selection). „Zaznaczenie" bez zaznaczonych komórek → bezpieczny fallback do aktywnej komórki.
- Zakresy są **addytywne** (dokładają krawędzie, jak przyciski w Word); „Brak" czyści, „Domyślne obramowanie" przywraca siatkę 1px. Zachowanie treści gwarantowane (modyfikujemy tylko `border*`).
- Galeria presetów świadomie wycofana w tej iteracji (priorytet: precyzyjna edycja linii). Patrz ADR-0007 (zaktualizowane).

## 2026-05-27 — Stylizacja tabel (gotowe style, obramowania, opcje, reset) + zakładki w panelu

### Changed
- GUI: nowy model `models/table-style.model.ts` — `TableStylePresetId`, `TableStyleOptions` (headerRow/bandedRows/firstColumn/lastColumn), `TableBorderScope`, `TableBorderSettings`, `TableStylePreset` + `TABLE_STYLE_PRESETS` (6 neutralnych stylów: Prosty, Siatka, Nagłówek, Naprzemienny, Biznesowy, Minimalny).
- GUI: nowy util `core/utils/table-style.util.ts` — czyste, testowalne funkcje DOM: `applyTablePreset` (idempotentne przeliczenie wyglądu), `applyBorderScope` (all/none/outer/inner), `resetTableStyle` (przywraca domyślny wygląd, zachowuje treść), `readTableStyleState` (odczyt presetu/opcji z markerów `data-*`). Styl utrwalany jako **style inline** na `<table>/<tr>/<td>` — spójnie z istniejącymi akcjami tabeli, przeżywa zapis (HTML→DOCX) i ponowne otwarcie. Patrz ADR-0007.
- GUI: `table-properties-panel` rozbudowany o **dwie zakładki** — „Układ" (akcje strukturalne: wiersze/kolumny, komórki, rozmiar, wygląd, usuń) i „Style" (gotowe style z miniaturami, opcje stylu, obramowania z kolorem/grubością, reset). Panel pozostaje czysto prezentacyjny (`@Input` stanu, `@Output` per akcja). Dodano wewnętrzny pionowy scroll body (`min-height:0` + `overflow-y:auto`, `:host` stretch).
- GUI: `document-editor.ts` — sygnały `activeTablePresetId`/`tableStyleOptions`/`tableBorderColor`/`tableBorderWidth`; metody `applyTableStylePreset`/`toggleTableStyleOption`/`applyTableBorderScope`/`setTableBorderColor`/`setTableBorderWidth`/`resetTableStyle` (glue: util na `activeTable()` + `notifyEditorChange()` → auto-save). `detectTableContext()` czyta stan stylu z aktywnej tabeli (`readActiveTableStyle`).
- GUI: testy `core/utils/table-style.util.spec.ts` (11 przypadków: borders none/outer, preset header/banded, zachowanie treści, recompute toggle, first/last column, reset, persist+read state, defaults) oraz rozszerzone `table-properties-panel.spec.ts` (zakładki, presety, opcje, obramowania, reset).

### Verified
- `npm run build` — OK (tylko istniejące ostrzeżenia budżetu).
- `npx ng test --watch=false` — 41 passed (panel + util + badge). Padają tylko 2 stare scaffoldowe `app.spec.ts` (niezwiązane).

### Notes
- Trwałość: style inline przeżywają round-trip HTML w edytorze i zapis do DOCX. Markery `data-table-style*` (zapamiętany preset/opcje dla panelu) mogą zostać usunięte przez konwersję DOCX — wtedy **wygląd pozostaje** (inline), a panel po ponownym otwarciu pokazuje stan domyślny presetu. Wierność DOCX↔HTML zależy od konwertera serwerowego (to samo ryzyko co istniejące cieniowanie/obramowania).
- Bez zmian backendu/API/modelu dokumentu (funkcja czysto frontendowa).
- MVP w pełni: gotowe style, wiersz nagłówka, wiersze naprzemienne, obramowania (kolor/grubość/zakres), reset. Pełne właściwości DOCX (styl linii dashed/double, padding per komórka, wyrównanie pionowe w UI) — etap późniejszy.

## 2026-05-27 — Konfiguracja tabeli przeniesiona do bocznego panelu

### Changed
- GUI: nowy komponent prezentacyjny `components/table-properties-panel` (`d2-table-properties-panel`) — boczny panel „Ustawienia tabeli" dokowany po lewej, spójny z panelem „Wyszukiwanie" (`.search-panel`). Czysto prezentacyjny: wejścia `hasActiveTable`/`gridLinesVisible`/`shadingColors`, wyjścia per akcja; cała logika operująca na `activeTable()`/`activeTableCell()` pozostaje w `document-editor`.
- GUI: `document-editor.html` — usunięto poziomy pasek `.table-toolbar` (pojawiający się przy `isInTable()`) oraz pływający `.shading-dropdown`. Wstawiono `<d2-table-properties-panel>` w doku obok panelu wyszukiwania; rozpórka poziomej linijki reaguje teraz na `showFindReplace() || showTablePanel()`.
- GUI: `document-editor.ts` — sygnał `showTablePanel` + flaga `tablePanelManuallyClosed`; `detectTableContext()` woła `syncTablePanel()` (auto-otwarcie w tabeli, zamknięcie po wyjściu, respekt ręcznego zamknięcia). `openFindReplace()` przejmuje dok (chowa panel tabeli), `closeFindReplace()` przywraca panel tabeli jeśli karetka nadal w tabeli. `tableDeleteTable()` zamyka panel. Panel wykluczony z czyszczenia zaznaczenia komórek w `onCellMouseDown` (scalanie działa). Wszystkie metody `table*()`, `setCellColor`, `clearCellColor` re-użyte bez zmian logiki.
- GUI (test infra): cel `test` w `angular.json` uzupełniony o `buildTarget`/`tsConfig`/`runner: vitest`/`setupFiles` + nowy `src/test-setup.ts` (Zone.js + zone.js/testing). Wcześniej cel testu nie miał `buildTarget`, więc `ng test` w ogóle się nie uruchamiał.
- GUI: testy `components/table-properties-panel/table-properties-panel.spec.ts` (stan pusty, render sekcji, emisja wszystkich akcji, kolor cieniowania, etykieta linii siatki, `preventDefault` mousedown).

### Verified
- `npm run build` — OK (tylko istniejące ostrzeżenia budżetu bundle/SCSS).
- `npx ng test --watch=false` — panel 6/6 i `document-classification-badge` 4/4 passed (27 passed łącznie). Padają jedynie 2 przestarzałe testy scaffoldowe `app.spec.ts` (asercja „Hello, frontend" + brak providera HttpClient) — niezwiązane z tą zmianą, ujawnione przez naprawę uruchamiania testów.

### Notes
- Decyzja UX dot. konfliktu z wyszukiwaniem: jeden dok po lewej; wyszukiwanie (jawnie wywołane) ma pierwszeństwo, panel tabeli (kontekstowy) nie wypiera go i wraca po zamknięciu wyszukiwania, jeśli karetka jest nadal w tabeli. Patrz ADR-0006.
- Stan pusty panelu („Kliknij tabelę…") jest zabezpieczeniem — przy normalnym flow panel zamyka się po opuszczeniu tabeli, więc rzadko widoczny.
- `app.spec.ts` to stary scaffold (`Hello, frontend`) niepasujący do realnego `App` — do przepisania osobno; nie ruszane w tej zmianie.

## 2026-05-27 — „Wklej tylko tekst" (Paste Text Only) w edytorze WYSIWYG

### Changed
- GUI: nowy util `core/utils/paste-text.util.ts` — czyste, testowalne funkcje `normalizeWhitespace`, `htmlToText` (paragrafy/nagłówki → nowe linie, listy z markerem `•`/`1.`, komórki tabel rozdzielane tabulatorem, `<a>` → tekst bez URL, pomijanie `script`/`style`) oraz `resolvePlainText` (preferuje `text/plain`, fallback z HTML).
- GUI: `wysiwyg-editor.ts` — skrót `Ctrl/Cmd+Shift+V` oznacza najbliższe wklejenie jako „tylko tekst" (okno czasowe 1 s, by nieużyty skrót nie wpłynął na kolejne zwykłe Ctrl+V). `handlePaste` przepuszcza wklejenie przez util; ścieżka „tylko tekst" wstawia przez istniejące `insertText` (zachowuje natywne undo i emituje `onContentChange` → auto-save/`contentChange`). Zwykłe Ctrl+V bez zmian (HTML po `sanitizeHtml`).
- GUI: testy `core/utils/paste-text.util.spec.ts` (17 przypadków: whitespace, nbsp, znaki zero-width, taby/TSV, listy, tabele, linki, script/style).

### Verified
- `npm run build` — OK (tylko istniejące ostrzeżenia budżetu bundle/SCSS).
- `npx vitest run --environment jsdom src/app/core/utils/paste-text.util.spec.ts` — 17/17 passed.

### Notes
- Sanitizacja HTML przy zwykłym wklejaniu nadal opiera się na regexie (`sanitizeHtml`) — patrz R-09 (rekomendacja DOMPurify, wymaga decyzji o zależności).
- Toolbar/menu kontekstowe „Wklej tylko tekst" (poza zdarzeniem paste, przez `navigator.clipboard.readText()`) — nie wdrożone, kolejny inkrement.

## 2026-05-26 — Microsoft Entra ID: uwierzytelnianie + role (docelowo)

### Changed
- **Backend (Internal API):** JwtBearer (`Microsoft.AspNetCore.Authentication.JwtBearer`), sekcja `AzureAd` (Authority/Audience/CorporateKeyClaim/role), `RoleClaimType="roles"`. Policy `RequireAppEmployee`/`RequireAppAdmin`. `[Authorize]` na `BaseApiController` (cała aplikacja za logowaniem; Health anonimowy). Endpointy admina (`GET /`, `deliveries`, `deliveries/{id}/retry`) → `RequireAppAdmin`.
- **Tożsamość:** `ICurrentUserProvider` + `IsAdmin`; `ClaimsCurrentUserProvider` (CorporateKey z konfigurowalnego claimu access tokena, admin z roli). `HttpHeaderCurrentUserProvider` jako DEV. `SystemCurrentUserProvider` w External API.
- **Dostęp do dokumentu:** `DocumentAccessGuard` — **APP_Admin omija `allowedCorporateKeys`** (ADR-0011).
- **Frontend:** MSAL (`@azure/msal-angular`/`-browser`) — `msal.config`, providery w `app.config` (MsalInterceptor + init), `MsalGuard` na całej aplikacji, `appAdminGuard` (UX) na `/admin`, obsługa redirect w `App`. `CurrentUserService` czyta rolę/`ck` z konta MSAL (UX). Usunięto `corporateKeyInterceptor`. Sekcja `auth` w `environment(.development).ts`.

### Verified
- Backend: Domain (63) + Application (155) zielone wcześniej; dodano test bypassu admina. Build pełnej solucji blokowany środowiskowo (DLL zajęte przez API/Rider) — weryfikacja przez projekty testowe.
- **Niezweryfikowane (kroki deployowe):** `dotnet restore` JwtBearer; **frontend wymaga `npm install`** MSAL (do tego czasu „Cannot find module '@azure/msal-angular'" jest oczekiwane); `app.spec.ts` będzie wymagał providerów MSAL w TestBed.

### Notes
- ADR-0011: App Roles (nie groups), CorporateKey z access tokena (claim-only), logowanie dla całej aplikacji, admin omija `allowedCorporateKeys`. Wartości Entra to placeholdery per środowisko (nie sekrety).

## 2026-05-26 — Kontrola dostępu do dokumentu (allowedCorporateKeys) — v1

### Changed
- Domena: `DocumentAccessPolicy` (czysta reguła: brak/null/pusta lista → publiczny dla posiadacza linku; niepusta → tylko pasujący `CorporateKey`; porównanie Trim + Ordinal-IgnoreCase).
- Application: `ICurrentUserProvider` (seam tożsamości), `IDocumentAccessGuard`/`DocumentAccessGuard` (parsuje `allowedCorporateKeys` z `documents.metadata`, stosuje politykę). Rejestracja w DI.
- Api: `HttpHeaderCurrentUserProvider` (v1: czyta nagłówek `X-Corporate-Key`; seam pod Entra ID) + `AddHttpContextAccessor`.
- `Result`/`Result<T>`: nowy stan `Forbidden`/`IsForbidden`.
- Egzekwowanie (403) w handlerach zwracających treść/metadane: `GetDocumentMetadata` (brama GUI), `GetDocument`, `GetDocumentBaseContent`, `GetDocumentVersionContent`. Kontroler mapuje `Forbidden → 403`.
- GUI: `CurrentUserService` (seam, opcjonalny `localStorage('corporateKey')`), `corporateKeyInterceptor` (dodaje `X-Corporate-Key`), `documentAccessGuard` (CanActivate na `/editor` i `/viewer` — blokuje wejście przed inicjalizacją edytora), wspólny `DocumentAccessDeniedComponent` + trasa `/access-denied`. `http-error.interceptor` nie pokazuje toasta dla 403 (obsługuje guard/widok).

### Verified
- `dotnet build D2ViewerEditor.Application.UnitTests` — 0 błędów (testy referencują Domain+Application → potwierdza kompilację warstw). Pozostałe MSB3021/3026 w solucji to blokady DLL przez działające API/Rider, nie błędy kodu.
- Nowe testy: `DocumentAccessPolicyTests` (Domain), `DocumentAccessGuardTests` (Application). Zaktualizowano konstruktory w `GetDocument*`/`GetDocumentMetadata*` testach (nowa zależność guarda).

### Notes
- v1 nie ma realnej tożsamości — `CorporateKey` z nagłówka; dokumenty bez `allowedCorporateKeys` pozostają publiczne (brak regresji). Wszystkie odmowy → 403 (401 zarezerwowany na fazę Entra ID). Backend = źródło prawdy; front tylko blokuje wejście i pokazuje widok.

## 2026-05-25 — Panel admina „Wysyłki" + uruchomienie migracji DB

### Changed
- GUI: nowy widok `/admin/deliveries` (`admin-deliveries` component) — lista zadań wysyłki ze statusem, liczbą prób, datami (utworzono/ostatnia/następna/deadline), `locked_by` i ostatnim błędem; przycisk „Ponów" dla `DeadLettered`/`FailedPermanently`. Filtr per kolumna + dropdown statusu (steruje zapytaniem do API) + paginacja klienta, wizualnie spójny z `/admin/files`. Statusy tłumaczone na PL w warstwie prezentacji (enum bez zmian). Link w `admin-shell` nav. Trasa-dziecko w `app.routes.ts`.
- GUI: `document-storage.service.ts` — `getDeliveriesByStatus(status, skip, take)` + `retryDelivery(id)` + interfejsy `DeliveryListItem`, `RequeueDeliveryResult`.
- Application: `DeliveryListItemDto` rozszerzony o `LockedUntil` i `LockedBy` (monitoring claimu workera); mapowanie w `GetDeliveriesByStatusQueryHandler`.
- DB: wykonano skrypty `infra/sql/001`–`007` na lokalnym PostgreSQL (Podman `d2viewereditor_postgres`). 001–006 już istniały (idempotentne), `007` utworzył tabelę `document_deliveries`.

### Verified
- `dotnet build D2ViewerEditor.sln` — 0 błędów (ostrzeżenia tylko MSB3026 — zablokowane DLL przez działające API/Rider).
- Weryfikacja schematu: `document_deliveries` z 20 kolumnami, 4 indeksami (w tym partial `ux_..._active_per_document`, `ix_..._due`, `ix_..._stuck`), CHECK statusu, FK do `documents`.

### Notes
- Zamyka rekomendację „panel admina dla `DeadLettered`" z `TASK_HANDOFF`. Filtry tekstowe działają po stronie klienta na pobranej stronie (jak `admin-files`).

## 2026-05-25 — Funkcja „Zakończ i wyślij" (asynchroniczna wysyłka na returnUrl)

### Changed
- Domena: nowy agregat `DocumentDelivery` (+ enum `DeliveryStatus`: Pending/Sending/RetryScheduled/Sent/FailedPermanently/DeadLettered), interfejsy `IDocumentDeliveryRepository`, `IDeliverySender`, `IBackoffStrategy`. `IDocumentStorageService.UploadRawAsync` (snapshot pod dowolną nazwą).
- Application: `FinishAndSendDocumentCommand` (+handler+validator), `GetDeliveryStatusQuery`, `RequeueDeliveryCommand`, `GetDeliveriesByStatusQuery`. Idempotencja wielokrotnego kliknięcia (aktywne zadanie per dokument).
- Infrastructure: `DocumentDeliveryRepository` (claim `FOR UPDATE SKIP LOCKED` + lease), `HttpDeliverySender` (Idempotency-Key), `ExponentialJitterBackoff`, `DeliveryAttemptRunner`, `DocumentDeliveryWorker` (BackgroundService, bounded concurrency, reclaim po crashu). Rejestracja w DI + `DeliveryWorkerOptions` (sekcja `DeliveryWorker`).
- DB: `infra/sql/007_add_document_deliveries.sql` (tabela + indeksy + unique partial idempotencji + check status). Mapowanie EF `DocumentDeliveryConfiguration` + `DbSet`.
- API (`DocumentStorageController`): `POST {masterId}/versions/{versionId}/finish` (202), `GET deliveries/{id}`, `GET deliveries?status=`, `POST deliveries/{id}/retry`.
- GUI: `finishDocument()` (utrwala stan + wysyłka + polling co 4 s do stanu końcowego), serwis `finishAndSend`/`getDeliveryStatus`, sygnały `isFinishing/deliveryStatus/deliveryId`.

### Verified
- `dotnet build D2ViewerEditor.sln` — OK (0 błędów).
- `dotnet test` — nowe testy (DocumentDelivery, handler FinishAndSend, validator, backoff) zielone. Jeden WCZEŚNIEJSZY błąd niezwiązany: `SaveDocumentVersionCommandHandlerTests.Handle_ValidCommand_ShouldAddNewVersion` (oczekuje `UpdateAsync`, handler go nie woła po przejściu na płaski zapis) — nie ruszany.
- GUI: TypeScript kompiluje się; `npm run build` blokuje WCZEŚNIEJSZY budżet `document-editor.scss` (35.7 kB > 32 kB), niezwiązany z tą zmianą.

### Notes
- Snapshot finalny zamrażany w GCS jako `deliveries/{deliveryId}` (nie wysyłamy później zmodyfikowanej v2). At-least-once + Idempotency-Key. Worker odporny na restart/multi-instancję; `DeliveryWorker:Enabled=false` pozwala wydzielić wysyłkę na dedykowany host.

## 2026-05-25 — Pionowa linijka per strona + lupa otwiera panel

### Changed
- Pionowa linijka: zamiast jednej „zamrożonej" renderowana jest OSOBNA linijka na każdą stronę (`pageList`), z geometrią kartek (segment = wysokość strony*zoom, odstęp = separator 8px*zoom). Pasek scrolluje 1:1, więc linijka restartuje się na granicy stron. Segment ma `overflow:hidden` (ucina ~1px nadmiaru axisPx 1123 vs strona 1122).
- Lupa w toolbarze otwiera panel „Wyszukiwanie" (nowy output `openSearch`), usunięto stary pasek wyszukiwania toolbaru (`showSearchBar`).

### Verified
- `ng build` OK.

### Notes
- `onEditorScroll` używa przybliżonego `PAGE_GAP=40` do wskaźnika „Strona X z Y" (realny separator to 8px) — nietykane, dotyczy tylko zaokrąglenia numeru strony, nie linijki.

## 2026-05-25 — Fix 400 /save (EF) + panel „Wyszukiwanie" zamiast dialogu

### Changed
- BACKEND FIX: `DocumentVersionConfiguration` — `Id` ma `ValueGeneratedNever()`. Bez tego EF (konwencja ValueGeneratedOnAdd dla Guid) traktował nową wersję z ręcznie ustawionym kluczem dodaną do śledzonego dokumentu jako `Modified` → UPDATE w 0 wierszy → `DbUpdateConcurrencyException` → POST `/{master}/save` zwracał 400. To blokowało tworzenie wersji edytowalnej przy ręcznym otwarciu z dysku/dashboardu. (Zmiana tylko w modelu EF, bez migracji schematu.) **Wymaga restartu API.**
- Dialog Znajdź/zamień → lewy panel `.search-panel` „Wyszukiwanie" (jak panel Nawigacja w Word): tylko wyszukiwanie (bez zamiany), bez zakładek Nagłówki/Strony — tylko sekcja „Wyniki". Lista wyników z fragmentami kontekstu; klik = skok do trafienia (`goToResult`→`editor.goToMatch`). Nowe API edytora: `getSearchSnippets()`, `goToMatch(index)`.
- EDYTUJ → pozycja „Znajdź" (Ctrl+F) w obu trybach.

### Verified
- Frontend `ng build` OK; backend `dotnet build` Infrastructure OK. Przyczynę 400 potwierdzono w logach API (DbUpdateConcurrencyException, UPDATE nowej wersji zamiast INSERT).

## 2026-05-25 — Wyszukiwanie po wszystkich stronach + naprawa dialogu Znajdź

### Changed
- `WysiwygEditorComponent.searchText` przeszukuje teraz WSZYSTKIE strony (`pageEditorRefs`) w kolejności dokumentu, nie tylko aktywną; agreguje trafienia do jednej listy `searchHighlights` (findNext/findPrevious nawigują między stronami, `scrollIntoView` przewija). Dotyczy też wyszukiwarki z toolbaru.
- `emitContentChange` po zamianie agreguje treść ze wszystkich stron (`getContent()`), więc zamiana na nieaktywnej stronie jest utrwalana.
- Dialog Znajdź (`DocumentEditorComponent`) przepięty z ułomnego `window.find()` na API edytora: wyszukiwanie na żywo (`onFindInput`), licznik „x z y", przyciski „Poprzednie/Następne", czyszczenie podświetleń przy zamknięciu (`closeFindReplace`).

### Verified
- `ng build --configuration development` OK.

## 2026-05-25 — Prowadnice marginesów + wyszukiwanie w read-only

### Changed
- „Linie marginesów" (martwy przełącznik — styl bez markupu) zaimplementowane: `wysiwyg-editor.html` renderuje `.margin-guides` (4 przerywane linie) per strona z `pageMargins()`, sterowane `[showMarginGuides]`.
- Read-only: cały `d2-editor-toolbar` ukryty (znika lupa). Wyszukiwanie przez EDYTUJ → „Znajdź" (zamiast „Znajdź i zamień") oraz skrót Ctrl+F (`onGlobalKeydown` w `DocumentEditorComponent`; Ctrl+H tylko gdy edycja dozwolona).
- Dialog `showFindReplace`: w trybie zablokowanym tylko pole „Znajdź" + „Znajdź następny" (ukryte „Zamień na"/„Zamień"/„Zamień wszystko"); nagłówek „Znajdź".

### Verified
- `ng build --configuration development` OK.

### Notes
- `showMarginGuides` domyślnie `true` → prowadnice widoczne domyślnie (można wyłączyć w WIDOK/FORMATUJ). Renderu nie potwierdzałem wizualnie.

## 2026-05-25 — Ukrywanie funkcji edycyjnych w trybie read-only / „zajęty"

### Changed
- `DocumentEditorComponent`: nowy `lockedByOther` (hook na backendowy `DocumentStatus`=Editing, domyślnie false) i computed `editingDisabled = readOnly() || lockedByOther()`.
- `EditorToolbarComponent`: nowy `@Input() readOnly` — przy true ukrywa undo/redo i wszystkie grupy edycyjne (styl, czcionka, format, wyrównanie, listy, wstawianie); zostaje wyszukiwarka.
- Menu edytora: ukryte WSTAW/FORMATUJ/NARZĘDZIA w całości; w PLIK ukryte Zapisz + Ustawienia strony; w EDYTUJ ukryte cofnij/ponów/wytnij/wklej/usuń (zostają Kopiuj, Zaznacz wszystko, Znajdź i zamień, Właściwości, Podpisy). WIDOK bez zmian. Toolbar tabeli ukryty (`!editingDisabled()`). Plakietka rozróżnia „ktoś inny edytuje".

### Verified
- `ng build --configuration development` OK.

### Notes
- Blokada „ktoś inny edytuje" wymaga podpięcia statusu dokumentu do edytora (`lockedByOther.set(...)`); obecnie edytor zna tylko read-only z braku `versionId`. Skróty klawiszowe nie są blokowane — chroni je `readOnly` na contenteditable.

## 2026-05-25 — Linijka: linia prowadząca + obrazowanie nagłówka/stopki

### Changed
- `RulerComponent` emituje `dragGuideChange` (active + axis + offset px @100% od krawędzi strony). Linia prowadząca (jak w MS Word) renderowana w `.paper-container` edytora, bo wewnątrz `.ruler-h-bar` (22px, `overflow:hidden`) była przycinana.
- `WysiwygEditorComponent` emituje `editingSectionChange` ('header'|'footer'|'body').
- `DocumentEditorComponent`: `verticalRulerMargins` przełącza obrazowanie pionowej linijki na pasmo nagłówka/stopki podczas ich edycji.
- Pasmo nagłówka/stopki ma `min-height` i rośnie z treścią (obraz), więc położenie na linijce pochodzi z POMIARU DOM: `WysiwygEditorComponent.sectionGeometryChange` (cm od góry strony, `ResizeObserver` re-emituje przy zmianie wysokości; skala liczona z szerokości strony — odporna na wzrost pasma).
- Pionowa linijka w trybie nagłówka/stopki: przeciągnięcie uchwytu zmienia WYSOKOŚĆ pasma (`setHeaderHeight`/`setFooterHeight`), nie marginesy strony. W trybie `body` działa jak wcześniej.

### Verified
- `ng build --configuration development` OK (bez błędów).

### Notes
- Marginesy góra/dół: ekstrakcja w `DocxToHtmlConverter` poprawna (1/567 ≈ 2.54/1440). Treść startuje na `marginTop` (header band + padding). Rozbieżność z Wordem dotyczy raczej POZYCJI pasma nagłówka (renderowane przy `top:0`, nie na „header from edge"=pgMar.Header), nie wysokości marginesu treści. Do potwierdzenia na konkretnym pliku DOCX.

## 2026-05-23 — Dostosowanie `.ai/` do realnego projektu

### Changed
- Przyjęto strukturę `.ai/` wg `template/` (INDEX, PROJECT_CONTEXT, TECH_STACK, ARCHITECTURE, DOMAIN, FEATURES, DATABASE, API_CONTRACTS, SECURITY, TESTING_QUALITY, DEVOPS_DEPLOYMENT, DECISIONS, RISKS_ASSUMPTIONS, GLOSSARY, CURRENT_STATE, TASK_HANDOFF, CHANGELOG, PRODUCT_GOALS + pliki-wytyczne i `templates/`).
- Wypełniono faktami z repo; oznaczono niepewności (port External API, brak compose/CI).
- Dodano `CLAUDE.md` i `AGENTS.md` w root.
- Usunięto stare `.ai/{CONTEXT,API,BACKEND,FRONTEND}.md` (treść przeniesiona do nowej struktury).

### Verified
- Pliki zapisane; fakty zebrane z `*.csproj`, `package.json`, `docker/*`, `launchSettings`, `appsettings`, `infra/sql/`.

### Notes
- Pliki ogólne (BACKEND_DOTNET, FRONTEND_ANGULAR, CODING_STANDARDS, AI_AGENT_WORKFLOW, PROMPTS, README, templates) pozostawiono zgodne ze wzorcem.

## 2026-05-23 — Płaski zapis wersji + auto-save + ujednolicony „Zapisz"

### Changed
- Backend: `Document.UpdateVersion`, `DocumentVersion.UpdateContent`/`ModifiedAt`, `UpdateDocumentVersionCommand` + `PUT /api/documentstorage/{masterId}/versions/{versionId}`, `GetDocumentMetadataQuery` + `GET .../{masterId}/metadata`, skrypt `infra/sql/005_add_version_modified_at.sql` + mapowanie EF.
- Frontend: mechanizm auto-save (timer + `environment.autoSave`), switch „AutoSave", ujednolicony `saveDocument()` (API), `downloadDocument()` (pobranie lokalne), usunięto `saveDocumentAs()`.

### Verified
- `dotnet build D2ViewerEditor.sln` — OK (0 błędów).
- `npm run build` (GUI) — OK.

### Notes
- Tryb podglądu (Krok 2) wciąż ładuje aktywną wersję zamiast v1 — do dokończenia.
- `finishDocument()` (zwrot na returnUrl) — TODO.
