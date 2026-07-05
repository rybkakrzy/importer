# Audyt zgodności z Microsoft Word — edytor D2 ViewerEditor

> Data: 2026-07-05. Zakres: pełny pipeline DOCX → Open XML → model → HTML/CSS → edycja →
> zapis → eksport DOCX. Analiza kodu bez zmian w implementacji.
>
> Źródła zweryfikowane w kodzie: `DocxToHtmlConverter.cs` (4138 l.), `HtmlToDocxConverter.cs`
> (3341 l.), `wysiwyg-editor.ts` (5172 l.), handlery Save/Sign/Finish/Download, model
> `DocumentContent`. Dokumenty referencyjne: `DOCX_CONVERSION.md`, `FIDELITY_REPORT.md`,
> `EDITOR_KEYBOARD.md`, `RISKS_ASSUMPTIONS.md`.
>
> Oznaczenia: **[POTWIERDZONE]** = zweryfikowane w kodzie (podane miejsce);
> **[POTENCJALNE]** = wynika z braku obsługi w kodzie / wymaga dokumentu testowego do potwierdzenia objawu.

## 0. Architektura problemu — ustalenie nadrzędne

Wszystkie ścieżki zapisu poza ręcznym „Pobierz dokument" używają **pełnej regeneracji pakietu
DOCX od zera** (`IHtmlToDocxConverter.Convert`):

| Ścieżka | Konwersja | Skutek |
|---|---|---|
| Auto-save / „Zapisz" (v2 nadpisywana w miejscu) | `SaveDocumentCommandHandler` → `Convert` **[POTWIERDZONE]** | pakiet zregenerowany; wszystko, czego reader nie przeniósł do HTML, znika z v2 **przy pierwszym autosave** |
| „Zakończ i wyślij" (plik do aplikacji zewnętrznej) | GUI `document-editor.ts:2024` → `documentService.saveDocument` → `Convert` **[POTWIERDZONE]** | plik dostarczany na zewnątrz = najgorsza ścieżka wierności |
| Podpis (Sign) | `SignDocumentCommandHandler:23` → `Convert` **bez** `SectionHeadersFooters` **[POTWIERDZONE]** | podpisany plik dodatkowo gubi sekcyjne nagłówki/stopki |
| „Pobierz dokument" | `DownloadEditedDocumentCommandHandler:72` → `ConvertPreservingPackage` (styles.xml + theme + fontTable z v1) | jedyna ścieżka z częściowym pass-through |

Reader ma przy tym **cichy fallback „drop"**: nieznany element → pusty string
(`ConvertElementToHtml` — tylko Paragraph/Table/SdtBlock; `ConvertRunChildToHtml` — `default:
return string.Empty`; pętla dzieci akapitu — tylko Run/Hyperlink/SimpleField/SdtRun)
**[POTWIERDZONE: DocxToHtmlConverter.cs:1527-1536, 2605-2632, 1044-1061]**.

Kombinacja „cichy drop w readerze" + „regeneracja pakietu w autosave" oznacza: **każdy
nieobsługiwany element DOCX to nie tylko brak renderu — to trwała utrata danych w v2 po
pierwszym autosave**, bez ostrzeżenia użytkownika. Oryginał v1 jest jedyną kopią bezpieczeństwa
(immutable — to działa i jest kluczową mitygacją).

---

## 1. Problemy KRYTYCZNE — utrata danych / uszkodzenie dokumentu

#### KR-01 · Regeneracja pakietu DOCX w autosave i „Zakończ i wyślij"
- **Kategoria:** round-trip / architektura zapisu
- **Różnica vs Word:** Word przy zapisie zachowuje cały pakiet OOXML; edytor buduje nowy pakiet z HTML.
- **Reprodukcja:** otwórz DOCX z niestandardowymi stylami/motywem/przypisami → poczekaj na autosave → pobierz v2 albo wyślij „Zakończ".
- **Oczekiwane:** części pakietu nieobjęte edycją pozostają bez zmian.
- **Obecne:** styles.xml (hardkod Normal/Heading1-6/Header/Footer/ListParagraph, `pl-PL`, spacing 160/259 — `HtmlToDocxConverter.AddDocumentStyles:672`), theme, fontTable, settings.xml, numbering.xml, footnotes/endnotes/comments — regenerowane lub pomijane. **[POTWIERDZONE]**
- **Przyczyna:** `Convert` tworzy `WordprocessingDocument` od zera; `ConvertPreservingPackage` wpięty tylko w Download.
- **Warstwa:** backend .NET / eksport
- **OOXML:** cały pakiet (parts + relacje)
- **Ryzyko:** krytyczne · **Wpływ:** utrata danych, round-trip
- **Kierunek:** rozszerzyć pass-through na ścieżkę Save/Finish (R-20); docelowo edycja in-place body w oryginalnym pakiecie zamiast pełnej regeneracji.
- **Odwzorowanie w HTML:** n/d (problem eksportu)
- **Testy:** round-trip realnych dokumentów z porównaniem listy parts + diff styles.xml/settings.xml.

#### KR-02 · Przypisy dolne i końcowe — całkowity brak obsługi
- **Kategoria:** pola i elementy dynamiczne
- **Różnica vs Word:** przypisy niewidoczne w edytorze; po autosave znikają z dokumentu.
- **Reprodukcja:** DOCX z przypisem → otwórz → brak odnośnika i treści; autosave → v2 bez footnotes.xml.
- **Obecne:** `FootnoteReference`/`EndnoteReference` wpadają w `default: string.Empty`; part footnotes/endnotes nigdy nieczytany; writer nie ma żadnej ścieżki przypisów. **[POTWIERDZONE — brak symboli w obu konwerterach]**
- **Przyczyna:** brak implementacji.
- **Warstwa:** import + eksport
- **OOXML:** `w:footnoteReference`, `w:endnoteReference`, parts `footnotes.xml`/`endnotes.xml`
- **Ryzyko:** krytyczne · **Wpływ:** utrata danych (treść przypisów!)
- **Kierunek:** minimum = round-trip bez renderu (zachować referencje + part); docelowo render `<sup>` + panel przypisów.
- **Odwzorowanie w HTML:** przybliżone (sup + lista na dole strony); paginacyjne umieszczenie na właściwej stronie — trudne.
- **Testy:** golden DOCX z przypisami, asercja przeżycia part + referencji po round-trip.

#### KR-03 · Komentarze — całkowity brak obsługi
- **Kategoria:** recenzja i współpraca
- **Różnica vs Word:** komentarze i zakresy znikają.
- **Obecne:** `CommentRangeStart/End`, `CommentReference` niedopasowane w żadnym switchu → drop; part comments.xml nieczytany. **[POTWIERDZONE]**
- **Warstwa:** import + eksport · **OOXML:** `w:commentRangeStart/End`, `w:commentReference`, `comments.xml`
- **Ryzyko:** krytyczne (dla dokumentów recenzowanych) · **Wpływ:** utrata danych
- **Kierunek:** minimum round-trip pass-through zakresów; render opcjonalny.
- **Testy:** round-trip z komentarzem wielorunowym i odpowiedzią.

#### KR-04 · Śledzenie zmian — wstawiony tekst ZNIKA z treści
- **Kategoria:** recenzja i współpraca
- **Różnica vs Word:** Word pokazuje `w:ins` jako treść (z adiustacją); edytor go w ogóle nie renderuje.
- **Reprodukcja:** DOCX z niezaakceptowanymi zmianami (wstawienia) → otwórz w edytorze → wstawiony tekst **nieobecny**; autosave → tekst trwale usunięty z v2, `w:del` znika bez śladu (kasowanie „zatwierdzone" bez decyzji użytkownika).
- **Obecne:** `InsertedRun` (`w:ins`) i `DeletedRun` (`w:del`) nie są dziećmi typu `Run` — pętle po dzieciach akapitu je pomijają. **[POTWIERDZONE — brak `Inserted`/`Deleted` w readerze]**
- **Przyczyna:** switch po typach bez gałęzi rewizji; brak `AcceptRevisions` przed konwersją.
- **Warstwa:** import · **OOXML:** `w:ins`, `w:del`, `w:rPrChange`, `w:pPrChange`, `w:moveFrom/To`
- **Ryzyko:** krytyczne · **Wpływ:** utrata danych (fragmenty treści!)
- **Kierunek:** krok 1 (szybki): traktować `w:ins` jak kontener runów (renderować treść), `w:del` pomijać = zachowanie „zaakceptuj wszystko"; jawnie to komunikować. Krok 2: prawdziwy round-trip rewizji.
- **Testy:** golden z ins/del/rPrChange; asercja obecności tekstu z `w:ins` w HTML.

#### KR-05 · Pola złożone inne niż PAGE/NUMPAGES/DATE tracą kod ORAZ wartość
- **Kategoria:** pola i elementy dynamiczne
- **Różnica vs Word:** Word pokazuje ostatnią obliczoną wartość pola i zachowuje instrukcję; edytor gubi obie.
- **Reprodukcja:** DOCX ze spisem treści (TOC) albo `REF`/`MERGEFIELD` → otwórz → treść pola (np. cały spis treści) nieobecna; autosave → trwała utrata.
- **Obecne:** w `FieldChar Separate` emitowany placeholder tylko dla PAGE/NUMPAGES/SECTIONPAGES/DATE/TIME; dla innych instrukcji nic. Po separatorze wszystkie runy wartości są **pomijane** (`if (fieldSeparated) continue;`). **[POTWIERDZONE: DocxToHtmlConverter.cs:1895-1956]**
- **Przyczyna:** whitelist instrukcji + skip cached result.
- **Warstwa:** import · **OOXML:** `w:fldChar`, `w:instrText`, `w:fldSimple` (TOC, REF, SEQ, MERGEFIELD, HYPERLINK, FORM*)
- **Ryzyko:** krytyczne · **Wpływ:** utrata danych, funkcjonalny
- **Kierunek:** dla nieznanej instrukcji emitować **wartość cache jako zwykły tekst** (zachowanie treści) + `data-field-instr` do round-tripu; TOC docelowo jako blok read-only.
- **Odwzorowanie w HTML:** wartość — pełne; dynamika pól — tylko round-trip.
- **Testy:** golden z TOC/REF/SEQ; asercja, że tekst wartości przeżywa import i eksport.

#### KR-06 · Kształty, pola tekstowe, WordArt, SmartArt, wykresy, OLE, równania — drop
- **Kategoria:** grafika / elementy osadzone
- **Różnica vs Word:** Word renderuje; edytor nie pokazuje nic i trwale usuwa przy autosave.
- **Reprodukcja:** DOCX z polem tekstowym (`wps:txbx`) z treścią → otwórz → brak; autosave → **tekst z pola tekstowego bezpowrotnie stracony**.
- **Obecne:** `ConvertDrawingToHtml` wymaga `a:blip` — `if (blip?.Embed?.Value == null) return string.Empty;` **[POTWIERDZONE: DocxToHtmlConverter.cs:3001-3002]**. Wykres/SmartArt/kształt/txbx nie mają blipa → drop. `m:oMath` nigdzie nieobsłużony. `mc:AlternateContent` nie ma dedykowanej obsługi.
- **Warstwa:** import · **OOXML:** `wps:wsp`, `wps:txbx`, `a:graphicData` (chart/diagram), `w:object`/OLE, `m:oMath`, `mc:AlternateContent`
- **Ryzyko:** krytyczne (txbx/równania = treść), wysokie (wykresy) · **Wpływ:** utrata danych, wizualny
- **Kierunek:** (1) tekst z `wps:txbx` renderować jako blok/absolutnie pozycjonowany div; (2) chart/SmartArt → rasteryzacja lub placeholder z pass-through partu; (3) oMath → MathML/obraz + pass-through.
- **Odwzorowanie w HTML:** txbx — przybliżone; chart/SmartArt — tylko obraz statyczny; oMath — MathML częściowo.
- **Testy:** golden z txbx/chart/oMath; asercja nie-utraty treści tekstowej.

#### KR-07 · Zakładki (bookmarks) tracone
- **Kategoria:** pola / struktura
- **Różnica vs Word:** cele odwołań i hiperłączy wewnętrznych przestają istnieć.
- **Obecne:** `BookmarkStart/End` niedopasowane → drop. **[POTWIERDZONE — brak w readerze]** Hiperłącze wewnętrzne (`w:hyperlink w:anchor`) po eksporcie wskazuje nieistniejącą zakładkę.
- **Warstwa:** import + eksport · **OOXML:** `w:bookmarkStart/End`, `w:hyperlink/@w:anchor`
- **Ryzyko:** wysokie · **Wpływ:** utrata danych, funkcjonalny (REF/TOC/odsyłacze przestają działać w Wordzie)
- **Kierunek:** mapować na `<a id>`/`<span data-bookmark>` — tani pełny round-trip.
- **Testy:** round-trip bookmark + hyperlink anchor.

#### KR-08 · Pole DATE zamrażane datą serwera
- **Kategoria:** pola dynamiczne
- **Różnica vs Word:** Word pokazuje wartość cache i aktualizuje przy druku; edytor podstawia `DateTime.Now` serwera w formacie `dd.MM.yyyy`, ignorując format z instrukcji (`\@ "..."`) i wartość zapisaną. **[POTWIERDZONE: DocxToHtmlConverter.cs:1918-1921]**
- **Reprodukcja:** dokument z `DATE \@ "d MMMM yyyy"` z 2024 r. → otwórz → dzisiejsza data w innym formacie; po eksporcie pole traci semantykę.
- **Warstwa:** import · **OOXML:** `w:fldSimple`/`w:instrText` DATE/TIME/CREATEDATE/SAVEDATE
- **Ryzyko:** wysokie (dokumenty umowne!) · **Wpływ:** utrata danych, strukturalny
- **Kierunek:** czytać wartość cache z pola (run po Separate), nie generować własnej; zachować instrukcję w data-*.
- **Testy:** golden DATE z formatem niestandardowym.

#### KR-09 · Układ wielokolumnowy (w:cols) nieobsługiwany
- **Kategoria:** układ strony
- **Obecne:** `w:cols` nieczytane (brak `Columns` w readerze poza tabelami) → treść dwukolumnowa renderuje się jednokolumnowo; po autosave sectPr bez `w:cols` (marker sekcji niesie tylko geometrię strony/marginesy). **[POTWIERDZONE — brak w kodzie]**
- **Warstwa:** import + eksport + frontend · **OOXML:** `w:cols` (`num`, `space`, `equalWidth`, `w:col`)
- **Ryzyko:** wysokie · **Wpływ:** wizualny, strukturalny, round-trip
- **Kierunek:** round-trip w data-* markera sekcji (tanio); render CSS `column-count` (przybliżenie — brak łamania kolumn jak Word).
- **Testy:** round-trip sectPr z cols; wizualny test 2 kolumn.

#### KR-10 · Numeracja stron sekcji (pgNumType) tracona
- **Kategoria:** układ strony / pola
- **Obecne:** `w:pgNumType` (start, format rzymski/literowy, chapter) nieczytany/niezapisywany. **[POTWIERDZONE — brak w kodzie]** Placeholder `{page}` w edytorze zawsze liczy od 1 arabsko.
- **Ryzyko:** średnie · **Wpływ:** wizualny, round-trip
- **Kierunek:** data-* na markerze sekcji + uwzględnienie w `_pageNumberHtml`.

#### KR-11 · Ochrona dokumentu i ograniczenia edycji ignorowane i tracone
- **Kategoria:** recenzja / bezpieczeństwo dokumentu
- **Obecne:** `w:documentProtection`, `w:permStart/End` (settings.xml regenerowany, permy dropowane) — edytor pozwala edytować wszystko; po zapisie ochrona znika. **[POTWIERDZONE — settings.xml nie jest czytany]**
- **Ryzyko:** wysokie (dokumenty z ograniczeniami formalnymi) · **Wpływ:** strukturalny, funkcjonalny
- **Kierunek:** minimum: wykrywać ochronę przy imporcie i ostrzegać/blokować; pass-through settings.xml.

#### KR-12 · Ukryty tekst (w:vanish) renderowany jak zwykły i tracony jako atrybut
- **Kategoria:** formatowanie znaków
- **Obecne:** `w:vanish` nieczytany **[POTWIERDZONE — brak]** → tekst ukryty w Wordzie jest widoczny w edytorze; po eksporcie atrybut znika → tekst staje się trwale widoczny także w Wordzie.
- **Ryzyko:** wysokie (treści robocze/instrukcje szablonów stają się widoczne w dokumencie wysyłanym!) · **Wpływ:** utrata danych (atrybutu), wizualny
- **Kierunek:** czytać → `display:none`+data-* (lub styl „ukryty" z podglądem); writer odtwarza `w:vanish`.

#### KR-13 · Znak wodny / tło dokumentu / kształty w nagłówku — tracone
- **Kategoria:** grafika / nagłówki
- **Obecne:** watermark to VML shape (WordArt) w nagłówku — `ConvertPictureToHtml` obsługuje tylko `ImageData` (obraz); tekstowy watermark ginie. `w:background` dokumentu nieczytany. **[POTWIERDZONE dla background; watermark POTENCJALNE — zależnie od postaci VML]**
- **Ryzyko:** średnie · **Wpływ:** wizualny, round-trip

#### KR-14 · Podpisywanie dokumentu gubi sekcyjne nagłówki/stopki
- **Kategoria:** eksport
- **Obecne:** `SignDocumentCommandHandler` woła `Convert` bez parametru `SectionHeadersFooters` **[POTWIERDZONE: SignDocumentCommandHandler.cs:23]** — dokument wielosekcyjny po podpisaniu ma nagłówki tylko z kompletu głównego.
- **Ryzyko:** wysokie · **Wpływ:** utrata danych, round-trip
- **Kierunek:** przekazać sekcyjne nagłówki jak w Save (znane ograniczenie ADR-0025 — domknąć).

---

## 2. Problemy STRUKTURALNE

#### ST-01 · Style nazwane spłaszczane do formatowania bezpośredniego
- **Kategoria:** style i dziedziczenie
- **Różnica vs Word:** po round-trip akapity nie odwołują się do oryginalnych stylów; zmiana definicji stylu w Wordzie po edycji nie działa; panel stylów pokazuje „Normal+direct".
- **Obecne:** reader rozwiązuje `basedOn` do **inline CSS** (sklejanie stringów + `DeduplicateCss` regex); writer emituje styles.xml z hardkodem i mapuje tylko h1-h6 → `Heading1..6` (z ustalonymi kolorami `2F5496`). `next`/`link`/`w:styleId` inne niż heading — tracone. **[POTWIERDZONE: LoadDocumentStyles / AddDocumentStyles:672-770]**
- **Warstwa:** import + eksport · **OOXML:** `w:styles`, `w:pStyle`, `w:rStyle`, `basedOn`, `next`, `link`
- **Ryzyko:** wysokie · **Wpływ:** strukturalny, round-trip
- **Kierunek:** nieść `data-style-id` na blokach (jak zrobiono dla tabel `data-tbl-style`) + pass-through styles.xml w Save; writer emituje `w:pStyle` zamiast pełnego inline.
- **Testy:** round-trip zachowania `w:pStyle` na akapicie stylowanym.

#### ST-02 · `w:pageBreakBefore` zamieniany na twardy `w:br type=page`
- **Kategoria:** akapity / paginacja
- **Obecne:** reader emituje `div.page-break` dla `pageBreakBefore` **[POTWIERDZONE: DocxToHtmlConverter.cs:1613-1618]**; writer mapuje div → `w:br` (`CreatePageBreak`). Właściwość akapitu staje się osobnym twardym podziałem — przy dalszej edycji w Wordzie zachowuje się inaczej (break nie „podąża" za akapitem).
- **Ryzyko:** średnie · **Wpływ:** strukturalny, round-trip
- **Kierunek:** znacznik `data-page-break-before` na akapicie zamiast osobnego diva.

#### ST-03 · Marker sekcji jest zwykłym elementem contenteditable — kasowalny
- **Kategoria:** sekcje / edycja
- **Obecne:** użytkownik może usunąć `div.docx-section-break` (Backspace/zaznaczenie) → dokument degraduje do 1 sekcji przy zapisie. **[POTWIERDZONE — udokumentowane ograniczenie ADR-0023]**
- **Ryzyko:** wysokie · **Wpływ:** utrata danych (geometria sekcji, nagłówki), round-trip
- **Kierunek:** `contenteditable=false` na markerze + ochrona w handlerach Backspace/Delete + test.

#### ST-04 · Formanty SDT — treść przenoszona, semantyka formantu uproszczona
- **Kategoria:** pola formularzy
- **Obecne:** SdtBlock/SdtRun konwertowane (`ConvertSdtBlockToHtml/ConvertSdtRunToHtml`, writer `BuildSdtBlockFromHtml/BuildSdtRunFromHtml` + `BuildSdtProperties`). **[POTWIERDZONE — istnieje]** **[POTENCJALNE]:** typy formantów (dropdown listy pozycji, date picker z formatem, checkbox stany, placeholder, lock) prawdopodobnie nie round-tripują w pełni — `BuildSdtProperties` buduje właściwości z okrojonych data-*.
- **Ryzyko:** średnie · **Wpływ:** funkcjonalny, round-trip
- **Testy:** round-trip każdego typu SDT (plain/rich/combo/dropdown/date/checkbox) — obecnie brak.

#### ST-05 · Writer przetasowuje strukturę dla nietypowego HTML
- **Kategoria:** eksport
- **Obecne:** luźny tekst na poziomie body → osobny akapit; `span`/`b`/... na poziomie body → nowy akapit (`ConvertInlineElement`); nieznane tagi → spłaszczenie dzieci. **[POTWIERDZONE: HtmlToDocxConverter.cs:895-1006]** Po nieczystym paste/edycji może przybywać akapitów.
- **Ryzyko:** średnie · **Wpływ:** strukturalny

#### ST-06 · Listy: nowe listy bez tożsamości, sufiks lvlText nierenderowany
- **Kategoria:** listy
- **Obecne (po fixie 2026-07-05):** tożsamość `numId/abstractNumId` round-tripuje przez data-*; ALE (a) listy tworzone w edytorze dostają domyślną drabinkę formatów (decimal/lowerLetter/lowerRoman) niezależnie od konwencji dokumentu; (b) `lvlText` sufiks („1)", „Art. %1.") i numeracja złożona „1.1.2" nie są renderowane wizualnie (wracają do DOCX); (c) numbering.xml nie jest pass-through (R-19) — picture bullets i niestandardowe poziomy odtwarzane z data-*, reszta definicji ginie. **[POTWIERDZONE — CURRENT_STATE + kod]**
- **Ryzyko:** średnie · **Wpływ:** wizualny, round-trip
- **Kierunek:** render markerów listy po stronie edytora z `data-lvl-text` (CSS counters); pass-through numbering.xml z remapem numId.

#### ST-07 · Puste akapity wypełniane `&nbsp;`
- **Kategoria:** akapity
- **Obecne:** reader emituje `&nbsp;` w pustych akapitach/li **[POTWIERDZONE: 1063-1066]**; writer dekoduje InnerText → NBSP może trafić do `w:t` jako realny znak. **[POTENCJALNE]:** po kilku round-tripach puste akapity zawierają U+00A0 (różnica przy zaznaczaniu/końcu wiersza w Wordzie).
- **Ryzyko:** niskie · **Wpływ:** strukturalny (drift)
- **Testy:** round-trip pustego akapitu ×2, asercja braku NBSP w `w:t`.

#### ST-08 · Ustawienia dokumentu (settings.xml) w całości tracone
- **Kategoria:** ustawienia wpływające na render
- **Obecne:** `w:settings` nieczytane i regenerowane: `defaultTabStop`, `autoHyphenation`, `mirrorMargins`, `gutter`, `evenAndOddHeaders` (odtwarzany tylko z modelu HeaderFooter), `compat` (tryb zgodności — wpływa na layout w Wordzie!), `docVars`. **[POTWIERDZONE — brak odczytu settings]**
- **Ryzyko:** średnie · **Wpływ:** round-trip, wizualny (compat potrafi zmienić łamanie w Wordzie)
- **Kierunek:** pass-through settings.xml w `ConvertPreservingPackage` i docelowo w Save.

---

## 3. PAGINACJA i UKŁAD STRONY

#### PG-01 · Silnik paginacji ≠ Word — inna liczba stron i miejsca łamania
- **Kategoria:** paginacja
- **Obecne:** paginacja frontendowa mierzy bloki w px (96 DPI, 37.8 px/cm — `_repaginateNow` **[POTWIERDZONE: wysiwyg-editor.ts:3461-3627]**), na metrykach fontów przeglądarki. Word łamie po liniach z własnymi metrykami. Ta sama treść daje inne granice stron; `{page}/{pages}` w edytorze ≠ numeracja wydruku Worda.
- **Ryzyko:** wysokie (percepcja klienta) · **Wpływ:** wizualny, funkcjonalny
- **Odwzorowanie w HTML:** pełna zgodność **niemożliwa**; cel = minimalizacja rozjazdu (fonty, marginesy, interlinia).
- **Testy:** porównawcze liczby stron dla korpusu realnych dokumentów (harness fidelity).

#### PG-02 · Akapit nigdy nie jest dzielony między strony
- **Kategoria:** paginacja
- **Obecne:** granulacja blokowa — cały akapit przechodzi na następną stronę (`pushMeasured`); blok wyższy niż strona zostaje i przepełnia (overflow). Word dzieli akapit liniami. **[POTWIERDZONE: 3536-3542]**
- **Reprodukcja:** akapit 40-liniowy na końcu strony → w edytorze cała kartka pusta + akapit od nowej strony (lub przepełnienie); w Wordzie podział w środku akapitu.
- **Ryzyko:** wysokie · **Wpływ:** wizualny, funkcjonalny (duże dokumenty prawnicze)
- **Kierunek:** splitting akapitu po liniach (Range.getClientRects) — złożone, ale wykonalne.

#### PG-03 · keepNext / keepLines / widowControl nieobsługiwane
- **Kategoria:** akapity/paginacja
- **Obecne:** właściwości nieczytane w ogóle **[POTWIERDZONE — brak w readerze]** → (a) nie honorowane przy paginacji edytora, (b) **tracone przy eksporcie** → Word też przestaje je stosować.
- **Ryzyko:** wysokie (b = trwała utrata) · **Wpływ:** utrata danych (atrybuty), wizualny
- **Kierunek:** minimum: round-trip przez data-*; opcjonalnie honorowanie keepNext w `_repaginateNow` (przesuwaj razem z następnym blokiem).

#### PG-04 · Wiersz nagłówkowy tabeli i cantSplit nie działają przy łamaniu stron
- **Obecne:** `tblHeader`/`cantSplit` round-tripują w data-*, ale `_splitTableForPagination` nie powtarza nagłówka na kolejnych stronach ani nie chroni wiersza przed podziałem. **[POTWIERDZONE — udokumentowane ADR-0026]**
- **Ryzyko:** średnie · **Wpływ:** wizualny

#### PG-05 · Wysokość nagłówka/stopki obcinana do 0.8–8 cm
- **Obecne:** `Math.Max(0.8, Math.Min(8, height))` **[POTWIERDZONE: DocxToHtmlConverter.cs:356,426,497,539]** — dokument z nagłówkiem ~0 lub bardzo wysokim renderuje się inaczej niż w Wordzie.
- **Ryzyko:** niskie · **Wpływ:** wizualny

#### PG-06 · Pola PAGE/NUMPAGES = placeholder, wartości z paginacji edytora
- **Obecne:** wartości liczone przez edytor (inna paginacja → inne liczby niż wydruk Worda); brak obsługi formatów (`\* ROMAN`), `SECTIONPAGES` zlewane z NUMPAGES. **[POTWIERDZONE: 1910-1916]**
- **Ryzyko:** średnie · **Wpływ:** wizualny, funkcjonalny

#### PG-07 · Pionowe wyrównanie strony (sectPr vAlign) brak
- **[POTWIERDZONE — brak]** Strona tytułowa wyśrodkowana pionowo w Wordzie renderuje się od góry. Atrybut tracony na eksporcie.
- **Ryzyko:** niskie · **Wpływ:** wizualny, round-trip

#### PG-08 · Kolaps marginesów CSS vs sumowanie before+after w Wordzie
- **[POTWIERDZONE — udokumentowane EDITOR_KEYBOARD §4]** Odstęp między akapitami w edytorze bywa mniejszy (większy margines wygrywa) niż w Wordzie (suma). Świadomie zostawione.
- **Ryzyko:** średnie · **Wpływ:** wizualny (kumuluje się w rozjazd paginacji)

#### PG-09 · Interlinia „single" ciaśniejsza niż w Wordzie
- **[POTWIERDZONE — udokumentowane]** Word single ≈ 1.15 (line-gap metryki fontu); CSS `line-height:1` dokładnie 1.0. Dotyczy każdej linii → największy pojedynczy wkład w rozjazd wysokości dokumentu.
- **Kierunek:** kalibracja mnożnika per rodzina fontu (np. 1.149 dla Calibri) — przybliżenie mierzalne.

#### PG-10 · lineRule=atLeast renderowany jak exact
- **[POTWIERDZONE]** CSS nie ma „co najmniej"; wysoka treść w linii (obraz inline) w Wordzie rozpycha linię, w edytorze nachodzi/przycina. Round-trip zachowany markerem `--w-line-rule:atLeast`.
- **Ryzyko:** niskie · **Wpływ:** wizualny

#### PG-11 · Dzielenie wyrazów (autoHyphenation) nieobsługiwane
- **[POTWIERDZONE — brak]** Word z włączonym dzieleniem łamie słowa → krótsze akapity; edytor nie (inne zawijanie, inna paginacja). Ustawienie tracone (ST-08).
- **Kierunek:** `hyphens:auto` + `lang` przybliża, ale słowniki przeglądarki ≠ Word.

#### PG-12 · Gutter / mirrorMargins / docGrid nieobsługiwane
- **[POTWIERDZONE — brak]** · **Ryzyko:** niskie · **Wpływ:** wizualny, round-trip.

#### PG-13 · Tabele pływające (tblpPr) — brak obsługi
- **[POTENCJALNE — brak `tblpPr` w kodzie]** Tabela pozycjonowana z opływem tekstu renderuje się jako zwykła blokowa; pozycjonowanie tracone na eksporcie.
- **Ryzyko:** średnie · **Wpływ:** wizualny, round-trip

---

## 4. TABELE i LISTY (uzupełnienie — stan po ADR-0026)

#### TB-01 · Przekątne obramowania komórek (tl2br/tr2bl) brak — **[POTWIERDZONE — udokumentowane]**; wizualne, niskie.
#### TB-02 · Warunkowe formatowanie TEKSTU ze stylu tabeli (bold w firstRow itd.) nieprzenoszone; regiony narożne NW/NE/SW/SE brak — **[POTWIERDZONE]**; wizualne, średnie.
#### TB-03 · fitText / hideMark nieobsługiwane — **[POTWIERDZONE]**; wizualne, niskie.
#### TB-04 · Rozwiązane wartości stylu tabeli zapiekane w inline CSS → po round-trip stają się formatowaniem bezpośrednim (referencja stylu przeżywa w `data-tbl-style`, ale późniejsza zmiana definicji stylu w Wordzie nie wpłynie na zapieczone wartości) — **[POTWIERDZONE — świadoma decyzja ADR-0026]**; strukturalne, średnie.
#### TB-05 · Szerokości kolumn: px(int)→twips→px — kumulacja zaokrągleń przy każdym cyklu zapisu **[POTENCJALNE]** — po N autosave'ach kolumny mogą dryfować o pojedyncze twipy; test wielokrotnego round-tripu brak.
#### LS-01 · Punktory Wingdings/Symbol: font celowo nieaplikowany w HTML (`fontCss` pomija wingdings/symbol) **[POTWIERDZONE: 1036-1041]** — znak z obszaru PUA (np. F0B7) renderuje się jako inny glif/tofu zamiast punktora Worda.
- **Kierunek:** mapa PUA→Unicode (F0B7→•, F0A7→▪ itd.).

---

## 5. FORMATOWANIE ZNAKÓW

#### CH-01 · Style podkreśleń spłaszczane do pojedynczego
- **Obecne:** każdy `w:u` ≠ none → `<u>`; double/dotted/dash/wavy i **kolor podkreślenia** tracone (render + eksport). **[POTWIERDZONE: GetRunSemanticFlags:2596]**
- **Ryzyko:** niskie · **Wpływ:** wizualny, round-trip
- **Kierunek:** `text-decoration-style/color` + data-*.

#### CH-02 · Kerning, skala znaków (w:w), pozycja (w:position), efekty (outline/shadow/emboss/imprint) — brak
- **[POTWIERDZONE — brak w readerze]** Tekst rozstrzelony przez `w:w=150%` renderuje się normalną szerokością; właściwości tracone na eksporcie.
- **Ryzyko:** niskie–średnie · **Wpływ:** wizualny, round-trip
- **Kierunek:** `w:w` → `transform:scaleX()` lub `font-stretch`; `w:position` → `vertical-align: Npt`.

#### CH-03 · RTL i języki mieszane — brak jakiejkolwiek obsługi
- **[POTWIERDZONE — brak `bidi`/`rtl` w kodzie]** `w:bidi` (akapit), `w:rtl` (run), `rtlGutter` nieczytane → tekst hebrajski/arabski: zła kolejność wizualna, złe wyrównanie; atrybuty tracone na eksporcie.
- **Ryzyko:** wysokie *jeśli* dokumenty RTL wystąpią; w praktyce projektu (pl-PL) — niskie prawdopodobieństwo.
- **Kierunek:** `dir="rtl"`+`unicode-bidi` — dobre przybliżenie.

#### CH-04 · Ligatury / funkcje OpenType (w14:ligatures, numForm...) — brak; przeglądarka stosuje własne domyślne ligatury. **[POTWIERDZONE — brak]**; niskie.
#### CH-05 · Fonty niezainstalowane → fallback generyczny (serif/mono/sans) — inne metryki → inne zawijanie **[POTWIERDZONE — udokumentowane]**; średnie; mitygacja: dostarczenie webfontów firmowych.
#### CH-06 · `w:color` z `themeTint/themeShade` na runach — **[POTENCJALNE]**: `ResolveThemeColor` rozwiązuje kolor bazowy motywu; tint/shade potwierdzone dla tabel, dla runów do weryfikacji (możliwy zbyt ciemny/jasny kolor).
#### CH-07 · Jawne wyłączenie formatowania (`w:b w:val="0"` na runie przy stylu bold) nie nadpisuje stylu — **[POTWIERDZONE — udokumentowane FIDELITY §3]**; rzadkie; niskie.

---

## 6. OBRAZY I GRAFIKA (uzupełnienie)

#### IM-01 · Rotacja (`a:xfrm/@rot`) nieobsługiwana — obraz obrócony w Wordzie renderuje się prosto; rotacja tracona na eksporcie. **[POTWIERDZONE — udokumentowane]**; wysokie wizualnie dla dokumentów z obróconymi elementami.
#### IM-02 · Crop nie skaluje do boxa jak Word (`clip-path` bez powiększenia) — **[POTWIERDZONE]**; średnie, wizualne.
#### IM-03 · Zawijanie tekstu wokół obrazu: tylko front/behind + przybliżenie square; `wrapTight/wrapThrough/wrapTopAndBottom` niemożliwe/nieobsługiwane — **[POTWIERDZONE]**; `wrapSquare` z odsunięciami (distT/B/L/R) do weryfikacji **[POTENCJALNE]**.
#### IM-04 · Kotwice względne (`relativeFrom=margin/page/column`, `align=center/right`, `percentOffset`) — reader czyta tylko `PositionOffset` (offset bezwzględny) **[POTWIERDZONE: 3059-3067]** → obraz wyrównany „do środka strony" w Wordzie może wylądować w innym miejscu; semantyka kotwicy tracona.
#### IM-05 · Kolejność warstw (relativeHeight) wielu obiektów pływających — **[POTENCJALNE]**: brak śladu odczytu z-order; nakładające się obiekty mogą renderować się w złej kolejności.
#### IM-06 · Czysty wektor EMF/WMF → niewidoczny w edytorze (pusty SVG); oryginał pass-through do DOCX **tylko** na ścieżce z pass-through obrazów — przy autosave przez `Convert` obraz jest re-osadzany z data URI **[POTENCJALNE — do weryfikacji, który content-type trafia do v2 po autosave]**. Ryzyko podmiany EMF→placeholder w v2.
#### IM-07 · Obrazy połączone (linked, external relationship) — **[POTENCJALNE — brak obsługi External rel]**: prawdopodobnie drop.

---

## 7. ZACHOWANIE EDYTORA (funkcjonalne różnice)

#### ED-01 · Zaznaczenie nie przekracza granicy strony
- **Obecne:** każda strona = osobny `contenteditable` **[POTWIERDZONE — EDITOR_KEYBOARD §1]** → nie da się zaznaczyć/skopiować/usunąć fragmentu obejmującego 2 strony ani Ctrl+A całego dokumentu. Word: dokument ciągły.
- **Ryzyko:** wysokie (codzienna edycja) · **Wpływ:** funkcjonalny
- **Kierunek:** trudny w architekturze multi-contenteditable; rozważyć globalne operacje (Ctrl+A, kopiowanie międzystronicowe) obsługiwane programowo.

#### ED-02 · Enter nie realizuje stylu „następny akapit" (`w:next`)
- W Wordzie Enter po Heading1 daje Normal; w edytorze domyślne `contenteditable` klonuje bieżący blok (kolejny H1). **[POTWIERDZONE — brak custom handlera Enter]**; średnie.

#### ED-03 · Tab = indent/outdent bloku, nie znak tabulacji / demote listy jak Word — **[POTWIERDZONE — EDITOR_KEYBOARD]**; średnie; w tabeli Word przechodzi do następnej komórki — do weryfikacji **[POTENCJALNE]**.

#### ED-04 · Wklejanie z Worda: sanitizacja regexowa, mso-markup niemapowany
- **Obecne:** paste `text/html` czyszczony regexami (script/style/on*) **[POTWIERDZONE — R-09]**; `mso-*` style, `<o:p>`, listy Worda (mso-list) nie są tłumaczone na semantykę edytora → wklejona lista z Worda staje się akapitami z ręcznymi numerami; ryzyko ominięcia sanitizacji regexowej (R-09, rekomendowany DOMPurify).
- **Ryzyko:** średnie (funkcjonalne) + średnie (bezpieczeństwo) · **Wpływ:** funkcjonalny, strukturalny

#### ED-05 · Undo/redo a repaginacja przebudowująca DOM
- **Obecne:** `pageContents.set` → `[innerHTML]` przebudowuje strony po każdej repaginacji (debounce 250/600 ms) **[POTWIERDZONE: 3616-3622]**; natywny stos undo przeglądarki nie przeżywa wymiany innerHTML → **[POTENCJALNE]:** Ctrl+Z po repaginacji nie cofa albo cofa nieprzewidywalnie (brak własnego historia-managera w kodzie).
- **Ryzyko:** wysokie (codzienna edycja) · **Testy:** e2e undo po przelaniu strony — brak.

#### ED-06 · Tab-stopy: przypisanie k-ty tab → k-ty stop (Word: „następny stop za bieżącą pozycją x"); leader (kropki) nierysowany w edytorze — **[POTWIERDZONE — udokumentowane ADR-0025]**; średnie, wizualne.

#### ED-07 · Nagłówek stron parzystych: brak UI edycji (round-trip działa) — **[POTWIERDZONE]**; niskie.

#### ED-08 · Kursor/karetka: kotwica `{block,offset}` może się przesunąć, gdy tabela przed karetką podzieli się inaczej — **[POTWIERDZONE — udokumentowane]**; niskie.

#### ED-09 · Drag&drop treści między stronami — **[POTENCJALNE]**: osobne contenteditable; przeciągnięcie na inną stronę może duplikować/gubić formatowanie; brak testów.

---

## 8. Ograniczenia trwałe HTML/CSS (z klasyfikacją wykonalności)

| Obszar | Możliwe przybliżenie? | Rekomendacja |
|---|---|---|
| Paginacja identyczna z Wordem | Nie (inny silnik layoutu) | minimalizować rozjazd: kalibracja line-height (PG-09), suma marginesów, hyphenation; wartości PAGE traktować jako orientacyjne |
| Metryki/kerning fontów | Nie | webfonty firmowe + fallback generyczny (już jest) |
| Zawijanie tight/through wokół kształtu | Nie natywnie (`shape-outside` tylko dla float) | zachować atrybuty round-trip; render jako square |
| line-height „atLeast" | Nie w czystym CSS | obecny marker round-trip wystarcza; ewent. JS-measure |
| Sumowanie odstępów akapitów (Word) vs kolaps (CSS) | Tak (padding albo `margin-trim`/wyliczanie) | rozważyć wyliczanie efektywnych marginesów w readerze (margin-top = before + after poprzednika) |
| Kolumny tekstowe z łamaniem jak Word | Częściowo (`column-count`) | round-trip + przybliżony render |
| Pola dynamiczne (TOC/REF) | Tylko wartość statyczna | render read-only + zachowanie instrukcji |
| contenteditable (selection wielostronicowa, undo) | Częściowo (własny model selekcji/historii) | docelowo własny command-manager undo |
| Jednostki: twips↔px zaokrąglenia | Tak | trzymać oryginalne twips w data-* (wzorzec już stosowany w tabelach) i preferować je przy eksporcie |

---

## 9. Priorytetyzacja

### 9.1 Krytyczne — utrata danych / uszkodzenie dokumentu
1. **KR-04** — tekst z `w:ins` znika (treść!).
2. **KR-05** — TOC/REF/MERGEFIELD tracą kod i wartość (treść spisu treści znika).
3. **KR-06** — treść pól tekstowych/kształtów/równań znika.
4. **KR-01** — regeneracja pakietu w autosave/„Zakończ" (wzmacniacz wszystkich pozostałych strat).
5. **KR-02/KR-03** — przypisy, komentarze.
6. **KR-12** — ukryty tekst staje się widoczny (ryzyko treści roboczych w dokumencie wychodzącym).
7. **KR-08** — DATE nadpisywany datą bieżącą (dokumenty umowne).
8. **KR-14** — Sign gubi sekcyjne nagłówki.
9. **ST-03** — kasowalny marker sekcji.
10. **KR-07, PG-03** (utrata atrybutów keep*), **KR-11** (ochrona dokumentu).

### 9.2 Największe różnice wizualne
PG-01/PG-02 (paginacja, niedzielone akapity), PG-09 (interlinia single), PG-08 (kolaps
marginesów), CH-05 (fonty), KR-09 (kolumny), IM-01 (rotacja), IM-04 (kotwice względne),
LS-01 (punktory Wingdings), TB-02 (bold nagłówka tabeli ze stylu), PG-04 (powtarzany
nagłówek tabeli).

### 9.3 Problemy codziennej edycji
ED-01 (zaznaczenie przez granicę stron), ED-05 (undo po repaginacji), ED-04 (paste z Worda),
ED-02 (Enter po nagłówku), ED-03 (Tab), ST-06 (nowe listy), ED-06 (tab-stopy/leader).

### 9.4 Ograniczenia HTML/przeglądarki (niezawinione przez implementację)
paginacja 1:1, metryki fontów, tight/through wrap, atLeast, natywny brak stron w HTML,
contenteditable (selection/undo/IME), clipboard API (uprawnienia), window.close().

### 9.5 Nieobsługiwane elementy DOCX (lista zbiorcza)
przypisy dolne/końcowe · komentarze · śledzenie zmian (ins/del/moveFrom/moveTo/rPrChange) ·
pola poza PAGE/NUMPAGES/DATE (TOC/REF/SEQ/MERGEFIELD/FORM*) · zakładki · równania (oMath) ·
kształty/WordArt/SmartArt/wykresy/OLE/pola tekstowe · kolumny (w:cols) · pgNumType · vAlign
sekcji · documentProtection/permStart · vanish · kerning/w:w/w:position/efekty znakowe ·
style podkreśleń · RTL/bidi · ligatury · hyphenation · gutter/mirrorMargins/docGrid ·
tblpPr (tabele pływające) · przekątne obramowania · fitText/hideMark · rotacja obrazu ·
kotwice relativeFrom/align · linked images · watermark/tło dokumentu · settings.xml (compat!).

### 9.6 Obszary nieoceniane bez dodatkowych dokumentów testowych
- dokumenty RTL i wschodnioazjatyckie (CH-03/CH-04),
- SDT wszystkich typów (ST-04),
- wieloobiektowe układy pływające / z-order (IM-05),
- EMF/WMF w autosave (IM-06 — co realnie trafia do v2),
- dokumenty w trybach zgodności (compat) Word 2007/2010,
- duże dokumenty (100+ stron) — wydajność repaginacji i stabilność kotwicy karetki,
- wklejanie z realnego Worda/Outlooka/Excela (korpus paste),
- run-level themeTint/themeShade (CH-06), wrapSquare z odsunięciami (IM-03).

### 9.7 Rekomendowana kolejność implementacji

**Etap 1 — stop utracie danych (najpierw):**
1. Reader: `w:ins` renderowany jako treść, `w:del` pomijany świadomie (KR-04).
2. Reader: nieznane pole złożone → emituj wartość cache jako tekst + `data-field-instr` (KR-05); DATE z wartości cache (KR-08).
3. Reader: tekst z `wps:txbx` ekstrahowany do bloku (KR-06 minimum).
4. Zakładki → `<a id>` round-trip (KR-07); `w:vanish` → `display:none`+data-* (KR-12).
5. keepNext/keepLines/widowControl/pgNumType/cols → round-trip data-* (PG-03, KR-10, KR-09).
6. Sign: przekazać `SectionHeadersFooters` (KR-14). Marker sekcji `contenteditable=false` (ST-03).
7. Detekcja nieobsługiwanych elementów przy imporcie + ostrzeżenie użytkownika („dokument zawiera przypisy/komentarze, które nie będą zachowane") — tani bezpiecznik do czasu pełnego wsparcia.

**Etap 2 — struktura dokumentu:**
8. Pass-through pakietu w ścieżce Save/Finish (R-20/KR-01): styles.xml, settings.xml, numbering.xml (remap numId — R-19), footnotes/comments parts.
9. `data-style-id` na akapitach + `w:pStyle` na eksporcie (ST-01).
10. `data-page-break-before` zamiast twardego breaka (ST-02). NBSP w pustych akapitach (ST-07).

**Etap 3 — paginacja i układ:**
11. Kalibracja interlinii single per font (PG-09) + strategia sumowania odstępów (PG-08).
12. Dzielenie akapitu między strony (PG-02); honorowanie keepNext/cantSplit/tblHeader w paginacji (PG-03/PG-04).
13. pgNumType w `_pageNumberHtml`; kolumny render `column-count` (KR-09).

**Etap 4 — tabele i listy:**
14. Render markerów listy z `data-lvl-text` (CSS counters) — sufiksy i „1.1.2" (ST-06).
15. Mapa PUA punktorów Wingdings/Symbol (LS-01). Test dryfu twips↔px (TB-05).
16. Warunkowy TEKST ze stylów tabel (TB-02).

**Etap 5 — różnice formatowania:**
17. Style podkreśleń + kolor (CH-01); `w:w`/`w:position` (CH-02); tint/shade na runach (CH-06).
18. Rotacja obrazów (IM-01), kotwice relativeFrom/align (IM-04), crop ze skalowaniem (IM-02).

**Etap 6 — zaawansowane/rzadkie:**
19. Render przypisów/komentarzy (KR-02/03 pełne), TOC read-only, SDT pełne typy (ST-04).
20. RTL (CH-03), hyphenation (PG-11), watermark (KR-13), ochrona dokumentu egzekwowana (KR-11).
21. Własny undo-manager + globalna selekcja międzystronicowa (ED-01/ED-05).

## 10. Braki w testach (wymagane nowe)

- round-trip **wielokrotny** (×3 autosave) z diffem XML — wykrywanie dryfu (TB-05, ST-07);
- korpus realnych dokumentów Word z asercją „części pakietu nie znikają" (lista parts przed/po);
- golden: przypisy, komentarze, track-changes, TOC, bookmarki, txbx, oMath, cols, pgNumType,
  vanish, underline styles, tblpPr, kotwice align/relativeFrom — dziś **zero** pokrycia;
- e2e: undo po repaginacji, zaznaczenie na granicy stron, paste z Worda (zrzuty clipboardu),
  kasowanie przy markerze sekcji;
- test porównawczy liczby stron edytor vs Word (manualny protokół w §6b DOCX_CONVERSION już
  istnieje — rozszerzyć o powyższe elementy).
