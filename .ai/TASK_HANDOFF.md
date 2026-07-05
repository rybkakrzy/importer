# Task Handoff

> Bezpieczne przekazanie pracy kolejnej sesji/agentowi.

## Ostatnia aktualizacja (2026-07-05)

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
