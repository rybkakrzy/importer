# AI-Assisted Changelog

Istotne zmiany dla kontynuacji pracy (nie zastępuje changeloga produktu).

## Format

```md
## YYYY-MM-DD — <tytuł>
### Changed
### Verified
### Notes
```

## Entries

## 2026-05-28 — Import nagłówka/stopki: wybór wg referencji sekcji (default + first-page)
### Changed
- Backend `DocxToHtmlConverter`: `ExtractHeader`/`ExtractFooter` rozwiązują part przez `sectPr`/`HeaderReference`/`FooterReference` typu **Default** zamiast `HeaderParts.FirstOrDefault()` (kolejność partów była niezdefiniowana → mógł trafić pusty even/first). Dodano helpery `ResolveHeaderPart`/`ResolveFooterPart`/`HasTitlePage` oraz `ConvertHeaderPartToHtml`/`ConvertFooterPartToHtml`. Pierwsza strona (`titlePg`) → `DifferentFirstPage`+`FirstPageHtml` (model domeny już miał te pola). Fallback do `FirstOrDefault` gdy sekcja nie ma referencji.
- Frontend `document-editor.ts`: oba miejsca ładowania (`loadFromStorage` + ścieżka `openDocument`) spread'ują cały obiekt `content.header`/`content.footer`, zamiast rekonstruować tylko `{html,height}` → wariant first-page/odd-even nie jest gubiony na granicy TS.
### Verified
- Backend build OK; `dotnet test` filtr Header/Footer + SectionReference: **10/10 pass** (5 nowych w `DocxToHtmlConverterSectionReferenceTests`, 5 istniejących).
- Frontend `tsc --noEmit` OK.
- Diagnoza na realnym `stupki.docx`: 3 nagłówki (even/default/first) + 3 stopki, brak `titlePg`/`evenAndOddHeaders`; default=header2 (logo + „ING Bank … Obciągalski"), default footer=footer2 (8pt, #808080). Stary kod mógł renderować pusty even part.
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
