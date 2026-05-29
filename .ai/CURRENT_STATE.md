# Current State

> Plik krótki i aktualny. Czytaj na początku każdej sesji.

## Ostatnia aktualizacja

2026-05-29 — Word-like pozycjonowanie obrazów (MVP): czwarty docked panel **`d2-image-properties-panel`** (Rozmiar/Wyrównanie/Opływanie info-only/Usuń) w slocie Wyszukiwanie/Tabele/Nagłówek+Stopka. Stateless — wszystkie akcje przez nowe public methods na `wysiwyg-editor`: `setSelectedImageWidth/Height(px, lockAspect)`, `setSelectedImageAlignment(L/C/R/null)`, `resetSelectedImageAspect()`, `removeSelectedImage()`. Nowy `@Output() imageSelectionChange` emitowany przy select/deselect/resize-end/drop-end. Sygnał `selectedImage` + `imageLockAspect` (default true) w `document-editor`; `showImagePanel = selectedImage()!=null && !showFindReplace() && !showTablePanel()`; **panel obrazu ma pierwszeństwo nad HF panelem** (rozsądny UX: formatowanie logo w nagłówku w trakcie edycji). Round-trip rozmiaru przez istniejące `data-*-emu` na `<img>`; alignment przez `text-align` na akapicie. Drag/resize mechaniki bez zmian — tylko emit snapshot po każdej operacji. Tests: **160/160 pass** (+9 panel, +1 koordynacja). **Ograniczenie MVP**: floating wp:anchor (square/tight/behind/in-front wrap) odłożone — panel pokazuje sekcję info-only „W tekście (inline). Floating w przygotowaniu." Wcześniej:

2026-05-29 — Reguła `userDownload` (kontrola pobierania edytowanego pliku). Opcjonalne pole w `documents.metadata` JSON: default **false**; brak/null/non-true ⇒ blokada; tylko jawne `true` zezwala. **Local upload** (`UploadDocumentCommandHandler`) **automatycznie** ustawia `userDownload=true` (anti-tamper — system, nie klient). **External ingest** (`POST /api/v1/document`) dostaje opcjonalne `UserDownload: bool?`. Nowy gated endpoint `POST /api/documentstorage/{masterId}/user-download` (`DownloadEditedDocumentCommand` + handler) → 403 gdy off (sentinel `USER_DOWNLOAD_FORBIDDEN:`). Frontend: sygnał `userDownload` + `canUserDownload` computed; pozycja menu „Pobierz dokument" gated; `downloadDocument()` POST na nowy endpoint z obsługą 403. Wspólny `ExternalDocumentMetadata` (parser+serializer) zastąpił 4 lokalne kopie. DTO metadanych/statusu eksponują flagę. **325/325 backend pass, 150/150 GUI pass**. **Ograniczenie**: `GET .../download` i `GET .../versions/{vid}/download` (load do edytora) celowo nie gated — zwracają ostatnio zapisaną wersję, nie aktualny stan edytora; do follow-up. Wcześniej:

2026-05-29 — External API: 3 nowe endpointy (callback URL / unlock / status). Wszystkie na **`masterGuid`** (returnUrl w master metadata; brak osobnego user-locka — `Editing` to lock; status = atrybut master). Implementacja: domena (`Document.MarkSaved()` + `Document.UpdateMetadata()`); 2 commands + 1 query w `Application/Features/Documents/`; 3 akcje w `D2ServicesViewerEditor.Api/Controllers/DocumentController` (`PUT .../callback-url`, `POST .../unlock`, `GET .../status`). Walidacja URL współdzielona z `DocumentDelivery.IsValidRecipientUrl`. `callbackUrl` nigdy nie surfacowany w response (`hasCallbackUrl: bool`) ani w logach (token-safe). Idempotencja: callback-url (zapis tego samego URL), unlock (Saved → `Changed=false`). Stany wysyłki blokują obie zmiany (409). Testy: +25 (Application UnitTests); pełna solucja **305/305 pass**. `.ai/API_CONTRACTS` rozszerzony o tabelę endpointów + uzasadnienie decyzji identyfikatora. Wcześniej:

2026-05-29 — Menu „Pomoc" + status autozapisu w stopce: (a) nowa zakładka menu „Pomoc" po „Widok" z jedną pozycją „Zgłoś" delegującą do istniejącego `openReportEmail()` (bez duplikacji); sygnał `showHelpMenu` + `toggleHelpMenu()` + `closeAllMenus` rozszerzony; **stary przycisk „Zgłoś" usunięty** z `.header-right` (toolbar czystszy). (b) **Data/czas autozapisu przeniesione z toolbara do stopki** — `.autosave-status` (min-width:70px) usunięty; przy switchu zostaje tylko switch + label „Autozapis"; w `.status-right` nowy `.status-autosave` („Autozapis: HH:MM:SS" / „zapisywanie…" / „błąd zapisu" / „włączony"; ukryty gdy off). Bez zmian w logice (te same sygnały `autoSaveEnabled`/`autoSaveStatus`/`lastAutoSaveAt`). Testy: +4 w `document-editor.spec.ts`, **146/146 pass**. Wcześniej:

2026-05-29 — Porządek panelu nagłówka/stopki: (a) **usunięty primary CTA „Zamknij nagłówek i stopkę"** z `d2-header-footer-panel` (redundantny z X i ESC) — X jako jedyne zamknięcie z panelu (spójne z `Wyszukiwanie`/tabela); (b) **NG8107**: 9× `editor?.` → `editor.` w bindingach panelu w `document-editor.html` (`editor` jest `!:`, panel renderuje się dopiero po `editingSection() !== 'body'` emitowanym przez edytor — operand nigdy nie null). Tests **142/142**, `ng serve` czysty. Wcześniej:

2026-05-29 — 4 UX poprawki: (1) **autozapis stabilny** — `.autosave-status` rezerwuje przestrzeń (min-width:70px) niezależnie od stanu toggle, więc Zgłoś/Zapisz/Zakończ nie skaczą; (2) **prowadnice marginesów domyślnie ukryte** (`showMarginGuides=false` w obu komponentach); (3) **obraz w trybie edycji nagłówka NIE jest reskalowany** — `wrapExistingImages` nie usuwa już inline width/height; CSS `.editor-image-wrapper img` bez `width:100%`; obraz wygląda tak samo w preview i edit; (4) **pływający pasek nagłówka/stopki USUNIĘTY** — nowy komponent `d2-header-footer-panel` (dock po lewej, ten sam slot co Wyszukiwanie / Tabela), stateless, wszystkie akcje delegowane do `wysiwyg-editor` (Inna pierwsza strona / Wstaw obraz / Numery stron / Format / Usuń / **Primary „Zamknij nagłówek i stopkę"**). Koordynacja single-mode: `showHeaderFooterPanel = computed(() => editingSection()!='body' && !showFindReplace() && !showTablePanel())` — Find/Tables mają pierwszeństwo. Testy `header-footer-panel.spec` (6) + `document-editor.spec` +5 koordynacja. `npm test` **142/142 pass**. **Manual verify**: top-margin drag w trybie edycji nagłówka (overlay usunięty, ale interakcja z `d2-ruler` wymaga weryfikacji w przeglądarce). Wcześniej:

2026-05-28 — Faza 2: audyt trybu edycji nagłówka/stopki. Naprawione realne gapy: (a) **routing wariantu** — `startEditingHeader/Footer` + `onHeaderInput/Footer` + `onHeaderBlur/Footer` zapisują/ładują wariant aktywny na stronie 0 (`firstPageHtml` gdy `differentFirstPage`, inaczej `_headerHtml`); wcześniej edycja first-page niewidocznie nadpisywała default (łamanie rule 10); (b) `_computeHeaderContent/Footer` w trybie odd/even używa kanonicznego `_headerHtml/_footerHtml` dla odd (= default w OOXML); (c) **ESC** w `document-editor.ts` zamyka tryb edycji nag/stopki (priorytet nad bocznymi panelami); (d) przycisk **„Zamknij nagłówek i stopkę"** w obu toolbarach (spec funkcjonalny #3). Nowy `wysiwyg-editor.spec.ts` (7) + `document-editor.spec.ts` (+4 ESC). `npm test`: **131/131 pass**. Ograniczenie: edycja even-page nieosiągalna z UI (template renderuje contenteditable tylko dla strony 0; even round-tripuje z importu). Wcześniej:

2026-05-28 — Faza 1b: pomiary wierności importu nag/stopki na `stupki.docx`. Nowy `DocxToHtmlConverterFidelityTests.cs` (9): geometria (Height ≈ 1.25 cm, Margins ≈ 2.5 cm), font-family z `docDefaults`, image EMU 133×34 px + aspect (4:1 stays 4:1) + `data-width-emu`/`data-height-emu` do round-tripu, L/C/R alignment. Realny plik nie odsłonił nowych gapów na T1–T12. Pełny `Infrastructure.UnitTests` **43/43 pass**. R-11 pozostaje (brak sygnału z tego pliku). Wcześniej:

2026-05-28 — Round-trip first/even header/footer (zapis): `HtmlToDocxConverter` zrefaktorowany — `WriteHeaderPart`/`WriteFooterPart(html, type)` tworzą osobne party Default/First/Even + `HeaderReference`/`FooterReference` z `Type`; `EnsureTitlePage` (do sekcji) i `EnsureEvenAndOddHeaders` (do `settings.xml`). 6 testów round-trip w `HeaderFooterRoundTripTests` (write → read potwierdza pola `DifferentFirstPage`/`FirstPageHtml`/`DifferentOddEven`/`EvenHtml`). Pełny `Infrastructure.UnitTests` **34/34 pass**. Edycja wariantów first/even jest teraz trwała (reguła 10). Wcześniej:

2026-05-28 — Import nagłówka/stopki (Faza 1, fidelity): `DocxToHtmlConverter.ExtractHeader/ExtractFooter` wybierają part **default** wg `sectPr`/`HeaderReference`/`FooterReference` (nie `FirstOrDefault()`, który mógł zwrócić pusty even/first), **first-page** przez `titlePg` → `DifferentFirstPage`/`FirstPageHtml`, oraz **even/odd** przez `settings/evenAndOddHeaders` → `DifferentOddEven`/`EvenHtml` (model domeny rozszerzony o te 2 pola). Front (`document-editor.ts`) spread'uje cały obiekt header/footer na obu ścieżkach load. `.ai/FEATURES.md` uzupełnione (sekcja nagłówek/stopka). Testy backendu **13/13** (`DocxToHtmlConverterSectionReferenceTests` 8 + Header/Footer 5), `tsc --noEmit` OK. R-10 częściowo zamknięte (**pozostaje: round-trip first/even na zapisie + multi-section**). Diagnoza na `stupki.docx`. **Uwaga:** w drzewie roboczym są też 4 niezacommitowane poprawki GUI (env init via `provideAppInitializer`, switch „Autozapis", widoczność „Zakończ" wg `returnUrl`, niższe nagłówki/stopki dialogów) — patrz niżej. Wcześniej:

2026-05-27 — 4 poprawki GUI: (1) admin nazewnictwo `Wysyłki`→`Pliki do wysłania` (`admin-shell`/`admin-deliveries`); (2) dok boczny `syncTablePanel` — ponowne zaznaczenie tabeli przy otwartym wyszukiwaniu przełącza dok na formatowanie tabeli (zamyka search; jeden aktywny tryb; nadpisuje ADR-0006); (3) menu „Wstaw" — QR/Linia pozioma/Podział strony owinięte `@if (false)` (ukryte, logika zostaje); (4) toolbar — przycisk „Akapit" (`editor-toolbar` `@Output openParagraph` → istniejący `openParagraphDialog()`, ta sama ikona co menu „Narzędzia", bez duplikacji). Testy 101/101 (nowe: `admin-shell.spec` 3, toolbar +1, document-editor +5). Build OK. Wcześniej:

2026-05-27 — Globalne bannery górne: nowy kontener `d2-global-banners` w `App` (flow flex, kolejność: środowisko → offline/API), banner środowiska `d2-environment-banner` (źródło: `BuildInfoService.environment()`; mapowanie w `core/utils/environment-banner.util.ts` — Local niebieski / DEV zielony / UAT,TST jasnożółty / PRE purpurowy / PROD,PRD ukryty / nieznane fallback). Pasek braku API (`offline-banner.scss`): usunięty arbitralny `z-index:10000`, w flow → nie przykrywa edytora; logika `ConnectionStatusService` bez zmian. **Uwaga:** brak `fileReplacements` w `angular.json` → `environment.ts` zawsze `PRD`, dlatego env bierzemy z health-check API. Naprawiony stary scaffold `app.spec.ts`. Testy: 92/92 (nowe: util 10, env-banner 10, global-banners 4). Build OK. Wcześniej:

2026-05-27 — ESC zamyka aktywny boczny panel (Wyszukiwanie / właściwości+stylizacja tabeli) tą samą logiką co × — `document-editor.ts` `@HostListener('document:keydown.escape')` → `closeActiveSidePanel()` (deleguje do `closeFindReplace`/`closeTablePanel`); guardy: `event.defaultPrevented` (deselekcja obrazu w edytorze ma pierwszeństwo), otwarty dialog/menu kontekstowe, brak otwartego panelu → no-op; focus wraca do edytora tylko gdy był w panelu. Przycisk „Cofnij" (undo) zweryfikowany — działa, ten sam `undo()` co Ctrl+Z; operacje tabeli wchodzą do historii przez `notifyEditorChange()`→`input`→`onContentChange`→`saveToUndoStack`. Nowe testy: `editor-toolbar.spec.ts` (6) + `document-editor.spec.ts` (7), 13/13 OK. Build OK. Wcześniej:

2026-05-27 — Fix nagłówka/stopki DOCX→HTML: kontener `.header-footer-content` z domyślnym rozmiarem/krojem z `docDefaults` (backend `DocxToHtmlConverter`) + fallback CSS frontu (11pt, `#000`) — naprawia „za duży tekst" i zgubiony kolor. Testy backendu 5/5. Otwarte: R-10 (first-page/even-odd/wiele sekcji — bierzemy tylko pierwszy header part), R-11 (round-trip rozmiaru). Wcześniej:

2026-05-27 — Bugfix cyklu życia aktywnej tabeli: `detectTableContext` nie czyści tabeli przy selekcji poza edytorem (klik w panel boczny ≠ opuszczenie tabeli; model „last known context" w `core/utils/table-context.util.ts`), podpięto `selectionChange` edytora (klik akapitu poza tabelą czyści kontekst + wizualne zaznaczenie komórek). Build OK; `ng test` 55 passed. Wcześniej:

2026-05-27 — Auto-zakres obramowań: usunięto ręczny wybór celu; zakres wnioskowany z zaznaczenia (`classifyBorderTarget`: komórka/wiersz/kolumna/tabela/zakres), panel pokazuje podpis „Zastosowanie: …". Domyślnie (sama karetka) → aktywna komórka. Build OK; `ng test` 49 passed. Wcześniej:

2026-05-27 — Szczegółowy edytor obramowań tabeli (zakładka „Obramowania" w `d2-table-properties-panel`): rodzaj linii (ciągła/przerywana/kropkowana/podwójna/brak), grubość, kolor (paleta+picker+reset), 10 ikon miejsca (wszystkie/brak/zewn./wewn./poziome/pionowe/góra/dół/lewa/prawa) oraz **cel: cała tabela / komórka / zaznaczenie**. Galeria presetów usunięta. Logika = `core/utils/table-style.util.ts` (`applyBorderToCells` względem prostokąta zbioru komórek); styl inline (ADR-0007). Build OK; `ng test` 42 passed. Wcześniej:

2026-05-27 — Stylizacja tabel w panelu `d2-table-properties-panel`: zakładki „Układ"/„Style", gotowe style (`TABLE_STYLE_PRESETS`), obramowania (kolor/grubość/zakres), wiersz nagłówka, wiersze naprzemienne, pierwsza/ostatnia kolumna, reset. Logika = czyste funkcje `core/utils/table-style.util.ts`; styl jako inline na `<table>/<tr>/<td>` (przeżywa zapis HTML→DOCX — ADR-0007). Bez zmian backendu. Build OK; `ng test` 41 passed. Wcześniej:

2026-05-27 — Konfiguracja tabeli przeniesiona z poziomego paska `.table-toolbar` do bocznego panelu `d2-table-properties-panel` (dokowany po lewej, spójny z panelem „Wyszukiwanie"; auto-otwierany po kliknięciu tabeli, wyszukiwanie ma pierwszeństwo w doku — ADR-0006). Naprawiono uruchamianie testów GUI (`angular.json` cel `test` + `src/test-setup.ts` z Zone.js). Build OK; `ng test` panel 6/6 (2 stare scaffoldowe `app.spec.ts` padają niezależnie). Wcześniej:

2026-05-27 — funkcja „Wklej tylko tekst" w edytorze WYSIWYG: util `core/utils/paste-text.util.ts` (htmlToText/normalizeWhitespace/resolvePlainText + 17 testów) i skrót `Ctrl/Cmd+Shift+V` w `wysiwyg-editor.ts`. Build + vitest OK. Otwarte: słaby regexowy `sanitizeHtml` (R-09, rekomendacja DOMPurify), brak przycisku toolbara/menu. Wcześniej:

2026-05-25 — panel admina `/admin/deliveries` (lista wysyłek + filtry + retry + `locked_by`, statusy PL) + uruchomienie migracji `infra/sql/001..007` na lokalnym PostgreSQL (Podman). Wcześniej: funkcja „Zakończ i wyślij" (kolejka `document_deliveries` + worker) + audyt/synchronizacja `.ai/` z kodem.

## Aktualny cel prac

Domknięcie przepływu Krok 1–4 (ingest → podgląd → edycja → zakończ i wyślij). Krok 4 zaimplementowany; Krok 2 (podgląd v1) wciąż w toku.

## Aktualny etap

- Branch: `feature/baz-astral-obciaga-mb`
- Główne obszary kodu:
  - `D2ApiViewerEditor/` (internal API + Application/Domain/Infrastructure)
  - `D2ServicesViewerEditor/` (external ingest API)
  - `D2GuiViewerEditor/` (Angular)
  - `infra/sql/` (schemat)

## Co zostało zrobione (ta i poprzednie sesje)

- Ingest zewnętrzny: `IngestExternalDocumentCommand` + endpoint `POST /api/v1/document` (DOCX → v1+v2, PDF → v1).
- Płaski zapis: `Document.UpdateVersion` + `UpdateDocumentVersionCommand` + `PUT /api/documentstorage/{masterId}/versions/{versionId}` (nadpisanie GCS pod tym samym versionId).
- `DocumentVersion.ModifiedAt` + skrypt `infra/sql/005_add_version_modified_at.sql` + mapowanie EF.
- Endpoint metadanych: `GetDocumentMetadataQuery` + `GET .../{masterId}/metadata`.
- GUI: mechanizm auto-save (timer, `environment.autoSave { enabled, intervalSeconds: 30 }`), switch „AutoSave", sygnały `documentVersionId/autoSaveEnabled/autoSaveStatus/lastAutoSaveAt`.
- GUI: ujednolicony „Zapisz" (API, PUT/POST) + „Pobierz dokument" (dawne „Zapisz" = download lokalny). Usunięto `saveDocumentAs()`.
- Dokumentacja `.ai/` przepisana wg wzorca z `template/`.
- „Zakończ i wyślij": `FinishAndSendDocumentCommand` + tabela `document_deliveries` (SQL 007) + worker `DocumentDeliveryWorker` (claim `FOR UPDATE SKIP LOCKED`, retry/backoff do 24h, snapshot GCS) + endpointy `finish`/`deliveries`/`retry` + GUI `finishDocument()` z pollingiem. Patrz `DECISIONS.md` ADR-0005.
- `documents.status` (`DocumentStatus`) + skrypt SQL 006 (dodane przed tą sesją, teraz udokumentowane).
- Panel admina `/admin/deliveries` (`admin-deliveries`): lista zadań wysyłki (filtr per kolumna + dropdown statusu + paginacja, spójny z `/admin/files`), `locked_by`/`locked_until`, przycisk „Ponów" (`POST deliveries/{id}/retry`), statusy tłumaczone na PL. `DeliveryListItemDto` rozszerzony o `LockedUntil`/`LockedBy`.
- Migracje `infra/sql/001..007` wykonane na lokalnym PostgreSQL (Podman `d2viewereditor_postgres`); `007` utworzył `document_deliveries`.

## Jak uruchomić projekt lokalnie

Patrz `DEVOPS_DEPLOYMENT.md`. Skrót:

```bash
cd D2ApiViewerEditor && dotnet run --project D2ViewerEditor.Api      # 5190
cd D2ServicesViewerEditor && dotnet run --project D2ServicesViewerEditor.Api  # 15112 (swagger: /swagger)
cd D2GuiViewerEditor && npm install && npm start                     # 4200
# Wymaga: PostgreSQL (skrypty infra/sql/ 001..005) + GCS/fake-gcs-server (bucket d2viewereditor-documents)
```

## Jak zweryfikować stan

```bash
dotnet build D2ApiViewerEditor/D2ViewerEditor.sln   # ostatnio: OK (0 błędów)
cd D2GuiViewerEditor && npm run build               # ostatnio: OK
```

## Znane problemy / niedokończone

| Problem | Wpływ | Status |
|---|---|---|
| Tryb podglądu (Krok 2) ładuje aktywną wersję (po DOCX = v2), nie v1 oryginał | Medium | Open — przełączyć GUI na `/download` |
| `finishDocument()` — zaimplementowane: async wysyłka na returnUrl (kolejka `document_deliveries` + worker) | — | Zamknięte |
| `SaveDocumentVersionCommandHandlerTests` oczekuje `UpdateAsync` (płaski zapis go nie woła) | Low | Zamknięte — test asercją `DidNotReceive().UpdateAsync` + `Received(1).SaveChangesAsync` |
| Budżet SCSS `document-editor.scss` przekraczał limit | Low | Zamknięte — `angular.json` `anyComponentStyle` podniesiony (error 48 kB / warn 24 kB) |
| Ręczny „Zapisz" przy włączonym auto-save jest częściowo redundantny | Low | Akceptowalne (wymuszenie zapisu) |
| Port External API potwierdzony: 15112 (`appsettings.json` → `Urls`) | — | Zamknięte |

## Ostatni bezpieczny punkt

Build backendu i GUI przechodzą po zmianach auto-save / „Zapisz/Pobierz".
