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
