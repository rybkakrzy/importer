# Current State

> Plik krótki i aktualny. Czytaj na początku każdej sesji.

## Ostatnia aktualizacja

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
