# Features

## Cel pliku

Funkcje systemu z perspektywy produktu i implementacji. Aktualizuj przy zmianie funkcji.

## Status funkcji

| Funkcja | Status | Backend | Frontend | API |
|---|---|---|---|---|
| Ingest dokumentu od aplikacji zewnętrznej (Krok 1) | Implemented | `IngestExternalDocumentCommand` | n/d | `POST /api/v1/document` (Services) |
| Tryb podglądu DOCX/PDF (Krok 2) | Implemented | endpointy gotowe | `?masterId=` bez `versionId` ładuje **v1** przez `downloadBaseVersion`; routing PDFViewer/DocxEditor po metadanych | `GET .../{id}/download`, `GET .../{id}/metadata` |
| Tryb edycji DOCX (Krok 3) | Implemented | gotowe | `document-editor` | `GET .../versions/{vid}/download` |
| Auto-save (nadpisanie v2 w miejscu) | Implemented | `UpdateDocumentVersionCommand` | timer + switch AutoSave | `PUT .../versions/{vid}` |
| Ujednolicony „Zapisz" + „Pobierz dokument" | Implemented | — | `saveDocument()`/`downloadDocument()` | `PUT`/`POST save` |
| Metadane (returnUrl, classification) | Implemented | `GetDocumentMetadataQuery` | — | `GET .../{id}/metadata` |
| Wersjonowanie + przywracanie | Implemented | `RestoreDocumentVersionCommand` | admin/historia | `POST .../restore/{vid}` |
| Podpisy cyfrowe (Custom XML Part) | Implemented | `DigitalSignatureService` | dialog podpisu | `POST /api/document/sign`, `/verify-signatures` |
| Kody kreskowe / QR | Implemented | `BarcodeGeneratorService` | `barcode-dialog` | `POST /api/barcode/generate` |
| Szablony dokumentów | Implemented | queries templates | menu szablonów | `GET /api/document/templates` |
| Konfiguracja tabeli (boczny panel) | Implemented | — | `d2-table-properties-panel` (dok lewy, jak wyszukiwanie; ADR-0006) | n/d (frontend) |
| Obramowania/linie tabeli (rodzaj, grubość, kolor, miejsce; auto-zakres z zaznaczenia) | Implemented | — | zakładka „Obramowania" w panelu; `table-style.util` `applyBorderToCells`/`classifyBorderTarget` (style inline; ADR-0007) | n/d (frontend) |
| PDF viewer | Implemented | n/d (statyczny plik) | `pdf-viewer` (lazy, pdfjs) | `GET .../download` |
| Eksport PDF | Deprecated/Placeholder | `501 Not Implemented` | — | `POST /api/document/export-pdf` |
| Zakończ i wyślij (zwrot na returnUrl) | Implemented | `FinishAndSendDocumentCommand`: **synchroniczna 1. próba** (inline) + worker `DocumentDeliveryWorker` dla kontynuacji/retry; `AbortSend`/`ContinueDelivery` | `finishDocument()`: modal wysyłki z „Przerwij"/„Zamknij"; stan `Sending` ≠ błąd | `POST .../versions/{vid}/finish` (**200** `{deliveryId,status,documentStatus,delivered,error?}`), `POST .../abort-send`, `POST .../continue-delivery`, `GET .../deliveries/...` |
| Panel admina wysyłek (monitoring + akcje) | Implemented | `GetDeliveriesByStatusQuery`, `RequeueDeliveryCommand`, `CancelDeliveryCommand`, `UpdateDeliveryRecipientUrlCommand` | `admin-deliveries` (`/admin/deliveries`): lista, filtry, autoodświeżanie 3 s, szczegóły wiersza, akcje Wznów/Anuluj/Edycja adresu/Kopiuj link | `GET .../deliveries?status=`, `POST .../deliveries/{id}/retry`, `POST .../deliveries/{id}/cancel`, `PUT .../deliveries/{id}/recipient-url` |
| Edycja nagłówka/stopki (Word-like) | Implemented | konwersja DOCX↔HTML (`DocxToHtmlConverter`/`HtmlToDocxConverter`) | `wysiwyg-editor`: wejście (db)klik, kontekstowy toolbar, „Zamknij", first/odd/even, numery stron | część `save`/`PUT` (header/footer w `SaveDocumentRequest`) |
| Import nagłówka/stopki: wybór wariantu sekcji | Implemented (import: default+first+even) | `ExtractHeader/ExtractFooter` wg `sectPr`/refs + `titlePg`/`evenAndOddHeaders` | spread pełnego obiektu header/footer | — |
| Round-trip first/even header/footer (zapis) | Implemented (single-section) | `HtmlToDocxConverter.WriteHeaderPart`/`WriteFooterPart(html, type)` + `EnsureTitlePage`/`EnsureEvenAndOddHeaders` | bez zmian (model TS już ma pola) | część `PUT .../versions/{vid}` (przez `SaveDocumentRequest`) |
| Kontrola dostępu do dokumentu (`allowedCorporateKeys`) | Implemented | `DocumentAccessPolicy` + `IDocumentAccessGuard`; 403 w handlerach treści/metadanych; admin omija listę | `documentAccessGuard` + `DocumentAccessDeniedComponent` | 403 na `GET /{id}`, `/metadata`, `/download`, `/versions/{vid}/download` |
| Uwierzytelnianie + role Entra ID (`Operator`/`Administrator`) | Implemented (wymaga konfiguracji Entra + restore/npm install) | **Microsoft.Identity.Web** (`AddMicrosoftIdentityWebApi`), klasowy `RequireAppOperator` na `DocumentController`/`DocumentStorageController` (ADR-0022), `RequireAppAdmin` na module admina, `ClaimsCurrentUserProvider` (claim `corpKey`, case-insensitive) | MSAL + guardy `authGuard`/`resourceGuard`/`homeRedirectGuard`/`documentAccessGuard`, `MsalInterceptor`; runtime `config.json` | 401 (brak tokena), 403 (rola/dostęp); admin: `GET /`, `deliveries*`, `GET /api/identity/users` |
| RBAC obszarów UI (Operator vs Administrator rozłączne) | Implemented | `ResourcesProvider.GetForUser`: Operator→`[dashboard,editor,viewer]`, Admin→`[admin]`; `GET /api/identity/resources` (`RequireAppOperator`) | `resourceGuard` (po `data.resource`/ścieżce), `homeRedirectGuard` na root (Operator→dashboard, Admin→`/admin`, brak roli→`/brak-uprawnien`) | 403 → „brak uprawnień" |
| Security upload + returnUrl (centralne polityki) | Implemented | `IFileUploadSecurityService` (MIME/magic-bytes/ZIP-hardening/VBA), `IReturnUrlValidator` (SSRF: loopback/private IP/allowlista), `IFileScanner` (`NoOpFileScanner` — R-28); sekcje `Security:Upload`/`Security:ReturnUrl` | — | egzekwowane w handlerach upload/callback |
| `userDownload` / „Pobierz dokument" (gate pobierania) | Implemented | `ExternalDocumentMetadata.IsUserDownloadAllowed`; 403 gdy brak flagi | pozycja menu ukrywana | `POST .../user-download` |
| `showSaveState` (ukrycie UI zapisu) | Implemented | metadane `showSaveState=false` | ukrywa autozapis + „Zapisz" + „Ostatnia modyfikacja" | `POST /api/v1/document` (pole `ShowSaveState`) |
| External: callback-url / unlock / status | Implemented | `UpdateCallbackUrlCommand`, `Document.MarkSaved()`, `GetDocumentStatusQuery` (odpowiedź okrojona do `{masterId,status}`) | n/d | `PUT/POST/GET /api/v1/document/{masterId}/...` |
| Konwersja binarnego `.doc` → DOCX | Implemented | `LegacyDocBinaryConverter` (FIB + piece table) w `DocumentInputNormalizer` | n/d | część `POST /open` / ingest |
| Grafiki legacy (VML/EMF/WMF/TIFF) → web | Implemented | `GraphicConversionService` (strategie, cache SHA-256, blank-fallback zamiast placeholdera — ADR-0020) | `data-original-src` round-trip | patrz `GRAPHICS_CONVERSION.md` |
| Wiele sekcji DOCX (geometria per sekcja) | Implemented | reader/writer ADR-0023: markery `div.docx-section-break`, pierwsza sekcja bazowa | `wysiwyg-editor`: `pageGeometries` per strona | patrz `DOCX_CONVERSION.md` |
| „Kto modyfikował" (LastModifiedBy/corpKey) | Implemented | `Document.LastModifiedBy` (SQL 011), corpKey z tokenu (ADR-0021) | kolumny+filtry w `admin-files`/`admin-deliveries` | — |
| Mapowanie grup→role Entra (Qutas, ADR-0012) | Implemented (placeholdery; wymaga grup `KUTAS_200_*`) | `RolesOptions` (sekcja `Roles`) + `ClaimsTransformer` (claim `groups`→`APP_*`; App Roles współistnieją) | role z runtime `config.json` (`adminRole`) | — |
| Microsoft Graph — lookup userów (Qutas, ADR-0012) | Implemented (app-only; wymaga ClientSecret z GCP) | Graph v5 `GraphUserService` + `IdentityController` | — | `GET /api/identity/users?query=` (RequireAppAdmin) |
| Sekrety w GCP Secret Manager (Qutas, ADR-0012) | Implemented (off lokalnie) | `EntraSecretLoader` → `AzureAd:ClientSecret` z `SMC01{ENV}2_APP00404_entra_secret` | — | — |

## Statusy

`Unknown` · `Planned` · `In Progress` · `Blocked` · `Implemented` · `Verified` · `Deprecated`

## Feature: Ingest zewnętrzny (Krok 1)

### Cel
Przyjąć dokument od aplikacji źródłowej, zapisać oryginał i (dla DOCX) utworzyć kopię edytowalną.

### Backend
- Endpoint: `POST /api/v1/document` (`D2ServicesViewerEditor`, multipart).
- Use case: `IngestExternalDocumentCommand` / `...Handler` (warstwa Application `D2ApiViewerEditor`).
- Walidacja: kontroler + handler (typ MIME, rozmiar ≤100 MB, JSON metadanych, opcjonalna classification C1..C4, ReturnUrl dla DOCX) + centralne `IFileUploadSecurityService`/`IReturnUrlValidator`. (Osobny `IngestExternalDocumentCommandValidator` już nie istnieje.)
- Reguły: BR-003, BR-004, BR-006, BR-007 (patrz `DOMAIN.md`).

### Otwarte pytania
- Czy wersja edytowalna ma być realną konwersją/normalizacją DOCX, czy dosłowną kopią bajtów? Obecnie kopia bajtów.

## Feature: Auto-save / ujednolicony zapis (Krok 3)

### Cel
Zapisywać zmiany w trybie edycji bez mnożenia wersji.

### Backend
- `PUT /api/documentstorage/{masterId}/versions/{versionId}` → `UpdateDocumentVersionCommand` → `Document.UpdateVersion` (nadpisanie GCS pod tym samym `versionId`). Reguły BR-002, BR-005.

### Frontend
- `components/document-editor/document-editor.ts`: timer `rxjs timer(intervalMs)`, sygnały `autoSaveEnabled/autoSaveStatus/lastAutoSaveAt`, switch „AutoSave" w nagłówku (widoczny w trybie edycji).
- `saveDocument()` = zapis przez API (PUT gdy versionId, inaczej POST). `downloadDocument()` = pobranie pliku lokalnie (dawne „Zapisz").
- Konfiguracja: `environment.autoSave { enabled, intervalSeconds: 30 }`.

## Feature: Zakończ i wyślij (Krok 4)

### Cel
Zakończyć pracę nad dokumentem, utrwalić aktualny stan edytora i wysłać finalny plik na `ReturnUrl` z metadanych — z natychmiastową informacją zwrotną (synchroniczna pierwsza próba), odpornie na restart i wiele instancji (kontynuacja w tle).

### Backend (przeprojektowane 2026-06-21, doszczelnione 2026-06-27/28)
- `POST /api/documentstorage/{masterId}/versions/{versionId}/finish` → `FinishAndSendDocumentCommand` / `...Handler`:
  1. nadpisuje wersję edytowalną treścią z edytora (jak auto-save),
  2. zamraża niezmienny snapshot w GCS (`deliveries/{deliveryId}`, SHA-256),
  3. tworzy zadanie `DocumentDelivery` i wykonuje **synchroniczną pierwszą próbę** (`BeginInlineAttempt` → `IDeliverySender.SendAsync`); rekord trafia do bazy od razu jako `Sending` bez lease (nigdy jako claimowalny `Pending` — brak okna podwójnej wysyłki z workerem).
  4. Sukces → `delivery.MarkSent` + `document.Sent` → **200** `{ deliveryId, status, documentStatus, delivered: true }`. Błąd → `HoldAfterFailedInlineAttempt` (`RetryScheduled` „zaparkowane" na `DeadlineAt`) + `document.DeliveryFailed` → 200 z `delivered: false` + `error`.
- Decyzja użytkownika po błędzie: `POST .../abort-send` (delivery `Cancelled` + dokument `SendAborted`, wraca do edycji) lub `POST .../continue-delivery` (`Requeue` + dokument `Queued` — dokańcza worker). `OperationCanceledException` = anulowanie (nie błąd).
- Idempotencja: jeśli istnieje aktywne zadanie dla dokumentu, zwraca je (nie tworzy nowego); unique partial index chroni przed wyścigiem. Reguły BR-010..BR-013.
- Walidator: `FinishAndSendDocumentCommandValidator` (MasterId/VersionId niepuste, Content niepusty ≤ 100 MB, CreatedBy ≤ 255).
- Worker `DocumentDeliveryWorker` (BackgroundService): claim `FOR UPDATE SKIP LOCKED` + lease, równoległość z limitem, `HttpDeliverySender` (POST multipart `file`/`masterId`/`versionId`/`corporateKey` + `Idempotency-Key`/`X-Content-SHA256`), retry `ExponentialJitterBackoff` (cap 15 min) do 24 h → `DeadLettered`; błąd non-retryable → `FailedPermanently`. Po sukcesie `Document.Status=Sent`, po porażce `DeliveryFailed`.
- Status/monitoring: `GET .../deliveries/{deliveryId}`, `GET .../deliveries?status=`, ręczne ponowienie `POST .../deliveries/{deliveryId}/retry`, anulowanie `POST .../deliveries/{deliveryId}/cancel`, zmiana adresu `PUT .../deliveries/{deliveryId}/recipient-url`.

### Frontend
- `components/document-editor/document-editor.ts`: `finishDocument()` serializuje DOCX i woła `finishAndSend`; trzy wyniki: `delivered` → ekran końcowy, `status==='Sending'` → okno informacyjne „Trwa wysyłka" (nie błąd), realny błąd → modal problemu z „Przerwij" (`abort-send`) / „Kontynuuj wysyłkę w tle" (`continue-delivery`). Modal wysyłki ma „Przerwij wysyłkę" (anulacja w locie) i „Zamknij" (wysyłka leci dalej w tle).
- Serwis `document-storage.service.ts`: `finishAndSend(masterId, versionId, { content })`, `abortSend`, `continueDelivery`, `getDeliveryStatus(deliveryId)`.

### Uwaga dla agenta
Brak (na dzień aktualizacji) testów integracyjnych claimu na realnym PostgreSQL (SKIP LOCKED / reclaim) — logikę pokrywają testy jednostkowe domeny, handlera, backoffu i walidatora. Patrz `RISKS_ASSUMPTIONS.md` (R-06).

## Feature: Tryb podglądu (Krok 2) — Implemented

### Uwaga dla agenta
GUI w trybie `?masterId=` (bez `versionId`) ładuje **wersję bazową v1** przez `documentStorageService.downloadBaseVersion(masterId)` (`GET .../{masterId}/download`); z `versionId` — wskazaną wersję (`downloadVersion`). Routing PDFViewer/DocxEditor po metadanych. Dawne ryzyko R-02 (podgląd pokazywał v2) — zamknięte.

## Feature: Nagłówek i stopka (import + edycja)

### Cel
Wierne odwzorowanie nagłówka/stopki z DOCX oraz edycja w trybie zbliżonym do Worda (wejście klikiem/dwuklikiem, kontekstowy toolbar „Nagłówek"/„Stopka", „Zamknij nagłówek i stopkę", wariant pierwszej strony / parzysty-nieparzysty, numery stron).

> Pełny reference konwersji DOCX↔HTML (pipeline, model `DocumentContent`, macierz statusów wszystkich obszarów, roadmapa, diagramy): `DOCX_CONVERSION.md`.

### Import (DOCX → HTML) — `DocxToHtmlConverter`
- **Wybór wariantu wg sekcji, nie kolejności partów.** `ExtractHeader`/`ExtractFooter` rozwiązują part przez `sectPr` → `HeaderReference`/`FooterReference` typu **Default** (helpery `ResolveHeaderPart`/`ResolveFooterPart`). Fallback do `HeaderParts.FirstOrDefault()` tylko gdy sekcja nie deklaruje referencji. **Nie używać `FirstOrDefault()` jako głównej ścieżki** — kolejność partów jest niezdefiniowana i może trafić pusty even/first (gubi logo/tekst).
- **Pierwsza strona:** czytana tylko gdy `sectPr/titlePg` (helper `HasTitlePage`) → `DifferentFirstPage` + `FirstPageHtml`.
- **Parzysty/nieparzysty:** czytany tylko gdy `settings.xml/evenAndOddHeaders` (helper `HasEvenAndOddHeaders`) → `DifferentOddEven` + `EvenHtml`. „Default" = strona nieparzysta/podstawowa.
- Rozmiar/krój: kontener `.header-footer-content` niesie domyślny font z `docDefaults` (patrz R-11). Run bez własnego `w:sz`/`w:rFonts` dziedziczy rozmiar dokumentu, nie edytora.
- Model: `HeaderFooterContent { Html, Height, DifferentFirstPage, FirstPageHtml, DifferentOddEven, EvenHtml }` (Domain) ↔ TS `HeaderFooterContent` (te same pola + `oddHtml`). Front (`document-editor.ts`) **spread'uje cały obiekt** na obu ścieżkach load — nie rekonstruować `{html,height}`, bo gubi warianty.

### Edycja (frontend) — `wysiwyg-editor` + `d2-header-footer-panel`
- `editingSection` ('header'|'footer'|'body'); `getActiveEditable()` kieruje komendy formatowania do aktywnego regionu (nie body).
- Wejście: `startEditingHeader`/`startEditingFooter` (klik/dblclik) — ładuje wariant aktywny na stronie 0 (`firstPageHtml` gdy `differentFirstPage`, inaczej `_headerHtml`/`_footerHtml`).
- **Panel boczny `d2-header-footer-panel`** (dock po lewej, ten sam slot co `Wyszukiwanie` i `d2-table-properties-panel`) zastąpił dawny pływający pasek `.header-toolbar`/`.footer-toolbar` nad dokumentem. Otwierany przez `showHeaderFooterPanel = computed(() => editingSection() !== 'body' && !showFindReplace() && !showTablePanel())` — Find / Tables mają pierwszeństwo (single-mode dock); po ich zamknięciu HF wraca, jeśli edycja trwa. Stateless: każda akcja (Inna pierwsza strona / Wstaw obraz / Numery stron / Format / Usuń) delegowana do publicznych metod `wysiwyg-editor` (rule 9). Brak osobnego primary CTA — zamknięcie przez X w nagłówku panelu, spójne z `Wyszukiwanie` i panelem tabeli.
- Wyjście: `stopEditingHeaderFooter()` — wywoływane przez (1) **X** w nagłówku panelu bocznego, (2) **ESC** (`onEscapeKeydown` w `document-editor.ts` deleguje do editora; tryb ma pierwszeństwo nad innymi bocznymi panelami), (3) klik w body editor.
- Routing wariantu (rule 10 — bez pozornej edycji): `onHeaderInput`/`onFooterInput` i `onHeaderBlur`/`onFooterBlur` zapisują do tego samego sygnału, z którego załadowano (`_headerFirstPageHtml` gdy `differentFirstPage`, inaczej `_headerHtml`). Blur dodatkowo woła `emitHeaderFooterChanges()` — partial emit by zgubił inne warianty u parenta.
- W trybie odd/even strona nieparzysta używa kanonicznego `_headerHtml`/`_footerHtml` (= „default" w OOXML; backend nie posyła osobnego `oddHtml`). Even — osobny `_headerEvenHtml`.
- Outputs: `headerChange`/`footerChange` emitują **pełny obiekt** ze wszystkimi wariantami (`html`, `firstPageHtml`, `evenHtml`, flagi); `editingSectionChange` synchronizuje stan w `document-editor`.
- Persist: `buildSaveRequest()` → `header`/`footer` w `SaveDocumentRequest` → `HtmlToDocxConverter.AddHeaderAndFooter` (emituje Default + opcjonalnie First/Even jako osobne party + `titlePg`/`evenAndOddHeaders`).

### Round-trip (HTML → DOCX) — `HtmlToDocxConverter`
- `AddHeaderAndFooter` zawsze emituje wariant **Default**. Jeśli `HeaderFooterContent.DifferentFirstPage && FirstPageHtml` — emituje dodatkowy `HeaderPart`/`FooterPart` typu **First** i ustawia `TitlePage` w `sectPr` (`EnsureTitlePage`). Jeśli `DifferentOddEven && EvenHtml` — emituje dodatkowy part typu **Even** i ustawia `EvenAndOddHeaders` w `settings.xml` (`EnsureEvenAndOddHeaders`).
- Pomocnicze: `WriteHeaderPart(html, type)`/`WriteFooterPart(html, type)` (jedna ścieżka konwersji per part, scoping obrazów do tego konkretnego HeaderPart/FooterPart); `AddHeaderReference(id, type)`/`AddFooterReference(id, type)`.
- Edycja wariantu first/even w UI jest trwała (rule 10): pole modelu → osobny part → ref → re-import.

### Ograniczenia / otwarte
- **Edycja even-page nieosiągalna z UI** — template `wysiwyg-editor.html` renderuje contenteditable tylko dla strony 0 (`@if (editingSection() === 'header' && $index === 0)`). Wariant even jest *czytany* z DOCX i *zachowywany* na zapisie (round-trip), ale nie da się go edytować bezpośrednio w GUI. Wymagałoby dodatkowego punktu wejścia (np. „Edytuj nagłówek parzystej strony" w menu Opcje albo edytowalność strony 2).
- **Wiele sekcji** — geometria (rozmiar/orientacja/marginesy) per sekcja jest round-tripowana i renderowana per strona (ADR-0023, 2026-07-05: markery `div.docx-section-break`, pierwsza sekcja bazowa). **Nagłówki/stopki per sekcja nadal jedno-kompletowe** (model `DocumentContent` niesie jeden zestaw — refs czytane/zapisywane z PIERWSZEJ sekcji, dziedziczone przez kolejne). Patrz `DOCX_CONVERSION.md` i `RISKS_ASSUMPTIONS.md` R-10.
- Round-trip rozmiaru z `docDefaults` (`.header-footer-content` wrapper) — patrz R-11.
