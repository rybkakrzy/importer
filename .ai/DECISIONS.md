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

## ADR-0001: `.ai/` jako pamięć projektu dla agentów AI

- Date: 2026-05-23
- Status: Accepted

### Context
Projekt rozwijany z pomocą agentów AI; potrzebny trwały, jawny kontekst między sesjami.

### Decision
`.ai/` jako pamięć projektu (struktura wg `template/`), plus `CLAUDE.md` i `AGENTS.md` w root.

### Consequences
Stały kontekst i handoff; wymaga dyscypliny aktualizacji.

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

## ADR-0012: Pełny wzorzec Doc2/D2WebCore dla Entra ID (Identity.Web + grupy→role + Graph + Secret Manager + Keycloak)

- Date: 2026-06-10
- Status: Accepted (kod gotowy; aktywacja per-środowisko wymaga realnych wartości Entra/GCP/Keycloak)

### Context
Po wdrożeniu lekkiej integracji Entra (ADR-0011: goły JwtBearer + App Roles) padła decyzja, by ViewerEditor przyjął **pełny wzorzec referencyjny Doc2/D2WebCore** (analiza: `analiza_implementacji_entra_id_pelna.md`). Wybór użytkownika: pełny wzorzec, **mapowanie grup→role**, praca na bieżącym drzewie. Wzorzec Doc2 to serwerowy web-app+API; ViewerEditor to SPA+API — adoptujemy części pasujące do tego kształtu (bez serwerowego OIDC/cookie).

### Decision
- **Biblioteka:** `Microsoft.Identity.Web` 3.12.0 (`AddMicrosoftIdentityWebApi`) zamiast gołego `JwtBearer`. JwtBearer NIE jest już pinowany (Identity.Web dostarcza go tranzytywnie per-TFM — pin 8.0.12 dawał NU1605 vs wymóg 9.x na konsumentach net9.0).
- **Mapowanie grup→role (Doc2):** `RolesOptions` (sekcja `Roles`: `GroupPrefix` + `Roles[]{RoleName,GroupNames}`) + `Doc2ClaimsTransformer : IClaimsTransformation` — claim `groups` → role aplikacyjne (`APP_Pracownik`/`APP_Admin`). **App Roles zachowane** (claim `roles` z tokena przeżywa) → grupy i App Roles **współistnieją**. Polityki `RequireAppEmployee`/`RequireAppAdmin` bez zmian (wymagają tych samych nazw ról).
- **Microsoft Graph** (v5, **5.103.0** jak Doc2): `IGraphUserService`/`GraphUserService` app-only (`ClientSecretCredential` + `.default`), `GET /api/identity/users?query=` (RequireAppAdmin). Aktywny tylko gdy jest ClientSecret; inaczej `DisabledGraphUserService` (no-op) → lokalnie bez sekretu działa. **Nie** użyto `Microsoft.Identity.Web.MicrosoftGraph` (to Graph v4).
- **GCP Secret Manager** (`Google.Cloud.SecretManager.V1` 2.6.0): `EntraSecretLoader` wstrzykuje `AzureAd:ClientSecret` z `SMC01{ENV}2_APP00404_entra_secret` na starcie. Guard `Enabled` + try/catch → nigdy nie wywala startu (lokalnie wyłączone).
- **Keycloak (legacy) dual-auth:** gdy `Keycloak:Enabled`, drugi schemat `JwtBearer` obok Entra + **policy scheme** `EntraOrKeycloak` wybierający schemat po **issuerze tokena** (`AuthSchemes.SelectByIssuer`). Wyłączony → tylko Entra (bez zmian zachowania).
- **Frontend runtime config (Doc2):** `assets/configs/config.json` ładowany w `main.ts` **przed** bootstrapem → `RUNTIME_AUTH_CONFIG` (token z root-factory = `environment.auth` jako fallback). Fabryki MSAL + `appAdminGuard` + `CurrentUserService` czytają z runtime-configu. Jeden build na wszystkie środowiska.
- **Wszystkie wartości środowiskowe to placeholdery** (tenant, clientId, grupy `GSAPW4D_DOC2_*`, sekrety, GCP project, Keycloak) — nigdy realnych sekretów w repo.

### Consequences
Architektura zbieżna z Doc2 w częściach pasujących do SPA+API. Mapowanie grup→role odwraca decyzję „App Roles only" z ADR-0011, ale je zachowuje (hybryda). Build backendu OK; testy: Api.UnitTests **50** (`Doc2ClaimsTransformerTests` 6, `AuthSchemesTests` 5, `IdentityControllerTests` 4), GUI **240** (`runtime-config.spec` 4), AOT build OK. **Aktywacja wymaga** realnych wartości w `appsettings.{ENV}.json`/`config.json` + (dla Graph) sekretu z GCP + (dla dual-auth) realnego Keycloaka. Graph używa app-only — wymaga uprawnień aplikacyjnych (`User.Read.All`) + admin consent. **Niezweryfikowane runtime** (brak tenanta/GCP/Keycloak lokalnie): start aplikacji z realnym Identity.Web, walidacja tokenów Entra, faktyczne mapowanie grup z realnego tokena, pobranie sekretu z GCP, selekcja schematu Keycloak.

**Hardening (analiza problemów):** (1) `RoleClaimType="roles"` ustawiony przez **`PostConfigure`** (rejestrowany po `AddMicrosoftIdentityWebApi`) — gwarantuje, że wygrywa nad konfiguracją Identity.Web, inaczej `RequireRole` mogłoby szukać `ClaimTypes.Role`. (2) `AuthSchemes.SelectByIssuer` w try/catch — niepoprawny token → fallback Entra (→401). (3) `RUNTIME_AUTH_CONFIG` ma root-factory default (=`environment.auth`) → injection zawsze się rozwiązuje (testy + bezpiecznik); `main.ts` nadpisuje wartością runtime; fetch `config.json` z fallbackiem przy 404/niepoprawnym JSON.

**Znane ograniczenia:** (a) **Group overage** — user w > ~200 grupach: Entra pomija claim `groups` (emituje `_claim_names`/`_claim_sources` → Graph). Transformer NIE rozwiązuje overage przez Graph → tacy użytkownicy nie dostają ról z grup; ratują App Roles (współistnieją) lub przyszłe rozwiązanie overage. (b) `groups` domyślnie niesie **object-id** (GUID), nie nazwy — konfiguracja Entra (optional claim) musi emitować nazwy `GSAPW4D_DOC2_*` albo w `RolesOptions.GroupNames` trzeba wpisać GUID-y. (c) start z PUSTYM `AzureAd:ClientId` (bazowy `appsettings.json` PRD) może rzucić walidacją Identity.Web przy pierwszym żądaniu — uruchamiać z env DEV (placeholdery niepuste) lub realną konfiguracją.

### Alternatives considered
Serwerowy OIDC (`AddMicrosoftIdentityWebApp` + cookie) jak w Doc2 — **pominięty**: ViewerEditor to SPA (interaktywny login robi MSAL w przeglądarce), backend pozostaje czystym resource-serverem. `Microsoft.Identity.Web.MicrosoftGraph` (Graph v4) — odrzucone na rzecz Graph v5 (zgodność z Doc2 5.103.0 + nowocześniejsze API). Delegated Graph zamiast app-only — odrzucone (lookup userów to funkcja admina, app-only prostsze).

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
