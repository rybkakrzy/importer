# Risks and Assumptions

## Cel pliku

Niepewności, ryzyka i założenia. Aktualizuj, gdy czegoś nie da się potwierdzić w repo.

## Aktywne założenia

| ID | Założenie | Wpływ | Jak potwierdzić | Status |
|---|---|---|---|---|
| A-01 | External API (`D2ServicesViewerEditor`) działa na porcie 15112 | — | **Potwierdzone**: `appsettings.json` `Urls=https://0.0.0.0:15112`, `appsettings.local.json` `http://localhost:15112`. Swagger: `/swagger` | Closed |
| A-02 | Lokalny PostgreSQL i GCS uruchamiane są poza repo (brak docker-compose) | Medium | Potwierdzić procedurę bootstrapu z zespołem | Open |
| A-03 | Schemat bazy bootstrapowany skryptami `infra/sql/` (001..007) w kolejności | Medium | Sprawdzić, czy istnieje runner migracji poza repo | Open |
| A-06 | Odbiorca `returnUrl` akceptuje POST z plikiem (octet-stream) i obsługuje `Idempotency-Key` do deduplikacji | Medium | Potwierdzić kontrakt z zespołem aplikacji źródłowej | Open |
| A-04 | Wersja edytowalna (v2) to dosłowna kopia bajtów oryginału, nie konwersja | Low | Kod `IngestExternalDocumentCommandHandler` (kopia bajtów) — potwierdzić wymaganie | Open |
| A-05 | Mechanizm auth — niepotwierdzony w tej sesji | Medium | Przejrzeć `Program.cs`/middleware obu API | Open |

## Aktywne ryzyka

| ID | Ryzyko | Pr. | Wpływ | Mitigacja | Status |
|---|---|---|---|---|---|
| R-01 | Brak CI/CD i docker-compose w repo → niespójny build/deploy między osobami | Medium | Medium | Spisać i dodać pipeline + compose | Open |
| R-02 | Tryb podglądu (Krok 2) ładuje aktywną wersję (v2), nie v1 → użytkownik podglądu widzi edytowalną, nie oryginał | Medium | Medium | Przełączyć GUI na `GET .../{masterId}/download` dla read-only | Open |
| R-03 | Auto-save i ręczny „Zapisz" dzielą ścieżkę; przy braku versionId zapis idzie jako nowa wersja (POST) | Low | Low | Świadome zachowanie; udokumentowane w `FEATURES.md` | Open |
| R-04 | Port External API niepewny → błędne URL integracji | — | — | Zamknięte: port 15112 potwierdzony w `appsettings.json` (`Urls`) | Closed |
| R-05 | Nieaktualizowana pamięć `.ai/` może wprowadzać w błąd | Medium | High | Aktualizować `CURRENT_STATE`/`TASK_HANDOFF`/`CHANGELOG` po pracy | Open |
| R-06 | Brak testów integracyjnych claimu wysyłki (`FOR UPDATE SKIP LOCKED` / reclaim po crashu) na realnym PostgreSQL — pokrycie tylko jednostkowe | Medium | Medium | Dodać testy integracyjne repo (Testcontainers/lokalny Postgres) | Open |
| R-07 | `RecipientUrl` z metadanych zewnętrznych użyty do żądań serwerowych workera → ryzyko SSRF (brak allowlisty/blokady adresów prywatnych) | Medium | Medium | Dodać walidację hosta/schematu poza samym http(s) | Open |
| R-08 | Ręczny „Zapisz"/auto-save może nadpisać v2 po utworzeniu zadania wysyłki — wysłany jest jednak niezmienny snapshot, więc treść wysyłki się nie zmienia (świadome) | Low | Low | Udokumentowane (BR-012) | Open |
| R-09 | `WysiwygEditorComponent.sanitizeHtml` (zwykłe Ctrl+V) to sanitizacja regexem — słaba na XSS (np. `<img onerror>` po dziwnym whitespace, `javascript:` w href, osadzone SVG). Treść trafia do `bypassSecurityTrustHtml`. | Medium | High | Wprowadzić DOMPurify (whitelist tagów/atrybutów/protokołów) — wymaga decyzji o zależności; sanitizacja także po stronie backendu przed zapisem/serwowaniem | Open |
| R-10 | Nagłówek/stopka: `ExtractHeader/ExtractFooter` biorą `HeaderParts.FirstOrDefault()` — tylko PIERWSZY part, bez rozróżnienia default/first-page/even-odd i bez obsługi wielu sekcji. Dokumenty z innym nagłówkiem pierwszej strony / parzystych-nieparzystych / wieloma sekcjami pokażą tylko jeden nagłówek/stopkę. | Medium | Medium | **ZAMKNIĘTE dla single-section (2026-05-28):** import czyta wariant **default** (refs sekcji), **first-page** (`titlePg`) i **even/odd** (`evenAndOddHeaders`); model domeny rozszerzony o `DifferentOddEven`/`EvenHtml`; front spread'uje cały obiekt header/footer. **Round-trip na zapisie:** `HtmlToDocxConverter` zapisuje wszystkie 3 warianty jako osobne `HeaderPart`/`FooterPart` typu Default/First/Even + ustawia `titlePg` w sekcji i `evenAndOddHeaders` w settings (testy `HeaderFooterRoundTripTests` 6/6). **Pozostaje OPEN:** wiele sekcji (`sectPr` per-section, każda ze swoimi refs) — brany pierwszy `SectionProperties` z body. | Open (multi-section) |
| R-11 | Round-trip nagłówka/stopki: rozmiar z `docDefaults` jest niesiony przez kontener `.header-footer-content` (nie per-run `w:sz`). Po zapisie HTML→DOCX wrapper jest spłaszczany; przy ponownym otwarciu rozmiar zależy od `docDefaults` regenerowanego DOCX. Jeśli się różni, rozmiar może dryfować. | Low | Medium | Front zapewnia spójny fallback (11pt). Docelowo: zachować `w:sz` per-run lub wymusić docDefaults przy zapisie. | Open |
| R-12 | Snapshoty Golden (Etap 0) lokują **obecne** wyjście HTML konwertera, które miejscami nie jest zgodne z Wordem. To regression-guard, nie golden-correctness — łatwo pomylić „snapshot zielony" z „wiernością z Wordem". | Low | Medium | ADR-0008 dokumentuje intencję; po świadomej poprawie wierności (Etap 3–8) baseline aktualizować (`delete *.approved.html` + re-run) i przeglądać diff w review. | Open |
| R-13 | Etap 1 zmienił cm↔twips z przybliżenia `567` na dokładny `1440/2.54` (566.929). Dla marginesów po zaokrągleniu do 0.01 cm różnica jest pomijalna (testy tolerancyjne zielone), ale zapis (`HtmlToDocxConverter.CmToTwips`) może dać o 1 twip inny wynik niż wcześniej (np. 2.5 cm: 1418→1417 tw). | Low | Low | Brak testu na dokładne twipsy marginesów; różnica < 0,002 cm. Monitorować przy round-tripie marginesów wielu sekcji (Etap 4). | Open |
| R-14 | Etap 4 backend: reader zwraca już `DocumentContent.PageSize`, a writer potrafi go zapisać (`Convert(... pageSize)`), ale handlery Save/Sign/Download nie przekazują jeszcze `PageSize`. | Medium | Medium | **ZAMKNIĘTE 2026-06-03:** pełny full-stack — `PageSize` w 3 komendach/handlerach/kontrolerach + `document.model.ts` + sygnał `documentPageSize` (load capture + `buildSaveRequest`) + 2 testy Vitest. Round-trip działa; brak `pageSize` → fallback A4. | Closed |
| R-15 | Round-trip nie przywraca pojedynczego **oryginalnego** twardego page-breaka (`<w:br type=page>`): reader emituje `<div class="page-break">`, ale writer go nie konwertuje z powrotem (AFTER=0 vs GOOD=1). Po naprawie marginesów/tabel Word paginuje naturalnie, więc wpływ mały. | Low | Low | Dodać w writerze mapowanie `div.page-break` → `w:br type=page`; test round-trip page-break. | Open |
| R-16 | Model `DocumentContent` jest **stratny**: nie niesie `styles.xml` (writer regeneruje 16 stylów vs 164 w oryginale, gubi 100 stylów tabel + latent), `numbering.xml`, theme; font treści `Cambria`(theme minor)→`Calibri`. Wpływa na fidelity (wygląd), nie na liczbę stron tego dokumentu (`numPr`=0). | Medium | Medium | Docelowo: przenosić oryginalne party (`styles.xml`/`numbering.xml`/theme) przez model albo zachowywać oryginalny pakiet i wstrzykiwać tylko zmienioną treść (pass-through). Patrz `DOCX_CONVERSION.md`. | Open |
| R-17 | Tabela dzielona przez edytor między strony (`_splitTableForPagination`) zapisywana jest jako 2 sąsiednie `<table>` zamiast jednej (pre-existing; `getContent` plain-concat nie pogarsza, usuwa tylko nadmiarowy break). | Low | Medium | Scalać kolejne fragmenty tej samej tabeli przy serializacji do zapisu (po `id`/kolejności). | Open |
| R-18 | Wysokości wierszy tabel są „zapiekane" jako `tr style="height:px"` z `getBoundingClientRect` w `_serializeSingleEditor` → zależne od renderu edytora (line-height/CSS). Po naprawie paddingu (0 góra/dół) wiersze są ciaśniejsze, ale wartość nadal pochodzi z DOM, nie z DOCX. | Low | Low | Rozważyć zapis `trHeight` tylko gdy oryginał miał jawną wysokość; w przeciwnym razie pominąć (Word policzy z treści). | Open |

## Zamknięte

| ID | Opis | Wynik | Data |
|---|---|---|---|
| — | brak | — | 2026-05-23 |

## Kiedy dodawać wpis

- agent zgaduje wersję/port/konfigurację,
- brak testów dla zmienianego obszaru,
- zmiana może być breaking,
- decyzja zależy od środowiska produkcyjnego,
- ryzyko migracji danych,
- niejasne wymaganie biznesowe.
