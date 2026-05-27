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
| R-10 | Nagłówek/stopka: `ExtractHeader/ExtractFooter` biorą `HeaderParts.FirstOrDefault()` — tylko PIERWSZY part, bez rozróżnienia default/first-page/even-odd i bez obsługi wielu sekcji. Dokumenty z innym nagłówkiem pierwszej strony / parzystych-nieparzystych / wieloma sekcjami pokażą tylko jeden nagłówek/stopkę. | Medium | Medium | **CZĘŚCIOWO ZAMKNIĘTE (2026-05-28):** import rozwiązuje part przez `sectPr` refs typu **Default** (fallback `FirstOrDefault` gdy brak refs), **first-page** przez `titlePg` → `DifferentFirstPage`/`FirstPageHtml`, oraz **even/odd** przez `settings/evenAndOddHeaders` → `DifferentOddEven`/`EvenHtml` (model domeny rozszerzony; front spread'uje cały obiekt). **Pozostaje OPEN:** (a) **round-trip na zapisie** — `HtmlToDocxConverter.AddHeaderAndFooter` zapisuje tylko default, nie serializuje first/even do osobnych partów + `titlePg`/`evenAndOddHeaders`; (b) **wiele sekcji** (`sectPr` per-section). | Open (partial) |
| R-11 | Round-trip nagłówka/stopki: rozmiar z `docDefaults` jest niesiony przez kontener `.header-footer-content` (nie per-run `w:sz`). Po zapisie HTML→DOCX wrapper jest spłaszczany; przy ponownym otwarciu rozmiar zależy od `docDefaults` regenerowanego DOCX. Jeśli się różni, rozmiar może dryfować. | Low | Medium | Front zapewnia spójny fallback (11pt). Docelowo: zachować `w:sz` per-run lub wymusić docDefaults przy zapisie. | Open |

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
