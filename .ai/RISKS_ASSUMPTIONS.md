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
