# Risks and Assumptions

## Cel pliku

Niepewności, ryzyka i założenia. Aktualizuj, gdy czegoś nie da się potwierdzić w repo.

## Aktywne założenia

| ID | Założenie | Wpływ | Jak potwierdzić | Status |
|---|---|---|---|---|
| A-01 | External API (`D2ServicesViewerEditor`) działa na porcie 15112 | Medium | `launchSettings.json` pokazuje 5000/7000; 15112 z rozmowy. Sprawdzić konfigurację deploymentu / reverse proxy | Open |
| A-02 | Lokalny PostgreSQL i GCS uruchamiane są poza repo (brak docker-compose) | Medium | Potwierdzić procedurę bootstrapu z zespołem | Open |
| A-03 | Schemat bazy bootstrapowany skryptami `infra/sql/` (001..005) w kolejności | Medium | Sprawdzić, czy istnieje runner migracji poza repo | Open |
| A-04 | Wersja edytowalna (v2) to dosłowna kopia bajtów oryginału, nie konwersja | Low | Kod `IngestExternalDocumentCommandHandler` (kopia bajtów) — potwierdzić wymaganie | Open |
| A-05 | Mechanizm auth — niepotwierdzony w tej sesji | Medium | Przejrzeć `Program.cs`/middleware obu API | Open |

## Aktywne ryzyka

| ID | Ryzyko | Pr. | Wpływ | Mitigacja | Status |
|---|---|---|---|---|---|
| R-01 | Brak CI/CD i docker-compose w repo → niespójny build/deploy między osobami | Medium | Medium | Spisać i dodać pipeline + compose | Open |
| R-02 | Tryb podglądu (Krok 2) ładuje aktywną wersję (v2), nie v1 → użytkownik podglądu widzi edytowalną, nie oryginał | Medium | Medium | Przełączyć GUI na `GET .../{masterId}/download` dla read-only | Open |
| R-03 | Auto-save i ręczny „Zapisz" dzielą ścieżkę; przy braku versionId zapis idzie jako nowa wersja (POST) | Low | Low | Świadome zachowanie; udokumentowane w `FEATURES.md` | Open |
| R-04 | Port External API niepewny (A-01) → błędne URL integracji | Medium | Medium | Potwierdzić źródło prawdy portu | Open |
| R-05 | Nieaktualizowana pamięć `.ai/` może wprowadzać w błąd | Medium | High | Aktualizować `CURRENT_STATE`/`TASK_HANDOFF`/`CHANGELOG` po pracy | Open |

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
