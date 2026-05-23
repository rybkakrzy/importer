# Product Goals

## Cel produktu

Umożliwić aplikacjom zewnętrznym oddanie dokumentu (DOCX/PDF) do bezpiecznego podglądu i edycji w przeglądarce, z zachowaniem oryginału i kontrolowanym zwrotem zmienionej wersji — bez konieczności posiadania lokalnego pakietu biurowego.

## Problem, który rozwiązujemy

- edycja dokumentów Word w przeglądarce bez psucia oryginału (osobna wersja edytowalna),
- spójny tryb podglądu PDF i edycji DOCX z jednego GUI,
- integracja z systemami źródłowymi (ingest + URL zwrotny + klasyfikacja dokumentu),
- wersjonowanie i możliwość przywrócenia wersji,
- podpisy cyfrowe i kody kreskowe wewnątrz dokumentu.

## Najważniejsze wartości dla użytkownika

- Oryginał (v1) zawsze nietknięty.
- Auto-save bez mnożenia wersji (nadpisanie v2 w miejscu) — brak zaśmiecania GCS.
- Przewidywalny tryb: podgląd vs edycja zależny od URL.
- Dobre komunikaty stanu (zapisywanie / zapisano / błąd).

## Cele krótkoterminowe

| Cel | Status | Uwagi |
|---|---|---|
| Ingest zewnętrzny (Krok 1) | Implemented | `D2ServicesViewerEditor` + `IngestExternalDocumentCommand` |
| Płaski zapis edytowalnej wersji + auto-save | Implemented | `PUT .../versions/{versionId}`, switch AutoSave w GUI |
| Endpoint metadanych (returnUrl, classification) | Implemented | `GET .../{masterId}/metadata` |
| Tryb podglądu (Krok 2) — ładowanie v1 dla DOCX | In Progress | GUI wciąż ładuje aktywną wersję; patrz RISKS/FEATURES |
| Funkcja „Zakończ" (zwrot pliku na returnUrl) | Planned | `finishDocument()` w GUI to TODO |

## Cele długoterminowe

| Cel | Status | Uwagi |
|---|---|---|
| Spójność architektury backend/frontend | Active | stosować istniejące wzorce |
| Minimalizacja regresji przez testy | Active | NUnit + Vitest |
| Bezpieczny, powtarzalny deployment | Planned | brak compose/CI w repo — do uzupełnienia |

## Poza zakresem bez jawnej decyzji

- redesign UI, zmiana frameworka frontu, zmiana architektury backendu,
- migracja bazy / zmiana provider'a auth,
- breaking changes w API,
- aktualizacja major .NET / Angular,
- eksport do PDF (obecnie `501 Not Implemented` — celowy placeholder).

## Priorytety przy konflikcie

1. Poprawność domenowa. 2. Bezpieczeństwo. 3. Stabilność/kompatybilność. 4. Testowalność. 5. Czytelność. 6. Wydajność. 7. Estetyka.
