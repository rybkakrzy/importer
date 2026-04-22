# MVP - D2ViewerEditor

## Główny problem

Potrzebują otwierać, edytować i zapisywać dokumenty DOCX w przeglądarce bez desktopowego Office, przy jednoczesnym zachowaniu struktury dokumentu i wsparciu scenariuszy regulowanych (podpisy, śledzalność, wersje).

## Minimalny zestaw funkcji

### 1. Otwieranie i edycja DOCX
- Upload pliku `.docx` z interfejsu (`/api/document/open`)
- Konwersja DOCX -> edytowalny HTML w przeglądarce
- Edycja WYSIWYG z podstawowym formatowaniem (pogrubienie, kursywa, podkreślenie, listy, wyrównanie)

### 2. Zapis z powrotem do DOCX
- Konwersja edytowanego HTML -> DOCX (`/api/document/save`)
- Pobranie wygenerowanego pliku DOCX
- Zachowanie kluczowych metadanych, nagłówka/stopki i marginesów

### 3. Przechowywanie dokumentów z wersjonowaniem
- Upload początkowej binarki do storage (`/api/documentstorage/upload`)
- Zapis tworzy kolejną wersję (`/api/documentstorage/{masterId}/save`)
- Odczyt aktywnej wersji i historii wersji
- Przywracanie poprzedniej wersji (`/api/documentstorage/{masterId}/restore/{versionId}`)

### 4. Wstawianie kodów kreskowych / QR
- Pobranie wspieranych typów kodów (`/api/barcode/types`)
- Generowanie obrazu kodu kreskowego (`/api/barcode/generate`)
- Wstawienie wygenerowanego obrazu do treści dokumentu

### 5. Podpisy cyfrowe (podstawowy przepływ)
- Podpisanie dokumentu certyfikatem X.509 (`/api/document/sign`)
- Weryfikacja podpisów z uploadowanego DOCX (`/api/document/verify-signatures`)

### 6. Podstawowe szablony
- Odczyt listy szablonów (`/api/document/templates`)
- Otwarcie wybranego szablonu (`/api/document/templates/{templateId}`)

### 7. Gotowość operacyjna
- Endpoint health z metadanymi builda (`/api/health`)
- Włączony Swagger do testowania API
- Środowisko DEV przez Podman compose (Postgres + fake-gcs + API + GUI)

## Poza zakresem MVP

- Edycja współbieżna w czasie rzeczywistym (wielu użytkowników, równoległe kursory)
- Pełna implementacja eksportu PDF (endpoint istnieje, zwraca 501)
- Zaawansowany silnik workflow (ścieżki akceptacji)
- Kontrola dostępu oparta o role i integracja SSO
- Dashboard audytowy i analityka długoterminowa
- Tryb edycji offline
- Natywna aplikacja mobilna

## Kryteria sukcesu MVP

- 95% plików DOCX użytych w pilocie otwiera się i zapisuje poprawnie
- Przywracanie wersji działa w 100% testowanych scenariuszy
- Mediana czasu API dla generowania kodu kreskowego < 300 ms w DEV/UAT
- Opóźnienie P95 API dla kluczowych operacji edytora < 1.5 s w UAT
- Zero krytycznych incydentów utraty danych podczas zapisu/przywracania

