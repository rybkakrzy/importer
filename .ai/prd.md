# Dokument wymagań produktu (PRD) - D2ViewerEditor

## 1. Przegląd produktu

D2ViewerEditor to przeglądarkowy edytor DOCX z konwersją po stronie backendu oraz wersjonowaniem dokumentów.
Produkt łączy frontend Angular i API ASP.NET Core, aby zapewnić:

- cykl open/edit/save dla DOCX
- storage i historię wersji
- wstawianie kodów barcode/QR
- wsparcie podpisów cyfrowych

Główni użytkownicy: zespoły biznesowe, które potrzebują kontrolowanej edycji dokumentów bez zależności od lokalnego Office.

## 2. Problemy użytkowników

- Edycja DOCX w przeglądarce często psuje formatowanie albo wymaga kosztownych narzędzi.
- Zarządzanie wersjami bywa ręczne i podatne na błędy.
- Przepływy dotyczące podpisów i barcode są rozproszone między wieloma narzędziami.

## 3. Wymagania funkcjonalne

- FR-001: Użytkownik może uploadować DOCX i otrzymać edytowalną treść dokumentu w HTML.
- FR-002: System waliduje rozszerzenie uploadowanego pliku i pusty payload.
- FR-003: Użytkownik może zapisać edytowany HTML i pobrać poprawny DOCX.
- FR-004: Użytkownik może utworzyć nowy pusty dokument.
- FR-005: Użytkownik może uploadować obraz i wstawiać go do dokumentu.
- FR-006: Użytkownik może pobrać listę dostępnych typów barcode.
- FR-007: Użytkownik może generować barcode/QR jako payload PNG.
- FR-008: Użytkownik może podpisać dokument danymi certyfikatu.
- FR-009: Użytkownik może zweryfikować podpisy w uploadowanym DOCX.
- FR-010: Użytkownik może uploadować oryginalny dokument do storage i otrzymać `masterId` + `versionId`.
- FR-011: Użytkownik może zapisać nową wersję dla istniejącego `masterId`.
- FR-012: System utrzymuje dokładnie jedną aktywną wersję na dokument.
- FR-013: Użytkownik może listować wszystkie wersje dokumentu posortowane malejąco po version number.
- FR-014: Użytkownik może przywrócić dowolną poprzednią wersję jako aktywną.
- FR-015: Użytkownik może pobrać binarki dokumentu: oryginalną/base i konkretną wersję.
- FR-016: Użytkownik może listować dostępne szablony dokumentów.
- FR-017: Użytkownik może wczytać szablon po id.
- FR-018: Endpoint health zwraca status, environment, build number i build date.
- FR-019: API zwraca spójne payloady błędów dla walidacji i błędów biznesowych.
- FR-020: System wspiera profile appsettings DEV, UAT i PRD.

## 4. Wymagania niefunkcjonalne

- NFR-001: API musi działać na .NET 8.
- NFR-002: Kluczowe endpointy dokumentowe powinny mieć dostępność 99.5% w UAT/PRD.
- NFR-003: Duże uploady muszą mieć limity (50 MB dla open, 10 MB dla upload image).
- NFR-004: Trwałość danych musi przetrwać restarty kontenerów (volumes Postgres).
- NFR-005: Operacje wersjonowania muszą być atomowe i spójne.

## 5. Granice produktu

Poza aktualnym zakresem:

- Real-time co-authoring
- Track changes and comments parity with desktop Word
- Full-text search indexing across all stored documents
- Full PDF export pipeline
- Tenant-level IAM/SSO admin console

## 6. User stories

- US-001: Jako editor chcę otworzyć DOCX w przeglądarce, aby szybko modyfikować treść.
- US-002: Jako editor chcę zapisać zmiany HTML z powrotem do DOCX, aby plik pozostał kompatybilny z Word.
- US-003: Jako operator chcę, aby każdy save tworzył nową wersję, żebym mógł śledzić zmiany.
- US-004: Jako operator chcę przywrócić starszą wersję, aby odzyskać dokument po błędzie.
- US-005: Jako użytkownik chcę wstawiać barcode/QR, aby generować dokumenty gotowe do druku.
- US-006: Jako użytkownik compliance chcę podpisywać i weryfikować dokumenty cyfrowo.
- US-007: Jako support engineer chcę mieć metadata health/build, aby szybko diagnozować wdrożenia.

## 7. Kryteria akceptacji (krytyczne przepływy)

### Open -> Edit -> Save
- Given poprawny upload DOCX, when użytkownik go otwiera, then API zwraca model dokumentu z HTML.
- Given zmodyfikowany HTML, when użytkownik zapisuje, then przeglądarka otrzymuje DOCX do pobrania.

### Versioning
- Given istniejący `masterId`, when użytkownik zapisuje wersję, then version number zwiększa się o 1.
- Given przywrócona wersja, when użytkownik pobiera aktywny dokument, then przywrócona wersja jest aktywna.

### Barcode
- Given poprawny barcode type i content, when użytkownik wywołuje generate, then API zwraca payload PNG.

### Signature
- Given poprawne cert data i hasło, when użytkownik podpisuje, then zwracany jest podpisany DOCX.

## 8. Metryki sukcesu

- Współczynnik powodzenia open/save dokumentu >= 95% w zestawie pilotażowym
- Współczynnik powodzenia restore version = 100% w automatycznej regresji
- Czas odpowiedzi P95 dla `/api/document/open` < 2.5s (zestaw plików testowych)
- Czas odpowiedzi P95 dla `/api/document/save` < 2.0s
- < 1% nieudanych requestów w profilu testu obciążeniowego `all_endpoints_locust.py`

