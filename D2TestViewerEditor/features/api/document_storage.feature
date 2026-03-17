# language: pl
Właściwość: Document Storage API — wersjonowanie dokumentów
  Jako klient API
  Chcę przechowywać dokumenty z historią wersji
  Aby móc przywracać poprzednie wersje i śledzić historię zmian

  # ────────────────────────────────────────────
  # Smoke — upload i podstawowy CRUD
  # ────────────────────────────────────────────

  @api @smoke
  Scenariusz: Upload nowego dokumentu kończy się sukcesem
    Kiedy wysyłam uploadDocument z nazwą "smoke_test.txt" i typem "text/plain"
    Wtedy status odpowiedzi to 200
    I odpowiedź zawiera pole "masterId"
    I odpowiedź zawiera pole "versionId"
    I odpowiedź zawiera pole "fileName"

  @api @smoke
  Scenariusz: Pobranie przesłanego dokumentu po masterId
    Zakładając że przesłano dokument o nazwie "retrieve_test.pdf" i typie "application/pdf"
    Kiedy pobieram dokument po masterId
    Wtedy status odpowiedzi to 200
    I odpowiedź zawiera pole "name"
    I odpowiedź zawiera pole "content"
    I odpowiedź zawiera pole "versionNumber"
    I wersja dokumentu wynosi 1

  # ────────────────────────────────────────────
  # Wersjonowanie
  # ────────────────────────────────────────────

  @api @regression
  Scenariusz: Zapis nowej wersji zwiększa numer wersji
    Zakładając że przesłano dokument o nazwie "versioning.docx" i typie "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    Kiedy zapisuję nową wersję dokumentu
    Wtedy status odpowiedzi to 200
    I odpowiedź zawiera pole "versionNumber"
    I numer wersji wynosi 2

  @api @regression
  Scenariusz: Pobranie historii wersji zwraca posortowaną listę
    Zakładając że przesłano dokument o nazwie "history_test.txt" i typie "text/plain"
    I zapisano 2 dodatkowe wersje dokumentu
    Kiedy pobieram listę wersji dokumentu
    Wtedy status odpowiedzi to 200
    I odpowiedź jest listą
    I lista zawiera co najmniej 3 element
    I wersje są posortowane malejąco
    I tylko najnowsza wersja jest aktywna

  @api @regression
  Scenariusz: Przywrócenie poprzedniej wersji
    Zakładając że przesłano dokument o nazwie "restore_test.md" i typie "text/markdown"
    I zapisano 2 dodatkowe wersje dokumentu
    Kiedy przywracam pierwszą wersję dokumentu
    Wtedy status odpowiedzi to 200
    I odpowiedź zawiera pole "message"

  # ────────────────────────────────────────────
  # Obsługa błędów
  # ────────────────────────────────────────────

  @api @regression
  Scenariusz: Pobranie nieistniejącego dokumentu zwraca 404
    Kiedy wysyłam GET na "/documentstorage/00000000-0000-0000-0000-000000000001"
    Wtedy status odpowiedzi to 404

  @api @regression
  Scenariusz: Upload z nieprawidłowym typem MIME zwraca 400
    Kiedy wysyłam uploadDocument z nazwą "bad.txt" i typem "invalid/mime-type"
    Wtedy status odpowiedzi to 400

  @api @regression
  Scenariusz: Upload z pustą zawartością zwraca 400
    Kiedy wysyłam uploadDocument z pustą zawartością
    Wtedy status odpowiedzi to 400
