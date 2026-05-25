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
