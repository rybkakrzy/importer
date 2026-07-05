# Diagram sekwencji: cykl pracy z dokumentem

> Źródło prawdy: `.ai/` + kod (stan 2026-05-26; zaktualizowano 2026-07-05: auth Entra ID, synchroniczna 1. próba „Zakończ" — 200 zamiast 202, podgląd ładuje v1). Elementy niepotwierdzone oznaczono **[Wymaga potwierdzenia]**.

## 1. Cel diagramu

Diagram pokazuje realny przepływ pracy z dokumentem w D2 ViewerEditor: od wejścia w URL, przez pobranie metadanych i treści, rozdzielenie ścieżek DOCX/PDF, edycję z auto-zapisem, aż po „Zakończ i wyślij" (finalizacja + asynchroniczna wysyłka). Przeznaczony dla developerów, architektów, DevOps i integratorów.

## 2. Zakres procesu

- wczytanie dokumentu (metadane + treść/plik),
- rozdział ścieżek DOCX (edytor) / PDF (podgląd),
- prezentacja klasyfikacji z metadanych,
- edycja DOCX,
- auto-zapis (nadpisanie v2 w miejscu) i ręczny „Zapisz",
- zakończenie pracy: finalizacja + utworzenie zadania wysyłki + polling statusu,
- asynchroniczna wysyłka przez worker na `ReturnUrl`,
- obsługa błędów na kluczowych etapach.

## 3. Założenia i źródła

**Pliki `.ai`:** `ARCHITECTURE.md`, `API_CONTRACTS.md`, `DOMAIN.md`, `DATABASE.md`, `DECISIONS.md` (ADR-0003/0005), `SECURITY.md`, `PROJECT_CONTEXT.md`, `FEATURES.md`.

**Potwierdzone w kodzie:**
- `pages/pdf-viewer/pdf-viewer.ts` — load przez `getDocument(masterId)` + niezależny `getDocumentMetadata(masterId)` (klasyfikacja).
- `components/document-editor/document-editor.ts` — `loadFromStorage`: `getDocumentMetadata` → decyzja PDF/DOCX → `downloadVersion`/`downloadBaseVersion` → `documentService.openDocument` (POST `/api/document/open`); `finishDocument()` z pollingiem; `documentClassification` z metadanych.
- `services/document-storage.service.ts` — `updateDocumentVersion` (PUT), `finishAndSend` (POST `.../finish`), `getDeliveryStatus`.
- `Api/Controllers/DocumentStorageController.cs` — endpointy `finish` (200, synchroniczna 1. próba), `abort-send`, `continue-delivery`, `deliveries/{id}`, `deliveries`, `retry`, `cancel`, `recipient-url`.
- `Infrastructure/Services/Delivery/*` + `DocumentDeliveryRepository` — worker, claim `FOR UPDATE SKIP LOCKED`, wysyłka HTTP.

**Rozstrzygnięte po snapshotcie (2026-07-05):**
- **Sprawdzenie dostępu / autoryzacja** — zaimplementowane: token Entra ID + role `Operator`/`Administrator` (ADR-0022) + `allowedCorporateKeys` (BR-014); A-05 zamknięte. Adnotacje „brak warstwy auth" w diagramach poniżej są historyczne.
- Krok 2 (podgląd): ładuje v1 przez `downloadBaseVersion` (R-02 zamknięte).
- POST zwrotny: `multipart/form-data` — części `file`/`masterId`/`versionId`/`corporateKey` (`HttpDeliverySender`).

---

## 4. Diagram sekwencji

### 4.1 Diagram główny (wysoki poziom)

```mermaid
sequenceDiagram
    autonumber
    actor U as Użytkownik
    box Frontend (Angular)
        participant GUI as SPA / Document Page
        participant SVC as document-storage.service
    end
    box Internal API (.NET 8)
        participant API as DocumentStorageController
        participant APP as Application (MediatR)
    end
    participant DB as PostgreSQL
    participant GCS as Google Cloud Storage

    U->>GUI: Wejście w URL ?masterId=[&versionId=]
    activate GUI
    GUI->>SVC: getDocumentMetadata(masterId)
    SVC->>API: GET /{masterId}/metadata
    API->>APP: GetDocumentMetadataQuery
    APP->>DB: GetByIdAsync(masterId)
    alt Dokument nie istnieje
        DB-->>APP: null
        APP-->>API: NotFound
        API-->>GUI: 404 { error }
        GUI-->>U: „Nie znaleziono dokumentu"
    else Dokument istnieje
        DB-->>APP: Document(metadata)
        APP-->>API: { mimeType, returnUrl?, classification? }
        API-->>GUI: 200 metadata
        Note over GUI: dostęp: token Entra + rola (401/403)<br/>+ allowedCorporateKeys (BR-014)
        GUI->>GUI: classification → badge (wspólny komponent)
        alt mimeType = PDF
            GUI->>GUI: render PDF (patrz 4.3)
        else mimeType = DOCX
            GUI->>GUI: edytor DOCX (patrz 4.2)
        end
    end
    deactivate GUI
```

**Objaśnienie:** wspólny punkt startowy dla obu typów to pobranie metadanych. mimeType decyduje o ścieżce; klasyfikacja jest prezentowana niezależnie od typu pliku.

---

### 4.2 Ścieżka DOCX — wczytanie, edycja, auto-zapis, zakończenie

```mermaid
sequenceDiagram
    autonumber
    actor U as Użytkownik
    participant ED as Document Editor (DOCX)
    participant SVC as document-storage.service
    participant DOCSVC as document.service
    participant API as Internal API
    participant APP as Application/Handler
    participant DOM as Document (Aggregate)
    participant DB as PostgreSQL
    participant GCS as GCS

    Note over ED,API: Wczytanie (po metadanych z diagramu głównego)
    alt Tryb edycji (versionId podany)
        ED->>SVC: downloadVersion(masterId, versionId)
        SVC->>API: GET /{masterId}/versions/{versionId}/download
    else Tryb podglądu (brak versionId)
        ED->>SVC: downloadBaseVersion(masterId)
        SVC->>API: GET /{masterId}/download
        Note over ED: podgląd bez versionId = wersja bazowa v1 (R-02 zamknięte)
    end
    API->>GCS: DownloadAsync(documents/{versionId})
    alt Błąd GCS / brak pliku
        GCS-->>API: error
        API-->>ED: 5xx/404
        ED-->>U: komunikat błędu pobrania
    else OK
        GCS-->>API: bytes
        API-->>ED: plik (blob)
        ED->>DOCSVC: openDocument(file)
        DOCSVC->>API: POST /api/document/open (DOCX→HTML)
        API-->>DOCSVC: DocumentContent (HTML)
        DOCSVC-->>ED: HTML + style + nagłówek/stopka
        ED-->>U: render edytora WYSIWYG
    end

    Note over U,DB: Edycja + auto-zapis (płaski zapis v2 — ADR-0003)
    loop Auto-save (timer, intervalSeconds=30, gdy włączony i jest versionId)
        U->>ED: zmiany treści
        ED->>SVC: updateDocumentVersion(masterId, versionId, content)
        SVC->>API: PUT /{masterId}/versions/{versionId}
        API->>APP: UpdateDocumentVersionCommand
        APP->>DOM: Document.UpdateVersion (BR-002: v1 → wyjątek/400)
        APP->>GCS: UploadAsync(documents/{versionId}) (nadpisanie)
        APP->>DB: SaveChanges (size, ModifiedAt, Status=Editing)
        alt Błąd zapisu (GCS/DB)
            APP-->>API: Failure
            API-->>ED: 4xx/5xx
            ED-->>U: autoSaveStatus = error
        else OK
            API-->>ED: UpdateDocumentVersionResult
            ED-->>U: autoSaveStatus = saved (HH:mm:ss)
        end
    end

    opt Ręczny „Zapisz"
        U->>ED: klik „Zapisz"
        ED->>SVC: PUT (versionId) lub POST /{masterId}/save
        SVC->>API: zapis wersji
        API-->>ED: wynik
    end

    Note over U,GCS: Zakończenie pracy — „Zakończ i wyślij" (patrz 4.4)
    U->>ED: klik „Zakończ"
    ED->>ED: finishDocument() → przejście do finalizacji
```

**Objaśnienie:** edytor pobiera plik z GCS, konwertuje DOCX→HTML przez bezstanowy `/api/document/open`, a auto-save nadpisuje tę samą wersję (v2) w miejscu — nie mnoży wersji.

---

### 4.3 Ścieżka PDF — podgląd + klasyfikacja

```mermaid
sequenceDiagram
    autonumber
    actor U as Użytkownik
    participant PV as PDF Viewer
    participant SVC as document-storage.service
    participant API as Internal API
    participant APP as Application/Handler
    participant DB as PostgreSQL
    participant GCS as GCS

    U->>PV: /viewer?masterId=
    activate PV
    par Treść PDF
        PV->>SVC: getDocument(masterId)
        SVC->>API: GET /{masterId}
        API->>APP: pobranie dokumentu + aktywnej wersji
        APP->>GCS: DownloadAsync(documents/{versionId})
        alt Błąd GCS / brak
            GCS-->>API: error
            API-->>PV: 5xx/404
            PV-->>U: „Nie można pobrać dokumentu"
        else OK
            GCS-->>APP: bytes
            APP-->>API: DocumentDto { name, content(base64) }
            API-->>PV: 200
            PV->>PV: pdfjs render (canvas + text layer)
            PV-->>U: podgląd PDF
        end
    and Klasyfikacja (niezależnie)
        PV->>SVC: getDocumentMetadata(masterId)
        SVC->>API: GET /{masterId}/metadata
        API-->>PV: { classification? }
        opt classification niepuste
            PV->>PV: badge klasyfikacji
        end
    end
    deactivate PV
```

**Objaśnienie:** treść i metadane (klasyfikacja) pobierane są niezależnie — awaria metadanych nie blokuje renderu PDF, a brak/puste `classification` nie renderuje pustego badge.

---

### 4.4 Zakończenie pracy: finalizacja, zadanie wysyłki, worker

```mermaid
sequenceDiagram
    autonumber
    actor U as Użytkownik
    participant ED as Editor (DOCX)
    participant SVC as document-storage.service
    participant API as Internal API
    participant APP as FinishAndSend Handler
    participant DB as PostgreSQL (document_deliveries)
    participant GCS as GCS
    participant W as DocumentDeliveryWorker
    participant R as ReturnUrl (odbiorca)

    U->>ED: „Zakończ"
    ED->>SVC: finishAndSend(masterId, versionId, content)
    SVC->>API: POST /{masterId}/versions/{versionId}/finish
    API->>APP: FinishAndSendDocumentCommand
    alt Brak/zły ReturnUrl w metadanych (BR-010)
        APP-->>API: Failure
        API-->>ED: 400 { error }
        ED-->>U: błąd finalizacji
    else OK
        APP->>GCS: UploadAsync(v2) + UploadRawAsync(deliveries/{id}) snapshot (SHA-256)
        APP->>DB: Status=Queued + INSERT document_deliveries(Sending, bez lease)
        Note over APP,DB: Idempotencja: unique partial index<br/>(jedno aktywne zadanie/dokument) → ponowny klik zwraca istniejące
        APP->>R: SYNCHRONICZNA 1. próba (IDeliverySender.SendAsync)
        alt Sukces
            APP->>DB: delivery.MarkSent + Document.Status=Sent
            APP-->>API: { deliveryId, delivered: true }
            API-->>ED: 200 { deliveryId, status, documentStatus, delivered }
        else Błąd 1. próby
            APP->>DB: HoldAfterFailedInlineAttempt (RetryScheduled „zaparkowane") + DeliveryFailed
            API-->>ED: 200 { delivered: false, error }
            ED-->>U: modal problemu: „Przerwij" (POST .../abort-send → Cancelled + SendAborted)<br/>lub „Kontynuuj w tle" (POST .../continue-delivery → Requeue + Queued)
        end
    end

    Note over W,R: Kontynuacja w tle (BackgroundService) — po „Kontynuuj w tle" lub retry
    W->>DB: ClaimDueBatch (FOR UPDATE SKIP LOCKED + lease)
    W->>GCS: DownloadAsync(deliveries/{id})
    W->>R: POST snapshot (Idempotency-Key=deliveryId, X-Content-SHA256)
    alt Sukces (2xx)
        R-->>W: 2xx
        W->>DB: MarkSent + Document.Status=Sent
    else Błąd retryable (5xx/timeout)
        R-->>W: błąd
        W->>DB: ScheduleRetry (backoff+jitter, cap 15 min)
        Note over W,DB: Po przekroczeniu 24 h → DeadLettered + DeliveryFailed
    else Błąd non-retryable (400/401/403/404/405/422)
        R-->>W: 4xx
        W->>DB: FailedPermanently + DeliveryFailed
    end
```

**Objaśnienie:** finalizacja jest atomowa (jeden `SaveChanges`: status dokumentu + zadanie). Wysyłany jest **niezmienny snapshot** zamrożony w chwili „Zakończ" (BR-012) — auto-save po zakończeniu nie zmienia wysyłanego pliku. Wysyłka jest at-least-once z `Idempotency-Key`.

---

## 5. Objaśnienie krok po kroku

**Wczytanie dokumentu:** GUI pobiera `GET /{masterId}/metadata`; brak dokumentu → 404. mimeType decyduje o ścieżce (PDF/DOCX); `classification` zasila wspólny badge.

**Sprawdzenie dostępu:** (zaktualizowane) całe Internal API wymaga tokenu Entra ID (401 bez tokenu) + roli aplikacyjnej `Operator`/`Administrator` (403, ADR-0022); endpointy treści/metadanych dodatkowo sprawdzają `allowedCorporateKeys` względem `CorporateKey` z claimu tokenu (403; admin omija — BR-014).

**Ścieżka DOCX:** pobranie pliku z GCS (`downloadVersion` w trybie edycji / `downloadBaseVersion` w podglądzie) → konwersja DOCX→HTML (`POST /api/document/open`) → render edytora.

**Ścieżka PDF:** `GET /{masterId}` (treść base64) → render pdfjs; metadane pobierane równolegle.

**Auto-zapis:** pętla `loop` (timer 30 s, tylko gdy włączony i jest `versionId`) → `PUT /{masterId}/versions/{versionId}` → nadpisanie v2 w GCS + `SaveChanges` (Status=Editing). v1 jest chroniona (BR-002).

**Zakończenie pracy:** `POST .../finish` → snapshot + **synchroniczna pierwsza próba** (200 `{ deliveryId, status, documentStatus, delivered, error? }`); sukces → `Sent`, błąd → decyzja użytkownika: `POST .../abort-send` (przerwij) lub `POST .../continue-delivery` (worker dokańcza w tle, retry do 24 h).

**Obsługa błędów:** 404 (brak dokumentu), 5xx/404 (błąd GCS/brak pliku), 4xx/5xx (błąd zapisu → `autoSaveStatus=error`), 400 (zły/brak `ReturnUrl` przy finalizacji), klasyfikacja błędów wysyłki (retry vs `FailedPermanently` vs `DeadLettered`).

---

## 6. Warianty procesu

| Wariant | Zachowanie |
|---|---|
| Dokument DOCX | edytor + auto-save + finalizacja (4.2, 4.4) |
| Dokument PDF | podgląd read-only + klasyfikacja (4.3) |
| Użytkownik bez dostępu | **[Wymaga potwierdzenia]** brak warstwy auth — obecnie nieobsługiwane odrębnie |
| Dokument bez klasyfikacji | badge nie renderuje się (brak/empty/whitespace) — brak błędu |
| Błąd GCS | komunikat pobrania; brak utraty danych |
| Błąd zapisu | `autoSaveStatus=error`; możliwy ręczny „Zapisz"/ponowienie |
| Zakończenie z wysyłką | DOCX z poprawnym `ReturnUrl` → zadanie + worker |
| Zakończenie bez `ReturnUrl` | 400 — finalizacja odrzucona (BR-010) |

---

## 7. Punkty integracyjne

**Endpointy (Internal API `/api/documentstorage`):** `GET /{masterId}/metadata`, `GET /{masterId}`, `GET /{masterId}/download`, `GET /{masterId}/versions/{versionId}/download`, `PUT /{masterId}/versions/{versionId}`, `POST /{masterId}/save`, `POST /{masterId}/versions/{versionId}/finish`, `GET /deliveries/{deliveryId}`, `GET /deliveries?status=`, `POST /deliveries/{deliveryId}/retry`. Bezstanowy: `POST /api/document/open`.

**Frontend:** `pages/pdf-viewer`, `components/document-editor`, `components/document-classification-badge`, `services/{document-storage.service, document.service}`.

**Backend handlery:** `GetDocumentMetadataQuery`, `UpdateDocumentVersionCommand`, `FinishAndSendDocumentCommand`, `GetDeliveryStatusQuery`, `RequeueDeliveryCommand`.

**Storage/DB/Worker:** GCS (`documents/{versionId}`, `deliveries/{deliveryId}`), PostgreSQL (`documents`, `document_versions`, `document_deliveries`), `DocumentDeliveryWorker`.

**System zewnętrzny:** odbiorca `ReturnUrl` (POST snapshotu z `Idempotency-Key`).

---

## 8. Rekomendacje dotyczące czytelności procesu

- **Dostęp:** (zrobione) auth Entra ID + role + `allowedCorporateKeys` — patrz `SECURITY.md`, ADR-0011/0012/0022.
- **Krok 2 (podgląd):** (zrobione) tryb read-only ładuje v1 przez `/download` (R-02 zamknięte).
- **Logi/correlation:** propagować `correlation_id` z `document_deliveries` do logów finalizacji i wysyłki dla śledzenia end-to-end.
- **Kontrakt zwrotny:** udokumentować `Content-Type`/body POST-u na `ReturnUrl` (potwierdzić w `HttpDeliverySender`).
```
