# Flow: Współbieżne otwieranie dokumentu w systemie D2 ViewerEditor

> Dokument projektowy opracowany z perspektywy doświadczonego architekta.
> Opisuje pełen przepływ: upload przez system zewnętrzny → otwarcie przez
> użytkownika U1 → równoległe otwarcie przez użytkownika U2.

## 1. Upload pliku przez aplikację zewnętrzną → otrzymanie GUID

**Endpoint:** `POST /api/documents` (multipart: plik + metadane JSON)

**Krok po kroku:**
1. Aplikacja zewnętrzna przesyła `file` (DOCX) + `metadata` (autor, tytuł, system źródłowy, tagi itd.) wraz z nagłówkiem `Idempotency-Key` (żeby retry nie tworzyło duplikatów).
2. API waliduje: typ MIME, rozmiar, schemat metadanych, antywirus (opcjonalnie ClamAV w `docker/`).
3. Backend tworzy:
   - rekord `Document` z `DocumentId = GUID` (to jest **stabilny identyfikator logicznego dokumentu**, nie wersji),
   - rekord `DocumentVersion` z `VersionId = GUID`, `VersionNumber = 0` (wersja bazowa, immutable), `Hash = SHA-256(content)`, `CreatedBy`, `CreatedAt`,
   - zapis binarki do storage'u (np. blob / filesystem), kluczem jest `VersionId`.
4. Odpowiedź: `201 Created` z body `{ documentId, currentVersionId, versionNumber: 0, etag }` + nagłówek `Location: /api/documents/{documentId}`.

**Co zwracamy zewnętrznej aplikacji?** Wyłącznie `documentId` (GUID logiczny). Zewnętrzny system nie powinien znać GUID-ów wersji — to szczegół wewnętrzny.

---

## 2. Użytkownik U1 otwiera dokument — `GET /api/documents/{documentId}`

**Logika serwera:**
1. Pobierz `Document` po `documentId` → 404 jeśli brak, 403 jeśli brak uprawnień.
2. Wyznacz **aktualną opublikowaną wersję** (`HEAD`) — to wersja bazowa v0.
3. Sprawdź, czy istnieje aktywny **lock / edit session** na tym dokumencie (tabela `DocumentEditSession`: `documentId`, `ownerUserId`, `sessionId = GUID`, `startedAt`, `expiresAt`, `lastHeartbeatAt`, `draftVersionId`).
4. Brak aktywnej sesji → utwórz nową:
   - `sessionId = GUID`, `ownerUserId = U1`, TTL np. 5 min, odnawiany heartbeatem co 30 s,
   - utwórz **draft** = `DocumentVersion` `VersionNumber = 1` w stanie `Draft`, sklonowany z v0; jego `VersionId` jest **prywatny** dla sesji edycyjnej.
5. Zwróć U1:
   ```json
   {
     "documentId": "...",
     "versionId": "<v0 lub draft v1 jeśli edytor>",
     "versionNumber": 0,
     "mode": "edit",
     "editSession": { "sessionId": "...", "expiresAt": "..." },
     "etag": "..."
   }
   ```

**Kluczowa decyzja projektowa:** rozróżniamy dwa tryby otwarcia:
- `?mode=view` — czytanie, zwraca ostatnią **opublikowaną** wersję (v0), bez locka.
- `?mode=edit` — edycja, próbuje wziąć lock i utworzyć / odzyskać draft.

---

## 3. Użytkownik U2 strzela pod ten sam `documentId` w tym samym czasie

Tu jest jądro problemu. Trzy realistyczne strategie — wybór zależy od wymagań biznesowych.
**Rekomendowana: strategia C (optimistic + warning).**

### Strategia A — Pessimistic lock (klasyka, najprostsza)
- U2 dostaje `200 OK` z **wersją opublikowaną v0** w trybie **read-only**.
- Response zawiera:
  ```json
  {
    "documentId": "...",
    "versionId": "<v0>",
    "versionNumber": 0,
    "mode": "readonly",
    "lockedBy": { "userId": "U1", "displayName": "...", "since": "..." },
    "retryAfter": "<expiresAt sesji U1>"
  }
  ```
- HTTP `200` + nagłówek `X-Document-Lock: user=U1`.
- Frontend wyświetla banner: *„Dokument jest edytowany przez Jana Kowalskiego od 13:15. Możesz go przeglądać, ale nie edytować. Spróbuj ponownie za ~3 min lub poproś o przejęcie.”*
- Opcja `POST /api/documents/{id}/lock/takeover` (z uprawnieniem) — wymusza release.

**Kiedy:** dokumenty regulacyjne, prawne, gdzie merge jest niedopuszczalny.

### Strategia B — Współbieżna edycja (CRDT / OT)
- Wymaga WebSocketów, serwera OT (np. ShareDB) lub CRDT (Yjs / Automerge).
- U2 dołącza do tej samej sesji, dostaje **ten sam `versionId` draftu v1**, zmiany są merge'owane w locie.
- Zbyt duży narzut dla obecnej architektury Clean Architecture + REST → **odradzam jako pierwszy krok**.

### Strategia C — Optimistic concurrency + jawny warning (REKOMENDOWANA)
- U2 dostaje `200 OK` z **wersją opublikowaną v0**, w trybie `edit`, ale z ostrzeżeniem.
- Backend tworzy dla U2 **osobny draft** np. `VersionNumber = 1` ze swoim własnym `VersionId` (każdy edytor ma swoją gałąź).
- Response:
  ```json
  {
    "documentId": "...",
    "versionId": "<draft-U2>",
    "baseVersionId": "<v0>",
    "versionNumber": 1,
    "mode": "edit",
    "concurrentEditors": [
      { "userId": "U1", "sessionId": "...", "since": "...", "draftVersionId": "..." }
    ],
    "etag": "<hash v0>"
  }
  ```
- Frontend wyświetla nieblokujący banner: *„Jan Kowalski edytuje ten dokument równolegle. Po zapisie może być wymagane scalenie zmian.”*
- Przy zapisie obowiązuje **optimistic concurrency** — request `PUT` musi wysłać `If-Match: <etag>` zawierający hash bazowej wersji.
  - Pierwszy, który zapisze (np. U1) → publikuje v1, ETag zmienia się.
  - Drugi (U2) → dostaje `409 Conflict` z body `{ currentVersionId: <v1 U1>, yourDraftVersionId, diffUrl }`. UI pokazuje merge tool / 3-way diff albo „nadpisz / odrzuć / scal”.

---

## 4. Jaki GUID dostaje U2? — odpowiedź wprost

| Kontekst                                | `documentId` (logiczny) | `versionId` w odpowiedzi                                  |
|-----------------------------------------|-------------------------|------------------------------------------------------------|
| Upload przez system zewnętrzny          | nowy GUID dokumentu     | versionId v0 (zwykle niezwracany na zewnątrz)              |
| U1 otwiera w trybie view                | ten sam documentId      | versionId v0                                               |
| U1 otwiera w trybie edit                | ten sam documentId      | versionId draftu (v1 należący do sesji U1)                 |
| **U2 otwiera, gdy U1 trzyma lock (A)**  | **ten sam documentId**  | **versionId v0 (bazowa, opublikowana)** + flaga readonly   |
| **U2 otwiera równolegle (C)**           | **ten sam documentId**  | **versionId nowego, własnego draftu U2 z bazą v0**          |

**Złota zasada:** U2 **NIGDY nie dostaje versionId draftu U1**. Draft U1 to niespójny, nieopublikowany stan należący do prywatnej sesji edycyjnej U1. U2 widzi tylko ostatnią **opublikowaną** wersję (v0) — albo jako readonly snapshot, albo jako bazę pod własny draft.

Wersja U1 staje się widoczna dla U2 dopiero, gdy U1 wywoła `POST /api/documents/{id}/versions/{draftId}/publish` — wtedy v1 zostaje promowane na HEAD i dostaje nowy publiczny `etag`. Od tego momentu kolejne `GET` zwracają v1 jako bazę.

---

## 5. Co konkretnie wyświetlić użytkownikowi U2 (UX)

Zależnie od strategii:

**Strategia A (pessimistic):**
- Tryb tylko-do-odczytu z widocznym znacznikiem 🔒.
- Banner u góry: nazwa edytora, czas od kiedy edytuje, przewidywany czas odblokowania, przycisk *„Powiadom mnie, gdy będzie wolny”* (SignalR / WebSocket push), opcjonalnie *„Poproś o przejęcie”*.
- Wyłączony toolbar edycyjny.

**Strategia C (optimistic — rekomendowana):**
- Pełny edytor, ale z presence indicator (lista awatarów aktywnych edytorów + ich kursorów, jeśli zaimplementujemy SignalR).
- Banner informacyjny (nie blokujący): *„2 osoby edytują ten dokument. Twoja kopia bazuje na wersji 0. Przy zapisie sprawdzimy konflikty.”*
- Przy zapisie konfliktowym — modal 3-way diff (baza v0 / moje zmiany / wersja opublikowana przez U1) z opcjami: *Scal automatycznie / Wybierz ręcznie / Zapisz jako kopię (fork) / Anuluj*.

---

## 6. Komponenty do dodania w obecnej architekturze

W `D2ViewerEditor.Domain`:
- Encje `Document` (agregat root), `DocumentVersion`, `DocumentEditSession`.
- Value object `DocumentEtag` (hash zawartości).

W `D2ViewerEditor.Application` (CQRS / MediatR):
- `UploadDocumentCommand` → zwraca `documentId`.
- `OpenDocumentQuery` (with `mode`, `userId`) → zwraca DTO z `versionId`, `mode`, `concurrentEditors`.
- `SaveDocumentDraftCommand` (z `If-Match`).
- `PublishDocumentVersionCommand`.
- `HeartbeatEditSessionCommand`.

W `D2ViewerEditor.Infrastructure`:
- Repozytoria, blob storage (klucz `versionId`), tabela sesji z TTL (Redis lub kolumna `ExpiresAt` + background cleanup).

W `D2ViewerEditor.Api`:
- Endpointy REST + `ETag` / `If-Match` / `Location`.
- Opcjonalnie SignalR hub `/hubs/documents/{id}` do presence i powiadomień o publikacji.

Frontend (Angular 20):
- Serwis `DocumentSessionService` z signalami: `currentVersionId`, `mode`, `concurrentEditors`, `conflict`.
- Komponent `CollaborationBannerComponent`.
- Modal `ConflictResolutionDialog` (3-way diff).

---

## 7. Stany brzegowe, które trzeba przewidzieć

1. **U1 zamknął przeglądarkę bez save** → sesja wygasa po TTL (heartbeat ustaje), draft v1 przechodzi w stan `Abandoned` (zachowujemy do recovery 24 h, potem GC).
2. **U2 zaczął edycję, U1 publikuje v1** → U2 dostaje push SignalR `document.published` i banner *„Pojawiła się nowa wersja — pobierz / pracuj dalej z konfliktem przy zapisie”*.
3. **Sieć padła w trakcie save** → idempotency key na save chroni przed dublami.
4. **Aplikacja zewnętrzna wysłała ten sam plik dwa razy** → `Idempotency-Key` zwraca pierwotny `documentId` zamiast tworzyć nowy.
5. **Lock zombie** → kolumna `lastHeartbeatAt` + reaper job zwalniający locki po > 2× TTL.
6. **Przejęcie locka** → wymaga uprawnienia, zapisuje audyt, U1 przy próbie save dostaje 409 i propozycję zapisu jako fork.

---

## Podsumowanie odpowiedzi na zadane pytania

- **Co wyświetlić U2?** Pełny widok dokumentu z bazowej wersji v0, z jawnym banerem informującym, że U1 równolegle edytuje. W wariancie pessimistic — tryb read-only z informacją o blokadzie i ETA odblokowania.
- **Jaki GUID dostaje U2?** Logiczny `documentId` jest ten sam. `versionId` zwracany w body to **versionId bazowej, opublikowanej wersji v0** (nie draft U1!). Jeśli U2 też edytuje, dostaje dodatkowo własny `draftVersionId` zbudowany na bazie v0. Draft U1 (jego prywatne v1) jest niewidoczny dla U2 do momentu publikacji.
