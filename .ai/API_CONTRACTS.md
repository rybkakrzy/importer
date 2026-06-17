# API Contracts

## Cel pliku

Zasady projektowania/zmiany API oraz katalog realnych endpointów obu backendów.

## Styl API

REST over HTTP/JSON (+ multipart dla plików). Dwa backendy:
- **Internal API** `D2ApiViewerEditor` — base `http://localhost:5190/api` (DEV). Swagger UI: `http://localhost:5190/swagger`. Konsument: GUI.
- **External API** `D2ServicesViewerEditor` — base `http://localhost:15112/api/v1` (DEV; `Urls` w `appsettings.json`). Swagger UI: `http://localhost:15112/swagger` (Swagger włączony w dev/local). Konsument: aplikacje źródłowe.

## Zasada kompatybilności

Nie zmieniaj istniejącego kontraktu w sposób breaking bez świadomej decyzji. Breaking = zmiana nazwy/typu/znaczenia pola, usunięcie pola, zmiana statusów HTTP, formatu błędu, wymagań walidacyjnych, paginacji.

## Format błędu

Obecny styl projektu: `{ "error": "komunikat" }` (sprawdź `BaseApiController` przy zmianach). Trzymaj się go dla spójności.

## Wersjonowanie

- Internal API: bez wersji w ścieżce (`/api/...`).
- External API: wersjonowanie w URL (`/api/v1/...`).

---

# Katalog endpointów — Internal API (`/api/documentstorage`, `/api/document`, `/api/barcode`)

## DocumentStorageController — `/api/documentstorage` (CRUD: DB + GCS)

| Metoda | Ścieżka | Opis | Odpowiedź |
|---|---|---|---|
| GET | `/?skip=0&take=200` | Lista dokumentów (admin) | `DocumentListItemDto[]` |
| POST | `/upload` | Upload nowego dokumentu (pierwszy zapis) | `UploadDocumentResult { masterId, versionId }` |
| POST | `/{masterId}/save` | Nowa wersja (`AddVersion`) | `SaveDocumentVersionResult { versionId }` |
| **PUT** | `/{masterId}/versions/{versionId}` | **Nadpisanie wersji w miejscu** (auto-save / „Zapisz"). v1 → 400 | `UpdateDocumentVersionResult { versionId, versionNumber, sizeInBytes, modifiedAt }` |
| GET | `/{masterId}` | Dokument + aktywna wersja (z contentem) | `DocumentDto` |
| **GET** | `/{masterId}/metadata` | Metadane (returnUrl, classification) | `{ masterId, mimeType, returnUrl?, classification? }` |
| GET | `/{masterId}/versions` | Historia wersji (bez contentu) | `DocumentVersionDto[]` |
| GET | `/{masterId}/download` | Bajty wersji bazowej (v1, sort po VersionNumber) | plik |
| GET | `/{masterId}/versions/{versionId}/download` | Bajty konkretnej wersji | plik |
| POST | `/{masterId}/restore/{versionId}` | Przywróć wersję jako aktywną | `{ message, versionId }` |
| **POST** | `/{masterId}/versions/{versionId}/finish` | **„Zakończ i wyślij"**: utrwala stan edytora, zamraża snapshot, tworzy zadanie wysyłki. Idempotentne (zwraca istniejące aktywne zadanie). Brak/zły returnUrl → 400 | **202 Accepted** `{ deliveryId, status, statusUrl }` |
| **POST** | `/{masterId}/user-download` | **„Pobierz dokument"** — konwertuje aktualny stan edytora (HTML+header/footer/margins) na DOCX dla użytkownika. Egzekwuje regułę domenową: tylko gdy `documents.metadata.userDownload == true`. Body: `SaveDocumentRequest`. → 200 plik DOCX / 403 (gate) / 404 / 400 |
| GET | `/deliveries/{deliveryId}` | Status zadania wysyłki (polling z GUI) | `DeliveryStatusDto { deliveryId, documentId, status, attemptCount, lastAttemptAt?, nextAttemptAt?, lastError?, updatedAt }` |
| GET | `/deliveries?status=&skip=&take=` | Lista zadań w danym statusie (monitoring/admin; domyślnie `DeadLettered`; widok GUI `/admin/deliveries`) | `DeliveryListItemDto[] { deliveryId, documentId, status, attemptCount, createdAt, lastAttemptAt?, nextAttemptAt?, deadlineAt, lastError?, lockedUntil?, lockedBy }` |
| POST | `/deliveries/{deliveryId}/retry` | Ręczne wznowienie zadania (`DeadLettered`/`FailedPermanently`/`RetryScheduled`/`Cancelled` → kolejka, wysyłka teraz) | `RequeueDeliveryResult { deliveryId, status }` |
| POST | `/deliveries/{deliveryId}/cancel` | Ręczne anulowanie zadania oczekującego/zaplanowanego (`Pending`/`RetryScheduled` → `Cancelled`) | `CancelDeliveryResult { deliveryId, status }` |
| PUT | `/deliveries/{deliveryId}/recipient-url` | Zmiana adresu odbiorcy (returnUrl) zadania; body `{ recipientUrl }`; dozwolone dla zadań ≠ `Sent`/`Sending` | `UpdateDeliveryRecipientUrlResult { deliveryId, recipientUrl, status }` |

Request DTO zapisu: `{ content: byte[]/base64, createdBy?: string }` (ten sam DTO `SaveDocumentVersionRequest` używany też przez `finish`).

**Uwierzytelnianie (Entra ID):** Internal API (via **Microsoft.Identity.Web**) wymaga **access tokena** (`Authorization: Bearer …`) dla wszystkich endpointów poza `GET /api/health`. Brak/nieważny token → **401**. Token dołącza front przez `MsalInterceptor`. Role: App Roles + **mapowanie grup→role** (`groups`→`APP_*`). Admin: `GET /api/identity/users?query=` (Graph user lookup, `RequireAppAdmin`). Patrz `SECURITY.md`, `DECISIONS.md` ADR-0011 + **ADR-0012**.

**Autoryzacja (role):** endpointy administracyjne — `GET /api/documentstorage`, `GET /api/documentstorage/deliveries`, `POST /api/documentstorage/deliveries/{id}/retry` — wymagają roli `Administrator` (policy `RequireAppAdmin`); inaczej **403**.

**Kontrola dostępu do dokumentu:** endpointy treści/metadanych (`GET /{masterId}`, `/{masterId}/metadata`, `/{masterId}/download`, `/{masterId}/versions/{versionId}/download`) sprawdzają `allowedCorporateKeys` względem `CorporateKey` z claimu access tokena (`AzureAd:CorporateKeyClaim`). Brak uprawnień → **403** `{ error }`. Dokument bez `allowedCorporateKeys` = każdy zalogowany. **`Administrator` omija listę.** GUI: `documentAccessGuard` → `/access-denied`.

## DocumentController — `/api/document` (operacje bezstanowe, bez DB/GCS)

| Metoda | Ścieżka | Opis |
|---|---|---|
| POST | `/open` | DOCX (multipart) → `DocumentContent` (HTML) |
| POST | `/save` | HTML → DOCX (download) |
| GET | `/new` | Pusty `DocumentContent` |
| POST | `/upload-image` | Obraz → Base64 (`ImageUploadResponse`) |
| GET | `/templates` / `/templates/{id}` | Szablony |
| POST | `/sign` | HTML → DOCX → podpisany DOCX |
| POST | `/verify-signatures` | DOCX → `DigitalSignatureInfo[]` |
| POST | `/export-pdf` | **501 Not Implemented** (placeholder) |

## BarcodeController — `/api/barcode`

| Metoda | Ścieżka | Opis |
|---|---|---|
| POST | `/generate` | `{ content, barcodeType, width, height, showText }` → `{ imageBase64, mimeType }` |
| GET | `/types` | `string[]` typów (QR_CODE, CODE_128, EAN_13, ...) |

## HealthController — `/api/health`

Standardowe health checki.

---

# Katalog endpointów — External API (`/api/v1/document`)

## DocumentController (`D2ServicesViewerEditor`)

| Metoda | Ścieżka | Opis |
|---|---|---|
| POST | `/api/v1/document` | Ingest DOCX/PDF (multipart). Pola: `File`, `ReturnUrl` (wymagany dla DOCX), `Classification` (C1..C4, obligatoryjna), **opcjonalnie `UserDownload: bool?`** (domyślnie `false`; tylko jawne `true` zezwala użytkownikowi na pobranie edytowanego pliku — patrz reguła `userDownload` poniżej). Nagłówek opcjonalny `X-Created-By`. → 201 `CreateDocumentResponse { masterId, versionId? }` (versionId tylko DOCX) |
| GET | `/api/v1/document/{documentId}` | Placeholder (read flow niezaimplementowany) |
| **PUT** | `/api/v1/document/{masterId}/callback-url` | **Aktualizacja URL do wysyłki po „Zakończ"**. Body: `{ "url": "https://..." }`. Walidacja przez `DocumentDelivery.IsValidRecipientUrl` (absolutny http/https, ≤ 2048 znaków). Zapisuje w `documents.metadata.returnUrl` (zachowuje `classification`). Idempotentny. Blokowany w stanach `Sending`/`Sent`/`DeliveryFailed`. → 204 / 400 / 404 / 409 |
| **POST** | `/api/v1/document/{masterId}/unlock` | **Odblokowanie dokumentu**. Body opcjonalne: `{ "reason": "..." }` (logowane). W tej domenie `DocumentStatus.Editing` = „trzymany przez edytora", więc unlock = `Editing → Saved` (dodano `Document.MarkSaved()`). Idempotentny: `Saved` → 200 z `Changed=false`. Blokowany dla stanów wysyłki (409). → 200 `UnlockDocumentResult { masterId, changed }` / 404 / 409 |
| **GET** | `/api/v1/document/{masterId}/status` | **Status dokumentu**. → 200 `DocumentStatusDto { masterId, status, isLocked, hasCallbackUrl, activeVersionId?, activeVersionNumber?, activeVersionModifiedAt?, latestDelivery? { deliveryId, status, attemptCount, lastAttemptAt?, nextAttemptAt?, deadlineAt } }`. Pełny `callbackUrl` celowo NIE jest zwracany (może zawierać token — surfacowany tylko jako boolean `hasCallbackUrl`). / 404 |

### Decyzja `masterGuid` vs `versionGuid` dla trzech nowych endpointów

Wszystkie trzy używają **wyłącznie `masterGuid`** — uzasadnienie:

| Endpoint | Identyfikator | Dlaczego nie `versionGuid` |
|---|---|---|
| `PUT .../callback-url` | `masterGuid` | `returnUrl` żyje w `documents.metadata` (master), wspólny dla wszystkich wersji. Po „Zakończ" worker pracuje na **zamrożonej** `RecipientUrl` w `DocumentDelivery` (snapshot) — edycja metadanych później nie psuje już-zakolejkowanej wysyłki, ale też nie ma sensu wymuszać konkretnej wersji. |
| `POST .../unlock` | `masterGuid` | Brak osobnego user-locka w domenie. „Lock" = `Document.Status == Editing` (frontend pokazuje to jako `lockedByOther`). Status na poziomie master → `versionGuid` nic nie wnosi. Worker-lease `LockedUntil`/`LockedBy` na `DocumentDelivery` to inny mechanizm i nie powinien być odblokowywany z zewnątrz. |
| `GET .../status` | `masterGuid` | Status jest atrybutem master; wersje nie mają własnego statusu. Odpowiedź niesie `activeVersionId` + `latestDelivery` dla pełnego obrazu cyklu życia. |

Metadane trafiają do `documents.metadata` jako `{ "returnUrl": "...", "classification": "C2", "userDownload": true | null }`.

### Reguła domenowa: `userDownload` (kontrola pobierania pliku)

`userDownload` jest opcjonalnym polem w metadanych dokumentu kontrolującym, czy użytkownik może pobrać edytowany plik na komputer.

| Stan w metadanych | Interpretacja | Pozycja menu „Pobierz dokument" |
|---|---|---|
| brak pola | `false` | ukryta |
| `null` | `false` | ukryta |
| `false` | `false` | ukryta |
| `"true"` (string) / liczba / inne | `false` (bool deserializacja restrykcyjna) | ukryta |
| `true` | `true` | widoczna |

**Lokalny upload** (`POST /api/documentstorage/upload`) — backend (`UploadDocumentCommandHandler`) **automatycznie** zapisuje `{"userDownload":true}` w metadanych. Klient nie może tego ustawić ani nadpisać — system jest jedynym źródłem prawdy dla tego flow (rule 12 — anti-tamper).

**External ingest** (`POST /api/v1/document`) — aplikacja źródłowa przekazuje `UserDownload: bool?` w form-data. Domyślnie `false`. Tylko jawne `true` aktywuje pobieranie.

**Egzekwowanie** — `POST /api/documentstorage/{masterId}/user-download` jest jedyną ścieżką dla użytkownika do pobrania edytowanego pliku z aktualnego stanu edytora. Endpoint sprawdza flagę przez `ExternalDocumentMetadata.IsUserDownloadAllowed` i zwraca **403** gdy nie spełniona. Frontend dodatkowo ukrywa pozycję menu, ale nie jest źródłem zabezpieczenia.

**Niezależność od `returnUrl`** — `userDownload` i `returnUrl` to dwa odrębne mechanizmy:
- `returnUrl` → automatyczny zwrot do systemu źródłowego po „Zakończ"
- `userDownload` → ręczne pobranie edytowanego pliku przez użytkownika

Dokument może mieć: tylko `returnUrl` (zewnętrzny, bez pobrania), tylko `userDownload=true` (lokalny upload), albo oba.

## Checklist przed zmianą API

- Czy frontend / aplikacja zewnętrzna są zaktualizowane?
- Czy testy przechodzą? Czy stary klient się nie psuje?
- Czy statusy HTTP i format błędu są spójne ze stylem?
- Czy zmiana jest w `CHANGELOG.md`?
