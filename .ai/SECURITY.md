# Security

## Zasada podstawowa

Bezpieczeństwo ma pierwszeństwo przed wygodą implementacji.

## Sekrety

- Sekrety trzymane są w `appsettings.{ENV}.secrets.json` (DEV/UAT/PRD) — **poza repo**. Nie commitować, nie drukować, nie kopiować do dokumentacji/testów.
- Connection stringi i klucze GCS to sekrety. Nie loguj ich.
- Jeśli sekret pojawi się w pliku versionowanym — zgłoś i nie powielaj wartości.

## Auth i autoryzacja

- Status: **do potwierdzenia w repo** — sprawdź `Program.cs`/middleware obu API przed zmianami auth. Nie zakładaj konkretnego mechanizmu.
- Backend musi egzekwować uprawnienia niezależnie od GUI.
- External API (`D2ServicesViewerEditor`) jest powierzchnią dla aplikacji zewnętrznych — traktuj wejście jako niezaufane.

## Walidacja (realne mechanizmy)

- Komendy mają walidatory FluentValidation (np. `IngestExternalDocumentCommandValidator`).
- Ingest waliduje: niepusty plik, typ MIME ∈ {DOCX, PDF}, rozmiar ≤ 100 MB, metadata = poprawny JSON ≤ 8000 znaków, classification ∈ {C1..C4}, ReturnUrl wymagany dla DOCX.
- Limity uploadu w kontrolerze: `RequestSizeLimit` ~110 MB, `MultipartBodyLengthLimit`.
- Nie ufaj `masterId`/`versionId` z requestu bez sprawdzenia istnienia (handlery zwracają `NotFound`).

## Pliki / treść

- Walidacja typu pliku po MIME + rozszerzeniu (`ResolveMimeType`).
- Uwaga na XSS w edytorze WYSIWYG — treść HTML pochodzi od użytkownika; przy renderowaniu używaj bezpiecznych mechanizmów Angulara, nie ręcznego `innerHTML` bez sanitizacji.

## Podpisy cyfrowe

- Format własny (Custom XML Part, RSA-SHA256) — nie standard OOXML. Certyfikaty/hasła przekazywane do `/api/document/sign` są wrażliwe: nie loguj, nie zapisuj.
- Hash liczony z `MainDocumentPart`; `VerifySignatures` raportuje osobno ważność podpisu RSA i zgodność hasha.

## Wysyłka „Zakończ i wyślij" (returnUrl)

- Realne mechanizmy: `RecipientUrl` (z `documents.metadata`, ustawiany przez aplikację zewnętrzną przy ingeście) jest walidowany jako **absolutny adres http(s)** (`DocumentDelivery.IsValidRecipientUrl`) — inaczej operacja `finish` zwraca 400.
- Wysyłka jest at-least-once; POST niesie nagłówek `Idempotency-Key = deliveryId` (dedup po stronie odbiorcy) i `X-Content-SHA256`.
- Klasyfikacja błędów: 4xx wybrane (400/401/403/404/405/422) → `FailedPermanently`; pozostałe/timeout → retry do 24 h.
- Wysyłany jest niezmienny snapshot (`deliveries/{deliveryId}`), nie bieżąca v2 — nie da się podmienić treści po zakończeniu.
- **Rekomendacja (nie istniejący stan):** `RecipientUrl` pochodzi z danych zewnętrznych, a worker wykonuje do niego żądania serwerowe — rozważyć ochronę przed SSRF (allowlista hostów/schematów, blokada adresów prywatnych/metadata-endpoint GCP `169.254.169.254`). Obecnie walidowany jest tylko format http(s).

## Logowanie

Nie loguj: haseł, tokenów, certyfikatów, danych osobowych, surowych payloadów z danymi wrażliwymi. Preferuj structured logging (`ILogger`). Brak `Console.WriteLine` w kodzie aplikacyjnym. (Uwaga: w GUI `document-editor` jest blok diagnostyczny `console.group` przy pobieraniu pliku — nie rozszerzaj go o dane wrażliwe.)

## Frontend

- `environment.*` (w tym `apiUrl`, `autoSave`) NIE są sekretami — trafiają do bundla.
- Nie przechowuj sekretów w kodzie frontu.

## Checklist security przy zmianie

- Czy backend waliduje i sprawdza uprawnienia/istnienie zasobu?
- Czy limity rozmiaru/MIME są zachowane?
- Czy błędy nie ujawniają stack trace / szczegółów infrastruktury?
- Czy logi nie zawierają sekretów?
- Czy treść użytkownika jest bezpiecznie renderowana (XSS)?
