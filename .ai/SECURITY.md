# Security

## Zasada podstawowa

Bezpieczeństwo ma pierwszeństwo przed wygodą implementacji.

## Sekrety

- Sekrety trzymane są w `appsettings.{ENV}.secrets.json` (DEV/UAT/PRD) — **poza repo**. Nie commitować, nie drukować, nie kopiować do dokumentacji/testów.
- Connection stringi i klucze GCS to sekrety. Nie loguj ich.
- Jeśli sekret pojawi się w pliku versionowanym — zgłoś i nie powielaj wartości.

## Auth i autoryzacja (Microsoft Entra ID)

- **Uwierzytelnianie:** Entra ID przez **`Microsoft.Identity.Web`** (`AddMicrosoftIdentityWebApi`, sekcja `AzureAd` — `ClientId`/`TenantId`/`Instance`/`Audience`). **Cała aplikacja za logowaniem** — `[Authorize]` na `BaseApiController`; Health anonimowy. Front: MSAL (redirect), token dołączany `MsalInterceptor`. Patrz ADR-0012.
- **Role:** dwa źródła współistnieją (ADR-0012): **App Roles** (claim `roles`, `Operator`/`Administrator`) **oraz mapowanie grup→role** (`ClaimsTransformer`: claim `groups` + sekcja `Roles` `GroupNames`→`RoleName`). `RoleClaimType="roles"`. Policy: `RequireAppOperator` (Operator lub Administrator) egzekwowana **klasowo na `DocumentController` i `DocumentStorageController`** (cały edytor + cykl życia dokumentu) oraz na `GET /api/identity/resources`; `RequireAppAdmin` (tylko Administrator) na module admina (`GET /api/documentstorage`, `deliveries`, `deliveries/{id}/retry`, `GET /api/identity/users`) — atrybut na metodzie łączy się AND z polityką klasy, więc te endpointy wymagają Administratora. **Samo `[Authorize]` (dowolny zalogowany) NIE wystarcza do edytora** — patrz ADR-0022. Egzekwowanie backendowe zweryfikowane testami `D2ViewerEditor.Api.IntegrationTests` (401 bez tokenu, 403 bez roli). Patrz ADR-0011 + ADR-0012 + ADR-0022.
- **Microsoft Graph (app-only):** `GraphUserService` (`ClientSecretCredential`, scope `.default`) — lookup userów dla admina; aktywny tylko z `AzureAd:ClientSecret` (z GCP Secret Manager), inaczej no-op. **GCP Secret Manager:** `EntraSecretLoader` wstrzykuje `ClientSecret` z `SMC01{ENV}2_APP00404_entra_secret`. Sekrety nigdy w repo (placeholdery).
- **Backend = źródło prawdy.** Angular guardy (`authGuard` → MsalGuard/bypass, `resourceGuard` po liście z `GET /api/identity/resources`, `homeRedirectGuard`, `documentAccessGuard`) to tylko UX — ukrycie/blokada wejścia, nie zabezpieczenie. (Dawny `appAdminGuard` usunięty — front nie zna nazw ról, ADR-0016.)
- External API (`D2ServicesViewerEditor`) — app-to-app (nie per user); rejestruje `SystemCurrentUserProvider`. Traktuj wejście jako niezaufane.

### Kontrola dostępu do dokumentu (`allowedCorporateKeys`)

- Reguła (`DocumentAccessPolicy`, Domain): metadane mogą zawierać `allowedCorporateKeys`. Brak/null/pusta → dokument dla każdego **zalogowanego** (authenticated-public); niepusta → tylko użytkownik z pasującym `CorporateKey`. Porównanie: `Trim()` + Ordinal **ignore-case**.
- **`CorporateKey` z access tokena** — `ClaimsCurrentUserProvider` czyta konfigurowalny claim (`AzureAd:CorporateKeyClaim`, domyślnie **`corpKey`**) **case-insensitive po nazwie claimu** (token Qutalo-AD niesie `corpkey` małymi literami, a `CaseSensitiveClaimsIdentity` z Identity.Web nie dopasowałby `FindFirst`); brak/pusty → null → dokument ograniczony = **403**. corpKey NIE jest transportowany z GUI — jedynym źródłem jest token (ADR-0021). (Dev: `HttpHeaderCurrentUserProvider` z nagłówka `X-Corporate-Key`, fallback `DEV-LOCAL` — nie rejestrowany w produkcji.)
- **`Administrator` omija `allowedCorporateKeys`** — pełny wgląd (`DocumentAccessGuard` zwraca true dla admina). Decyzja biznesowa ADR-0011.
- Egzekwowanie: `IDocumentAccessGuard` w handlerach treści/metadanych → `Result.Forbidden` → **403**; brak/nieważny token → **401** (JwtBearer). GCS tylko przez backend.
- Frontend: `documentAccessGuard` (po `authGuard` + `resourceGuard`) — 403 → `/brak-uprawnien` (`AccessForbiddenComponent`; stara trasa `/access-denied` przekierowuje).
- Nie loguj `allowedCorporateKeys` ani `CorporateKey` (dane wrażliwe) — polityka jest czysta i nie loguje.

## Walidacja (realne mechanizmy)

- Komendy mają walidatory FluentValidation (np. `FinishAndSendDocumentCommandValidator`); ingest waliduje w kontrolerze + handlerze (osobny walidator ingestu został usunięty).
- Ingest waliduje: niepusty plik, typ MIME ∈ {DOCX, PDF}, rozmiar ≤ 100 MB, metadata = poprawny JSON, `Classification` **opcjonalna** (jeśli podana: C1..C4, inaczej 400), `ReturnUrl` wymagany dla DOCX.
- **Centralne polityki uploadu (ADR-0024, 2026-06-22):** `IFileUploadSecurityService` (`FileUploadSecurityService`, sekcja `Security:Upload`) — spójność extension↔MIME↔magic-bytes, inspekcja archiwum DOCX (zip-slip, limity entry/uncompressed/compression-ratio, wymagane part-y OOXML, blokada makr VBA), sanitizacja MIME obrazów; abstrakcja skanera AV `IFileScanner` (domyślnie `NoOpFileScanner` — patrz R-28) z polityką fail-closed. Egzekwowane w `UploadDocument`/`UploadImage`/`IngestExternalDocument`.
- Limity uploadu w kontrolerze: `RequestSizeLimit` ~110 MB, `MultipartBodyLengthLimit`; Kestrel: globalny limit body 150 MB.
- Nie ufaj `masterId`/`versionId` z requestu bez sprawdzenia istnienia (handlery zwracają `NotFound`).

## Pliki / treść

- Walidacja typu pliku po MIME + rozszerzeniu (`ResolveMimeType`).
- Uwaga na XSS w edytorze WYSIWYG — treść HTML pochodzi od użytkownika; przy renderowaniu używaj bezpiecznych mechanizmów Angulara, nie ręcznego `innerHTML` bez sanitizacji.

## Podpisy cyfrowe

- Format własny (Custom XML Part, RSA-SHA256) — nie standard OOXML. Certyfikaty/hasła przekazywane do `/api/document/sign` są wrażliwe: nie loguj, nie zapisuj.
- Hash liczony z `MainDocumentPart`; `VerifySignatures` raportuje osobno ważność podpisu RSA i zgodność hasha.

## Wysyłka „Zakończ i wyślij" (returnUrl)

- Realne mechanizmy: `RecipientUrl` (z `documents.metadata`, ustawiany przez aplikację zewnętrzną przy ingeście) jest walidowany jako **absolutny adres http(s)** (`DocumentDelivery.IsValidRecipientUrl`) — inaczej operacja `finish` zwraca 400.
- **Ochrona SSRF (ADR-0024; `IReturnUrlValidator`/`ReturnUrlValidator`, sekcja `Security:ReturnUrl`):** normalizacja URL + kontrola schematu/znaków, blokada protocol-relative i user-info, blokada loopback/adresów prywatnych, opcjonalna allowlista hostów (`AllowedHosts`). Egzekwowane w `IngestExternalDocument`/`UpdateCallbackUrl`/`UpdateDeliveryRecipientUrl`/`FinishAndSendDocument`. **Pozostaje otwarte (R-07 Partial):** weryfikacja DNS (rebind) i obowiązkowa allowlista na prod/UAT.
- Wysyłka jest at-least-once; żądanie to POST **`multipart/form-data`** z częściami `file` (DOCX, filename `document.docx`), `masterId`, `versionId`, `corporateKey` oraz nagłówkami `Idempotency-Key = deliveryId` (dedup po stronie odbiorcy) i `X-Content-SHA256`. Odbiorca musi czytać plik z pola formularza `file`, nie z surowego body.
- Klasyfikacja błędów: 4xx wybrane (400/401/403/404/405/422) → `FailedPermanently`; pozostałe/timeout → retry do 24 h.
- Wysyłany jest niezmienny snapshot (`deliveries/{deliveryId}`), nie bieżąca v2 — nie da się podmienić treści po zakończeniu.

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
