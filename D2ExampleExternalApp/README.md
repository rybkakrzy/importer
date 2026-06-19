# D2ExampleExternalApp

Przykładowa aplikacja **zewnętrzna**, symulująca integratora D2 ViewerEditor. Pokazuje
end-to-end pełen obieg dokumentu i pozwala podejrzeć, co realnie wymieniamy po HTTP.

## Co robi

1. **Wysyłka do D2** — `POST /api/integration/upload` (multipart, pole `File`): aplikacja
   przesyła plik do `D2ServicesViewerEditor` (`POST /api/v1/document`) z `ReturnUrl`
   wskazującym z powrotem na własny endpoint callback.
2. **Edycja** — użytkownik edytuje dokument w edytorze i klika „Zakończ i wyślij".
3. **Odbiór zwrotki** — `POST /api/integration/callback` (returnUrl): D2 odsyła gotowy plik
   jako `multipart/form-data` z częściami `file`, `masterId`, `versionId`, `corporateKey`
   oraz nagłówkami `Idempotency-Key` i `X-Content-SHA256`. Plik jest zapisywany na dysk,
   liczony jest SHA-256 i porównywany z nagłówkiem.
4. **Podgląd** — `GET /api/integration/received` pokazuje, co wróciło; opcjonalnie
   `GET /api/integration/status/{masterId}` odpytuje status w D2.

## Uruchomienie

```bash
cd D2ExampleExternalApp
dotnet run
```

Swagger: <http://localhost:15120/swagger>

Wymaga uruchomionego `D2ServicesViewerEditor` (domyślnie `http://localhost:15112`).

## Konfiguracja (`appsettings.json`)

| Klucz | Znaczenie |
|---|---|
| `D2Services:BaseUrl` | adres D2Services (domyślnie `http://localhost:15112`) |
| `D2Services:ApiKey` | opcjonalny klucz — jeśli ustawiony, leci jako `X-Api-Key` |
| `ExternalApp:PublicBaseUrl` | publiczny URL TEJ aplikacji; z niego budujemy `returnUrl` |
| `ExternalApp:ReceivedFilesPath` | katalog na odebrane pliki (domyślnie `received/`) |

> **Ważne:** `returnUrl` musi być osiągalny z poziomu D2Services. Lokalnie wystarczy ten sam
> host (`localhost:15120`). Jeśli D2 działa w kontenerze, ustaw `PublicBaseUrl` na adres
> widoczny z kontenera (np. `http://host.docker.internal:15120`).

## Kontrakt zwrotki (to, czego oczekuje callback)

`POST {returnUrl}` — `multipart/form-data`:

| Pole | Typ | Opis |
|---|---|---|
| `file` | binary | gotowy DOCX (`filename=document.docx`) |
| `masterId` | text | GUID dokumentu master |
| `versionId` | text | GUID wersji edytowalnej |
| `corporateKey` | text | corporate key użytkownika kończącego edycję |

Nagłówki: `Idempotency-Key` (= deliveryId, do deduplikacji at-least-once),
`X-Content-SHA256` (hex SHA-256 zawartości). Endpoint **musi** zwrócić `2xx`, inaczej D2
ponawia wysyłkę.

> Uwaga: D2 wysyła pole użytkownika jako `corporateKey` (nie `userCk`).
