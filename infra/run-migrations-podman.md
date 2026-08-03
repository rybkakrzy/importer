# Uruchamianie migracji DB na Podman (PowerShell)

Migracje to surowe pliki SQL w `infra/sql/` (`001_*.sql` … `010_*.sql`). Nie ma EF Migrations.
Wszystkie pliki są **idempotentne** (`ADD COLUMN IF NOT EXISTS` / `DROP CONSTRAINT IF EXISTS` / `COMMENT`),
więc można je bezpiecznie puścić ponownie.

| Parametr | Wartość |
|---|---|
| Kontener | `d2viewereditor_postgres` |
| Baza | `d2viewereditor` |
| User | `postgres` |
| Hasło | `postgres` |
| Port | `5432` |

> W PowerShell **nie** działa `< plik.sql` (jak w bashu) — używamy `Get-Content -Raw … | psql -f -`.
> `ON_ERROR_STOP=1` sprawia, że błąd SQL kończy proces niezerowym kodem (inaczej `psql` zignorowałby błędy w środku skryptu).

## Pojedyncza migracja (np. `010_extend_document_status.sql`)

```powershell
# --- Migracja DB: 010_extend_document_status.sql (Podman) ---
$container = "d2viewereditor_postgres"
$db        = "d2viewereditor"
$dbUser    = "postgres"
$sqlDir    = "sciezka_do_projektu"
$migration = "010_extend_document_status.sql"

# 1. Sprawdź, czy kontener Postgres działa
$running = podman ps --filter "name=$container" --format "{{.Names}}"
if ($running -ne $container) {
    Write-Error "Kontener '$container' nie jest uruchomiony. Wystartuj go: podman start $container"
    return
}

# 2. Zaczekaj aż baza przyjmuje połączenia
podman exec $container pg_isready -U $dbUser -d $db

# 3. Zastosuj migrację (stdin → psql -f -), zatrzymaj się na pierwszym błędzie
Get-Content -Raw (Join-Path $sqlDir $migration) |
    podman exec -i $container psql -U $dbUser -d $db -v ON_ERROR_STOP=1 -f -

if ($LASTEXITCODE -eq 0) { Write-Host "OK: zastosowano $migration" -ForegroundColor Green }
else { Write-Error "Migracja $migration nie powiodla sie (exit $LASTEXITCODE)" }

# 4. Weryfikacja — komentarz kolumny powinien wymieniać Queued i SendAborted
podman exec $container psql -U $dbUser -d $db -c "SELECT col_description('documents'::regclass, (SELECT attnum FROM pg_attribute WHERE attrelid='documents'::regclass AND attname='status')) AS status_comment;"
```

## Wszystkie migracje od nowa (idempotentne)

```powershell
$container = "d2viewereditor_postgres"
$db        = "d2viewereditor"
$dbUser    = "postgres"
$sqlDir    = "sciezka_do_projektu"

Get-ChildItem $sqlDir -Filter "*.sql" | Sort-Object Name | ForEach-Object {
    Write-Host "==> $($_.Name)" -ForegroundColor Cyan
    Get-Content -Raw $_.FullName |
        podman exec -i $container psql -U $dbUser -d $db -v ON_ERROR_STOP=1 -f -
    if ($LASTEXITCODE -ne 0) { Write-Error "Blad w $($_.Name)"; break }
}
```

## Uwagi

- `010_extend_document_status.sql` nie zmienia schematu — kolumna `documents.status` to `VARCHAR(40)` bez `CHECK`,
  więc aplikacja zapisze nowe statusy `Queued` / `SendAborted` bez żadnej zmiany w bazie (to tylko aktualizacja komentarza).
- Statusy delivery (`Cancelled`) były już dodane w `008_add_delivery_cancelled_status.sql`.
- Jak wystartować/zatrzymać kontener i pełna konfiguracja: patrz [`postgres-setup.md`](./postgres-setup.md).
```
