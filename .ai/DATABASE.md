# Database

## Cel pliku

Zasady pracy z bazą danych, migracjami i zapytaniami.

## Baza danych

| Obszar | Wartość |
|---|---|
| Engine | PostgreSQL |
| ORM | EF Core 8 + Npgsql (`Npgsql.EntityFrameworkCore.PostgreSQL` 8.0.11) |
| Migracje | **Ręczny SQL** w `infra/sql/` — NIE EF Migrations (brak folderu `Migrations/`) |
| DbContext | `DocumentDbContext` (`Infrastructure/Persistence`) z `DbSet<Document>`, `DbSet<DocumentVersion>`, `DbSet<DocumentDelivery>` |
| Konfiguracje EF | `DocumentConfiguration`, `DocumentVersionConfiguration`, `DocumentDeliveryConfiguration` (enumy mapowane `HasConversion<string>()`) |
| Repo | `DocumentRepository : IDocumentRepository` (odczyty `AsNoTracking()`, filtr soft-delete), `DocumentDeliveryRepository : IDocumentDeliveryRepository` (claim `FOR UPDATE SKIP LOCKED`) |
| Storage plików | binaria NIE w bazie — w GCS (bucket `d2viewereditor-documents`); w bazie tylko `storage_path` |
| Connection string | `appsettings.{ENV}.json` → `ConnectionStrings` (wartości wrażliwe w `*.secrets.json`, poza repo) |

## Skrypty SQL (`infra/sql/`)

| Plik | Co robi |
|---|---|
| `001_init_schema.sql` | Schemat początkowy (documents, document_versions) |
| `002_init_indexes.sql` | Indeksy |
| `003_migrate_storage_to_gcs.sql` | Przejście storage na GCS |
| `004_add_document_metadata.sql` | `documents.metadata TEXT` (JSON od aplikacji zewnętrznej) |
| `005_add_version_modified_at.sql` | `document_versions.modified_at TIMESTAMPTZ` (znacznik nadpisania v2) |
| `006_add_document_status.sql` | `documents.status VARCHAR(40) NOT NULL DEFAULT 'Saved'` (cykl życia dokumentu) |
| `007_add_document_deliveries.sql` | tabela `document_deliveries` (kolejka „Zakończ i wyślij") + indeksy + check status |
| `008_add_delivery_cancelled_status.sql` | rozszerzenie CHECK `document_deliveries.status` o `Cancelled` |
| `009_add_delivery_corporate_key.sql` | `document_deliveries.corporate_key` (kto kończył/wysyłał) |
| `010_extend_document_status.sql` | nowe wartości `documents.status`: `Queued`, `SendAborted` (tylko COMMENT — kolumna bez CHECK) |
| `011_add_document_last_modified_by.sql` | `documents.last_modified_by` (corpKey ostatnio modyfikującego) |

## Tabele (model EF)

**`documents`**: `id` (PK, guid_master), `name`, `mime_type`, `created_at`, `created_by`, `is_deleted` (default false), `metadata` (text, JSON), `status` (varchar(40), default `Saved`; wartości enum `DocumentStatus`: `Saved`/`Editing`/`Queued`/`Sending`/`SendAborted`/`DeliveryFailed`/`Sent` — `Queued`/`SendAborted` od SQL 010), `last_modified_by` (nullable, corpKey — SQL 011). Indeksy: `created_at`, `is_deleted`.

**`document_versions`**: `id` (PK, guid_wersji), `document_id` (FK → documents, cascade), `storage_path` (≤500), `size_in_bytes`, `version_number`, `created_at`, `created_by`, `is_active`, `modified_at` (nullable). Indeksy: `document_id`, `(document_id, is_active)`, `created_at`.

**`document_deliveries`** (kolejka wysyłki finalnego pliku na returnUrl): `id` (PK), `document_id` (FK → documents, cascade), `source_version_id`, `snapshot_object_name` (niezmienny obiekt GCS `deliveries/{id}`), `snapshot_size_bytes`, `snapshot_sha256`, `recipient_url`, `status` (varchar(32), default `Pending`; enum `DeliveryStatus`: `Pending`/`Sending`/`RetryScheduled`/`Sent`/`FailedPermanently`/`DeadLettered`/`Cancelled` — `Cancelled` od SQL 008), `attempt_count`, `created_at`, `updated_at`, `first_attempt_at`, `last_attempt_at`, `next_attempt_at`, `deadline_at` (created_at + 24 h), `locked_until` + `locked_by` (lease techniczny claimu — NIE status biznesowy), `last_error`, `correlation_id`, `created_by`, `corporate_key` (SQL 009).
- `CHECK ck_document_deliveries_status` ogranicza dozwolone wartości `status`.
- Indeksy: `ix_..._due` (partial: `next_attempt_at WHERE status IN ('Pending','RetryScheduled')`), `ix_..._stuck` (partial: `locked_until WHERE status='Sending'`), `ux_..._active_per_document` (**unique partial**: `document_id WHERE status IN ('Pending','Sending','RetryScheduled')` — jedno aktywne zadanie na dokument = idempotencja kliknięcia), `ix_..._status`.
- Pobieranie zadań przez workera: `SELECT ... FOR UPDATE SKIP LOCKED` w `DocumentDeliveryRepository.ClaimDueBatchAsync` (surowy SQL `FromSqlRaw`, bezpieczny dla wielu instancji).

## Zasady migracji

- **Nie generuj migracji EF.** Dodaj nowy ponumerowany skrypt w `infra/sql/` ORAZ mapowanie w odpowiednim `*Configuration`.
- Kolumny dodawaj `ADD COLUMN IF NOT EXISTS`; dokumentuj `COMMENT ON COLUMN`.
- Zmiany destrukcyjne wymagają osobnej decyzji (`DECISIONS.md`) i analizy danych.
- Migracja nie miesza się z refaktoryzacją domeny.

## Zasady zapytań

- Odczyty bez modyfikacji: `AsNoTracking()`.
- Unikaj N+1; dla list rozważ projekcję do DTO i paginację (`GetDocumentsQuery(skip, take)`).
- Pobieranie dokumentu z wersjami: `GetByIdWithVersionsAsync`.

## Dane testowe / lokalny setup

- Do potwierdzenia: brak `docker-compose.yml` w repo, więc lokalna baza + GCS uruchamiane są poza repo (Postgres + fake-gcs-server). Sposób bootstrapu schematu = uruchomienie skryptów `infra/sql/` w kolejności.

## Checklist przed zmianą modelu danych

- Czy zmiana jest wymagana przez funkcję?
- Czy dodano skrypt SQL w `infra/sql/` + mapowanie w `*Configuration`?
- Czy migracja jest bezpieczna dla istniejących danych?
- Czy trzeba indeks? Czy uwzględniono rollback?
- Czy backend i frontend rozumieją nowy model?
