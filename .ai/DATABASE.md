# Database

## Cel pliku

Zasady pracy z bazą danych, migracjami i zapytaniami.

## Baza danych

| Obszar | Wartość |
|---|---|
| Engine | PostgreSQL |
| ORM | EF Core 8 + Npgsql (`Npgsql.EntityFrameworkCore.PostgreSQL` 8.0.11) |
| Migracje | **Ręczny SQL** w `infra/sql/` — NIE EF Migrations (brak folderu `Migrations/`) |
| DbContext | `DocumentDbContext` (`Infrastructure/Persistence`) z `DbSet<Document>`, `DbSet<DocumentVersion>` |
| Konfiguracje EF | `DocumentConfiguration`, `DocumentVersionConfiguration` |
| Repo | `DocumentRepository : IDocumentRepository` (odczyty `AsNoTracking()`, filtr soft-delete) |
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

## Tabele (model EF)

**`documents`**: `id` (PK, guid_master), `name`, `mime_type`, `created_at`, `created_by`, `is_deleted` (default false), `metadata` (text, JSON). Indeksy: `created_at`, `is_deleted`.

**`document_versions`**: `id` (PK, guid_wersji), `document_id` (FK → documents, cascade), `storage_path` (≤500), `size_in_bytes`, `version_number`, `created_at`, `created_by`, `is_active`, `modified_at` (nullable). Indeksy: `document_id`, `(document_id, is_active)`, `created_at`.

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
