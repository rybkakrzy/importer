# Database Schema - D2ViewerEditor

## Przegląd

Backend używa PostgreSQL do przechowywania metadanych dokumentów i historii wersji.
Binarna zawartość dokumentów jest przechowywana w object storage (GCS lub fake-gcs w DEV),
natomiast dane relacyjne są przechowywane w tabelach SQL.

## Tabele

### `documents`

Główny aggregate root dla logicznego dokumentu.

```sql
CREATE TABLE IF NOT EXISTS documents (
	id UUID PRIMARY KEY,
	name VARCHAR(500) NOT NULL,
	mime_type VARCHAR(255) NOT NULL,
	created_at TIMESTAMP WITHOUT TIME ZONE NOT NULL,
	created_by VARCHAR(255) NOT NULL,
	is_deleted BOOLEAN NOT NULL DEFAULT FALSE,

	CONSTRAINT chk_name_not_empty CHECK (name <> ''),
	CONSTRAINT chk_mime_type_not_empty CHECK (mime_type <> '')
);
```

### `document_versions`

Każda operacja zapisu tworzy nowy wiersz w tej tabeli.

```sql
CREATE TABLE IF NOT EXISTS document_versions (
	id UUID PRIMARY KEY,
	document_id UUID NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
	version_number INT NOT NULL,

	-- legacy binary payload (still present in current migration set)
	content BYTEA NOT NULL,

	-- object storage fields (added in migration 003)
	storage_path VARCHAR(500) NOT NULL,
	size_in_bytes BIGINT NOT NULL,

	created_at TIMESTAMP WITHOUT TIME ZONE NOT NULL,
	created_by VARCHAR(255) NOT NULL,
	is_active BOOLEAN NOT NULL DEFAULT FALSE,

	CONSTRAINT chk_version_number_positive CHECK (version_number > 0)
);
```

## Indeksy i ograniczenia

```sql
CREATE INDEX IF NOT EXISTS idx_documents_created_at
	ON documents(created_at DESC);

CREATE INDEX IF NOT EXISTS idx_documents_is_deleted
	ON documents(is_deleted)
	WHERE is_deleted = FALSE;

CREATE INDEX IF NOT EXISTS idx_document_versions_document_active
	ON document_versions(document_id, is_active)
	WHERE is_active = TRUE;

CREATE INDEX IF NOT EXISTS idx_document_versions_document_version
	ON document_versions(document_id, version_number DESC);

CREATE UNIQUE INDEX IF NOT EXISTS idx_document_versions_unique_active
	ON document_versions(document_id)
	WHERE is_active = TRUE;
```

Reguła biznesowa gwarantowana przez DB: tylko jedna aktywna wersja na dokument.

## Przepływ danych

1. Upload nowego pliku -> utworzenie wiersza w `documents` + pierwszego wiersza w `document_versions`.
2. Save z edytora -> dodanie kolejnego wiersza w `document_versions`, deaktywacja poprzedniej aktywnej wersji.
3. Restore version -> wybrana wersja staje się aktywna bez usuwania historii.

## Uwagi dotyczące migracji storage

- Migracja `003_migrate_storage_to_gcs.sql` wprowadza `storage_path` i `size_in_bytes`.
- Kolumna `content` jest oznaczona do usunięcia, ale instrukcja DROP jest celowo zakomentowana.
- Przed usunięciem `content` cała historyczna zawartość musi zostać przeniesiona do object storage.

## Typowe zapytania

### Pobranie metadanych aktywnej wersji

```sql
SELECT d.id,
	   d.name,
	   d.mime_type,
	   v.id AS version_id,
	   v.version_number,
	   v.storage_path,
	   v.size_in_bytes
FROM documents d
JOIN document_versions v ON v.document_id = d.id
WHERE d.id = :master_id
  AND d.is_deleted = FALSE
  AND v.is_active = TRUE;
```

### Pobranie pełnej historii wersji

```sql
SELECT id, document_id, version_number, created_at, created_by, is_active, size_in_bytes
FROM document_versions
WHERE document_id = :master_id
ORDER BY version_number DESC;
```

### Soft delete dokumentu

```sql
UPDATE documents
SET is_deleted = TRUE
WHERE id = :master_id;
```

