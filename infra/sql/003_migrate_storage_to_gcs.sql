-- Migracja: przeniesienie zawartości binarnej z bazy do Google Cloud Storage
-- Kolumna 'content' (bytea) zastąpiona przez 'storage_path' (varchar) + 'size_in_bytes' (bigint)

-- 1. Dodaj nowe kolumny
ALTER TABLE document_versions ADD COLUMN IF NOT EXISTS storage_path VARCHAR(500);
ALTER TABLE document_versions ADD COLUMN IF NOT EXISTS size_in_bytes BIGINT DEFAULT 0;

-- 2. Wypełnij size_in_bytes na podstawie istniejącego contentu (jeśli istnieje)
UPDATE document_versions
SET size_in_bytes = COALESCE(octet_length(content), 0)
WHERE size_in_bytes = 0 OR size_in_bytes IS NULL;

-- 3. Ustaw tymczasowy storage_path dla istniejących rekordów
--    UWAGA: Przed uruchomieniem tej migracji na produkcji należy najpierw
--    przeprowadzić migrację danych z kolumny 'content' do GCS bucketów
--    i ustawić prawidłowe wartości storage_path.
UPDATE document_versions
SET storage_path = 'documents/' || id
WHERE storage_path IS NULL;

-- 4. Ustaw kolumny jako NOT NULL po wypełnieniu
ALTER TABLE document_versions ALTER COLUMN storage_path SET NOT NULL;
ALTER TABLE document_versions ALTER COLUMN size_in_bytes SET NOT NULL;

-- 5. Usuń kolumnę content (UWAGA: nieodwracalne!)
--    Odkomentuj po zweryfikowaniu że dane zostały zmigrowane do GCS
-- ALTER TABLE document_versions DROP COLUMN IF EXISTS content;
