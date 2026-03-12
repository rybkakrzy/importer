-- Performance indexes for document management

-- Index dla wyszukiwania dokumentów po dacie utworzenia
CREATE INDEX IF NOT EXISTS idx_documents_created_at 
    ON documents(created_at DESC);

-- Index dla filtrowania nie-usuniętych dokumentów
CREATE INDEX IF NOT EXISTS idx_documents_is_deleted 
    ON documents(is_deleted) 
    WHERE is_deleted = FALSE;

-- Composite index dla pobierania aktywnej wersji dokumentu
CREATE INDEX IF NOT EXISTS idx_document_versions_document_active 
    ON document_versions(document_id, is_active) 
    WHERE is_active = TRUE;

-- Index dla wersji dokumentu (sortowanie po numerze wersji)
CREATE INDEX IF NOT EXISTS idx_document_versions_document_version 
    ON document_versions(document_id, version_number DESC);

-- Index dla daty utworzenia wersji (historia zmian)
CREATE INDEX IF NOT EXISTS idx_document_versions_created_at 
    ON document_versions(created_at DESC);

-- Unique constraint - tylko jedna aktywna wersja na dokument
CREATE UNIQUE INDEX IF NOT EXISTS idx_document_versions_unique_active 
    ON document_versions(document_id) 
    WHERE is_active = TRUE;

-- Comments
COMMENT ON INDEX idx_documents_created_at IS 'Index dla sortowania dokumentów po dacie';
COMMENT ON INDEX idx_documents_is_deleted IS 'Partial index dla aktywnych dokumentów';
COMMENT ON INDEX idx_document_versions_document_active IS 'Index dla szybkiego pobierania aktywnej wersji';
COMMENT ON INDEX idx_document_versions_unique_active IS 'Zapewnia że tylko jedna wersja jest aktywna';
