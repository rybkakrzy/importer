-- Initial schema for document management with versioning
-- PostgreSQL 16+

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

CREATE TABLE IF NOT EXISTS document_versions (
    id UUID PRIMARY KEY,
    document_id UUID NOT NULL,
    version_number INT NOT NULL,
    content BYTEA NOT NULL,
    created_at TIMESTAMP WITHOUT TIME ZONE NOT NULL,
    created_by VARCHAR(255) NOT NULL,
    is_active BOOLEAN NOT NULL DEFAULT FALSE,
    
    CONSTRAINT fk_document_versions_document FOREIGN KEY (document_id) 
        REFERENCES documents(id) ON DELETE CASCADE,
    CONSTRAINT chk_version_number_positive CHECK (version_number > 0),
    CONSTRAINT chk_content_not_empty CHECK (length(content) > 0)
);

-- Comments for documentation
COMMENT ON TABLE documents IS 'Główna tabela dokumentów (aggregate root)';
COMMENT ON TABLE document_versions IS 'Wersje dokumentów - każde zapisanie z GUI tworzy nową wersję';

COMMENT ON COLUMN documents.id IS 'GUID mastera (guid_master)';
COMMENT ON COLUMN documents.name IS 'Nazwa dokumentu';
COMMENT ON COLUMN documents.mime_type IS 'Typ MIME (application/pdf, image/jpeg, etc.)';
COMMENT ON COLUMN documents.is_deleted IS 'Soft delete flag';

COMMENT ON COLUMN document_versions.id IS 'GUID wersji (guid_wersji)';
COMMENT ON COLUMN document_versions.document_id IS 'FK do documents.id';
COMMENT ON COLUMN document_versions.version_number IS 'Numer wersji (auto-increment przez aplikację)';
COMMENT ON COLUMN document_versions.content IS 'Zawartość binarna dokumentu (BYTEA)';
COMMENT ON COLUMN document_versions.is_active IS 'Czy to aktywna wersja (tylko jedna może być active=true)';
