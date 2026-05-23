-- Add metadata column to documents for external integration payload (returnUrl, classification, etc.)
-- Stored as raw JSON text — schema-less by design

ALTER TABLE documents
    ADD COLUMN IF NOT EXISTS metadata TEXT NULL;

COMMENT ON COLUMN documents.metadata IS 'Metadane od aplikacji zewnętrznej w formacie JSON (np. { "returnUrl": "...", "classification": "C2" })';
