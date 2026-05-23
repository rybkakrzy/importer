-- Add modified_at column to document_versions for in-place auto-save overwrites.
-- Auto-save nadpisuje wersję edytowalną (v2) w miejscu — created_at zostaje, modified_at śledzi ostatni zapis.

ALTER TABLE document_versions
    ADD COLUMN IF NOT EXISTS modified_at TIMESTAMPTZ NULL;

COMMENT ON COLUMN document_versions.modified_at IS 'Znacznik ostatniego nadpisania zawartości wersji (auto-save). NULL gdy wersja nie była modyfikowana po utworzeniu.';
