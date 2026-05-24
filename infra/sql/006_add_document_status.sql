-- Add lifecycle status to documents.
-- Values (string, matching enum DocumentStatus):
--   Saved, Editing, Sending, DeliveryFailed, Sent

ALTER TABLE documents
    ADD COLUMN IF NOT EXISTS status VARCHAR(40) NOT NULL DEFAULT 'Saved';

COMMENT ON COLUMN documents.status IS 'Document lifecycle status: Saved | Editing | Sending | DeliveryFailed | Sent';
