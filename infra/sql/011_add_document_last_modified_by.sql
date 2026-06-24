-- Records who last modified the document from the editor — the CorporateKey carried by the
-- Entra ID token (`corpKey` claim) and forwarded by the GUI on every save/auto-save/finish call.
-- Surfaced in the admin file list and the delivery (files-to-send) list. Nullable: a save may
-- arrive without a CorporateKey (e.g. local dev bypass), in which case the previous value is kept.
-- Not a secret — a corporate identifier, displayed only to administrators.

ALTER TABLE documents
    ADD COLUMN IF NOT EXISTS last_modified_by varchar(255) NULL;
