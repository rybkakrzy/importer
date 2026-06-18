-- Persists the finishing user's CorporateKey on the delivery so the async worker can include it as
-- an identifying field when posting the file to the recipient (returnUrl). Nullable: the finishing
-- user may have no CorporateKey claim. Not a secret — a corporate identifier — but kept out of logs
-- by value (logged as business context only where diagnostically justified).

ALTER TABLE document_deliveries
    ADD COLUMN IF NOT EXISTS corporate_key varchar(255) NULL;
