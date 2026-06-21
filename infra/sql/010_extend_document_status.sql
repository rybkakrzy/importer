-- Extend the documents.status lifecycle with the finish-and-send states.
-- The column is a free-form VARCHAR(40) (no CHECK constraint), so adding new enum values
-- requires no schema change — this migration only refreshes the documentation comment.
--
-- New values (string, matching enum DocumentStatus):
--   Queued       — "Zlecono do wysyłki" (delivery requested / continued in background)
--   SendAborted  — "UzytkownikPrzerwałWysyłkę" (user aborted after a failed first attempt)

COMMENT ON COLUMN documents.status IS
    'Document lifecycle status: Saved | Editing | Queued | Sending | DeliveryFailed | Sent | SendAborted';
