-- Dodaje status 'Cancelled' do kolejki wysyłki (ręczne „Anuluj" w panelu admina).
-- Status końcowy: zadanie anulowane nie jest claimowane przez workera (poza zbiorami due/active).
-- Indeksy częściowe (due/active) bazują na Pending/Sending/RetryScheduled, więc Cancelled jest z nich
-- naturalnie wykluczony — nie wymagają zmiany.

ALTER TABLE document_deliveries
    DROP CONSTRAINT IF EXISTS ck_document_deliveries_status;

ALTER TABLE document_deliveries
    ADD CONSTRAINT ck_document_deliveries_status
        CHECK (status IN ('Pending','Sending','RetryScheduled','Sent','FailedPermanently','DeadLettered','Cancelled'));
