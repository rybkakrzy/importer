-- Finish-and-send: kolejka zadań wysyłki finalnego dokumentu na returnUrl.
-- At-least-once delivery; claim przez FOR UPDATE SKIP LOCKED + lease (locked_until).
-- Statusy (string, zgodne z enum DeliveryStatus):
--   Pending, Sending, RetryScheduled, Sent, FailedPermanently, DeadLettered

CREATE TABLE IF NOT EXISTS document_deliveries (
    id                   UUID PRIMARY KEY,
    document_id          UUID NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
    source_version_id    UUID NOT NULL,
    snapshot_object_name TEXT NOT NULL,
    snapshot_size_bytes  BIGINT NOT NULL,
    snapshot_sha256      TEXT NOT NULL,
    recipient_url        TEXT NOT NULL,
    status               VARCHAR(32) NOT NULL DEFAULT 'Pending',
    attempt_count        INT NOT NULL DEFAULT 0,
    created_at           TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at           TIMESTAMPTZ NOT NULL DEFAULT now(),
    first_attempt_at     TIMESTAMPTZ NULL,
    last_attempt_at      TIMESTAMPTZ NULL,
    next_attempt_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    deadline_at          TIMESTAMPTZ NOT NULL,
    locked_until         TIMESTAMPTZ NULL,
    locked_by            VARCHAR(128) NULL,
    last_error           TEXT NULL,
    correlation_id       UUID NOT NULL,
    created_by           VARCHAR(255) NOT NULL,
    CONSTRAINT ck_document_deliveries_status
        CHECK (status IN ('Pending','Sending','RetryScheduled','Sent','FailedPermanently','DeadLettered'))
);

-- Claim gotowych zadań: szukane po next_attempt_at w stanach kolejkowanych.
CREATE INDEX IF NOT EXISTS ix_document_deliveries_due
    ON document_deliveries (next_attempt_at)
    WHERE status IN ('Pending','RetryScheduled');

-- Reclaim zawieszonych (worker padł): Sending z wygasłym lease.
CREATE INDEX IF NOT EXISTS ix_document_deliveries_stuck
    ON document_deliveries (locked_until)
    WHERE status = 'Sending';

-- Idempotencja: tylko JEDNO aktywne (nieterminalne) zadanie na dokument.
CREATE UNIQUE INDEX IF NOT EXISTS ux_document_deliveries_active_per_document
    ON document_deliveries (document_id)
    WHERE status IN ('Pending','Sending','RetryScheduled');

-- Monitoring (liczność per status).
CREATE INDEX IF NOT EXISTS ix_document_deliveries_status
    ON document_deliveries (status);

COMMENT ON TABLE  document_deliveries IS 'Kolejka wysyłki finalnego dokumentu na returnUrl (finish-and-send).';
COMMENT ON COLUMN document_deliveries.snapshot_object_name IS 'Niezmienny snapshot pliku w GCS zamrożony w chwili "Zakończ".';
COMMENT ON COLUMN document_deliveries.locked_until IS 'Lease techniczny claimu; NIE jest statusem biznesowym.';
COMMENT ON COLUMN document_deliveries.deadline_at IS 'created_at + okno ponawiania (24 h); po przekroczeniu retryable -> DeadLettered.';
COMMENT ON COLUMN document_deliveries.correlation_id IS 'Identyfikator korelacji do śledzenia w logach.';
