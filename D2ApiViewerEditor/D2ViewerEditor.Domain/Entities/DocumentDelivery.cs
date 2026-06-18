using D2ViewerEditor.Domain.Interfaces;

namespace D2ViewerEditor.Domain.Entities;

/// <summary>
/// Zadanie wysyłki finalnego dokumentu na returnUrl (finish-and-send) — aggregate root kolejki wysyłkowej.
/// Trzyma inwarianty cyklu życia (dozwolone przejścia statusów, planowanie retry, dead-letter po deadline).
/// Sama fizyczna wysyłka (HTTP/GCS) jest infrastrukturą — encja nie wie "jak", tylko "w jaki stan przejść".
/// </summary>
public class DocumentDelivery
{
    private const int MaxErrorLength = 4000;

    private DocumentDelivery() { } // EF Core

    public static DocumentDelivery Create(
        Guid id,
        Guid documentId,
        Guid sourceVersionId,
        string snapshotObjectName,
        long snapshotSizeBytes,
        string snapshotSha256,
        string recipientUrl,
        string createdBy,
        Guid correlationId,
        TimeSpan retentionWindow,
        string? corporateKey = null)
    {
        if (string.IsNullOrWhiteSpace(snapshotObjectName))
            throw new ArgumentException("Snapshot object name nie może być pusty", nameof(snapshotObjectName));

        if (string.IsNullOrWhiteSpace(snapshotSha256))
            throw new ArgumentException("Snapshot hash nie może być pusty", nameof(snapshotSha256));

        if (!IsValidRecipientUrl(recipientUrl))
            throw new ArgumentException("recipientUrl musi być absolutnym adresem http(s)", nameof(recipientUrl));

        if (retentionWindow <= TimeSpan.Zero)
            throw new ArgumentException("Okno ponawiania musi być dodatnie", nameof(retentionWindow));

        var now = DateTime.UtcNow;

        return new DocumentDelivery
        {
            Id = id,
            DocumentId = documentId,
            SourceVersionId = sourceVersionId,
            SnapshotObjectName = snapshotObjectName,
            SnapshotSizeBytes = snapshotSizeBytes,
            SnapshotSha256 = snapshotSha256,
            RecipientUrl = recipientUrl,
            Status = DeliveryStatus.Pending,
            AttemptCount = 0,
            CreatedAt = now,
            UpdatedAt = now,
            NextAttemptAt = now,
            DeadlineAt = now.Add(retentionWindow),
            CorrelationId = correlationId,
            CreatedBy = createdBy,
            CorporateKey = string.IsNullOrWhiteSpace(corporateKey) ? null : corporateKey.Trim()
        };
    }

    public Guid Id { get; private set; }
    public Guid DocumentId { get; private set; }

    /// <summary>Wersja edytowalna, której stan został zamrożony do snapshotu.</summary>
    public Guid SourceVersionId { get; private set; }

    /// <summary>Niezmienny obiekt GCS ze snapshotem treści (deliveries/{id}).</summary>
    public string SnapshotObjectName { get; private set; } = string.Empty;
    public long SnapshotSizeBytes { get; private set; }
    public string SnapshotSha256 { get; private set; } = string.Empty;

    public string RecipientUrl { get; private set; } = string.Empty;
    public DeliveryStatus Status { get; private set; }
    public int AttemptCount { get; private set; }

    public DateTime CreatedAt { get; private set; }
    public DateTime UpdatedAt { get; private set; }
    public DateTime? FirstAttemptAt { get; private set; }
    public DateTime? LastAttemptAt { get; private set; }
    public DateTime NextAttemptAt { get; private set; }

    /// <summary>created_at + okno ponawiania (24 h). Po przekroczeniu błąd retryable → DeadLettered.</summary>
    public DateTime DeadlineAt { get; private set; }

    /// <summary>Lease techniczny claimu (NIE status biznesowy). Po wygaśnięciu zadanie może przejąć inny worker.</summary>
    public DateTime? LockedUntil { get; private set; }
    public string? LockedBy { get; private set; }

    public string? LastError { get; private set; }
    public Guid CorrelationId { get; private set; }
    public string CreatedBy { get; private set; } = string.Empty;

    /// <summary>CorporateKey of the user who finished the document; sent to the recipient as an
    /// identifying field. Null when the finishing user had no CorporateKey claim.</summary>
    public string? CorporateKey { get; private set; }

    /// <summary>Czy zadanie jest w stanie końcowym (nie podlega dalszemu przetwarzaniu).</summary>
    public bool IsTerminal =>
        Status is DeliveryStatus.Sent or DeliveryStatus.FailedPermanently
               or DeliveryStatus.DeadLettered or DeliveryStatus.Cancelled;

    public void MarkSent()
    {
        Status = DeliveryStatus.Sent;
        ClearLease();
        LastError = null;
        UpdatedAt = DateTime.UtcNow;
    }

    public void MarkPermanentFailure(string error)
    {
        Status = DeliveryStatus.FailedPermanently;
        ClearLease();
        LastError = Truncate(error);
        UpdatedAt = DateTime.UtcNow;
    }

    /// <summary>
    /// Planuje kolejną próbę albo — jeśli wyliczony termin przekracza deadline — oznacza zadanie jako DeadLettered.
    /// Zwraca true, gdy zaplanowano retry; false, gdy zadanie trafiło do stanu końcowego.
    /// </summary>
    public bool ScheduleRetryOrDeadLetter(string error, IBackoffStrategy backoff)
    {
        var now = DateTime.UtcNow;
        var next = backoff.NextAttempt(Math.Max(1, AttemptCount), now);

        if (next > DeadlineAt)
        {
            Status = DeliveryStatus.DeadLettered;
        }
        else
        {
            Status = DeliveryStatus.RetryScheduled;
            NextAttemptAt = next;
        }

        ClearLease();
        LastError = Truncate(error);
        UpdatedAt = now;
        return Status == DeliveryStatus.RetryScheduled;
    }

    /// <summary>
    /// Ręczne wznowienie zadania ("Wznów") — wraca do kolejki i wysyła od razu. Dotyczy zadań
    /// nieudanych (DeadLettered / FailedPermanently), zaplanowanych na później (RetryScheduled — wyślij
    /// teraz zamiast czekać na backoff) oraz anulowanych (Cancelled). Wydłuża deadline o podane okno,
    /// żeby zadanie miało realną szansę na ponowną wysyłkę. NIE dotyczy Sent (już wysłane), Sending
    /// (próba w toku) ani Pending (już w kolejce).
    /// </summary>
    public void Requeue(TimeSpan retentionWindow)
    {
        if (Status is not (DeliveryStatus.DeadLettered or DeliveryStatus.FailedPermanently
                        or DeliveryStatus.RetryScheduled or DeliveryStatus.Cancelled))
            throw new InvalidOperationException(
                "Wznowienie dotyczy zadań nieudanych, zaplanowanych lub anulowanych (nie: wysłane/w toku/oczekujące)");

        var now = DateTime.UtcNow;
        Status = DeliveryStatus.Pending;
        NextAttemptAt = now;
        DeadlineAt = now.Add(retentionWindow);
        ClearLease();
        LastError = null;
        UpdatedAt = now;
    }

    /// <summary>
    /// Ręczne anulowanie zadania ("Anuluj") — zatrzymuje wysyłkę i przechodzi w stan końcowy Cancelled.
    /// Dozwolone tylko dla zadań jeszcze nieprzetworzonych i niezablokowanych przez workera
    /// (Pending / RetryScheduled). Próby w toku (Sending) i stany końcowe nie podlegają anulowaniu.
    /// </summary>
    public void Cancel()
    {
        if (Status is not (DeliveryStatus.Pending or DeliveryStatus.RetryScheduled))
            throw new InvalidOperationException(
                "Anulować można tylko zadanie oczekujące lub zaplanowane (nie: w toku/wysłane/zakończone)");

        Status = DeliveryStatus.Cancelled;
        ClearLease();
        LastError = "Anulowano ręcznie (administrator).";
        UpdatedAt = DateTime.UtcNow;
    }

    /// <summary>
    /// Zmiana adresu odbiorcy ("Edytuj returnUrl") — np. korekta błędnego adresu na zadaniu, które
    /// nie doszło. Dozwolona dla zadań niewysłanych i nieblokowanych przez workera (NIE: Sent/Sending).
    /// Po zmianie zadanie zwykle wymaga „Wznów", by ponowić wysyłkę na nowy adres.
    /// </summary>
    public void UpdateRecipientUrl(string url)
    {
        if (Status is DeliveryStatus.Sent or DeliveryStatus.Sending)
            throw new InvalidOperationException(
                "Adresu odbiorcy nie można zmienić dla zadania wysłanego ani w trakcie wysyłki");

        if (!IsValidRecipientUrl(url))
            throw new ArgumentException("recipientUrl musi być absolutnym adresem http(s)", nameof(url));

        RecipientUrl = url;
        UpdatedAt = DateTime.UtcNow;
    }

    public static bool IsValidRecipientUrl(string? url) =>
        !string.IsNullOrWhiteSpace(url)
        && Uri.TryCreate(url, UriKind.Absolute, out var uri)
        && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps);

    private void ClearLease()
    {
        LockedUntil = null;
        LockedBy = null;
    }

    private static string Truncate(string s) =>
        string.IsNullOrEmpty(s) ? string.Empty : (s.Length > MaxErrorLength ? s[..MaxErrorLength] : s);
}
