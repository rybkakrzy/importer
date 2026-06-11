namespace D2ViewerEditor.Domain.Entities;

/// <summary>
/// Status zadania wysyłki finalnego dokumentu na returnUrl (finish-and-send).
/// Trzymany w bazie jako string (zob. DocumentDeliveryConfiguration).
/// Uwaga: lease techniczny (locked_until) jest osobny od statusu biznesowego — brak statusu "Locked".
/// </summary>
public enum DeliveryStatus
{
    /// <summary>Zadanie utworzone, czeka na pierwszą próbę.</summary>
    Pending,

    /// <summary>Próba wysyłki w toku (zadanie zaclaimowane przez workera).</summary>
    Sending,

    /// <summary>Próba nieudana (błąd retryable) — zaplanowano kolejną próbę (next_attempt_at).</summary>
    RetryScheduled,

    /// <summary>Dokument skutecznie wysłany na returnUrl.</summary>
    Sent,

    /// <summary>Błąd non-retryable (np. zły adres odbiorcy) — nie ma sensu ponawiać.</summary>
    FailedPermanently,

    /// <summary>Wyczerpano okno ponawiania (24 h) bez skutecznej wysyłki.</summary>
    DeadLettered,

    /// <summary>Zadanie anulowane ręcznie (administrator) — nie podlega dalszemu przetwarzaniu.</summary>
    Cancelled
}
