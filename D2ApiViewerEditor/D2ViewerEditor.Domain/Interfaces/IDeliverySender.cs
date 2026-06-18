namespace D2ViewerEditor.Domain.Interfaces;

/// <summary>
/// Klasyfikacja wyniku pojedynczej próby wysyłki.
/// </summary>
public enum DeliveryOutcome
{
    Succeeded,
    RetryableError,
    PermanentError
}

/// <summary>
/// Dane jednej próby dostarczenia: niezmienny snapshot + adres odbiorcy + hash do idempotencji,
/// oraz dane identyfikacyjne wysyłane obok pliku (masterId, versionId, corporateKey).
/// </summary>
public sealed record DeliveryDispatch(
    Guid DeliveryId,
    string RecipientUrl,
    byte[] Content,
    string Sha256,
    Guid MasterId,
    Guid VersionId,
    string? CorporateKey);

/// <summary>
/// Wynik próby wysyłki z perspektywy infrastruktury (klasyfikacja + ewentualny komunikat błędu).
/// </summary>
public sealed record DeliveryResult(DeliveryOutcome Outcome, string? Error)
{
    public static DeliveryResult Succeeded() => new(DeliveryOutcome.Succeeded, null);
    public static DeliveryResult Retryable(string error) => new(DeliveryOutcome.RetryableError, error);
    public static DeliveryResult Permanent(string error) => new(DeliveryOutcome.PermanentError, error);
}

/// <summary>
/// Abstrakcja fizycznej wysyłki dokumentu do odbiorcy (returnUrl). At-least-once.
/// </summary>
public interface IDeliverySender
{
    Task<DeliveryResult> SendAsync(DeliveryDispatch dispatch, CancellationToken cancellationToken = default);
}
