using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDeliveryStatus;

/// <summary>
/// Zapytanie o status zadania wysyłki (polling z GUI).
/// </summary>
public record GetDeliveryStatusQuery(Guid DeliveryId) : IRequest<Result<DeliveryStatusDto>>;

/// <summary>
/// Status zadania wysyłki widoczny dla klienta.
/// </summary>
public record DeliveryStatusDto(
    Guid DeliveryId,
    Guid DocumentId,
    string Status,
    int AttemptCount,
    DateTime? LastAttemptAt,
    DateTime? NextAttemptAt,
    string? LastError,
    DateTime UpdatedAt
);
