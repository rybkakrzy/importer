using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDeliveriesByStatus;

/// <summary>
/// Lista zadań wysyłki w danym statusie (monitoring / panel admina).
/// </summary>
public record GetDeliveriesByStatusQuery(string Status, int Skip = 0, int Take = 100)
    : IRequest<Result<IReadOnlyList<DeliveryListItemDto>>>;

public record DeliveryListItemDto(
    Guid DeliveryId,
    Guid DocumentId,
    string Status,
    int AttemptCount,
    DateTime CreatedAt,
    DateTime? LastAttemptAt,
    DateTime? NextAttemptAt,
    DateTime DeadlineAt,
    string? LastError);

public class GetDeliveriesByStatusQueryHandler
    : IRequestHandler<GetDeliveriesByStatusQuery, Result<IReadOnlyList<DeliveryListItemDto>>>
{
    private readonly IDocumentDeliveryRepository _deliveryRepository;

    public GetDeliveriesByStatusQueryHandler(IDocumentDeliveryRepository deliveryRepository)
    {
        _deliveryRepository = deliveryRepository;
    }

    public async Task<Result<IReadOnlyList<DeliveryListItemDto>>> Handle(
        GetDeliveriesByStatusQuery request, CancellationToken cancellationToken)
    {
        if (!Enum.TryParse<DeliveryStatus>(request.Status, ignoreCase: true, out var status))
            return Result<IReadOnlyList<DeliveryListItemDto>>.Failure(
                $"Nieznany status: {request.Status}");

        var items = await _deliveryRepository.GetByStatusAsync(
            status, request.Skip, request.Take, cancellationToken);

        var dtos = items
            .Select(d => new DeliveryListItemDto(
                d.Id, d.DocumentId, d.Status.ToString(), d.AttemptCount,
                d.CreatedAt, d.LastAttemptAt, d.NextAttemptAt, d.DeadlineAt, d.LastError))
            .ToList();

        return Result<IReadOnlyList<DeliveryListItemDto>>.Success(dtos);
    }
}
