using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDeliveryStatus;

public class GetDeliveryStatusQueryHandler
    : IRequestHandler<GetDeliveryStatusQuery, Result<DeliveryStatusDto>>
{
    private readonly IDocumentDeliveryRepository _deliveryRepository;

    public GetDeliveryStatusQueryHandler(IDocumentDeliveryRepository deliveryRepository)
    {
        _deliveryRepository = deliveryRepository;
    }

    public async Task<Result<DeliveryStatusDto>> Handle(
        GetDeliveryStatusQuery request, CancellationToken cancellationToken)
    {
        var delivery = await _deliveryRepository.GetByIdAsync(request.DeliveryId, cancellationToken);
        if (delivery == null)
            return Result<DeliveryStatusDto>.NotFound("Nie znaleziono zadania wysyłki");

        return Result<DeliveryStatusDto>.Success(new DeliveryStatusDto(
            DeliveryId: delivery.Id,
            DocumentId: delivery.DocumentId,
            Status: delivery.Status.ToString(),
            AttemptCount: delivery.AttemptCount,
            LastAttemptAt: delivery.LastAttemptAt,
            NextAttemptAt: delivery.NextAttemptAt,
            LastError: delivery.LastError,
            UpdatedAt: delivery.UpdatedAt));
    }
}
