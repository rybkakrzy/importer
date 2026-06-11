using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UpdateDeliveryRecipientUrl;

public class UpdateDeliveryRecipientUrlCommandHandler
    : IRequestHandler<UpdateDeliveryRecipientUrlCommand, Result<UpdateDeliveryRecipientUrlResult>>
{
    private readonly IDocumentDeliveryRepository _deliveryRepository;

    public UpdateDeliveryRecipientUrlCommandHandler(IDocumentDeliveryRepository deliveryRepository)
    {
        _deliveryRepository = deliveryRepository;
    }

    public async Task<Result<UpdateDeliveryRecipientUrlResult>> Handle(
        UpdateDeliveryRecipientUrlCommand request, CancellationToken cancellationToken)
    {
        var delivery = await _deliveryRepository.GetByIdAsync(request.DeliveryId, cancellationToken);
        if (delivery == null)
            return Result<UpdateDeliveryRecipientUrlResult>.NotFound("Nie znaleziono zadania wysyłki");

        try
        {
            delivery.UpdateRecipientUrl(request.RecipientUrl?.Trim() ?? string.Empty);
            await _deliveryRepository.SaveChangesAsync(cancellationToken);

            return Result<UpdateDeliveryRecipientUrlResult>.Success(
                new UpdateDeliveryRecipientUrlResult(delivery.Id, delivery.RecipientUrl, delivery.Status.ToString()));
        }
        catch (Exception ex) when (ex is InvalidOperationException or ArgumentException)
        {
            return Result<UpdateDeliveryRecipientUrlResult>.Failure(ex.Message);
        }
    }
}
