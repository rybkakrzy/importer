using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Application.Common.Security;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UpdateDeliveryRecipientUrl;

public class UpdateDeliveryRecipientUrlCommandHandler
    : IRequestHandler<UpdateDeliveryRecipientUrlCommand, Result<UpdateDeliveryRecipientUrlResult>>
{
    private readonly IDocumentDeliveryRepository _deliveryRepository;
    private readonly IReturnUrlValidator _returnUrlValidator;

    public UpdateDeliveryRecipientUrlCommandHandler(
        IDocumentDeliveryRepository deliveryRepository,
        IReturnUrlValidator returnUrlValidator)
    {
        _deliveryRepository = deliveryRepository;
        _returnUrlValidator = returnUrlValidator;
    }

    public async Task<Result<UpdateDeliveryRecipientUrlResult>> Handle(
        UpdateDeliveryRecipientUrlCommand request, CancellationToken cancellationToken)
    {
        var delivery = await _deliveryRepository.GetByIdAsync(request.DeliveryId, cancellationToken);
        if (delivery == null)
            return Result<UpdateDeliveryRecipientUrlResult>.NotFound("Nie znaleziono zadania wysyłki");

        try
        {
            var urlValidation = _returnUrlValidator.Validate(request.RecipientUrl);
            if (!urlValidation.IsValid)
                return Result<UpdateDeliveryRecipientUrlResult>.Failure(urlValidation.Error!);

            delivery.UpdateRecipientUrl(urlValidation.NormalizedUrl!);
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
