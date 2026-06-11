using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.CancelDelivery;

public class CancelDeliveryCommandHandler
    : IRequestHandler<CancelDeliveryCommand, Result<CancelDeliveryResult>>
{
    private readonly IDocumentDeliveryRepository _deliveryRepository;
    private readonly IDocumentRepository _documentRepository;

    public CancelDeliveryCommandHandler(
        IDocumentDeliveryRepository deliveryRepository,
        IDocumentRepository documentRepository)
    {
        _deliveryRepository = deliveryRepository;
        _documentRepository = documentRepository;
    }

    public async Task<Result<CancelDeliveryResult>> Handle(
        CancelDeliveryCommand request, CancellationToken cancellationToken)
    {
        var delivery = await _deliveryRepository.GetByIdAsync(request.DeliveryId, cancellationToken);
        if (delivery == null)
            return Result<CancelDeliveryResult>.NotFound("Nie znaleziono zadania wysyłki");

        try
        {
            delivery.Cancel();

            // Anulowanie aktywnej wysyłki zdejmuje dokument ze stanu „Sending" → wraca do „Saved".
            var document = await _documentRepository.GetByIdWithVersionsAsync(delivery.DocumentId, cancellationToken);
            document?.MarkSaved();

            await _deliveryRepository.SaveChangesAsync(cancellationToken);

            return Result<CancelDeliveryResult>.Success(
                new CancelDeliveryResult(delivery.Id, delivery.Status.ToString()));
        }
        catch (InvalidOperationException ex)
        {
            return Result<CancelDeliveryResult>.Failure(ex.Message);
        }
    }
}
