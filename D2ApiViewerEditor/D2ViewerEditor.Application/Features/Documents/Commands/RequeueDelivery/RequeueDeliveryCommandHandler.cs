using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.RequeueDelivery;

public class RequeueDeliveryCommandHandler
    : IRequestHandler<RequeueDeliveryCommand, Result<RequeueDeliveryResult>>
{
    private static readonly TimeSpan RetentionWindow = TimeSpan.FromHours(24);

    private readonly IDocumentDeliveryRepository _deliveryRepository;
    private readonly IDocumentRepository _documentRepository;

    public RequeueDeliveryCommandHandler(
        IDocumentDeliveryRepository deliveryRepository,
        IDocumentRepository documentRepository)
    {
        _deliveryRepository = deliveryRepository;
        _documentRepository = documentRepository;
    }

    public async Task<Result<RequeueDeliveryResult>> Handle(
        RequeueDeliveryCommand request, CancellationToken cancellationToken)
    {
        var delivery = await _deliveryRepository.GetByIdAsync(request.DeliveryId, cancellationToken);
        if (delivery == null)
            return Result<RequeueDeliveryResult>.NotFound("Nie znaleziono zadania wysyłki");

        try
        {
            delivery.Requeue(RetentionWindow);

            var document = await _documentRepository.GetByIdWithVersionsAsync(delivery.DocumentId, cancellationToken);
            document?.MarkSending();

            await _deliveryRepository.SaveChangesAsync(cancellationToken);

            return Result<RequeueDeliveryResult>.Success(
                new RequeueDeliveryResult(delivery.Id, delivery.Status.ToString()));
        }
        catch (InvalidOperationException ex)
        {
            return Result<RequeueDeliveryResult>.Failure(ex.Message);
        }
    }
}
