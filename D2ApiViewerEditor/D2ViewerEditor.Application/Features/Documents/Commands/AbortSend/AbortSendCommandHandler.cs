using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.AbortSend;

public class AbortSendCommandHandler
    : IRequestHandler<AbortSendCommand, Result<AbortSendResult>>
{
    private readonly IDocumentRepository _documentRepository;
    private readonly IDocumentDeliveryRepository _deliveryRepository;

    public AbortSendCommandHandler(
        IDocumentRepository documentRepository,
        IDocumentDeliveryRepository deliveryRepository)
    {
        _documentRepository = documentRepository;
        _deliveryRepository = deliveryRepository;
    }

    public async Task<Result<AbortSendResult>> Handle(
        AbortSendCommand request, CancellationToken cancellationToken)
    {
        var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
        if (document == null)
            return Result<AbortSendResult>.NotFound();

        try
        {
            var active = await _deliveryRepository.GetActiveByDocumentIdAsync(document.Id, cancellationToken);
            string? deliveryStatus = null;
            if (active is not null)
            {
                var delivery = await _deliveryRepository.GetByIdAsync(active.Id, cancellationToken);
                delivery?.Cancel();
                deliveryStatus = delivery?.Status.ToString();
            }

            document.MarkSendAborted();
            await _deliveryRepository.SaveChangesAsync(cancellationToken);

            return Result<AbortSendResult>.Success(
                new AbortSendResult(document.Id, document.Status.ToString(), deliveryStatus));
        }
        catch (InvalidOperationException ex)
        {
            // np. zadanie już w trakcie wysyłki przez workera (Sending) — nie można anulować.
            return Result<AbortSendResult>.Failure(ex.Message);
        }
    }
}
