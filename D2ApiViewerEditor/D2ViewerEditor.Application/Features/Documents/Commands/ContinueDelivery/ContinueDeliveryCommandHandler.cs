using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.ContinueDelivery;

public class ContinueDeliveryCommandHandler
    : IRequestHandler<ContinueDeliveryCommand, Result<ContinueDeliveryResult>>
{
    // Zgodne z oknem ponawiania zadania (24 h) — Requeue odświeża deadline na nowo.
    private static readonly TimeSpan RetentionWindow = TimeSpan.FromHours(24);

    private readonly IDocumentRepository _documentRepository;
    private readonly IDocumentDeliveryRepository _deliveryRepository;

    public ContinueDeliveryCommandHandler(
        IDocumentRepository documentRepository,
        IDocumentDeliveryRepository deliveryRepository)
    {
        _documentRepository = documentRepository;
        _deliveryRepository = deliveryRepository;
    }

    public async Task<Result<ContinueDeliveryResult>> Handle(
        ContinueDeliveryCommand request, CancellationToken cancellationToken)
    {
        var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
        if (document == null)
            return Result<ContinueDeliveryResult>.NotFound();

        var active = await _deliveryRepository.GetActiveByDocumentIdAsync(document.Id, cancellationToken);
        if (active == null)
            return Result<ContinueDeliveryResult>.Failure("Brak zadania wysyłki do kontynuowania");

        var delivery = await _deliveryRepository.GetByIdAsync(active.Id, cancellationToken);
        if (delivery == null)
            return Result<ContinueDeliveryResult>.Failure("Brak zadania wysyłki do kontynuowania");

        try
        {
            delivery.Requeue(RetentionWindow);
            document.MarkQueued();
            await _deliveryRepository.SaveChangesAsync(cancellationToken);

            return Result<ContinueDeliveryResult>.Success(new ContinueDeliveryResult(
                document.Id, document.Status.ToString(), delivery.Id, delivery.Status.ToString()));
        }
        catch (InvalidOperationException ex)
        {
            return Result<ContinueDeliveryResult>.Failure(ex.Message);
        }
    }
}
