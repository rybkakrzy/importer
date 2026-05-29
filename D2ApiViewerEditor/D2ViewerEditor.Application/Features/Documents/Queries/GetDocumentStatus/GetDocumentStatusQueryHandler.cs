using D2ViewerEditor.Application.Features.Documents.Common;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentStatus;

public class GetDocumentStatusQueryHandler
    : IRequestHandler<GetDocumentStatusQuery, Result<DocumentStatusDto>>
{
    private readonly IDocumentRepository _documentRepository;
    private readonly IDocumentDeliveryRepository _deliveryRepository;

    public GetDocumentStatusQueryHandler(
        IDocumentRepository documentRepository,
        IDocumentDeliveryRepository deliveryRepository)
    {
        _documentRepository = documentRepository;
        _deliveryRepository = deliveryRepository;
    }

    public async Task<Result<DocumentStatusDto>> Handle(
        GetDocumentStatusQuery request,
        CancellationToken cancellationToken)
    {
        var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
        if (document == null)
            return Result<DocumentStatusDto>.NotFound();

        var meta = ExternalDocumentMetadata.Parse(document.Metadata);
        var activeVersion = document.GetActiveVersion();
        var delivery = await _deliveryRepository.GetActiveByDocumentIdAsync(document.Id, cancellationToken);

        return Result<DocumentStatusDto>.Success(new DocumentStatusDto(
            MasterId: document.Id,
            Status: document.Status.ToString(),
            IsLocked: document.Status == DocumentStatus.Editing,
            HasCallbackUrl: !string.IsNullOrWhiteSpace(meta.ReturnUrl),
            UserDownload: meta.IsUserDownloadAllowed,
            ActiveVersionId: activeVersion?.Id,
            ActiveVersionNumber: activeVersion?.VersionNumber,
            ActiveVersionModifiedAt: activeVersion?.ModifiedAt,
            LatestDelivery: delivery is null
                ? null
                : new DocumentStatusDeliveryDto(
                    DeliveryId: delivery.Id,
                    Status: delivery.Status.ToString(),
                    AttemptCount: delivery.AttemptCount,
                    LastAttemptAt: delivery.LastAttemptAt,
                    NextAttemptAt: delivery.NextAttemptAt,
                    DeadlineAt: delivery.DeadlineAt)
        ));
    }
}
