using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Application.Features.Documents.Common;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentMetadata;

/// <summary>
/// Handler zwracający zdeserializowane metadane dokumentu (returnUrl, classification, userDownload).
/// Pełni też rolę bramy dostępu dla GUI: GUI woła ten endpoint jako pierwszy,
/// więc brak uprawnień (CorporateKey) jest tu wykrywany przed pobraniem treści.
/// </summary>
public class GetDocumentMetadataQueryHandler
    : IRequestHandler<GetDocumentMetadataQuery, Result<DocumentMetadataDto>>
{
    private readonly IDocumentRepository _documentRepository;
    private readonly IDocumentAccessGuard _accessGuard;

    public GetDocumentMetadataQueryHandler(
        IDocumentRepository documentRepository,
        IDocumentAccessGuard accessGuard)
    {
        _documentRepository = documentRepository;
        _accessGuard = accessGuard;
    }

    public async Task<Result<DocumentMetadataDto>> Handle(
        GetDocumentMetadataQuery request,
        CancellationToken cancellationToken)
    {
        var document = await _documentRepository.GetByIdAsync(request.MasterId, cancellationToken);
        if (document == null)
            return Result<DocumentMetadataDto>.NotFound();

        if (!_accessGuard.IsViewAllowed(document.Metadata))
            return Result<DocumentMetadataDto>.Forbidden("Brak uprawnień do podglądu tego dokumentu.");

        var meta = ExternalDocumentMetadata.Parse(document.Metadata);

        return Result<DocumentMetadataDto>.Success(new DocumentMetadataDto(
            MasterId: document.Id,
            MimeType: document.MimeType,
            ReturnUrl: meta.ReturnUrl,
            Classification: meta.Classification,
            UserDownload: meta.IsUserDownloadAllowed
        ));
    }
}
