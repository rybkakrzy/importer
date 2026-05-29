using D2ViewerEditor.Application.Features.Documents.Common;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentMetadata;

public class GetDocumentMetadataQueryHandler
    : IRequestHandler<GetDocumentMetadataQuery, Result<DocumentMetadataDto>>
{
    private readonly IDocumentRepository _documentRepository;

    public GetDocumentMetadataQueryHandler(IDocumentRepository documentRepository)
    {
        _documentRepository = documentRepository;
    }

    public async Task<Result<DocumentMetadataDto>> Handle(
        GetDocumentMetadataQuery request,
        CancellationToken cancellationToken)
    {
        var document = await _documentRepository.GetByIdAsync(request.MasterId, cancellationToken);
        if (document == null)
            return Result<DocumentMetadataDto>.NotFound();

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
