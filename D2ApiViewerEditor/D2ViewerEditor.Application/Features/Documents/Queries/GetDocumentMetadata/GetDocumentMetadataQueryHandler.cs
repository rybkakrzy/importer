using System.Text.Json;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentMetadata;

/// <summary>
/// Handler zwracający zdeserializowane metadane dokumentu (returnUrl, classification).
/// </summary>
public class GetDocumentMetadataQueryHandler
    : IRequestHandler<GetDocumentMetadataQuery, Result<DocumentMetadataDto>>
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

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

        string? returnUrl = null;
        string? classification = null;

        if (!string.IsNullOrWhiteSpace(document.Metadata))
        {
            try
            {
                var parsed = JsonSerializer.Deserialize<ExternalMetadata>(document.Metadata, JsonOptions);
                returnUrl = parsed?.ReturnUrl;
                classification = parsed?.Classification;
            }
            catch (JsonException)
            {
                // Metadane zapisane w nieoczekiwanym kształcie — zwracamy puste pola zamiast wywalać request.
            }
        }

        return Result<DocumentMetadataDto>.Success(new DocumentMetadataDto(
            MasterId: document.Id,
            MimeType: document.MimeType,
            ReturnUrl: returnUrl,
            Classification: classification
        ));
    }

    private sealed record ExternalMetadata(string? ReturnUrl, string? Classification);
}
