using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentMetadata;

/// <summary>
/// Zwraca metadane przysłane przez aplikację zewnętrzną (returnUrl, classification).
/// Wykorzystywane m.in. w kroku „Zakończ" do odesłania pliku.
/// </summary>
public record GetDocumentMetadataQuery(Guid MasterId) : IRequest<Result<DocumentMetadataDto>>;

/// <summary>
/// Zdeserializowane metadane dokumentu.
/// </summary>
public record DocumentMetadataDto(
    Guid MasterId,
    string MimeType,
    string? ReturnUrl,
    string? Classification
);
