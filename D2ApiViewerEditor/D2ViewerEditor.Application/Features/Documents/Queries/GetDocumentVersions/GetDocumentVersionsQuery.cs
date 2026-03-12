using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentVersions;

/// <summary>
/// Query pobierania listy wersji dokumentu
/// </summary>
public record GetDocumentVersionsQuery(
    Guid MasterId
) : IRequest<Result<List<DocumentVersionDto>>>;

/// <summary>
/// DTO wersji dokumentu (bez content - tylko metadane)
/// </summary>
public record DocumentVersionDto(
    Guid VersionId,
    int VersionNumber,
    DateTime CreatedAt,
    string CreatedBy,
    bool IsActive,
    long SizeInBytes
);
