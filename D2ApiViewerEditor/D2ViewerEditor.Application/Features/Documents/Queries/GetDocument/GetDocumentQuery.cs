using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocument;

/// <summary>
/// Query pobierania aktywnej wersji dokumentu
/// </summary>
public record GetDocumentQuery(
    Guid MasterId
) : IRequest<Result<DocumentDto>>;

/// <summary>
/// DTO dokumentu z aktywną wersją
/// </summary>
public record DocumentDto(
    Guid MasterId,
    string Name,
    string MimeType,
    DateTime CreatedAt,
    Guid ActiveVersionId,
    byte[] Content,
    int VersionNumber
);
