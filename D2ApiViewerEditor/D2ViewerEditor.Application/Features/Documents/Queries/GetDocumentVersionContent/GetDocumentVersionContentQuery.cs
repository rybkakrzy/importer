using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentVersionContent;

/// <summary>
/// Query pobierania fizycznej zawartości konkretnej wersji dokumentu
/// </summary>
public record GetDocumentVersionContentQuery(
    Guid MasterId,
    Guid VersionId
) : IRequest<Result<DocumentVersionContentDto>>;

/// <summary>
/// DTO z zawartością wersji dokumentu
/// </summary>
public record DocumentVersionContentDto(
    string FileName,
    string MimeType,
    byte[] Content
);
