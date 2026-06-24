using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocuments;

public record GetDocumentsQuery(
    int Skip = 0,
    int Take = 200
) : IRequest<Result<List<DocumentListItemDto>>>;

public record DocumentListItemDto(
    Guid MasterId,
    string Name,
    string MimeType,
    DateTime CreatedAt,
    Guid ActiveVersionId,
    int VersionNumber,
    string Status,
    string? LastModifiedBy
);
