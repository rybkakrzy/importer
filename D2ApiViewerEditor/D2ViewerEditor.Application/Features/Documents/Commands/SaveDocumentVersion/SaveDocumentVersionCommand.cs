using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.SaveDocumentVersion;

/// <summary>
/// Komenda zapisu nowej wersji dokumentu (z GUI)
/// </summary>
public record SaveDocumentVersionCommand(
    Guid MasterId,
    byte[] Content,
    string CreatedBy,
    string? CorporateKey = null
) : IRequest<Result<SaveDocumentVersionResult>>;

/// <summary>
/// Wynik zapisu nowej wersji
/// </summary>
public record SaveDocumentVersionResult(
    Guid VersionId,
    int VersionNumber,
    DateTime CreatedAt
);
