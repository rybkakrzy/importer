using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UpdateDocumentVersion;

/// <summary>
/// Komenda nadpisania istniejącej wersji dokumentu w miejscu (auto-save edytora).
/// Plik w GCS jest podmieniany pod tym samym obiektem (documents/{versionId}) — nie tworzy nowych wersji.
/// Wersja oryginalna (v1) jest nietykalna.
/// </summary>
public record UpdateDocumentVersionCommand(
    Guid MasterId,
    Guid VersionId,
    byte[] Content,
    string? CorporateKey = null
) : IRequest<Result<UpdateDocumentVersionResult>>;

/// <summary>
/// Wynik nadpisania wersji.
/// </summary>
public record UpdateDocumentVersionResult(
    Guid VersionId,
    int VersionNumber,
    long SizeInBytes,
    DateTime ModifiedAt
);
