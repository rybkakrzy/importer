using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.RestoreDocumentVersion;

/// <summary>
/// Komenda przywrócenia wcześniejszej wersji dokumentu
/// </summary>
public record RestoreDocumentVersionCommand(
    Guid MasterId,
    Guid VersionId
) : IRequest<Result<RestoreDocumentVersionResult>>;

/// <summary>
/// Wynik przywrócenia wersji
/// </summary>
public record RestoreDocumentVersionResult(
    Guid VersionId,
    int VersionNumber,
    string Message
);
