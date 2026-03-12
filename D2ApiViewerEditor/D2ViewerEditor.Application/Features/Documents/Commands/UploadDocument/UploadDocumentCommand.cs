using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UploadDocument;

/// <summary>
/// Komenda upload dokumentu - tworzy nowy dokument w bazie
/// </summary>
public record UploadDocumentCommand(
    byte[] Content,
    string FileName,
    string MimeType,
    string CreatedBy
) : IRequest<Result<UploadDocumentResult>>;

/// <summary>
/// Wynik upload dokumentu
/// </summary>
public record UploadDocumentResult(
    Guid MasterId,
    Guid VersionId,
    string FileName,
    DateTime CreatedAt
);
