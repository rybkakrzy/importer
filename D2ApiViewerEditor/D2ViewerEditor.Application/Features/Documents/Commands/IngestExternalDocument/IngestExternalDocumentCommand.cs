using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.IngestExternalDocument;

/// <summary>
/// Komenda przyjęcia dokumentu od aplikacji zewnętrznej.
/// Dla DOCX tworzy master + wersję oryginalną + wersję edytowalną (kopię).
/// Dla PDF tworzy master + jedną wersję (oryginał).
/// </summary>
public record IngestExternalDocumentCommand(
    byte[] Content,
    string FileName,
    string MimeType,
    string CreatedBy,
    string? Metadata = null
) : IRequest<Result<IngestExternalDocumentResult>>;

/// <summary>
/// Wynik przyjęcia dokumentu zewnętrznego.
/// VersionId jest ustawiony tylko dla DOCX (id wersji edytowalnej).
/// Dla PDF VersionId pozostaje null — aplikacja zewnętrzna otrzymuje tylko MasterId.
/// </summary>
public record IngestExternalDocumentResult(
    Guid MasterId,
    Guid? VersionId,
    string FileName,
    DateTime CreatedAt
);
