using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentBaseContent;

/// <summary>
/// Query pobierania bazowej (pierwszej) wersji dokumentu — oryginalny plik po uploadzie.
/// </summary>
public record GetDocumentBaseContentQuery(
    Guid MasterId
) : IRequest<Result<DocumentBaseContentDto>>;

/// <summary>
/// DTO z bazową zawartością dokumentu gotową do bezpośredniego otwarcia.
/// </summary>
public record DocumentBaseContentDto(
    string FileName,
    string MimeType,
    byte[] Content
);
