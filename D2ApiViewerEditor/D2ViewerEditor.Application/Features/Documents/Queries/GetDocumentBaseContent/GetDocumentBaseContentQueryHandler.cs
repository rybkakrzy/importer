using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentBaseContent;

/// <summary>
/// Handler pobierający pierwszą (bazową) wersję dokumentu — oryginał przesłanego pliku.
/// </summary>
public class GetDocumentBaseContentQueryHandler
    : IRequestHandler<GetDocumentBaseContentQuery, Result<DocumentBaseContentDto>>
{
    private readonly IDocumentRepository _documentRepository;

    public GetDocumentBaseContentQueryHandler(IDocumentRepository documentRepository)
    {
        _documentRepository = documentRepository;
    }

    public async Task<Result<DocumentBaseContentDto>> Handle(
        GetDocumentBaseContentQuery request,
        CancellationToken cancellationToken)
    {
        var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
        if (document == null)
            return Result<DocumentBaseContentDto>.NotFound();

        var baseVersion = document.Versions
            .OrderBy(v => v.VersionNumber)
            .FirstOrDefault();

        if (baseVersion == null)
            return Result<DocumentBaseContentDto>.Failure($"Dokument {request.MasterId} nie ma żadnych wersji");

        var extension = GetExtension(document.MimeType);
        var dto = new DocumentBaseContentDto(
            FileName: $"{document.Name}_v{baseVersion.VersionNumber}{extension}",
            MimeType: document.MimeType,
            Content: baseVersion.Content
        );

        return Result<DocumentBaseContentDto>.Success(dto);
    }

    private static string GetExtension(string mimeType) => mimeType switch
    {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => ".docx",
        "application/msword" => ".doc",
        "application/pdf" => ".pdf",
        _ => string.Empty
    };
}
