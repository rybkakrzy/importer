using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentVersionContent;

/// <summary>
/// Handler dla pobierania zawartości konkretnej wersji dokumentu
/// </summary>
public class GetDocumentVersionContentQueryHandler
    : IRequestHandler<GetDocumentVersionContentQuery, Result<DocumentVersionContentDto>>
{
    private readonly IDocumentRepository _documentRepository;
    private readonly IDocumentStorageService _storageService;
    private readonly IDocumentAccessGuard _accessGuard;

    public GetDocumentVersionContentQueryHandler(
        IDocumentRepository documentRepository,
        IDocumentStorageService storageService,
        IDocumentAccessGuard accessGuard)
    {
        _documentRepository = documentRepository;
        _storageService = storageService;
        _accessGuard = accessGuard;
    }

    public async Task<Result<DocumentVersionContentDto>> Handle(
        GetDocumentVersionContentQuery request,
        CancellationToken cancellationToken)
    {
        try
        {
            var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
            if (document == null)
                return Result<DocumentVersionContentDto>.NotFound();

            if (!_accessGuard.IsViewAllowed(document.Metadata))
                return Result<DocumentVersionContentDto>.Forbidden("Brak uprawnień do podglądu tego dokumentu.");

            var version = document.Versions.FirstOrDefault(v => v.Id == request.VersionId);
            if (version == null)
                return Result<DocumentVersionContentDto>.NotFound();

            // Pobierz content z GCS
            var content = await _storageService.DownloadAsync(version.StoragePath, cancellationToken);

            var dto = new DocumentVersionContentDto(
                FileName: $"{document.Name}_v{version.VersionNumber}{GetExtension(document.MimeType)}",
                MimeType: document.MimeType,
                Content: content
            );

            return Result<DocumentVersionContentDto>.Success(dto);
        }
        catch (Exception ex)
        {
            return Result<DocumentVersionContentDto>.Failure($"Błąd podczas pobierania wersji: {ex.Message}");
        }
    }

    private static string GetExtension(string mimeType) => mimeType switch
    {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => ".docx",
        "application/msword" => ".doc",
        "application/pdf" => ".pdf",
        _ => string.Empty
    };
}
