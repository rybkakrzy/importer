using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocument;

/// <summary>
/// Handler dla pobierania dokumentu
/// </summary>
public class GetDocumentQueryHandler : IRequestHandler<GetDocumentQuery, Result<DocumentDto>>
{
    private readonly IDocumentRepository _documentRepository;

    public GetDocumentQueryHandler(IDocumentRepository documentRepository)
    {
        _documentRepository = documentRepository;
    }

    public async Task<Result<DocumentDto>> Handle(GetDocumentQuery request, CancellationToken cancellationToken)
    {
        try
        {
            // Pobierz dokument z wersjami
            var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
            if (document == null)
                return Result<DocumentDto>.Failure($"Dokument {request.MasterId} nie istnieje");

            // Pobierz aktywną wersję
            var activeVersion = document.GetActiveVersion();
            if (activeVersion == null)
                return Result<DocumentDto>.Failure($"Dokument {request.MasterId} nie ma aktywnej wersji");

            var dto = new DocumentDto(
                MasterId: document.Id,
                Name: document.Name,
                MimeType: document.MimeType,
                CreatedAt: document.CreatedAt,
                ActiveVersionId: activeVersion.Id,
                Content: activeVersion.Content,
                VersionNumber: activeVersion.VersionNumber
            );

            return Result<DocumentDto>.Success(dto);
        }
        catch (Exception ex)
        {
            return Result<DocumentDto>.Failure($"Błąd podczas pobierania dokumentu: {ex.Message}");
        }
    }
}
