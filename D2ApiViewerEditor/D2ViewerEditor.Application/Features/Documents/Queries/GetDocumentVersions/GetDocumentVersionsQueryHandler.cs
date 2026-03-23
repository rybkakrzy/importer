using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentVersions;

/// <summary>
/// Handler dla pobierania listy wersji dokumentu
/// </summary>
public class GetDocumentVersionsQueryHandler : IRequestHandler<GetDocumentVersionsQuery, Result<List<DocumentVersionDto>>>
{
    private readonly IDocumentRepository _documentRepository;

    public GetDocumentVersionsQueryHandler(IDocumentRepository documentRepository)
    {
        _documentRepository = documentRepository;
    }

    public async Task<Result<List<DocumentVersionDto>>> Handle(GetDocumentVersionsQuery request, CancellationToken cancellationToken)
    {
        try
        {
            // Pobierz dokument z wersjami
            var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
            if (document == null)
                return Result<List<DocumentVersionDto>>.NotFound();

            // Mapuj wersje na DTO (sortuj od najnowszej)
            var versions = document.Versions
                .OrderByDescending(v => v.VersionNumber)
                .Select(v => new DocumentVersionDto(
                    DocumentId: document.Id,
                    VersionId: v.Id,
                    VersionNumber: v.VersionNumber,
                    CreatedAt: v.CreatedAt,
                    CreatedBy: v.CreatedBy,
                    IsActive: v.IsActive,
                    SizeInBytes: v.SizeInBytes
                ))
                .ToList();

            return Result<List<DocumentVersionDto>>.Success(versions);
        }
        catch (Exception ex)
        {
            return Result<List<DocumentVersionDto>>.Failure($"Błąd podczas pobierania wersji: {ex.Message}");
        }
    }
}
