using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocuments;

public class GetDocumentsQueryHandler : IRequestHandler<GetDocumentsQuery, Result<List<DocumentListItemDto>>>
{
    private readonly IDocumentRepository _documentRepository;

    public GetDocumentsQueryHandler(IDocumentRepository documentRepository)
    {
        _documentRepository = documentRepository;
    }

    public async Task<Result<List<DocumentListItemDto>>> Handle(GetDocumentsQuery request, CancellationToken cancellationToken)
    {
        try
        {
            var documents = await _documentRepository.GetAllAsync(request.Skip, request.Take, cancellationToken);

            var dtos = documents.Select(d =>
            {
                var activeVersion = d.GetActiveVersion();
                return new DocumentListItemDto(
                    MasterId: d.Id,
                    Name: d.Name,
                    MimeType: d.MimeType,
                    CreatedAt: d.CreatedAt,
                    ActiveVersionId: activeVersion?.Id ?? Guid.Empty,
                    VersionNumber: activeVersion?.VersionNumber ?? 0,
                    Status: d.Status.ToString()
                );
            }).ToList();

            return Result<List<DocumentListItemDto>>.Success(dtos);
        }
        catch (Exception ex)
        {
            return Result<List<DocumentListItemDto>>.Failure($"Błąd podczas pobierania dokumentów: {ex.Message}");
        }
    }
}
