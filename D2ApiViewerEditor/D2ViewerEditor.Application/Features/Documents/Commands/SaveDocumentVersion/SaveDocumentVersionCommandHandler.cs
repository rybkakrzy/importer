using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.SaveDocumentVersion;

/// <summary>
/// Handler dla zapisu nowej wersji dokumentu
/// </summary>
public class SaveDocumentVersionCommandHandler : IRequestHandler<SaveDocumentVersionCommand, Result<SaveDocumentVersionResult>>
{
    private readonly IDocumentRepository _documentRepository;

    public SaveDocumentVersionCommandHandler(IDocumentRepository documentRepository)
    {
        _documentRepository = documentRepository;
    }

    public async Task<Result<SaveDocumentVersionResult>> Handle(SaveDocumentVersionCommand request, CancellationToken cancellationToken)
    {
        try
        {
            // Pobierz dokument z wersjami
            var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
            if (document == null)
                return Result<SaveDocumentVersionResult>.Failure($"Dokument {request.MasterId} nie istnieje");

            // Dodaj nową wersję
            var newVersion = document.AddVersion(
                content: request.Content,
                createdBy: request.CreatedBy
            );

            // Zapisz zmiany
            await _documentRepository.UpdateAsync(document, cancellationToken);
            await _documentRepository.SaveChangesAsync(cancellationToken);

            return Result<SaveDocumentVersionResult>.Success(new SaveDocumentVersionResult(
                VersionId: newVersion.Id,
                VersionNumber: newVersion.VersionNumber,
                CreatedAt: newVersion.CreatedAt
            ));
        }
        catch (Exception ex)
        {
            return Result<SaveDocumentVersionResult>.Failure($"Błąd podczas zapisu wersji: {ex.Message}");
        }
    }
}
