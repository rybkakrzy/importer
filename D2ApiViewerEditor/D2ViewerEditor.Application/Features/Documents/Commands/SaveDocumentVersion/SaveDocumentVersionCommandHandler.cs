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
    private readonly IDocumentStorageService _storageService;

    public SaveDocumentVersionCommandHandler(IDocumentRepository documentRepository, IDocumentStorageService storageService)
    {
        _documentRepository = documentRepository;
        _storageService = storageService;
    }

    public async Task<Result<SaveDocumentVersionResult>> Handle(SaveDocumentVersionCommand request, CancellationToken cancellationToken)
    {
        try
        {
            // Pobierz dokument z wersjami
            var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
            if (document == null)
                return Result<SaveDocumentVersionResult>.NotFound();

            // Upload pliku do GCS
            var versionId = Guid.NewGuid();
            var storagePath = await _storageService.UploadAsync(
                versionId, request.Content, document.MimeType, cancellationToken);

            // Dodaj nową wersję z referencją do GCS (version.Id == klucz obiektu w GCS)
            var newVersion = document.AddVersion(
                id: versionId,
                storagePath: storagePath,
                sizeInBytes: request.Content.Length,
                createdBy: request.CreatedBy
            );

            // Zapis z edytora → utrwal CorporateKey ostatniego modyfikującego (z claimu `corpKey`).
            document.SetLastModifiedBy(request.CorporateKey);

            // Encja jest śledzona — SaveChanges utrwala nową wersję. Nie wołamy _context.Update na całym
            // agregacie, bo wymusiłby pełny UPDATE documents (created_at jako Kind=Unspecified → błąd Npgsql).
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
