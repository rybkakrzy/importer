using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UploadDocument;

/// <summary>
/// Handler dla upload dokumentu
/// </summary>
public class UploadDocumentCommandHandler : IRequestHandler<UploadDocumentCommand, Result<UploadDocumentResult>>
{
    private readonly IDocumentRepository _documentRepository;
    private readonly IDocumentStorageService _storageService;

    public UploadDocumentCommandHandler(IDocumentRepository documentRepository, IDocumentStorageService storageService)
    {
        _documentRepository = documentRepository;
        _storageService = storageService;
    }

    public async Task<Result<UploadDocumentResult>> Handle(UploadDocumentCommand request, CancellationToken cancellationToken)
    {
        try
        {
            if (request.Content == null || request.Content.Length == 0)
                return Result<UploadDocumentResult>.Failure("Zawartość dokumentu nie może być pusta");

            // Generuj GUID master dla dokumentu
            var masterId = Guid.NewGuid();
            var versionId = Guid.NewGuid();

            // Upload pliku do GCS
            var storagePath = await _storageService.UploadAsync(
                versionId, request.Content, request.MimeType, cancellationToken);

            // Utwórz dokument (aggregate root)
            var document = new Document(
                id: masterId,
                name: request.FileName,
                mimeType: request.MimeType,
                createdBy: request.CreatedBy
            );

            // Dodaj pierwszą wersję z referencją do GCS
            var version = document.AddVersion(
                storagePath: storagePath,
                sizeInBytes: request.Content.Length,
                createdBy: request.CreatedBy
            );

            // Zapisz metadane w bazie
            await _documentRepository.AddAsync(document, cancellationToken);
            await _documentRepository.SaveChangesAsync(cancellationToken);

            return Result<UploadDocumentResult>.Success(new UploadDocumentResult(
                MasterId: document.Id,
                VersionId: version.Id,
                FileName: document.Name,
                CreatedAt: document.CreatedAt
            ));
        }
        catch (Exception ex)
        {
            return Result<UploadDocumentResult>.Failure($"Błąd podczas upload dokumentu: {ex.Message}");
        }
    }
}
