using D2ViewerEditor.Application.Features.Documents.Common;
using D2ViewerEditor.Application.Common.Security;
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
    private readonly IFileUploadSecurityService _uploadSecurityService;

    public UploadDocumentCommandHandler(
        IDocumentRepository documentRepository,
        IDocumentStorageService storageService,
        IFileUploadSecurityService uploadSecurityService)
    {
        _documentRepository = documentRepository;
        _storageService = storageService;
        _uploadSecurityService = uploadSecurityService;
    }

    public async Task<Result<UploadDocumentResult>> Handle(UploadDocumentCommand request, CancellationToken cancellationToken)
    {
        try
        {
            if (request.Content == null || request.Content.Length == 0)
                return Result<UploadDocumentResult>.Failure("Zawartość dokumentu nie może być pusta");

            var uploadValidation = await _uploadSecurityService.ValidateDocumentAsync(
                request.Content,
                request.FileName,
                request.MimeType,
                cancellationToken);

            if (!uploadValidation.IsValid)
            {
                // Bug 13625398: dla defektów PLIKU komunikaty uzgodnione z QA — identyczne jak na
                // ścieżce edytora (POST /open). Pozostałe odrzucenia (skaner, limity) bez zmian.
                return Result<UploadDocumentResult>.Failure(uploadValidation.Code switch
                {
                    UploadRejectionCode.SignatureMismatch => "Nieprawidłowy format dokumentu.",
                    UploadRejectionCode.DocxRequiredPartMissing or UploadRejectionCode.DocxInvalidArchive
                        => "Dokument jest uszkodzony.",
                    _ => $"Upload odrzucony ({uploadValidation.Code}): {uploadValidation.Error}"
                });
            }

            // Generuj GUID master dla dokumentu
            var masterId = Guid.NewGuid();
            var versionId = Guid.NewGuid();

            // Upload pliku do GCS
            var storagePath = await _storageService.UploadAsync(
                versionId, request.Content, request.MimeType, cancellationToken);

            // Local-upload flow (user-chosen file from disk): there is no external app and no
            // return URL, so the only way for the user to retrieve the edited file is to download
            // it. The system itself sets userDownload=true here — NOT the client request, so a
            // tampered upload cannot suppress it (rule 12).
            var localUploadMetadata = new ExternalDocumentMetadata(
                ReturnUrl: null,
                Classification: null,
                UserDownload: true).Serialize();

            var document = new Document(
                id: masterId,
                name: request.FileName,
                mimeType: request.MimeType,
                createdBy: request.CreatedBy,
                metadata: localUploadMetadata
            );

            // Dodaj pierwszą wersję z referencją do GCS (version.Id == klucz obiektu w GCS)
            var version = document.AddVersion(
                id: versionId,
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
