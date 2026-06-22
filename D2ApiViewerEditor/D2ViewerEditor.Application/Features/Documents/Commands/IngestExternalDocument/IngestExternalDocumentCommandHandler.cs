using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Application.Features.Documents.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.IngestExternalDocument;

public class IngestExternalDocumentCommandHandler
    : IRequestHandler<IngestExternalDocumentCommand, Result<IngestExternalDocumentResult>>
{
    public const string DocxMimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    public const string PdfMimeType = "application/pdf";

    private readonly IDocumentRepository _documentRepository;
    private readonly IDocumentStorageService _storageService;
    private readonly IFileUploadSecurityService _uploadSecurityService;
    private readonly IReturnUrlValidator _returnUrlValidator;

    public IngestExternalDocumentCommandHandler(
        IDocumentRepository documentRepository,
        IDocumentStorageService storageService,
        IFileUploadSecurityService uploadSecurityService,
        IReturnUrlValidator returnUrlValidator)
    {
        _documentRepository = documentRepository;
        _storageService = storageService;
        _uploadSecurityService = uploadSecurityService;
        _returnUrlValidator = returnUrlValidator;
    }

    public async Task<Result<IngestExternalDocumentResult>> Handle(
        IngestExternalDocumentCommand request,
        CancellationToken cancellationToken)
    {
        try
        {
            if (request.Content == null || request.Content.Length == 0)
                return Result<IngestExternalDocumentResult>.Failure("Zawartość dokumentu nie może być pusta");

            var isDocx = string.Equals(request.MimeType, DocxMimeType, StringComparison.OrdinalIgnoreCase);
            var isPdf = string.Equals(request.MimeType, PdfMimeType, StringComparison.OrdinalIgnoreCase);

            if (!isDocx && !isPdf)
                return Result<IngestExternalDocumentResult>.Failure(
                    $"Nieobsługiwany typ pliku: {request.MimeType}. Wspierane: DOCX, PDF.");

            var uploadValidation = await _uploadSecurityService.ValidateDocumentAsync(
                request.Content,
                request.FileName,
                request.MimeType,
                cancellationToken);

            if (!uploadValidation.IsValid)
                return Result<IngestExternalDocumentResult>.Failure(
                    $"Upload odrzucony ({uploadValidation.Code}): {uploadValidation.Error}");

            var metadata = ExternalDocumentMetadata.Parse(request.Metadata);
            if (isDocx && string.IsNullOrWhiteSpace(metadata.ReturnUrl))
                return Result<IngestExternalDocumentResult>.Failure("Brak poprawnego adresu odbiorcy (returnUrl) w metadanych dokumentu");

            if (!string.IsNullOrWhiteSpace(metadata.ReturnUrl))
            {
                var returnUrlValidation = _returnUrlValidator.Validate(metadata.ReturnUrl);
                if (!returnUrlValidation.IsValid)
                    return Result<IngestExternalDocumentResult>.Failure(returnUrlValidation.Error!);
            }

            var document = new Document(
                id: Guid.NewGuid(),
                name: request.FileName,
                mimeType: request.MimeType,
                createdBy: request.CreatedBy,
                metadata: request.Metadata
            );

            // Wersja oryginalna (v1) — to co przysłała aplikacja zewnętrzna, niedotykalna
            var originalVersionId = Guid.NewGuid();
            var originalStoragePath = await _storageService.UploadAsync(
                originalVersionId, request.Content, request.MimeType, cancellationToken);

            var originalVersion = document.AddVersion(
                id: originalVersionId,
                storagePath: originalStoragePath,
                sizeInBytes: request.Content.Length,
                createdBy: request.CreatedBy
            );

            Guid? editableVersionId = null;

            if (isDocx)
            {
                // Wersja edytowalna (v2) — kopia oryginału, na niej działa edytor + auto-save
                var copyVersionId = Guid.NewGuid();
                var copyStoragePath = await _storageService.UploadAsync(
                    copyVersionId, request.Content, request.MimeType, cancellationToken);

                var editableVersion = document.AddVersion(
                    id: copyVersionId,
                    storagePath: copyStoragePath,
                    sizeInBytes: request.Content.Length,
                    createdBy: request.CreatedBy
                );

                editableVersionId = editableVersion.Id;
            }

            await _documentRepository.AddAsync(document, cancellationToken);
            await _documentRepository.SaveChangesAsync(cancellationToken);

            return Result<IngestExternalDocumentResult>.Success(new IngestExternalDocumentResult(
                MasterId: document.Id,
                VersionId: editableVersionId,
                FileName: document.Name,
                CreatedAt: document.CreatedAt
            ));
        }
        catch (Exception ex)
        {
            // Rozwiń łańcuch inner exceptions — DbUpdateException chowa realną przyczynę (np. brak kolumny, naruszenie constraintu).
            var details = ex.Message;
            for (var inner = ex.InnerException; inner != null; inner = inner.InnerException)
                details += $" -> {inner.Message}";

            return Result<IngestExternalDocumentResult>.Failure(
                $"Błąd podczas przyjęcia dokumentu zewnętrznego: {details}");
        }
    }
}
