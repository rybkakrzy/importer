using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UpdateDocumentVersion;

/// <summary>
/// Handler nadpisania wersji edytowalnej w miejscu (auto-save).
/// Re-upload pod tym samym versionId podmienia obiekt w GCS — bez przyrostu liczby wersji.
/// </summary>
public class UpdateDocumentVersionCommandHandler
    : IRequestHandler<UpdateDocumentVersionCommand, Result<UpdateDocumentVersionResult>>
{
    private readonly IDocumentRepository _documentRepository;
    private readonly IDocumentStorageService _storageService;

    public UpdateDocumentVersionCommandHandler(
        IDocumentRepository documentRepository,
        IDocumentStorageService storageService)
    {
        _documentRepository = documentRepository;
        _storageService = storageService;
    }

    public async Task<Result<UpdateDocumentVersionResult>> Handle(
        UpdateDocumentVersionCommand request,
        CancellationToken cancellationToken)
    {
        try
        {
            if (request.Content == null || request.Content.Length == 0)
                return Result<UpdateDocumentVersionResult>.Failure("Zawartość dokumentu nie może być pusta");

            var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
            if (document == null)
                return Result<UpdateDocumentVersionResult>.NotFound();

            var version = document.Versions.FirstOrDefault(v => v.Id == request.VersionId);
            if (version == null)
                return Result<UpdateDocumentVersionResult>.NotFound();

            // Nadpisanie w miejscu: UploadAsync(versionId) buduje tę samą nazwę obiektu (documents/{versionId}),
            // więc istniejący plik w buckecie zostaje zastąpiony — nie powstaje nowy obiekt.
            await _storageService.UploadAsync(
                version.Id, request.Content, document.MimeType, cancellationToken);

            // Domena pilnuje, że v1 (oryginał) jest nietykalna.
            document.UpdateVersion(request.VersionId, request.Content.Length);

            // Encja jest już śledzona (GetByIdWithVersionsAsync bez AsNoTracking), więc SaveChanges
            // utrwala samą zmianę wersji. NIE wołamy UpdateAsync/_context.Update — to oznaczyłoby
            // cały agregat jako Modified i wygenerowało pełny UPDATE documents (z created_at odczytanym
            // jako Kind=Unspecified), co Npgsql odrzuca przy zapisie do timestamptz.
            await _documentRepository.SaveChangesAsync(cancellationToken);

            return Result<UpdateDocumentVersionResult>.Success(new UpdateDocumentVersionResult(
                VersionId: version.Id,
                VersionNumber: version.VersionNumber,
                SizeInBytes: version.SizeInBytes,
                ModifiedAt: version.ModifiedAt ?? version.CreatedAt
            ));
        }
        catch (InvalidOperationException ex)
        {
            // np. próba nadpisania wersji oryginalnej (v1)
            return Result<UpdateDocumentVersionResult>.Failure(ex.Message);
        }
        catch (Exception ex)
        {
            // Rozwiń łańcuch inner exceptions — DbUpdateException chowa realną przyczynę.
            var details = ex.Message;
            for (var inner = ex.InnerException; inner != null; inner = inner.InnerException)
                details += $" -> {inner.Message}";

            return Result<UpdateDocumentVersionResult>.Failure($"Błąd podczas nadpisywania wersji: {details}");
        }
    }
}
