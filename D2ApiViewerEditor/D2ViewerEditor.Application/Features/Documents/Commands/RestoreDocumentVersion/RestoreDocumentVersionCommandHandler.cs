using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.RestoreDocumentVersion;

/// <summary>
/// Handler dla przywrócenia wersji dokumentu
/// </summary>
public class RestoreDocumentVersionCommandHandler : IRequestHandler<RestoreDocumentVersionCommand, Result<RestoreDocumentVersionResult>>
{
    private readonly IDocumentRepository _documentRepository;

    public RestoreDocumentVersionCommandHandler(IDocumentRepository documentRepository)
    {
        _documentRepository = documentRepository;
    }

    public async Task<Result<RestoreDocumentVersionResult>> Handle(RestoreDocumentVersionCommand request, CancellationToken cancellationToken)
    {
        try
        {
            // Pobierz dokument z wersjami
            var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
            if (document == null)
                return Result<RestoreDocumentVersionResult>.Failure($"Dokument {request.MasterId} nie istnieje");

            // Znajdź wersję do przywrócenia
            var versionToRestore = document.Versions.FirstOrDefault(v => v.Id == request.VersionId);
            if (versionToRestore == null)
                return Result<RestoreDocumentVersionResult>.Failure($"Wersja {request.VersionId} nie należy do dokumentu {request.MasterId}");

            // Przywróć wersję (ustaw jako aktywną)
            document.RestoreVersion(request.VersionId);

            // Zapisz zmiany
            await _documentRepository.UpdateAsync(document, cancellationToken);
            await _documentRepository.SaveChangesAsync(cancellationToken);

            return Result<RestoreDocumentVersionResult>.Success(new RestoreDocumentVersionResult(
                VersionId: versionToRestore.Id,
                VersionNumber: versionToRestore.VersionNumber,
                Message: $"Przywrócono wersję {versionToRestore.VersionNumber}"
            ));
        }
        catch (Exception ex)
        {
            return Result<RestoreDocumentVersionResult>.Failure($"Błąd podczas przywracania wersji: {ex.Message}");
        }
    }
}
