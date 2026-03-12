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

    public UploadDocumentCommandHandler(IDocumentRepository documentRepository)
    {
        _documentRepository = documentRepository;
    }

    public async Task<Result<UploadDocumentResult>> Handle(UploadDocumentCommand request, CancellationToken cancellationToken)
    {
        try
        {
            // Generuj GUID master dla dokumentu
            var masterId = Guid.NewGuid();

            // Utwórz dokument (aggregate root)
            var document = new Document(
                id: masterId,
                name: request.FileName,
                mimeType: request.MimeType,
                createdBy: request.CreatedBy
            );

            // Dodaj pierwszą wersję
            var version = document.AddVersion(
                content: request.Content,
                createdBy: request.CreatedBy
            );

            // Zapisz w bazie
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
