using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;
using Microsoft.Extensions.Logging;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UnlockDocument;

public class UnlockDocumentCommandHandler
    : IRequestHandler<UnlockDocumentCommand, Result<UnlockDocumentResult>>
{
    private readonly IDocumentRepository _documentRepository;
    private readonly ILogger<UnlockDocumentCommandHandler> _logger;

    public UnlockDocumentCommandHandler(
        IDocumentRepository documentRepository,
        ILogger<UnlockDocumentCommandHandler> logger)
    {
        _documentRepository = documentRepository;
        _logger = logger;
    }

    public async Task<Result<UnlockDocumentResult>> Handle(
        UnlockDocumentCommand request,
        CancellationToken cancellationToken)
    {
        var document = await _documentRepository.GetByIdAsync(request.MasterId, cancellationToken);
        if (document == null)
            return Result<UnlockDocumentResult>.NotFound();

        if (document.Status is DocumentStatus.Sending or DocumentStatus.Sent or DocumentStatus.DeliveryFailed)
            return Result<UnlockDocumentResult>.Failure(
                $"Nie można odblokować dokumentu w stanie {document.Status}: wysyłka jest w toku lub zakończona.");

        if (document.Status != DocumentStatus.Editing)
        {
            // Already Saved — idempotent no-op so retries don't fail.
            return Result<UnlockDocumentResult>.Success(new UnlockDocumentResult(document.Id, Changed: false));
        }

        document.MarkSaved();
        await _documentRepository.SaveChangesAsync(cancellationToken);

        _logger.LogInformation(
            "Document {MasterId} unlocked via external API (reason: {Reason}).",
            document.Id,
            string.IsNullOrWhiteSpace(request.Reason) ? "(none)" : request.Reason);

        return Result<UnlockDocumentResult>.Success(new UnlockDocumentResult(document.Id, Changed: true));
    }
}
