using D2ViewerEditor.Application.Features.Documents.Common;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UpdateCallbackUrl;

public class UpdateCallbackUrlCommandHandler
    : IRequestHandler<UpdateCallbackUrlCommand, Result>
{
    /// <summary>
    /// Belt-and-suspenders cap (real-world callback URLs sit well below this; longer values
    /// usually mean an embedded token we should not be storing).
    /// </summary>
    private const int MaxUrlLength = 2048;

    private readonly IDocumentRepository _documentRepository;

    public UpdateCallbackUrlCommandHandler(IDocumentRepository documentRepository)
    {
        _documentRepository = documentRepository;
    }

    public async Task<Result> Handle(UpdateCallbackUrlCommand request, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.CallbackUrl))
            return Result.Failure("Callback URL jest wymagany.");

        var url = request.CallbackUrl.Trim();
        if (url.Length > MaxUrlLength)
            return Result.Failure($"Callback URL nie może być dłuższy niż {MaxUrlLength} znaków.");

        if (!DocumentDelivery.IsValidRecipientUrl(url))
            return Result.Failure("Callback URL musi być absolutnym adresem http(s).");

        var document = await _documentRepository.GetByIdAsync(request.MasterId, cancellationToken);
        if (document == null)
            return Result.NotFound();

        if (document.Status is DocumentStatus.Sending or DocumentStatus.Sent or DocumentStatus.DeliveryFailed)
            return Result.Failure(
                $"Nie można zaktualizować callback URL dla dokumentu w stanie {document.Status}: wysyłka już została zlecona.");

        // Preserve every other field on the metadata blob — we only own ReturnUrl here.
        var existing = ExternalDocumentMetadata.Parse(document.Metadata);
        var newMetadata = (existing with { ReturnUrl = url }).Serialize();

        document.UpdateMetadata(newMetadata);
        await _documentRepository.SaveChangesAsync(cancellationToken);

        return Result.Success();
    }
}
