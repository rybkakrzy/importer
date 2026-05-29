using System.Text.Json;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UpdateCallbackUrl;

public class UpdateCallbackUrlCommandHandler
    : IRequestHandler<UpdateCallbackUrlCommand, Result>
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

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

        // Reuse the same validator the delivery worker uses to call the URL — keeps the contract
        // consistent ("if it passes here, it can be delivered").
        if (!DocumentDelivery.IsValidRecipientUrl(url))
            return Result.Failure("Callback URL musi być absolutnym adresem http(s).");

        var document = await _documentRepository.GetByIdAsync(request.MasterId, cancellationToken);
        if (document == null)
            return Result.NotFound();

        // After "Zakończ" the delivery has already snapshotted its own RecipientUrl — changing
        // the master metadata at this point would be misleading (worker keeps the old URL).
        if (document.Status is DocumentStatus.Sending or DocumentStatus.Sent or DocumentStatus.DeliveryFailed)
            return Result.Failure(
                $"Nie można zaktualizować callback URL dla dokumentu w stanie {document.Status}: wysyłka już została zlecona.");

        var classification = TryReadClassification(document.Metadata);
        var newMetadata = JsonSerializer.Serialize(new ExternalMetadata(url, classification), JsonOptions);

        document.UpdateMetadata(newMetadata);
        await _documentRepository.SaveChangesAsync(cancellationToken);

        return Result.Success();
    }

    private static string? TryReadClassification(string? metadata)
    {
        if (string.IsNullOrWhiteSpace(metadata)) return null;
        try
        {
            return JsonSerializer.Deserialize<ExternalMetadata>(metadata, JsonOptions)?.Classification;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private sealed record ExternalMetadata(string? ReturnUrl, string? Classification);
}
