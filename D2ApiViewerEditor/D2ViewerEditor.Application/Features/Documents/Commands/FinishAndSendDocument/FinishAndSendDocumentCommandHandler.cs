using System.Security.Cryptography;
using System.Text.Json;
using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.FinishAndSendDocument;

public class FinishAndSendDocumentCommandHandler
    : IRequestHandler<FinishAndSendDocumentCommand, Result<FinishAndSendResult>>
{
    // Twardy limit ponawiania zgodny z wymaganiem (24 h od utworzenia zadania).
    private static readonly TimeSpan RetentionWindow = TimeSpan.FromHours(24);
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    private readonly IDocumentRepository _documentRepository;
    private readonly IDocumentDeliveryRepository _deliveryRepository;
    private readonly IDocumentStorageService _storage;
    private readonly ICurrentUserProvider _currentUser;

    public FinishAndSendDocumentCommandHandler(
        IDocumentRepository documentRepository,
        IDocumentDeliveryRepository deliveryRepository,
        IDocumentStorageService storage,
        ICurrentUserProvider currentUser)
    {
        _documentRepository = documentRepository;
        _deliveryRepository = deliveryRepository;
        _storage = storage;
        _currentUser = currentUser;
    }

    public async Task<Result<FinishAndSendResult>> Handle(
        FinishAndSendDocumentCommand request, CancellationToken cancellationToken)
    {
        try
        {
            if (request.Content == null || request.Content.Length == 0)
                return Result<FinishAndSendResult>.Failure("Zawartość dokumentu nie może być pusta");

            var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
            if (document == null)
                return Result<FinishAndSendResult>.NotFound();

            var version = document.Versions.FirstOrDefault(v => v.Id == request.VersionId);
            if (version == null)
                return Result<FinishAndSendResult>.NotFound();

            var recipientUrl = ReadReturnUrl(document.Metadata);
            if (!DocumentDelivery.IsValidRecipientUrl(recipientUrl))
                return Result<FinishAndSendResult>.Failure(
                    "Brak poprawnego adresu odbiorcy (returnUrl) w metadanych dokumentu");

            // Idempotencja wielokrotnego kliknięcia: jeśli jest aktywne zadanie — zwróć je.
            var active = await _deliveryRepository.GetActiveByDocumentIdAsync(document.Id, cancellationToken);
            if (active != null)
                return Result<FinishAndSendResult>.Success(new FinishAndSendResult(active.Id, active.Status.ToString()));

            // 1. Utrwal stan edytora (nadpisanie wersji edytowalnej w miejscu — domena pilnuje v1-immutable).
            await _storage.UploadAsync(version.Id, request.Content, document.MimeType, cancellationToken);
            document.UpdateVersion(version.Id, request.Content.Length);

            // 2. Zamroź niezmienny snapshot finalnego pliku w GCS.
            var deliveryId = Guid.NewGuid();
            var snapshotObjectName = $"deliveries/{deliveryId}";
            var sha256 = Convert.ToHexString(SHA256.HashData(request.Content));
            await _storage.UploadRawAsync(snapshotObjectName, request.Content, document.MimeType, cancellationToken);

            // 3. Status dokumentu + zadanie wysyłki — atomowo (jeden SaveChanges, ten sam DbContext).
            document.MarkSending();
            var delivery = DocumentDelivery.Create(
                id: deliveryId,
                documentId: document.Id,
                sourceVersionId: version.Id,
                snapshotObjectName: snapshotObjectName,
                snapshotSizeBytes: request.Content.Length,
                snapshotSha256: sha256,
                recipientUrl: recipientUrl!,
                createdBy: request.CreatedBy ?? document.CreatedBy,
                correlationId: Guid.NewGuid(),
                retentionWindow: RetentionWindow,
                corporateKey: _currentUser.CorporateKey);

            await _deliveryRepository.AddAsync(delivery, cancellationToken);

            try
            {
                await _deliveryRepository.SaveChangesAsync(cancellationToken);
            }
            catch (Exception)
            {
                // Wyścig: równoległe drugie kliknięcie mogło wygrać unique index
                // (jedno aktywne zadanie na dokument). Jeśli zwycięzca istnieje — zwróć go; inaczej propaguj.
                var winner = await _deliveryRepository.GetActiveByDocumentIdAsync(document.Id, cancellationToken);
                if (winner != null)
                    return Result<FinishAndSendResult>.Success(new FinishAndSendResult(winner.Id, winner.Status.ToString()));
                throw;
            }

            return Result<FinishAndSendResult>.Success(new FinishAndSendResult(delivery.Id, delivery.Status.ToString()));
        }
        catch (InvalidOperationException ex)
        {
            // np. próba finalizacji na wersji oryginalnej (v1)
            return Result<FinishAndSendResult>.Failure(ex.Message);
        }
        catch (Exception ex)
        {
            var details = ex.Message;
            for (var inner = ex.InnerException; inner != null; inner = inner.InnerException)
                details += $" -> {inner.Message}";

            return Result<FinishAndSendResult>.Failure($"Błąd podczas kończenia i wysyłki dokumentu: {details}");
        }
    }

    private static string? ReadReturnUrl(string? metadata)
    {
        if (string.IsNullOrWhiteSpace(metadata))
            return null;

        try
        {
            var parsed = JsonSerializer.Deserialize<ExternalMetadata>(metadata, JsonOptions);
            return parsed?.ReturnUrl;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private sealed record ExternalMetadata(string? ReturnUrl, string? Classification);
}
