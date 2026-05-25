using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using Microsoft.Extensions.Logging;

namespace D2ViewerEditor.Infrastructure.Services.Delivery;

/// <summary>
/// Wykonuje pojedynczą próbę wysyłki zaclaimowanego zadania i utrwala jego nowy stan.
/// Działa w dedykowanym scope (własny DbContext) — bezpieczne przy równoległym przetwarzaniu batcha.
/// </summary>
public class DeliveryAttemptRunner
{
    private readonly IDocumentDeliveryRepository _deliveryRepository;
    private readonly IDocumentRepository _documentRepository;
    private readonly IDocumentStorageService _storage;
    private readonly IDeliverySender _sender;
    private readonly IBackoffStrategy _backoff;
    private readonly ILogger<DeliveryAttemptRunner> _logger;

    public DeliveryAttemptRunner(
        IDocumentDeliveryRepository deliveryRepository,
        IDocumentRepository documentRepository,
        IDocumentStorageService storage,
        IDeliverySender sender,
        IBackoffStrategy backoff,
        ILogger<DeliveryAttemptRunner> logger)
    {
        _deliveryRepository = deliveryRepository;
        _documentRepository = documentRepository;
        _storage = storage;
        _sender = sender;
        _backoff = backoff;
        _logger = logger;
    }

    public async Task RunAsync(Guid deliveryId, CancellationToken cancellationToken)
    {
        var delivery = await _deliveryRepository.GetByIdAsync(deliveryId, cancellationToken);
        if (delivery is null || delivery.Status != DeliveryStatus.Sending)
            return; // już przetworzone / przejęte przez inny worker

        using var scope = _logger.BeginScope(new Dictionary<string, object>
        {
            ["deliveryId"] = delivery.Id,
            ["documentId"] = delivery.DocumentId,
            ["correlationId"] = delivery.CorrelationId,
            ["attempt"] = delivery.AttemptCount
        });

        try
        {
            var bytes = await _storage.DownloadAsync(delivery.SnapshotObjectName, cancellationToken);
            var dispatch = new DeliveryDispatch(delivery.Id, delivery.RecipientUrl, bytes, delivery.SnapshotSha256);
            var result = await _sender.SendAsync(dispatch, cancellationToken);

            switch (result.Outcome)
            {
                case DeliveryOutcome.Succeeded:
                    delivery.MarkSent();
                    await SetDocumentStatusAsync(delivery.DocumentId, DocumentStatus.Sent, cancellationToken);
                    _logger.LogInformation("Delivery sent to recipient");
                    break;

                case DeliveryOutcome.PermanentError:
                    delivery.MarkPermanentFailure(result.Error ?? "Permanent error");
                    await SetDocumentStatusAsync(delivery.DocumentId, DocumentStatus.DeliveryFailed, cancellationToken);
                    _logger.LogWarning("Delivery permanently failed: {Error}", result.Error);
                    break;

                default:
                    var scheduled = delivery.ScheduleRetryOrDeadLetter(result.Error ?? "Retryable error", _backoff);
                    if (!scheduled)
                    {
                        await SetDocumentStatusAsync(delivery.DocumentId, DocumentStatus.DeliveryFailed, cancellationToken);
                        _logger.LogWarning("Delivery dead-lettered after {Attempts} attempts: {Error}",
                            delivery.AttemptCount, result.Error);
                    }
                    else
                    {
                        _logger.LogWarning("Delivery retry scheduled at {NextAttemptAt}: {Error}",
                            delivery.NextAttemptAt, result.Error);
                    }
                    break;
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            // Shutdown w trakcie próby — nie utrwalamy stanu; lease wygaśnie i zadanie zostanie przejęte.
            _logger.LogInformation("Delivery attempt cancelled (shutdown), will be reclaimed after lease expiry");
            return;
        }
        catch (Exception ex)
        {
            // Np. brak snapshotu w GCS albo nieoczekiwany błąd — traktuj jako retryable do deadline.
            delivery.ScheduleRetryOrDeadLetter(ex.Message, _backoff);
            _logger.LogError(ex, "Delivery attempt threw, rescheduled to {NextAttemptAt}", delivery.NextAttemptAt);
        }

        await _deliveryRepository.SaveChangesAsync(cancellationToken);
    }

    private async Task SetDocumentStatusAsync(Guid documentId, DocumentStatus status, CancellationToken cancellationToken)
    {
        // Ten sam scope => ten sam DbContext co delivery; jeden SaveChanges utrwala oba.
        var document = await _documentRepository.GetByIdWithVersionsAsync(documentId, cancellationToken);
        if (document is null)
            return;

        switch (status)
        {
            case DocumentStatus.Sent: document.MarkSent(); break;
            case DocumentStatus.DeliveryFailed: document.MarkDeliveryFailed(); break;
        }
    }
}
