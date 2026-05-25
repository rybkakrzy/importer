using D2ViewerEditor.Domain.Interfaces;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Infrastructure.Services.Delivery;

/// <summary>
/// Worker wysyłki: cyklicznie claimuje paczki gotowych zadań (FOR UPDATE SKIP LOCKED)
/// i przetwarza je równolegle z ograniczoną współbieżnością. Odporny na restart (lease)
/// i na wiele instancji (claim atomowy). Jedno wolne zadanie nie blokuje pozostałych.
/// </summary>
public class DocumentDeliveryWorker : BackgroundService
{
    private readonly IServiceScopeFactory _scopeFactory;
    private readonly DeliveryWorkerOptions _options;
    private readonly ILogger<DocumentDeliveryWorker> _logger;
    private readonly string _workerId = $"{Environment.MachineName}:{Guid.NewGuid():N}";

    public DocumentDeliveryWorker(
        IServiceScopeFactory scopeFactory,
        IOptions<DeliveryWorkerOptions> options,
        ILogger<DocumentDeliveryWorker> logger)
    {
        _scopeFactory = scopeFactory;
        _options = options.Value;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _logger.LogInformation("DocumentDeliveryWorker started ({WorkerId})", _workerId);
        using var timer = new PeriodicTimer(_options.PollInterval);

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                // Drenaż: dopóki batch jest pełny, czerp dalej bez czekania na tick.
                int processed;
                do { processed = await ProcessBatchAsync(stoppingToken); }
                while (processed == _options.BatchSize && !stoppingToken.IsCancellationRequested);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                break;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "DocumentDeliveryWorker batch failed");
            }

            if (!await SafeWaitAsync(timer, stoppingToken))
                break;
        }

        _logger.LogInformation("DocumentDeliveryWorker stopped ({WorkerId})", _workerId);
    }

    private async Task<int> ProcessBatchAsync(CancellationToken cancellationToken)
    {
        using var scope = _scopeFactory.CreateScope();
        var repository = scope.ServiceProvider.GetRequiredService<IDocumentDeliveryRepository>();

        var batch = await repository.ClaimDueBatchAsync(
            _options.BatchSize, _options.Lease, _workerId, cancellationToken);

        if (batch.Count == 0)
            return 0;

        _logger.LogInformation("Claimed {Count} delivery job(s)", batch.Count);

        await Parallel.ForEachAsync(
            batch,
            new ParallelOptions { MaxDegreeOfParallelism = _options.MaxConcurrency, CancellationToken = cancellationToken },
            async (claimed, token) =>
            {
                using var jobScope = _scopeFactory.CreateScope();
                var runner = jobScope.ServiceProvider.GetRequiredService<DeliveryAttemptRunner>();
                await runner.RunAsync(claimed.Id, token);
            });

        return batch.Count;
    }

    private static async Task<bool> SafeWaitAsync(PeriodicTimer timer, CancellationToken cancellationToken)
    {
        try { return await timer.WaitForNextTickAsync(cancellationToken); }
        catch (OperationCanceledException) { return false; }
    }
}
